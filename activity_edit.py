"""Apply the context-clue and text-multiple-choice fixes from the error report.

    uv run python activity_edit.py CC             # dry run: says what it would change
    uv run python activity_edit.py TMC --limit 5  # dry run over the first five books
    uv run python activity_edit.py CC --apply     # save

Both activities work the same way: find the box a finding names, check it still holds
the text the report quoted, replace it, and commit the book with one Update. Neither
page needs a destructive button clicked, which makes them simpler than the open-ended
questions -- but both carry buttons that delete an item outright, so a field is always
located through its own heading and never by counting.

CONTEXT CLUE holds ten cloze sentences, each a textarea under a heading "Item N", and a
finding on Q4 means the fourth of them. The heading is what is trusted; the position is
checked against it and a disagreement is reported rather than resolved, because the two
only ever disagree if the page is not what this script thinks it is.

TEXT MULTIPLE CHOICE holds a question and its answers under a heading "Question N".
The page is not laid out the way the report's field names suggest:

    the question text  a <textarea>, the only one in the question's block
    each answer        an <input placeholder="Answer text">, not a textarea, and with
                       no heading of its own -- AnsA4 is the first answer input inside
                       Question 4's block, AnsD4 the fourth

    Some questions have only three answers. A finding on AnsD there is asking for an
    option to be added rather than corrected, which also means choosing whether the new
    option is a distractor or the right answer. The report does not say which, and a
    wrong choice makes a wrong answer correct, so those are held back and listed.

An AnsAll finding quotes all four options at once and is expanded into four ordinary
edits, one per letter, so that a fix touching two options does not rewrite the other
two. One AnsAll finding is not a rewrite at all -- book 2841's question belongs to a
different book -- and it is held back with the rest.

The guards are the ones the vocabulary and open-ended-question runs arrived at:

    Dry run is the default.      --apply is required before anything is saved.
    The page is checked first.   The box must still hold what the report quoted, or the
                                 field is skipped: the fix was written against text that
                                 is no longer there. A box that already reads the
                                 suggested text is finished, not a conflict.
    The fill is read back.       React wires its change handlers up after the values
                                 appear, and text typed into that gap sits in the box
                                 without the form noticing.
    Update proves it too.        The button enables only once something really changed.
    The answer is read.          The save endpoint returns HTTP 200 with any failure in
                                 the body, so the body is read rather than the status.
    Every book is journalled.    data/<activity>_edits.jsonl records the before and
                                 after of each field, so a run can resume and any change
                                 can be traced or put back by hand.
"""

from __future__ import annotations

import argparse
import re
import sys
import time
from pathlib import Path

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import TimeoutError as PlaywrightTimeout
from playwright.sync_api import sync_playwright

import edit_common as ec
import make_reports
import manual_qa
import rs_scrape as rs
from word_ids import norm

WORKBOOKS = (Path("data/2plus_check.xlsx"), Path("data/2plus_check_batch2.xlsx"))
WRITE_DELAY_S = 5.0
SETTLE_MS = 1500
UPDATE_TEXT = "Update"

# --------------------------------------------------------------------------------------
# Decisions taken by hand
# --------------------------------------------------------------------------------------

# Where two findings want to rewrite the same box, the run refuses to choose. Each entry
# below was read against both suggestions and the item they replace. The rule applied
# throughout: keep the item and fix the fault. Where two rewrites both correct what was
# flagged, the one that leaves the original sentence standing is the one taken, because
# a context-clue item is written to carry its own target word and the shorter edit keeps
# that intact.
CHOSEN = {
    ("CC", "697", "Q2"): "The coach will remove him from the team.",
    ("CC", "798", "Q5"): "Her clothes were ripped, torn into pieces.",
    ("CC", "798", "Q7"): "The stories are similar, looking almost the same but not "
                         "exactly the same.",
    ("CC", "798", "Q8"): "They marry powerful men, ones with great strength and control.",
    ("CC", "808", "Q5"): "Things will get better soon, turning out nicer than they "
                         "are now.",
    ("CC", "808", "Q10"): "It turns out the school is nice, ending up better than I "
                          "expected.",
    ("CC", "1270", "Q9"): "The sewing machine stitches cloth far faster than by hand.",
    # The sibling question in the same book quotes its target phrase, and the Punctuation
    # finding asks for that to be matched; the alternative rewrote the question stem too.
    ("TMC", "808", "Q6"): "What is the closest meaning of “be able to”?",
}

# Findings whose "fix" is a note to a person rather than replacement text. There is one,
# and it is not a rewrite at all: the question and all four of its options belong to a
# different book. Nothing here can put that right.
UNUSABLE = {
    ("TMC", "2841", "AnsAll4"): "the question and all four options belong to a "
                                "different book -- needs rewriting by hand",
}

# --------------------------------------------------------------------------------------
# Finding the boxes
# --------------------------------------------------------------------------------------

# Context clue: the textarea under "Item N". Its position among all the textareas is
# returned as well, so the caller can check it against the N the report used.
CC_JS = r"""(n) => {
  const h3 = Array.from(document.querySelectorAll('h3'))
      .find(h => h.innerText.trim() === 'Item ' + n);
  if (!h3) return null;
  let box = h3;
  while (box && !box.querySelector('textarea')) box = box.parentElement;
  if (!box) return null;
  const ta = box.querySelector('textarea');
  const all = Array.from(document.querySelectorAll('textarea'));
  return {kind: 'textarea', index: all.indexOf(ta), value: ta.value,
          position: all.indexOf(ta) + 1};
}"""

# Text multiple choice: slot 0 is the question's own textarea, 1-4 its answer inputs.
# The question's block is the nearest ancestor of the heading that holds an answer
# input; that block holds exactly one textarea, which is checked rather than assumed.
TMC_JS = r"""({n, slot}) => {
  const h3 = Array.from(document.querySelectorAll('h3'))
      .find(h => h.innerText.trim() === 'Question ' + n);
  if (!h3) return null;
  const SEL = 'input[placeholder="Answer text"]';
  let box = h3;
  while (box && !box.querySelector(SEL)) box = box.parentElement;
  if (!box) return null;
  const own = Array.from(box.querySelectorAll(SEL));
  if (slot === 0) {
    const mine = Array.from(box.querySelectorAll('textarea'));
    if (mine.length !== 1) return {confused: mine.length};
    const all = Array.from(document.querySelectorAll('textarea'));
    return {kind: 'textarea', index: all.indexOf(mine[0]), value: mine[0].value,
            answers: own.length};
  }
  const el = own[slot - 1];
  if (!el) return {missing: true, answers: own.length};
  const all = Array.from(document.querySelectorAll(SEL));
  return {kind: 'input', index: all.indexOf(el), value: el.value, answers: own.length};
}"""

ANSWER_SELECTOR = 'input[placeholder="Answer text"]'


class Activity:
    def __init__(self, key, sheet, segment, marker, ready):
        self.key, self.sheet, self.segment = key, sheet, segment
        self.marker, self.ready = marker, ready
        self.journal = Path(f"data/{key.lower()}_edits.jsonl")

    def url(self, book_id: str) -> str:
        return f"{rs.BASE_URL}/activities/{book_id}/{self.segment}/edit"


ACTIVITIES = {
    "CC": Activity("CC", "CC", "context-clue", "ContextClue", "textarea"),
    "TMC": Activity("TMC", "TMC", "text-multiple-choice", "TextMultipleChoice",
                    ANSWER_SELECTOR),
}

TARGET = re.compile(r"^(?:Q|Ans(All|[A-D]))(\d+)$")


def parse_target(target: str) -> tuple[int, int] | None:
    """"AnsC4" -> (question 4, slot 3); "Q4" -> (4, 0). None if it is neither."""
    m = TARGET.fullmatch(target)
    if not m:
        return None
    letter, n = m.group(1), int(m.group(2))
    if letter is None:
        return n, 0
    if letter == "All":
        return None
    return n, "ABCD".index(letter) + 1


SPLIT = re.compile(r"(?:^|\s)([A-D])[.)]\s+")


def split_options(text: str) -> dict[str, str]:
    """The four options out of an AnsAll cell, newline- or space-separated.

    Four of the fifteen write the options on one line, so splitting on newlines alone
    silently yields a single option and would push all four answers into answer A.
    """
    flat = " ".join(line.strip() for line in (text or "").splitlines() if line.strip())
    parts = SPLIT.split(flat)
    out: dict[str, str] = {}
    for i in range(1, len(parts) - 1, 2):
        out[parts[i]] = parts[i + 1].strip()
    return out


# --------------------------------------------------------------------------------------
# The findings
# --------------------------------------------------------------------------------------


def load_edits(act: Activity, paths, titles) -> tuple[dict, list[str]]:
    """book id -> {field: {...}}, plus the notes explaining everything left out."""
    found: dict[tuple[str, str], list[dict]] = {}
    notes: list[str] = []
    for path in paths:
        if not path.exists():
            print(f"  ! {path} not found -- skipped")
            continue
        for f in make_reports.load_findings(path, titles):
            if f["sheet"] != act.sheet:
                continue
            bid, target = f["book_id"], f["target"]
            if (act.key, bid, target) in UNUSABLE:
                continue
            if not f["fix"].strip():
                notes.append(f"book {bid} {target} [{f['type']}]: no suggested fix "
                             f"-- {f['details'][:60]}")
                continue
            if f["prefix"] == "AnsAll":
                # Expand into one edit per letter so a fix touching two options does not
                # rewrite the other two with text nobody proposed.
                n = f["n"]
                now, fix = split_options(f["current"]), split_options(f["fix"])
                if set(now) != set("ABCD") or set(fix) != set("ABCD"):
                    notes.append(f"book {bid} {target}: the four options could not be "
                                 f"read out of the cell -- skipped")
                    continue
                for letter in "ABCD":
                    if now[letter].strip() == fix[letter].strip():
                        continue
                    found.setdefault((bid, f"Ans{letter}{n}"), []).append(
                        {**f, "target": f"Ans{letter}{n}", "prefix": f"Ans{letter}",
                         "current": now[letter], "fix": fix[letter],
                         "from": f["target"]})
                continue
            if parse_target(target) is None:
                notes.append(f"book {bid} {target} [{f['type']}]: not a field this page "
                             f"can edit -- skipped")
                continue
            found.setdefault((bid, target), []).append(f)

    books: dict[str, dict] = {}
    for (bid, target), fs in sorted(found.items()):
        fixes = {f["fix"].strip() for f in fs}
        if len(fixes) > 1:
            picked = CHOSEN.get((act.key, bid, target))
            if not picked:
                notes.append(f"book {bid} {target}: {len(fixes)} findings want different "
                             f"text and none is recorded in CHOSEN -- skipped")
                continue
            fix = picked
        else:
            fix = fixes.pop()
        spot = parse_target(target)
        books.setdefault(bid, {"book_id": bid, "book": fs[0]["book"], "fields": {}})
        books[bid]["fields"][target] = {
            "target": target, "n": spot[0], "slot": spot[1], "fix": fix,
            "current": fs[0]["current"].strip(),
            "types": sorted({f["type"] for f in fs}),
            "from": fs[0].get("from"),
        }
    return books, notes


def outstanding(book: dict, done: dict[str, set[str]]) -> dict:
    have = done.get(book["book_id"], set())
    return {**book, "fields": {t: q for t, q in book["fields"].items() if t not in have}}


# --------------------------------------------------------------------------------------
# The page
# --------------------------------------------------------------------------------------


def find_field(page, act: Activity, field: dict) -> dict | None:
    if act.key == "CC":
        return page.evaluate(CC_JS, field["n"])
    return page.evaluate(TMC_JS, {"n": field["n"], "slot": field["slot"]})


def box_for(page, spot: dict):
    if spot["kind"] == "textarea":
        return page.locator("textarea").nth(spot["index"])
    return page.locator(ANSWER_SELECTOR).nth(spot["index"])


def update_button(page):
    return page.locator("button", has_text=UPDATE_TEXT).first


def open_activity(page, act: Activity, url: str) -> None:
    rs.polite_goto(page, url)
    page.wait_for_selector(act.ready, timeout=rs.SELECTOR_TIMEOUT_MS)
    page.wait_for_timeout(SETTLE_MS)


def fill_box(page, spot: dict, text: str) -> bool:
    box = box_for(page, spot)
    box.fill("")
    box.fill(text)
    page.wait_for_timeout(400)
    if ec.same(box.input_value(), text):
        return True
    box.fill("")
    box.fill(text)
    page.wait_for_timeout(800)
    return ec.same(box.input_value(), text)


# --------------------------------------------------------------------------------------
# The run
# --------------------------------------------------------------------------------------


def run(args) -> None:
    sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
    act = ACTIVITIES[args.activity]
    _, rows = manual_qa.load_books()
    titles = {b: t for b, t, _ in rows}
    books, notes = load_edits(act, args.workbook, titles)
    for note in notes:
        print(f"  ! {note}")
    for (key, bid, target), why in sorted(UNUSABLE.items()):
        if key == act.key:
            print(f"  ! book {bid} {target}: {why}")
    if notes:
        print()

    order = sorted(books.values(), key=lambda b: int(b["book_id"]))
    if args.only:
        keep = {s.strip() for s in args.only.split(",") if s.strip()}
        order = [b for b in order if b["book_id"] in keep]
    if not args.redo:
        done = ec.settled_fields(act.journal)
        order = [outstanding(b, done) for b in order]
        already = sum(1 for b in order if not b["fields"])
        if already:
            print(f"{already} book(s) already have every flagged field written "
                  f"-- skipping\n")
        order = [b for b in order if b["fields"]]
    if args.limit:
        order = order[:args.limit]
    if not order:
        sys.exit("nothing to do")

    total = sum(len(b["fields"]) for b in order)
    mode = "APPLYING" if args.apply else "DRY RUN (nothing will be saved)"
    print(f"{mode}: {act.key}, {total} field(s) across {len(order)} book(s), about "
          f"{len(order) * (WRITE_DELAY_S + 4) / 60:.0f} min at current pacing\n")

    counts = {"saved": 0, "would save": 0, "already fixed": 0, "skipped": 0,
              "failed": 0, "held": 0}
    consecutive = 0
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=rs.HEADLESS)
        state = str(rs.STORAGE_STATE) if rs.STORAGE_STATE.exists() else None
        context = browser.new_context(storage_state=state)
        rs.block_heavy_resources(context)
        page = context.new_page()
        page.set_default_navigation_timeout(rs.NAV_TIMEOUT_MS)
        rs.ensure_logged_in(page, *rs.credentials())
        rs.save_session(context)
        mutation = ec.Mutation(page, act.marker)

        try:
            for i, book in enumerate(order, 1):
                bid = book["book_id"]
                label = f"[{i}/{len(order)}] {bid}"
                url = act.url(bid)
                try:
                    open_activity(page, act, url)
                except (PlaywrightTimeout, PlaywrightError):
                    # A run outlives its login, and an expired session looks exactly
                    # like a dead server: the page never finishes loading.
                    print(f"{label}: page did not load -- logging in again")
                    try:
                        rs.ensure_logged_in(page, *rs.credentials())
                        rs.save_session(context)
                        open_activity(page, act, url)
                    except (PlaywrightTimeout, PlaywrightError) as exc:
                        consecutive += 1
                        counts["failed"] += 1
                        print(f"{label}: {type(exc).__name__} loading the page")
                        ec.journal(act.journal, {"book_id": bid, "status": "failed",
                                                 "why": f"{type(exc).__name__} loading"})
                        if consecutive >= rs.MAX_CONSECUTIVE_FAILURES:
                            print(f"\n{consecutive} failures in a row -- stopping.")
                            break
                        continue

                applied, problems, settled, record = [], [], [], {}
                for target in sorted(book["fields"]):
                    fld = book["fields"][target]
                    spot = find_field(page, act, fld)
                    if not spot:
                        problems.append(f"{target}: no such item on the page")
                        continue
                    if spot.get("confused"):
                        problems.append(f"{target}: the question's block holds "
                                        f"{spot['confused']} question boxes, not one")
                        continue
                    if spot.get("missing"):
                        # A fourth option asked for on a three-option question. Adding
                        # one also means saying whether it is a distractor or the right
                        # answer, which the report does not record.
                        counts["held"] += 1
                        problems.append(
                            f"{target}: this question has only {spot['answers']} "
                            f"options, so the fix would add one rather than correct it "
                            f"-- held back, it needs an answer type choosing")
                        continue
                    if act.key == "CC" and spot.get("position") != fld["n"]:
                        problems.append(f"{target}: 'Item {fld['n']}' is textarea "
                                        f"{spot['position']} on the page -- not editing")
                        continue
                    # Asked first, and without reference to what the report quoted. A
                    # fix that differs from the original only in the shape of its quotes
                    # normalises equal to it, so a box already holding the suggested
                    # text would otherwise be refilled with what it already says and the
                    # book skipped for a form that rightly reports nothing changed.
                    if ec.same(spot["value"], fld["fix"]):
                        settled.append(target)
                        print(f"{label}: {target} already reads the suggested text")
                        continue
                    if norm(spot["value"]) != norm(fld["current"]):
                        problems.append(
                            f"{target}: the page reads {spot['value'][:55]!r}, but the "
                            f"report recorded {fld['current'][:55]!r} -- edited since")
                        continue
                    if not fill_box(page, spot, fld["fix"]):
                        problems.append(f"{target}: the replacement would not take "
                                        f"in the form")
                        continue
                    record[target] = {"before": spot["value"].strip(),
                                      "after": fld["fix"], "types": fld["types"],
                                      "from": fld.get("from")}
                    applied.append(target)

                for p in problems:
                    print(f"{label}: ! {p}")

                if settled and args.apply:
                    ec.journal(act.journal,
                               {"book_id": bid, "status": "already", "fields": settled,
                                "why": "the page already reads the suggested text"})

                if not applied:
                    consecutive = 0
                    clean = settled and not problems
                    counts["already fixed" if clean else "skipped"] += 1
                    if not clean:
                        print(f"{label}: SKIPPED -- nothing could be applied")
                        ec.journal(act.journal,
                                   {"book_id": bid, "status": "skipped",
                                    "why": "; ".join(problems) or "nothing to change"})
                    continue

                if update_button(page).is_disabled():
                    consecutive = 0
                    counts["skipped"] += 1
                    print(f"{label}: SKIPPED -- the form says nothing changed")
                    ec.journal(act.journal,
                               {"book_id": bid, "status": "skipped", "fields": applied,
                                "why": "update stayed disabled after filling"})
                    continue

                rec = {"book_id": bid, "book": book["book"], "fields": applied,
                       "detail": record, "problems": problems}
                if not args.apply:
                    consecutive = 0
                    counts["would save"] += 1
                    print(f"{label}: would change {', '.join(applied)}"
                          f"  -- {book['book']}")
                    for t in applied:
                        d = record[t]
                        src = f" (from {d['from']})" if d.get("from") else ""
                        print(f"        {t} [{', '.join(d['types'])}]{src}")
                        print(f"          now: {d['before']}")
                        print(f"          fix: {d['after']}")
                    continue

                try:
                    mutation.take()
                    update_button(page).click()
                    page.wait_for_timeout(2000)
                    refused = mutation.take()
                except (PlaywrightTimeout, PlaywrightError) as exc:
                    consecutive += 1
                    counts["failed"] += 1
                    print(f"{label}: {type(exc).__name__} clicking update")
                    ec.journal(act.journal, {**rec, "status": "failed",
                                             "why": f"{type(exc).__name__} on update"})
                    if consecutive >= rs.MAX_CONSECUTIVE_FAILURES:
                        print(f"\n{consecutive} failures in a row -- stopping.")
                        break
                    continue

                if refused:
                    # HTTP 200 with the objection in the body. Saying what the server
                    # objected to beats reporting that the text did not change.
                    consecutive += 1
                    counts["failed"] += 1
                    print(f"{label}: REFUSED -- {refused[0]}")
                    ec.journal(act.journal,
                               {**rec, "status": "failed", "why": refused[0]})
                    if consecutive >= rs.MAX_CONSECUTIVE_FAILURES:
                        print(f"\n{consecutive} failures in a row -- stopping.")
                        break
                    continue

                if args.verify:
                    # The save has gone through by now, so a page that will not come
                    # back is a failure to confirm rather than a failure to save.
                    try:
                        open_activity(page, act, url)
                    except (PlaywrightTimeout, PlaywrightError) as exc:
                        consecutive += 1
                        counts["saved"] += 1
                        print(f"{label}: saved {', '.join(applied)}, but could not "
                              f"reload to confirm ({type(exc).__name__})")
                        ec.journal(act.journal,
                                   {**rec, "status": "saved",
                                    "why": f"saved but not verified: "
                                           f"{type(exc).__name__} on reload"})
                        if consecutive >= rs.MAX_CONSECUTIVE_FAILURES:
                            print(f"\n{consecutive} failures in a row -- stopping.")
                            break
                        time.sleep(WRITE_DELAY_S)
                        continue
                    wrong = []
                    for t in applied:
                        spot = find_field(page, act, book["fields"][t])
                        got = spot.get("value") if spot else None
                        if got is None or not ec.same(got, book["fields"][t]["fix"]):
                            wrong.append(f"{t} reads {(got or '<gone>')[:55]!r}")
                    if wrong:
                        consecutive += 1
                        counts["failed"] += 1
                        print(f"{label}: SAVE DID NOT STICK -- {'; '.join(wrong)}")
                        ec.journal(act.journal,
                                   {**rec, "status": "failed", "on_reload": wrong,
                                    "why": "reload did not show the new text"})
                        if consecutive >= rs.MAX_CONSECUTIVE_FAILURES:
                            print(f"\n{consecutive} failures in a row -- stopping.")
                            break
                        continue

                consecutive = 0
                counts["saved"] += 1
                ec.journal(act.journal, {**rec, "status": "saved"})
                print(f"{label}: saved {', '.join(applied)}  -- {book['book']}")
                time.sleep(WRITE_DELAY_S)
        finally:
            browser.close()

    print("\n" + "  ".join(f"{k}: {v}" for k, v in counts.items()))
    if not args.apply:
        print("\nDRY RUN -- nothing was saved. Add --apply to write these changes.")
    else:
        print(f"journal: {act.journal}")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("activity", choices=sorted(ACTIVITIES),
                    help="which activity to edit")
    ap.add_argument("--apply", action="store_true",
                    help="actually save; without it the run only reports")
    ap.add_argument("--workbook", type=Path, action="append",
                    help="findings workbook; repeatable, defaults to both batches")
    ap.add_argument("--only", help="comma-separated book ids")
    ap.add_argument("--limit", type=int, help="stop after this many books")
    ap.add_argument("--redo", action="store_true",
                    help="ignore the journal and revisit books already written")
    ap.add_argument("--no-verify", dest="verify", action="store_false",
                    help="skip the reload that confirms each save")
    args = ap.parse_args()
    if not args.workbook:
        args.workbook = list(WORKBOOKS)
    run(args)


if __name__ == "__main__":
    main()
