"""Apply the open-ended-question fixes from the error report to the activity pages.

    uv run python oeq_edit.py                     # dry run: says what it would change
    uv run python oeq_edit.py --limit 5           # dry run over the first five books
    uv run python oeq_edit.py --apply --limit 5   # actually save those five
    uv run python oeq_edit.py --apply             # every book with a flagged question

The second script here that writes, and the guards follow word_edit.py for the same
reason: a mistake overwrites text that children read. Dry run is the default and
--apply is required before anything is saved.

What makes this page different from the word form is that a wrong click destroys more
than a wrong value. Each question carries two buttons labelled "Remove":

    beside the <h3>          deletes the whole question
    beside "Change Audio"    removes the recorded audio, which is what we want

Only the second is ever touched, and it is found by looking for the Remove that shares
a parent with "Change Audio" -- never by position, because a question whose audio has
already been removed has just one Remove left and it is the destructive one. As a
backstop the questions are counted before and after editing, and a book that lost one
is abandoned unsaved.

The guards, in the order they run:

    The fix must be a question.  Three findings carry the reviewer's instruction
                                 ("Replace it with a question a student can answer...")
                                 where the replacement text belongs. Written out, that
                                 sentence becomes the question the child is asked, so
                                 anything that does not read as replacement text is
                                 dropped rather than typed.
    One fix per textarea.        Where two findings want to rewrite the same question,
                                 the run refuses to choose. CHOSEN records the ones
                                 settled by hand, with the reason.
    The page is checked first.   The question on the page must match what the report
                                 recorded, or the row is skipped: the suggested fix was
                                 written against text that is no longer there.
    The fill is read back.       React wires its change handlers up after the values
                                 appear, and text typed into that gap sits in the box
                                 without the form noticing. Reading the textarea back
                                 is what proves the edit registered.
    Update proves it too.        The button enables only once something really changed,
                                 so a disabled Update means the run would save nothing.
    Every book is journalled.    data/oeq_edits.jsonl records the before and after of
                                 each question, so a run can resume and any change can
                                 be traced or put back by hand.

Removing the audio is intended, not incidental: the recording reads the old question
aloud, so leaving it in place would have the narrator ask something the screen no longer
says. Both are staged in the browser and committed together by one Update, so a question
never ends up saved with a recording of the text it replaced.

    The audio half does not work at present, and not for a reason on this side. Asked
    to save a question with no audio, the server tries to re-record it and answers

        ELEVENLABS_API_KEY is not set

    with HTTP 200 and the failure in the body, so the page shows nothing and the save
    silently does not happen. It refuses the whole mutation, text included, which is
    what makes this a block rather than a nuisance: the text cannot be saved in the
    same visit that clears the recording.

    A text-only save works. --keep-audio does exactly that, and leaves a question
    reading correctly while the recording still speaks the old wording. Whether that
    trade is worth making is a decision about the product, not about this script, and
    the run was held until the key is configured. Once it is, the plain command does
    both halves together and no such state ever exists.
"""

from __future__ import annotations

import argparse
import json
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
JOURNAL = Path("data/oeq_edits.jsonl")
SHEET = "OEC"

UPDATE_TEXT = "Update"
CHANGE_AUDIO_TEXT = "Change Audio"
REMOVE_TEXT = "Remove"

# Saving is heavier on the server than reading, so it gets a longer gap than the scrape.
WRITE_DELAY_S = 5.0
# The textareas appear with their values before the page has wired up its change
# handlers; text filled into that gap lands in the box without the form registering it.
SETTLE_MS = 1500

# Two findings wanting to rewrite one question is a decision, not something to resolve by
# ordering. Each entry below was read against both findings and the book's own question.
CHOSEN = {
    ("940", "Q2"): "Why do you think animals make groups? "
                   "How does being in groups help animals?",
}
# Why: book 940's Q2 was flagged twice over, once as too hard for the level ("classified")
# and once as unanswerable before reading. The Too Hard rewrite, "How do we decide which
# animals belong in the same group?", drops the hard word but still asks for the answer
# the book gives. The one above is the only candidate that settles both, so it is the one
# that gets written.


# --------------------------------------------------------------------------------------
# The findings
# --------------------------------------------------------------------------------------


def usable(fix: str) -> bool:
    """True when the text reads as a replacement question rather than a note about one.

    The reviewer's own instruction sits in the fix column on three findings. It is
    ordinary prose, so nothing but its shape distinguishes it from a real rewrite, and
    the cost of not noticing is that the instruction becomes the question on screen.
    """
    fix = fix.strip()
    if not fix.endswith("?"):
        return False
    first = fix.split()[0].lower() if fix.split() else ""
    return first not in {"replace", "rewrite", "reword", "change", "swap", "use",
                         "make", "consider", "simplify", "add", "remove"}


def load_edits(paths, titles) -> tuple[dict, list[str]]:
    """book id -> {"Q1": {...}}, plus the notes explaining everything left out."""
    found: dict[tuple[str, str], list[dict]] = {}
    for path in paths:
        if not path.exists():
            print(f"  ! {path} not found -- skipped")
            continue
        for f in make_reports.load_findings(path, titles):
            if f["sheet"] != SHEET or not f["prefix"] == "Q":
                continue
            found.setdefault((f["book_id"], f["target"]), []).append(f)

    books: dict[str, dict] = {}
    notes: list[str] = []
    for (bid, target), fs in sorted(found.items()):
        good = [f for f in fs if usable(f["fix"])]
        for f in fs:
            if f not in good:
                notes.append(f"book {bid} {target} [{f['type']}]: the suggested fix "
                             f"reads as an instruction, not a question -- "
                             f"{f['fix'][:60]!r}")
        if not good:
            continue
        fixes = {f["fix"].strip() for f in good}
        if len(fixes) > 1:
            picked = CHOSEN.get((bid, target))
            if not picked:
                notes.append(f"book {bid} {target}: {len(fixes)} findings want different "
                             f"text and none is recorded in CHOSEN -- skipped")
                continue
            fix = picked
        else:
            fix = fixes.pop()
        # Every finding on a textarea quotes the same text it is replacing; that was
        # checked when the sheet was read, so the first is as good as any.
        books.setdefault(bid, {"book_id": bid, "book": good[0]["book"], "questions": {}})
        books[bid]["questions"][target] = {
            "target": target, "n": good[0]["n"], "fix": fix,
            "current": good[0]["current"].strip(),
            "types": sorted({f["type"] for f in good}),
        }
    return books, notes


def saved_questions() -> dict[str, set[str]]:
    """book id -> the questions the journal records as already written."""
    out: dict[str, set[str]] = {}
    if not JOURNAL.exists():
        return out
    for line in JOURNAL.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        rec = json.loads(line)
        # "already" is a question somebody else had put right before we got there. It
        # needs no work either, so it counts as settled for the purpose of resuming.
        if rec.get("status") not in ("saved", "already"):
            continue
        out.setdefault(str(rec["book_id"]), set()).update(rec.get("questions", []))
    return out


def outstanding(book: dict, done: dict[str, set[str]]) -> dict:
    """The book with the questions already written taken out of it.

    Asking only whether the book has been saved before would strand the second question
    of any book saved once for the first -- the same fault that once left three
    Vietnamese boxes empty for good.
    """
    have = done.get(book["book_id"], set())
    rest = {t: q for t, q in book["questions"].items() if t not in have}
    return {**book, "questions": rest}


def journal(rec: dict) -> None:
    JOURNAL.parent.mkdir(parents=True, exist_ok=True)
    with JOURNAL.open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(rec, ensure_ascii=False) + "\n")


# --------------------------------------------------------------------------------------
# The page
# --------------------------------------------------------------------------------------

# Finding is done in the page, acting is done through Playwright: the walk from a heading
# to its own textarea is easier to express here, but a value has to be typed through the
# real locator or the events React listens for never fire.
FIND_JS = r"""(qn) => {
  const heads = Array.from(document.querySelectorAll('h3'));
  const h3 = heads.find(h => h.innerText.trim() === 'Question ' + qn);
  if (!h3) return null;
  // The nearest ancestor holding a textarea is the question's own block; going by a
  // class name would tie this to the current markup.
  let box = h3;
  while (box && !box.querySelector('textarea')) box = box.parentElement;
  if (!box) return null;
  const areas = Array.from(document.querySelectorAll('textarea'));
  const ta = box.querySelector('textarea');
  const buttons = Array.from(document.querySelectorAll('button'));
  const change = Array.from(box.querySelectorAll('button'))
      .find(b => b.innerText.trim() === 'Change Audio');
  // The Remove that matters shares a parent with Change Audio. The other one lives up
  // beside the heading and deletes the question, so it is never looked for.
  let remove = null;
  if (change) {
    remove = Array.from(change.parentElement.querySelectorAll('button'))
        .find(b => b.innerText.trim() === 'Remove') || null;
  }
  return {
    textarea: areas.indexOf(ta),
    value: ta.value,
    hasAudio: Boolean(change),
    audioName: change ? (change.parentElement.innerText.trim().split('\n')[0] || '') : '',
    removeButton: remove ? buttons.indexOf(remove) : -1,
  };
}"""

COUNT_JS = r"""() => Array.from(document.querySelectorAll('h3'))
    .filter(h => /^Question \d+$/.test(h.innerText.trim())).length"""


def find_question(page, n: int) -> dict | None:
    return page.evaluate(FIND_JS, n)


def question_count(page) -> int:
    return page.evaluate(COUNT_JS)


def update_button(page):
    return page.locator("button", has_text=UPDATE_TEXT).first


def open_book(page, url: str) -> None:
    rs.polite_goto(page, url)
    page.wait_for_selector("textarea", timeout=rs.SELECTOR_TIMEOUT_MS)
    page.wait_for_timeout(SETTLE_MS)


def fill_question(page, spot: dict, text: str) -> bool:
    """Type the replacement in and confirm the form took it."""
    box = page.locator("textarea").nth(spot["textarea"])
    box.fill("")
    box.fill(text)
    page.wait_for_timeout(400)
    if ec.same(box.input_value(), text):
        return True
    # One retry: the first fill can land before the handlers are wired.
    box.fill("")
    box.fill(text)
    page.wait_for_timeout(800)
    return ec.same(box.input_value(), text)


# --------------------------------------------------------------------------------------
# The run
# --------------------------------------------------------------------------------------


def run(args) -> None:
    sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
    _, rows = manual_qa.load_books()
    titles = {b: t for b, t, _ in rows}
    books, notes = load_edits(args.workbook, titles)
    for note in notes:
        print(f"  ! {note}")
    if notes:
        print()

    order = sorted(books.values(), key=lambda b: int(b["book_id"]))
    if args.only:
        keep = {s.strip() for s in args.only.split(",") if s.strip()}
        order = [b for b in order if b["book_id"] in keep]
    if not args.redo:
        done = saved_questions()
        order = [outstanding(b, done) for b in order]
        skipped = [b for b in order if not b["questions"]]
        if skipped:
            print(f"{len(skipped)} book(s) already have every flagged question "
                  f"written -- skipping\n")
        order = [b for b in order if b["questions"]]
    if args.limit:
        order = order[:args.limit]
    if not order:
        sys.exit("nothing to do")

    total_q = sum(len(b["questions"]) for b in order)
    mode = "APPLYING" if args.apply else "DRY RUN (nothing will be saved)"
    print(f"{mode}: {total_q} question(s) across {len(order)} book(s), about "
          f"{len(order) * (WRITE_DELAY_S + 4) / 60:.0f} min at current pacing\n")

    counts = {"saved": 0, "would save": 0, "already fixed": 0,
              "skipped": 0, "failed": 0}
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
        mutation = ec.Mutation(page, "OpenEndedQuestions")

        try:
            for i, book in enumerate(order, 1):
                bid = book["book_id"]
                label = f"[{i}/{len(order)}] {bid}"
                url = f"{rs.BASE_URL}/activities/{bid}/open-ended-questions/edit"
                try:
                    open_book(page, url)
                except (PlaywrightTimeout, PlaywrightError):
                    # A run outlives its login, and an expired session looks exactly
                    # like a dead server: the page simply never loads. Log in again and
                    # try once more before believing the server is in trouble.
                    print(f"{label}: page did not load -- logging in again")
                    try:
                        rs.ensure_logged_in(page, *rs.credentials())
                        rs.save_session(context)
                        open_book(page, url)
                    except (PlaywrightTimeout, PlaywrightError) as exc:
                        consecutive += 1
                        counts["failed"] += 1
                        print(f"{label}: {type(exc).__name__} loading the page")
                        journal({"book_id": bid, "status": "failed",
                                 "why": f"{type(exc).__name__} loading the page"})
                        if consecutive >= rs.MAX_CONSECUTIVE_FAILURES:
                            print(f"\n{consecutive} failures in a row -- stopping.")
                            break
                        continue

                before_count = question_count(page)
                applied, problems, record = [], [], {}
                settled: list[str] = []
                for target in sorted(book["questions"]):
                    q = book["questions"][target]
                    spot = find_question(page, q["n"])
                    if not spot:
                        problems.append(f"{target}: no question {q['n']} on the page")
                        continue
                    # Part of this list has already been fixed by hand. A page that
                    # now reads exactly what the report asked for is finished, not a
                    # conflict, and saying so keeps the real mismatches visible. Asked
                    # before the comparison with the report, because a fix differing
                    # only in the shape of its quotes normalises equal to the text it
                    # replaces, and the box would then be refilled with what it already
                    # says.
                    if ec.same(spot["value"], q["fix"]):
                        settled.append(target)
                        note = ("already reads the suggested text"
                                + (" (its audio is still in place)"
                                   if spot["hasAudio"] else ""))
                        print(f"{label}: {target} {note}")
                        continue
                    if norm(spot["value"]) != norm(q["current"]):
                        problems.append(
                            f"{target}: the page reads {spot['value'][:60]!r}, but the "
                            f"report recorded {q['current'][:60]!r} -- edited since")
                        continue
                    if not fill_question(page, spot, q["fix"]):
                        problems.append(f"{target}: the replacement would not take "
                                        f"in the form")
                        continue
                    record[target] = {"before": spot["value"].strip(),
                                      "after": q["fix"], "types": q["types"],
                                      "audio": spot["audioName"]}
                    # Re-read: removing the audio re-renders the block, so the button
                    # index found before the fill may no longer be the same button.
                    after_fill = find_question(page, q["n"])
                    if not args.audio:
                        record[target]["audio_removed"] = False
                    elif after_fill and after_fill["hasAudio"]:
                        if after_fill["removeButton"] < 0:
                            problems.append(f"{target}: audio present but no Remove "
                                            f"beside Change Audio -- left in place")
                        else:
                            page.locator("button").nth(
                                after_fill["removeButton"]).click()
                            page.wait_for_timeout(600)
                            still = find_question(page, q["n"])
                            if still and still["hasAudio"]:
                                problems.append(f"{target}: the audio did not clear")
                            else:
                                record[target]["audio_removed"] = True
                    else:
                        record[target]["audio_removed"] = False
                    if not args.audio and after_fill and after_fill["hasAudio"]:
                        record[target]["audio_kept"] = after_fill["audioName"]
                    applied.append(target)

                for p in problems:
                    print(f"{label}: ! {p}")

                # The backstop for the two Remove buttons. Nothing above should ever
                # touch the one that deletes a question, and if the count has dropped
                # then something did, so the form is abandoned rather than saved.
                after_count = question_count(page)
                if after_count != before_count:
                    counts["failed"] += 1
                    consecutive += 1
                    why = (f"the page had {before_count} questions before editing and "
                           f"{after_count} after -- not saving")
                    print(f"{label}: ABANDONED -- {why}")
                    journal({"book_id": bid, "status": "failed", "why": why})
                    if consecutive >= rs.MAX_CONSECUTIVE_FAILURES:
                        print(f"\n{consecutive} failures in a row -- stopping.")
                        break
                    continue

                if settled and args.apply:
                    # Recorded so a later run does not open the book again to re-read a
                    # question somebody else has already put right.
                    journal({"book_id": bid, "status": "already", "questions": settled,
                             "why": "the page already reads the suggested text"})

                if not applied:
                    consecutive = 0
                    counts["already fixed" if settled and not problems
                           else "skipped"] += 1
                    if not (settled and not problems):
                        print(f"{label}: SKIPPED -- nothing could be applied")
                        journal({"book_id": bid, "status": "skipped",
                                 "why": "; ".join(problems) or "nothing to change"})
                    continue

                if update_button(page).is_disabled():
                    consecutive = 0
                    counts["skipped"] += 1
                    print(f"{label}: SKIPPED -- the form says nothing changed")
                    journal({"book_id": bid, "status": "skipped",
                             "why": "update stayed disabled after filling",
                             "questions": applied})
                    continue

                rec = {"book_id": bid, "book": book["book"], "questions": applied,
                       "detail": record, "problems": problems}
                if not args.apply:
                    consecutive = 0
                    counts["would save"] += 1
                    print(f"{label}: would change {', '.join(applied)}"
                          f"  -- {book['book']}")
                    for t in applied:
                        d = record[t]
                        print(f"        {t} [{', '.join(d['types'])}]")
                        print(f"          now: {d['before']}")
                        print(f"          fix: {d['after']}")
                        if d.get("audio_removed"):
                            kept = f"would remove {d['audio']}"
                        elif d.get("audio_kept"):
                            kept = f"left in place: {d['audio_kept']}"
                        else:
                            kept = "none to remove"
                        print(f"          audio: {kept}")
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
                    journal({**rec, "status": "failed",
                             "why": f"{type(exc).__name__} clicking update"})
                    if consecutive >= rs.MAX_CONSECUTIVE_FAILURES:
                        print(f"\n{consecutive} failures in a row -- stopping.")
                        break
                    continue

                if refused:
                    # The server answered 200 and put its objection in the body. Saying
                    # what it objected to beats reporting that the text did not change.
                    consecutive += 1
                    counts["failed"] += 1
                    print(f"{label}: REFUSED -- {refused[0]}")
                    journal({**rec, "status": "failed", "why": refused[0]})
                    if any("ELEVENLABS" in e for e in refused):
                        print("\nThe server re-records a question left without audio, "
                              "and its speech key is unset,\nso every audio removal "
                              "will fail the same way. Stopping rather than asking\n"
                              "sixty more times. Re-run with --keep-audio to save the "
                              "text alone.")
                        break
                    if consecutive >= rs.MAX_CONSECUTIVE_FAILURES:
                        print(f"\n{consecutive} failures in a row -- stopping.")
                        break
                    continue

                if args.verify:
                    # The save has gone through by now, so a page that will not come
                    # back is a failure to confirm rather than a failure to save.
                    try:
                        open_book(page, url)
                    except (PlaywrightTimeout, PlaywrightError):
                        try:
                            rs.ensure_logged_in(page, *rs.credentials())
                            rs.save_session(context)
                            open_book(page, url)
                        except (PlaywrightTimeout, PlaywrightError) as exc:
                            consecutive += 1
                            counts["saved"] += 1
                            print(f"{label}: saved {', '.join(applied)}, but could not "
                                  f"reload to confirm ({type(exc).__name__})")
                            journal({**rec, "status": "saved",
                                     "why": "saved but not verified: "
                                            f"{type(exc).__name__} on reload"})
                            if consecutive >= rs.MAX_CONSECUTIVE_FAILURES:
                                print(f"\n{consecutive} failures in a row -- stopping.")
                                break
                            time.sleep(WRITE_DELAY_S)
                            continue
                    wrong = []
                    for t in applied:
                        q = book["questions"][t]
                        spot = find_question(page, q["n"])
                        if not spot or not ec.same(spot["value"], q["fix"]):
                            got = spot["value"] if spot else "<gone>"
                            wrong.append(f"{t} reads {got[:60]!r}")
                        elif spot["hasAudio"] and record[t].get("audio_removed"):
                            wrong.append(f"{t} still has its audio")
                    if wrong:
                        consecutive += 1
                        counts["failed"] += 1
                        print(f"{label}: SAVE DID NOT STICK -- {'; '.join(wrong)}")
                        journal({**rec, "status": "failed",
                                 "why": "reload did not show the new text",
                                 "on_reload": wrong})
                        if consecutive >= rs.MAX_CONSECUTIVE_FAILURES:
                            print(f"\n{consecutive} failures in a row -- stopping.")
                            break
                        continue

                consecutive = 0
                counts["saved"] += 1
                journal({**rec, "status": "saved"})
                print(f"{label}: saved {', '.join(applied)}  -- {book['book']}")
                time.sleep(WRITE_DELAY_S)
        finally:
            browser.close()

    print("\n" + "  ".join(f"{k}: {v}" for k, v in counts.items()))
    if not args.apply:
        print("\nDRY RUN -- nothing was saved. Add --apply to write these changes.")
    else:
        print(f"journal: {JOURNAL}")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
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
    ap.add_argument("--keep-audio", dest="audio", action="store_false",
                    help="save the question text but leave the recording alone; the "
                         "server refuses to save a question with no audio while its "
                         "speech key is unset")
    args = ap.parse_args()
    if not args.workbook:
        args.workbook = list(WORKBOOKS)
    run(args)


if __name__ == "__main__":
    main()
