"""Detect mechanical vocabulary faults and emit them as candidates for review.

This deliberately does NOT write findings. A rule proposes; a person disposes. Every
candidate carries the evidence that triggered it so it can be confirmed or rejected
quickly, and each check reports its own hit list rather than being folded into one
number -- a check that silently stops matching should be visible as a count going to
zero, not hidden inside a total.

Where one test would not catch most cases, several are applied and their results merged:
`no_definition` looks for an empty box, an exact repeat of the word, a repeat of just the
base form, and a suspiciously short non-phrase, because the authoring fault shows up in
all four shapes.

    uv run python vocab_checks.py --staged vocab_staged.json            # summary
    uv run python vocab_checks.py --staged ... --check no_definition    # one check
    uv run python vocab_checks.py --staged ... --out candidates.json    # for review
"""

from __future__ import annotations

import argparse
import collections
import json
import re
import sys
from pathlib import Path

WS = re.compile(r"[^a-z0-9]+")


def norm(text: str) -> str:
    return WS.sub(" ", (text or "").lower()).strip()


def head(word: str) -> str:
    """The base form: `stammer, (stammered)` -> `stammer`, `crate(s)` -> `crate`."""
    return norm(re.split(r"[,(]", word or "", maxsplit=1)[0])


# --------------------------------------------------------------------------------------
# Checks. Each takes one entry and returns (details, suggested fix) or None. `fix` may be
# empty where the fault needs a human to write the replacement.
# --------------------------------------------------------------------------------------


def no_definition(w: str, pos: str, dfn: str, sent: str):
    """The definition box carries no meaning. Four shapes of the same authoring fault."""
    if not dfn.strip():
        return ("the definition box is empty", "")
    if norm(dfn) == norm(w):
        return ("the definition box only repeats the word; no meaning is given", "")
    if norm(dfn) == head(w):
        return ("the definition box only repeats the base form of the word; "
                "no meaning is given", "")
    # A real definition is a phrase. One or two words with no article and no verb is
    # almost always the word echoed back in a slightly different form.
    if len(norm(dfn).split()) <= 2 and head(w).split()[:1] == norm(dfn).split()[:1]:
        return ("the definition box repeats the word rather than defining it", "")
    return None


# Endings that make a word the *same* word in another form. Anything else that merely
# starts with the base is a different word: `healthy` is not a form of `heal`, and
# `hippopotamus` is not a form of `hippo` -- defining an abbreviation by its full form
# is correct, not a giveaway.
INFLECTIONS = ("", "s", "es", "d", "ed", "ing", "ies", "er", "ers")


def word_in_definition(w: str, pos: str, dfn: str, sent: str):
    """The definition uses the word it is defining, so the match is free.

    Two guards, both added after reviewing false positives rather than assumed up front:

      * an entry whose definition box merely repeats the word is `no_definition`, not
        this one -- reporting it under both inflates each count and hides how much of
        the 296 was really the 206;
      * only a genuine inflection of the base counts, not any longer word that happens
        to start with it.
    """
    base = head(w)
    if not base or not dfn:
        return None
    if no_definition(w, pos, dfn, sent):
        return None
    # A two-letter base matches far too much to be evidence of anything.
    if len(base) < 3:
        return None
    if " " in base:
        # A multi-word base appearing whole cannot be coincidence.
        if re.search(rf"\b{re.escape(base)}\b", norm(dfn)):
            return (f'the definition uses the vocabulary word itself ("{base}")', "")
        return None
    forms = {f"{base}{suffix}" for suffix in INFLECTIONS}
    for token in norm(dfn).split():
        if token in forms:
            shown = "itself" if token == base else f'the form "{token}"'
            return (f'the definition uses the vocabulary word {shown} ("{base}")', "")
    return None


def noun_defined_as_verb(w: str, pos: str, dfn: str, sent: str):
    """Tagged noun, but the definition opens with an infinitive."""
    if pos.strip().lower() == "noun" and dfn.strip().lower().startswith("to "):
        return ("tagged noun, but the definition is written as a verb", "verb")
    return None


def sentence_run_together(w: str, pos: str, dfn: str, sent: str):
    """A missing space after a full stop inside the story sentence."""
    m = re.search(r"([a-z][.!?])([A-Z])", sent or "")
    if not m:
        return None
    fixed = re.sub(r"([a-z][.!?])([A-Z])", r"\1 \2", sent)
    return (f'no space after the full stop in "{m.group(0)}"', fixed)


def missing_sentence(w: str, pos: str, dfn: str, sent: str):
    if (sent or "").strip():
        return None
    return ("no story sentence has been written for this word", "")


CHECKS = {
    "no_definition": (no_definition, "Other"),
    "word_in_definition": (word_in_definition, "Answer Given"),
    "noun_defined_as_verb": (noun_defined_as_verb, "Part of Speech"),
    "sentence_run_together": (sentence_run_together, "Story Sentence"),
    "missing_sentence": (missing_sentence, "Story Sentence"),
}
# Which field the finding points at, per check.
TARGET_FIELD = {
    "no_definition": "DEF", "word_in_definition": "DEF",
    "noun_defined_as_verb": "POS", "sentence_run_together": "SENT",
    "missing_sentence": "SENT",
}


def run(staged: dict, only: str | None = None) -> list[dict]:
    out = []
    for book, items in staged.items():
        for slot, it in enumerate(items, 1):
            w = it.get("W", "")
            pos, dfn, sent = it.get("POS", ""), it.get("DEF", ""), it.get("SENT", "")
            for name, (fn, err_type) in CHECKS.items():
                if only and name != only:
                    continue
                hit = fn(w, pos, dfn, sent)
                if not hit:
                    continue
                details, fix = hit
                out.append({
                    "check": name, "book_id": book, "slot": slot,
                    "target": f"{TARGET_FIELD[name]}{slot}", "type": err_type,
                    "details": details, "fix": fix,
                    "W": w, "POS": pos, "DEF": dfn, "SENT": sent,
                })
    return out


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--staged", type=Path, required=True)
    ap.add_argument("--check", choices=sorted(CHECKS))
    ap.add_argument("--books", default="", help="comma-separated ids to limit to")
    ap.add_argument("--out", type=Path)
    ap.add_argument("--show", type=int, default=0, help="print this many candidates")
    args = ap.parse_args()

    staged = json.loads(args.staged.read_text(encoding="utf-8"))
    if args.books:
        wanted = {b.strip() for b in args.books.split(",") if b.strip()}
        staged = {b: v for b, v in staged.items() if b in wanted}

    found = run(staged, args.check)
    by_check = collections.Counter(f["check"] for f in found)
    total = sum(len(v) for v in staged.values())
    print(f"{total} entries across {len(staged)} book(s)\n")
    for name in sorted(CHECKS):
        n = by_check.get(name, 0)
        books = len({f["book_id"] for f in found if f["check"] == name})
        print(f"{n:5}  {name:24} across {books:3} book(s)")
    print(f"\n{len(found)} candidate(s) -- every one needs review before it is recorded")

    for f in found[:args.show]:
        print(f"\n  [{f['check']}] {f['book_id']} {f['target']}")
        print(f"     W    : {f['W']!r}  [{f['POS']}]")
        print(f"     DEF  : {f['DEF']!r}")
        print(f"     SENT : {f['SENT'][:90]!r}")
        print(f"     -> {f['details']}")

    if args.out:
        args.out.write_text(json.dumps(found, ensure_ascii=False), encoding="utf-8")
        print(f"\nwrote {len(found)} candidate(s) to {args.out}")


if __name__ == "__main__":
    main()
