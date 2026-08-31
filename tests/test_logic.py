"""Tests for the decision logic behind the reports and the edit run.

Nothing here opens a browser or touches the network. What is covered is the reasoning a
run depends on -- which rows are held back, which edits the database would refuse,
whether a word still has work outstanding -- because that is where the faults have
actually been. Each test below stands for a bug that reached the live word list or came
within one run of doing so.

    uv run python tests/test_logic.py
"""

from __future__ import annotations

import sys
import tempfile
import traceback
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import Workbook  # noqa: E402

import word_edit as we  # noqa: E402
import word_ids as W  # noqa: E402
from report.plain import plain  # noqa: E402

CASES: list = []


def case(fn):
    CASES.append(fn)
    return fn


def sheet(rows: list[dict], *, columns: list[str] | None = None,
          split: list[str] | None = None) -> Path:
    """A minimal Vocab Changes workbook, written where load_rows can read it."""
    cols = columns or ["Book", "Book ID", "word_id", "current_word", "fixed_word",
                       "current_pos", "fixed_pos", "current_def", "fixed_def",
                       "current_kor", "fixed_kor", "current_vie", "fixed_vie",
                       "Error Types", "Details"]
    wb = Workbook()
    ws = wb.active
    ws.title = we.SHEET
    for c, h in enumerate(cols, 1):
        ws.cell(we.HEADER_ROW, c, h)
    for i, row in enumerate(rows):
        for c, h in enumerate(cols, 1):
            ws.cell(we.HEADER_ROW + 1 + i, c, row.get(h, ""))
    sp = wb.create_sheet(we.SPLIT_SHEET)
    for c, h in enumerate(["Book", "Book ID", "word_id"], 1):
        sp.cell(we.HEADER_ROW, c, h)
    for i, wid in enumerate(split or []):
        sp.cell(we.HEADER_ROW + 1 + i, 3, wid)
    path = Path(tempfile.mkdtemp()) / "report.xlsx"
    wb.save(path)
    return path


# --------------------------------------------------------------------------------------
# What the database will and will not accept
# --------------------------------------------------------------------------------------


@case
def key_is_case_sensitive():
    """A unique constraint in Postgres compares case. Lowercasing first made the check
    refuse seven edits the database would have taken."""
    assert we._key("bonnet(s)", "noun", "a hat") != we._key("Bonnet(s)", "noun", "a hat")
    assert we._key("bonnet(s)", "noun", "a hat") == we._key("bonnet(s)", "Noun", "a hat")


@case
def collision_ignores_the_row_itself():
    """An entry always matches its own key; only another id is a collision."""
    row = {"word_id": "10", "fixed_word": "", "current_word": "pea",
           "fixed_pos": "", "current_pos": "noun",
           "fixed_definition": "a vegetable", "current_definition": "pea"}
    key = we._key("pea", "noun", "a vegetable")
    assert we.collides(row, {key: ["10"]}) == []
    assert we.collides(row, {key: ["10", "99"]}) == ["99"]


# --------------------------------------------------------------------------------------
# Which rows a run may touch
# --------------------------------------------------------------------------------------


@case
def wrong_sense_is_held_back():
    path = sheet([{"word_id": "1", "current_def": "old", "fixed_def": "new",
                   "Error Types": "Wrong Sense"},
                  {"word_id": "2", "current_def": "old", "fixed_def": "new",
                   "Error Types": "Grammar"}])
    got = [r["word_id"] for r in we.load_rows(path, we.SHEET)]
    assert got == ["2"], got


@case
def cleared_ids_come_through_the_hold():
    wid = sorted(we.WRONG_SENSE_CLEARED)[0]
    path = sheet([{"word_id": wid, "current_def": "old", "fixed_def": "new",
                   "Error Types": "Wrong Sense"}])
    assert [r["word_id"] for r in we.load_rows(path, we.SHEET)] == [wid]


@case
def missing_error_types_column_is_refused():
    """The hold reads that column, and a lookup for an absent column returns empty. A
    report without it would once have written every entry the hold exists to keep back."""
    cols = ["word_id", "current_word", "fixed_word", "current_pos", "fixed_pos",
            "current_def", "fixed_def"]
    path = sheet([{"word_id": "1", "fixed_def": "new"}], columns=cols)
    try:
        we.load_rows(path, we.SHEET)
    except SystemExit as exc:
        assert "Error Types" in str(exc), exc
    else:
        raise AssertionError("a report with no Error Types column was accepted")


@case
def split_entries_are_never_edited_in_place():
    path = sheet([{"word_id": "7", "current_def": "old", "fixed_def": "new",
                   "Error Types": "Wrong Entry"}], split=["7A", "7B"])
    assert we.load_rows(path, we.SHEET) == []


@case
def rows_with_nothing_to_change_are_dropped():
    path = sheet([{"word_id": "1", "current_def": "same", "Error Types": "Grammar"}])
    assert we.load_rows(path, we.SHEET) == []


# --------------------------------------------------------------------------------------
# Resuming
# --------------------------------------------------------------------------------------


@case
def a_word_is_done_only_when_every_pending_field_is_written():
    """Treating a word as finished because it was saved once left three entries with an
    empty Vietnamese box that no later run would ever pick up."""
    row = {"word_id": "5", "fixed_word": "", "fixed_pos": "",
           "fixed_definition": "d", "fixed_kor": "", "fixed_vie": "v"}
    assert we.already_done(row, {"5": {"definition": "d", "vie": "v"}})
    assert not we.already_done(row, {"5": {"definition": "d"}})
    assert not we.already_done(row, {"5": {"definition": "other", "vie": "v"}})
    assert not we.already_done(row, {})


# --------------------------------------------------------------------------------------
# Refusing to write over an entry that has moved on
# --------------------------------------------------------------------------------------


@case
def identity_check_compares_what_the_report_recorded():
    row = {"word_id": "3", "current_word": "pea", "current_pos": "noun",
           "current_definition": "a vegetable", "current_kor": "", "current_vie": "",
           "fixed_word": "", "fixed_pos": "", "fixed_definition": "x",
           "fixed_kor": "", "fixed_vie": ""}
    same = {"word_id": "3", "word": "pea", "pos": "noun",
            "definition": "a vegetable", "kor": "", "vie": ""}
    assert we.check_matches(row, same) is None
    assert we.check_matches(row, {**same, "word_id": "4"})
    assert we.check_matches(row, {**same, "definition": "moved on"})


@case
def a_translation_is_only_checked_when_it_is_being_replaced():
    """The translations were read in a separate pass, so a Korean edited by someone else
    must not block an unrelated English fix."""
    base = {"word_id": "3", "current_word": "pea", "current_pos": "noun",
            "current_definition": "a vegetable", "current_kor": "A", "current_vie": "",
            "fixed_word": "", "fixed_pos": "", "fixed_definition": "x", "fixed_vie": ""}
    live = {"word_id": "3", "word": "pea", "pos": "noun",
            "definition": "a vegetable", "kor": "B", "vie": ""}
    assert we.check_matches({**base, "fixed_kor": ""}, live) is None
    assert we.check_matches({**base, "fixed_kor": "C"}, live)


# --------------------------------------------------------------------------------------
# Matching a vocabulary entry to its global id
# --------------------------------------------------------------------------------------


def index(*words):
    return W.build_index([{"id": i, "word": w, "pos": p, "definition": d}
                          for i, w, p, d in words])


@case
def exact_match_wins():
    tiers = index(("1", "root", "noun", "the part of a plant"))
    assert W.resolve("root", "the part of a plant", tiers)[0] == "1"


@case
def part_of_speech_breaks_a_tie():
    """The list holds 522 word-and-definition pairs twice, nearly all of them filed under
    two different parts of speech; the workbook's own tag settles which is meant."""
    tiers = index(("1", "root", "verb", "the part of a plant"),
                  ("2", "root", "noun", "the part of a plant"))
    assert W.resolve("root", "the part of a plant", tiers, "noun")[0] == "2"
    assert W.resolve("root", "the part of a plant", tiers)[0] is None


@case
def a_true_duplicate_resolves_to_nothing():
    """Same word, part of speech and definition twice over: an arbitrary id here is the
    confusion the ids exist to remove."""
    tiers = index(("1", "be afraid of", "phrase", "to feel fear"),
                  ("2", "be afraid of", "phrase", "to feel fear"))
    assert W.resolve("be afraid of", "to feel fear", tiers, "phrase")[0] is None


@case
def word_alone_never_picks_a_sense():
    """With the definition already failed to match, word plus part of speech would choose
    a meaning rather than confirm one."""
    tiers = index(("1", "bark", "noun", "a tree's covering"),
                  ("2", "bark", "noun", "a dog's sound"))
    assert W.resolve("bark", "something else entirely", tiers, "noun")[0] is None


# --------------------------------------------------------------------------------------
# Rewriting finding text into plain language
# --------------------------------------------------------------------------------------


@case
def quoted_text_is_left_exactly_as_written():
    """Rewriting inside quotation marks corrupts the very text a fixer must search for:
    "a large circular structure" once became "a large explains itself in a circle"."""
    out = plain('above level: "a large circular structure"')
    assert '"a large circular structure"' in out, out


if __name__ == "__main__":
    failed = []
    for fn in CASES:
        try:
            fn()
            print(f"  ok   {fn.__name__}")
        except Exception:                                          # noqa: BLE001
            failed.append(fn.__name__)
            print(f"  FAIL {fn.__name__}")
            traceback.print_exc()
    print(f"\n{len(CASES) - len(failed)}/{len(CASES)} passed")
    sys.exit(1 if failed else 0)
