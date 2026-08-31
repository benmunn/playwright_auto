---
name: two-level-qa
description: QA a Reading Star activity spreadsheet — check CC, Vocab, OEC and TMC sheets for grammar, punctuation, capitalization, spacing, awkward phrasing and activity-specific faults, and record findings in Err[#] columns beside each row. Use when the user asks to "run 2-level QA", "QA the CC/Vocab/OEC/TMC sheet", "check the activity spreadsheet for errors", "error-check this workbook", or "make the error report" for a scraped activity workbook.
---

# Two-Level QA

Judge every question, sentence, word and definition in an activity workbook against its
activity's error taxonomy, and record findings in `Err[#]_target / _type / _details /
_fix` columns beside the source row.

**You do the judging. `qa_sheet.py` does everything else.** Never edit the workbook by
hand — the script owns backups, header creation, free-slot seeking, resume state and
validation. Editing around it will corrupt the resume state and silently lose findings.

Full taxonomies: [reference/error-types.md](reference/error-types.md). Read it before
judging anything; it also documents four places where it deliberately corrects the source
instructions.

## 1. Ask first — never guess these three

1. **Which workbook.** Ask for the path. Do not assume a file, even if one is obviously
   present in `data/`.
2. **Target reading level.** Required. "Too Hard" and the awkward-phrasing bar are
   meaningless without it, and it changes run to run. Ask plainly: *what reading level
   are these books for?* (A recent workbook was US 2nd-grade native equivalent — an
   example, not a default.)
3. **Scope.** Which sheets, and all rows or a subset.

Then run `status` and show the user what the job actually is before starting:

```bash
uv run python .claude/skills/two-level-qa/qa_sheet.py status <workbook>
```

**If OEC is in scope, check it has `Main_Character` and `Prev_Text` first.** The student
sees both while answering, so they are what `Requires Reading` is judged against — and
without them that type cannot be judged at all. If the columns are missing, the sheet
predates them; offer to populate them before starting:

```bash
uv run python book_fields.py --workbook <workbook>
```

That adds both columns if absent and fills them from each book's `/books/<id>/edit` page.
It is a ~13-minute run for 250 books. Judging OEC without it is possible, but say so in
the report rather than letting `Requires Reading` look checked.

**If Vocab is in scope, check it has `SENT1`.** The story sentence is what makes
`Wrong Sense` and `Part of Speech` checkable instead of inferred from the book's other
vocabulary. If the column is missing, the sheet predates it and needs re-scraping with
`rs_scrape.py`; without it, say in the report that those two types were judged without the
sentence rather than letting them look verified.

Vocab rows are wide — a book's word list runs to about twenty entries, not the eight the
Word Accuracy Check activity used to expose — so keep Vocab batches small.

**Vocab entries are shared between books, and their translations are not in this sheet.**
The word, part of speech and definition belong to a global word entry that any number of
books link to; only the story sentence is this book's. The same entry also carries a
Korean and a Vietnamese definition that the workbook does not hold. Judge what is in the
row and leave the translations alone — they are reviewed once per entry on the
**Vocab Changes** sheet of the by-type report, not once per book that happens to use the
word. See `reference/error-types.md`.

A full pass over a 253-book workbook is several thousand judgment calls. Tell the user the
scale up front and agree the scope rather than silently starting a job that runs for hours.

## 2. Prepare

```bash
uv run python .claude/skills/two-level-qa/qa_sheet.py prepare <workbook> --sheets CC Vocab
```

Takes a timestamped backup, then appends `Err1..Err5` sets plus `QA_pass1` / `QA_pass2`
after the last used column. Safe to re-run: it never duplicates headers.

`QA_pass1` / `QA_pass2` are two columns beyond the 20 the instructions specify. They hold
the finding count for that pass (`0` for a clean row) and are what makes the run
resumable — without them a clean row is indistinguishable from an unjudged one.

## 3. Pass 1 — find

Loop until `batch` returns no rows. One sheet at a time.

```bash
uv run python .claude/skills/two-level-qa/qa_sheet.py batch <workbook> \
    --sheet CC --size 15 --out batch.json
```

Read `batch.json`, judge every item, then write findings back:

```bash
uv run python .claude/skills/two-level-qa/qa_sheet.py write <workbook> \
    --sheet CC --findings findings.json
```

`findings.json` — **every row from the batch must appear**, including clean ones with an
empty `findings` list. That is what marks them judged:

```json
{"rows": [
  {"row": 5, "findings": [
    {"target": "Q2", "type": "Awkward Phrasing",
     "details": "unnatural phrase: \"they enjoy together\"",
     "fix": "They enjoy playing together."}
  ]},
  {"row": 6, "findings": []}
]}
```

Batch sizing: start at **15 rows** for CC and TMC, **25** for OEC and **8** for Vocab —
twenty words a row adds up fast — and drop it if a batch feels crowded. Judgment quality
matters more than throughput: a rushed batch produces findings that pass 2 has to retract.

`write` validates the whole batch before touching the workbook: unknown `type`, a
`target` that isn't a real column, or empty `details` rejects the batch entirely, so
nothing half-applies. It extends to `Err6`, `Err7`… automatically when a row needs more
than five. Findings are never truncated.

## 4. Pass 2 — proof

Only after pass 1 covers the whole sheet. Same loop with `--pass 2`; each row now also
carries its `recorded` findings, each with a `slot` number.

```bash
uv run python .claude/skills/two-level-qa/qa_sheet.py batch <workbook> \
    --sheet CC --size 15 --pass 2 --out proof.json
```

Re-read each row against the source text and do two things:

- **Retract false positives.** Reference findings by their `slot`:

  ```bash
  uv run python .claude/skills/two-level-qa/qa_sheet.py retract <workbook> \
      --sheet CC --retractions retractions.json
  ```
  ```json
  {"retractions": [{"row": 5, "slot": 2}]}
  ```
  Remaining findings close up so slots stay contiguous.

- **Add what pass 1 missed**, with `write --pass 2`.

Order matters: retract first, then write, or the slot numbers you were given shift under
you.

Report every retraction to the user with its reason. A withdrawn finding is information,
not an embarrassment to hide.

## 5. Report

Per sheet: rows judged, findings by type, rows that needed `Err6+`, retractions with
reasons, and anything you could not check. Two `null` columns have to be called out by
name rather than passed over:

- `Correct[#]` null → TMC answer-key correctness was not verified; the sheet predates the
  answer-key scrape.
- `Main_Character` / `Prev_Text` null → OEC `Requires Reading` was not verified; the sheet
  predates `book_fields.py`.
- `SENT[#]` null → Vocab `Wrong Sense` and `Part of Speech` were judged without the story
  sentence; the sheet predates the word-list scrape.

A workbook may also carry `LR` / `LRA` sheets. Those are **not** part of this process:
they are flat, one-row-per-error records of a human listening pass over Listen & Read,
with no scraped content to judge, so `qa_sheet.py` ignores them and neither pass touches
them. `make_reports.py` picks them up separately, so leave them alone rather than trying
to fold them into an `Err[#]` sheet.

## Judgment discipline

- **Only flag what you would actually change.** The bar is "clearly wrong", not
  "could be phrased differently". A noisy report is worse than a short one.
- **Never invent a `fix`.** If you cannot write a confident paste-ready replacement, the
  finding probably isn't solid enough to record.
- **Write `details` for the person who will fix it.** They are a proficient non-native
  English speaker, not a linguist. Everyday classroom grammar words are fine — *singular*,
  *plural*, *article*, *clause*, *uncountable*, *comma*. Linguistics vocabulary is not:
  write "the vocabulary word" not *headword*, "the base of the word" not *stem*, "nothing
  to refer to" not *no referent*, "unnecessary comma" not *spurious comma*, "a short
  explanation" not *gloss*, "an -ing form" not *gerund*, "added at the end" not
  *appended*. If a term would send the reader to a dictionary, it costs more than it saves.
- **Use the row's context.** CC's `word_bank` is supplied because "Lack of Context" and
  "Answer Given" cannot be judged from one sentence alone. OEC's `Main_Character` and
  `Prev_Text` are supplied for the same reason: they are everything the student can see
  before reading, so a question is only "Requires Reading" if those two don't cover it.
  Both are also valid `target` values — a typo in the preview text is a finding.
  Consistency rules are judged within the row, never across the sheet.
- **`null` is not `""`.** In a batch payload, `null` means that column does not exist in
  the sheet; `""` means the cell is empty. Never report a check you could not perform.
- **Stop and ask** if a sheet's content doesn't match its expected shape rather than
  forcing findings into the wrong taxonomy.

## Failure modes

| Symptom | Cause |
|---|---|
| `is open in Excel` | Close the workbook. openpyxl cannot write to a locked file. |
| `has no QA_pass1 column` | `prepare` wasn't run for that sheet. |
| `Refusing to write` | A finding has a bad type, target or empty details. Fix and resend — nothing was written. |
| `batch` returns no rows | That sheet is finished for that pass. |
| OEC rows have `"Main_Character": null` | The sheet predates `book_fields.py`. Run it, or report `Requires Reading` as unchecked. |
