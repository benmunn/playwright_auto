# playwright_auto

Scrapes authored activity content out of the Reading Space admin site into an existing
Excel workbook. Implements the spec in [pseudo-code.md](pseudo-code.md).

For each book id it opens four pages and records what it finds into the matching sheet.
Three are activity-edit pages; `Vocab` is the book's own word list, because that list —
not the *Word Accuracy Check* or *Word Meaning Match* activity that each show a subset of
it — is the authoritative record, and the only place the story sentence exists:

| Sheet | Page | Columns filled |
|---|---|---|
| `OEC` | `/activities/{id}/open-ended-questions/edit` | `Q1..Q10` |
| `CC` | `/activities/{id}/context-clue/edit` | `Q1..Q10`, `A1..A10` |
| `Vocab` | `/books/{id}/words` | `W1..W20`, `POS1..POS20`, `DEF1..DEF20`, `SENT1..SENT20` |
| `TMC` | `/activities/{id}/text-multiple-choice/edit` | `Q1..Q10`, `AnsA1..AnsD10`, `Correct1..Correct10` |

Two OEC columns come from the book record rather than an activity page, and so have
their own script, [book_fields.py](book_fields.py):

| Sheet | Page | Columns filled |
|---|---|---|
| `OEC` | `/books/{id}/edit` | `Main_Character`, `Prev_Text` |

The student sees `Main_Character` and `Prev_Text` while answering the OEC questions, which
is why they belong beside the questions in the sheet.

```bash
uv run python book_fields.py --dry-run          # report, write nothing
uv run python book_fields.py                    # every id in the OEC sheet
uv run python book_fields.py --books 697,1204   # just these
```

Unlike `rs_scrape.py` it takes command-line arguments, and will add its
two columns to an OEC sheet that lacks them (inserted after `id`, so the QA skill's `Err`
block stays at the right-hand edge); pass `--no-add-columns` to require them up front.
Pacing, login and session reuse are imported from `rs_scrape.py`, so every script here
hits the server identically.

A third script compares two activities against each other rather than recording one,
[pw_compare.py](pw_compare.py). Picture-Word Accuracy Check and Picture-Word Match pick
their words from the same global word list, so a word in the second that is missing from
the first is a word the student meets but is never checked on:

| Page | Read |
|---|---|
| `/activities/{id}/picture-word-accuracy-check/edit` | the selected-words table |
| `/activities/{id}/picture-word-match/edit` | the selected-words table |

```bash
uv run python pw_compare.py --scrape          # fill data/pw_words.json, resumable
uv run python pw_compare.py --write           # build the sheet from the cache
```

The scrape and the write are separate steps because the scrape is long — two page loads
per book — and caching the raw word lists means an interrupted run resumes instead of
starting over, and the sheet can be rebuilt without touching the server again. `--write`
adds a `PWAC - PWM` sheet to each workbook, one row per extra word, plus a row for any
book that has one activity and not the other. Books with neither are silently skipped;
most Level-2 books have neither.

Books scraped in separate batches end up in separate workbooks.
[merge_batch.py](merge_batch.py) folds one into another:

```bash
uv run python merge_batch.py --into data/2plus_check.xlsx \
    --from data/2plus_check_batch2.xlsx --dry-run
```

Columns are matched by header name, and the run aborts rather than merging if a book id
appears in both files — that would mean the two are versions of the same rows and
appending would double them — or if the source has a column the target lacks. Appended
rows are filled lavender and a legend is written beside the header, so a reviewer can
see which books arrived from elsewhere and still want their eyes. Pass `--no-highlight`
to skip the colouring.

`highlight_changes.py --flag-books 53,65,71` paints those books' rows the same lavender
in the reports, taking precedence over the new/renumbered colours.

## The global word list

Vocabulary is not book-owned. A book's word list links to entries in one global list, so
the same entry is shared by however many books chose it — and the same spelling can exist
several times over as separate entries. `bark` the sound a dog makes and `bark` on a tree
are two entries, and only the id tells them apart.

That matters because the word, its part of speech and its three definitions belong to the
entry. Editing any of them changes every book linked to it at once.

[word_ids.py](word_ids.py) walks the whole list and matches each vocabulary entry in the
workbook to the id it came from:

```bash
uv run python word_ids.py --scrape          # fill data/global_words.json, resumable
uv run python word_ids.py --match           # report how the two sides line up
uv run python word_ids.py --match --write   # add WID1..WID25 to the Vocab sheet
```

Walking the list costs 284 page loads against roughly four thousand for one search per
word, it can be re-matched without touching the server again, and it makes the near
misses visible instead of silently resolving them.

Matching runs in tiers, strictest first, and each entry records the tier that resolved it
so the weak ones can be audited:

| Tier | Resolves on |
|---|---|
| 1 | word + definition, exactly |
| 2 | word + definition, with the part of speech breaking a tie |
| 3 | word + definition ignoring case and punctuation |
| 4 | word, where one definition is the start of the other |
| 5 | word alone, when the list holds only one entry with that spelling |

The tie-break in tier 2 is a lookup rather than a guess: the workbook's part of speech was
scraped from the same record, and the list holds 522 word-and-definition pairs twice, all
but 15 of them filed under two different parts of speech. An entry that stays ambiguous
gets no id at all and is written to `data/word_id_unresolved.json`, because an arbitrary
id here is exactly the confusion the ids exist to remove.

The scrape also records the **Korean and Vietnamese definitions**, which live on the same
list rows. They are read here rather than from the per-word edit form for the same reason
as everything else: 284 pages against several thousand.

## Error reports

[make_reports.py](make_reports.py) turns a QA'd workbook into three views of the same
findings, plus a consistency check that all three hold identical rows:

```bash
uv run python make_reports.py --workbook data/2plus_check.xlsx     --out-prefix references/0820_error-reports
```

| View | Shape |
|---|---|
| `_long` | one row per finding, in the RnD x TFT feedback layout |
| `_recurring` | grouped by recurring kind, each kind with its own sheet |
| `_by-type` | grouped by the type each was logged under |

The by-type workbook carries two extra sheets that are organised by word rather than by
book, because a fix to a shared entry is one edit no matter how many books wanted it:

- **Vocab Changes** — one row per global word entry, with `current_` and `fixed_` columns
  for the word, part of speech, English definition and both translations. A blank `fixed_`
  cell means that field needs no change.
- **Vocab Changes - Split** — entries whose books mean genuinely different things by the
  same word. Ids are suffixed `A`, `B`, `C`, one row per sense. These cannot be fixed by
  editing the shared entry: it has to be split and each book relinked. Rows with no
  suggested fix are senses nobody flagged that the other row's fix would break.

Which entries split is a judgment about meaning, not a rule about text — two QA passes
routinely word the same fix differently — so the decisions are recorded by hand in
[report/vocab_changes.py](report/vocab_changes.py) alongside the reading that produced
them. Translation corrections live the same way, in
[report/translations.py](report/translations.py).

## Applying the fixes

[word_edit.py](word_edit.py) is the only script here that writes. Everything else reads,
where a mistake costs a re-run; a mistake here overwrites the live word list.

```bash
uv run python word_edit.py                     # dry run: says what it would change
uv run python word_edit.py --apply --limit 5   # actually save those five
uv run python word_edit.py --apply             # the whole sheet
```

It reads the **Vocab Changes** sheet and edits `/words/{id}/edit`. The split sheet is
refused outright, and so is any id that appears on it: those rows need new entries, and
editing the shared one in place is the exact mistake that sheet exists to prevent.

| Guard | What it prevents |
|---|---|
| Dry run is the default | Saving something nobody has read yet |
| The form is checked first | Writing to an entry that has been edited since the report was generated — the fix was written against text that is gone |
| Filled fields are read back | A fill that lands in the box before the page wires up its handlers, which registers nowhere and would otherwise save half an edit |
| The save button must enable | The form only allows a save once something really changed |
| Every word is journalled | No record of what was altered; `data/word_edits.jsonl` holds each field's before and after, and makes a run resumable |
| Saves are verified | A save that silently did not take |

Saving is heavier than reading, so it waits 5s between words on top of the usual
throttle, and stops after three consecutive failures like every other script here.

The image, the audio, and any translation the sheet has no replacement for are never
touched.

## Reviewing the vocabulary

[trans_dump.py](trans_dump.py) builds the review set for the Korean and Vietnamese
definitions: every global entry the books use, judged against the definition and part of
speech it will have *after* the English fixes are applied rather than the ones it has
now, since reviewing against text that is about to change approves translations that go
wrong the moment they are written.

```bash
uv run python trans_dump.py stats      # sizes, and defects a rule can see
uv run python trans_dump.py batch 3    # print one batch to read
```

Batches come riskiest-first: entries whose English changes *meaning* first, since their
translations are about to describe a word that no longer exists at that id; then entries
whose English is merely reworded; then the rest. Decisions go to
`data/translation_fixes.json`, written per entry so an interrupted review keeps what it
has already decided.

## Rewriting the open-ended questions

`oeq_edit.py` applies the OEC findings to the activity pages, one book per visit, both
questions in a single Update. It covers every OEQ subtype -- spelling, capitalization,
spacing, grammar, punctuation, awkward phrasing, unclear, too hard and requires-reading
-- because they all come to the same operation: replace the text and drop the recording
that reads the old wording.

```bash
uv run python oeq_edit.py            # dry run over both batches
uv run python oeq_edit.py --apply    # save
```

Three things about this page are worth knowing before changing the script.

**There are two buttons labelled "Remove" per question.** The one beside the `<h3>`
deletes the whole question; the one beside "Change Audio" drops the recording. Only the
second is ever clicked, and it is found by asking for the Remove that shares a parent
with "Change Audio". Never by position -- a question whose audio is already gone has
only the destructive one left. The questions are counted before and after editing as a
backstop, and a book that lost one is abandoned unsaved.

**Three findings carry an instruction where the replacement belongs** ("Replace it with
a question a student can answer..."). Typed in, that sentence becomes the question on
screen, so `usable()` drops anything that does not read as a question.

**The audio removal is blocked server-side.** Asked to save a question with no audio,
the backend tries to re-record it and answers `ELEVENLABS_API_KEY is not set` -- HTTP
200, failure in the body, nothing shown on the page. It rejects the whole mutation, so
the text cannot be saved in the same visit that clears the recording. `--keep-audio`
saves the text alone and works; the full run is held until the key is configured.

Of the 96 books with a flagged question, a dry run currently reports 68 to change, 16
already carrying the suggested text, and 12 rewritten on the site into something other
than what we proposed -- those last are left alone, since the live wording reads well.

## Re-scraping after the fixes have been applied

`word_edit.py` writes to the global word list, so once it has run the workbook and the
live list have diverged: 859 definitions currently differ. Two things follow, and
neither is obvious from the scraper on its own.

**A re-scrape will not reconcile them.** `rs_scrape.py` only ever writes into empty
cells, and `SKIP_ALREADY_RECORDED` skips a book whose values are already there. Running
it again over a populated workbook changes nothing. To pull the corrected text down, the
Vocab data cells have to be cleared first -- and then the `Err[#]` findings beside them
describe text that is gone, so they need re-QAing rather than re-reading.

**Clearing the word cells means re-matching the ids.** `WID[#]` is written per slot by
`word_ids.py`, and a re-scrape can put a different word in a given slot. An id left
behind from the previous occupant is worse than a blank one, because everything
downstream treats that column as naming the entry. `word_ids.py --match --write` now
clears the cell of any slot it cannot resolve, so a stale id cannot survive; run it
again after any re-scrape of the Vocab sheet.

The reports themselves stay honest in the meantime: `make_reports.py` replays
`data/word_edits.jsonl` over the current-value columns, so they describe the list as it
now stands rather than as it was scraped.

## Tests

```bash
uv run python tests/test_logic.py
```

No browser, no network. What is covered is the reasoning the runs depend on -- which
rows are held back, which edits the database would refuse, whether a word still has work
outstanding -- because that is where the faults have actually been. Every test stands
for a bug that reached the live word list or came within one run of doing so.

## Setup

```bash
uv sync && uv run playwright install chromium
```

Generate a blank workbook with every column the scraper expects:

```bash
uv run make_template.py
```

That writes `data/activities_template.xlsx` — headers only. Add `--ids 697 --out data/test_697.xlsx` to seed book ids into the id column.

Copy `.env.example` to `.env` and fill it in:

```
RS_USERNAME=your.username
RS_PASSWORD=your.password
RS_WORKBOOK=data/activities.xlsx
```

`RS_WORKBOOK` is the workbook to write into — it must already exist with its sheets and
headers. A real environment variable of the same name takes precedence over `.env`, which
is handy for a one-off run against a different file:

```bash
RS_WORKBOOK=data/test_697.xlsx uv run main.py
```

`.env` is gitignored. The login session is cached to `.auth/state.json` so later runs
skip the login form — delete that file to force a fresh login.

## Running

The workbook path and credentials come from `.env`. Everything else lives in the
constants block at the top of [rs_scrape.py](rs_scrape.py) — there are no command-line
arguments.

```bash
uv run main.py
```

Useful constants while getting started:

- `DRY_RUN = True` — report every intended write, save nothing.
- `HEADLESS = False` — watch the browser work.
- `BOOK_IDS = ["101"]` — one book instead of every id in the sheet (`None` = all).
- `ACTIVITIES_TO_RUN = ("OEC",)` — one activity at a time.

## Going easy on the server

The admin panel is a low-resource back-office tool, so the script is deliberately slow
and defensive. All of this is tunable in the politeness block in
[rs_scrape.py](rs_scrape.py):

| Setting | Default | Effect |
|---|---|---|
| `REQUEST_DELAY_S` | `3.0` | Minimum seconds between page loads. Every navigation goes through one throttle, so nothing bypasses it. |
| `RETRY_BACKOFF_S` | `20.0` | Wait after a failed page load, before the single retry. |
| `MAX_RETRIES` | `1` | Retries per page. `0` disables them. |
| `MAX_CONSECUTIVE_FAILURES` | `3` | Abort the whole run after this many books fail back-to-back, rather than keep pushing a struggling server. |
| `BLOCKED_RESOURCES` | images, media, fonts | Never fetched — a page load is many requests, and these aren't needed to read text. Scripts and stylesheets are still loaded, since the panel is a React app. |

The run is strictly sequential: one page at a time, never concurrent. Expect roughly
`books × activities × REQUEST_DELAY_S` seconds — 100 books across all four activities at
the default pacing is about 20 minutes. Raise `REQUEST_DELAY_S` if the server still feels
stressed; nothing breaks if you set it to 10.

If a run does abort, progress is already saved and re-running picks up only what's
missing.

## Behaviour worth knowing

- **The workbook must already exist**, with its sheets and header row. The script locates
  columns by header name and never creates or restructures anything. A missing sheet or
  missing `Q1`-style column is a hard error rather than a guess.
- **Re-running is safe.** Values already recorded for a book are not written a second
  time, so a re-run after a crash picks up only what is missing. Set
  `SKIP_ALREADY_RECORDED = False` for plain append-at-first-empty-slot behaviour.
- **Existing cells are never overwritten.** New data goes into the first free slot.
- **Nothing is truncated silently.** If a book has more items than free slots, each
  dropped value is printed.
- Progress is saved after every book, so an interrupted run keeps its finished work.
- Pages that yield no data are logged and screenshotted to `debug/`.
- The workbook must not be open in Excel while the script runs — Windows locks the file.

## A note on selectors

Most elements are found structurally (placeholder text, table cell position, input
ordering). Two have no stable handle and are matched on their Tailwind classes — the CC
answer pills and the TMC question cards, both defined near the top of
[rs_scrape.py](rs_scrape.py). If an activity suddenly extracts nothing after an admin-UI
restyle, those two constants are the first place to look.
