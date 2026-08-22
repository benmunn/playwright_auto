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
