# Error types by activity

Source of truth: `data/Instructions - 2level QA.md`. Where this file departs from that
document it is correcting a copy-paste fault — each departure is marked **[corrected]**
with the reasoning, so nobody silently restores the original.

`qa_sheet.py write` rejects any `type` outside the sheet's list, so these are closed sets.

---

## Shared rules

**Too Hard applies to every sheet.** The source instructions list it only under OEC,
but above-level vocabulary appears in all four activities. Judge it against the reading
level given at the start of the run, and start `details` with `above level:`.

**Answer Given is scoped to the sentence or entry itself**, never across a row. A word
bank answer appearing in a *different* CC sentence is out of scope, by decision — the
instructions say "in the sentence" and that reading stands.

**Severity bar.** Slightly awkward is fine. Flag what is clearly wrong. A finding should
be something an editor would actually change, not a stylistic preference.

**`Err[#]_target`** — the column header the error is in (`Q2`, `DEF3`, `AnsB1`).
When the problem is the *relationship* between two cells, target the primary one:

| Sheet | Relationship error targets |
|---|---|
| Vocab | `W[#]`, not `DEF[#]` |
| TMC | `Q[#]`, not `AnsA-D[#]` |

**`Err[#]_details`** — a very short explanation. Example: `unnatural phrase: "they enjoy
together"`.
**[corrected]** Three of the four sections say "write the type of the error" here. That
is copy-pasted from the `_type` bullet; the worked example in every section is an
explanation, and the OEC section states it correctly.

**`Err[#]_fix`** — a paste-ready replacement, nothing else. No explanation, no "consider
changing X to Y", no quotation marks around it unless they belong in the cell. Someone
should be able to copy the cell straight into the admin panel.

**`Other`** — a real error that fits no listed type. Never a hedge for "I'm not sure
this is an error". If you are not confident, do not record a finding.

**One row is one book.** Every consistency rule ("standard across each book") is judged
*within a single row*, never across the sheet. Book 101 starting definitions with
capitals and book 102 starting lowercase is not an error; one book doing both is.

---

## CC — Context Clue

Students see the sentence with the `A[#]` word blanked out and pick from the word bank of
all non-empty `A1..A10` values in that row. The batch payload gives you that word bank.

| Type | Flag when |
|---|---|
| Grammar | Improper tense, subject-verb disagreement, dangling modifier |
| Punctuation | Extra or missing. Sentences need terminal punctuation; non-sentences must not have it |
| Capitalization | Sentences start capitalised. Non-sentences: consistent within the row |
| Spacing | Leading, trailing or double spaces |
| Awkward Phrasing | Not idiomatic, natural English |
| Lack of Context | Not enough context to pick this answer over the other word-bank options |
| Answer Given | The sentence gives the answer away |
| Too Hard | Vocabulary, grammar or concepts beyond the target reading level |
| Other | — |

**Answer Given** is the subtle one. Check for the answer's *stem* surviving elsewhere in
the sentence after the blank is cut. The doc's example: `"In the northern regions, the
regions up north, it is very cold."` with answer `northern` — blanking `northern` still
leaves `north`.

**Lack of Context** must be judged against the *whole* word bank. A sentence is only
faulty if more than one bank word would genuinely fit.

---

## Vocab

The book's vocabulary set: `W[#]` (the word), `POS[#]` (its part of speech), `DEF[#]` (its
definition) and `SENT[#]` (the sentence showing how this book uses it). Students match
`W[#]` to `DEF[#]`. `POS[#]` and `SENT[#]` are not shown to students but are both checked.

**Where this comes from, and what to call it.** The sheet is scraped from the book's own
word list at `/books/<id>/words`. Two activities each show a subset of that list: Word
Accuracy Check shows the words alone, Word Meaning Match shows words with definitions. So
a definition or part-of-speech error can only reach a student through **Word Meaning
Match**, and that is the activity to name when reporting one — never Word Accuracy Check.

Because the list is the source, it holds words no activity uses: about twenty entries per
book against the eight the activity exposes. Judge every word in the row.

**The entry is shared, and its translations are not in this sheet.** `W[#]`, `POS[#]` and
`DEF[#]` belong to an entry in the global word list, which any number of books link to;
only `SENT[#]` belongs to this book. The same entry also carries a Korean and a
Vietnamese definition, which the workbook does not hold at all.

Do not try to QA the translations from here. A book-by-book pass would judge one shared
entry once per book that uses it — the duplication the global word id exists to remove —
and the workbook has no column to record the finding in. Translations are reviewed once
per entry, on the **Vocab Changes** sheet of the by-type report, against the corrected
English definition. Log what belongs to this book: the word, its part of speech, its
definition and its sentence.

### The story sentence settles Wrong Sense and Part of Speech

`SENT[#]` is the evidence, not a hint. It shows the word as this book actually uses it, so
judge `POS[#]` and `DEF[#]` **against the sentence** before reasoning from the book's other
vocabulary:

- The sentence uses the word as a different word class than `POS[#]` claims → **Part of
  Speech**.
- The definition describes a different meaning than the sentence shows → **Wrong Sense**.
- Both wrong → **two findings on the same word**. That is expected and permitted.

`null` means the sheet predates the sentence scrape; `""` means no sentence has been
authored for that word yet, not that the scrape missed it. In either case say so rather
than reporting these two types as checked.

The sentence is evidence about *this book's* usage, not a definition. A definition that
covers a wider meaning than the sentence shows is fine — flag only when the sentence and
the definition are about genuinely different meanings.

| Type | Flag when |
|---|---|
| Grammar | Improper tense, agreement, dangling modifier in the definition |
| Punctuation | Extra or missing; definitions written with sentence-like punctuation |
| **Part of Speech** | The POS contradicts the word or its definition |
| **Wrong Sense** | The definition is correct English, but for a different meaning of the word than the one this book teaches |
| **Spelling** | The word tile itself is misspelled — `aike` for *alike*, `hittern` for *hitter* |
| **Wrong Entry** | The whole entry belongs to a different word or sense than the book uses, not just its tag |
| Capitalization | Definitions consistent within the row. Words capitalised only if proper nouns or full sentences |
| Spacing | Leading, trailing or double spaces |
| Awkward Phrasing | Definition is not idiomatic, natural English |
| Lack of Context | Definition too thin to identify the word, or to tell it from the other options in the row |
| Answer Given | Definition contains the word or another form of it |
| Too Hard | Vocabulary, grammar or concepts beyond the target reading level |
| Other | — |

**Wrong Entry vs Part of Speech vs Wrong Sense.** All three say the entry does not match
the book, and they are fixed in different places. *Part of Speech* is only the tag being
wrong. *Wrong Sense* is the definition describing another meaning. *Wrong Entry* is both
at once — the tag and the definition belong to a different word spelled the same way, so
the whole entry has to be rewritten. `bark(ing)`, tagged verb and defined as a dog's
sound, in a book about trees is a Wrong Entry, not a Part of Speech slip.

**Wrong Sense vs Too Hard.** These look alike and are not. *Too Hard* means the
definition is about the right meaning but uses words above the reading level — the fix is
simpler wording. *Wrong Sense* means the definition is about the **wrong meaning
entirely**, and no amount of simplifying helps:

- `hemisphere` defined as *"either of the two halves of the brain"* in a book about
  continents → **Wrong Sense**. The fix is a new definition, not an easier one.
- `waste` defined as *"an area of land that cannot be used"* in a book about garbage →
  **Wrong Sense**.
- `constitute` used inside the definition of *make up* → **Too Hard**.

Where the word genuinely has several senses and the book's one is merely second, put the
book's sense first rather than deleting the others.

**[corrected] "Part of Speech"**, not "Inflection". The doc's body checks part of speech
while its Types line offers "Inflection"; the body wins. Word-form problems — a headword
given in the wrong form, or inconsistently with the row's other entries, such as
`stammer, (stammered)` — also file under Part of Speech.

**[corrected] Lack of Context** reads "Flag definitions **with** enough context" in the
doc. That is inverted: flag definitions **without** enough context.

**Answer Given** examples: word `northern` / def `in the north`; word `scary` / def
`tending to scare people`.

**Sense order.** Multiple senses in one definition are fine and often deliberate — these
words get reused in other books with a different sense. What is wrong is **leading with
the uncommon or irrelevant sense**. Flag `Too Hard` only when the first sense given is not
the one this book teaches, and write the `fix` as a **reordering** that keeps every sense:

> `matter` — "a subject or situation that you must think about or deal with; the physical
> substance that makes up all things" → put the physical-substance sense first.

Do **not** flag a definition merely for carrying an extra sense after a correct leading
one. `court` leading with the legal sense in a government book is right, even though a
sports sense follows.

Definitions should also not be verbose. Flag `Too Hard` for wording that is longer or
more abstract than the level needs, independently of sense order.

---

## OEC — Open-Ended Comprehension

Warm-up questions asked **before** the student reads the book. That framing drives two of
the types below.

### What the student can see

Two per-book fields appear on screen alongside the questions, so they are the whole of
what a student knows before answering:

| Column | Source | Content |
|---|---|---|
| `Main_Character` | book record | comma-separated character names, e.g. *Bill, Sandy, Anubis* |
| `Prev_Text` | book record | one-or-two-sentence teaser for the book |

Both are scraped by `book_fields.py`, not `rs_scrape.py` — they live on `/books/<id>/edit`,
not on the activity page. In a batch payload they sit at the top level of the row, beside
`items`.

`null` means the sheet has no such column (it predates `book_fields.py`); `""` means the
book genuinely left the field blank. **Never judge `Requires Reading` from a `null`** —
say in the report that it could not be checked, and have the user run `book_fields.py`.

Non-fiction books routinely have an empty `Main_Character`. That is normal and is not a
finding.

### Types

| Type | Flag when |
|---|---|
| Grammar | Improper tense, agreement, dangling modifier |
| Punctuation | Extra or missing |
| Capitalization | First letter capitalised; otherwise only proper nouns |
| Spacing | Leading, trailing or double spaces |
| Awkward Phrasing | Not idiomatic, natural English |
| Unclear | A student could not tell how they are meant to answer |
| Requires Reading | Needs information the student does not have yet — see below |
| Too Hard | Vocabulary, grammar, concepts, or world knowledge beyond the target level |
| Other | — |

`Main_Character` and `Prev_Text` are themselves valid `Err[#]_target` values and take the
same types as a question, minus `Requires Reading` (they are the source, so they cannot
presuppose it). A typo, a missing capital, or a preview text that gives away the ending
is a finding against that column.

**[corrected]** The doc's OEC Types line reads "Lack of Context, Answer Given" — copied
from CC, and neither appears anywhere in the OEC body. The body's own three types
(Unclear, Requires Reading, Too Hard) are the correct list.

### Requires Reading, judged against the preview

A question is `Requires Reading` when answering it needs something **neither the student's
own experience nor `Main_Character` / `Prev_Text` supplies**. The two fields are part of
the row's context, exactly as CC's word bank is: check them before flagging.

- *"If you were John, how would you feel about Jessica's plan?"* where neither name is in
  `Main_Character` and the preview never mentions a plan — flag.
- *"What are the two differences between reptiles and mammals mentioned in the story?"* —
  flag. "Mentioned in the story" presupposes the text outright, whatever the preview says.
- *"How do you think Bill feels about his new friend?"* where `Main_Character` lists
  *Bill, Sandy* and `Prev_Text` says Bill must *"tell his parents the truth about his new
  friend"* — **do not flag**. Every noun in the question is on screen, and the student is
  being asked to predict, which is the point of the activity.

The names matter both ways: a question naming a character who appears in **neither** field
is worth flagging even if it reads like a prediction prompt, because the student has no
way to know who that is.

**The referent test.** Most of these turn on one question: does a definite or comparative
phrase — *the* leak, *other* animals, *each* community, *the next* note — have something
in `Main_Character` or `Prev_Text` to point at?

- Preview says *"Your five senses help keep you safe. You can smell smoke, hear a warning,
  and see where you step."* → *"Can you think of **other ways** our senses keep us safe?"*
  is fine; the preview gave the ways that "other" contrasts with.
- Preview says only *"learn why soil is so important for life"* → *"Can you think of
  **other uses** for soil?"* still flags; no use was named, so "other" points at nothing.

Same word, opposite verdicts, decided entirely by the preview. Run this test before
reaching for the type.

**Wrong type, real fault.** A question can be answerable from the preview and still be
broken — *"If you were Nick, how would you feel when you discovered the treasure was not
gold but a secret recipe?"* hands the student the ending. That is `Other` (a spoiler), not
`Requires Reading`. When retiring a `Requires Reading` finding, check whether a different
type should replace it rather than leaving the row clean.

**Too Hard** is the one type that depends entirely on the target reading level you were
given at the start of the run. Judge against that level, not against adult fluency.

---

## TMC — Text Multiple Choice

One `Q[#]` with four options `AnsA[#]`–`AnsD[#]`. `Correct[#]` holds the letter of the
keyed answer.

### Targeting all four options at once

When the same fault runs through every option of one question — periods on all four, a
whole option set pasted from the wrong book — write the target as **`AnsAll[#]`** rather
than four near-identical findings. Whoever fixes it needs to know this is one problem,
not four.

`Err[#]_current` and `Err[#]_fix` for an `AnsAll` finding must show **all four options**,
labelled and one per line, so the row can be applied without opening the admin page:

```
A. six baby dinosaurs.          A. six baby dinosaurs
B. two grown-up dinosaurs.  ->  B. two grown-up dinosaurs
C. one T-rex.                   C. one T-rex
D. two large dinosaurs.         D. two large dinosaurs
```

Use it only when *every* option changes. Two bad options out of four are two findings.

| Type | Flag when |
|---|---|
| Grammar | In question or answers; also tense/pronoun/plurality disagreement *between* them |
| Punctuation | Extra or missing; sentence answers without punctuation; non-sentence answers with it |
| Capitalization | Sentence answers capitalised. Non-sentence answers consistent within the row |
| Spacing | Leading, trailing or double spaces |
| Awkward Phrasing | Not idiomatic, natural English |
| Unclear | Question or answer vague, or open to more than one reading |
| Too Hard | Vocabulary, grammar or concepts beyond the target reading level |
| Other | — |

**The answer key.** `Correct[#]` is `A`, `B`, `C` or `D`.

- **`null`** means the column does not exist — the sheet predates the answer-key scrape.
  Say so in the report; do not imply key correctness was checked.
- **`""`** means the scraper found no single option marked correct: either none was set,
  or the page's dropdowns were never touched and all four defaulted to "correct". Record
  this as `Unclear` targeting `Q[#]`, details `no single answer marked correct`.
- A keyed answer that is wrong, or a question where another option is equally defensible,
  is `Unclear` on `Q[#]` — but only when the question's own wording settles it. The book
  text is not available, so do not flag a key merely because you cannot confirm it; say
  in the report that key correctness was checked only where wording made it decidable.
