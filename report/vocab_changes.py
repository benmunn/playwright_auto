"""One consolidated list of every change wanted to the global vocabulary entries.

The word, its part of speech and its definition belong to the global word list, not to a
book, so a fix to any of them lands on every book linked to that entry. The per-book
reports therefore list the same fix several times over, and give no warning when two
books want the entry to say different things.

These rows are the other way round: one per global word. Where every book that uses an
entry wants the same thing, that is one row and one edit, and it goes on the main sheet.
Where they do not -- `bark` the dog sound against `bark` on a tree -- the entry cannot
serve both, so the row is split into 587A / 587B, one per sense, each naming the books it
covers, and those go on a second sheet of their own. Splitting is a judgment about meaning
rather than a rule about text, so the splits below were decided by reading the story
sentence of every book involved; SPLITS records that reading.

Story-sentence findings are deliberately absent: they are faults in a book's own
sentence, not in the shared entry, and there is no column here that could hold them.
"""

from __future__ import annotations

import collections

from . import translations

# --------------------------------------------------------------------------------------
# Reviewed decisions
# --------------------------------------------------------------------------------------

# word id -> ordered (what this sense means, the books that use it that way).
# Every one of these was checked against the story sentence of each book listed. A group
# whose books raised no finding still gets a row: the point of splitting is that the other
# group's fix would leave this one wrong, which is exactly what needs saying.
SPLITS: dict[str, list[tuple[str, set[str]]]] = {
    "33624": [("to stop trying", {"81", "751", "146", "1118"}),
              ("to let something go in order to have something else",
               {"728"})],
    "33775": [("to push something into a place", {"43", "2819"}),
              ("to become firmly attached to a surface", {"2832", "73"})],
    "33851": [("to open out, or move over a wider area", {"26", "2831"}),
              ("news reaching more and more people", {"64"})],
    "34814": [("to get on a bus, train or other transport",
               {"1268", "1267", "1257", "1232", "1163"}),
              ("to draw something in and use it", {"1255"}),
              ("to carry something from one place to another", {"746"})],
    "34932": [("a noise you hear", {"26"}),
              ("a long narrow area of sea", {"835"})],
    "35086": [("matter that is firm and keeps its shape", {"1269", "1182"}),
              ("a three-dimensional shape such as a cube",
               {"1254", "1239", "1253"})],
    "35211": [("to pour water on plants", {"1079", "1042"}),
              ("(of the mouth) to fill with spit", {"706"})],
    "41020": [("to ask someone to pay a price", {"94"}),
              ("to rush forward in an attack", {"76", "155"})],
    "42715": [("to form something into a shape", {"983", "707"}),
              ("a shape such as a circle or square", {"1074", "1095", "1239"})],
    "46522": [("the answer to a problem", {"756"}),
              ("a liquid with something dissolved in it", {"749"})],
    "46579": [("a metal fastener", {"1101"}),
              ("to twist your face", {"751"})],
    "49366": [("to place things one on top of another", {"1147"}),
              ("one of several levels, one above another", {"1266"})],
    "32585": [("to feel sad that someone is absent", {"2834"}),
              ("to fail to see or hear something", {"43"})],
    "32841": [("to repair something broken", {"2819"}),
              ("to prepare food or a drink", {"2840"})],
    "32989": [("to collect someone in a car", {"856", "1057"}),
              ("to lift something up off the ground", {"606", "1254", "1272"})],
    "33682": [("deep in colour", {"26"}),
              ("with little or no light", {"606"})],
    "34682": [("a plan for a new law", {"696"}),
              ("a piece of paper money", {"949"})],
    "34837": [("a device for weighing", {"1052"}),
              ("one of the plates covering a fish or reptile", {"940"})],
    "41780": [("a photograph or drawing", {"1227"}),
              ("to see something in your mind", {"610", "960"})],
    "42707": [("to hold something and go away with it", {"983"}),
              ("to become so excited that you lose control", {"129", "2834"})],
    "43813": [("to make the sharp sound of a dog", {"816"}),
              ("the hard outer covering of a tree", {"1258"})],
    "44002": [("to break open suddenly", {"672"}),
              ("a short, sudden sound or action", {"107"})],
}

# Two QA passes sometimes logged the same cell twice with different suggestions. Where
# the two agree in substance the longer one is kept automatically; these are the ones
# that genuinely disagreed, resolved by reading the book's story sentence.
# (word id, book id, field) -> the suggestion to use.
CHOSEN: dict[tuple[str, str, str], str] = {
    ("3619", "77", "DEF"): "to shout happily to show that you support someone",
    ("51260", "1266", "DEF"): ("the top layer of a forest, where branches and leaves "
                               "spread out; also curtains hanging above a bed"),
    ("49366", "1266", "DEF"): "one of several levels of something, lying one above another",
    ("45035", "84", "POS"): "verb",
    ("42715", "1074", "DEF"): "a form such as a circle, a square, or a triangle",
    ("4106", "208", "DEF"): "a creature in stories that drinks blood and comes out at night",
    ("4107", "208", "DEF"): "a creature in stories that walks around after dying",
    ("4110", "208", "DEF"): "not in danger; free from harm",
    ("54198", "795", "DEF"): "to shake with small, quick movements",
    ("3563", "52", "DEF"): ("a boat that the wind pushes along by blowing on a large "
                            "piece of cloth"),
    ("3564", "52", "DEF"): "a short pole with a flat end, used to push a small boat through water",
    ("3565", "52", "DEF"): "a machine that makes something move, usually by burning fuel",
    ("3567", "52", "DEF"): "to make a fire stop burning",
    ("43267", "1234", "DEF"): ("thinking or showing that you think someone or something "
                               "is important"),
    ("55809", "1231", "DEF"): "the way things are arranged, one after another",
    ("3396", "91", "DEF"): "to travel on water in a boat pushed by the wind",
    ("3762", "91", "DEF"): "to give something to someone and get something else back",
    ("3764", "91", "DEF"): "the leader of a team",
    ("3765", "91", "DEF"): "unkind; treating other people badly",
    ("34932", "835", "DEF"): ("a long narrow area of sea between two areas of land; also "
                              "a noise you hear with your ears"),
    ("34263", "606", "DEF"): "to go down to the ground; to go into something by accident",
}

# Noticed while consolidating rather than during the per-book QA, so there is no finding
# to carry them: word id -> what a reviewer should know when editing this entry.
NOTES: dict[str, str] = {
    "4110": ("The story uses the adjective sense (\"safe from zombies\"), but the entry "
             "is tagged noun -- the part of speech needs changing too."),
    "3619": ("The story uses the verb (\"they cheered\"), but the entry is tagged noun -- "
             "the part of speech needs changing too."),
    "42825": ("Book 829 uses this for the Statue of Liberty, not a map, so keep the "
              "wording broad enough to cover a symbol that is not on a map."),
    "34682": ("No book was flagged for the money sense, but book 949 (\"the $100 bill\") "
              "matches neither the bird nor the law sense and needs its own entry."),
}

# A suggestion that was filed against the wrong field and cannot be used as written.
# (word id, field) -> why. The field is left blank, so nothing is edited from it.
UNUSABLE: dict[tuple[str, str], str] = {
    ("43110", "POS"): ("The suggestion for this field is a definition, not a part of "
                       "speech. The entry is already tagged noun, so only the "
                       "definition needs changing."),
    ("40188", "DEF"): ("The finding is about a different word. This entry is 'cause', "
                       "defined as 'to make something happen', in a sentence about "
                       "litter causing pollution; the finding objects to the word "
                       "'chemistry' and suggests a definition of a chemical. Applying "
                       "it would replace the definition of 'cause' with one for a "
                       "substance, so the slot it belongs to has to be found by hand."),
}

# The database refuses two entries sharing a word, part of speech and definition, and
# for these the wording the fix wanted is already taken by a duplicate entry. The
# duplicates cannot be deleted -- other people are working from the same list -- so the
# definition is reworded instead: the same meaning at the same reading level, phrased
# differently enough to be its own row. Without this the entry keeps whatever broken
# definition it has now, which is what the books linked to it still show.
REPHRASED: dict[str, str] = {
    "3339": "a small green vegetable shaped like a ball",
    "3347": "a small patch of water lying on the ground",
    "3407": "to run after someone or something so that you can catch them",
    "3437": "to do something again and again so that you get better at it",
    "3633": "soft and sticky to touch",
    "33956": "a building where plays and shows are performed",
    "40845": "to take a quick secret look at something",
    "43052": "in a place higher than something else",
    "43693": "a big special meal with lots of food, eaten to celebrate",
    "45035": "to save someone from danger or from a harmful situation",
    "46227": "to touch someone's skin lightly so that they laugh",
    "46820": "with no chance of success; feeling there is no hope ahead",
    "47450": "being held responsible for something bad that happened",
    "48583": "a large amount; more than you need",
    "48802": "covered with sharp points that can scratch you",
    "49055": "something you say or do to make people laugh",
    "49174": "the time of year when crops are gathered in from the fields",
    "55285": "to make a picture using a pencil, pen, or crayon",
    "55516": "a number that shows a part of a whole, like 1/2 or 1/4",
    "57090": "to find out how big, how long, or how much something is",
    "60985": "a big heavy gun on wheels that fires metal balls",
    "61238": "to draw around the edge of something to copy its shape",
}

FIELD = {"W": "word", "POS": "pos", "DEF": "def"}


# --------------------------------------------------------------------------------------
# Building the rows
# --------------------------------------------------------------------------------------


def _pick(field: str, wid: str, book: str, items: list[dict]) -> str:
    """The one suggestion to show for a field that may carry several findings."""
    fixes = [f["fix"].strip() for f in items if f["fix"].strip()]
    if not fixes:
        return ""
    if len(fixes) == 1:
        return fixes[0]
    chosen = CHOSEN.get((wid, book, field))
    if chosen:
        return chosen
    # Same suggestion worded twice: the fuller wording says strictly more.
    return max(fixes, key=len)


def _label(i: int, total: int, wid: str) -> str:
    return wid if total == 1 else f"{wid}{chr(ord('A') + i)}"


def build_rows(findings: list[dict], uses: dict[str, list[dict]],
               glossary: dict[str, dict] | None = None) -> list[dict]:
    """One row per global word entry, split by sense where the books disagree.

    `uses` maps a word id to every book that links to it, flagged or not, so a group
    whose books raised no finding can still be named. `glossary` maps a word id to its
    global record, which is where the Korean and Vietnamese definitions come from -- they
    belong to the entry, not to any book, so no per-book scrape carries them.
    """
    glossary = glossary or {}
    by_word: dict[str, dict[str, list[dict]]] = collections.defaultdict(
        lambda: collections.defaultdict(list))
    for f in findings:
        if f["sheet"] == "Vocab" and f["prefix"] in FIELD:
            by_word[f["ctx"]["word_id"]][f["book_id"]].append(f)

    # An entry can need a translation corrected without any English finding against it.
    # Those belong on the same sheet -- it is the list of changes wanted to word entries,
    # not the list of English findings -- so they are added here with no flagged books.
    for wid in translations.all_fixes():
        if wid not in by_word and wid in uses:
            by_word[wid] = {}

    # A finding on a word whose slot carries no id cannot be placed on this sheet: the
    # sheet is keyed by the global entry, and without an id there is nothing to key it
    # to. That happens when the workbook has not been through word_ids.py, or when a
    # re-scrape put a new word in a slot that has not been matched yet. Dropping it here
    # would lose it silently, so it is counted and named by the caller instead.
    unmatched = by_word.pop("", None)
    if unmatched:
        print(f"  ! {sum(len(v) for v in unmatched.values())} vocabulary finding(s) on "
              f"{len(unmatched)} book(s) have no word id and are left off the Vocab "
              f"Changes sheet -- run word_ids.py --match --write on this workbook")

    rows = []
    for wid, books in sorted(by_word.items(), key=lambda kv: int(kv[0])):
        entry = (uses.get(wid) or [{}])[0]
        groups = SPLITS.get(wid)
        if groups is None:
            groups = [("", {u["book"] for u in uses.get(wid, [])} or set(books))]
        for i, (sense, group_books) in enumerate(groups):
            mine = {b: fs for b, fs in books.items() if b in group_books}
            items = [f for fs in mine.values() for f in fs]
            fixed = {}
            for pfx, name in FIELD.items():
                per_book = [(b, [f for f in fs if f["prefix"] == pfx])
                            for b, fs in mine.items()]
                cands = [_pick(pfx, wid, b, fs) for b, fs in per_book if fs]
                cands = [c for c in cands if c]
                if (wid, pfx) in UNUSABLE:
                    cands = []
                fixed[name] = max(cands, key=len) if cands else ""
                if name == "def" and wid in REPHRASED:
                    fixed[name] = REPHRASED[wid]
            entry_g = glossary.get(wid, {})
            rows.append({
                "id": _label(i, len(groups), wid),
                "word_id": wid,
                "kor": entry_g.get("kor", ""),
                "vie": entry_g.get("vie", ""),
                "fixed_kor": translations.fixed(wid, "kor"),
                "fixed_vie": translations.fixed(wid, "vie"),
                "translation_note": translations.reason(wid),
                "sense": sense,
                "word": entry.get("word", ""),
                "pos": entry.get("pos", ""),
                "definition": entry.get("definition", ""),
                "fixed_word": fixed["word"],
                "fixed_pos": fixed["pos"],
                "fixed_def": fixed["def"],
                "types": (sorted({f["type"] for f in items})
                          or (["Translation"] if not groups[0][0] else
                              ["Sense Conflict"])),
                "items": items,
                "books": sorted(group_books, key=lambda b: int(b) if b.isdigit() else 0),
                "flagged": sorted(mine, key=lambda b: int(b) if b.isdigit() else 0),
                "uses": [u for u in uses.get(wid, []) if u["book"] in group_books],
                "note": NOTES.get(wid, ""),
                "split": len(groups) > 1,
            })
    return rows


def details(row: dict, titles: dict[str, str]) -> str:
    """The prose cell: what is wrong, and which books this row speaks for."""
    out = []
    if row["split"]:
        # The sheet itself says these are splits, so the row only has to say which sense.
        out.append(f"This row covers only the sense: {row['sense']}.")
        if not row["flagged"]:
            out.append("No fix was requested for this sense, but the other row's fix "
                       "would leave these books wrong. Give this sense its own word "
                       "entry and link these books to it instead.")
        # The QA wrote these suggestions expecting one shared entry, so several bolt the
        # rival sense on with "also". Once the entry is split that half belongs to the
        # other row.
        if any("; also" in row[k] for k in ("fixed_word", "fixed_pos", "fixed_def")):
            out.append("The suggested definition still carries the other sense after "
                       "\"also\" -- it was written before this entry was split. Keep "
                       "only the part that matches this row's sense.")
    seen = set()
    for f in row["items"]:
        d = f["details"].strip()
        if d and d.lower() not in seen:
            seen.add(d.lower())
            out.append(f"[{f['type']}] {d}")
    for (wid, pfx), why in UNUSABLE.items():
        if wid == row["word_id"]:
            out.append(f"Note: {why}")
    if row.get("translation_note"):
        out.append(f"Translation: {row['translation_note']}")
    if row["note"]:
        out.append(f"Note: {row['note']}")
    # The books themselves are in the first two columns now; what is worth repeating
    # here is how each one actually uses the word.
    for u in row["uses"][:3]:
        if u.get("sentence"):
            out.append(f"  {titles.get(u['book'], u['book'])}: “{u['sentence']}”")
    return "\n".join(out)
