# -*- coding: utf-8 -*-
"""Write the translation review set: every global word entry the 278 books use.

Each entry is judged against the definition and part of speech it will have *after* the
English fixes are applied, not the ones it has now -- reviewing against text that is
about to change would approve translations that go wrong the moment word_edit.py runs.

    uv run python trans_dump.py stats          # sizes and mechanical defects
    uv run python trans_dump.py batch 3        # print batch 3 for reading
"""
import collections
import json
import re
import sys

sys.path.insert(0, ".")
sys.stdout.reconfigure(encoding="utf-8")

from pathlib import Path  # noqa: E402

import make_reports  # noqa: E402
import manual_qa  # noqa: E402
from report import vocab_changes  # noqa: E402

WB = Path("data/2plus_check.xlsx")
OUT = Path("data/translation_entries.json")
BATCH = 120
HANGUL = re.compile(r"[가-힣ᄀ-ᇿ㄰-㆏]")
SENSE = {"Wrong Sense", "Part of Speech", "Wrong Entry"}


def build() -> list[dict]:
    _, brows = manual_qa.load_books()
    titles = {b: t for b, t, _ in brows}
    data = make_reports.load_findings(WB, titles)
    uses = make_reports.load_word_uses(WB)
    glossary = make_reports.load_glossary()

    # What the English fixes will make each entry say. Split entries are excluded: they
    # are becoming separate word records, so their translations follow that split rather
    # than this pass.
    planned, split = {}, set()
    for r in vocab_changes.build_rows(data, uses, glossary):
        if r["split"]:
            split.add(r["word_id"])
        else:
            planned[r["word_id"]] = r

    out = []
    for wid, books in sorted(uses.items(), key=lambda kv: int(kv[0])):
        if wid in split:
            continue
        g = glossary.get(wid)
        if not g:
            continue
        row = planned.get(wid)
        out.append({
            "word_id": wid,
            "word": (row or {}).get("fixed_word") or g.get("word", ""),
            "pos": (row or {}).get("fixed_pos") or g.get("pos", ""),
            "pos_was": g.get("pos", "") if row and row.get("fixed_pos") else "",
            "en": (row or {}).get("fixed_def") or g.get("definition", ""),
            "en_was": g.get("definition", "") if row and row.get("fixed_def") else "",
            "kor": g.get("kor", ""),
            "vie": g.get("vie", ""),
            "types": (row or {}).get("types", []),
            "books": len({u["book"] for u in books}),
            "sentence": next((u["sentence"] for u in books if u.get("sentence")), ""),
        })
    return out


def defects(e: dict) -> list[str]:
    out = []
    for lang, label in (("kor", "Korean"), ("vie", "Vietnamese")):
        text = (e[lang] or "").strip()
        if not text:
            out.append(f"{label} empty")
        elif text.lower() == e["word"].strip().lower():
            out.append(f"{label} is the English word")
        elif text.lower() == e["en"].strip().lower():
            out.append(f"{label} is the English definition")
        elif lang == "kor" and not HANGUL.search(text):
            out.append("Korean has no Hangul")
    return out


def render(e: dict, n: int) -> str:
    head = f"[{n}] #{e['word_id']} {e['word']!r} ({e['pos']}"
    head += f", was {e['pos_was']}" if e["pos_was"] else ""
    head += ")"
    if e["types"]:
        head += f"   [{', '.join(e['types'])}]"
    lines = [head]
    if e["en_was"]:
        lines.append(f"    EN was: {e['en_was']}")
        lines.append(f"    EN new: {e['en']}")
    else:
        lines.append(f"    EN: {e['en']}")
    lines.append(f"    KO: {e['kor'] or '(empty)'}")
    lines.append(f"    VI: {e['vie'] or '(empty)'}")
    return "\n".join(lines)


def main():
    if not OUT.exists() or (len(sys.argv) > 1 and sys.argv[1] == "stats"):
        entries = build()
        OUT.write_text(json.dumps(entries, ensure_ascii=False), encoding="utf-8")
    entries = json.loads(OUT.read_text(encoding="utf-8"))
    # Riskiest first: an entry whose meaning is changing has a translation that is about
    # to describe the wrong word, then one whose wording is changing, then the rest.
    # Reviewing in that order means stopping early still leaves the worst of it done.
    def risk(e):
        if set(e["types"]) & SENSE:
            return 0
        if e["en_was"] or e["pos_was"]:
            return 1
        return 2
    entries.sort(key=lambda e: (risk(e), int(e["word_id"])))

    cmd = sys.argv[1] if len(sys.argv) > 1 else "stats"
    if cmd == "stats":
        changing = [e for e in entries if e["en_was"] or e["pos_was"]]
        sense = [e for e in changing if set(e["types"]) & SENSE]
        print(f"{len(entries):,} word entries used by the books")
        print(f"  {len(changing):,} have their English definition or part of speech "
              f"changing")
        print(f"     of those, {len(sense):,} change MEANING -- the translation has to "
              f"be re-judged")
        print(f"  {len(entries) - len(changing):,} are unchanged; the translation is "
              f"judged against what is already there")
        found = collections.Counter(d for e in entries for d in defects(e))
        print("\nmechanical defects:")
        for why, n in found.most_common():
            print(f"  {n:>5}  {why}")
        print(f"\n{(len(entries) + BATCH - 1) // BATCH} batches of {BATCH}")
    elif cmd == "batch":
        i = int(sys.argv[2])
        chunk = entries[(i - 1) * BATCH:i * BATCH]
        print(f"### batch {i} of {(len(entries) + BATCH - 1) // BATCH} "
              f"-- {len(chunk)} entries\n")
        for j, e in enumerate(chunk, (i - 1) * BATCH + 1):
            print(render(e, j))
            print()


if __name__ == "__main__":
    main()
