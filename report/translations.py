"""Reviewed corrections to the Korean and Vietnamese definitions.

Every global word entry carries three definitions -- English, Korean, Vietnamese -- and
the English one is being rewritten. A reworded English definition usually leaves the
translations still true: turning "feeling anxiety or concern" into "feeling afraid that
something bad might happen" does not change what the word means, so the Korean stays
right. A re-sensed one does not: once `bark` stops meaning the noise a dog makes and
starts meaning what covers a tree, a Korean definition reading 짖다 describes a word that
no longer exists at that id.

The bar for changing a translation is that it is plainly wrong for the entry as it will
stand: a different sense, a different part of speech, the English copied into the box, or
a word that simply is not what the definition says. A translation that is merely clumsy,
narrower than the English, differently emphasised, or one of several reasonable phrasings
is left exactly as it is. Rewriting a language to read better is how a translation ends
up worse than the one it replaced.

Corrections live in data/translation_fixes.json rather than in this file: there are
thousands of entries to work through and they are recorded a batch at a time. Each is

    "34455": {"kor": "...", "vie": "...", "why": "..."}

with only the field that was wrong present -- a field left out is a field left alone.
`missing` lists boxes that are empty, which is a gap for a translator to fill rather than
an error to correct here, so nothing is invented for them.
"""

from __future__ import annotations

import json
from pathlib import Path

FIXES = Path("data/translation_fixes.json")

_cache: dict | None = None


def _load() -> dict:
    global _cache
    if _cache is None:
        if FIXES.exists():
            raw = json.loads(FIXES.read_text(encoding="utf-8"))
        else:
            raw = {}
        # Keep whatever else the file carries. Rebuilding it from two known keys
        # silently discarded anything recorded alongside them -- notes about English
        # definitions that are themselves wrong were written once and then dropped by
        # the next save.
        raw.setdefault("fixes", {})
        raw.setdefault("missing", {})
        raw.setdefault("english_errors", {})
        _cache = raw
    return _cache


def fixed(word_id: str, lang: str) -> str:
    return _load()["fixes"].get(str(word_id), {}).get(lang, "")


def reason(word_id: str) -> str:
    return _load()["fixes"].get(str(word_id), {}).get("why", "")


def all_fixes() -> dict[str, dict]:
    return dict(_load()["fixes"])


def missing() -> dict[str, list[str]]:
    return dict(_load()["missing"])


def english_errors() -> dict[str, str]:
    """Entries whose own English definition is wrong. No finding covers these, so no
    run will ever touch them; they are recorded here so they are not lost."""
    return dict(_load()["english_errors"])


def record(word_id: str, **fields: str) -> None:
    """Add or update one entry and write the file straight away.

    Written per entry rather than per batch so that stopping half way through a review
    leaves the decisions already made on disk.
    """
    data = _load()
    entry = data["fixes"].setdefault(str(word_id), {})
    entry.update({k: v for k, v in fields.items() if v})
    FIXES.parent.mkdir(parents=True, exist_ok=True)
    FIXES.write_text(json.dumps(data, ensure_ascii=False, indent=1, sort_keys=True),
                     encoding="utf-8")
