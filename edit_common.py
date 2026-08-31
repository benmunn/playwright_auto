"""The parts every activity edit run needs, kept in one place.

Three activities are edited by script now -- open-ended questions, context clue and
text multiple choice -- and each one saves through the same admin form, fails in the
same ways and needs the same guards. The report generators drifted apart when they were
copied between scripts; these are the pieces that would drift next, so they live here
instead.

Nothing in this module knows about a particular activity. Where the three differ -- the
URL, which findings apply, how a field is found on the page -- that stays with the
activity.
"""

from __future__ import annotations

import json
from pathlib import Path


class Mutation:
    """Catches what a save mutation actually answered.

    The endpoint replies HTTP 200 whether it saved or not, putting the failure in the
    body, so a rejected save looks exactly like a successful one from the outside. It
    hid a unique-constraint rejection through the whole first vocabulary run, and it
    hides an unset speech key on the open-ended-question page today. Reading the body is
    the only way to tell the two apart.
    """

    def __init__(self, page, path_marker: str):
        # Which mutation's errors matter -- the page fires several on save, and a
        # failure belonging to some other query is not this run's business.
        self.marker = path_marker
        self.errors: list[str] = []
        page.on("response", self._on)

    def _on(self, response) -> None:
        if response.request.method != "POST" or "graphql" not in response.url:
            return
        try:
            body = response.json()
        except Exception:                                          # noqa: BLE001
            return
        for err in (body or {}).get("errors") or []:
            if self.marker in str(err.get("path", "")):
                self.errors.append(str(err.get("message", "")).strip())

    def take(self) -> list[str]:
        out, self.errors = self.errors, []
        return out


def same(a: str, b: str) -> bool:
    """Exact once the outer whitespace is gone.

    Deliberately stricter than word_ids.norm(): a good number of these fixes correct
    spacing or punctuation and nothing else, and a comparison that collapsed whitespace
    would call such an edit verified without it ever having been made.
    """
    return (a or "").strip() == (b or "").strip()


def journal(path: Path, rec: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(rec, ensure_ascii=False) + "\n")


def settled_fields(path: Path) -> dict[str, set[str]]:
    """book id -> the fields the journal records as needing no further work.

    "already" is a field somebody else had put right before the run got there. It is
    finished in the only sense that matters here, so it counts alongside "saved".

    Asking merely whether a book has been saved before is not enough: a book saved for
    its first question would then be treated as done and a second still-pending field on
    the same book skipped for good. That is exactly what once left three Vietnamese
    boxes empty.
    """
    out: dict[str, set[str]] = {}
    if not path.exists():
        return out
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        rec = json.loads(line)
        if rec.get("status") not in ("saved", "already"):
            continue
        out.setdefault(str(rec["book_id"]), set()).update(rec.get("fields", []))
    return out
