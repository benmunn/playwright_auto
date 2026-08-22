"""Rewrite finding details into language a proficient non-native speaker can act on.

Kept as-is: singular, plural, article, clause, uncountable, comma, period, hyphen.
Those are everyday EFL teaching terms for this audience. Replaced: terms that are
linguistics jargon rather than classroom vocabulary.
"""
import re

SUBS = [
 # --- word-level jargon, longest patterns first -----------------------------
 (r"definition contains the headword stem", "the definition uses the base of the vocabulary word"),
 (r"definition contains the headword", "the definition uses the vocabulary word itself"),
 (r"definition field repeats the headword; no definition given",
  "the definition box only repeats the word; no meaning is given"),
 (r"definition repeats both parts of the headword", "the definition uses both forms of the vocabulary word"),
 (r"repeats the headword stem", "repeats the base of the vocabulary word"),
 (r"repeats the headword", "repeats the vocabulary word"),
 (r"plural (\S+) defines a singular headword",
  r"the definition is plural \1 but the vocabulary word is singular"),
 (r"headword inflection reads", "the word form is written as"),
 (r"\bheadwords\b", "vocabulary words"),
 (r"\bheadword\b", "vocabulary word"),
 (r"the stem of the answer", "the base of the answer word"),
 (r"the stem of\b", "the base of"),
 (r"shares its stem with", "has the same base as"),
 (r"\bstem\b(?! of)", "question text"),
 # --- reference / grammar jargon --------------------------------------------
 (r"is definite with no referent", "uses \u201cthe\u201d although nothing has been mentioned yet"),
 (r"is definite where", "uses \u201cthe\u201d where"),
 (r"is definite\b", "uses \u201cthe\u201d"),
 (r"has no referent before reading", "has nothing to refer to before the student reads the book"),
 (r"has no referent inside this sentence", "has nothing to refer to inside this sentence"),
 (r"has no referent", "has nothing to refer to"),
 (r"has no antecedent", "has nothing earlier to refer back to"),
 (r"which has no antecedent", "with nothing earlier to refer back to"),
 (r"\bantecedent\b", "earlier reference"),
 (r"spurious comma", "unnecessary comma"),
 (r"\bspurious\b", "unnecessary"),
 (r"fragment appended", "incomplete phrase added at the end"),
 (r"\bappended\b", "added at the end"),
 (r"a gerund defining a noun", "an -ing form used to define a noun"),
 (r"a gerund answering", "an -ing form answering"),
 (r"gerund (\S+) where", r"the -ing form \1 where"),
 (r"\bgerund\b", "-ing form"),
 (r"infinitive (\S+) defines", "the “to …” form \1 defines"),
 (r"\binfinitive\b", "\u201cto \u2026\u201d form"),
 (r"no terminal period", "no period at the end"),
 (r"a terminal period", "a period at the end"),
 (r"terminal period", "period at the end"),
 (r"needs a hyphen as a compound modifier", "needs a hyphen: the two words act as one describing word"),
 (r"misplaced modifier", "the describing phrase is in the wrong place"),
 (r"dangling modifier", "the describing phrase attaches to the wrong thing"),
 (r"the modifier is misplaced", "the describing phrase is in the wrong place"),
 (r"\bmodifier\b", "describing phrase"),
 (r"is not idiomatic", "is not natural English"),
 (r"\bidiomatic sense\b", "figurative meaning"),
 (r"\bidiomatic\b", "natural English"),
 (r"presupposes", "assumes"),
 (r"presuppose", "assume"),
 (r"is hard to parse", "is hard to read"),
 (r"does not parse", "does not make sense"),
 (r"hard to parse", "hard to read"),
 (r"is circular", "explains the word by using the word itself"),
 (r"\bcircular\b", "explains itself in a circle"),
 (r"typographic sense", "printing sense"),
 (r"a long run-on with", "one long sentence that should be split, with"),
 (r"is a run-on", "joins two sentences without the right punctuation"),
 (r"\brun-on\b", "two sentences joined without punctuation"),
 (r"comma splice between", "only a comma joins"),
 (r"\bcomma splice\b", "two sentences joined by only a comma"),
 (r"bare fragment where", "an incomplete phrase where"),
 (r"periods on fragments where", "periods on incomplete phrases where"),
 (r"sentence fragment with no main verb", "an incomplete sentence with no main verb"),
 (r"an unexplained fragment", "an incomplete phrase with no explanation"),
 (r"are fragments", "are incomplete phrases"),
 (r"\bfragment\b", "incomplete phrase"),
 (r"\binflection\b", "word form"),
 (r"no gloss to identify the word", "no short explanation to identify the word"),
 (r"no gloss to identify the phrase", "no short explanation to identify the phrase"),
 (r"no defining gloss", "no short explanation of the meaning"),
 (r"no example or gloss", "no example or short explanation"),
 (r"the gloss", "the short explanation"),
 (r"\bgloss\b", "short explanation"),
 (r"leads with the", "gives first the"),
 (r"\bgarbled\b", "jumbled and unclear"),
 (r"\bopaque\b", "impossible to understand"),
 (r"\bmeta-gloss\b", "explanation about the definition instead of the meaning"),
 (r"\bappositive\b", "explaining phrase"),
 (r"\bcopula\b", "linking verb"),
]

def plain(text: str) -> str:
    if not text:
        return text
    out = text
    for pat, rep in SUBS:
        out = re.sub(pat, rep, out, flags=re.I)
    return re.sub(r"\s{2,}", " ", out).strip()

BANNED = re.compile(r"\bheadword|\bgloss|\breferent|\bspurious|\bappend|\bgerund|\bantecedent"
                    r"|\bparse\b|\binflection\b|comma splice|\brun-on\b|\bcopula\b|\bappositive\b"
                    r"|typographic|\bmeta-", re.I)
