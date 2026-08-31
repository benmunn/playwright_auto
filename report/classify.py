"""Assign every finding to a recurring error kind. First matching rule wins."""
import re
LQ, RQ = "\u201c", "\u201d"

RULES = [
 # code, sheet, type (None=any), regex on details (None=catch-all), label, tab name
 ("W01","Vocab","Too Hard",r"above level","Above-level word used inside the definition","Vocab above-level word"),
 ("W02","Vocab","Wrong Sense",None,"Definition gives the wrong sense of the word for this book","Vocab wrong sense"),
 ("W17","Vocab","Too Hard",None,"Definition says the same thing twice","Vocab repeated wording"),
 ("W03","Vocab","Answer Given",r"repeats both parts","Definition uses both forms of a two-part vocabulary word","Vocab two-part word"),
 ("W04","Vocab","Answer Given",None,"Definition uses the vocabulary word or its base","Vocab word in definition"),
 ("W05","Vocab","Other",r"field repeats|repeats the word","Definition field repeats the word instead of defining it","Vocab definition is the word"),
 ("W06","Vocab","Grammar",r"^plural|plural \S+ defines|singular definition","Plural vocabulary word given a singular definition","Vocab plural mismatch"),
 ("W18","Vocab","Grammar",r"given a plural definition|describes several|definition is plural","Singular vocabulary word given a plural definition","Vocab singular-plural mismatch"),
 ("W08","Vocab","Grammar",r"gerund|noun phrase|starts with|defines a verb|defines a noun|word class|the word is tagged as|starts as an action|written as a "+LQ+r"when|an -ing form used to define|describes a person","Definition's grammatical form does not match the vocabulary word","Vocab form vs word class"),
 ("W09","Vocab","Grammar",None,"Other grammar error inside the definition","Vocab definition grammar"),
 ("W10","Vocab","Part of Speech",None,"Part-of-speech tag disagrees with the definition","Vocab part-of-speech tag"),
 ("W11","Vocab","Lack of Context",None,"Definition too vague or too short to identify the word","Vocab definition too vague"),
 ("W12","Vocab","Awkward Phrasing",None,"Definition written in unnatural English","Vocab unnatural definition"),
 ("W13","Vocab","Spacing",None,"Missing, doubled or zero-width space in an entry","Vocab spacing"),
 ("W14","Vocab","Capitalization",None,"Capitalization error in an entry","Vocab capitalization"),
 ("W15","Vocab","Other",None,"Factual error or wrong content in an entry","Vocab factual error"),
 ("W19","Vocab","Spelling",None,"The vocabulary word itself is misspelled","Vocab misspelled word"),
 ("W20","Vocab","Wrong Entry",None,"The entry is for a different word or sense than the book uses","Vocab wrong entry"),

 # The story sentence is authored beside the word but never shown to students, so its
 # faults are split out from the rest of the Vocab entry rather than mixed in with them.
 ("S01","Vocab","Story Sentence",r"no space after the (full stop|period)|run.?together","Two sentences run together with no space after the full stop","Sentence run-together"),
 ("S02","Vocab","Story Sentence",r"no full stop at the end|has no full stop","Story sentence has no end punctuation","Sentence missing end stop"),
 ("S03","Vocab","Story Sentence",r"not a whole sentence|stops in the middle|stops before|cut short|joined by only a comma","Story sentence is incomplete or spliced together","Sentence incomplete"),
 ("S04","Vocab","Story Sentence",r"repeats a whole line","Story sentence repeats a line of the book","Sentence repeated line"),
 ("S05","Vocab","Story Sentence",r"space before the comma|quotation mark|hyphen where a dash|two spaces","Punctuation or spacing error in the story sentence","Sentence punctuation"),
 ("S06","Vocab","Story Sentence",None,"Wrong or missing word in the story sentence","Sentence wrong word"),

 ("T01","TMC",None,r"no blank|has no blank","Fill-in-the-blank stem is missing its blank","TMC stem has no blank"),
 ("T02","TMC","Spacing",r"missing space|no space|run of spaces|zero-width","Missing space between words","TMC missing space"),
 ("T03","TMC","Spacing",None,"Double space, trailing space or space before punctuation","TMC extra space"),
 ("T04","TMC","Grammar",r"tense|\bpresent\b|\bpast\b","Tense disagreement between the question and its options","TMC tense mismatch"),
 ("T05","TMC","Grammar",r"missing an article|missing the article|is missing "+LQ+r"the|"+LQ+r"an |uncountable|article","Wrong or missing article","TMC article error"),
 ("T06","TMC","Grammar",r"disagree|singular|plural|agreement","Subject-verb or singular/plural disagreement","TMC agreement error"),
 ("T07","TMC","Grammar",r"cannot fill|answering "+LQ+r"|where the other (three )?options|the stem asks|does not fit the stem|phrase where|answers "+LQ+r"|does not answer|where the question asks|verb phrase|noun phrases?\b","Option does not fit or answer the question stem","TMC option does not fit stem"),
 ("T08","TMC","Grammar",None,"Wrong word, word form or preposition","TMC wrong word form"),
 ("T09","TMC","Too Hard",None,"Above-level vocabulary in a question or option","TMC above-level word"),
 ("T10","TMC","Punctuation",None,"Punctuation inconsistent or missing across the four options","TMC punctuation"),
 ("T11","TMC","Capitalization",None,"Capitalization inconsistent within the question set","TMC capitalization"),
 ("T12","TMC","Other",r"identical|duplicate|same as|nearly the same|both mean","Two options duplicate each other","TMC duplicate options"),
 ("T13","TMC","Other",r"question text is|different book|pasted|editorial note|is missing|left in","Wrong or missing content pasted into a field","TMC wrong content in field"),
 ("T14","TMC","Other",r"misspell|spelling|hyphen","Misspelling or hyphenation error","TMC misspelling"),
 ("T15","TMC","Other",None,"Factual or logical fault in a question or option","TMC factual or logic fault"),
 ("T16","TMC","Unclear",None,"Question or answer key is ambiguous","TMC ambiguous question or key"),
 ("T17","TMC","Awkward Phrasing",None,"Question or option written in unnatural English","TMC unnatural phrasing"),

 ("C01","CC","Answer Given",None,"Sentence contains the answer word or its stem","CC answer given away"),
 ("C02","CC","Lack of Context",None,"Not enough context in the sentence to choose the answer","CC not enough context"),
 ("C03","CC","Too Hard",None,"Above-level vocabulary in a sentence or bank word","CC above-level word"),
 ("C04","CC","Other",r"duplicate|repeats|near-duplicate|almost the same|same as","Sentence duplicates another in the same row","CC duplicate sentence"),
 ("C05","CC","Awkward Phrasing",r"points outside|no referent|refers outside|has nothing before","Sentence refers to something outside itself","CC dangling reference"),
 ("C06","CC","Awkward Phrasing",None,"Sentence written in unnatural English","CC unnatural sentence"),
 ("C07","CC","Grammar",None,"Grammar error in a sentence","CC sentence grammar"),
 ("C08","CC","Punctuation",None,"Punctuation error in a sentence","CC punctuation"),
 ("C09","CC","Spacing",None,"Spacing error in a sentence","CC spacing"),
 ("C10","CC","Other",None,"Other content fault in a sentence","CC other content fault"),

 ("O01","OEC","Too Hard",None,"Above-level vocabulary or concept in a warm-up question","OEC above-level word"),
 ("O02","OEC","Requires Reading",None,"Cannot be answered before reading the book","OEC requires reading"),
 ("O03","OEC","Unclear",r"yes/no","Bare yes/no question","OEC yes-no question"),
 ("O04","OEC","Unclear",None,"Question gives no hint how to answer","OEC unclear question"),
 ("O05","OEC","Grammar",None,"Grammar error in a warm-up question","OEC question grammar"),
 ("O06","OEC","Awkward Phrasing",None,"Question written in unnatural English","OEC unnatural question"),
 ("O07","OEC","Spacing",None,"Spacing error in a warm-up question","OEC spacing"),
 ("O08","OEC","Punctuation",None,"Punctuation error in a warm-up question","OEC punctuation"),
 ("O09","OEC","Other",None,"Other fault in a warm-up question","OEC other fault"),

 # Listen & Read and Listen & Read Along came entirely from the human audit, so their
 # kinds are the types the reviewer used rather than anything a detector produced.
 ("L01","LR","Severe TTS Cutoff",None,"Recorded audio stops well before the end of the sentence","LR severe cutoff"),
 ("L02","LR","Minor TTS Cutoff",None,"Recorded audio is clipped slightly at the end","LR minor cutoff"),
 ("L03","LR","Start TTS Cutoff",None,"Recorded audio starts late, losing the first word","LR start cutoff"),
 ("L04","LR","TTS Pronunciation",None,"Recorded audio mispronounces or distorts a word","LR pronunciation"),
 ("L05","LR","Text-TTS Mismatch",None,"Recorded audio does not match the printed text","LR text-audio mismatch"),
 ("L06","LR","Text",None,"Mistake in the printed text on the page","LR page text"),
 ("L07","LR","Other",None,"Other fault noticed while listening and reading","LR other fault"),

 ("A01","LRA","Severe TTS Cutoff",None,"Recorded audio stops well before the end of the sentence","LRA severe cutoff"),
 ("A02","LRA","Minor TTS Cutoff",None,"Recorded audio is clipped slightly at the end","LRA minor cutoff"),
 ("A03","LRA","Start TTS Cutoff",None,"Recorded audio starts late, losing the first word","LRA start cutoff"),
 ("A04","LRA","TTS Pronunciation",None,"Recorded audio mispronounces or distorts a word","LRA pronunciation"),
 ("A05","LRA","Text-TTS Mismatch",None,"Recorded audio does not match the printed text","LRA text-audio mismatch"),
 ("A06","LRA","Text",None,"Mistake in the printed text on the page","LRA page text"),
 ("A07","LRA","Other",None,"Other fault noticed while listening and reading along","LRA other fault"),
]
BY_CODE = {r[0]: r for r in RULES}

def classify(f):
    for code, sh, ty, pat, label, tab in RULES:
        if sh and f["sheet"] != sh: continue
        if ty and f["type"] != ty: continue
        if pat and not re.search(pat, f["details"], re.I): continue
        return code
    return "UNCLASSIFIED"
