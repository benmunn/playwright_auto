"""What the Feedback cell says about each by-type group.

The by-type summary gives one row per (activity, error type), and its Feedback cell has
to explain the class of error rather than any single instance: a sentence naming it, the
shapes it takes in this data, and one short example. None of that can be derived from a
finding, so it is written here, keyed the way `render_by_type` groups its rows.

Each entry is (description, bullets, example current, example suggested). The example is
trimmed to the part that shows the fault -- a reader comparing two long sentences has to
hunt for the difference, which defeats the point of an example.

Groups with no entry fall back to a description built from the recurring-error label, so
a newly-introduced error type still renders rather than breaking the report.
"""

# (sheet, type, recurring-kind code or None) -> (description, bullets, current, suggested)
BLURBS: dict[tuple, tuple[str, list[str], str, str]] = {

    # ---- Open-Ended Questions -------------------------------------------------------
    ("OEC", "Awkward Phrasing", None): (
        "OEQ written in unnatural or unclear English",
        ["unnatural collocations", "dangling modifiers", "confusing wording",
         "statements mixed with questions"],
        "Imagine if you raised a frog for a pet, would you be able to let it go free?",
        "Imagine if you raised a frog for a pet. Would you be able to let it go free?"),
    ("OEC", "Grammar", None): (
        "Grammar errors in OEQ",
        ["Wrong tenses", "Missing articles", "Subject-verb agreement"],
        "Who do you know a good citizen around you?",
        "Who do you know that is a good citizen?"),
    ("OEC", "Punctuation", None): (
        "Punctuation errors in a OEQ",
        ["extra punctuation", "missing punctuation", "wrong punctuation type"],
        "Which words rhyme with cat??",
        "Which words rhyme with cat?"),
    ("OEC", "Requires Reading", None): (
        "Some OEQ require knowing information from the book",
        [],
        "How did the city of Chicago become an important, fast-moving metropolis?",
        "What do you think helps a small town grow into a big city?"),
    ("OEC", "Spacing", None): (
        "Spacing errors in an OEQ",
        ["missing spaces"],
        "If you did not likethe plan, what would you do?",
        "If you did not like the plan, what would you do?"),
    ("OEC", "Too Hard", None): (
        "OEQ using words or ideas the book has not taught yet",
        ["above-level vocabulary", "topic words the book introduces later",
         "concepts a student cannot be expected to know beforehand"],
        "How can you describe sediment?",
        "What do you think happens to sand and mud at the bottom of a river?"),
    ("OEC", "Unclear", None): (
        "OEQ a student cannot tell how to answer",
        ["bare yes/no questions with no follow-up",
         "no hint about what kind of answer is wanted", "ambiguous wording"],
        "Do you know any Korean legends?",
        "Do you know any old legends from your own country? What happens in them?"),
    ("OEC", "Other", "O09"): (
        "Other faults in an OEQ",
        ["the second question assumes an answer to the first",
         "the question gives away the story", "factually wrong content"],
        "Have you ever been to a dinosaur museum? What dinosaur did you like seeing?",
        "Have you ever been to a dinosaur museum? If you have, which dinosaur did you "
        "like best?"),

    # ---- Context Clue ---------------------------------------------------------------
    ("CC", "Answer Given", None): (
        "The sentence contains the answer word, so the clue gives itself away",
        ["the answer word appears unblanked", "the answer's base or stem appears",
         "another form of the answer appears (plural, past tense)"],
        "I feel lonely when I am alone.",
        "I feel lonely when no one is with me."),
    ("CC", "Awkward Phrasing", None): (
        "Context Clue sentences written in unnatural or unclear English",
        ["unnatural collocations", "one long sentence that should be split",
         "a quoted line that does not connect to the rest",
         "the sentence points at something outside itself"],
        "Money is a resource we use to build.",
        "Money is a resource we use to buy the things we need."),
    ("CC", "Grammar", None): (
        "Grammar errors in a Context Clue sentence",
        ["missing articles", "wrong word order", "subject-verb agreement",
         "pronouns with no clear referent"],
        "Dad will pick up me by car.",
        "Dad will drive to our school to pick up my sister."),
    ("CC", "Lack of Context", None): (
        "The sentence does not give enough to work out the answer",
        ["another bank word fits the blank equally well",
         "the sentence names the word but never shows what it means",
         "no clue at all in the surrounding words"],
        "John's body gave out.",
        "John's body gave out, growing too weak to work anymore."),
    ("CC", "Punctuation", None): (
        "Punctuation errors in a Context Clue sentence",
        ["missing or unmatched quotation marks", "missing final period",
         "wrong punctuation type"],
        "I paid with a ten dollar bill.",
        "I paid with a ten-dollar bill."),
    ("CC", "Spacing", None): (
        "Spacing errors in a Context Clue sentence",
        ["no space after a period", "no space after a closing quotation mark"],
        "Ice can erode Earth.When the ice moves away...",
        "Ice can erode Earth. When the ice moves away..."),
    ("CC", "Too Hard", None): (
        "Context Clue sentences using words above the level of the book",
        ["above-level vocabulary in the sentence",
         "above-level vocabulary in the word bank"],
        "The stories are similar, looking almost the same but not exactly identical.",
        "The stories are similar; the most important parts are the same but a few small "
        "parts are different."),
    ("CC", "Other", "C04"): (
        "Two sentences in the same row are the same or nearly the same",
        ["a sentence repeated from another question in the row",
         "two sentences differing by only a word"],
        "Planes land at the airport.",
        "Travelers wait for their flights at the airport."),
    ("CC", "Other", "C10"): (
        "Other content faults in a Context Clue sentence",
        ["misspelled names", "a definition that does not match the story",
         "two sentences in the row giving each other away"],
        "Nerves are thin fibers that carry messages to and from your brain.",
        "Nerves are thin fibers that run all through your body."),

    # ---- Word Meaning Match ---------------------------------------------------------
    ("Vocab", "Answer Given", None): (
        "The definition contains the word it is defining",
        ["the word itself inside the definition", "the word's base or stem",
         "both halves of a two-part word", "“plural of…” openings"],
        "parking space — a marked space where a car can be parked",
        "parking space — a marked area where a car can be left"),
    ("Vocab", "Awkward Phrasing", None): (
        "Definitions written in unnatural English",
        ["word order an English speaker would not use",
         "comparisons that do not read naturally"],
        "in a not certain way",
        "in a way that shows you are not sure"),
    ("Vocab", "Capitalization", None): (
        "Definitions inconsistently capitalized within a book",
        ["a block of definitions starting with a capital while the rest start lowercase"],
        "Unable to see",
        "unable to see"),
    ("Vocab", "Grammar", None): (
        "Grammar errors inside the definition",
        ["a singular word given a plural definition",
         "a plural word given a singular definition",
         "a definition whose form does not match the word class",
         "missing articles, subject-verb agreement"],
        "small round marks",
        "a small round mark"),
    ("Vocab", "Lack of Context", None): (
        "Definitions too vague or too close to another word in the same book",
        ["two words given definitions that cannot be told apart",
         "a definition that uses another word from the same book",
         "too short to identify the word"],
        "germ — a tiny living thing that causes disease",
        "germ — a tiny living or non-living thing that spreads illness from person "
        "to person"),
    ("Vocab", "Part of Speech", None): (
        "The part-of-speech tag does not match the word or its definition",
        ["a verb tagged as a noun", "an adjective tagged as a noun",
         "a multi-word item tagged as a single word class",
         "the tag left at the default “noun”"],
        "hairy — tagged noun",
        "hairy — adjective"),
    ("Vocab", "Spacing", None): (
        "Spacing errors in a vocabulary entry",
        ["missing space after a comma", "two spaces in the middle of a sentence",
         "no space after a period"],
        "snatch,(snatching)",
        "snatch, (snatching)"),
    ("Vocab", "Too Hard", None): (
        "Definitions written above the level of the book",
        ["a gloss harder than the word it defines",
         "technical or grammar wording (“past participle of…”)",
         "a second sense the book never uses",
         "the same thing said twice"],
        "lemming — a small, short-tailed rodent found in cold northern regions",
        "lemming — a small animal like a mouse that lives in cold places"),
    ("Vocab", "Wrong Sense", None): (
        "The definition is right for some other meaning of the word, not the one this "
        "book uses",
        ["the everyday sense given where the book uses a technical one",
         "the wrong sense listed first",
         "a homograph defined instead of the word in the story"],
        "hold on — wait a moment",
        "hold on — to keep holding something tightly"),
    ("Vocab", "Other", "W05"): (
        "The definition box only repeats the word, so the entry gives no meaning",
        ["the word copied into the definition field",
         "the part of speech also left at the default “noun”"],
        "lawn mower(s) — lawn mower(s)",
        "lawn mower(s) — a machine that cuts grass"),
    ("Vocab", "Other", "W15"): (
        "Factually wrong or mismatched content in a vocabulary entry",
        ["a definition that contradicts the book",
         "an example from the wrong context",
         "the same word listed twice in one book",
         "a word tile that does not match its story sentence"],
        "alive — living; not dead",
        "alive — having life; able to grow, move and breathe"),

    ("Vocab", "Spelling", None): (
        "The word tile itself is misspelled",
        ["a letter missing or wrong in the word students see"],
        "aike",
        "alike"),
    ("Vocab", "Wrong Entry", None): (
        "The entry is for a different word or sense than the book uses",
        ["the tag and definition belong to a word spelled the same way",
         "the entry describes a sense the story never uses"],
        "bark(ing), verb — to make the sharp, loud sound of a dog",
        "bark, noun — the hard outer covering of a tree"),

    # ---- Book word list -------------------------------------------------------------
    ("Vocab", "Story Sentence", None): (
        "Faults in the story sentence on the book's word list (students never see this "
        "field, so these are lower priority than the rest)",
        ["two sentences run together with no space after the period",
         "a wrong or missing word", "no end punctuation",
         "an incomplete or spliced sentence", "a line repeated"],
        "Ants are insects. Its body has three parts.",
        "Ants are insects. Their bodies have three parts."),

    # ---- Text Multiple Choice -------------------------------------------------------
    ("TMC", "Awkward Phrasing", None): (
        "Questions or options written in unnatural English",
        ["unnatural collocations", "wording no English speaker would use",
         "an option phrased unlike the other three"],
        "Police office",
        "Police station"),
    ("TMC", "Capitalization", None): (
        "Capitalization inconsistent across a question set",
        ["some options capitalized and others not",
         "a proper noun left lowercase"],
        "mom's handbag",
        "Mom's handbag"),
    ("TMC", "Grammar", None): (
        "Grammar errors in a question or its options",
        ["wrong word form or preposition",
         "tense disagreement between the question and its options",
         "wrong or missing article", "subject-verb or singular/plural disagreement",
         "an option that does not fit the stem"],
        "On their faces.",
        "Using their faces."),
    ("TMC", "Punctuation", None): (
        "Punctuation inconsistent or missing across the four options",
        ["some options ending in a period and others not",
         "missing end punctuation", "quotation marks missing around a quoted phrase"],
        "A. baking cookies.  B. reading a book  C. playing outside  D. singing",
        "A. baking cookies  B. reading a book  C. playing outside  D. singing"),
    ("TMC", "Spacing", None): (
        "Spacing errors in a question or option",
        ["missing space between words", "double space",
         "space before punctuation"],
        "He feltangry.",
        "He felt angry."),
    ("TMC", "Too Hard", None): (
        "Questions or options using words above the level of the book",
        ["above-level vocabulary in the stem",
         "above-level vocabulary in an option",
         "a word the book never teaches"],
        "Its geography",
        "The land and water around it"),
    ("TMC", "Unclear", None): (
        "The question or its answer key is ambiguous",
        ["more than one option is correct",
         "the answer cannot be settled from the story",
         "the key is not the best answer"],
        "Sight and smell",
        "Sight and strength"),
    ("TMC", "Other", "T01"): (
        "A fill-in-the-blank question with no blank in it",
        [],
        "A map is a picture that shows where places are ?",
        "A map is a picture that shows where places are __________."),
    ("TMC", "Other", "T12"): (
        "Two options mean the same thing, so neither can be the single right answer",
        ["two options worded differently but identical in meaning"],
        "A. It rained a little.  B. It did not rain at all.",
        "A. It rained a little.  B. It rained all day."),
    ("TMC", "Other", "T13"): (
        "The wrong content, or none, was pasted into a field",
        ["an answer option sitting in the question field",
         "text from a different book", "an editorial note left in",
         "an empty field"],
        "\"WE DON'T NEED YOU\"",
        "What did Joe see by the lake in the park?"),
    ("TMC", "Other", "T14"): (
        "Misspellings and hyphenation errors",
        ["a misspelled word", "a missing hyphen in a compound modifier"],
        "Factory made sweaters are prettier.",
        "Factory-made sweaters are prettier."),
    ("TMC", "Other", "T15"): (
        "The question or an option is factually or logically wrong",
        ["an option that cannot answer the question asked",
         "a statement the book contradicts", "a name or number that is wrong"],
        "Bob is a T.rex.",
        "Bob is a T-rex."),

    # ---- Listen & Read --------------------------------------------------------------
    ("LR", "Minor TTS Cutoff", None): (
        "The recorded audio is clipped slightly at the end of the sentence",
        ["the last word cut short", "the sentence ending abruptly"],
        "One end is high.",
        "Re-record the audio for this sentence."),
    ("LR", "Severe TTS Cutoff", None): (
        "The recorded audio stops well before the end of the sentence",
        ["several words missing from the end",
         "the audio stopping mid-sentence"],
        "Tortoise was almost home.",
        "Re-record the audio for this sentence."),
    ("LR", "Start TTS Cutoff", None): (
        "The recorded audio starts late, so the beginning of the sentence is missing",
        ["the first word clipped", "the audio starting mid-word"],
        "Joy and learning are important.",
        "Adjust stop-start times for the audio."),
    ("LR", "TTS Pronunciation", None): (
        "The recorded audio mispronounces or distorts a word",
        ["a word read with the wrong sound",
         "a name read as though it were a common word",
         "distorted or unnatural delivery"],
        "The woman uses a simple machine.",
        "Re-record the audio for this sentence."),
    ("LR", "Text", None): (
        "A mistake in the printed text on the page",
        ["a misspelled or missing word", "missing punctuation",
         "a stray character left in", "a font that is hard to read"],
        "Tank goodness! Tere you are, Elly.",
        "Thank goodness! There you are, Elly."),
    ("LR", "Text-TTS Mismatch", None): (
        "The recorded audio does not say what the page shows",
        ["a word read in the singular where the page has a plural",
         "an article read that the page does not have"],
        "These bones are parts of a dinosaur.",
        "These bones are part of a dinosaur."),
    ("LR", "Other", "L07"): (
        "Other faults noticed while listening and reading",
        ["the whole page highlighted at once",
         "highlighting out of step with the audio",
         "unrelated audio playing"],
        "Who it will be?",
        "Check this page and correct the problem described."),
    ("LRA", "Severe TTS Cutoff", None): (
        "The recorded audio stops well before the end of the sentence",
        ["several words missing from the end",
         "the audio stopping mid-sentence"],
        "Fuzzy, flying mammal",
        "Adjust stop-start times for the audio."),
    ("LRA", "Other", "A07"): (
        "Other faults noticed while listening and reading along",
        ["the audio continuing after the highlighted text ends",
         "highlighting out of step with the audio",
         "unrelated audio playing"],
        "you are called a pinky\"",
        "Adjust stop-start times for the audio."),
}
