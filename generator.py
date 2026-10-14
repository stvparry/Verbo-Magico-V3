from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.comments import Comment


NAVY = "17365D"
BLUE = "2F75B5"
LIGHT = "D9EAF7"
PALE = "EEF5FB"
WHITE = "FFFFFF"
GREY = "E7E6E6"


DISPLAY_PERSONS = [
    ("yo", ["yo"]),
    ("tú", ["tú"]),
    ("él / ella / usted", ["él", "ella", "usted"]),
    ("nosotros / nosotras", ["nosotros", "nosotras"]),
    ("vosotros / vosotras", ["vosotros", "vosotras"]),
    ("ellos / ellas / ustedes", ["ellos", "ellas", "ustedes"]),
]


CORE_TENSES = [
    ("indicativo", "presente"),
    ("indicativo", "pretérito-perfecto-simple"),
    ("indicativo", "pretérito-imperfecto"),
    ("indicativo", "futuro"),
    ("condicional", "presente"),
]


TENSE_NAMES = {
    ("indicativo", "presente"): "Presente",
    ("indicativo", "pretérito-perfecto-simple"): "Pretérito perfecto simple",
    ("indicativo", "pretérito-imperfecto"): "Pretérito imperfecto",
    ("indicativo", "futuro"): "Futuro simple",
    ("condicional", "presente"): "Condicional simple",
}


def normalise(payload):

    if "value" not in payload:
        raise ValueError(
            f"Expected 'value' key. Received: {list(payload.keys())}"
        )

    value = payload["value"]

    if "moods" not in value:
        raise ValueError(
            f"Expected 'moods' key inside value. Received: {list(value.keys())}"
        )

    return value


def strip_pronoun(text):

    pronouns = [
        "yo",
        "tú",
        "vos",
        "él",
        "ella",
        "usted",
        "nosotros",
        "nosotras",
        "vosotros",
        "vosotras",
        "ellos",
        "ellas",
        "ustedes",
    ]

    text = text.strip()

    for p in sorted(pronouns, key=len, reverse=True):

        prefix = p + " "

        if text.lower().startswith(prefix):
            return text[len(prefix):]

    return text


def get_forms(data, mood, tense):

    entries = data["moods"][mood][tense]

    by_pronoun = {}

    for e in entries:

        pronoun = e.get("pr")

        if not pronoun:
            continue

        by_pronoun[pronoun] = strip_pronoun(
