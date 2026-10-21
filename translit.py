"""O'zbek lotin yozuvidan kirill yozuviga konvertor va alifbo yordamchilari.

Asosiy matnni Gemini tanlangan alifboda o'zi yozadi. Bu modul:
- doimiy yozuvlarni (sarlavhalar: "Reja", "Xulosa" va h.k.) kirillga o'tkazadi;
- Gemini lotinda yozib yuborsa, zaxira sifatida matnni kirillga o'tkazadi.
"""
import re

LATIN = "latin"
KIRILL = "kirill"
ALPHABETS = (LATIN, KIRILL)

APOSTROPHES = "'’ʻʼ‘`´"
_WORD_RE = re.compile("[A-Za-z" + re.escape(APOSTROPHES) + "]+")
_APOS_TABLE = {ord(ch): "'" for ch in APOSTROPHES}

_SINGLE = {
    "a": "а", "b": "б", "d": "д", "f": "ф", "g": "г", "h": "ҳ", "i": "и", "j": "ж",
    "k": "к", "l": "л", "m": "м", "n": "н", "o": "о", "p": "п", "q": "қ", "r": "р",
    "s": "с", "t": "т", "u": "у", "v": "в", "x": "х", "y": "й", "z": "з",
    "w": "в", "c": "ц", "e": "е",
}
_CYR_VOWELS = set("аеэиоуўёюя")
_CYRILLIC_LETTERS = "А-Яа-яЁёЎўҚқҒғҲҳ"
_CYR_RE = re.compile("[" + _CYRILLIC_LETTERS + "]")
_LAT_RE = re.compile("[A-Za-z]")


def _case(src: str, cyr: str) -> str:
    """Manba harflarining registriga qarab kirill harf(lar)ni moslaydi."""
    if len(src) > 1 and src.isupper():
        return cyr.upper()
    if src[:1].isupper():
        return cyr[:1].upper() + cyr[1:]
    return cyr


def _convert_word(word: str) -> str:
    low = word.lower()
    n = len(word)
    out: list[str] = []
    prev_vowel = False
    i = 0
    while i < n:
        two = low[i:i + 2]
        ch = low[i]

        if two == "o'":
            piece, step = _case(word[i:i + 2], "ў"), 2
        elif two == "g'":
            piece, step = _case(word[i:i + 2], "ғ"), 2
        elif two == "sh":
            piece, step = _case(word[i:i + 2], "ш"), 2
        elif two == "ch":
            piece, step = _case(word[i:i + 2], "ч"), 2
        elif two == "ts":
            piece, step = _case(word[i:i + 2], "ц"), 2
        elif two == "yo" and low[i + 2:i + 3] != "'":
            piece, step = _case(word[i:i + 2], "ё"), 2
        elif two == "yu":
            piece, step = _case(word[i:i + 2], "ю"), 2
        elif two == "ya":
            piece, step = _case(word[i:i + 2], "я"), 2
        elif two == "ye":
            piece, step = _case(word[i:i + 2], "е"), 2
        elif ch == "'":
            prev_upper = i > 0 and word[i - 1].isupper()
            next_upper = i + 1 >= n or word[i + 1].isupper()
            piece, step = ("Ъ" if prev_upper and next_upper else "ъ"), 1
        elif ch == "e":
            cyr = "э" if (i == 0 or prev_vowel) else "е"
            piece, step = _case(word[i], cyr), 1
        else:
            cyr = _SINGLE.get(ch)
            if cyr is None:
                piece, step = word[i], 1
            else:
                piece, step = _case(word[i], cyr), 1

        out.append(piece)
        prev_vowel = piece.lower() in _CYR_VOWELS
        i += step
    return "".join(out)


def latin_to_cyrillic(text: str) -> str:
    """Lotin yozuvidagi o'zbekcha matnni kirillga o'tkazadi (raqam va belgilar o'zgarmaydi)."""
    if not text:
        return text
    text = text.translate(_APOS_TABLE)
    return _WORD_RE.sub(lambda m: _convert_word(m.group(0)), text)


def cyrillic_ratio(text: str) -> float:
    """Matndagi harflarning qancha qismi kirill ekanini qaytaradi (harf bo'lmasa 1.0)."""
    cyr = len(_CYR_RE.findall(text or ""))
    lat = len(_LAT_RE.findall(text or ""))
    total = cyr + lat
    return 1.0 if total == 0 else cyr / total


def normalize_alphabet(value) -> str:
    return value if value in ALPHABETS else LATIN


def fix_text(text: str, alphabet: str) -> str:
    """Tanlangan alifbo kirill bo'lsa-yu matn lotinda qolgan bo'lsa, uni kirillga o'tkazadi."""
    if alphabet == KIRILL and text and cyrillic_ratio(text) < 0.5:
        return latin_to_cyrillic(text)
    return text


def ui(text: str, alphabet: str) -> str:
    """Doimiy (dasturdagi) yozuvlarni tanlangan alifboga o'tkazadi."""
    return latin_to_cyrillic(text) if alphabet == KIRILL else text


def alphabet_rule(alphabet: str) -> str:
    """Gemini ko'rsatmasiga qo'shiladigan satr (kirill uchun)."""
    if alphabet == KIRILL:
        return (
            "\n- Butun matnni o'zbek tilining KIRILL alifbosida (ўзбек кирилл ёзуви) yozing: "
            "ў, ғ, қ, ҳ harflaridan to'g'ri foydalaning. JSON kalitlari o'zgarmasin "
            "(title, sections, content kabi kalitlar lotincha qoladi)."
        )
    return ""
