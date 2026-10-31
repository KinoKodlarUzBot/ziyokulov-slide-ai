"""Word (.docx) hujjatni tarjima qilish uchun yordamchi funksiyalar.

Bu fayl main.py bilan bir papkada turishi kerak.
Bu yerda Gemini yoki Telegram yo'q: faqat hujjatni o'qish, tilni aniqlash
va tarjima matnini formatni buzmasdan joyiga qo'yish.
"""
import re
from dataclasses import dataclass

from docx.oxml.ns import qn
from lxml import etree

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


@dataclass
class Segment:
    """Hujjatdagi bitta tarjima qilinadigan bo'lak (odatda bitta abzats)."""
    t_elements: list      # shu bo'lakdagi <w:t> elementlari
    text: str             # asl matn


def detect_lang(text: str) -> str:
    """Matn tilini aniqlaydi: 'uz', 'ru' yoki 'en'."""
    sample = text[:4000]
    letters = [c for c in sample if c.isalpha()]
    if not letters:
        return "uz"

    cyr = sum("\u0400" <= c <= "\u04ff" for c in letters)
    low = sample.lower()

    if cyr / len(letters) > 0.5:
        uz_marks = sum(low.count(c) for c in "ўқғҳ")
        ru_marks = sum(low.count(c) for c in "ыщь")
        return "uz" if uz_marks > ru_marks else "ru"

    padded = " " + re.sub(r"[^a-z'ʻʼ‘’` ]", " ", low) + " "
    en_score = sum(padded.count(w) for w in
                   (" the ", " and ", " of ", " is ", " to ", " in ", " that ", " with "))
    uz_score = sum(padded.count(w) for w in
                   (" va ", " bilan ", " uchun ", " bu ", " ham ", " bir ", " lar ", " ning "))
    uz_score += sum(low.count(m) for m in ("o'", "g'", "oʻ", "gʻ", "o‘", "g‘", "o’", "g’"))
    return "en" if en_score > uz_score else "uz"


def _run_key(t_el) -> bytes:
    run = t_el.getparent()
    rpr = run.find(qn("w:rPr")) if run is not None else None
    return etree.tostring(rpr) if rpr is not None else b""


def _split_segments(t_elements: list) -> list[Segment]:
    """Abzatsni bo'laklarga ajratadi.

    Odatda butun abzats bitta bo'lak bo'ladi (jumla to'liq tarjima qilinadi).
    Faqat "Kalit so'zlar:" kabi qisqa qalin yorliq + oddiy matn bo'lsa,
    ikkiga bo'linadi, shunda yorliq qalin, matn oddiy qoladi.
    """
    groups: list[list] = []
    last_key = None
    for t in t_elements:
        key = _run_key(t)
        if groups and key == last_key:
            groups[-1].append(t)
        else:
            groups.append([t])
            last_key = key

    def text_of(els):
        return "".join(e.text or "" for e in els)

    if len(groups) == 2:
        first = text_of(groups[0])
        if first.strip().endswith(":") and len(first) <= 40:
            return [Segment(groups[0], first), Segment(groups[1], text_of(groups[1]))]

    return [Segment(t_elements, text_of(t_elements))]


def _paragraph_segments(p_el) -> list[Segment]:
    t_elements = [t for t in p_el.iter(qn("w:t"))]
    if not t_elements:
        return []
    segments = _split_segments(t_elements)
    return [s for s in segments if any(ch.isalpha() for ch in s.text)]


def collect_segments(doc, include_headers: bool = True) -> list[Segment]:
    """Hujjatdagi barcha matn bo'laklarini tartib bilan yig'adi
    (asosiy matn, jadvallar, matn qutilari, xohlansa - yuqori/pastki kolontitul)."""
    segments: list[Segment] = []
    for p_el in doc.element.body.iter(qn("w:p")):
        segments.extend(_paragraph_segments(p_el))

    if include_headers:
        for section in doc.sections:
            parts = (section.header, section.footer,
                     section.first_page_header, section.first_page_footer)
            for part in parts:
                try:
                    if part.is_linked_to_previous:
                        continue
                    for p_el in part._element.iter(qn("w:p")):
                        segments.extend(_paragraph_segments(p_el))
                except Exception:
                    continue
    return segments


def apply_translation(segment: Segment, new_text: str) -> None:
    """Tarjimani birinchi run'ga yozadi (uning formati saqlanadi), qolganlarini bo'shatadi."""
    old = segment.text
    lead = old[: len(old) - len(old.lstrip())]
    trail = old[len(old.rstrip()):]
    final = lead + new_text.strip() + trail

    first, rest = segment.t_elements[0], segment.t_elements[1:]
    first.text = final
    first.set(XML_SPACE, "preserve")
    for t in rest:
        t.text = ""
