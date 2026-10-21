"""Krossvord katagini avtomatik joylashtirish.

Gemini so'z va ta'riflarni beradi, bu modul esa so'zlarni o'zaro kesishtirib
katakka joylaydi. Lotin yozuvida o', g', sh, ch bitta katakka sig'adi.
"""
import random
import re
from dataclasses import dataclass

_LATIN_TOKEN = re.compile(r"o'|g'|sh|ch|[a-z]")
_CYRILLIC_OK = re.compile(r"^[А-ЯЁЎҚҒҲ]+$")
_APOS = str.maketrans({ch: "'" for ch in "’ʻʼ‘`´"})

MIN_WORD_LEN = 3
MAX_GRID = 15


def tokenize(word: str, alphabet: str) -> list[str] | None:
    """So'zni katak harflariga ajratadi. Yaroqsiz so'z uchun None qaytaradi."""
    word = (word or "").strip().translate(_APOS)
    if not word or " " in word or "-" in word:
        return None
    if alphabet == "kirill":
        upper = word.upper()
        if not _CYRILLIC_OK.match(upper):
            return None
        return list(upper)
    low = word.lower()
    tokens = _LATIN_TOKEN.findall(low)
    if "".join(tokens) != low:
        return None
    return [t.upper() for t in tokens]


@dataclass
class Item:
    word: str
    clue: str
    tokens: list


def _can_place(grid, tokens, row, col, horizontal):
    """So'zni (row, col) dan boshlab qo'yish mumkinmi? Kesishishlar soni yoki None."""
    dr, dc = (0, 1) if horizontal else (1, 0)
    pr, pc = (1, 0) if horizontal else (0, 1)  # perpendikulyar yo'nalish

    # so'zdan oldin va keyin katak bo'sh bo'lishi kerak
    if (row - dr, col - dc) in grid:
        return None
    if (row + dr * len(tokens), col + dc * len(tokens)) in grid:
        return None

    crossings = 0
    previous_crossing = False
    for i, token in enumerate(tokens):
        r, c = row + dr * i, col + dc * i
        existing = grid.get((r, c))
        if existing is not None:
            if existing != token:
                return None
            if previous_crossing:  # ketma-ket kesishish = parallel ustma-ust tushish
                return None
            crossings += 1
            previous_crossing = True
        else:
            previous_crossing = False
            # bo'sh katakning yon qo'shnilari band bo'lmasligi kerak (tasodifiy so'z hosil bo'lmasin)
            if (r + pr, c + pc) in grid or (r - pr, c - pc) in grid:
                return None
    return crossings


def _bbox(cells):
    rows = [r for r, _ in cells]
    cols = [c for _, c in cells]
    return min(rows), max(rows), min(cols), max(cols)


def _layout(order, rng):
    grid: dict = {}
    placed = []  # (item, row, col, horizontal)

    for item in order:
        tokens = item.tokens
        if not placed:
            for i, token in enumerate(tokens):
                grid[(0, i)] = token
            placed.append((item, 0, 0, True))
            continue

        best = None  # (crossings, -area, jitter, row, col, horizontal)
        for other, orow, ocol, ohoriz in placed:
            horizontal = not ohoriz
            for oi, otoken in enumerate(other.tokens):
                for ti, token in enumerate(tokens):
                    if token != otoken:
                        continue
                    if ohoriz:
                        row, col = orow - ti, ocol + oi
                    else:
                        row, col = orow + oi, ocol - ti
                    crossings = _can_place(grid, tokens, row, col, horizontal)
                    if not crossings:
                        continue
                    cells = list(grid) + [
                        (row + (0 if horizontal else k), col + (k if horizontal else 0))
                        for k in range(len(tokens))
                    ]
                    r0, r1, c0, c1 = _bbox(cells)
                    height, width = r1 - r0 + 1, c1 - c0 + 1
                    if height > MAX_GRID or width > MAX_GRID:
                        continue
                    score = (crossings, -(height * width), rng.random(), row, col, horizontal)
                    if best is None or score[:3] > best[:3]:
                        best = score
        if best is None:
            continue
        _, _, _, row, col, horizontal = best
        for i, token in enumerate(tokens):
            grid[(row + (0 if horizontal else i), col + (i if horizontal else 0))] = token
        placed.append((item, row, col, horizontal))
    return grid, placed


def build_crossword(entries, alphabet: str, attempts: int = 80, seed: int = 2026):
    """entries: [(so'z, ta'rif), ...]. Natija: katak, raqamlar va savollar ro'yxati yoki None."""
    items, seen = [], set()
    for word, clue in entries:
        tokens = tokenize(word, alphabet)
        clue = (clue or "").strip()
        if not tokens or not clue or len(tokens) < MIN_WORD_LEN or len(tokens) > MAX_GRID:
            continue
        key = "".join(tokens)
        if key in seen:
            continue
        seen.add(key)
        items.append(Item(word="".join(tokens), clue=clue, tokens=tokens))
    if len(items) < 5:
        return None

    rng = random.Random(seed)
    best_grid, best_placed, best_key = None, None, None
    for attempt in range(attempts):
        order = sorted(items, key=lambda it: -len(it.tokens))
        if attempt:
            # eng uzun so'zni boshida qoldirib, qolganini aralashtiramiz
            head, tail = order[:1], order[1:]
            rng.shuffle(tail)
            if attempt % 2:
                tail.sort(key=lambda it: -len(it.tokens) + rng.random() * 3)
            order = head + tail
        grid, placed = _layout(order, rng)
        r0, r1, c0, c1 = _bbox(grid)
        key = (len(placed), -((r1 - r0 + 1) * (c1 - c0 + 1)))
        if best_key is None or key > best_key:
            best_grid, best_placed, best_key = grid, placed, key
        if len(placed) == len(items) and attempt >= 10:
            break

    if best_grid is None or len(best_placed) < 5:
        return None

    r0, _, c0, _ = _bbox(best_grid)
    cells = {(r - r0, c - c0): t for (r, c), t in best_grid.items()}
    rows = max(r for r, _ in cells) + 1
    cols = max(c for _, c in cells) + 1

    starts = sorted({(row - r0, col - c0) for _, row, col, _ in best_placed})
    numbers = {pos: n for n, pos in enumerate(starts, start=1)}

    across, down = [], []
    for item, row, col, horizontal in best_placed:
        pos = (row - r0, col - c0)
        entry = {"number": numbers[pos], "word": item.word, "clue": item.clue,
                 "row": pos[0], "col": pos[1], "length": len(item.tokens)}
        (across if horizontal else down).append(entry)
    across.sort(key=lambda e: e["number"])
    down.sort(key=lambda e: e["number"])

    return {"rows": rows, "cols": cols, "cells": cells, "numbers": numbers,
            "across": across, "down": down, "placed": len(best_placed), "total": len(items)}


def validate(result) -> list[str]:
    """Sinov uchun: katakdagi har bir harf ketma-ketligi ro'yxatdagi so'z ekanini tekshiradi."""
    problems = []
    cells = result["cells"]
    words = {("A", e["row"], e["col"]): e for e in result["across"]}
    words.update({("D", e["row"], e["col"]): e for e in result["down"]})

    for r in range(result["rows"]):
        for c in range(result["cols"]):
            if (r, c) not in cells:
                continue
            if (r, c - 1) not in cells and (r, c + 1) in cells and ("A", r, c) not in words:
                problems.append(f"gorizontal so'z ro'yxatda yo'q: {(r, c)}")
            if (r - 1, c) not in cells and (r + 1, c) in cells and ("D", r, c) not in words:
                problems.append(f"vertikal so'z ro'yxatda yo'q: {(r, c)}")

    for (kind, r, c), e in words.items():
        dr, dc = (0, 1) if kind == "A" else (1, 0)
        letters = [cells.get((r + dr * i, c + dc * i)) for i in range(e["length"])]
        if None in letters:
            problems.append(f"{e['word']}: katakda bo'sh joy bor")
            continue
        tokens = tokenize(e["word"], "kirill" if re.match(r"^[А-ЯЁЎҚҒҲ]+$", e["word"]) else "latin")
        if letters != tokens:
            problems.append(f"{e['word']}: harflar mos emas")
        if (r - dr, c - dc) in cells or (r + dr * e["length"], c + dc * e["length"]) in cells:
            problems.append(f"{e['word']}: so'z uzayib ketgan")
    return problems
