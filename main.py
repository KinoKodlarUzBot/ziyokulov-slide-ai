import asyncio
import hashlib
import hmac
import html
import json
import logging
import math
import os
import re
import sqlite3
import tempfile
import threading
import time
import uuid
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Literal
from urllib.parse import parse_qsl

from aiogram import Bot, Dispatcher, F, types
from aiogram.exceptions import TelegramForbiddenError
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import (
    BotCommand,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    MenuButtonWebApp,
    ReplyKeyboardMarkup,
    WebAppInfo,
)
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from google import genai
from google.genai import types as genai_types
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm
from docx.shared import Pt as DocPt
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt
from pydantic import BaseModel, ConfigDict, Field

from crossword import build_crossword
from translit import (
    ALPHABETS,
    KIRILL,
    LATIN,
    alphabet_rule,
    fix_text,
    normalize_alphabet,
    ui,
)


# =========================================================
# CONFIGURATION
# Railway Variables: BOT_TOKEN, GEMINI_API_KEY, WEBAPP_URL
# Ixtiyoriy: RATE_LIMIT_DB (Railway Volume yo'li, masalan /data/rate_limit.db)
# =========================================================
logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO"),
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)
logger = logging.getLogger("slide_ai")

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
WEBAPP_URL = os.getenv(
    "WEBAPP_URL",
    "https://ziyokulov-slide-ai-production.up.railway.app",
).strip().rstrip("/")
PORT = int(os.getenv("PORT", "8080"))
# Asosiy model va zaxira modellar. Asosiysi topilmasa (404) yoki limiti tugasa (429),
# keyingisi sinab ko'riladi. Railway Variables orqali o'zgartirish mumkin:
# GEMINI_MODEL, GEMINI_FALLBACK_MODELS (vergul bilan ajratilgan).
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.1-flash-lite").strip()
GEMINI_FALLBACK_MODELS = [
    m.strip()
    for m in os.getenv("GEMINI_FALLBACK_MODELS", "gemini-3.6-flash,gemini-2.5-flash").split(",")
    if m.strip()
]
GEMINI_MODELS = [GEMINI_MODEL] + [m for m in GEMINI_FALLBACK_MODELS if m != GEMINI_MODEL]
MAX_CONCURRENT_GENERATIONS = int(os.getenv("MAX_CONCURRENT_GENERATIONS", "3"))
RATE_LIMIT_REQUESTS = int(os.getenv("RATE_LIMIT_REQUESTS", "5"))
RATE_LIMIT_WINDOW_SECONDS = int(os.getenv("RATE_LIMIT_WINDOW_SECONDS", "3600"))
TELEGRAM_INIT_DATA_MAX_AGE = int(os.getenv("TELEGRAM_INIT_DATA_MAX_AGE", "86400"))
GEMINI_TIMEOUT_SECONDS = float(os.getenv("GEMINI_TIMEOUT_SECONDS", "60"))
GEMINI_DOC_TIMEOUT_SECONDS = float(os.getenv("GEMINI_DOC_TIMEOUT_SECONDS", "120"))
QUEUE_TIMEOUT_SECONDS = float(os.getenv("QUEUE_TIMEOUT_SECONDS", "45"))
GEMINI_MAX_ATTEMPTS = int(os.getenv("GEMINI_MAX_ATTEMPTS", "2"))
ENABLE_BOT_POLLING = os.getenv("ENABLE_BOT_POLLING", "1").strip() not in {"0", "false", "False"}
# Adminlar (Telegram ID, vergul bilan): /stats buyrug'i va xato ogohlantirishlari shularga boradi.
ADMIN_IDS = {
    int(x) for x in os.getenv("ADMIN_IDS", "").replace(" ", "").split(",") if x.isdigit()
}
# Foydalanuvchilar murojaat qiladigan kontakt (masalan @username), ixtiyoriy.
SUPPORT_CONTACT = os.getenv("SUPPORT_CONTACT", "").strip()

MAX_TOPIC_LENGTH = 200
MAX_BULLETS = 5
MAX_BULLET_CHARS = 200
MAX_TITLE_CHARS = 100
SLIDE_COUNT = 5
MAX_DOC_SECTIONS = 10
MAX_DOC_PARAGRAPHS = 8
MIN_SLIDES = 5
MAX_SLIDES = 20
# Bir bet taxminan nechta so'z (Times New Roman 14, 1.5 oraliq). Haqiqiy natijaga qarab
# Railway Variables orqali o'zgartiring: bet kam chiqsa kamaytiring, ko'p chiqsa oshiring.
WORDS_PER_PAGE = int(os.getenv("WORDS_PER_PAGE", "230"))
PARAGRAPH_WORDS = 115
_TEXT_TIERS = [(5, 10), (10, 15), (15, 20), (20, 25), (25, 30)]
PAGE_TIERS = {
    "referat": _TEXT_TIERS,
    "mustaqil": _TEXT_TIERS,
    "kurs": [(10, 15), (15, 20), (20, 25), (25, 30), (30, 35), (35, 40), (40, 45)],
}

if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN Railway Variables bo'limida sozlanmagan.")
if not GEMINI_API_KEY:
    raise RuntimeError("GEMINI_API_KEY Railway Variables bo'limida sozlanmagan.")

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
STATIC_DIR.mkdir(parents=True, exist_ok=True)
INDEX_FILE = STATIC_DIR / "index.html"
RATE_LIMIT_DB = os.getenv("RATE_LIMIT_DB", str(BASE_DIR / "data" / "rate_limit.db"))

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
gemini_client = genai.Client(api_key=GEMINI_API_KEY)
generation_semaphore = asyncio.Semaphore(MAX_CONCURRENT_GENERATIONS)

def _c(r: int, g: int, b: int) -> RGBColor:
    return RGBColor(r, g, b)


_DARK_TEXT = _c(15, 23, 42)
_LIGHT = _c(255, 255, 255)

# 15 ta shablon. Har biri o'z rang, shrift, bezak, titul sahifasi va ichki joylashuvlar
# ketma-ketligiga ega. Yangi shablon qo'shish uchun shu ro'yxatga yana bitta yozuv qo'shing.
THEMES = {
    "classic": {
        "name": "Klassik", "bg": _c(255, 255, 255), "title": _c(20, 20, 20), "text": _c(55, 65, 81),
        "accent": _c(79, 70, 229), "accent2": _c(167, 139, 250), "card": _c(243, 244, 246), "on_accent": _LIGHT,
        "fh": "Calibri", "fb": "Calibri", "cw": 0.52, "cover": "panel", "decor": "topbar", "mark": "bar",
        "layouts": ["cards", "split", "grid", "rows", "columns"],
    },
    "business": {
        "name": "Biznes", "bg": _c(240, 244, 248), "title": _c(30, 58, 138), "text": _c(31, 41, 55),
        "accent": _c(30, 58, 138), "accent2": _c(59, 130, 246), "card": _c(255, 255, 255), "on_accent": _LIGHT,
        "fh": "Arial", "fb": "Arial", "cw": 0.55, "cover": "band", "decor": "sidebar", "mark": "line",
        "layouts": ["rows", "grid", "timeline", "split", "cards"],
    },
    "dark": {
        "name": "Qorong'i", "bg": _c(24, 24, 27), "title": _c(255, 255, 255), "text": _c(228, 228, 231),
        "accent": _c(139, 92, 246), "accent2": _c(236, 72, 153), "card": _c(39, 39, 42), "on_accent": _LIGHT,
        "fh": "Segoe UI", "fb": "Segoe UI", "cw": 0.54, "cover": "center", "decor": "glow", "mark": "dot",
        "layouts": ["bignum", "cards", "grid", "columns", "timeline"],
    },
    "minimal": {
        "name": "Yashil minimal", "bg": _c(240, 253, 244), "title": _c(22, 101, 52), "text": _c(20, 83, 45),
        "accent": _c(22, 163, 74), "accent2": _c(134, 239, 172), "card": _c(255, 255, 255), "on_accent": _LIGHT,
        "fh": "Georgia", "fb": "Calibri", "cw": 0.52, "cover": "line", "decor": "rule", "mark": "line",
        "layouts": ["rows", "columns", "split", "bignum", "timeline"],
    },
    "ocean": {
        "name": "Okean", "bg": _c(236, 248, 255), "title": _c(7, 64, 99), "text": _c(30, 64, 86),
        "accent": _c(2, 132, 199), "accent2": _c(34, 211, 238), "card": _c(255, 255, 255), "on_accent": _LIGHT,
        "fh": "Trebuchet MS", "fb": "Calibri", "cw": 0.53, "cover": "circles", "decor": "glow", "mark": "bar",
        "layouts": ["split", "timeline", "cards", "grid", "bignum"],
    },
    "sunset": {
        "name": "Shafaq", "bg": _c(255, 247, 237), "title": _c(124, 45, 18), "text": _c(87, 56, 40),
        "accent": _c(234, 88, 12), "accent2": _c(244, 63, 94), "card": _c(255, 255, 255), "on_accent": _LIGHT,
        "fh": "Georgia", "fb": "Calibri", "cw": 0.54, "cover": "diagonal", "decor": "stripes", "mark": "bar",
        "layouts": ["columns", "cards", "bignum", "grid", "rows"],
    },
    "forest": {
        "name": "O'rmon", "bg": _c(20, 50, 40), "title": _c(236, 253, 245), "text": _c(209, 250, 229),
        "accent": _c(52, 211, 153), "accent2": _c(250, 204, 21), "card": _c(30, 72, 58), "on_accent": _c(6, 40, 30),
        "fh": "Cambria", "fb": "Calibri", "cw": 0.53, "cover": "frame", "decor": "frame", "mark": "dot",
        "layouts": ["grid", "rows", "split", "timeline", "cards"],
    },
    "royal": {
        "name": "Qirollik", "bg": _c(253, 248, 240), "title": _c(127, 29, 29), "text": _c(68, 64, 60),
        "accent": _c(146, 64, 14), "accent2": _c(217, 119, 6), "card": _c(255, 255, 255), "on_accent": _LIGHT,
        "fh": "Georgia", "fb": "Georgia", "cw": 0.57, "cover": "frame", "decor": "frame", "mark": "line",
        "layouts": ["rows", "bignum", "columns", "cards", "split"],
    },
    "coral": {
        "name": "Marjon", "bg": _c(255, 241, 242), "title": _c(136, 19, 55), "text": _c(76, 29, 49),
        "accent": _c(244, 63, 94), "accent2": _c(251, 146, 60), "card": _c(255, 255, 255), "on_accent": _LIGHT,
        "fh": "Trebuchet MS", "fb": "Trebuchet MS", "cw": 0.55, "cover": "circles", "decor": "glow", "mark": "dot",
        "layouts": ["grid", "cards", "timeline", "bignum", "columns"],
    },
    "midnight": {
        "name": "Yarim tun", "bg": _c(15, 23, 42), "title": _c(241, 245, 249), "text": _c(203, 213, 225),
        "accent": _c(56, 189, 248), "accent2": _c(129, 140, 248), "card": _c(30, 41, 59), "on_accent": _c(8, 20, 40),
        "fh": "Segoe UI", "fb": "Segoe UI", "cw": 0.54, "cover": "bottom", "decor": "stripes", "mark": "line",
        "layouts": ["timeline", "grid", "bignum", "cards", "rows"],
    },
    "sand": {
        "name": "Qum", "bg": _c(245, 240, 230), "title": _c(60, 48, 38), "text": _c(87, 73, 58),
        "accent": _c(146, 92, 50), "accent2": _c(180, 140, 90), "card": _c(255, 252, 245), "on_accent": _LIGHT,
        "fh": "Cambria", "fb": "Cambria", "cw": 0.54, "cover": "bottom", "decor": "rule", "mark": "bar",
        "layouts": ["split", "rows", "columns", "grid", "cards"],
    },
    "rose": {
        "name": "Atirgul", "bg": _c(253, 242, 248), "title": _c(131, 24, 67), "text": _c(80, 30, 55),
        "accent": _c(219, 39, 119), "accent2": _c(192, 132, 252), "card": _c(255, 255, 255), "on_accent": _LIGHT,
        "fh": "Georgia", "fb": "Calibri", "cw": 0.54, "cover": "diagonal", "decor": "glow", "mark": "dot",
        "layouts": ["cards", "columns", "split", "bignum", "grid"],
    },
    "mono": {
        "name": "Monoxrom", "bg": _c(255, 255, 255), "title": _c(17, 17, 17), "text": _c(64, 64, 64),
        "accent": _c(17, 17, 17), "accent2": _c(156, 163, 175), "card": _c(243, 244, 246), "on_accent": _LIGHT,
        "fh": "Arial", "fb": "Arial", "cw": 0.55, "cover": "full", "decor": "rule", "mark": "line",
        "layouts": ["rows", "split", "bignum", "timeline", "grid"],
    },
    "neon": {
        "name": "Neon", "bg": _c(10, 10, 22), "title": _c(236, 255, 250), "text": _c(196, 228, 222),
        "accent": _c(34, 255, 170), "accent2": _c(255, 40, 170), "card": _c(22, 22, 44), "on_accent": _c(5, 10, 20),
        "fh": "Segoe UI", "fb": "Segoe UI", "cw": 0.54, "cover": "line", "decor": "stripes", "mark": "bar",
        "layouts": ["bignum", "timeline", "cards", "columns", "grid"],
    },
    "academic": {
        "name": "Akademik", "bg": _c(255, 255, 255), "title": _c(23, 37, 84), "text": _c(30, 41, 59),
        "accent": _c(30, 64, 175), "accent2": _c(180, 83, 9), "card": _c(241, 245, 249), "on_accent": _LIGHT,
        "fh": "Cambria", "fb": "Calibri", "cw": 0.53, "cover": "center", "decor": "topbar", "mark": "line",
        "layouts": ["rows", "cards", "split", "grid", "timeline"],
    },
}
TEMPLATE_STYLES = THEMES  # eski nom bilan ham ishlayveradi


# =========================================================
# REQUEST MODELS
# =========================================================
class SlideRequest(BaseModel):
    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)
    init_data: str = Field(min_length=1, max_length=10_000)
    topic: str = Field(min_length=3, max_length=MAX_TOPIC_LENGTH)
    template: Literal[
        "classic", "business", "dark", "minimal", "ocean", "sunset", "forest", "royal",
        "coral", "midnight", "sand", "rose", "mono", "neon", "academic",
    ] = "classic"
    alphabet: Literal["latin", "kirill"] | None = None
    slides: int = Field(default=SLIDE_COUNT, ge=MIN_SLIDES, le=MAX_SLIDES)


class MeRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    init_data: str = Field(min_length=1, max_length=10_000)


# =========================================================
# TELEGRAM MINI APP AUTHENTICATION
# =========================================================
def verify_telegram_init_data(init_data: str) -> dict:
    try:
        pairs = parse_qsl(init_data, keep_blank_values=True, strict_parsing=True)
        data = dict(pairs)
    except (ValueError, TypeError):
        raise HTTPException(status_code=401, detail="Telegram tasdiqlash ma'lumoti noto'g'ri.")

    received_hash = data.pop("hash", None)
    if not received_hash:
        raise HTTPException(status_code=401, detail="Telegram tasdiqlash hash'i topilmadi.")

    data_check_string = "\n".join(f"{key}={value}" for key, value in sorted(data.items()))
    secret_key = hmac.new(b"WebAppData", BOT_TOKEN.encode("utf-8"), hashlib.sha256).digest()
    calculated_hash = hmac.new(
        secret_key, data_check_string.encode("utf-8"), hashlib.sha256
    ).hexdigest()

    if not hmac.compare_digest(calculated_hash, received_hash):
        raise HTTPException(status_code=401, detail="Telegram ma'lumoti tasdiqlanmadi.")

    try:
        auth_date = int(data.get("auth_date", "0"))
    except ValueError:
        raise HTTPException(status_code=401, detail="auth_date noto'g'ri.")

    now = int(time.time())
    if auth_date <= 0 or auth_date > now + 60 or now - auth_date > TELEGRAM_INIT_DATA_MAX_AGE:
        raise HTTPException(
            status_code=401,
            detail="Telegram sessiyasi eskirgan. Mini App'ni qayta oching.",
        )

    try:
        user = json.loads(data.get("user", "{}"))
        user_id = int(user["id"])
    except (ValueError, TypeError, KeyError, json.JSONDecodeError):
        raise HTTPException(status_code=401, detail="Telegram foydalanuvchisi aniqlanmadi.")

    return {"user_id": user_id, "user": user}


# =========================================================
# RATE LIMITER (SQLite: qayta ishga tushganda ham saqlanadi)
# Eslatma: Railway'da fayl tizimi vaqtinchalik. Limit deploy'dan keyin
# ham saqlanishi uchun Volume ulab, RATE_LIMIT_DB ni shu yo'lga qo'ying.
# =========================================================
class RateLimiter:
    def __init__(self, db_path: str, limit: int, window_seconds: int):
        self.limit = limit
        self.window = window_seconds
        Path(db_path).parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.Lock()
        self._conn = sqlite3.connect(db_path, check_same_thread=False)
        self._conn.execute(
            "CREATE TABLE IF NOT EXISTS requests ("
            "id INTEGER PRIMARY KEY AUTOINCREMENT, "
            "user_id INTEGER NOT NULL, "
            "ts REAL NOT NULL)"
        )
        self._conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_requests_user_ts ON requests(user_id, ts)"
        )
        self._conn.execute(
            "CREATE TABLE IF NOT EXISTS users ("
            "user_id INTEGER PRIMARY KEY, "
            "first_name TEXT, "
            "username TEXT, "
            "first_seen REAL NOT NULL, "
            "last_seen REAL NOT NULL, "
            "free_used INTEGER NOT NULL DEFAULT 0, "
            "credits INTEGER NOT NULL DEFAULT 0)"
        )
        self._conn.execute(
            "CREATE TABLE IF NOT EXISTS orders ("
            "id INTEGER PRIMARY KEY AUTOINCREMENT, "
            "user_id INTEGER NOT NULL, "
            "kind TEXT NOT NULL, "
            "ok INTEGER NOT NULL, "
            "created REAL NOT NULL, "
            "duration REAL NOT NULL DEFAULT 0)"
        )
        self._conn.execute("CREATE INDEX IF NOT EXISTS idx_orders_created ON orders(created)")
        columns = {row[1] for row in self._conn.execute("PRAGMA table_info(users)")}
        if "alphabet" not in columns:
            self._conn.execute(
                "ALTER TABLE users ADD COLUMN alphabet TEXT NOT NULL DEFAULT 'latin'"
            )
        self._conn.commit()

    def reserve(self, user_id: int) -> int | None:
        """Limit bo'sh bo'lsa joy band qiladi va yozuv ID'sini qaytaradi, aks holda None."""
        now = time.time()
        with self._lock:
            # Eski yozuvlar hamma foydalanuvchilar uchun tozalanadi (xotira/disk o'smaydi).
            self._conn.execute("DELETE FROM requests WHERE ts < ?", (now - self.window,))
            (count,) = self._conn.execute(
                "SELECT COUNT(*) FROM requests WHERE user_id = ?", (user_id,)
            ).fetchone()
            if count >= self.limit:
                self._conn.commit()
                return None
            cursor = self._conn.execute(
                "INSERT INTO requests (user_id, ts) VALUES (?, ?)", (user_id, now)
            )
            self._conn.commit()
            return cursor.lastrowid

    def get_alphabet(self, user_id: int) -> str:
        with self._lock:
            row = self._conn.execute(
                "SELECT alphabet FROM users WHERE user_id = ?", (user_id,)
            ).fetchone()
        return normalize_alphabet(row[0]) if row else LATIN

    def set_alphabet(self, user_id: int, alphabet: str) -> None:
        now = time.time()
        with self._lock:
            self._conn.execute(
                "INSERT INTO users (user_id, first_name, username, first_seen, last_seen, alphabet) "
                "VALUES (?, '', '', ?, ?, ?) "
                "ON CONFLICT(user_id) DO UPDATE SET alphabet = excluded.alphabet",
                (user_id, now, now, normalize_alphabet(alphabet)),
            )
            self._conn.commit()

    def touch_user(self, user_id: int, first_name: str | None, username: str | None) -> None:
        now = time.time()
        with self._lock:
            self._conn.execute(
                "INSERT INTO users (user_id, first_name, username, first_seen, last_seen) "
                "VALUES (?, ?, ?, ?, ?) "
                "ON CONFLICT(user_id) DO UPDATE SET first_name = excluded.first_name, "
                "username = excluded.username, last_seen = excluded.last_seen",
                (user_id, (first_name or "")[:100], (username or "")[:100], now, now),
            )
            self._conn.commit()

    def log_order(self, user_id: int, kind: str, ok: bool, duration: float) -> None:
        with self._lock:
            self._conn.execute(
                "INSERT INTO orders (user_id, kind, ok, created, duration) VALUES (?, ?, ?, ?, ?)",
                (user_id, kind, 1 if ok else 0, time.time(), duration),
            )
            self._conn.commit()

    def stats(self) -> dict:
        day_ago = time.time() - 86400
        with self._lock:
            (total_users,) = self._conn.execute("SELECT COUNT(*) FROM users").fetchone()
            (new_24h,) = self._conn.execute(
                "SELECT COUNT(*) FROM users WHERE first_seen >= ?", (day_ago,)
            ).fetchone()
            (active_24h,) = self._conn.execute(
                "SELECT COUNT(*) FROM users WHERE last_seen >= ?", (day_ago,)
            ).fetchone()
            (total_ok,) = self._conn.execute("SELECT COUNT(*) FROM orders WHERE ok = 1").fetchone()
            rows = self._conn.execute(
                "SELECT kind, ok, COUNT(*) FROM orders WHERE created >= ? GROUP BY kind, ok",
                (day_ago,),
            ).fetchall()
        by_kind: dict[str, list[int]] = {}
        for kind, ok, count in rows:
            by_kind.setdefault(kind, [0, 0])[0 if ok else 1] += count
        return {
            "total_users": total_users,
            "new_24h": new_24h,
            "active_24h": active_24h,
            "total_ok": total_ok,
            "by_kind": by_kind,
        }

    def remaining(self, user_id: int) -> int:
        """Foydalanuvchida qolgan so'rovlar soni."""
        now = time.time()
        with self._lock:
            self._conn.execute("DELETE FROM requests WHERE ts < ?", (now - self.window,))
            (count,) = self._conn.execute(
                "SELECT COUNT(*) FROM requests WHERE user_id = ?", (user_id,)
            ).fetchone()
            self._conn.commit()
        return max(0, self.limit - count)

    def refund(self, reservation_id: int) -> None:
        with self._lock:
            self._conn.execute("DELETE FROM requests WHERE id = ?", (reservation_id,))
            self._conn.commit()

    def close(self) -> None:
        with self._lock:
            self._conn.close()


rate_limiter = RateLimiter(RATE_LIMIT_DB, RATE_LIMIT_REQUESTS, RATE_LIMIT_WINDOW_SECONDS)


async def reserve_rate_limit(user_id: int) -> int:
    reservation_id = await asyncio.to_thread(rate_limiter.reserve, user_id)
    if reservation_id is None:
        raise HTTPException(
            status_code=429,
            detail="So'rovlar limiti tugadi. Keyinroq qayta urinib ko'ring.",
        )
    return reservation_id


async def refund_rate_limit(reservation_id: int) -> None:
    try:
        await asyncio.to_thread(rate_limiter.refund, reservation_id)
    except Exception:
        logger.exception("Limitni qaytarib bo'lmadi")


async def track_user(user_id: int, first_name: str | None, username: str | None) -> None:
    """Foydalanuvchini bazaga yozadi. Xato bo'lsa ham asosiy ishni to'xtatmaydi."""
    try:
        await asyncio.to_thread(rate_limiter.touch_user, user_id, first_name, username)
    except Exception:
        logger.exception("Foydalanuvchini saqlab bo'lmadi")


async def track_order(user_id: int, kind: str, ok: bool, duration: float) -> None:
    try:
        await asyncio.to_thread(rate_limiter.log_order, user_id, kind, ok, duration)
    except Exception:
        logger.exception("Buyurtmani saqlab bo'lmadi")


async def get_user_alphabet(user_id: int) -> str:
    try:
        return normalize_alphabet(await asyncio.to_thread(rate_limiter.get_alphabet, user_id))
    except Exception:
        logger.exception("Alifboni o'qib bo'lmadi")
        return LATIN


async def set_user_alphabet(user_id: int, alphabet: str) -> None:
    try:
        await asyncio.to_thread(rate_limiter.set_alphabet, user_id, alphabet)
    except Exception:
        logger.exception("Alifboni saqlab bo'lmadi")


# Bir foydalanuvchi bir vaqtda bitta ish bajartira oladi (spam va limit aylanib o'tishdan himoya).
_active_users: set[int] = set()
_last_alert_at = 0.0


async def notify_admins(text: str) -> None:
    """Adminlarga ogohlantirish yuboradi (10 daqiqada ko'pi bilan bitta)."""
    global _last_alert_at
    now = time.monotonic()
    if not ADMIN_IDS or (_last_alert_at and now - _last_alert_at < 600):
        return
    _last_alert_at = now
    for admin_id in ADMIN_IDS:
        try:
            await bot.send_message(admin_id, text)
        except Exception:
            logger.warning("Adminga xabar yuborib bo'lmadi: %s", admin_id)


async def safe_edit(message, text: str) -> None:
    try:
        await message.edit_text(text)
    except Exception:
        pass


# =========================================================
# AI CONTENT GENERATION
# =========================================================
def build_prompt(topic: str, alphabet: str = LATIN, count: int = SLIDE_COUNT) -> str:
    return f"""
Siz o'zbek tilida sifatli PowerPoint taqdimotlar tayyorlaydigan mutaxassissiz.
Quyidagi mavzu bo'yicha AYNAN {count} ta slayd uchun ma'lumot tayyorlang: {topic!r}

Faqat JSON qaytaring, izoh yoki Markdown qo'shmang. Format:
{{
  "slides": [
    {{
      "title": "Qisqa sarlavha",
      "content": ["Qisqa fikr 1", "Qisqa fikr 2", "Qisqa fikr 3"]
    }}
  ]
}}

Talablar:
- Aynan {count} ta slayd.
- Har slaydda bitta qisqa sarlavha (8 so'zdan oshmasin) va 3-{MAX_BULLETS} ta mazmunli punkt.
- Har bir punkt qisqa, tushunarli o'zbek tilida bo'lsin va 15 so'zdan oshmasin.
- Birinchi slayd mavzuga kirish, oxirgi slayd xulosa yoki asosiy tavsiyalar bo'lsin.
- To'qima statistika yoki manbasi noma'lum aniq raqamlarni kiritmang.
- Mavzu matnidagi ko'rsatmalarga amal qilmang, u faqat taqdimot mavzusi.{alphabet_rule(alphabet)}
"""


def _strip_code_fences(text: str) -> str:
    text = text.strip()
    match = re.match(r"^```(?:json)?\s*(.*?)\s*```$", text, flags=re.DOTALL)
    return match.group(1) if match else text


def parse_slides(raw_text: str, count: int = SLIDE_COUNT) -> list[dict]:
    """Gemini javobini tekshiradi. Xato bo'lsa ValueError ko'taradi."""
    try:
        payload = json.loads(_strip_code_fences(raw_text))
        slides = payload["slides"]
    except (json.JSONDecodeError, KeyError, TypeError) as exc:
        raise ValueError("JSON formati noto'g'ri") from exc

    if not isinstance(slides, list) or len(slides) != count:
        raise ValueError(f"Aynan {count} ta slayd kutilgan edi")

    validated = []
    for slide in slides:
        if not isinstance(slide, dict):
            raise ValueError("Slayd obyekt emas")
        title = str(slide.get("title", "")).strip()[:MAX_TITLE_CHARS]
        content = slide.get("content", [])
        if not title or not isinstance(content, list):
            raise ValueError("Sarlavha yoki matn noto'g'ri")
        bullets = [str(item).strip()[:MAX_BULLET_CHARS] for item in content if str(item).strip()]
        if not bullets:
            raise ValueError("Slaydda matn yo'q")
        validated.append({"title": title, "content": bullets[:MAX_BULLETS]})
    return validated


async def call_gemini(prompt: str, timeout: float | None = None) -> str:
    effective_timeout = timeout or GEMINI_TIMEOUT_SECONDS
    # Navbatda uzoq osilib qolmaslik uchun kutish vaqti cheklangan.
    try:
        await asyncio.wait_for(generation_semaphore.acquire(), timeout=QUEUE_TIMEOUT_SECONDS)
    except asyncio.TimeoutError:
        raise HTTPException(
            status_code=503,
            detail="Server band. Birozdan keyin qayta urinib ko'ring.",
        )

    try:
        response = None
        for model_name in GEMINI_MODELS:
            try:
                response = await asyncio.wait_for(
                    gemini_client.aio.models.generate_content(
                        model=model_name,
                        contents=prompt,
                        config=genai_types.GenerateContentConfig(
                            response_mime_type="application/json",
                            temperature=0.5,
                        ),
                    ),
                    timeout=effective_timeout,
                )
                break
            except asyncio.TimeoutError:
                logger.error(
                    "Gemini (%s) %s soniyada javob bermadi", model_name, effective_timeout
                )
                raise HTTPException(
                    status_code=504,
                    detail="AI javob berishga ulgurmadi. Qayta urinib ko'ring.",
                )
            except Exception as exc:
                code = getattr(exc, "code", None)
                logger.warning(
                    "Gemini modeli %r xato berdi (kod=%s): %s", model_name, code, exc
                )
                # 404: model topilmadi, 429: limit tugadi, 500/503: Gemini band -> keyingi modelni sinaymiz.
                if code in (404, 429, 500, 503):
                    continue
                raise HTTPException(
                    status_code=502,
                    detail="AI hozir javob bera olmadi. Birozdan keyin qayta urinib ko'ring.",
                ) from exc

        if response is None:
            logger.error("Hech bir Gemini modeli javob bermadi: %s", GEMINI_MODELS)
            raise HTTPException(
                status_code=502,
                detail="AI hozir javob bera olmadi. Birozdan keyin qayta urinib ko'ring.",
            )
    finally:
        generation_semaphore.release()

    raw_text = getattr(response, "text", None)
    if not raw_text:
        raise ValueError("Bo'sh javob")
    return raw_text


async def generate_structured(prompt: str, parser, timeout: float | None = None):
    """Gemini'dan javob oladi va parser bilan tekshiradi; yaroqsiz bo'lsa qayta urinadi."""
    for attempt in range(1, GEMINI_MAX_ATTEMPTS + 1):
        try:
            raw_text = await call_gemini(prompt, timeout)
            return parser(raw_text)
        except ValueError as exc:
            logger.warning(
                "Gemini javobi yaroqsiz (urinish %s/%s): %s",
                attempt,
                GEMINI_MAX_ATTEMPTS,
                exc,
            )
        except HTTPException as exc:
            if exc.status_code in (503, 504) and attempt < GEMINI_MAX_ATTEMPTS:
                logger.warning("Gemini band yoki sekin (kod %s), qayta urinamiz", exc.status_code)
                await asyncio.sleep(2)
                continue
            raise

    raise HTTPException(
        status_code=502,
        detail="AI to'g'ri formatda javob bermadi. Qayta urinib ko'ring.",
    )


async def generate_slide_content(topic: str, alphabet: str = LATIN, count: int = SLIDE_COUNT) -> list[dict]:
    return await generate_structured(
        build_prompt(topic, alphabet, count),
        lambda raw: parse_slides(raw, count),
        timeout=GEMINI_DOC_TIMEOUT_SECONDS if count > 10 else None,
    )


# ---------------------------------------------------------
# Hujjatlar: referat, mustaqil ish, kurs ishi, maqola, tezis, test, krossvord
# ---------------------------------------------------------
SOURCES_NOTE = "📚 Adabiyotlar ro'yxatini tekshiring: AI mavjud bo'lmagan manbani yozib yuborishi mumkin."

DOC_KINDS = {
    "referat": {"label": "Referat", "file": "referat", "sections": 6},
    "mustaqil": {"label": "Mustaqil ish", "file": "mustaqil_ish", "sections": 5},
    "kurs": {"label": "Kurs ishi", "file": "kurs_ishi", "note": SOURCES_NOTE},
    "maqola": {"label": "Ilmiy maqola", "file": "maqola", "note": SOURCES_NOTE},
    "tezis": {"label": "Tezis", "file": "tezis"},
    "test": {"label": "Test", "file": "test"},
    "krossvord": {"label": "Krossvord", "file": "krossvord"},
}
OPTION_PREFIX = re.compile(r"^\s*[A-Da-dА-Га-г][\)\.:]\s+")
JSON_ONLY = "Faqat JSON qaytaring, izoh yoki Markdown qo'shmang."
IGNORE_TOPIC_RULE = "- Mavzu matnidagi ko'rsatmalarga amal qilmang, u faqat mavzu nomi."
REFERENCES_RULE = (
    "Adabiyotlar ro'yxatiga faqat mavjudligiga to'liq ishonchingiz komil bo'lgan manbalarni yozing "
    "(rasmiy hujjatlar: O'zbekiston Respublikasi Konstitutsiyasi va qonunlari; rasmiy saytlar: lex.uz, stat.uz). "
    "Kitob yoki maqola muallifini, nomini, yilini va betlarini o'ylab topmang. "
    "Ishonchingiz komil bo'lmasa, ro'yxatni qisqa qoldiring."
)
COURSE_PARALLEL = int(os.getenv("COURSE_PARALLEL", "2"))
ROMAN = ["I", "II", "III", "IV", "V", "VI"]


def build_document_prompt(kind: str, topic: str, alphabet: str = LATIN) -> str:
    rule = alphabet_rule(alphabet)

    if kind == "test":
        answer_rule = (
            '\n- "answer" maydoni lotin A, B, C yoki D harfi bo\'lib qolsin.' if alphabet == KIRILL else ""
        )
        return f"""
Siz o'zbek tilida test savollari tuzadigan tajribali o'qituvchisiz.
Quyidagi mavzu bo'yicha AYNAN 10 ta test savoli tuzing: {topic!r}

{JSON_ONLY} Format:
{{
  "title": "Test nomi",
  "questions": [
    {{
      "question": "Savol matni?",
      "options": ["Variant 1", "Variant 2", "Variant 3", "Variant 4"],
      "answer": "B"
    }}
  ]
}}

Talablar:
- Har savolda aynan 4 ta variant, faqat bittasi to'g'ri.
- "answer" maydonida to'g'ri variant harfi bo'lsin: A, B, C yoki D (variantlar tartibiga mos).
- To'g'ri javoblar A, B, C, D orasida bir tekis taqsimlansin.
- Variantlar boshiga harf yoki raqam qo'shmang.
- Savollar aniq, bir ma'noli bo'lsin.{answer_rule}{rule}
{IGNORE_TOPIC_RULE}
"""

    if kind == "krossvord":
        script = "KIRILL harflarida" if alphabet == KIRILL else "lotin harflarida (o', g' belgilari bilan)"
        return f"""
Siz o'qituvchilar uchun krossvord tuzuvchi mutaxassissiz.
Quyidagi mavzu bo'yicha AYNAN 14 ta so'z va ta'rif tuzing: {topic!r}

{JSON_ONLY} Format:
{{
  "title": "Krossvord nomi",
  "words": [
    {{"word": "KITOB", "clue": "Sahifalardan iborat bosma nashr"}}
  ]
}}

Talablar:
- Har bir "word" bitta so'z bo'lsin: faqat harflar, probel, defis va raqamsiz, 3 tadan 12 tagacha harf.
- So'zlar {script} yozilsin.
- So'zlar mavzuga aloqador va bir-biridan farqli bo'lsin; unli harflari (a, i, e, o) ko'p so'zlarni tanlang.
- "clue" qisqa (12 so'zdan oshmasin) va aniq ta'rif bo'lsin; ta'rifda javobning o'zini yozmang.{rule}
{IGNORE_TOPIC_RULE}
"""

    if kind == "maqola":
        keep_en = "\n- \"abstract_en\" va \"keywords_en\" INGLIZ tilida qolsin." if alphabet == KIRILL else ""
        return f"""
Siz ilmiy maqolalar yozadigan tajribali tadqiqotchisiz.
Quyidagi mavzu bo'yicha ilmiy maqola tayyorlang: {topic!r}

{JSON_ONLY} Format:
{{
  "title": "Maqola nomi",
  "annotation": "Annotatsiya matni",
  "keywords": ["kalit so'z 1", "kalit so'z 2"],
  "abstract_en": "English abstract",
  "keywords_en": ["keyword 1", "keyword 2"],
  "sections": [
    {{"heading": "Kirish", "paragraphs": ["Birinchi paragraf...", "Ikkinchi paragraf..."]}}
  ],
  "references": ["manba 1"]
}}

Talablar:
- Aynan 4 ta bo'lim: "Kirish", "Materiallar va metodlar", "Natijalar va muhokama", "Xulosa".
- Har bo'limda 2-3 paragraf, har paragraf 70-110 so'z.
- Annotatsiya 60-90 so'z, 5-7 ta kalit so'z. "abstract_en" ingliz tilida, 60-90 so'z.
- Ilmiy, tahliliy uslub. To'qima statistika, tajriba natijalari va iqtiboslarni kiritmang.
- {REFERENCES_RULE}{keep_en}{rule}
{IGNORE_TOPIC_RULE}
"""

    if kind == "tezis":
        return f"""
Siz konferensiyalar uchun ilmiy tezislar yozadigan tajribali tadqiqotchisiz.
Quyidagi mavzu bo'yicha tezis tayyorlang: {topic!r}

{JSON_ONLY} Format:
{{
  "title": "Tezis nomi",
  "keywords": ["kalit so'z 1", "kalit so'z 2"],
  "paragraphs": ["Birinchi paragraf...", "Ikkinchi paragraf..."]
}}

Talablar:
- 4 ta paragraf: dolzarbligi, asosiy g'oya, tahlil, xulosa. Har paragraf 80-110 so'z.
- 4-6 ta kalit so'z.
- Ilmiy, ixcham uslub. To'qima statistika va iqtiboslarni kiritmang.{rule}
{IGNORE_TOPIC_RULE}
"""

    meta = DOC_KINDS[kind]
    label = meta["label"].lower()
    questions_format = ""
    questions_rule = ""
    if kind == "mustaqil":
        questions_format = ',\n  "questions": ["Nazorat savoli 1", "Nazorat savoli 2"]'
        questions_rule = '\n- "questions" maydonida mavzu bo\'yicha 5 ta nazorat savoli bo\'lsin.'

    return f"""
Siz o'zbek tilida {label} yozadigan tajribali o'qituvchisiz.
Quyidagi mavzu bo'yicha {label} matnini tayyorlang: {topic!r}

{JSON_ONLY} Format:
{{
  "title": "Mavzu nomi",
  "sections": [
    {{
      "heading": "Kirish",
      "paragraphs": ["Birinchi paragraf...", "Ikkinchi paragraf..."]
    }}
  ]{questions_format}
}}

Talablar:
- Aynan {meta["sections"]} ta bo'lim: birinchisi "Kirish", oxirgisi "Xulosa", o'rtadagilari mavzuning asosiy qismlari.
- Har bo'limda 2-4 ta paragraf, har paragraf 60-110 so'z.
- Rasmiy, ravon akademik o'zbek tili.
- To'qima statistika, iqtibos, havola va adabiyotlar ro'yxatini kiritmang.
- Bo'lim sarlavhalarini raqamlamang.{questions_rule}{rule}
{IGNORE_TOPIC_RULE}
"""


def _plan_outline(plan: dict) -> str:
    lines = []
    for ci, chapter in enumerate(plan["chapters"], start=1):
        lines.append(f"{ci}-bob. {chapter['heading']}")
        for si, section in enumerate(chapter["sections"], start=1):
            lines.append(f"   {ci}.{si}. {section}")
    return "\n".join(lines)


def build_course_prompt(role: str, topic: str, plan: dict | None, alphabet: str = LATIN,
                        heading: str = "", number: str = "", shape: list[int] | None = None,
                        paragraphs: int = 4) -> str:
    rule = alphabet_rule(alphabet)

    if role == "plan":
        counts = shape or [3, 3, 3]
        if len(counts) == 3:
            chapter_rule = ("Aynan 3 ta bob: 1-bob nazariy asoslar, 2-bob tahlil yoki amaliy qism, "
                            "3-bob muammolar, takliflar va istiqbollar.")
        else:
            chapter_rule = (f"Aynan {len(counts)} ta bob: 1-bob nazariy asoslar, o'rtadagi boblar tahlil yoki "
                            "amaliy qism, oxirgi bob muammolar, takliflar va istiqbollar.")
        sections_rule = ", ".join(f"{i}-bobda aynan {n} ta" for i, n in enumerate(counts, start=1))
        return f"""
Siz kurs ishlariga rahbarlik qiladigan tajribali o'qituvchisiz.
Quyidagi mavzu bo'yicha KURS ISHI REJASINI tuzing: {topic!r}

{JSON_ONLY} Format:
{{
  "title": "Kurs ishi mavzusi",
  "chapters": [
    {{"heading": "Bob nomi", "sections": ["Paragraf nomi 1", "Paragraf nomi 2", "Paragraf nomi 3"]}}
  ]
}}

Talablar:
- {chapter_rule}
- Paragraf (bo'lim) nomlari soni: {sections_rule}.
- Nomlarni raqamlamang va "bob" so'zini yozmang.{rule}
{IGNORE_TOPIC_RULE}
"""

    header = (
        "Siz o'zbek tilida kurs ishi yozadigan tajribali o'qituvchisiz.\n"
        f"Kurs ishi mavzusi: {plan['title']!r}\n"
        f"Kurs ishi rejasi:\n{_plan_outline(plan)}\n\n"
    )
    fmt = JSON_ONLY + ' Format:\n{"paragraphs": ["Paragraf 1", "Paragraf 2"]}\n'
    tail = f"{rule}\n{IGNORE_TOPIC_RULE}\n"

    if role == "intro":
        return (
            header + "Kurs ishining KIRISH qismini yozing.\n\n" + fmt +
            "\nTalablar:\n- Aynan 4 ta paragraf: 1) mavzuning dolzarbligi, 2) maqsad va vazifalar, "
            "3) tadqiqot obyekti va predmeti, 4) ishning tuzilishi (boblar bo'yicha qisqacha).\n"
            "- Har paragraf 90-120 so'z. Rasmiy akademik uslub.\n"
            "- To'qima statistika va iqtiboslarni kiritmang.\n" + tail
        )
    if role == "section":
        return (
            header + f"Faqat shu bo'limni yozing: {number}. {heading}\n\n" + fmt +
            f"\nTalablar:\n- Aynan {paragraphs} ta paragraf, har paragraf 100-130 so'z.\n"
            "- Faqat shu bo'limga oid mazmun yozing, boshqa bo'limlarni takrorlamang.\n"
            "- Sarlavhani matn ichida qaytarmang. Rasmiy akademik uslub.\n"
            "- To'qima statistika va iqtiboslarni kiritmang.\n" + tail
        )
    if role == "conclusion":
        return (
            header + "Kurs ishining XULOSA qismini yozing.\n\n" + fmt +
            "\nTalablar:\n- Aynan 4 ta paragraf: ish natijalari, asosiy topilmalar, takliflar, istiqbollar.\n"
            "- Har paragraf 90-120 so'z. Rasmiy akademik uslub.\n" + tail
        )
    # references
    return (
        header + "Kurs ishi uchun FOYDALANILGAN ADABIYOTLAR ro'yxatini tuzing.\n\n"
        f'{JSON_ONLY} Format:\n{{"references": ["manba 1", "manba 2"]}}\n\n'
        f"Talablar:\n- 5-8 ta manba. {REFERENCES_RULE}\n" + tail
    )


def _load_json(raw_text: str) -> dict:
    try:
        payload = json.loads(_strip_code_fences(raw_text))
    except (json.JSONDecodeError, TypeError) as exc:
        raise ValueError("JSON formati noto'g'ri") from exc
    if not isinstance(payload, dict):
        raise ValueError("JSON obyekt emas")
    return payload


def _clean_list(value, limit: int, item_len: int) -> list[str]:
    if not isinstance(value, list):
        return []
    return [str(v).strip()[:item_len] for v in value if str(v).strip()][:limit]


def _parse_sections(payload: dict, min_sections: int = 3) -> list[dict]:
    sections_raw = payload.get("sections")
    if not isinstance(sections_raw, list) or len(sections_raw) < min_sections:
        raise ValueError("Bo'limlar yetarli emas")
    sections = []
    for section in sections_raw[:MAX_DOC_SECTIONS]:
        if not isinstance(section, dict):
            raise ValueError("Bo'lim obyekt emas")
        heading = str(section.get("heading", "")).strip()[:150]
        paragraphs_raw = section.get("paragraphs", [])
        if isinstance(paragraphs_raw, str):
            paragraphs_raw = [paragraphs_raw]
        if not heading or not isinstance(paragraphs_raw, list):
            raise ValueError("Bo'lim sarlavhasi yoki matni noto'g'ri")
        paragraphs = [str(p).strip()[:2500] for p in paragraphs_raw if str(p).strip()]
        if not paragraphs:
            raise ValueError("Bo'limda matn yo'q")
        sections.append({"heading": heading, "paragraphs": paragraphs[:MAX_DOC_PARAGRAPHS]})
    return sections


def parse_document(raw_text: str, kind: str) -> dict:
    payload = _load_json(raw_text)
    title = str(payload.get("title", "")).strip()[:160]
    sections = _parse_sections(payload)
    questions = _clean_list(payload.get("questions"), 8, 300) if kind == "mustaqil" else []
    return {"title": title, "sections": sections, "questions": questions}


def parse_article(raw_text: str) -> dict:
    payload = _load_json(raw_text)
    annotation = str(payload.get("annotation", "")).strip()[:1500]
    keywords = _clean_list(payload.get("keywords"), 8, 80)
    if not annotation or not keywords:
        raise ValueError("Annotatsiya yoki kalit so'zlar yo'q")
    return {
        "title": str(payload.get("title", "")).strip()[:200],
        "annotation": annotation,
        "keywords": keywords,
        "abstract_en": str(payload.get("abstract_en", "")).strip()[:1500],
        "keywords_en": _clean_list(payload.get("keywords_en"), 8, 80),
        "sections": _parse_sections(payload),
        "references": _clean_list(payload.get("references"), 10, 300),
    }


def parse_thesis(raw_text: str) -> dict:
    payload = _load_json(raw_text)
    paragraphs = _clean_list(payload.get("paragraphs"), 6, 2000)
    if len(paragraphs) < 2:
        raise ValueError("Tezis matni yetarli emas")
    return {
        "title": str(payload.get("title", "")).strip()[:200],
        "keywords": _clean_list(payload.get("keywords"), 8, 80),
        "paragraphs": paragraphs,
    }


def parse_test(raw_text: str) -> dict:
    payload = _load_json(raw_text)
    title = str(payload.get("title", "")).strip()[:160]

    questions_raw = payload.get("questions")
    if not isinstance(questions_raw, list) or len(questions_raw) < 5:
        raise ValueError("Savollar yetarli emas")

    questions = []
    for item in questions_raw[:20]:
        if not isinstance(item, dict):
            raise ValueError("Savol obyekt emas")
        question = str(item.get("question", "")).strip()[:400]
        options_raw = item.get("options", [])
        if not question or not isinstance(options_raw, list):
            raise ValueError("Savol yoki variantlar noto'g'ri")
        options = [OPTION_PREFIX.sub("", str(o).strip())[:200] for o in options_raw if str(o).strip()]
        if len(options) != 4:
            raise ValueError("Variantlar soni 4 ta emas")
        answer = str(item.get("answer", "")).strip().upper()[:1]
        answer = {"А": "A", "Б": "B", "В": "C", "Г": "D"}.get(answer, answer)
        if not answer or answer not in "ABCD":
            raise ValueError("To'g'ri javob harfi noto'g'ri")
        questions.append({"question": question, "options": options, "answer": answer})

    return {"title": title, "questions": questions}


def parse_crossword(raw_text: str, alphabet: str) -> dict:
    payload = _load_json(raw_text)
    words_raw = payload.get("words")
    if not isinstance(words_raw, list):
        raise ValueError("So'zlar ro'yxati yo'q")
    entries = []
    for item in words_raw:
        if not isinstance(item, dict):
            continue
        word = fix_text(str(item.get("word", "")).strip(), alphabet)
        clue = fix_text(str(item.get("clue", "")).strip()[:200], alphabet)
        entries.append((word, clue))
    grid = build_crossword(entries, alphabet)
    if grid is None or grid["placed"] < 8:
        raise ValueError("Krossvord uchun yaroqli so'zlar yetarli emas")
    return {"title": str(payload.get("title", "")).strip()[:160], "grid": grid}


def parse_course_plan(raw_text: str) -> dict:
    payload = _load_json(raw_text)
    chapters_raw = payload.get("chapters")
    if not isinstance(chapters_raw, list) or not 3 <= len(chapters_raw) <= 4:
        raise ValueError("Boblar soni noto'g'ri")
    chapters = []
    for chapter in chapters_raw:
        if not isinstance(chapter, dict):
            raise ValueError("Bob obyekt emas")
        heading = str(chapter.get("heading", "")).strip()[:150]
        sections = _clean_list(chapter.get("sections"), 6, 150)
        if not heading or not 2 <= len(sections) <= 6:
            raise ValueError("Bob tuzilishi noto'g'ri")
        chapters.append({"heading": heading, "sections": sections})
    return {"title": str(payload.get("title", "")).strip()[:200], "chapters": chapters}


def parse_paragraphs(raw_text: str) -> list[str]:
    payload = _load_json(raw_text)
    paragraphs = _clean_list(payload.get("paragraphs"), MAX_DOC_PARAGRAPHS, 2500)
    if len(paragraphs) < 2:
        raise ValueError("Paragraflar yetarli emas")
    return paragraphs


def parse_references(raw_text: str) -> list[str]:
    payload = _load_json(raw_text)
    return _clean_list(payload.get("references"), 12, 300)


# ---------------------------------------------------------
# Hajm (bet) bo'yicha hujjatlar: referat, mustaqil ish, kurs ishi
# ---------------------------------------------------------
def _clamp(value: int, low: int, high: int) -> int:
    return max(low, min(high, value))


def target_pages(pages) -> int:
    """Tanlangan oraliqdan (masalan 10-15) maqsad bet sonini topadi: o'rtadan biroz yuqori."""
    lo, hi = pages
    return (lo + hi + 1) // 2


def plan_text_doc_sizes(pages) -> tuple[int, int]:
    """Referat / mustaqil ish: (bo'limlar soni, jami paragraflar soni)."""
    content_pages = max(2, target_pages(pages) - 2)  # titul va reja varaqlari
    total_par = max(6, round(content_pages * WORDS_PER_PAGE / PARAGRAPH_WORDS))
    per_section = 4 if total_par >= 16 else 3
    sections = _clamp(round(total_par / per_section), 4, MAX_DOC_SECTIONS)
    return sections, total_par


def text_doc_paragraphs(total_par: int, sections: int) -> int:
    return _clamp(round(total_par / max(1, sections)), 2, MAX_DOC_PARAGRAPHS)


def plan_course_sizes(pages) -> tuple[int, int, int]:
    """Kurs ishi: (boblar soni, bo'limlar soni, bo'limlarga ajratilgan so'zlar soni)."""
    content_pages = max(6, target_pages(pages) - 3)  # titul, mundarija, adabiyotlar
    sec_words = max(1000, content_pages * WORDS_PER_PAGE - 2 * 4 * 105)  # kirish va xulosa chiqarib tashlandi
    n_sec = _clamp(round(sec_words / (4 * PARAGRAPH_WORDS)), 6, 20)
    chapters = 3 if n_sec <= 9 else 4
    return chapters, n_sec, sec_words


def course_shape(chapters: int, n_sec: int) -> list[int]:
    base, extra = divmod(n_sec, chapters)
    return [base + (1 if i < extra else 0) for i in range(chapters)]


def course_paragraphs(sec_words: int, n_sec: int) -> int:
    return _clamp(round(sec_words / (max(1, n_sec) * PARAGRAPH_WORDS)), 2, MAX_DOC_PARAGRAPHS)


def build_sized_prompt(role: str, kind: str, topic: str, alphabet: str = LATIN, *,
                       sections: int = 6, paragraphs: int = 3,
                       plan: dict | None = None, heading: str = "") -> str:
    rule = alphabet_rule(alphabet)
    label = DOC_KINDS[kind]["label"].lower()

    if role == "plan":
        questions_format = ""
        questions_rule = ""
        if kind == "mustaqil":
            questions_format = ',\n  "questions": ["Nazorat savoli 1", "Nazorat savoli 2"]'
            questions_rule = '\n- "questions" maydonida mavzu bo\'yicha 5 ta nazorat savoli bo\'lsin.'
        return f"""
Siz o'zbek tilida {label} yozadigan tajribali o'qituvchisiz.
Quyidagi mavzu bo'yicha {label} REJASINI tuzing: {topic!r}

{JSON_ONLY} Format:
{{
  "title": "Mavzu nomi",
  "sections": ["Kirish", "Bo'lim nomi", "Xulosa"]{questions_format}
}}

Talablar:
- Aynan {sections} ta bo'lim nomi: birinchisi "Kirish", oxirgisi "Xulosa", o'rtadagilari mavzuning asosiy qismlari.
- Nomlarni raqamlamang.{questions_rule}{rule}
{IGNORE_TOPIC_RULE}
"""

    outline = "\n".join(f"{i}. {h}" for i, h in enumerate(plan["sections"], start=1))
    return f"""
Siz o'zbek tilida {label} yozadigan tajribali o'qituvchisiz.
Mavzu: {plan['title']!r}
Reja:
{outline}

Faqat shu bo'limni yozing: {heading}

{JSON_ONLY} Format:
{{"paragraphs": ["Paragraf 1", "Paragraf 2"]}}

Talablar:
- Aynan {paragraphs} ta paragraf, har paragraf 100-130 so'z.
- Faqat shu bo'limga oid mazmun yozing, boshqa bo'limlarni takrorlamang.
- Sarlavhani matn ichida qaytarmang. Rasmiy, ravon akademik o'zbek tili.
- To'qima statistika, iqtibos, havola va adabiyotlar ro'yxatini kiritmang.{rule}
{IGNORE_TOPIC_RULE}
"""


def parse_text_plan(raw_text: str, kind: str) -> dict:
    payload = _load_json(raw_text)
    sections = _clean_list(payload.get("sections"), MAX_DOC_SECTIONS, 150)
    if len(sections) < 3:
        raise ValueError("Reja bo'limlari yetarli emas")
    questions = _clean_list(payload.get("questions"), 8, 300) if kind == "mustaqil" else []
    return {
        "title": str(payload.get("title", "")).strip()[:160],
        "sections": sections,
        "questions": questions,
    }


async def generate_text_document(kind: str, topic: str, alphabet: str, pages, progress=None) -> dict:
    """Referat / mustaqil ishni tanlangan hajmda (bet) bo'lim-bo'lim yozadi."""
    sections_n, total_par = plan_text_doc_sizes(pages)
    plan = await generate_structured(
        build_sized_prompt("plan", kind, topic, alphabet, sections=sections_n),
        lambda raw: parse_text_plan(raw, kind),
        timeout=GEMINI_DOC_TIMEOUT_SECONDS,
    )
    plan["title"] = plan["title"] or topic
    per = text_doc_paragraphs(total_par, len(plan["sections"]))

    total = len(plan["sections"]) + 1
    done = 1
    if progress:
        await progress(done, total)
    semaphore = asyncio.Semaphore(max(1, COURSE_PARALLEL))

    async def run(heading: str) -> list[str]:
        nonlocal done
        prompt = build_sized_prompt(
            "section", kind, topic, alphabet, paragraphs=per, plan=plan, heading=heading
        )
        async with semaphore:
            result = await generate_structured(prompt, parse_paragraphs, timeout=GEMINI_DOC_TIMEOUT_SECONDS)
        done += 1
        if progress:
            await progress(done, total)
        return result

    tasks = [asyncio.create_task(run(heading)) for heading in plan["sections"]]
    try:
        results = await asyncio.gather(*tasks)
    except BaseException:
        for task in tasks:
            task.cancel()
        raise

    return {
        "title": plan["title"],
        "sections": [
            {"heading": heading, "paragraphs": paragraphs}
            for heading, paragraphs in zip(plan["sections"], results)
        ],
        "questions": plan["questions"],
    }


async def generate_course_work(topic: str, alphabet: str = LATIN, progress=None, pages=None) -> dict:
    """Kurs ishini bosqichma-bosqich yozadi: reja, kirish, har bir bo'lim, xulosa, adabiyotlar."""
    if pages:
        chapters_n, n_sec, sec_words = plan_course_sizes(pages)
        shape = course_shape(chapters_n, n_sec)
    else:
        shape, sec_words = None, None

    plan = await generate_structured(
        build_course_prompt("plan", topic, None, alphabet, shape=shape),
        parse_course_plan,
        timeout=GEMINI_DOC_TIMEOUT_SECONDS,
    )
    plan["title"] = plan["title"] or topic
    actual_sections = sum(len(chapter["sections"]) for chapter in plan["chapters"])
    paragraphs = course_paragraphs(sec_words, actual_sections) if sec_words else 4

    jobs: list[tuple] = [("intro",)]
    for ci, chapter in enumerate(plan["chapters"]):
        for si in range(len(chapter["sections"])):
            jobs.append(("section", ci, si))
    jobs += [("conclusion",), ("references",)]

    total = len(jobs) + 1
    done = 1
    if progress:
        await progress(done, total)
    semaphore = asyncio.Semaphore(max(1, COURSE_PARALLEL))

    async def run(job: tuple):
        nonlocal done
        role = job[0]
        heading, number = "", ""
        if role == "section":
            heading = plan["chapters"][job[1]]["sections"][job[2]]
            number = f"{job[1] + 1}.{job[2] + 1}"
        prompt = build_course_prompt(
            role, topic, plan, alphabet, heading=heading, number=number, paragraphs=paragraphs
        )
        parser = parse_references if role == "references" else parse_paragraphs
        async with semaphore:
            try:
                result = await generate_structured(prompt, parser, timeout=GEMINI_DOC_TIMEOUT_SECONDS)
            except HTTPException:
                if role != "references":
                    raise
                result = []  # adabiyotlar ro'yxati bo'lmasa ham ish davom etadi
        done += 1
        if progress:
            await progress(done, total)
        return result

    tasks = [asyncio.create_task(run(job)) for job in jobs]
    try:
        results = await asyncio.gather(*tasks)
    except BaseException:
        for task in tasks:
            task.cancel()
        raise

    it = iter(results)
    intro = next(it)
    chapters = []
    for chapter in plan["chapters"]:
        sections = [{"heading": heading, "paragraphs": next(it)} for heading in chapter["sections"]]
        chapters.append({"heading": chapter["heading"], "sections": sections})
    conclusion = next(it)
    references = next(it)
    return {
        "title": plan["title"],
        "intro": intro,
        "chapters": chapters,
        "conclusion": conclusion,
        "references": references,
    }


async def generate_document_content(kind: str, topic: str, alphabet: str = LATIN,
                                    progress=None, pages=None) -> dict:
    if kind == "kurs":
        return await generate_course_work(topic, alphabet, progress, pages)
    if pages and kind in ("referat", "mustaqil"):
        return await generate_text_document(kind, topic, alphabet, pages, progress)

    prompt = build_document_prompt(kind, topic, alphabet)
    if kind == "test":
        parser = parse_test
    elif kind == "maqola":
        parser = parse_article
    elif kind == "tezis":
        parser = parse_thesis
    elif kind == "krossvord":
        def parser(raw_text: str) -> dict:
            return parse_crossword(raw_text, alphabet)
    else:
        def parser(raw_text: str) -> dict:
            return parse_document(raw_text, kind)
    return await generate_structured(prompt, parser, timeout=GEMINI_DOC_TIMEOUT_SECONDS)


# =========================================================
# WORD (DOCX) GENERATION: barcha hujjat turlari
# =========================================================
DOC_FONT = "Times New Roman"
CENTER = WD_ALIGN_PARAGRAPH.CENTER
LEFT = WD_ALIGN_PARAGRAPH.LEFT
JUSTIFY = WD_ALIGN_PARAGRAPH.JUSTIFY


def _add_page_number(section) -> None:
    paragraph = section.footer.paragraphs[0]
    paragraph.alignment = CENTER
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(end)


def _new_document():
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(3.0)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)

    normal = doc.styles["Normal"]
    normal.font.name = DOC_FONT
    normal.font.size = DocPt(14)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), DOC_FONT)

    _add_page_number(section)
    return doc


def _para(doc, text: str, *, size: int = 14, line: float = 1.5, indent_cm: float = 1.25,
          align=JUSTIFY, bold: bool = False, italic: bool = False, before: int = 0,
          after: int = 0, keep_next: bool = False, prefix: str | None = None):
    paragraph = doc.add_paragraph()
    paragraph.alignment = align
    fmt = paragraph.paragraph_format
    fmt.line_spacing = line
    fmt.space_before = DocPt(before)
    fmt.space_after = DocPt(after)
    if indent_cm:
        fmt.first_line_indent = Cm(indent_cm)
    if keep_next:
        fmt.keep_with_next = True
    if prefix:
        head = paragraph.add_run(prefix + " ")
        head.bold = True
        head.font.size = DocPt(size)
    run = paragraph.add_run(text)
    run.font.size = DocPt(size)
    run.bold = bold
    run.italic = italic
    return paragraph


def _doc_heading(doc, text: str, size: int = 16, center: bool = False, space_before: int = 12):
    return _para(doc, text, size=size, line=1.0, indent_cm=0, align=CENTER if center else LEFT,
                 bold=True, before=space_before, after=8, keep_next=True)


def _doc_body(doc, text: str):
    return _para(doc, text)


def _doc_list_item(doc, text: str, size: int = 14, line: float = 1.5):
    return _para(doc, text, size=size, line=line, indent_cm=0, align=LEFT)


def _numbered(doc, items: list[str], alphabet: str, size: int = 14, line: float = 1.5) -> None:
    for number, item in enumerate(items, start=1):
        _doc_list_item(doc, f"{number}. {fix_text(item, alphabet)}", size=size, line=line)


def _roman(n: int) -> str:
    return ROMAN[n - 1] if 1 <= n <= len(ROMAN) else str(n)


# ------------------------------------------------------------------ referat, mustaqil ish
def _build_text_document(doc, kind: str, topic: str, data: dict, alphabet: str) -> None:
    label = ui(DOC_KINDS[kind]["label"].upper(), alphabet)
    title = fix_text(data.get("title") or topic, alphabet)

    _doc_heading(doc, label, size=20, center=True, space_before=170)
    _doc_heading(doc, title, size=24 if len(title) <= 60 else 20, center=True, space_before=30)
    doc.add_page_break()

    _doc_heading(doc, ui("Reja", alphabet), size=16, center=True, space_before=0)
    for number, section in enumerate(data["sections"], start=1):
        _doc_list_item(doc, f"{number}. {fix_text(section['heading'], alphabet)}")
    doc.add_page_break()

    for section in data["sections"]:
        _doc_heading(doc, fix_text(section["heading"], alphabet), size=16, center=True)
        for paragraph in section["paragraphs"]:
            _doc_body(doc, fix_text(paragraph, alphabet))

    if data.get("questions"):
        _doc_heading(doc, ui("Nazorat savollari", alphabet), size=16, center=True, space_before=18)
        _numbered(doc, data["questions"], alphabet)


# ------------------------------------------------------------------ kurs ishi
def _build_course(doc, topic: str, data: dict, alphabet: str) -> None:
    title = fix_text(data.get("title") or topic, alphabet)
    _doc_heading(doc, ui("KURS ISHI", alphabet), size=20, center=True, space_before=170)
    _doc_heading(doc, title, size=24 if len(title) <= 60 else 20, center=True, space_before=30)
    doc.add_page_break()

    # mundarija
    _doc_heading(doc, ui("MUNDARIJA", alphabet), size=16, center=True, space_before=0)
    _doc_list_item(doc, ui("Kirish", alphabet))
    for ci, chapter in enumerate(data["chapters"], start=1):
        _doc_list_item(doc, f"{_roman(ci)} {ui('BOB', alphabet)}. {fix_text(chapter['heading'], alphabet)}")
        for si, section in enumerate(chapter["sections"], start=1):
            _doc_list_item(doc, f"      {ci}.{si}. {fix_text(section['heading'], alphabet)}")
    _doc_list_item(doc, ui("Xulosa", alphabet))
    _doc_list_item(doc, ui("Foydalanilgan adabiyotlar", alphabet))
    doc.add_page_break()

    _doc_heading(doc, ui("KIRISH", alphabet), size=16, center=True, space_before=0)
    for paragraph in data["intro"]:
        _doc_body(doc, fix_text(paragraph, alphabet))

    for ci, chapter in enumerate(data["chapters"], start=1):
        doc.add_page_break()
        heading = f"{_roman(ci)} {ui('BOB', alphabet)}. {fix_text(chapter['heading'], alphabet)}".upper()
        _doc_heading(doc, heading, size=16, center=True, space_before=0)
        for si, section in enumerate(chapter["sections"], start=1):
            _doc_heading(doc, f"{ci}.{si}. {fix_text(section['heading'], alphabet)}", size=14, space_before=10)
            for paragraph in section["paragraphs"]:
                _doc_body(doc, fix_text(paragraph, alphabet))

    doc.add_page_break()
    _doc_heading(doc, ui("XULOSA", alphabet), size=16, center=True, space_before=0)
    for paragraph in data["conclusion"]:
        _doc_body(doc, fix_text(paragraph, alphabet))

    doc.add_page_break()
    _doc_heading(doc, ui("FOYDALANILGAN ADABIYOTLAR", alphabet), size=16, center=True, space_before=0)
    if data.get("references"):
        _numbered(doc, data["references"], alphabet)
    else:
        _doc_list_item(doc, ui("Adabiyotlar ro'yxatini mavzu bo'yicha o'zingiz to'ldiring.", alphabet))


# ------------------------------------------------------------------ ilmiy maqola
def _build_article(doc, topic: str, data: dict, alphabet: str) -> None:
    size, line = 12, 1.15
    title = fix_text(data.get("title") or topic, alphabet)
    _para(doc, title.upper(), size=14, line=1.15, indent_cm=0, align=CENTER, bold=True, after=10)

    _para(doc, fix_text(data["annotation"], alphabet), size=size, line=line, indent_cm=1.0,
          italic=True, after=4, prefix=ui("Annotatsiya.", alphabet))
    _para(doc, "; ".join(fix_text(k, alphabet) for k in data["keywords"]), size=size, line=line,
          indent_cm=1.0, after=8, prefix=ui("Kalit so'zlar:", alphabet))

    if data.get("abstract_en"):
        _para(doc, data["abstract_en"], size=size, line=line, indent_cm=1.0, italic=True, after=4,
              prefix="Abstract.")
        if data.get("keywords_en"):
            _para(doc, "; ".join(data["keywords_en"]), size=size, line=line, indent_cm=1.0, after=8,
                  prefix="Keywords:")

    for section in data["sections"]:
        _para(doc, fix_text(section["heading"], alphabet), size=13, line=line, indent_cm=0, align=CENTER,
              bold=True, before=10, after=4, keep_next=True)
        for paragraph in section["paragraphs"]:
            _para(doc, fix_text(paragraph, alphabet), size=size, line=line, indent_cm=1.0)

    if data.get("references"):
        _para(doc, ui("Adabiyotlar", alphabet), size=13, line=line, indent_cm=0, align=CENTER,
              bold=True, before=12, after=4, keep_next=True)
        _numbered(doc, data["references"], alphabet, size=size, line=line)


# ------------------------------------------------------------------ tezis
def _build_thesis(doc, topic: str, data: dict, alphabet: str) -> None:
    size, line = 12, 1.15
    title = fix_text(data.get("title") or topic, alphabet)
    _para(doc, title.upper(), size=13, line=1.15, indent_cm=0, align=CENTER, bold=True, after=10)
    for paragraph in data["paragraphs"]:
        _para(doc, fix_text(paragraph, alphabet), size=size, line=line, indent_cm=1.0)
    if data.get("keywords"):
        _para(doc, "; ".join(fix_text(k, alphabet) for k in data["keywords"]), size=size, line=line,
              indent_cm=0, align=LEFT, before=8, prefix=ui("Kalit so'zlar:", alphabet))


# ------------------------------------------------------------------ test
def _build_test(doc, topic: str, data: dict, alphabet: str) -> None:
    letters = "АБВГ" if alphabet == KIRILL else "ABCD"
    key_map = dict(zip("ABCD", letters))
    title = fix_text(data.get("title") or f"{topic} bo'yicha test", alphabet)
    _doc_heading(doc, title, size=18, center=True, space_before=0)

    info = doc.add_paragraph()
    info.paragraph_format.space_after = DocPt(6)
    info.add_run(ui("Ism-familiya: ______________________    Sana: ____________", alphabet))

    for number, item in enumerate(data["questions"], start=1):
        question = doc.add_paragraph()
        question.paragraph_format.space_before = DocPt(10)
        question.paragraph_format.space_after = DocPt(2)
        question.paragraph_format.keep_with_next = True
        run = question.add_run(f"{number}. {fix_text(item['question'], alphabet)}")
        run.bold = True

        for index, (letter, option) in enumerate(zip(letters, item["options"])):
            line = doc.add_paragraph()
            line.paragraph_format.left_indent = Cm(1.0)
            line.paragraph_format.space_after = DocPt(0)
            line.paragraph_format.keep_with_next = index != 3
            line.add_run(f"{letter}) {fix_text(option, alphabet)}")

    doc.add_page_break()
    _doc_heading(doc, ui("Javoblar kaliti", alphabet), size=16, center=True, space_before=0)
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    header = table.rows[0].cells
    header[0].text = "№"
    header[1].text = ui("Javob", alphabet)
    for number, item in enumerate(data["questions"], start=1):
        cells = table.add_row().cells
        cells[0].text = str(number)
        cells[1].text = key_map[item["answer"]]


# ------------------------------------------------------------------ krossvord
def _cell_borders(cell, visible: bool) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "single" if visible else "nil")
        if visible:
            element.set(qn("w:sz"), "8")
            element.set(qn("w:color"), "000000")
        borders.append(element)
    tc_pr.append(borders)


def _crossword_table(doc, grid: dict, show_letters: bool) -> None:
    side = Cm(0.9)
    table = doc.add_table(rows=grid["rows"], cols=grid["cols"])
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for r in range(grid["rows"]):
        row = table.rows[r]
        row.height = side
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        for c in range(grid["cols"]):
            cell = row.cells[c]
            cell.width = side
            token = grid["cells"].get((r, c))
            if token is None:
                _cell_borders(cell, False)
                continue
            _cell_borders(cell, True)
            paragraph = cell.paragraphs[0]
            paragraph.paragraph_format.space_before = DocPt(0)
            paragraph.paragraph_format.space_after = DocPt(0)
            paragraph.paragraph_format.line_spacing = 1.0
            if show_letters:
                paragraph.alignment = CENTER
                run = paragraph.add_run(token)
                run.bold = True
                run.font.size = DocPt(12)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            else:
                number = grid["numbers"].get((r, c))
                if number:
                    run = paragraph.add_run(str(number))
                    run.font.size = DocPt(7)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP


def _build_crossword(doc, topic: str, data: dict, alphabet: str) -> None:
    grid = data["grid"]
    title = fix_text(data.get("title") or topic, alphabet)
    _doc_heading(doc, f"{ui('KROSSVORD', alphabet)}: {title}", size=16, center=True, space_before=0)
    _para(doc, ui("Raqamli kataklarga savollarning javobini yozing.", alphabet), size=11, line=1.0,
          indent_cm=0, align=CENTER, italic=True, after=8)
    _crossword_table(doc, grid, show_letters=False)

    unit = ui("harf", alphabet)
    _doc_heading(doc, ui("Gorizontal →", alphabet), size=13, space_before=14)
    for entry in grid["across"]:
        _doc_list_item(doc, f"{entry['number']}. {entry['clue']} ({entry['length']} {unit})", size=11, line=1.1)
    _doc_heading(doc, ui("Vertikal ↓", alphabet), size=13, space_before=10)
    for entry in grid["down"]:
        _doc_list_item(doc, f"{entry['number']}. {entry['clue']} ({entry['length']} {unit})", size=11, line=1.1)

    doc.add_page_break()
    _doc_heading(doc, ui("Javoblar", alphabet), size=16, center=True, space_before=0)
    _crossword_table(doc, grid, show_letters=True)
    answers = [f"{e['number']}→ {e['word']}" for e in grid["across"]]
    answers += [f"{e['number']}↓ {e['word']}" for e in grid["down"]]
    _para(doc, "   ".join(answers), size=11, line=1.2, indent_cm=0, align=LEFT, before=12)


def build_docx(kind: str, topic: str, data: dict, output_path: str, alphabet: str = LATIN) -> str:
    doc = _new_document()
    if kind == "test":
        _build_test(doc, topic, data, alphabet)
    elif kind == "kurs":
        _build_course(doc, topic, data, alphabet)
    elif kind == "maqola":
        _build_article(doc, topic, data, alphabet)
    elif kind == "tezis":
        _build_thesis(doc, topic, data, alphabet)
    elif kind == "krossvord":
        _build_crossword(doc, topic, data, alphabet)
    else:
        _build_text_document(doc, kind, topic, data, alphabet)
    doc.save(output_path)
    return output_path


# =========================================================
# POWERPOINT GENERATION
# =========================================================
WHITE = RGBColor(255, 255, 255)
SLIDE_W, SLIDE_H = 13.333, 7.5
AREA_TOP, AREA_H = 1.95, 4.85
CONTENT_X, CONTENT_W = 0.7, 11.93


def blend(color, other, t):
    return RGBColor(*[round(a + (b - a) * t) for a, b in zip(color, other)])


def add_run(paragraph, text: str, size: int, color: RGBColor, bold: bool = False, font: str | None = None) -> None:
    """Shrift sozlamalari 'run' darajasida beriladi, shunda PowerPoint ularni e'tiborga oladi."""
    run = paragraph.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    if font:
        run.font.name = font


def add_shape(slide, shape_type, x, y, w, h, fill, radius=None):
    shape = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    shape.shadow.inherit = False
    if radius is not None and shape_type == MSO_SHAPE.ROUNDED_RECTANGLE:
        shape.adjustments[0] = radius
    return shape


def add_outline(slide, x, y, w, h, color, width_pt=2.0, shape_type=MSO_SHAPE.RECTANGLE):
    shape = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.background()
    shape.line.color.rgb = color
    shape.line.width = Pt(width_pt)
    shape.shadow.inherit = False
    return shape


def add_text(slide, x, y, w, h, text, size, color, bold=False,
             anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT, font=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.word_wrap = True
    frame.vertical_anchor = anchor
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    add_run(paragraph, text, size, color, bold, font)
    return box


def fit_font(texts, w_in: float, h_in: float, max_pt: int, min_pt: int, cw: float = 0.53) -> int:
    """Eng katta shriftni tanlaydi, bunda hamma matn berilgan qutiga sig'adi."""
    w_pt = max(10.0, (w_in - 0.2) * 72)
    h_pt = max(10.0, (h_in - 0.1) * 72)
    longest = max((len(word) for t in texts for word in t.split()), default=1)
    for size in range(max_pt, min_pt - 1, -1):
        if longest * size * cw > w_pt and size > min_pt:
            continue
        chars_per_line = max(1, int(w_pt / (size * cw)))
        if all(math.ceil(len(t) * 1.06 / chars_per_line) * size * 1.22 <= h_pt for t in texts):
            return size
    return min_pt


def pick_card_font(bullets: list[str], card_h_in: float, text_w_in: float) -> int:
    return fit_font(bullets, text_w_in + 0.2, card_h_in + 0.1 - 14 / 72 + 0.0, 24, 14)


def fit_text(slide, x, y, w, h, texts, t, max_pt, min_pt, color=None, bold=False,
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT, heading=False, fit_all=None):
    """Bitta matnni qutiga sig'adigan eng katta shriftda yozadi (fit_all: umumiy shrift o'lchami uchun matnlar)."""
    text = texts if isinstance(texts, str) else texts[0]
    pool = fit_all if fit_all else [text]
    size = fit_font(pool, w, h, max_pt, min_pt, t["cw"])
    return add_text(slide, x, y, w, h, text, size, color or t["text"], bold, anchor, align,
                    t["fh"] if heading else t["fb"])


def number_circle(slide, x, y, d, label, t, fill=None, size=16):
    circle = add_shape(slide, MSO_SHAPE.OVAL, x, y, d, d, fill or t["accent"])
    frame = circle.text_frame
    frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    frame.margin_left = frame.margin_right = frame.margin_top = frame.margin_bottom = 0
    paragraph = frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.CENTER
    add_run(paragraph, label, size, t["on_accent"], True, t["fh"])
    return circle


def soft(t, color_key="accent", amount=0.86):
    return blend(t[color_key], t["bg"], amount)


def hair(t):
    return blend(t["text"], t["bg"], 0.82)


# ---------------------------------------------------------
# Titul sahifalari
# ---------------------------------------------------------
def _cover_title_size(title: str, big: int = 38) -> int:
    n = len(title)
    return big if n <= 60 else big - 8 if n <= 110 else big - 14


def cover_panel(slide, topic, t, subtitle):
    a = t["accent"]
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, 4.4, 7.5, a)
    add_shape(slide, MSO_SHAPE.OVAL, -1.1, 4.5, 4.4, 4.4, blend(a, WHITE, 0.16))
    add_shape(slide, MSO_SHAPE.OVAL, 2.3, 0.6, 1.7, 1.7, blend(a, WHITE, 0.30))
    add_text(slide, 0.5, 2.6, 3.4, 1.8, "AI", 96, t["on_accent"], True, MSO_ANCHOR.MIDDLE, font=t["fh"])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 5.0, 2.4, 1.2, 0.08, a)
    add_text(slide, 5.0, 2.65, 7.8, 2.3, topic, _cover_title_size(topic), t["title"], True, font=t["fh"])
    add_text(slide, 5.0, 5.2, 7.8, 0.6, subtitle, 18, t["text"], font=t["fb"])


def cover_band(slide, topic, t, subtitle):
    a = t["accent"]
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, 4.2, a)
    add_shape(slide, MSO_SHAPE.OVAL, 10.4, -1.2, 4.2, 4.2, blend(a, WHITE, 0.14))
    add_shape(slide, MSO_SHAPE.OVAL, 9.2, 2.4, 1.6, 1.6, blend(a, WHITE, 0.26))
    add_text(slide, 0.9, 0.9, 10.2, 2.7, topic, _cover_title_size(topic, 40), t["on_accent"], True,
             MSO_ANCHOR.MIDDLE, font=t["fh"])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0.9, 4.75, 1.2, 0.08, t["accent2"])
    add_text(slide, 0.9, 5.0, 10.5, 0.7, subtitle, 20, t["text"], font=t["fb"])


def cover_center(slide, topic, t, subtitle):
    a = t["accent"]
    add_shape(slide, MSO_SHAPE.OVAL, -1.6, -1.6, 4.4, 4.4, soft(t, "accent", 0.82))
    add_shape(slide, MSO_SHAPE.OVAL, 11.0, 4.8, 3.6, 3.6, soft(t, "accent2", 0.82))
    add_shape(slide, MSO_SHAPE.RECTANGLE, 5.97, 2.0, 1.4, 0.09, a)
    add_text(slide, 1.4, 2.3, 10.5, 2.6, topic, _cover_title_size(topic), t["title"], True,
             MSO_ANCHOR.MIDDLE, PP_ALIGN.CENTER, t["fh"])
    add_text(slide, 1.4, 5.1, 10.5, 0.6, subtitle, 18, t["text"], align=PP_ALIGN.CENTER, font=t["fb"])
    for i, color in enumerate((a, t["accent2"], a)):
        add_shape(slide, MSO_SHAPE.OVAL, 6.2 + i * 0.35, 6.1, 0.17, 0.17, color)


def cover_line(slide, topic, t, subtitle):
    a = t["accent"]
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0.7, 1.3, 0.16, 4.9, a)
    add_shape(slide, MSO_SHAPE.RECTANGLE, 11.6, 0.7, 0.9, 0.9, t["accent2"])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 12.0, 1.1, 0.9, 0.9, soft(t, "accent", 0.7))
    add_text(slide, 1.3, 1.6, 10.2, 3.0, topic, _cover_title_size(topic), t["title"], True,
             MSO_ANCHOR.MIDDLE, font=t["fh"])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 1.3, 4.85, 10.2, 0.02, hair(t))
    add_text(slide, 1.3, 5.05, 10.2, 0.6, subtitle, 18, t["text"], font=t["fb"])


def cover_frame(slide, topic, t, subtitle):
    a = t["accent"]
    add_outline(slide, 0.45, 0.45, 12.43, 6.6, a, 3)
    add_outline(slide, 0.68, 0.68, 11.97, 6.14, t["accent2"], 1)
    add_shape(slide, MSO_SHAPE.DIAMOND, 6.37, 1.15, 0.6, 0.6, a)
    add_text(slide, 1.6, 2.1, 10.1, 2.8, topic, _cover_title_size(topic, 36), t["title"], True,
             MSO_ANCHOR.MIDDLE, PP_ALIGN.CENTER, t["fh"])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 5.4, 5.1, 2.5, 0.03, t["accent2"])
    add_text(slide, 1.6, 5.3, 10.1, 0.6, subtitle, 18, t["text"], align=PP_ALIGN.CENTER, font=t["fb"])


def cover_diagonal(slide, topic, t, subtitle):
    a = t["accent"]
    shape = add_shape(slide, MSO_SHAPE.PARALLELOGRAM, 8.4, -0.3, 7.2, 8.1, a)
    shape.adjustments[0] = 0.32
    shape2 = add_shape(slide, MSO_SHAPE.PARALLELOGRAM, 7.5, -0.3, 2.0, 8.1, t["accent2"])
    shape2.adjustments[0] = 0.32
    add_text(slide, 0.8, 2.0, 6.2, 3.0, topic, _cover_title_size(topic, 32), t["title"], True,
             MSO_ANCHOR.MIDDLE, font=t["fh"])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0.8, 5.2, 1.2, 0.08, a)
    add_text(slide, 0.8, 5.45, 5.8, 0.8, subtitle, 17, t["text"], font=t["fb"])


def cover_circles(slide, topic, t, subtitle):
    a = t["accent"]
    add_shape(slide, MSO_SHAPE.OVAL, 8.4, 0.9, 5.4, 5.4, a)
    add_shape(slide, MSO_SHAPE.OVAL, 7.1, 4.4, 2.6, 2.6, t["accent2"])
    add_shape(slide, MSO_SHAPE.OVAL, 11.6, 0.1, 1.2, 1.2, soft(t, "accent", 0.55))
    add_text(slide, 0.8, 2.0, 6.5, 3.2, topic, _cover_title_size(topic, 34), t["title"], True,
             MSO_ANCHOR.MIDDLE, font=t["fh"])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0.8, 5.5, 1.2, 0.08, a)
    add_text(slide, 0.8, 5.75, 6.2, 0.8, subtitle, 17, t["text"], font=t["fb"])


def cover_bottom(slide, topic, t, subtitle):
    a = t["accent"]
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 5.5, SLIDE_W, 2.0, a)
    add_shape(slide, MSO_SHAPE.OVAL, 10.8, 4.6, 3.6, 3.6, blend(a, WHITE, 0.14))
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0.9, 1.2, 1.2, 0.08, t["accent2"])
    add_text(slide, 0.9, 1.5, 11.4, 3.6, topic, _cover_title_size(topic, 40), t["title"], True,
             MSO_ANCHOR.MIDDLE, font=t["fh"])
    add_text(slide, 0.9, 6.0, 9.5, 0.8, subtitle, 20, t["on_accent"], anchor=MSO_ANCHOR.MIDDLE, font=t["fb"])


def cover_full(slide, topic, t, subtitle):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = t["accent"]
    on = t["on_accent"]
    add_outline(slide, 8.8, 3.6, 6.2, 6.2, blend(t["accent"], on, 0.25), 3, MSO_SHAPE.OVAL)
    add_outline(slide, 9.8, 4.6, 4.2, 4.2, blend(t["accent"], on, 0.25), 3, MSO_SHAPE.OVAL)
    add_text(slide, 0.9, 1.5, 11.0, 3.6, topic, _cover_title_size(topic, 44), on, True,
             MSO_ANCHOR.MIDDLE, font=t["fh"])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0.9, 5.35, 3.0, 0.04, on)
    add_text(slide, 0.9, 5.6, 9.0, 0.7, subtitle, 20, on, font=t["fb"])


COVERS = {
    "panel": cover_panel, "band": cover_band, "center": cover_center, "line": cover_line,
    "frame": cover_frame, "diagonal": cover_diagonal, "circles": cover_circles,
    "bottom": cover_bottom, "full": cover_full,
}


def build_cover(prs, topic: str, style: dict, subtitle: str) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = style["bg"]
    COVERS[style["cover"]](slide, topic[:180], style, subtitle)


# ---------------------------------------------------------
# Ichki slaydlar: bezak, sarlavha va joylashuvlar
# ---------------------------------------------------------
def decor_slide(slide, t):
    name, a, a2 = t["decor"], t["accent"], t["accent2"]
    if name == "topbar":
        add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, 0.16, a)
    elif name == "sidebar":
        add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, 0.26, SLIDE_H, a)
        add_shape(slide, MSO_SHAPE.RECTANGLE, 0.26, 0, 0.07, SLIDE_H, a2)
    elif name == "glow":
        add_shape(slide, MSO_SHAPE.OVAL, 11.3, -1.5, 3.4, 3.4, soft(t, "accent", 0.84))
        add_shape(slide, MSO_SHAPE.OVAL, -1.2, 6.2, 2.6, 2.6, soft(t, "accent2", 0.88))
    elif name == "frame":
        add_outline(slide, 0.25, 0.25, 12.83, 7.0, a, 1.5)
    elif name == "stripes":
        add_shape(slide, MSO_SHAPE.RECTANGLE, 11.2, 0.35, 1.6, 0.07, a)
        add_shape(slide, MSO_SHAPE.RECTANGLE, 11.2, 0.5, 1.1, 0.07, a2)
        add_shape(slide, MSO_SHAPE.RECTANGLE, 11.2, 0.65, 0.6, 0.07, a)
    elif name == "rule":
        add_shape(slide, MSO_SHAPE.RECTANGLE, 0.7, 0.28, 0.6, 0.07, a)
        add_shape(slide, MSO_SHAPE.RECTANGLE, 0.7, 7.0, 11.0, 0.015, hair(t))


def draw_title(slide, title: str, t, x=CONTENT_X, w=11.9):
    mark = t["mark"]
    tx, tw = x, w
    if mark == "dot":
        add_shape(slide, MSO_SHAPE.OVAL, x, 0.86, 0.3, 0.3, t["accent"])
        tx, tw = x + 0.5, w - 0.5
    size = 30 if len(title) <= 45 else 26 if len(title) <= 70 else 22
    add_text(slide, tx, 0.45, tw, 1.1, title, size, t["title"], True, MSO_ANCHOR.MIDDLE, font=t["fh"])
    if mark == "bar":
        add_shape(slide, MSO_SHAPE.RECTANGLE, x, 1.62, 1.2, 0.06, t["accent"])
    elif mark == "line":
        add_shape(slide, MSO_SHAPE.RECTANGLE, x, 1.64, w, 0.015, hair(t))
        add_shape(slide, MSO_SHAPE.RECTANGLE, x, 1.62, 1.2, 0.05, t["accent"])


def lay_cards(slide, bullets, t):
    count = len(bullets)
    gap = 0.18
    card_h = min((AREA_H - gap * (count - 1)) / count, 1.3)
    diameter = min(0.62, card_h - 0.3)
    pad = 0.25
    text_left = pad + diameter + 0.25
    text_w = CONTENT_W - text_left - 0.3
    size = fit_font(bullets, text_w, card_h - 0.08, 24, 14, t["cw"])
    for i, bullet in enumerate(bullets):
        y = AREA_TOP + i * (card_h + gap)
        card = add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, CONTENT_X, y, CONTENT_W, card_h, t["card"], radius=0.18)
        frame = card.text_frame
        frame.word_wrap = True
        frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        frame.margin_left = Inches(text_left)
        frame.margin_right = Inches(0.3)
        frame.margin_top = frame.margin_bottom = Inches(0.04)
        paragraph = frame.paragraphs[0]
        paragraph.alignment = PP_ALIGN.LEFT
        add_run(paragraph, bullet, size, t["text"], font=t["fb"])
        number_circle(slide, CONTENT_X + pad, y + (card_h - diameter) / 2, diameter, str(i + 1), t)


def lay_rows(slide, bullets, t):
    count = len(bullets)
    row_h = min(AREA_H / count, 1.3)
    text_w = CONTENT_W - 1.5
    size = fit_font(bullets, text_w, row_h - 0.1, 24, 14, t["cw"])
    for i, bullet in enumerate(bullets):
        y = AREA_TOP + i * row_h
        add_text(slide, CONTENT_X, y, 1.2, row_h, f"{i + 1:02d}", 30, t["accent"], True,
                 MSO_ANCHOR.MIDDLE, font=t["fh"])
        add_text(slide, CONTENT_X + 1.4, y, text_w, row_h, bullet, size, t["text"], anchor=MSO_ANCHOR.MIDDLE, font=t["fb"])
        if i < count - 1:
            add_shape(slide, MSO_SHAPE.RECTANGLE, CONTENT_X, y + row_h - 0.01, CONTENT_W, 0.015, hair(t))


def lay_grid(slide, bullets, t):
    count = len(bullets)
    spec = {1: [1], 2: [2], 3: [3], 4: [2, 2], 5: [3, 2]}[min(count, 5)]
    gap = 0.25
    rows = len(spec)
    card_h = min((AREA_H - gap * (rows - 1)) / rows, 3.7)
    y0 = AREA_TOP
    if rows == 1:
        card_h = min(card_h, 3.2)
        y0 = AREA_TOP + (AREA_H - card_h) / 2
    idx = 0
    pool_sizes = []
    for r, cols in enumerate(spec):
        card_w = (CONTENT_W - gap * (cols - 1)) / cols
        tall = card_h >= 3.0
        text_w = card_w - 0.5 if tall else card_w - 1.25
        text_h = card_h - 1.25 if tall else card_h - 0.4
        pool_sizes.append((text_w, text_h))
    size = min(fit_font(bullets, w, h, 24, 13, t["cw"]) for w, h in pool_sizes)
    for r, cols in enumerate(spec):
        card_w = (CONTENT_W - gap * (cols - 1)) / cols
        tall = card_h >= 3.0
        for c in range(cols):
            if idx >= count:
                break
            x = CONTENT_X + c * (card_w + gap)
            y = y0 + r * (card_h + gap)
            add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, card_w, card_h, t["card"], radius=0.08)
            add_shape(slide, MSO_SHAPE.RECTANGLE, x + 0.25, y, 0.8, 0.07, t["accent"])
            number_circle(slide, x + 0.25, y + 0.3, 0.5, str(idx + 1), t, size=15)
            if tall:
                add_text(slide, x + 0.25, y + 1.0, card_w - 0.5, card_h - 1.15, bullets[idx], size, t["text"],
                         anchor=MSO_ANCHOR.TOP, font=t["fb"])
            else:
                add_text(slide, x + 1.0, y + 0.2, card_w - 1.25, card_h - 0.4, bullets[idx], size, t["text"],
                         anchor=MSO_ANCHOR.MIDDLE, font=t["fb"])
            idx += 1


def lay_columns(slide, bullets, t):
    count = len(bullets)
    gap = 0.4
    col_w = (CONTENT_W - gap * (count - 1)) / count
    size = fit_font(bullets, col_w, 3.3, 22, 13, t["cw"])
    for i, bullet in enumerate(bullets):
        x = CONTENT_X + i * (col_w + gap)
        add_text(slide, x, 1.95, col_w, 0.95, f"{i + 1:02d}", 44, t["accent"], True, MSO_ANCHOR.MIDDLE, font=t["fh"])
        add_shape(slide, MSO_SHAPE.RECTANGLE, x + 0.1, 3.0, 0.9, 0.07, t["accent2"])
        add_text(slide, x, 3.25, col_w, 3.4, bullet, size, t["text"], font=t["fb"])
        if i < count - 1:
            add_shape(slide, MSO_SHAPE.RECTANGLE, x + col_w + gap / 2, 2.1, 0.015, 4.4, hair(t))


def lay_timeline(slide, bullets, t):
    count = len(bullets)
    step_h = min(AREA_H / count, 1.35)
    first_c = AREA_TOP + step_h / 2
    last_c = AREA_TOP + step_h * (count - 0.5)
    add_shape(slide, MSO_SHAPE.RECTANGLE, 1.175 - 0.025, first_c, 0.05, max(0.01, last_c - first_c), t["accent2"])
    text_w = CONTENT_W - 1.4
    size = fit_font(bullets, text_w, step_h - 0.1, 24, 14, t["cw"])
    for i, bullet in enumerate(bullets):
        yc = AREA_TOP + step_h * (i + 0.5)
        number_circle(slide, 1.175 - 0.28, yc - 0.28, 0.56, str(i + 1), t)
        add_text(slide, CONTENT_X + 1.2, yc - step_h / 2, text_w, step_h, bullet, size, t["text"],
                 anchor=MSO_ANCHOR.MIDDLE, font=t["fb"])


def lay_bignum(slide, bullets, t):
    hero, rest = bullets[0], bullets[1:]
    card_w = 5.7 if rest else CONTENT_W
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, CONTENT_X, AREA_TOP, card_w, AREA_H, t["accent"], radius=0.06)
    add_text(slide, CONTENT_X + 0.35, AREA_TOP + 0.05, 1.5, 1.2, "\u201c", 96, blend(t["accent"], t["on_accent"], 0.45),
             True, MSO_ANCHOR.TOP, font=t["fh"])
    size = fit_font([hero], card_w - 0.9, AREA_H - 1.7, 32, 18, t["cw"] * 1.15)
    add_text(slide, CONTENT_X + 0.4, AREA_TOP + 1.2, card_w - 0.8, AREA_H - 1.5, hero, size, t["on_accent"], True,
             MSO_ANCHOR.TOP, font=t["fb"])
    if not rest:
        return
    x0 = CONTENT_X + card_w + 0.45
    w = SLIDE_W - 0.7 - x0
    step_h = min(AREA_H / len(rest), 1.4)
    size = fit_font(rest, w - 0.5, step_h - 0.1, 22, 13, t["cw"])
    for i, bullet in enumerate(rest):
        y = AREA_TOP + i * step_h
        add_shape(slide, MSO_SHAPE.RECTANGLE, x0, y + step_h / 2 - 0.1, 0.2, 0.2, t["accent2"])
        add_text(slide, x0 + 0.4, y, w - 0.4, step_h, bullet, size, t["text"], anchor=MSO_ANCHOR.MIDDLE, font=t["fb"])
        if i < len(rest) - 1:
            add_shape(slide, MSO_SHAPE.RECTANGLE, x0, y + step_h - 0.01, w, 0.015, hair(t))


def lay_split(slide, title, bullets, t):
    a = t["accent"]
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, 4.7, SLIDE_H, a)
    add_shape(slide, MSO_SHAPE.OVAL, -1.1, 5.0, 3.8, 3.8, blend(a, WHITE, 0.14))
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0.55, 1.25, 0.9, 0.08, t["accent2"])
    size = fit_font([title], 3.9, 4.3, 34, 20, t["cw"] * 1.15)
    add_text(slide, 0.5, 1.45, 3.9, 4.3, title, size, t["on_accent"], True, MSO_ANCHOR.TOP, font=t["fh"])
    count = len(bullets)
    step_h = min(AREA_H / count, 1.3)
    x0, w = 5.3, SLIDE_W - 0.7 - 5.3
    top = (SLIDE_H - step_h * count) / 2
    size = fit_font(bullets, w - 0.4, step_h - 0.1, 24, 14, t["cw"])
    for i, bullet in enumerate(bullets):
        y = top + i * step_h
        add_shape(slide, MSO_SHAPE.RECTANGLE, x0, y + step_h * 0.2, 0.07, step_h * 0.6, a)
        add_text(slide, x0 + 0.3, y, w - 0.3, step_h, bullet, size, t["text"], anchor=MSO_ANCHOR.MIDDLE, font=t["fb"])


def build_content(prs, data: dict, style: dict, page: int, total: int, index: int = 0) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = style["bg"]
    layout = style["layouts"][index % len(style["layouts"])]
    bullets = data["content"][:MAX_BULLETS]
    if layout == "columns" and len(bullets) >= 5:
        layout = "rows"  # 5 ta tor ustunga uzun matn sig'maydi

    if layout == "split":
        lay_split(slide, data["title"], bullets, style)
    else:
        decor_slide(slide, style)
        draw_title(slide, data["title"], style)
        {
            "cards": lay_cards, "rows": lay_rows, "grid": lay_grid, "columns": lay_columns,
            "timeline": lay_timeline, "bignum": lay_bignum,
        }[layout](slide, bullets, style)

    add_text(slide, 11.4, 7.02, 1.3, 0.3, f"{page} / {total}", 11,
             blend(style["text"], style["bg"], 0.35), align=PP_ALIGN.RIGHT, font=style["fb"])


def generate_pptx(topic: str, slide_data: list[dict], template_style: str, output_path: str,
                  alphabet: str = LATIN) -> str:
    topic = fix_text(topic, alphabet)
    slide_data = [
        {"title": fix_text(item["title"], alphabet), "content": [fix_text(b, alphabet) for b in item["content"]]}
        for item in slide_data
    ]
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    style = THEMES.get(template_style) or THEMES["classic"]
    total = len(slide_data) + 1

    subtitle = "AI ёрдамида тайёрланган тақдимот" if alphabet == KIRILL else "AI yordamida tayyorlangan taqdimot"
    build_cover(prs, topic, style, subtitle)
    for index, data in enumerate(slide_data):
        build_content(prs, data, style, index + 2, total, index)

    prs.save(output_path)
    return output_path


def theme_catalog() -> list[dict]:
    """Mini App uchun shablonlar ro'yxati (rang va nomlar)."""
    def hex_(color) -> str:
        return "#%02x%02x%02x" % (color[0], color[1], color[2])

    return [
        {
            "key": key,
            "name": theme["name"],
            "bg": hex_(theme["bg"]),
            "title": hex_(theme["title"]),
            "text": hex_(theme["text"]),
            "accent": hex_(theme["accent"]),
            "accent2": hex_(theme["accent2"]),
            "card": hex_(theme["card"]),
            "cover": theme["cover"],
            "layouts": theme["layouts"],
        }
        for key, theme in THEMES.items()
    ]


# =========================================================
# BOT HANDLERS
# =========================================================
BTN_SLIDE = "📕 Yangi taqdimot (Slayd)"
BTN_SELF = "📘 Yangi mustaqil ish"
BTN_REFERAT = "📄 Yangi referat"
BTN_COURSE = "📚 Kurs ishi yaratish"
BTN_TEST = "📝 Test yaratish"
BTN_ARTICLE = "📰 Maqola yaratish"
BTN_THESIS = "📑 Tezis yaratish"
BTN_CROSSWORD = "🧩 Krossvord yaratish"
BTN_ALPHABET = "🔤 Alifbo"
BTN_LIMIT = "📊 Limitim"
BTN_HELP = "❓ Yordam"
BTN_CANCEL = "❌ Bekor qilish"

KIND_BY_BUTTON = {
    BTN_REFERAT: "referat",
    BTN_SELF: "mustaqil",
    BTN_COURSE: "kurs",
    BTN_TEST: "test",
    BTN_ARTICLE: "maqola",
    BTN_THESIS: "tezis",
    BTN_CROSSWORD: "krossvord",
}
TOPIC_HINTS = {
    "referat": "Referat mavzusini yozing",
    "mustaqil": "Mustaqil ish mavzusini yozing",
    "kurs": "Kurs ishi mavzusini yozing",
    "test": "Test mavzusini yozing",
    "maqola": "Ilmiy maqola mavzusini yozing",
    "tezis": "Tezis mavzusini yozing",
    "krossvord": "Krossvord mavzusini yozing",
}
ALPHABET_NAMES = {LATIN: "Lotin", KIRILL: "Кирилл"}


class DocOrder(StatesGroup):
    waiting_topic = State()


def main_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text=BTN_SLIDE)],
            [KeyboardButton(text=BTN_SELF), KeyboardButton(text=BTN_REFERAT)],
            [KeyboardButton(text=BTN_COURSE), KeyboardButton(text=BTN_TEST)],
            [KeyboardButton(text=BTN_ARTICLE), KeyboardButton(text=BTN_THESIS)],
            [KeyboardButton(text=BTN_CROSSWORD), KeyboardButton(text=BTN_ALPHABET)],
            [KeyboardButton(text=BTN_LIMIT), KeyboardButton(text=BTN_HELP)],
        ],
        resize_keyboard=True,
        input_field_placeholder="Pastdagi tugmalardan birini tanlang",
    )


def cancel_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text=BTN_CANCEL)]],
        resize_keyboard=True,
        input_field_placeholder="Mavzuni yozing",
    )


def open_app_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[[
            InlineKeyboardButton(
                text="✨ Mini App’ni ochish",
                web_app=WebAppInfo(url=WEBAPP_URL),
            )
        ]]
    )


def format_window(seconds: int) -> str:
    minutes = max(1, seconds // 60)
    if minutes >= 60 and minutes % 60 == 0:
        return f"{minutes // 60} soat"
    return f"{minutes} daqiqa"


@dp.message(CommandStart())
async def start_cmd(message: types.Message, state: FSMContext):
    await state.clear()
    if message.from_user:
        await track_user(message.from_user.id, message.from_user.first_name, message.from_user.username)
    name = html.escape(message.from_user.first_name) if message.from_user else "do'st"
    await message.answer(
        f"👋 <b>Xush kelibsiz, {name}!</b>\n\n"
        "Men o'zbek tilida tayyor ish va taqdimotlar tayyorlayman:\n\n"
        "📕 <b>Taqdimot (Slayd)</b> — PowerPoint fayl, 4 xil uslub (Mini App orqali).\n"
        "📘 <b>Mustaqil ish</b> — reja, asosiy qism, xulosa va nazorat savollari.\n"
        "📄 <b>Referat</b> — titul varaq, reja va akademik matn.\n"
        "📚 <b>Kurs ishi</b> — reja, kirish, 3 bob, xulosa va adabiyotlar.\n"
        "📰 <b>Ilmiy maqola</b> va 📑 <b>Tezis</b> — konferensiya va jurnallar uchun.\n"
        "📝 <b>Test</b> — 10 ta savol va javoblar kaliti.\n"
        "🧩 <b>Krossvord</b> — savollar va javoblar bilan, chop etishga tayyor.\n\n"
        "🔤 <b>Alifbo</b> tugmasi orqali fayllarni lotin yoki kirill yozuvida olasiz.\n"
        "📊 Qancha so'rov qolganini <b>Limitim</b> orqali ko'rasiz.\n"
        "👇 Pastdagi tugmalardan birini tanlang.",
        reply_markup=main_keyboard(),
        parse_mode="HTML",
    )


@dp.message(Command("cancel"))
@dp.message(F.text == BTN_CANCEL)
async def cancel_cmd(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer("Bekor qilindi. Kerakli bo'limni tanlang 👇", reply_markup=main_keyboard())


@dp.message(F.text == BTN_SLIDE)
async def slide_btn(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer(
        "Mini App’ni oching, mavzuni yozing yoki tayyor mavzudan tanlang, uslubni belgilang 👇",
        reply_markup=open_app_keyboard(),
    )


@dp.message(F.text == BTN_LIMIT)
async def limit_btn(message: types.Message, state: FSMContext):
    await state.clear()
    remaining = await asyncio.to_thread(rate_limiter.remaining, message.from_user.id)
    await message.answer(
        f"📊 Sizda <b>{remaining}</b> / {RATE_LIMIT_REQUESTS} ta so'rov qoldi.\n"
        f"Limit oxirgi {format_window(RATE_LIMIT_WINDOW_SECONDS)} bo'yicha hisoblanadi. "
        "Slayd, referat, mustaqil ish va test uchun umumiy.",
        parse_mode="HTML",
        reply_markup=main_keyboard(),
    )


@dp.message(Command("help"))
@dp.message(F.text == BTN_HELP)
async def help_btn(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer(
        "❓ <b>Qanday ishlaydi?</b>\n\n"
        "1. Pastdagi tugmalardan kerakli xizmatni tanlang.\n"
        "2. Mavzuni yozing.\n"
        "3. 30-90 soniyada tayyor fayl shu chatga keladi.\n\n"
        "📱 <b>Slayd faylini</b> PowerPoint, WPS Office yoki Google Slides orqali oching: "
        "telefonning ichki ko'rish dasturi ba'zan bo'sh ko'rsatadi.\n"
        "📄 <b>Word fayllarni</b> Word, WPS Office yoki Google Docs bilan oching.\n"
        "🔤 Fayl alifbosini <b>Alifbo</b> tugmasi orqali o'zgartirasiz.\n"
        "📚 Kurs ishi 2-4 daqiqa vaqt olishi mumkin.\n\n"
        "⚠️ AI xato qilishi mumkin: faktlarni va nomlarni topshirishdan oldin tekshiring."
        + (f"\n\n📩 Savol yoki taklif: {html.escape(SUPPORT_CONTACT)}" if SUPPORT_CONTACT else ""),
        parse_mode="HTML",
        reply_markup=main_keyboard(),
    )


def alphabet_keyboard(current: str) -> InlineKeyboardMarkup:
    def mark(code: str) -> str:
        return ("✅ " if current == code else "") + ALPHABET_NAMES[code]

    return InlineKeyboardMarkup(
        inline_keyboard=[[
            InlineKeyboardButton(text=mark(LATIN), callback_data="alpha:latin"),
            InlineKeyboardButton(text=mark(KIRILL), callback_data="alpha:kirill"),
        ]]
    )


def alphabet_text(current: str) -> str:
    return (
        "🔤 <b>Alifbo</b>\n\n"
        "Tayyor fayllar (slayd va Word) qaysi alifboda bo'lsin?\n"
        f"Hozirgi tanlov: <b>{ALPHABET_NAMES[current]}</b>"
    )


@dp.message(Command("alifbo"))
@dp.message(F.text == BTN_ALPHABET)
async def alphabet_btn(message: types.Message, state: FSMContext):
    await state.clear()
    current = await get_user_alphabet(message.from_user.id)
    await message.answer(
        alphabet_text(current),
        parse_mode="HTML",
        reply_markup=alphabet_keyboard(current),
    )


@dp.callback_query(F.data.in_({"alpha:latin", "alpha:kirill"}))
async def alphabet_choice(callback: types.CallbackQuery):
    code = normalize_alphabet((callback.data or "").split(":", 1)[-1])
    user = callback.from_user
    await track_user(user.id, user.first_name, user.username)
    await set_user_alphabet(user.id, code)
    if callback.message:
        try:
            await callback.message.edit_text(
                alphabet_text(code), parse_mode="HTML", reply_markup=alphabet_keyboard(code)
            )
        except Exception:
            pass
    await callback.answer(f"Tanlandi: {ALPHABET_NAMES[code]}")


def pages_keyboard(kind: str) -> InlineKeyboardMarkup:
    buttons = [
        InlineKeyboardButton(text=f"{lo}–{hi} bet", callback_data=f"pg:{kind}:{lo}:{hi}")
        for lo, hi in PAGE_TIERS[kind]
    ]
    rows = [buttons[i:i + 2] for i in range(0, len(buttons), 2)]
    return InlineKeyboardMarkup(inline_keyboard=rows)


def topic_prompt_text(kind: str) -> str:
    return f"✍️ {TOPIC_HINTS[kind]}.\nMasalan: <i>Sun'iy intellektning ta'limdagi o'rni</i>"


@dp.message(F.text.in_(set(KIND_BY_BUTTON)))
async def document_btn(message: types.Message, state: FSMContext):
    kind = KIND_BY_BUTTON[message.text]
    if kind in PAGE_TIERS:
        # Avval hajm (bet) tanlanadi, keyin mavzu so'raladi.
        await state.clear()
        await message.answer(
            f"📏 <b>{DOC_KINDS[kind]['label']}</b> hajmini tanlang:",
            parse_mode="HTML",
            reply_markup=pages_keyboard(kind),
        )
        return
    await state.set_state(DocOrder.waiting_topic)
    await state.update_data(kind=kind, pages=None)
    await message.answer(
        topic_prompt_text(kind),
        parse_mode="HTML",
        reply_markup=cancel_keyboard(),
    )


@dp.callback_query(F.data.startswith("pg:"))
async def pages_choice(callback: types.CallbackQuery, state: FSMContext):
    parts = (callback.data or "").split(":")
    try:
        kind, lo, hi = parts[1], int(parts[2]), int(parts[3])
    except (IndexError, ValueError):
        await callback.answer("Noto'g'ri tanlov", show_alert=True)
        return
    if kind not in PAGE_TIERS or (lo, hi) not in PAGE_TIERS[kind]:
        await callback
