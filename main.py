import asyncio
import hashlib
import hmac
import json
import logging
import os
import time
import tempfile
import uuid
from contextlib import asynccontextmanager
from pathlib import Path
from collections import defaultdict, deque
from typing import Literal

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field, ConfigDict
from aiogram import Bot, Dispatcher, types
from aiogram.filters import CommandStart
from aiogram.types import WebAppInfo, InlineKeyboardMarkup, InlineKeyboardButton
from google import genai
from google.genai import types as genai_types
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor


# =========================================================
# CONFIGURATION
# Railway Variables: BOT_TOKEN, GEMINI_API_KEY, WEBAPP_URL
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
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")
MAX_CONCURRENT_GENERATIONS = int(os.getenv("MAX_CONCURRENT_GENERATIONS", "3"))
RATE_LIMIT_REQUESTS = int(os.getenv("RATE_LIMIT_REQUESTS", "5"))
RATE_LIMIT_WINDOW_SECONDS = int(os.getenv("RATE_LIMIT_WINDOW_SECONDS", "3600"))
MAX_TOPIC_LENGTH = 200
TELEGRAM_INIT_DATA_MAX_AGE = int(os.getenv("TELEGRAM_INIT_DATA_MAX_AGE", "86400"))

if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN Railway Variables bo'limida sozlanmagan.")
if not GEMINI_API_KEY:
    raise RuntimeError("GEMINI_API_KEY Railway Variables bo'limida sozlanmagan.")

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
STATIC_DIR.mkdir(parents=True, exist_ok=True)
INDEX_FILE = STATIC_DIR / "index.html"

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
gemini_client = genai.Client(api_key=GEMINI_API_KEY)
generation_semaphore = asyncio.Semaphore(MAX_CONCURRENT_GENERATIONS)

_request_history = defaultdict(deque)
_rate_lock = asyncio.Lock()

TEMPLATE_STYLES = {
    "classic": {
        "bg_color": RGBColor(255, 255, 255),
        "title_color": RGBColor(20, 20, 20),
        "text_color": RGBColor(60, 60, 60),
    },
    "business": {
        "bg_color": RGBColor(240, 244, 248),
        "title_color": RGBColor(30, 58, 138),
        "text_color": RGBColor(31, 41, 55),
    },
    "dark": {
        "bg_color": RGBColor(24, 24, 27),
        "title_color": RGBColor(255, 255, 255),
        "text_color": RGBColor(228, 228, 231),
    },
    "minimal": {
        "bg_color": RGBColor(240, 253, 244),
        "title_color": RGBColor(22, 101, 52),
        "text_color": RGBColor(20, 83, 45),
    },
}


# =========================================================
# REQUEST MODELS
# =========================================================
class SlideRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    init_data: str = Field(min_length=1, max_length=10_000)
    topic: str = Field(min_length=3, max_length=MAX_TOPIC_LENGTH)
    template: Literal["classic", "business", "dark", "minimal"] = "classic"


# =========================================================
# TELEGRAM MINI APP AUTHENTICATION
# =========================================================
def verify_telegram_init_data(init_data: str) -> dict:
    from urllib.parse import parse_qsl

    try:
        pairs = parse_qsl(init_data, keep_blank_values=True, strict_parsing=True)
        data = dict(pairs)
    except (ValueError, TypeError):
        raise HTTPException(status_code=401, detail="Telegram tasdiqlash ma'lumoti noto'g'ri.")

    received_hash = data.pop("hash", None)
    if not received_hash:
        raise HTTPException(status_code=401, detail="Telegram tasdiqlash hash'i topilmadi.")

    data_check_string = "\n".join(
        f"{key}={value}" for key, value in sorted(data.items())
    )
    secret_key = hmac.new(
        b"WebAppData",
        BOT_TOKEN.encode("utf-8"),
        hashlib.sha256,
    ).digest()
    calculated_hash = hmac.new(
        secret_key,
        data_check_string.encode("utf-8"),
        hashlib.sha256,
    ).hexdigest()

    if not hmac.compare_digest(calculated_hash, received_hash):
        raise HTTPException(status_code=401, detail="Telegram ma'lumoti tasdiqlanmadi.")

    try:
        auth_date = int(data.get("auth_date", "0"))
    except ValueError:
        raise HTTPException(status_code=401, detail="auth_date noto'g'ri.")

    now = int(time.time())
    if auth_date <= 0 or auth_date > now + 60 or now - auth_date > TELEGRAM_INIT_DATA_MAX_AGE:
        raise HTTPException(status_code=401, detail="Telegram sessiyasi eskirgan. Mini App'ni qayta oching.")

    try:
        user = json.loads(data.get("user", "{}"))
        user_id = int(user["id"])
    except (ValueError, TypeError, KeyError, json.JSONDecodeError):
        raise HTTPException(status_code=401, detail="Telegram foydalanuvchisi aniqlanmadi.")

    return {"user_id": user_id, "user": user}


async def enforce_rate_limit(user_id: int) -> None:
    now = time.monotonic()
    cutoff = now - RATE_LIMIT_WINDOW_SECONDS
    async with _rate_lock:
        history = _request_history[user_id]
        while history and history[0] < cutoff:
            history.popleft()
        if len(history) >= RATE_LIMIT_REQUESTS:
            raise HTTPException(
                status_code=429,
                detail="So'rovlar limiti tugadi. Keyinroq qayta urinib ko'ring.",
            )
        history.append(now)


# =========================================================
# AI CONTENT GENERATION
# =========================================================
async def generate_slide_content(topic: str) -> list[dict]:
    prompt = f"""
Siz o'zbek tilida sifatli PowerPoint taqdimotlar tayyorlaydigan mutaxassissiz.
Quyidagi mavzu bo'yicha AYNAN 5 ta slayd uchun ma'lumot tayyorlang: {topic!r}

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
- Aynan 5 ta slayd.
- Har slaydda bitta qisqa sarlavha va 3-5 ta mazmunli punkt.
- Har bir punkt qisqa, tushunarli o'zbek tilida bo'lsin.
- Birinchi slayd mavzuga kirish, oxirgi slayd xulosa yoki asosiy tavsiyalar bo'lsin.
- To'qima statistika yoki manbasi noma'lum aniq raqamlarni kiritmang.
"""
    async with generation_semaphore:
        try:
            response = await gemini_client.aio.models.generate_content(
                model=GEMINI_MODEL,
                contents=prompt,
                config=genai_types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.5,
                ),
            )
        except Exception as exc:
            logger.exception("Gemini API so'rovi muvaffaqiyatsiz tugadi")
            raise HTTPException(
                status_code=502,
                detail="AI hozir javob bera olmadi. Birozdan keyin qayta urinib ko'ring.",
            ) from exc

    raw_text = getattr(response, "text", None)
    if not raw_text:
        raise HTTPException(status_code=502, detail="AI bo'sh javob qaytardi.")

    try:
        payload = json.loads(raw_text)
        slides = payload["slides"]
    except (json.JSONDecodeError, KeyError, TypeError):
        logger.error("Gemini javobi kutilgan JSON formatida emas: %r", raw_text[:500])
        raise HTTPException(status_code=502, detail="AI javobini o'qib bo'lmadi. Qayta urinib ko'ring.")

    if not isinstance(slides, list) or len(slides) != 5:
        raise HTTPException(status_code=502, detail="AI aynan 5 ta slayd yaratmadi. Qayta urinib ko'ring.")

    validated = []
    for slide in slides:
        if not isinstance(slide, dict):
            raise HTTPException(status_code=502, detail="AI slayd ma'lumotlari noto'g'ri.")
        title = str(slide.get("title", "")).strip()[:160]
        content = slide.get("content", [])
        if not title or not isinstance(content, list):
            raise HTTPException(status_code=502, detail="Slayd sarlavhasi yoki matni noto'g'ri.")
        bullets = [str(item).strip()[:350] for item in content if str(item).strip()]
        if not bullets:
            raise HTTPException(status_code=502, detail="Slaydlardan birida matn yo'q.")
        validated.append({"title": title, "content": bullets[:6]})
    return validated


# =========================================================
# POWERPOINT GENERATION
# =========================================================
def generate_pptx(topic: str, slide_data: list[dict], template_style: str, output_path: str) -> str:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    style = TEMPLATE_STYLES[template_style]

    cover = prs.slides.add_slide(prs.slide_layouts[6])
    cover.background.fill.solid()
    cover.background.fill.fore_color.rgb = style["bg_color"]
    title_box = cover.shapes.add_textbox(Inches(0.9), Inches(1.7), Inches(11.5), Inches(2.0))
    title_frame = title_box.text_frame
    title_frame.word_wrap = True
    title_frame.text = topic[:180]
    title_p = title_frame.paragraphs[0]
    title_p.font.bold = True
    title_p.font.size = Pt(32)
    title_p.font.color.rgb = style["title_color"]
    subtitle = cover.shapes.add_textbox(Inches(0.95), Inches(4.0), Inches(11.3), Inches(0.8))
    subtitle.text_frame.text = "AI yordamida tayyorlangan taqdimot"
    subtitle.text_frame.paragraphs[0].font.size = Pt(18)
    subtitle.text_frame.paragraphs[0].font.color.rgb = style["text_color"]

    for index, data in enumerate(slide_data, start=1):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = style["bg_color"]

        title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.45), Inches(12.0), Inches(1.15))
        title_frame = title_box.text_frame
        title_frame.word_wrap = True
        title_frame.text = data["title"]
        title_p = title_frame.paragraphs[0]
        title_p.font.bold = True
        title_p.font.size = Pt(28 if len(data["title"]) < 55 else 23)
        title_p.font.color.rgb = style["title_color"]

        body_box = slide.shapes.add_textbox(Inches(0.9), Inches(1.85), Inches(11.6), Inches(4.9))
        body_frame = body_box.text_frame
        body_frame.clear()
        body_frame.word_wrap = True
        body_frame.margin_left = Inches(0.12)
        body_frame.margin_right = Inches(0.12)
        body_frame.margin_top = Inches(0.08)

        bullets = data["content"][:6]
        font_size = 21 if len(bullets) <= 4 else 18
        for i, bullet in enumerate(bullets):
            paragraph = body_frame.paragraphs[0] if i == 0 else body_frame.add_paragraph()
            paragraph.text = bullet
            paragraph.level = 0
            paragraph.font.size = Pt(font_size)
            paragraph.font.color.rgb = style["text_color"]
            paragraph.space_after = Pt(14)

        footer = slide.shapes.add_textbox(Inches(11.8), Inches(7.0), Inches(0.8), Inches(0.25))
        footer.text_frame.text = str(index + 1)
        footer.text_frame.paragraphs[0].font.size = Pt(10)
        footer.text_frame.paragraphs[0].font.color.rgb = style["text_color"]

    prs.save(output_path)
    return output_path


# =========================================================
# BOT HANDLERS
# =========================================================
@dp.message(CommandStart())
async def start_cmd(message: types.Message):
    keyboard = InlineKeyboardMarkup(
        inline_keyboard=[[
            InlineKeyboardButton(
                text="✨ Mini App’da Slayd Yaratish",
                web_app=WebAppInfo(url=WEBAPP_URL),
            )
        ]]
    )
    await message.answer(
        "👋 <b>Xush kelibsiz!</b>\n\n"
        "PowerPoint slaydlarini yaratish uchun pastdagi "
        "<b>Mini App</b> tugmasini bosing!",
        reply_markup=keyboard,
        parse_mode="HTML",
    )


# =========================================================
# FASTAPI LIFESPAN
# =========================================================
@asynccontextmanager
async def lifespan(app: FastAPI):
    polling_task = None
    try:
        await bot.delete_webhook(drop_pending_updates=False)
        polling_task = asyncio.create_task(dp.start_polling(bot, handle_signals=False))
        logger.info("FastAPI ishga tushdi; bot polling vazifasi boshlandi.")
        yield
    finally:
        if polling_task:
            polling_task.cancel()
            try:
                await polling_task
            except asyncio.CancelledError:
                pass
            except Exception:
                logger.exception("Bot polling to'xtashida xatolik")
        await bot.session.close()
        logger.info("Ilova to'xtatildi.")


app = FastAPI(title="Ziyokulov Slide AI", version="2.0.0", lifespan=lifespan)

allowed_origin = os.getenv("MINI_APP_ORIGIN", WEBAPP_URL).strip().rstrip("/")
app.add_middleware(
    CORSMiddleware,
    allow_origins=[allowed_origin],
    allow_credentials=False,
    allow_methods=["GET", "POST"],
    allow_headers=["Content-Type"],
)

app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.get("/")
async def serve_index():
    if not INDEX_FILE.is_file():
        raise HTTPException(
            status_code=503,
            detail="static/index.html topilmadi.",
        )
    return FileResponse(str(INDEX_FILE), media_type="text/html")


@app.post("/api/generate-slide")
async def api_generate_slide(req: SlideRequest):
    auth = verify_telegram_init_data(req.init_data)
    user_id = auth["user_id"]
    await enforce_rate_limit(user_id)

    slide_data = await generate_slide_content(req.topic)
    file_path = os.path.join(tempfile.gettempdir(), f"slide_{uuid.uuid4().hex}.pptx")

    try:
        await asyncio.to_thread(
            generate_pptx,
            req.topic,
            slide_data,
            req.template,
            file_path,
        )
        document = types.FSInputFile(file_path, filename=f"taqdimot_{uuid.uuid4().hex[:8]}.pptx")
        try:
            await bot.send_document(
                chat_id=user_id,
                document=document,
                caption=f"🎉 Slaydingiz tayyor!\n📌 Mavzu: {req.topic}\n🎨 Uslub: {req.template.capitalize()}",
            )
        except Exception as exc:
            logger.exception("Tayyor faylni Telegram foydalanuvchisiga yuborib bo'lmadi")
            raise HTTPException(
                status_code=409,
                detail="Faylni Telegram'ga yuborib bo'lmadi. Botga /start yuboring.",
            ) from exc

        return {
            "success": True,
            "message": "Slayd Telegram orqali yuborildi.",
        }
    except HTTPException:
        raise
    except Exception as exc:
        logger.exception("Slayd yaratishda kutilmagan xatolik")
        raise HTTPException(
            status_code=500,
            detail="Slayd yaratishda xatolik yuz berdi.",
        ) from exc
    finally:
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
        except OSError:
            logger.warning("Vaqtinchalik faylni o'chirib bo'lmadi: %s", file_path)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=PORT, workers=1)
