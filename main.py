 import os
import logging
import asyncio
from typing import List
from fastapi import FastAPI, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from aiogram import Bot, Dispatcher, types
from aiogram.filters import CommandStart
from aiogram.types import WebAppInfo, InlineKeyboardMarkup, InlineKeyboardButton
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from google import genai

# =========================================================
# 1. SOZLAMALAR VA ENVIRONMENT O'ZGARUVCHILARI
# =========================================================
BOT_TOKEN = os.getenv("BOT_TOKEN", "YOUR_TELEGRAM_BOT_TOKEN")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "YOUR_GEMINI_API_KEY")
WEBAPP_URL = os.getenv("WEBAPP_URL", "https://your-railway-app.up.railway.app")

# Google Gemini AI ni to'g'ridan-to'g'ri API kalit bilan ishga tushirish
client = genai.Client(api_key=GEMINI_API_KEY)

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Frontend fayllarni ulash
if not os.path.exists("static"):
    os.makedirs("static")
app.mount("/static", StaticFiles(directory="static"), name="static")

logging.basicConfig(level=logging.INFO)

# =========================================================
# 2. SLAYD RAMKALARI VA RANGLAR PALITRASI
# =========================================================
TEMPLATE_STYLES = {
    "classic": {
        "bg_color": RGBColor(255, 255, 255),
        "title_color": RGBColor(20, 20, 20),
        "text_color": RGBColor(60, 60, 60)
    },
    "business": {
        "bg_color": RGBColor(240, 244, 248),
        "title_color": RGBColor(30, 58, 138),
        "text_color": RGBColor(31, 41, 55)
    },
    "dark": {
        "bg_color": RGBColor(24, 24, 27),
        "title_color": RGBColor(255, 255, 255),
        "text_color": RGBColor(228, 228, 231)
    },
    "minimal": {
        "bg_color": RGBColor(240, 253, 244),
        "title_color": RGBColor(22, 101, 52),
        "text_color": RGBColor(20, 83, 45)
    }
}

# =========================================================
# 3. PPTX SLAYD YARATISH FUNKSIYASI
# =========================================================
def generate_pptx(topic: str, slide_data: list, template_style: str = "classic") -> str:
    prs = Presentation()
    style = TEMPLATE_STYLES.get(template_style, TEMPLATE_STYLES["classic"])
    
    for data in slide_data:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = style["bg_color"]

        tx_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(8.4), Inches(1))
        tf = tx_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = data.get("title", "Taqdimot")
        p.font.bold = True
        p.font.size = Pt(28)
        p.font.color.rgb = style["title_color"]
        
        tx_box_body = slide.shapes.add_textbox(Inches(0.8), Inches(2.0), Inches(8.4), Inches(4.5))
        tf_body = tx_box_body.text_frame
        tf_body.word_wrap = True
        
        bullets = data.get("content", [])
        for i, bullet in enumerate(bullets):
            p_body = tf_body.add_paragraph() if i > 0 else tf_body.paragraphs[0]
            p_body.text = f"• {bullet}"
            p_body.font.size = Pt(18)
            p_body.font.color.rgb = style["text_color"]
            p_body.space_after = Pt(12)

    safe_topic = "".join(c for c in topic if c.isalnum() or c in (' ', '_')).rstrip()
    file_name = f"slide_{safe_topic.replace(' ', '_')}.pptx"
    prs.save(file_name)
    return file_name

# =========================================================
# 4. FASTAPI API ENDPOINT (MINI APP UCHUN)
# =========================================================
class SlideRequest(BaseModel):
    user_id: int
    topic: str
    template: str = "classic"

@app.post("/api/generate-slide")
async def api_generate_slide(req: SlideRequest):
    try:
        prompt = f"""
        "{req.topic}" mavzusida 5 ta slayddan iborat taqdimot matnini o'zbek tilida tayyorlang.
        Javobni quyidagi aniq va qat'iy strukturada bering (ortiqcha kirish so'zlar yozmang):

        SLAYD 1
        Sarlavha: [Slayd sarlavhasi]
        - [Nuqta 1]
        - [Nuqta 2]
        - [Nuqta 3]

        SLAYD 2
        Sarlavha: [Slayd sarlavhasi]
        - [Nuqta 1]
        - [Nuqta 2]
        - [Nuqta 3]
        """
        
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=prompt,
        )
        ai_text = response.text

        slides_raw = ai_text.split("SLAYD")
        slide_list = []

        for slide_block in slides_raw:
            if not slide_block.strip():
                continue
            lines = [line.strip() for line in slide_block.strip().split("\n") if line.strip()]
            title = "Taqdimot"
            bullets = []
            for line in lines:
                if line.startswith("Sarlavha:"):
                    title = line.replace("Sarlavha:", "").strip()
                elif line.startswith("-"):
                    bullets.append(line.replace("-", "").strip())
            if bullets or title != "Taqdimot":
                slide_list.append({"title": title, "content": bullets})

        file_path = generate_pptx(req.topic, slide_list, req.template)

        file_to_send = types.FSInputFile(file_path)
        await bot.send_document(
            chat_id=req.user_id,
            document=file_to_send,
            caption=f"🎉 **Slaydingiz tayyor!**\n📌 Mavzu: *{req.topic}*\n🎨 Ramka stili: *{req.template.capitalize()}*"
        )

        if os.path.exists(file_path):
            os.remove(file_path)

        return {"success": True, "message": "Slayd muvaffaqiyatli yuborildi!"}

    except Exception as e:
        logging.error(f"Xatolik yuz berdi: {e}")
        raise HTTPException(status_code=500, detail=str(e))

# =========================================================
# 5. TELEGRAM BOT HANDLERLARI
# =========================================================
@dp.message(CommandStart())
async def start_cmd(message: types.Message):
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(
            text="✨ Mini App’da Slayd Yaratish", 
            web_app=WebAppInfo(url=f"{WEBAPP_URL}/static/index.html")
        )]
    ])
    await message.answer(
        "👋 **Xush kelibsiz!**\n\n"
        "PowerPoint slaydlarini yaratish uchun pastdagi **Mini App** tugmasini bosing!",
        reply_markup=keyboard
    )

@app.on_event("startup")
async def on_startup():
    asyncio.create_task(dp.start_polling(bot))
