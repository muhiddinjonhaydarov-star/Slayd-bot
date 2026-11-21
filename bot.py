import asyncio
import io
import json
import logging
import os
import re
import sqlite3
import threading
import time
from contextlib import closing
from datetime import datetime
from http.server import BaseHTTPRequestHandler, HTTPServer

import httpx
from dotenv import load_dotenv
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, Update
from telegram.ext import (
    ApplicationBuilder,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)
from telegram.request import HTTPXRequest

load_dotenv()
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logging.getLogger("httpx").setLevel(logging.WARNING)
log = logging.getLogger("slayd-bot")

# Gemini ulanishi: IPv4 majburlanadi (Telegram uchun ham shunday qilingan) va vaqt chegarasi bor,
# shuning uchun so'rov "osilib" qolsa ham bot cheksiz kutib o'tirmaydi.
AI_TIMEOUT_MS = int(os.getenv("AI_TIMEOUT", "120")) * 1000
client = genai.Client(
    api_key=os.getenv("GEMINI_API_KEY"),
    http_options=types.HttpOptions(
        timeout=AI_TIMEOUT_MS,
        client_args={"transport": httpx.HTTPTransport(local_address="0.0.0.0")},
    ),
)
# AI_MODEL bo'sh yoki yo'q bo'lsa ham standart nom ishlaydi.
MODEL = os.getenv("AI_MODEL") or "gemini-flash-latest"
# Asosiy model band bo'lsa, shu yengilroq Gemini modeli sinab ko'riladi.
FALLBACK_MODEL = os.getenv("AI_FALLBACK_MODEL") or "gemini-flash-lite-latest"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "bot.db")
STYLES_PREVIEW = os.path.join(BASE_DIR, "styles.png")  # uslublar rasmi (bo'lmasa, o'tkazib yuboriladi)

SLIDE_CHOICES = (5, 10, 15, 20)  # jami slayd soni (sarlavha + reja + xulosa ham ichida)
HISTORY_LIMIT = 10

# ============================================================
# MENYU (panel) tugmalari
# ============================================================
BTN_NEW = "🎞 Yangi slayd"
BTN_HISTORY = "📚 Tarixim"
BTN_ASK = "❓ Savol berish"
BTN_PROFILE = "👤 Ma'lumotlarim"
MENU = ReplyKeyboardMarkup(
    [[BTN_NEW, BTN_HISTORY], [BTN_ASK, BTN_PROFILE]],
    resize_keyboard=True,
)

CHAT_SYSTEM = (
    "Sen o'quvchi va talabalarga yordam beradigan AI yordamchisan. "
    "Doim o'zbek tilida (lotin yozuvida) sodda va aniq javob ber. "
    "Inglizcha atamalarni qavs ichida o'zbekcha tarjimasi bilan yoz. "
    "Javobing qisqa va tushunarli bo'lsin."
)

SLIDE_SYSTEM = (
    "Sen taqdimot (slayd) matnlarini tayyorlaydigan yordamchisan. "
    "Faqat bitta JSON obyekt qaytar, boshqa hech narsa yozma. "
    "Til: o'zbek (lotin yozuvi). Faktlarni to'qima, aniq bilmasang umumiy yoz."
)

SLIDE_PROMPT = """Mavzu: {topic}
Kontent slaydlari soni: {n} ta (sarlavha, reja va xulosa slaydlari bundan tashqari).

Quyidagi JSON tuzilmasini qaytar:
{{
  "title": "taqdimot sarlavhasi (qisqa)",
  "subtitle": "bir jumlali tavsif",
  "label": "KATTA HARFLARDA fan/yo'nalish nomi, masalan BIOLOGIYA",
  "plan": ["reja bandi 1", "reja bandi 2", "... (4-7 ta)"],
  "slides": [
    {{"title": "slayd sarlavhasi", "bullets": ["fikr 1", "fikr 2", "fikr 3"]}}
  ],
  "conclusion": ["xulosa 1", "xulosa 2", "xulosa 3"]
}}

Qoidalar:
- "slides" ichida aynan {n} ta slayd bo'lsin.
- Har slaydda 3-5 ta band, har band 15 so'zdan oshmasin.
- Sarlavhalar 8 so'zdan oshmasin.
- Reja bandlari slayd sarlavhalariga mos bo'lsin."""


# ============================================================
# MA'LUMOTLAR BAZASI (SQLite): foydalanuvchilar va slaydlar tarixi
# ============================================================
def db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with closing(db()) as conn, conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS users (
                user_id    INTEGER PRIMARY KEY,
                full_name  TEXT NOT NULL,
                group_name TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS history (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id     INTEGER NOT NULL,
                topic       TEXT,
                title       TEXT,
                slide_count INTEGER,
                file_id     TEXT,
                filename    TEXT,
                created_at  TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_history_user ON history(user_id, id DESC);
            """
        )


def get_user(user_id: int):
    with closing(db()) as conn:
        return conn.execute("SELECT * FROM users WHERE user_id = ?", (user_id,)).fetchone()


def save_user(user_id: int, full_name: str, group_name: str):
    with closing(db()) as conn, conn:
        conn.execute(
            "INSERT INTO users (user_id, full_name, group_name) VALUES (?, ?, ?) "
            "ON CONFLICT(user_id) DO UPDATE SET full_name = excluded.full_name, "
            "group_name = excluded.group_name",
            (user_id, full_name, group_name),
        )


def add_history(user_id, topic, title, slide_count, file_id, filename):
    with closing(db()) as conn, conn:
        conn.execute(
            "INSERT INTO history (user_id, topic, title, slide_count, file_id, filename, created_at) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            (user_id, topic, title, slide_count, file_id, filename,
             datetime.now().strftime("%Y-%m-%d %H:%M")),
        )


def get_history(user_id: int, limit: int = HISTORY_LIMIT):
    with closing(db()) as conn:
        return conn.execute(
            "SELECT * FROM history WHERE user_id = ? ORDER BY id DESC LIMIT ?", (user_id, limit)
        ).fetchall()


def get_history_item(user_id: int, item_id: int):
    with closing(db()) as conn:
        return conn.execute(
            "SELECT * FROM history WHERE id = ? AND user_id = ?", (item_id, user_id)
        ).fetchone()


# ============================================================
# AI
# ============================================================
def _generate_gemini(prompt: str, system: str, max_tokens: int, as_json: bool = False) -> str:
    use_thinking_off = True  # "o'ylash" rejimini o'chirsak, javob ancha tez keladi
    last_error = None
    # 2 marta asosiy model, keyin 1 marta zaxira Gemini modeli
    models = [MODEL, MODEL, FALLBACK_MODEL]
    for attempt, model in enumerate(models):
        extra = {"thinking_config": types.ThinkingConfig(thinking_budget=0)} if use_thinking_off else {}
        config = types.GenerateContentConfig(
            system_instruction=system,
            max_output_tokens=max_tokens,
            response_mime_type="application/json" if as_json else None,
            **extra,
        )
        try:
            resp = client.models.generate_content(model=model, contents=prompt, config=config)
            return (resp.text or "").strip()
        except Exception as e:
            last_error = e
            msg = str(e).lower()
            if use_thinking_off and "thinking" in msg:
                use_thinking_off = False  # bu model bu sozlamani qabul qilmadi, usiz urinamiz
                continue
            if any(k in msg for k in ("429", "503", "504", "unavailable", "deadline", "timed out", "timeout")):
                log.warning("Gemini band, %s modeli, urinish %d/%d", model, attempt + 1, len(models))
                time.sleep(3 * (attempt + 1))
                continue
            raise
    raise last_error


# Zaxira xizmatlar (OpenAI uslubidagi API). Kaliti .env da bo'lmasa, o'tkazib yuboriladi.
OPENAI_PROVIDERS = [
    ("groq", "https://api.groq.com/openai/v1/chat/completions",
     os.getenv("GROQ_API_KEY"), os.getenv("GROQ_MODEL") or "llama-3.3-70b-versatile"),
    ("openrouter", "https://openrouter.ai/api/v1/chat/completions",
     os.getenv("OPENROUTER_API_KEY"),
     os.getenv("OPENROUTER_MODEL") or "meta-llama/llama-3.3-70b-instruct:free"),
]


def _openai_style(url, key, model, prompt, system, max_tokens, as_json):
    body = {
        "model": model,
        "messages": [{"role": "system", "content": system}, {"role": "user", "content": prompt}],
        "max_tokens": min(max_tokens, 7000),
    }
    if as_json:
        body["response_format"] = {"type": "json_object"}
    with httpx.Client(transport=httpx.HTTPTransport(local_address="0.0.0.0"),
                      timeout=AI_TIMEOUT_MS / 1000) as http:
        r = http.post(url, json=body, headers={"Authorization": f"Bearer {key}"})
        r.raise_for_status()
        return (r.json()["choices"][0]["message"]["content"] or "").strip()


def _generate(prompt: str, system: str, max_tokens: int, as_json: bool = False) -> str:
    """Avval Gemini, ishlamasa zaxira xizmatlar (Groq, OpenRouter) ketma-ket sinab ko'riladi."""
    errors = []
    try:
        return _generate_gemini(prompt, system, max_tokens, as_json)
    except Exception as e:
        log.warning("Gemini ishlamadi: %s", str(e)[:150])
        errors.append(f"gemini: {str(e)[:80]}")
    for name, url, key, model in OPENAI_PROVIDERS:
        if not key:
            continue
        try:
            log.info("Zaxira xizmatga o'tildi: %s", name)
            return _openai_style(url, key, model, prompt, system, max_tokens, as_json)
        except Exception as e:
            log.warning("%s ishlamadi: %s", name, str(e)[:150])
            errors.append(f"{name}: {str(e)[:80]}")
    raise RuntimeError("Hamma AI xizmatlari band yoki ishlamayapti. " + " | ".join(errors))


def ask(text: str) -> str:
    return _generate(text, CHAT_SYSTEM, 2000)


def make_slide_content(topic: str, n: int) -> dict:
    # Slaydlar ko'p bo'lsa, javob ham uzun bo'ladi, shuning uchun limit n ga qarab oshadi
    raw = _generate(SLIDE_PROMPT.format(topic=topic, n=n), SLIDE_SYSTEM, 8000 + 500 * n, as_json=True)
    match = re.search(r"\{.*\}", raw, re.S)
    if not match:
        raise ValueError("AI JSON qaytarmadi")
    data = json.loads(match.group(0))
    for key in ("title", "plan", "slides", "conclusion"):
        if key not in data:
            raise ValueError(f"JSON da '{key}' yo'q")
    return data


# ============================================================
# PPTX YASASH (5 xil uslub)
# ============================================================
DEFAULT_STYLE = "modern"
STYLE_ORDER = ("classic", "modern", "dark", "nature", "minimal")

# mode: sarlavha joylashuvi ("bar" - tepada keng chiziq, "plain" - kartochka,
#       "side" - chapda yon panel, "line" - oddiy sarlavha + chiziq)
THEMES = {
    "classic": {
        "label": "📘 Klassik", "head": "Georgia", "body": "Arial", "mode": "plain", "deco": "bar",
        "title_bg": "141B2D", "title_fg": "FFFFFF", "title_sub": "CFE8E2", "title_accent": "F2B84B",
        "body_bg": "F6F3EC", "text": "1C2333", "head_fg": "1C2333",
        "header_bg": "0B6558", "header_fg": "FFFFFF", "accent": "0B6558",
        "card": "CFE8E2", "muted": "6B7280", "end_bg": "141B2D", "end_fg": "FFFFFF",
    },
    "modern": {
        "label": "💠 Zamonaviy", "head": "Calibri", "body": "Calibri", "mode": "bar", "deco": "bar",
        "title_bg": "141B2D", "title_fg": "FFFFFF", "title_sub": "A9B4C8", "title_accent": "3B82F6",
        "body_bg": "F5F7FB", "text": "1F2937", "head_fg": "FFFFFF",
        "header_bg": "141B2D", "header_fg": "FFFFFF", "accent": "3B82F6",
        "card": None, "muted": "6B7280", "end_bg": "141B2D", "end_fg": "FFFFFF",
    },
    "dark": {
        "label": "🌙 Qorong'i", "head": "Calibri", "body": "Calibri", "mode": "plain", "deco": "bar",
        "title_bg": "0B1020", "title_fg": "FFFFFF", "title_sub": "94A3B8", "title_accent": "F59E0B",
        "body_bg": "0F172A", "text": "E5E7EB", "head_fg": "FFFFFF",
        "header_bg": "F59E0B", "header_fg": "0F172A", "accent": "F59E0B",
        "card": "1E293B", "muted": "64748B", "end_bg": "0B1020", "end_fg": "FFFFFF",
    },
    "nature": {
        "label": "🌿 Tabiat", "head": "Cambria", "body": "Calibri", "mode": "side", "deco": "bar",
        "title_bg": "1B4332", "title_fg": "FFFFFF", "title_sub": "B7E4C7", "title_accent": "95D5B2",
        "body_bg": "F4FAF5", "text": "1B2E22", "head_fg": "1B4332",
        "header_bg": "2D6A4F", "header_fg": "FFFFFF", "accent": "40916C",
        "card": None, "muted": "5B7A66", "end_bg": "1B4332", "end_fg": "FFFFFF",
    },
    "minimal": {
        "label": "⬜ Minimal", "head": "Arial", "body": "Arial", "mode": "line", "deco": "line",
        "title_bg": "FFFFFF", "title_fg": "111111", "title_sub": "6B7280", "title_accent": "E11D48",
        "body_bg": "FFFFFF", "text": "111111", "head_fg": "111111",
        "header_bg": "111111", "header_fg": "FFFFFF", "accent": "E11D48",
        "card": None, "muted": "9CA3AF", "end_bg": "111111", "end_fg": "FFFFFF",
    },
}


def _c(hexstr: str) -> RGBColor:
    return RGBColor.from_string(hexstr)


def _bg(slide, hexstr):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = _c(hexstr)


def _rect(slide, x, y, w, h, hexstr, rounded=False):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(kind, x, y, w, h)
    if rounded:
        shape.adjustments[0] = 0.04
    shape.fill.solid()
    shape.fill.fore_color.rgb = _c(hexstr)
    shape.line.fill.background()
    return shape


def _text(slide, x, y, w, h, text, size, color, font, bold=False, align=PP_ALIGN.LEFT,
          anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = str(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = _c(color)
    run.font.name = font
    return box


def _bullets(slide, area, items, size, color, font, anchor=MSO_ANCHOR.TOP):
    x, y, w, h = area
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(12)
        run = p.add_run()
        run.text = f"•  {item}"
        run.font.size = Pt(size)
        run.font.color.rgb = _c(color)
        run.font.name = font
    return box


def _as_list(value) -> list:
    if isinstance(value, list):
        return [str(v) for v in value]
    return [str(value)]


def _title_slide(prs, th, data, author):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s, th["title_bg"])
    if th["deco"] == "line":
        _rect(s, Inches(1.2), Inches(1.0), Inches(1.6), Inches(0.1), th["title_accent"])
    else:
        _rect(s, Inches(0.8), Inches(2.2), Inches(0.12), Inches(2.6), th["title_accent"])
    label = str(data.get("label", "")).upper()
    if label:
        _text(s, Inches(1.2), Inches(1.5), Inches(10), Inches(0.5), label, 16,
              th["title_accent"], th["body"], bold=True)
    _text(s, Inches(1.2), Inches(2.2), Inches(11), Inches(1.8), data["title"], 44,
          th["title_fg"], th["head"], bold=True, anchor=MSO_ANCHOR.MIDDLE)
    if data.get("subtitle"):
        _text(s, Inches(1.2), Inches(4.1), Inches(11), Inches(1), data["subtitle"], 20,
              th["title_sub"], th["body"])
    if author:
        _text(s, Inches(1.2), Inches(6.2), Inches(11), Inches(0.6), author, 18,
              th["title_sub"], th["body"])


def _content_slide(prs, th, title, items, num, topic, total):
    """Reja va kontent slaydlari uchun umumiy shakl. Uslubga qarab joylashuv o'zgaradi."""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s, th["body_bg"])
    W, H = prs.slide_width, prs.slide_height
    mode = th["mode"]
    anchor = MSO_ANCHOR.TOP
    size = 22
    footer_x = Inches(0.7)

    if mode == "bar":
        _rect(s, 0, 0, W, Inches(1.3), th["header_bg"])
        _rect(s, 0, Inches(1.3), W, Inches(0.06), th["accent"])
        _text(s, Inches(0.8), Inches(0.2), Inches(11.2), Inches(0.9), title, 30,
              th["head_fg"], th["head"], bold=True, anchor=MSO_ANCHOR.MIDDLE)
        _text(s, Inches(12.0), Inches(0.2), Inches(1), Inches(0.9), str(num), 20,
              th["accent"], th["head"], bold=True, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        area = (Inches(1), Inches(1.9), Inches(11.3), Inches(4.9))
    elif mode == "plain":
        _rect(s, Inches(0.7), Inches(0.62), Inches(0.12), Inches(0.8), th["header_bg"])
        _text(s, Inches(1.0), Inches(0.5), Inches(11.6), Inches(1.1), title, 32,
              th["head_fg"], th["head"], bold=True, anchor=MSO_ANCHOR.MIDDLE)
        _rect(s, Inches(0.7), Inches(1.9), Inches(11.93), Inches(4.7), th["card"], rounded=True)
        area = (Inches(1.1), Inches(2.0), Inches(11.1), Inches(4.5))
        anchor = MSO_ANCHOR.MIDDLE
    elif mode == "side":
        _rect(s, 0, 0, Inches(2.4), H, th["header_bg"])
        _text(s, 0, Inches(0.5), Inches(2.4), Inches(1.4), str(num), 54,
              th["header_fg"], th["head"], bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        _text(s, Inches(3.0), Inches(0.5), Inches(9.8), Inches(1.2), title, 30,
              th["head_fg"], th["head"], bold=True, anchor=MSO_ANCHOR.MIDDLE)
        _rect(s, Inches(3.0), Inches(1.8), Inches(1.2), Inches(0.06), th["accent"])
        area = (Inches(3.0), Inches(2.1), Inches(9.8), Inches(4.7))
        size = 20
        footer_x = Inches(3.0)
    else:  # "line"
        _text(s, Inches(0.9), Inches(0.45), Inches(10.8), Inches(1.0), title, 30,
              th["head_fg"], th["head"], bold=True, anchor=MSO_ANCHOR.MIDDLE)
        _text(s, Inches(11.9), Inches(0.45), Inches(1), Inches(1.0), str(num), 20,
              th["accent"], th["head"], bold=True, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        _rect(s, Inches(0.9), Inches(1.55), Inches(1.5), Inches(0.06), th["accent"])
        area = (Inches(0.9), Inches(1.95), Inches(11.5), Inches(4.9))

    _bullets(s, area, items, size, th["text"], th["body"], anchor)
    _text(s, footer_x, Inches(6.95), Inches(9), Inches(0.35), f"{short(topic, 40)} · {num} / {total}",
          11, th["muted"], th["body"])
    return s


def _end_slide(prs, th, items):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s, th["end_bg"])
    _text(s, Inches(1), Inches(0.8), Inches(11), Inches(1), "Xulosa", 38, th["end_fg"], th["head"],
          bold=True)
    _rect(s, Inches(1), Inches(1.8), Inches(1.2), Inches(0.08), th["title_accent"])
    _bullets(s, (Inches(1), Inches(2.2), Inches(11), Inches(3.8)), items, 22, th["end_fg"], th["body"])
    _text(s, Inches(1), Inches(6.3), Inches(11.3), Inches(0.8), "E'tiboringiz uchun rahmat!", 24,
          th["title_accent"], th["head"], bold=True)


def build_pptx(data: dict, author: str = "", style: str = DEFAULT_STYLE) -> io.BytesIO:
    th = THEMES.get(style, THEMES[DEFAULT_STYLE])
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    topic = data["title"]
    slides = data["slides"]
    total = len(slides) + 3

    _title_slide(prs, th, data, author)
    _content_slide(prs, th, "Reja", _as_list(data["plan"])[:7], 2, topic, total)
    for idx, sl in enumerate(slides, 1):
        _content_slide(prs, th, sl.get("title", ""), _as_list(sl.get("bullets", []))[:5],
                       idx + 2, topic, total)
    _end_slide(prs, th, _as_list(data["conclusion"])[:4])

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


def safe_filename(title: str) -> str:
    name = re.sub(r"[^\w\- ]", "", title, flags=re.U).strip().replace(" ", "_")
    return (name[:50] or "taqdimot") + ".pptx"


# ============================================================
# TELEGRAM: yordamchi funksiyalar
# ============================================================
def short(text: str, limit: int = 38) -> str:
    text = str(text or "")
    return text if len(text) <= limit else text[: limit - 1] + "…"


def fmt_date(created_at: str) -> str:
    try:
        return datetime.strptime(created_at, "%Y-%m-%d %H:%M").strftime("%d.%m")
    except (TypeError, ValueError):
        return ""


async def ask_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["step"] = "ask_name"
    await update.effective_message.reply_text(
        "👤 Tanishib olaylik!\n\nIsm va familiyangizni yozing.\nMasalan: Aliyev Vali"
    )


# ============================================================
# TELEGRAM: buyruqlar
# ============================================================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    intro = (
        "👋 Salom! Men mavzu bo'yicha tayyor taqdimot (slayd) yasab beraman.\n\n"
        "Qanday ishlaydi:\n"
        "1️⃣ «🎞 Yangi slayd» tugmasini bosing\n"
        "2️⃣ Mavzuni yozing\n"
        "3️⃣ Slayd sonini tanlang (5, 10, 15 yoki 20)\n"
        "4️⃣ Tayyor PowerPoint (.pptx) faylni oling"
    )
    if get_user(update.effective_user.id):
        await update.message.reply_text(intro + "\n\nMenyudan tanlang 👇", reply_markup=MENU)
    else:
        await update.message.reply_text(intro)
        await ask_name(update, context)


async def menu_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["step"] = None
    await update.message.reply_text("Menyudan tanlang 👇", reply_markup=MENU)


async def cancel_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("Bekor qilindi.", reply_markup=MENU)


async def question_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = " ".join(context.args).strip()
    if not q:
        return await start_question_mode(update, context)
    await answer_question(update, q)


async def history_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await show_history(update, context)


# ============================================================
# TELEGRAM: menyu bo'limlari
# ============================================================
async def new_slide(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not get_user(update.effective_user.id):
        await update.message.reply_text("Avval tanishib olaylik 👇")
        return await ask_name(update, context)
    context.user_data["step"] = "ask_topic"
    await update.message.reply_text(
        "📝 Slayd mavzusini yozing.\nMasalan: Fotosintez jarayoni",
        reply_markup=MENU,
    )


async def got_topic(update: Update, context: ContextTypes.DEFAULT_TYPE, text: str):
    if len(text) < 3:
        return await update.message.reply_text("Mavzu juda qisqa. Iltimos, to'liqroq yozing.")
    context.user_data["topic"] = text[:200]
    context.user_data["step"] = None

    # Uslublar rasmi bo'lsa, avval shuni ko'rsatamiz
    if os.path.exists(STYLES_PREVIEW):
        try:
            with open(STYLES_PREVIEW, "rb") as f:
                await update.message.reply_photo(f, caption="🎨 Slayd uslublari (1-5)")
        except Exception:
            log.exception("Uslublar rasmini yuborib bo'lmadi")

    btn = lambda k: InlineKeyboardButton(THEMES[k]["label"], callback_data=f"s:{k}")
    kb = InlineKeyboardMarkup(
        [
            [btn("classic"), btn("modern")],
            [btn("dark"), btn("nature")],
            [btn("minimal")],
            [InlineKeyboardButton("❌ Bekor qilish", callback_data="cancel")],
        ]
    )
    await update.message.reply_text(
        f"Mavzu: «{context.user_data['topic']}»\n\n🎨 Slayd uslubini tanlang 👇", reply_markup=kb
    )


def count_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        [
            [InlineKeyboardButton(f"{n} ta", callback_data=f"n:{n}") for n in SLIDE_CHOICES],
            [InlineKeyboardButton("❌ Bekor qilish", callback_data="cancel")],
        ]
    )


async def show_history(update: Update, context: ContextTypes.DEFAULT_TYPE):
    rows = get_history(update.effective_user.id)
    if not rows:
        return await update.effective_message.reply_text(
            "📚 Hali slayd yaratmagansiz.\n«🎞 Yangi slayd» tugmasini bosing.", reply_markup=MENU
        )
    buttons = [
        [InlineKeyboardButton(
            f"{fmt_date(r['created_at'])} · {short(r['title'])} ({r['slide_count']} ta)",
            callback_data=f"h:{r['id']}",
        )]
        for r in rows
    ]
    await update.effective_message.reply_text(
        f"📚 Oxirgi {len(rows)} ta slaydingiz.\nFaylni qayta olish uchun ustiga bosing 👇",
        reply_markup=InlineKeyboardMarkup(buttons),
    )


async def show_profile(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = get_user(update.effective_user.id)
    if not user:
        return await ask_name(update, context)
    kb = InlineKeyboardMarkup([[InlineKeyboardButton("✏️ O'zgartirish", callback_data="edit_profile")]])
    await update.message.reply_text(
        f"👤 Ism: {user['full_name']}\n👥 Guruh: {user['group_name']}\n\n"
        "Bu ma'lumotlar slaydning birinchi sahifasiga yoziladi.",
        reply_markup=kb,
    )


async def start_question_mode(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["step"] = "ask_question"
    await update.message.reply_text(
        "❓ Savolingizni yozing, AI javob beradi.\n"
        "Boshqa bo'limga o'tish uchun menyudagi tugmalardan birini bosing."
    )


async def answer_question(update: Update, q: str):
    await update.message.chat.send_action("typing")
    try:
        answer = await asyncio.to_thread(ask, q)
    except Exception as e:
        log.exception("AI savol xatosi")
        return await update.message.reply_text(f"❌ Xato: {str(e)[:300]}")
    await update.message.reply_text(answer[:4000] or "Javob topilmadi.")


# ============================================================
# TELEGRAM: matn va tugmalar
# ============================================================
async def on_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    step = context.user_data.get("step")

    # Menyu tugmalari har doim ishlaydi
    if text == BTN_NEW:
        return await new_slide(update, context)
    if text == BTN_HISTORY:
        context.user_data["step"] = None
        return await show_history(update, context)
    if text == BTN_ASK:
        return await start_question_mode(update, context)
    if text == BTN_PROFILE:
        context.user_data["step"] = None
        return await show_profile(update, context)

    if step == "ask_name":
        if not 2 <= len(text) <= 60:
            return await update.message.reply_text("Ism va familiyani to'g'ri yozing (2-60 harf).")
        context.user_data["tmp_name"] = text
        context.user_data["step"] = "ask_group"
        return await update.message.reply_text(
            "👥 Endi guruhingizni (yoki sinfingizni) yozing.\nMasalan: 101-guruh"
        )

    if step == "ask_group":
        if not 1 <= len(text) <= 30:
            return await update.message.reply_text("Guruh nomi juda uzun yoki bo'sh. Qisqaroq yozing.")
        name = context.user_data.pop("tmp_name", "")
        save_user(update.effective_user.id, name, text)
        context.user_data["step"] = None
        return await update.message.reply_text(
            f"✅ Rahmat, {name}!\nEndi «🎞 Yangi slayd» tugmasini bosib, birinchi slaydingizni yarating.",
            reply_markup=MENU,
        )

    if step == "ask_topic":
        return await got_topic(update, context, text)

    if step == "ask_question":
        return await answer_question(update, text)

    await update.message.reply_text(
        "Slayd yaratish uchun «🎞 Yangi slayd» tugmasini bosing 👇", reply_markup=MENU
    )


async def on_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    data = q.data or ""

    if data == "cancel":
        context.user_data.pop("topic", None)
        context.user_data.pop("style", None)
        return await q.edit_message_text("Bekor qilindi.")
    if data == "edit_profile":
        return await ask_name(update, context)
    if data.startswith("s:") and data[2:] in THEMES:
        context.user_data["style"] = data[2:]
        topic = context.user_data.get("topic")
        if not topic:
            return await q.edit_message_text("Mavzu topilmadi. «🎞 Yangi slayd» tugmasini bosing.")
        return await q.edit_message_text(
            f"Mavzu: «{topic}»\nUslub: {THEMES[data[2:]]['label']}\n\nNechta slayd kerak? 👇",
            reply_markup=count_keyboard(),
        )
    if data.startswith("n:"):
        total = int(data[2:]) if data[2:].isdigit() else 0
        if total in SLIDE_CHOICES:
            return await generate(update, context, total)
    if data.startswith("h:") and data[2:].isdigit():
        return await send_history_item(update, context, int(data[2:]))


# ============================================================
# KUTISH ANIMATSIYASI (spinner + progress chizig'i + soniyalar)
# ============================================================
SPINNER = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
EXPECTED_SECONDS = 45  # progress chizig'i shu vaqtga mo'ljallangan (taxminiy)


async def progress_updater(message, header: str):
    """Kutish vaqtida xabarni har 2 soniyada yangilab turadi."""
    start = time.monotonic()
    i = 0
    while True:
        await asyncio.sleep(2)
        i += 1
        sec = int(time.monotonic() - start)
        filled = min(9, int(sec / EXPECTED_SECONDS * 10))  # 100% ga yetmaydi, to'xtab qolmasligi uchun
        bar = "▰" * filled + "▱" * (10 - filled)
        text = f"{SPINNER[i % len(SPINNER)]} {header}\n\n{bar}\n⏱ {sec} soniya o'tdi..."
        try:
            await message.edit_text(text)
        except Exception:
            pass  # "message is not modified" yoki limit xatolari e'tiborsiz


async def generate(update: Update, context: ContextTypes.DEFAULT_TYPE, total: int):
    q = update.callback_query
    uid = update.effective_user.id
    topic = context.user_data.get("topic")
    user = get_user(uid)

    if not topic or not user:
        return await q.edit_message_text("Mavzu topilmadi. «🎞 Yangi slayd» tugmasini bosing.")
    if context.user_data.get("busy"):
        return  # ikki marta bosishdan himoya
    context.user_data["busy"] = True

    style = context.user_data.get("style", DEFAULT_STYLE)
    style_label = THEMES.get(style, THEMES[DEFAULT_STYLE])["label"]
    header = f"«{short(topic, 60)}»\n{style_label} · {total} ta slayd tayyorlanmoqda"
    await q.edit_message_text(f"{SPINNER[0]} {header}\n\n▱▱▱▱▱▱▱▱▱▱\n⏱ 0 soniya o'tdi...")
    progress_task = asyncio.create_task(progress_updater(q.message, header))

    async def stop_progress():
        progress_task.cancel()
        try:
            await progress_task
        except asyncio.CancelledError:
            pass

    try:
        # AI so'rovi va fayl yasash bloklovchi, shuning uchun alohida oqimda
        data = await asyncio.to_thread(make_slide_content, topic, total - 3)
        author = f"{user['full_name']} · {user['group_name']}"
        buf = await asyncio.to_thread(build_pptx, data, author, style)
        await stop_progress()

        filename = safe_filename(data["title"])
        count = len(data["slides"]) + 3
        msg = await q.message.reply_document(
            document=buf,
            filename=filename,
            caption=f"✅ Tayyor: {data['title']}\n📄 {count} ta slayd · {style_label}",
        )
        add_history(uid, topic, data["title"], count, msg.document.file_id, filename)
        context.user_data.pop("topic", None)
        context.user_data.pop("style", None)
        try:
            await q.message.delete()
        except Exception:
            pass
    except Exception as e:
        await stop_progress()
        log.exception("Slayd yaratishda xato")
        reason = str(e) or type(e).__name__
        if "timeout" in type(e).__name__.lower() or "timed out" in reason.lower():
            reason = "AI vaqtida javob bermadi (internet yoki Google serveri sekin)."
        if any(k in reason.lower() for k in ("503", "504", "unavailable", "deadline", "band")):
            reason = "AI xizmatlari hozir band yoki sekin. 1-2 daqiqadan keyin qayta urinib ko'ring (kamroq slayd tanlasangiz tezroq bo'ladi)."
        await q.edit_message_text(f"❌ Xato: {reason[:300]}\n\nQayta urinib ko'ring: «🎞 Yangi slayd».")
    finally:
        context.user_data["busy"] = False


async def send_history_item(update: Update, context: ContextTypes.DEFAULT_TYPE, item_id: int):
    q = update.callback_query
    row = get_history_item(update.effective_user.id, item_id)
    if not row:
        return await q.message.reply_text("Bu yozuv topilmadi.")
    try:
        await q.message.reply_document(
            document=row["file_id"],
            caption=f"📄 {row['title']} ({row['slide_count']} ta slayd)",
        )
    except Exception:
        log.exception("Tarixdagi faylni yuborib bo'lmadi")
        await q.message.reply_text("❌ Faylni qayta yuborib bo'lmadi. Yangi slayd yasab ko'ring.")


async def on_error(update, context: ContextTypes.DEFAULT_TYPE):
    log.error("Kutilmagan xato", exc_info=context.error)


# ============================================================
# ISHGA TUSHIRISH
# ============================================================
class _Health(BaseHTTPRequestHandler):
    """Render bepul tarifi veb-servis (port) talab qiladi. Bu oddiy server UptimeRobot uchun ham javob beradi."""

    def _ok(self, with_body: bool):
        self.send_response(200)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.end_headers()
        if with_body:
            self.wfile.write(b"Bot ishlayapti")

    def do_GET(self):
        self._ok(True)

    def do_HEAD(self):
        self._ok(False)

    def log_message(self, *args):  # logni to'ldirmaslik uchun
        pass


def start_health_server():
    port = os.getenv("PORT")  # Render PORT ni o'zi beradi; kompyuterda bo'lmaydi va kerak ham emas
    if not port:
        return
    server = HTTPServer(("0.0.0.0", int(port)), _Health)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    log.info("Health-server %s portda ishga tushdi", port)


def main():
    token = os.getenv("BOT_TOKEN")
    if not token:
        raise RuntimeError("BOT_TOKEN topilmadi. .env faylga BOT_TOKEN=... yozing")

    init_db()
    start_health_server()

    def make_request():
        return HTTPXRequest(
            connect_timeout=30,
            read_timeout=30,
            write_timeout=30,
            pool_timeout=30,
            httpx_kwargs={
                "transport": httpx.AsyncHTTPTransport(local_address="0.0.0.0")
            },
        )

    app = (
        ApplicationBuilder()
        .token(token)
        .request(make_request())
        .get_updates_request(make_request())
        .build()
    )
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("menu", menu_cmd))
    app.add_handler(CommandHandler("bekor", cancel_cmd))
    app.add_handler(CommandHandler("savol", question_cmd))
    app.add_handler(CommandHandler("tarix", history_cmd))
    app.add_handler(CallbackQueryHandler(on_callback))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text))
    app.add_error_handler(on_error)
    print("Bot ishga tushdi ...")
    app.run_polling()

  
if __name__ == "__main__":
    main()