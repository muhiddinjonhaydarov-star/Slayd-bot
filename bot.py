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

import anthropic
import httpx
from dotenv import load_dotenv
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, LabeledPrice, ReplyKeyboardMarkup, Update
from telegram.ext import (
    ApplicationBuilder,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    PreCheckoutQueryHandler,
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
client = None if not os.getenv("GEMINI_API_KEY") else genai.Client(
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
# CLAUDE + TO'LOV SOZLAMALARI (.env dan o'qiladi)
# ============================================================
CLAUDE_MODEL = os.getenv("CLAUDE_MODEL") or "claude-haiku-5-5"
claude_client = (
    anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"), timeout=AI_TIMEOUT_MS / 1000)
    if os.getenv("ANTHROPIC_API_KEY") else None
)
# TARIFLAR: to'lov summasi (so'm) -> beriladigan SLAYD ballari.
# Taqdimot narxi = uning jami slayd soni (5, 10, 15 yoki 20 ball). Tarifni o'zgartirmoqchi bo'lsangiz, shu yerni tahrirlang.
PLANS = {
    5000: 20,    # 1 ta taqdimot (20 talik yoki 15 talik)
    10000: 45,   # 3 ta 15 talik  yoki 2 ta 20 talik
    15000: 60,   # 4 ta 15 talik  yoki 3 ta 20 talik
    20000: 80,   # 5 ta 15 talik  yoki 4 ta 20 talik
    25000: 100,  # 6 ta 15 talik  yoki 5 ta 20 talik
    30000: 120,  # 8 ta 15 talik  yoki 6 ta 20 talik
    35000: 140,  # 9 ta 15 talik  yoki 7 ta 20 talik
    40000: 160,  # 10 ta 15 talik yoki 8 ta 20 talik
    45000: 180,  # 12 ta 15 talik yoki 9 ta 20 talik
    50000: 200,  # 13 ta 15 talik yoki 10 ta 20 talik
}
PROVIDER_TOKEN = os.getenv("PROVIDER_TOKEN") or ""   # BotFather -> Payments (Click/Payme)
CARD_NUMBER = os.getenv("CARD_NUMBER") or ""          # qo'lda to'lov uchun karta raqami
CARD_HOLDER = os.getenv("CARD_HOLDER") or ""         # karta egasi (ixtiyoriy)
ADMIN_ID = int(os.getenv("ADMIN_ID") or "0")         # sizning Telegram ID raqamingiz
# Doim tekin foydalanadigan profillar (vergul bilan): FREE_IDS=123456789,987654321
FREE_IDS_ENV = {int(x) for x in re.findall(r"\d+", os.getenv("FREE_IDS") or "")}

# ============================================================
# MENYU (panel) tugmalari
# ============================================================
BTN_NEW = "🎞 Yangi slayd"
BTN_HISTORY = "📚 Tarixim"
BTN_ASK = "❓ Savol berish"
BTN_PROFILE = "👤 Ma'lumotlarim"
BTN_BALANCE = "💰 Balans"
MENU = ReplyKeyboardMarkup(
    [[BTN_NEW, BTN_HISTORY], [BTN_ASK, BTN_PROFILE], [BTN_BALANCE]],
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
            CREATE TABLE IF NOT EXISTS wallets (
                user_id INTEGER PRIMARY KEY,
                balance INTEGER NOT NULL DEFAULT 0
            );
            CREATE TABLE IF NOT EXISTS payments (
                charge_id TEXT PRIMARY KEY,
                user_id   INTEGER,
                amount    INTEGER,
                created   TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS free_users (user_id INTEGER PRIMARY KEY);
            CREATE TABLE IF NOT EXISTS card_orders (
                id      INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                price   INTEGER NOT NULL,
                credits INTEGER NOT NULL,
                status  TEXT NOT NULL DEFAULT 'pending',
                created TEXT DEFAULT CURRENT_TIMESTAMP
            );
            """
        )


# ---- hamyon (balans) va tekin profillar ----
def get_balance(user_id: int) -> int:
    with closing(db()) as conn, conn:
        conn.execute("INSERT OR IGNORE INTO wallets (user_id) VALUES (?)", (user_id,))
        return conn.execute("SELECT balance FROM wallets WHERE user_id = ?", (user_id,)).fetchone()[0]


def add_balance(user_id: int, amount: int):
    get_balance(user_id)
    with closing(db()) as conn, conn:
        conn.execute("UPDATE wallets SET balance = balance + ? WHERE user_id = ?", (amount, user_id))


def try_charge(user_id: int, amount: int) -> bool:
    """Balansdan pulni atomik yechadi. Yetmasa False qaytaradi."""
    get_balance(user_id)
    with closing(db()) as conn, conn:
        cur = conn.execute(
            "UPDATE wallets SET balance = balance - ? WHERE user_id = ? AND balance >= ?",
            (amount, user_id, amount),
        )
        return cur.rowcount == 1


def record_payment(charge_id: str, user_id: int, amount: int) -> bool:
    try:
        with closing(db()) as conn, conn:
            conn.execute("INSERT INTO payments (charge_id, user_id, amount) VALUES (?, ?, ?)",
                         (charge_id, user_id, amount))
        return True
    except sqlite3.IntegrityError:  # shu to'lov ikki marta hisoblanmasin
        return False


def is_free(user_id: int) -> bool:
    if user_id in FREE_IDS_ENV or (ADMIN_ID and user_id == ADMIN_ID):
        return True
    with closing(db()) as conn:
        return conn.execute("SELECT 1 FROM free_users WHERE user_id = ?", (user_id,)).fetchone() is not None


def set_free(user_id: int, on: bool):
    with closing(db()) as conn, conn:
        if on:
            conn.execute("INSERT OR IGNORE INTO free_users (user_id) VALUES (?)", (user_id,))
        else:
            conn.execute("DELETE FROM free_users WHERE user_id = ?", (user_id,))


# ---- karta orqali qo'lda to'lov buyurtmalari ----
def create_order(user_id: int, price: int, credits: int) -> int:
    with closing(db()) as conn, conn:
        cur = conn.execute("INSERT INTO card_orders (user_id, price, credits) VALUES (?, ?, ?)",
                           (user_id, price, credits))
        return cur.lastrowid


def get_order(order_id: int):
    with closing(db()) as conn:
        return conn.execute("SELECT * FROM card_orders WHERE id = ?", (order_id,)).fetchone()


def set_order_status(order_id: int, old: str, new: str) -> bool:
    """Faqat bir marta o'tadi (ikki marta tasdiqlab, balansni ikki baravar qilib qo'ymaslik uchun)."""
    with closing(db()) as conn, conn:
        cur = conn.execute("UPDATE card_orders SET status = ? WHERE id = ? AND status = ?",
                           (new, order_id, old))
        return cur.rowcount == 1


def count_pending(user_id: int) -> int:
    with closing(db()) as conn:
        return conn.execute("SELECT COUNT(*) FROM card_orders WHERE user_id = ? AND status = 'pending'",
                            (user_id,)).fetchone()[0]


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
    if client is None:
        raise RuntimeError("GEMINI_API_KEY yo'q")
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


def _generate_claude(prompt: str, system: str, max_tokens: int, as_json: bool = False) -> str:
    """Asosiy xizmat: Claude (Anthropic API). 429/5xx/ulanish xatolarida 2 marta urinadi."""
    last_error = None
    for attempt in range(2):
        try:
            resp = claude_client.messages.create(
                model=CLAUDE_MODEL,
                max_tokens=min(max_tokens, 16000),
                system=system,
                messages=[{"role": "user", "content": prompt}],
            )
            return "".join(b.text for b in resp.content if b.type == "text").strip()
        except (anthropic.RateLimitError, anthropic.APIConnectionError, anthropic.InternalServerError) as e:
            last_error = e
            log.warning("Claude band (%s), urinish %d/2", type(e).__name__, attempt + 1)
            time.sleep(3 * (attempt + 1))
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
    if claude_client is not None:
        try:
            return _generate_claude(prompt, system, max_tokens, as_json)
        except Exception as e:
            log.warning("Claude ishlamadi: %s", str(e)[:150])
            errors.append(f"claude: {str(e)[:80]}")
    if client is not None:
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
def money(n: int) -> str:
    return f"{n:,}".replace(",", " ") + " so'm"


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
    if is_free(update.effective_user.id):
        intro += "\n\n🎁 Sizning profilingiz uchun hammasi TEKIN!"
    else:
        intro += "\n\n💵 Tarif 5 000 so'mdan boshlanadi (1 ta taqdimot). Tarifni «💰 Balans» bo'limida tanlaysiz."
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
    uid = update.effective_user.id
    price_note = ""
    if not is_free(uid):
        bal = get_balance(uid)
        if bal < min(SLIDE_CHOICES):
            return await update.message.reply_text(
                f"❗ Balansingiz: {bal} slayd. Taqdimot yaratish uchun tarifni tanlang 👇\n\n" + plans_help(),
                reply_markup=topup_keyboard(),
            )
        price_note = f"\n\n💰 Balansingiz: {bal} slayd. Tanlangan slayd soni balansdan yechiladi."
    context.user_data["step"] = "ask_topic"
    await update.message.reply_text(
        "📝 Slayd mavzusini yozing.\nMasalan: Fotosintez jarayoni" + price_note,
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


def count_keyboard(uid=None) -> InlineKeyboardMarkup:
    bal = None if uid is None or is_free(uid) else get_balance(uid)
    return InlineKeyboardMarkup(
        [
            [InlineKeyboardButton(f"{n} ta" + (" 🔒" if bal is not None and bal < n else ""),
                                  callback_data=f"n:{n}") for n in SLIDE_CHOICES],
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
        f"👤 Ism: {user['full_name']}\n👥 Guruh: {user['group_name']}\n"
        f"{'🎁 Tekin profil' if is_free(update.effective_user.id) else '💰 Balans: ' + str(get_balance(update.effective_user.id)) + ' slayd'}\n\n"
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
    if text == BTN_BALANCE:
        context.user_data["step"] = None
        return await show_balance(update, context)

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
            f"Mavzu: «{topic}»\nUslub: {THEMES[data[2:]]['label']}\n{balance_line(update.effective_user.id)}\n"
            "Nechta slayd kerak? (jami slayd soni balansdan yechiladi; 🔒 — balans yetmaydi) 👇",
            reply_markup=count_keyboard(update.effective_user.id),
        )
    if data.startswith("n:"):
        total = int(data[2:]) if data[2:].isdigit() else 0
        if total in SLIDE_CHOICES:
            return await generate(update, context, total)
    if data.startswith("t:") and data[2:].isdigit():
        return await choose_method(update, context, int(data[2:]))
    if data.startswith("ti:") and data[3:].isdigit():
        return await send_topup_invoice(update, context, int(data[3:]))
    if data.startswith("tc:") and data[3:].isdigit():
        return await card_instructions(update, context, int(data[3:]))
    if data[:3] in ("ap:", "rj:") and data[3:].isdigit():
        return await review_order(update, context, data.startswith("ap:"), int(data[3:]))
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

    # TO'LOV AVVAL: pul yechilmasa slayd yaratilmaydi (tekin profillar bundan mustasno)
    charged = False
    if not is_free(uid):
        if not try_charge(uid, total):
            context.user_data["busy"] = False
            return await q.edit_message_text(
                f"❗ Balans yetarli emas.\n{total} ta slayd kerak, balansingiz {get_balance(uid)} slayd.\n\n" + plans_help(),
                reply_markup=topup_keyboard(),
            )
        charged = True

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
            caption=f"✅ Tayyor: {data['title']}\n📄 {count} ta slayd · {style_label}"
            + (f"\n💰 Qoldi: {get_balance(uid)} slayd" if charged else ""),
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
        if charged:
            add_balance(uid, total)  # xato bo'lsa slaydlar qaytariladi
        reason = str(e) or type(e).__name__
        if "timeout" in type(e).__name__.lower() or "timed out" in reason.lower():
            reason = "AI vaqtida javob bermadi (internet yoki Google serveri sekin)."
        if any(k in reason.lower() for k in ("503", "504", "unavailable", "deadline", "band")):
            reason = "AI xizmatlari hozir band yoki sekin. 1-2 daqiqadan keyin qayta urinib ko'ring (kamroq slayd tanlasangiz tezroq bo'ladi)."
        await q.edit_message_text(
            f"❌ Xato: {reason[:300]}"
            + ("\n💰 Slaydlaringiz balansga qaytarildi." if charged else "")
            + "\n\nQayta urinib ko'ring: «🎞 Yangi slayd»."
        )
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


# ============================================================
# BALANS VA TO'LOV
# ============================================================
def plan_label(price: int) -> str:
    c = PLANS[price]
    return f"{money(price)} · 15 talik {c // 15} ta · 20 talik {c // 20} ta"


def plans_help() -> str:
    return "📋 Tariflar (taqdimot soni):\n" + "\n".join(plan_label(p) for p in PLANS)


def balance_line(uid: int) -> str:
    return "" if is_free(uid) else f"💰 Balans: {get_balance(uid)} slayd\n"


def topup_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([[InlineKeyboardButton(plan_label(p), callback_data=f"t:{p}")] for p in PLANS])


async def show_balance(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    if is_free(uid):
        return await update.message.reply_text(
            "🎁 Sizning profilingiz uchun barcha slaydlar TEKIN. Cheklov yo'q!", reply_markup=MENU
        )
    await update.message.reply_text(
        f"💰 Balansingiz: {get_balance(uid)} slayd\n"
        "Taqdimot narxi uning slayd soniga teng (5, 10, 15 yoki 20).\n\n"
        "Tarifni tanlang 👇",
        reply_markup=topup_keyboard(),
    )


async def send_topup_invoice(update: Update, context: ContextTypes.DEFAULT_TYPE, amount: int):
    q = update.callback_query
    if amount not in PLANS:
        return
    if not PROVIDER_TOKEN:
        return await q.message.reply_text(
            "To'lov tizimi hali ulanmagan. Hisobni to'ldirish uchun admin bilan bog'laning."
        )
    await context.bot.send_invoice(
        chat_id=q.message.chat_id,
        title="Hisobni to'ldirish",
        description=f"Balansga {PLANS[amount]} slayd qo'shiladi",
        payload=f"topup:{amount}",
        provider_token=PROVIDER_TOKEN,
        currency="UZS",
        prices=[LabeledPrice("Balans", amount * 100)],  # tiyinda
    )


async def on_pre_checkout(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.pre_checkout_query.answer(ok=True)


async def on_paid(update: Update, context: ContextTypes.DEFAULT_TYPE):
    pay = update.message.successful_payment
    uid = update.effective_user.id
    amount = pay.total_amount // 100
    credits = PLANS.get(amount, amount // 250)  # 250 so'm = 1 slayd (zaxira qoida)
    if record_payment(pay.telegram_payment_charge_id, uid, amount):
        add_balance(uid, credits)
    await update.message.reply_text(
        f"✅ To'lov qabul qilindi: {money(amount)}\n➕ {credits} slayd qo'shildi\n"
        f"💰 Balans: {get_balance(uid)} slayd", reply_markup=MENU
    )


# ============================================================
# ADMIN BUYRUQLARI
# ============================================================
def _admin(update: Update) -> bool:
    return bool(ADMIN_ID) and update.effective_user.id == ADMIN_ID


async def id_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(f"🆔 Sizning Telegram ID: {update.effective_user.id}")


async def addbalance_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not _admin(update):
        return
    try:
        uid, amount = int(context.args[0]), int(context.args[1])
    except (IndexError, ValueError):
        return await update.message.reply_text("Format: /addbalance <user_id> <slayd_soni>")
    add_balance(uid, amount)
    await update.message.reply_text(f"✅ {uid} balansiga {amount} slayd qo'shildi.")


async def free_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not _admin(update):
        return
    try:
        uid = int(context.args[0])
    except (IndexError, ValueError):
        return await update.message.reply_text("Format: /free <user_id>")
    set_free(uid, True)
    await update.message.reply_text(f"🎁 {uid} endi TEKIN foydalanadi.")


async def unfree_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not _admin(update):
        return
    try:
        uid = int(context.args[0])
    except (IndexError, ValueError):
        return await update.message.reply_text("Format: /unfree <user_id>")
    set_free(uid, False)
    await update.message.reply_text(f"{uid} tekin ro'yxatdan chiqarildi.")


# ============================================================
# KARTA ORQALI QO'LDA TO'LOV (chek yuboriladi, admin tasdiqlaydi)
# ============================================================
def method_keyboard(price: int) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("💳 Click / Payme (avtomatik)", callback_data=f"ti:{price}")],
        [InlineKeyboardButton("🏦 Karta orqali (chek yuborish)", callback_data=f"tc:{price}")],
    ])


async def choose_method(update: Update, context: ContextTypes.DEFAULT_TYPE, price: int):
    q = update.callback_query
    if price not in PLANS:
        return
    if PROVIDER_TOKEN and CARD_NUMBER:
        return await q.message.reply_text(
            f"{plan_label(price)}\n\nTo'lov usulini tanlang 👇", reply_markup=method_keyboard(price)
        )
    if PROVIDER_TOKEN:
        return await send_topup_invoice(update, context, price)
    if CARD_NUMBER:
        return await card_instructions(update, context, price)
    await q.message.reply_text("To'lov tizimi hali ulanmagan. Admin bilan bog'laning.")


async def card_instructions(update: Update, context: ContextTypes.DEFAULT_TYPE, price: int):
    q = update.callback_query
    if price not in PLANS or not CARD_NUMBER:
        return
    context.user_data["receipt_price"] = price
    holder = f"\n👤 Karta egasi: {CARD_HOLDER}" if CARD_HOLDER else ""
    await q.message.reply_text(
        f"🏦 Karta orqali to'lash\n\n"
        f"💵 Summa: {money(price)} ({PLANS[price]} slayd)\n"
        f"💳 Karta: {CARD_NUMBER}{holder}\n\n"
        "1) Aynan shu summani o'tkazing.\n"
        "2) To'lov chekini (skrinshot) shu yerga RASM qilib yuboring.\n"
        "3) Admin tekshirib tasdiqlagach, balansingiz to'ladi."
    )


async def on_receipt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    user = update.effective_user
    price = context.user_data.get("receipt_price")
    if not price:
        return await msg.reply_text(
            "Avval «💰 Balans» bo'limida tarifni tanlang va «Karta orqali» usulini bosing.", reply_markup=MENU
        )
    if not ADMIN_ID:
        return await msg.reply_text("Admin sozlanmagan. Keyinroq urinib ko'ring.")
    if count_pending(user.id) >= 3:
        return await msg.reply_text("Sizning 3 ta chekingiz hali tekshirilmoqda. Biroz kuting.")

    oid = create_order(user.id, price, PLANS[price])
    uname = f"@{user.username}" if user.username else "username yo'q"
    caption = (
        f"💳 To'lov cheki #{oid}\n👤 {user.full_name} ({uname})\n🆔 {user.id}\n"
        f"💵 {money(price)} → {PLANS[price]} slayd"
    )
    kb = InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ Tasdiqlash", callback_data=f"ap:{oid}"),
        InlineKeyboardButton("❌ Rad etish", callback_data=f"rj:{oid}"),
    ]])
    try:
        if msg.photo:
            await context.bot.send_photo(ADMIN_ID, msg.photo[-1].file_id, caption=caption, reply_markup=kb)
        else:
            await context.bot.send_document(ADMIN_ID, msg.document.file_id, caption=caption, reply_markup=kb)
    except Exception:
        log.exception("Chekni adminga yuborib bo'lmadi")
        set_order_status(oid, "pending", "failed")
        return await msg.reply_text("❌ Chekni adminga yuborib bo'lmadi. Keyinroq urinib ko'ring.")
    context.user_data.pop("receipt_price", None)
    await msg.reply_text("✅ Chek adminga yuborildi. Tasdiqlangach balansingiz to'ladi.", reply_markup=MENU)


async def review_order(update: Update, context: ContextTypes.DEFAULT_TYPE, approve: bool, oid: int):
    q = update.callback_query
    if not ADMIN_ID or q.from_user.id != ADMIN_ID:
        return
    order = get_order(oid)
    if not order:
        return await q.answer("Buyurtma topilmadi.", show_alert=True)
    if not set_order_status(oid, "pending", "approved" if approve else "rejected"):
        return await q.answer("Bu chek allaqachon ko'rib chiqilgan.", show_alert=True)
    if approve:
        add_balance(order["user_id"], order["credits"])
        user_text = (f"✅ To'lovingiz tasdiqlandi!\n➕ {order['credits']} slayd qo'shildi.\n"
                     f"💰 Balans: {get_balance(order['user_id'])} slayd")
    else:
        user_text = "❌ To'lovingiz tasdiqlanmadi. Chek noto'g'ri bo'lsa, qayta urinib ko'ring yoki admin bilan bog'laning."
    try:
        await context.bot.send_message(order["user_id"], user_text, reply_markup=MENU)
    except Exception:
        log.exception("Foydalanuvchiga xabar yuborib bo'lmadi")
    status = "✅ TASDIQLANDI" if approve else "❌ RAD ETILDI"
    try:
        await q.edit_message_caption(caption=(q.message.caption or "") + f"\n\n{status}")
    except Exception:
        pass


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
    app.add_handler(CommandHandler("id", id_cmd))
    app.add_handler(CommandHandler("addbalance", addbalance_cmd))
    app.add_handler(CommandHandler("free", free_cmd))
    app.add_handler(CommandHandler("unfree", unfree_cmd))
    app.add_handler(PreCheckoutQueryHandler(on_pre_checkout))
    app.add_handler(MessageHandler(filters.SUCCESSFUL_PAYMENT, on_paid))
    app.add_handler(MessageHandler(filters.PHOTO | filters.Document.IMAGE, on_receipt))
    app.add_handler(CallbackQueryHandler(on_callback))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text))
    app.add_error_handler(on_error)
    print("Bot ishga tushdi ...")
    app.run_polling()

  
if __name__ == "__main__":
    main()