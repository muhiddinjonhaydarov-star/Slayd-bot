import json
import os
import re
import time

from google import genai
from google.genai import types

client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))
# "-latest" nomi eskirib qolmasligi uchun tanlandi. Xato bersa, .env da AI_MODEL ni o'zgartiring.
MODEL = os.getenv("AI_MODEL", "gemini-flash-latest")

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
- Har slaydda 3-5 ta band, har band 20 so'zdan oshmasin.
- Sarlavhalar 8 so'zdan oshmasin.
- Reja bandlari slayd sarlavhalariga mos bo'lsin."""


def _generate(prompt: str, system: str, max_tokens: int, as_json: bool = False) -> str:
    config = types.GenerateContentConfig(
        system_instruction=system,
        max_output_tokens=max_tokens,
        response_mime_type="application/json" if as_json else None,
    )
    last_error = None
    for attempt in range(3):
        try:
            resp = client.models.generate_content(model=MODEL, contents=prompt, config=config)
            return (resp.text or "").strip()
        except Exception as e:  # bepul tarifda 429 (limit) tez-tez bo'ladi
            last_error = e
            if "429" in str(e) or "503" in str(e):
                time.sleep(6 * (attempt + 1))
                continue
            raise
    raise last_error


def ask(text: str) -> str:
    return _generate(text, CHAT_SYSTEM, 2000)


def make_slide_content(topic: str, n: int) -> dict:
    raw = _generate(SLIDE_PROMPT.format(topic=topic, n=n), SLIDE_SYSTEM, 8000, as_json=True)
    match = re.search(r"\{.*\}", raw, re.S)
    if not match:
        raise ValueError("AI JSON qaytarmadi")
    data = json.loads(match.group(0))
    for key in ("title", "plan", "slides", "conclusion"):
        if key not in data:
            raise ValueError(f"JSON da '{key}' yo'q")
    return data