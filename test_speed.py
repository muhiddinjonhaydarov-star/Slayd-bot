"""Gemini tezligini tekshirish: python test_speed.py"""
import os
import time

import httpx
from dotenv import load_dotenv
from google import genai
from google.genai import types

load_dotenv()
key = os.getenv("GEMINI_API_KEY")
model = os.getenv("AI_MODEL") or "gemini-flash-latest"
print("Model:", model)


def run(name, thinking_off=False, **http):
    client = genai.Client(api_key=key, http_options=types.HttpOptions(timeout=40000, **http))
    extra = {"thinking_config": types.ThinkingConfig(thinking_budget=0)} if thinking_off else {}
    cfg = types.GenerateContentConfig(max_output_tokens=400, **extra)
    t = time.time()
    try:
        r = client.models.generate_content(
            model=model, contents="Fotosintez nima? 2 gapda javob ber.", config=cfg
        )
        print(f"{name}: OK, {time.time() - t:.1f} soniya")
    except Exception as e:
        print(f"{name}: XATO, {time.time() - t:.1f} soniya -> {type(e).__name__}: {str(e)[:150]}")


ipv4 = {"client_args": {"transport": httpx.HTTPTransport(local_address="0.0.0.0")}}
run("1) Oddiy ulanish")
run("2) IPv4 ulanish", **ipv4)
run("3) IPv4 + o'ylashsiz", thinking_off=True, **ipv4)