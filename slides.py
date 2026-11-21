from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

CREAM = RGBColor(0xF6, 0xF3, 0xEC)
DARK = RGBColor(0x1C, 0x23, 0x33)
TEAL = RGBColor(0x0B, 0x65, 0x58)
CARD = RGBColor(0xCF, 0xE8, 0xE2)
GOLD = RGBColor(0xF2, 0xB8, 0x4B)
NAVY = RGBColor(0x14, 0x1B, 0x2D)
GRAY = RGBColor(0x6B, 0x72, 0x80)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

HEAD = "Georgia"
BODY = "Arial"


def _bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _rect(slide, x, y, w, h, color, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    return s


def _text(slide, x, y, w, h, text, font, size, color, bold=False, align=PP_ALIGN.LEFT,
          anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    return box


def _bullets(slide, x, y, w, h, items, size=18, color=DARK, marker="•"):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(10)
        r = p.add_run()
        r.text = f"{marker}  {item}"
        r.font.name = BODY
        r.font.size = Pt(size)
        r.font.color.rgb = color


def _footer(slide, topic, n, total):
    label = topic if len(topic) <= 40 else topic[:37] + "..."
    _text(slide, 0.7, 6.95, 9, 0.35, f"{label} · {n} / {total}", BODY, 11, GRAY)


def _content_slide(prs, title, items, topic, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s, CREAM)
    _rect(s, 0.7, 0.62, 0.12, 0.8, TEAL)
    _text(s, 1.0, 0.5, 11.6, 1.1, title, HEAD, 32, DARK, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    card = _rect(s, 0.7, 1.9, 11.93, 4.7, CARD, MSO_SHAPE.ROUNDED_RECTANGLE)
    card.adjustments[0] = 0.04
    _bullets(s, 1.1, 2.0, 11.1, 4.5, items)
    _footer(s, topic, n, total)
    return s


def build(content: dict, path: str, author: str = ""):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    topic = content["title"]
    slides = content["slides"]
    total = len(slides) + 3

    # 1. Sarlavha slaydi (to'q ko'k)
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s, NAVY)
    _text(s, 0.9, 1.3, 11, 0.5, content.get("label", "TAQDIMOT").upper(), BODY, 14, GOLD, bold=True)
    _text(s, 0.9, 1.9, 11.5, 2.4, topic, HEAD, 44, WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    _text(s, 0.9, 4.4, 11, 1.0, content.get("subtitle", ""), BODY, 20, CARD)
    if author:
        _text(s, 0.9, 6.4, 11, 0.5, author, BODY, 14, GRAY)

    # 2. Reja
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(s, CREAM)
    _rect(s, 0.7, 0.62, 0.12, 0.8, TEAL)
    _text(s, 1.0, 0.5, 11, 1.1, "Reja", HEAD, 32, DARK, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    plan = content["plan"][:7]
    box = s.shapes.add_textbox(Inches(1.0), Inches(1.9), Inches(11.3), Inches(4.7))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for i, item in enumerate(plan, 1):
        p = tf.paragraphs[0] if i == 1 else tf.add_paragraph()
        p.space_after = Pt(12)
        r1 = p.add_run()
        r1.text = f"{i}.  "
        r1.font.name = HEAD
        r1.font.size = Pt(20)
        r1.font.bold = True
        r1.font.color.rgb = TEAL
        r2 = p.add_run()
        r2.text = item
        r2.font.name = BODY
        r2.font.size = Pt(20)
        r2.font.color.rgb = DARK
    _footer(s, topic, 2, total)

    # Kontent slaydlari
    for i, sl in enumerate(slides):
        _content_slide(prs, sl["title"], sl["bullets"][:5], topic, i + 3, total)

    # Xulosa
    s = _content_slide(prs, "Xulosa", content["conclusion"][:4], topic, total, total)
    _text(s, 0.7, 6.55, 11.9, 0.4, "Eʼtiboringiz uchun rahmat!", HEAD, 16, TEAL, bold=True,
          align=PP_ALIGN.RIGHT)

    prs.save(path)
    return path