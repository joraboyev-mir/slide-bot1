"""
Professional PPTX slayd generator v2
- AI mavzuga qarab rang palitrasini tanlaydi
- Har bir slayd orqa foni AI yaratgan RASM bilan
- Har bir varaq HAR XIL dizaynda (6 xil layout aylanadi)
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from PIL import Image, ImageDraw, ImageFilter, ImageEnhance
import io
import re
import requests
import urllib.parse
import random
import logging
from concurrent.futures import ThreadPoolExecutor

log = logging.getLogger("slidebot")

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

DEFAULT_THEME = {
    "primary": RGBColor(0x0D, 0x47, 0xA1),
    "accent": RGBColor(0x00, 0xC2, 0xE0),
    "dark": RGBColor(0x0A, 0x18, 0x2E),
    "light": RGBColor(0xEA, 0xF3, 0xFF),
}


def hex_to_rgb(hex_str: str) -> RGBColor:
    """#RRGGBB -> RGBColor"""
    try:
        h = hex_str.lstrip('#')
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except Exception:
        return DEFAULT_THEME["primary"]


def build_theme(theme_colors: dict | None) -> dict:
    """AI tanlagan ranglardan tema yasash"""
    if not theme_colors:
        return dict(DEFAULT_THEME)
    theme = dict(DEFAULT_THEME)
    if theme_colors.get("primary"):
        theme["primary"] = hex_to_rgb(theme_colors["primary"])
    if theme_colors.get("accent"):
        theme["accent"] = hex_to_rgb(theme_colors["accent"])
    if theme_colors.get("dark"):
        theme["dark"] = hex_to_rgb(theme_colors["dark"])
    if theme_colors.get("light"):
        theme["light"] = hex_to_rgb(theme_colors["light"])
    return theme


# ==================== RASM YUKLASH ====================

def _is_image(resp) -> bool:
    """Javob haqiqiy rasmmi?"""
    ct = resp.headers.get('content-type', '')
    return resp.status_code == 200 and len(resp.content) > 5000 and 'image' in ct


def _fetch_stock_photo(keywords: str, width: int = 1024, height: int = 576, seed: int = 0) -> bytes | None:
    """LoremFlickr — HAQIQIY stok fotolar (juda ishonchli va tez zaxira)"""
    # kalit so'zlarni tozalash: faqat oddiy inglizcha so'zlar
    words = re.findall(r'[a-zA-Z]+', keywords)[:3]
    if not words:
        words = ["business"]
    tags = ",".join(words)
    try:
        url = f"https://loremflickr.com/{width}/{height}/{urllib.parse.quote(tags)}?lock={seed % 100}"
        resp = requests.get(url, timeout=15)
        if _is_image(resp):
            return resp.content
    except Exception:
        pass
    return None


def _fetch_ai_image(prompt: str, width: int = 1024, height: int = 576, seed: int = 0,
                    fast: bool = False) -> bytes | None:
    """Rasm olish: Pollinations AI (flux) → turbo → LoremFlickr (real foto).
    Har doim rasm qaytarishga harakat qiladi! fast=True — flux o'tkazib yuboriladi"""
    # 1-urinish: flux (eng sifatli AI rasm) — fast rejimda o'tkazib yuboriladi
    if not fast:
        try:
            enc = urllib.parse.quote(prompt[:200])
            url = (f"https://image.pollinations.ai/prompt/{enc}"
                   f"?width={width}&height={height}&nologo=true&seed={seed}&model=flux")
            resp = requests.get(url, timeout=40)
            if _is_image(resp):
                return resp.content
        except Exception:
            pass
    # 2-urinish: turbo (tez AI rasm)
    try:
        enc = urllib.parse.quote(prompt[:150])
        url = (f"https://image.pollinations.ai/prompt/{enc}"
               f"?width={width}&height={height}&nologo=true&seed={seed}&model=turbo")
        resp = requests.get(url, timeout=25)
        if _is_image(resp):
            return resp.content
    except Exception:
        pass
    # 3-urinish: LoremFlickr — haqiqiy stok foto (deyarli har doim ishlaydi)
    return _fetch_stock_photo(prompt, width, height, seed)


def _fallback_gradient(width: int, height: int, theme: dict) -> bytes:
    """Rasm yuklanmasa — chiroyli gradient fon"""
    img = Image.new('RGB', (width, height))
    draw = ImageDraw.Draw(img)
    c1 = (theme["dark"][0], theme["dark"][1], theme["dark"][2])
    c2 = (theme["primary"][0], theme["primary"][1], theme["primary"][2])
    for y in range(height):
        r = y / height
        draw.line([(0, y), (width, y)], fill=(
            int(c1[0] * (1 - r) + c2[0] * r),
            int(c1[1] * (1 - r) + c2[1] * r),
            int(c1[2] * (1 - r) + c2[2] * r),
        ))
    # dekorativ doiralar
    rnd = random.Random(42)
    overlay = Image.new('RGBA', (width, height), (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    ac = (theme["accent"][0], theme["accent"][1], theme["accent"][2])
    for _ in range(6):
        x, y = rnd.randint(0, width), rnd.randint(0, height)
        rr = rnd.randint(60, 220)
        od.ellipse([x - rr, y - rr, x + rr, y + rr], fill=(ac[0], ac[1], ac[2], 25))
    img = Image.alpha_composite(img.convert('RGBA'), overlay).convert('RGB')
    buf = io.BytesIO()
    img.save(buf, format='JPEG', quality=88)
    return buf.getvalue()


def _prepare_image(raw: bytes, width: int, height: int,
                   darken: float = 0.0, blur: float = 0.0) -> bytes:
    """Rasmni kerakli o'lchamga moslash (crop), qoraytirish va blur"""
    img = Image.open(io.BytesIO(raw)).convert('RGB')

    # Crop to aspect
    target_ratio = width / height
    w, h = img.size
    ratio = w / h
    if ratio > target_ratio:
        new_w = int(h * target_ratio)
        left = (w - new_w) // 2
        img = img.crop((left, 0, left + new_w, h))
    else:
        new_h = int(w / target_ratio)
        top = (h - new_h) // 2
        img = img.crop((0, top, w, top + new_h))
    img = img.resize((width, height), Image.LANCZOS)

    if blur > 0:
        img = img.filter(ImageFilter.GaussianBlur(radius=blur))
    if darken > 0:
        img = ImageEnhance.Brightness(img).enhance(1 - darken)

    buf = io.BytesIO()
    img.save(buf, format='JPEG', quality=88)
    return buf.getvalue()


def _rounded_image(raw: bytes, width: int, height: int, radius: int = 40) -> bytes:
    """Burchagi yumaloq rasm (PNG, shaffof burchak)"""
    prepared = _prepare_image(raw, width, height)
    img = Image.open(io.BytesIO(prepared)).convert('RGBA')
    mask = Image.new('L', (width, height), 0)
    md = ImageDraw.Draw(mask)
    md.rounded_rectangle([0, 0, width, height], radius=radius, fill=255)
    img.putalpha(mask)
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    return buf.getvalue()


# ==================== YORDAMCHI ====================

def _set_font(run, size=None, bold=None, color=None, font_name="Arial"):
    if size:
        run.font.size = size
    if bold is not None:
        run.font.bold = bold
    if color:
        run.font.color.rgb = color
    run.font.name = font_name
    rPr = run._r.get_or_add_rPr()
    rPr.set(qn('a:latin'), font_name)


def _text(slide, left, top, width, height, text, size=Pt(16), color=None,
          bold=False, align=PP_ALIGN.LEFT, anchor='t', line_spacing=1.15):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = line_spacing
    run = p.add_run()
    run.text = text
    _set_font(run, size=size, bold=bold, color=color)
    tf._txBody.bodyPr.set('anchor', anchor)
    return box


def _rect(slide, left, top, width, height, color, transparency: int = 0):
    """Rangli to'rtburchak. transparency: 0-100 (%)"""
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    sh.fill.solid()
    sh.fill.fore_color.rgb = color
    sh.line.fill.background()
    sh.shadow.inherit = False
    if transparency > 0:
        # XML orqali shaffoflik
        sPr = sh.fill._xPr.find(qn('a:solidFill'))
        srgb = sPr.find(qn('a:srgbClr'))
        alpha = srgb.makeelement(qn('a:alpha'), {'val': str(int((100 - transparency) * 1000))})
        srgb.append(alpha)
    return sh


def _full_bg_image(slide, img_bytes: bytes):
    """Slaydga to'liq fon rasmi"""
    slide.shapes.add_picture(io.BytesIO(img_bytes), 0, 0, SLIDE_W, SLIDE_H)


WHITE = RGBColor(0xFF, 0xFF, 0xFF)


# ==================== LAYOUTLAR ====================

def _layout_title(prs, slide_item, theme, img_raw, pres_title):
    """1. SARLAVHA: to'liq rasm fon + qoraytirilgan + katta sarlavha"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = _prepare_image(img_raw, 1280, 720, darken=0.55, blur=1.5)
    _full_bg_image(slide, bg)

    # aksent chiziq
    _rect(slide, Inches(1), Inches(2.1), Inches(1.6), Pt(6), theme["accent"])

    _text(slide, Inches(1), Inches(2.4), Inches(11), Inches(2.2),
          slide_item.get('title', pres_title), size=Pt(48), bold=True, color=WHITE)

    sub = slide_item.get('subtitle') or slide_item.get('content', '')[:180]
    if sub:
        _text(slide, Inches(1), Inches(4.7), Inches(10.5), Inches(1.5),
              sub, size=Pt(20), color=RGBColor(0xE8, 0xE8, 0xE8))

    # Muallif ismi (pastda)
    author = slide_item.get('author', '')
    if author:
        _rect(slide, Inches(1), Inches(6.55), Inches(0.08), Inches(0.5), theme["accent"])
        _text(slide, Inches(1.25), Inches(6.55), Inches(9), Inches(0.5),
              f"Tayyorladi: {author}", size=Pt(16), bold=True, color=WHITE)


def _layout_image_right(prs, slide_item, theme, img_raw, num, img2_raw=None):
    """2. MATN CHAP + RASM O'NG (yumaloq burchakli), pastda 2-rasm"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)
    # chap tomonda yupqa rangli panel
    _rect(slide, 0, 0, Inches(0.25), SLIDE_H, theme["primary"])

    _text(slide, Inches(0.8), Inches(0.5), Inches(6.6), Inches(1.2),
          slide_item.get('title', ''), size=Pt(30), bold=True, color=theme["dark"])
    _rect(slide, Inches(0.8), Inches(1.55), Inches(2.2), Pt(4), theme["accent"])

    if img2_raw:
        # Matn qisqaroq joy oladi, pastda 2-rasm
        _text(slide, Inches(0.8), Inches(1.95), Inches(6.4), Inches(3.1),
              slide_item.get('content', ''), size=Pt(14), color=RGBColor(0x33, 0x33, 0x3D))
        img2 = _rounded_image(img2_raw, 640, 360, radius=30)
        slide.shapes.add_picture(io.BytesIO(img2), Inches(0.8), Inches(5.0), Inches(4.2), Inches(2.1))
    else:
        _text(slide, Inches(0.8), Inches(1.95), Inches(6.4), Inches(5.1),
              slide_item.get('content', ''), size=Pt(15), color=RGBColor(0x33, 0x33, 0x3D))

    img = _rounded_image(img_raw, 720, 800, radius=40)
    slide.shapes.add_picture(io.BytesIO(img), Inches(7.7), Inches(0.9), Inches(5.0), Inches(5.6))

    _text(slide, Inches(12.5), Inches(7.05), Inches(0.6), Inches(0.35),
          str(num), size=Pt(12), color=theme["primary"], align=PP_ALIGN.RIGHT)


def _layout_image_left(prs, slide_item, theme, img_raw, num, img2_raw=None):
    """3. RASM CHAP (to'liq balandlik) + MATN O'NG rangli fonda, 2-rasm pastda"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(slide, 0, 0, SLIDE_W, SLIDE_H, theme["dark"])

    img = _prepare_image(img_raw, 640, 800)
    slide.shapes.add_picture(io.BytesIO(img), 0, 0, Inches(5.6), SLIDE_H)

    # rasm ustidan o'ngga o'tuvchi soyali panel
    _rect(slide, Inches(5.6), 0, Inches(0.12), SLIDE_H, theme["accent"])

    _text(slide, Inches(6.2), Inches(0.7), Inches(6.5), Inches(1.3),
          slide_item.get('title', ''), size=Pt(28), bold=True, color=WHITE)
    _rect(slide, Inches(6.2), Inches(1.85), Inches(1.8), Pt(4), theme["accent"])

    if img2_raw:
        _text(slide, Inches(6.2), Inches(2.25), Inches(6.5), Inches(2.6),
              slide_item.get('content', ''), size=Pt(14),
              color=RGBColor(0xE2, 0xE6, 0xEE))
        img2 = _rounded_image(img2_raw, 640, 340, radius=30)
        slide.shapes.add_picture(io.BytesIO(img2), Inches(6.2), Inches(5.0), Inches(4.4), Inches(2.2))
    else:
        _text(slide, Inches(6.2), Inches(2.25), Inches(6.5), Inches(4.6),
              slide_item.get('content', ''), size=Pt(15),
              color=RGBColor(0xE2, 0xE6, 0xEE))

    _text(slide, Inches(12.5), Inches(7.05), Inches(0.6), Inches(0.35),
          str(num), size=Pt(12), color=WHITE, align=PP_ALIGN.RIGHT)


def _layout_full_image(prs, slide_item, theme, img_raw, num):
    """4. TO'LIQ RASM FON + pastda matn paneli"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = _prepare_image(img_raw, 1280, 720, darken=0.25)
    _full_bg_image(slide, bg)

    # pastki qorong'i panel
    _rect(slide, 0, Inches(4.2), SLIDE_W, Inches(3.3), theme["dark"], transparency=18)

    _text(slide, Inches(0.9), Inches(4.5), Inches(11.5), Inches(0.9),
          slide_item.get('title', ''), size=Pt(28), bold=True, color=WHITE)
    _rect(slide, Inches(0.9), Inches(5.35), Inches(1.8), Pt(4), theme["accent"])
    _text(slide, Inches(0.9), Inches(5.6), Inches(11.5), Inches(1.7),
          slide_item.get('content', ''), size=Pt(14), color=RGBColor(0xE8, 0xE8, 0xEE))

    _text(slide, Inches(12.5), Inches(0.3), Inches(0.6), Inches(0.35),
          str(num), size=Pt(12), color=WHITE, align=PP_ALIGN.RIGHT)


def _layout_bullets(prs, slide_item, theme, img_raw, num):
    """5. NUQTALI KARTALAR: yuqorida rasm banner + pastda rangli kartalar"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(slide, 0, 0, SLIDE_W, SLIDE_H, RGBColor(0xF5, 0xF7, 0xFA))

    # yuqori banner rasm
    banner = _prepare_image(img_raw, 1280, 256, darken=0.45)
    slide.shapes.add_picture(io.BytesIO(banner), 0, 0, SLIDE_W, Inches(1.9))

    _text(slide, Inches(0.8), Inches(0.5), Inches(11.7), Inches(1.0),
          slide_item.get('title', ''), size=Pt(30), bold=True, color=WHITE)

    items = slide_item.get('bullet_points') or []
    if not items:
        items = [s.strip() for s in slide_item.get('content', '').split('.') if s.strip()][:6]
    items = [it.strip().lstrip('-•* ') for it in items if it.strip()]

    img2_raw = slide_item.get('_img2')
    if img2_raw:
        # O'ngda 2-rasm, kartalar chapda bitta ustun bo'lib
        items = items[:3]
        img2 = _rounded_image(img2_raw, 560, 600, radius=30)
        slide.shapes.add_picture(io.BytesIO(img2), Inches(8.6), Inches(2.25), Inches(4.2), Inches(4.5))
        card_w_override = Inches(7.6)
    else:
        items = items[:6]
        card_w_override = None

    # kartalar: 2 ustunli (yoki rasm bo'lsa 1 ustunli keng)
    top0 = Inches(2.25)
    card_w = card_w_override or Inches(6.0)
    card_h = Inches(1.5)
    gap = Inches(0.25)
    single_col = card_w_override is not None
    for i, item in enumerate(items):
        col = 0 if single_col else i % 2
        row = i if single_col else i // 2
        left = Inches(0.5) + col * (card_w + gap)
        top = top0 + row * (card_h + gap)
        card = _rect(slide, left, top, card_w, card_h, WHITE)
        card.line.color.rgb = RGBColor(0xE0, 0xE4, 0xEA)
        card.line.width = Pt(1)
        # raqamli aksent doira
        circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, left + Inches(0.18),
                                      top + Inches(0.45), Inches(0.55), Inches(0.55))
        circ.fill.solid()
        circ.fill.fore_color.rgb = theme["accent"] if i % 2 == 0 else theme["primary"]
        circ.line.fill.background()
        tfc = circ.text_frame
        pc = tfc.paragraphs[0]
        pc.alignment = PP_ALIGN.CENTER
        rc = pc.add_run()
        rc.text = str(i + 1)
        _set_font(rc, size=Pt(16), bold=True, color=WHITE)

        _text(slide, left + Inches(0.95), top + Inches(0.12), card_w - Inches(1.1),
              card_h - Inches(0.24), item, size=Pt(12.5),
              color=RGBColor(0x2A, 0x2A, 0x35), anchor='ctr')

    _text(slide, Inches(12.5), Inches(7.05), Inches(0.6), Inches(0.35),
          str(num), size=Pt(12), color=theme["primary"], align=PP_ALIGN.RIGHT)


def _layout_split_diagonal(prs, slide_item, theme, img_raw, num, img2_raw=None):
    """6. YUQORI RASM + PASTKI MATN + o'ngda 2-rasm"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)

    # yuqori yarmida keng rasm
    img = _prepare_image(img_raw, 1280, 448, darken=0.15)
    slide.shapes.add_picture(io.BytesIO(img), 0, 0, SLIDE_W, Inches(3.3))

    # rasm ustiga chapdan rangli sarlavha paneli
    _rect(slide, Inches(0.7), Inches(2.55), Inches(7.6), Inches(1.35), theme["primary"])
    _text(slide, Inches(1.0), Inches(2.7), Inches(7.0), Inches(1.05),
          slide_item.get('title', ''), size=Pt(24), bold=True, color=WHITE, anchor='ctr')

    if img2_raw:
        # Matn chapda, 2-rasm o'ngda
        _text(slide, Inches(0.8), Inches(4.25), Inches(7.6), Inches(2.9),
              slide_item.get('content', ''), size=Pt(14), color=RGBColor(0x33, 0x33, 0x3D))
        img2 = _rounded_image(img2_raw, 560, 420, radius=30)
        slide.shapes.add_picture(io.BytesIO(img2), Inches(8.8), Inches(4.2), Inches(3.9), Inches(2.9))
    else:
        _text(slide, Inches(0.8), Inches(4.25), Inches(11.7), Inches(2.9),
              slide_item.get('content', ''), size=Pt(15), color=RGBColor(0x33, 0x33, 0x3D))

    sub = slide_item.get('subtitle', '')
    if sub:
        _rect(slide, Inches(0.8), Inches(6.75), Inches(11.7), Pt(2), theme["accent"])
        _text(slide, Inches(0.8), Inches(6.85), Inches(11.7), Inches(0.5),
              "💡 " + sub, size=Pt(12), color=theme["primary"])

    _text(slide, Inches(12.5), Inches(7.05), Inches(0.6), Inches(0.35),
          str(num), size=Pt(12), color=theme["primary"], align=PP_ALIGN.RIGHT)


def _layout_conclusion(prs, slide_item, theme, img_raw):
    """7. XULOSA: to'liq rasm + markazda matn"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = _prepare_image(img_raw, 1280, 720, darken=0.6, blur=2.0)
    _full_bg_image(slide, bg)

    _rect(slide, Inches(5.85), Inches(1.15), Inches(1.6), Pt(6), theme["accent"])
    _text(slide, Inches(1.5), Inches(1.5), Inches(10.3), Inches(1.1),
          slide_item.get('title', 'Xulosa'), size=Pt(36), bold=True,
          color=WHITE, align=PP_ALIGN.CENTER)
    _text(slide, Inches(1.8), Inches(2.9), Inches(9.7), Inches(3.6),
          slide_item.get('content', ''), size=Pt(17),
          color=RGBColor(0xEC, 0xEC, 0xF2), align=PP_ALIGN.CENTER, line_spacing=1.3)


def _layout_thanks(prs, theme, img_raw, pres_title):
    """8. RAHMAT slaydi"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = _prepare_image(img_raw, 1280, 720, darken=0.65, blur=3.0)
    _full_bg_image(slide, bg)

    _text(slide, Inches(1.5), Inches(2.6), Inches(10.3), Inches(1.4),
          "E'tiboringiz uchun rahmat!", size=Pt(44), bold=True,
          color=WHITE, align=PP_ALIGN.CENTER)
    _rect(slide, Inches(5.6), Inches(4.15), Inches(2.1), Pt(5), theme["accent"])
    _text(slide, Inches(1.5), Inches(4.5), Inches(10.3), Inches(0.9),
          pres_title, size=Pt(18), color=RGBColor(0xD8, 0xD8, 0xE2),
          align=PP_ALIGN.CENTER)


# ==================== ASOSIY GENERATOR ====================

def generate_professional_pptx(slide_data: dict, image_prompts: list[str] = None,
                                progress_callback=None) -> bytes:
    """
    Professional PPTX yaratish.

    slide_data = {
      "title": "...",
      "theme_colors": {"primary": "#..", "accent": "#..", "dark": "#..", "light": "#.."},
      "slides": [ {"type": "...", "title": "...", "content": "...",
                   "subtitle": "...", "bullet_points": [], "image_prompt": "..."} ]
    }
    """
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    theme = build_theme(slide_data.get('theme_colors'))
    slides = slide_data.get('slides', [])
    pres_title = slide_data.get('title', 'Taqdimot')

    # ---- Rasmlarni PARALLEL yuklash (tezlik uchun) ----
    def _load_one(args):
        i, s = args
        prompt = s.get('image_prompt') or (image_prompts[i] if image_prompts and i < len(image_prompts) else pres_title)
        full_prompt = (f"{prompt}, ultra realistic professional stock photography, 4k quality, "
                       f"sharp focus, cinematic lighting, detailed, "
                       f"no text, no watermark, no cartoon, no illustration")
        # 12 tadan ko'p slaydda flux o'rniga turbo (tezlik uchun)
        fast_mode = len(slides) > 12
        raw = _fetch_ai_image(full_prompt, seed=1000 + i * 7, fast=fast_mode)
        if raw is None:
            log.warning(f"Rasm {i+1}: AI ham, stok ham ishlamadi")
        else:
            log.info(f"Rasm {i+1}/{len(slides)} tayyor")
        return i, raw

    images = [None] * len(slides)
    # ASOSIY rasmlar: 3 tadan parallel (Pollinations limitiga tushmaslik uchun!)
    # Ko'p parallel so'rov yuborsak Pollinations rad etadi va rasmlar chiqmaydi
    with ThreadPoolExecutor(max_workers=3) as pool:
        for i, raw in pool.map(_load_one, enumerate(slides)):
            images[i] = raw
    real_count = sum(1 for im in images if im is not None)
    log.info(f"Asosiy rasmlar: {real_count}/{len(slides)} ta yuklandi")

    # Yuklanmaganlarga: real stok foto → gradient
    for i in range(len(slides)):
        if images[i] is None:
            prompt = slides[i].get('image_prompt') or pres_title
            images[i] = _fetch_stock_photo(prompt, seed=i * 3 + 1) or _fallback_gradient(1024, 576, theme)

    # ---- QO'SHIMCHA (ikkinchi) rasmlar — real stok fotolardan (tez va ishonchli) ----
    def _load_secondary(args):
        i, s = args
        prompt = s.get('image_prompt') or pres_title
        raw = _fetch_stock_photo(prompt, seed=5000 + i * 13)
        if raw is None:
            raw = _fallback_gradient(1024, 576, theme)
        return i, raw

    images2 = {}
    middle = [(i, s) for i, s in enumerate(slides) if 0 < i < len(slides) - 1]
    if middle:
        with ThreadPoolExecutor(max_workers=8) as pool:
            for i, raw in pool.map(_load_secondary, middle):
                if raw:
                    images2[i] = raw
    log.info(f"Qo'shimcha rasmlar tayyor: {len(images2)} ta")

    # Rahmat slaydi uchun rasm — birinchi rasmni ishlatamiz
    thanks_img = images[0] if images else _fallback_gradient(1280, 720, theme)

    # ---- Layout rotatsiyasi: hech qaysi ketma-ket slayd bir xil emas ----
    content_layouts = [
        _layout_image_right,
        _layout_full_image,
        _layout_image_left,
        _layout_bullets,
        _layout_split_diagonal,
    ]

    author = slide_data.get('author', '')

    for idx, item in enumerate(slides):
        img = images[idx]
        img2 = images2.get(idx)
        if idx == 0:
            if author:
                item = dict(item)
                item['author'] = author
            _layout_title(prs, item, theme, img, pres_title)
        elif idx == len(slides) - 1 and len(slides) > 2:
            _layout_conclusion(prs, item, theme, img)
        else:
            # bullet_points bo'lsa — bullets layout, aks holda rotatsiya
            if item.get('bullet_points') and len(item['bullet_points']) >= 3:
                item = dict(item)
                item['_img2'] = img2
                _layout_bullets(prs, item, theme, img, idx + 1)
            else:
                # 2-rasmni qo'llab-quvvatlaydigan layoutlar rotatsiyasi
                non_bullet = [_layout_image_right, _layout_full_image,
                              _layout_image_left, _layout_split_diagonal]
                layout = non_bullet[(idx - 1) % len(non_bullet)]
                if layout is _layout_full_image:
                    layout(prs, item, theme, img, idx + 1)
                else:
                    layout(prs, item, theme, img, idx + 1, img2_raw=img2)

    _layout_thanks(prs, theme, thanks_img, pres_title)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.getvalue()
