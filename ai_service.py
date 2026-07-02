"""
AI xizmati - Google Gemini API orqali kontent generatsiyasi
"""
import json
from google import genai

# Default API kalit (bo'sh — foydalanuvchi o'zi qo'shadi)
_client = None
_model_name = "gemini-2.0-flash"


def init_ai(api_key: str, model: str = "gemini-2.0-flash"):
    """AI mijozni ishga tushirish"""
    global _client, _model_name
    _client = genai.Client(api_key=api_key)
    _model_name = model


def _generate(prompt: str, max_tokens: int = 4000) -> str:
    """AI orqali matn generatsiya qilish"""
    if _client is None:
        raise ValueError("AI API kaliti sozlanmagan! Admin panel orqali API kalitini qo'shing.")

    response = _client.models.generate_content(
        model=_model_name,
        contents=prompt,
        config={
            "max_output_tokens": max_tokens,
            "temperature": 0.7,
        }
    )
    return response.text


def generate_slide_content(topic: str, num_slides: int = 5, language: str = "uz") -> dict:
    """Slayd uchun kontent generatsiya qilish - professional va uzun.
    AI o'zi mavzuga mos rang palitrasini va har bir slayd uchun rasm promptini tanlaydi."""
    prompt = f"""Siz professional taqdimot dizayneri va kontent yaratuvchi ekspertsiz. Quyidagi mavzu bo'yicha {num_slides} ta slayd uchun TO'LIQ va UZUN kontent yarating.

MAVZU: {topic}
TIL: {"O'zbek" if language == "uz" else "Rus" if language == "ru" else "English"}

===== 1. RANG PALITRASI =====
Mavzuga eng mos keladigan professional rang palitrasini tanlang (HEX formatda):
- Tibbiyot/salomatlik → yashil-moviy tonlar
- Texnologiya/IT → ko'k/binafsha tonlar
- Tarix/madaniyat → jigarrang/oltin tonlar
- Tabiat/ekologiya → yashil tonlar
- Biznes/iqtisod → to'q ko'k/kulrang tonlar
- Va hokazo — mavzudan kelib chiqib O'ZINGIZ tanlang!

===== 2. HAR BIR SLAYD UCHUN =====
- Slayd sarlavhasi (qisqa, ta'sirli)
- Asosiy matn (KAMIDA 4-5 jumla, batafsil va uzun, 60-120 so'z)
- "image_prompt": shu slayd orqa foni uchun INGLIZ TILIDA rasm tavsifi.
  Rasm mavzuga va slayd mazmuniga aniq mos bo'lsin. Masalan slayd yurak
  kasalliklari haqida bo'lsa: "human heart medical illustration, cardiology,
  hospital". Har bir slaydda HAR XIL rasm prompti bo'lsin!
- Qo'shimcha fakt yoki statistika ("subtitle" maydonida)

MUHIM STRUKTURA:
- 1-slayd: Sarlavha slaydi — mavzu nomi va qisqacha kirish
- O'rta slaydlar: mavzuning TURLI jihatlari (tarixi, turlari, ahamiyati, muammolari, statistikasi, kelajagi...)
- Kamida 1-2 ta slaydda "bullet_points" massivi bo'lsin (5-6 ta punkt)
- Oxirgi slayd: Xulosa

Kontent professional, ilmiy uslubda va boy bo'lishi kerak.

JAVOB FORMATI - FAQAT JSON:
{{
  "title": "Taqdimot sarlavhasi",
  "theme_colors": {{
    "primary": "#1A5276",
    "accent": "#F39C12",
    "dark": "#0B2138",
    "light": "#EAF2F8"
  }},
  "slides": [
    {{
      "type": "title",
      "title": "Slayd sarlavhasi",
      "content": "Asosiy matn (uzun va batafsil)...",
      "subtitle": "Qo'shimcha izoh yoki fakt",
      "bullet_points": [],
      "image_prompt": "english image description for this slide background"
    }}
  ]
}}

Har bir slaydda "type": "title", "content", "bullet_list" yoki "conclusion".
"bullet_list" turida "bullet_points" massivini to'ldiring (5-6 ta punkt).
HAR BIR slaydda "image_prompt" MAJBURIY va har xil bo'lsin!
"""

    result = _generate(prompt, max_tokens=8000)
    result = result.strip()
    if result.startswith("```json"):
        result = result[7:]
    if result.startswith("```"):
        result = result[3:]
    if result.endswith("```"):
        result = result[:-3]
    return json.loads(result)


def generate_course_work_content(topic: str, language: str = "uz") -> dict:
    """Kurs ishi uchun kontent generatsiya qilish"""
    prompt = f"""Siz akademik kurs ishi yozuvchi professorsiz. Quyidagi mavzu bo'yicha TO'LIQ kurs ishi kontentini yarating.

MAVZU: {topic}
TIL: {"O'zbek" if language == "uz" else "Rus" if language == "ru" else "English"}

Kurs ishi quyidagi bo'limlardan iborat bo'lishi kerak:
1. Titul sahifasi
2. Mundarija
3. Kirish (dolzarbligi, maqsadi, vazifalari) — kamida 200 so'z
4. 1-bob: Nazariy asoslar — kamida 500 so'z, 2-3 paragraf
5. 2-bob: Amaliy tahlil — kamida 500 so'z, 2-3 paragraf
6. 3-bob: Taklif va tavsiyalar — kamida 400 so'z, 2-3 paragraf
7. Xulosa — kamida 150 so'z
8. Foydalanilgan adabiyotlar ro'yxati (kamida 8 ta)

JAVOB FORMATI - FAQAT JSON:
{{
  "title": "Kurs ishi sarlavhasi",
  "subject": "Fan nomi",
  "introduction": "Kirish matni...",
  "chapters": [
    {{
      "title": "1-bob. ...",
      "paragraphs": ["Paragraf 1...", "Paragraf 2...", "Paragraf 3..."]
    }}
  ],
  "conclusion": "Xulosa matni...",
  "references": ["Adabiyot 1", "Adabiyot 2"]
}}
"""

    result = _generate(prompt, max_tokens=8000)
    result = result.strip()
    if result.startswith("```json"):
        result = result[7:]
    if result.startswith("```"):
        result = result[3:]
    if result.endswith("```"):
        result = result[:-3]
    return json.loads(result)


def generate_slide_images_queries(topic: str, slides: list) -> list[str]:
    """Har bir slayd uchun rasm qidiruv so'rovlarini generatsiya qilish"""
    prompt = f"""Quyidagi taqdimot slaydlari uchun mos rasm qidiruv kalit so'zlarini yarating (ingliz tilida).

MAVZU: {topic}

Slaydlar:
{json.dumps([s.get('title', '') for s in slides], indent=2, ensure_ascii=False)}

Har bir slayd uchun 1 ta eng mos Unsplash/Pexels qidiruv kalit so'zi bering.
JAVOB FORMATI - FAQAT JSON massiv: ["keyword1", "keyword2", ...]"""

    result = _generate(prompt, max_tokens=500)
    result = result.strip()
    if result.startswith("```json"):
        result = result[7:]
    if result.startswith("```"):
        result = result[3:]
    if result.endswith("```"):
        result = result[:-3]
    try:
        return json.loads(result)
    except:
        return [topic] * len(slides)
