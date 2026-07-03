"""
AI xizmati - Google Gemini API orqali kontent generatsiyasi
"""
import json
import re
from google import genai


def _clean_json_text(text: str) -> str:
    """AI javobidan toza JSON ajratib olish"""
    text = text.strip()
    if text.startswith("```json"):
        text = text[7:]
    if text.startswith("```"):
        text = text[3:]
    if text.endswith("```"):
        text = text[:-3]
    text = text.strip()
    # JSON boshlanishini topish
    start = text.find('{')
    if start > 0:
        text = text[start:]
    return text


def _repair_truncated_json(text: str) -> str:
    """Kesilgan (tugallanmagan) JSONni tuzatishga urinish"""
    # Ochiq stringni yopish
    in_string = False
    escape = False
    for ch in text:
        if escape:
            escape = False
            continue
        if ch == '\\':
            escape = True
            continue
        if ch == '"':
            in_string = not in_string
    if in_string:
        text += '"'

    # Ochiq qavslarni yopish
    stack = []
    in_string = False
    escape = False
    for ch in text:
        if escape:
            escape = False
            continue
        if ch == '\\':
            escape = True
            continue
        if ch == '"':
            in_string = not in_string
            continue
        if in_string:
            continue
        if ch in '{[':
            stack.append(ch)
        elif ch == '}' and stack and stack[-1] == '{':
            stack.pop()
        elif ch == ']' and stack and stack[-1] == '[':
            stack.pop()

    # Oxiridagi vergulni olib tashlash
    text = re.sub(r',\s*$', '', text.strip())

    for br in reversed(stack):
        text += '}' if br == '{' else ']'

    return text


def _parse_json_safe(text: str) -> dict:
    """JSONni xavfsiz o'qish — kesilgan bo'lsa tuzatadi"""
    cleaned = _clean_json_text(text)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass
    # Tuzatishga urinish
    repaired = _repair_truncated_json(cleaned)
    try:
        return json.loads(repaired)
    except json.JSONDecodeError:
        pass
    # Oxirgi to'liq elementgacha kesish (oxirgi } yoki ] gacha)
    last_brace = max(cleaned.rfind('}'), cleaned.rfind(']'))
    if last_brace > 0:
        candidate = _repair_truncated_json(cleaned[:last_brace + 1])
        return json.loads(candidate)
    raise ValueError("AI javobini o'qib bo'lmadi. Qaytadan urinib ko'ring.")


def _parse_json_array_safe(text: str) -> list:
    """JSON massivni xavfsiz o'qish"""
    cleaned = text.strip()
    if cleaned.startswith("```json"):
        cleaned = cleaned[7:]
    if cleaned.startswith("```"):
        cleaned = cleaned[3:]
    if cleaned.endswith("```"):
        cleaned = cleaned[:-3]
    cleaned = cleaned.strip()
    start = cleaned.find('[')
    if start > 0:
        cleaned = cleaned[start:]
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass
    repaired = _repair_truncated_json(cleaned)
    try:
        return json.loads(repaired)
    except json.JSONDecodeError:
        return []

# Default API kalit (bo'sh — foydalanuvchi o'zi qo'shadi)
_client = None
_model_name = "gemini-2.0-flash"


def init_ai(api_key: str, model: str = "gemini-2.0-flash"):
    """AI mijozni ishga tushirish"""
    global _client, _model_name
    _client = genai.Client(api_key=api_key)
    _model_name = model


# Zaxira modellar — eng TEZLARI birinchi (flash-lite 2-3x tezroq javob beradi)
FALLBACK_MODELS = [
    "gemini-2.5-flash-lite",
    "gemini-flash-lite-latest",
    "gemini-2.0-flash-lite",
    "gemini-2.5-flash",
    "gemini-flash-latest",
    "gemini-2.0-flash",
]

# Tarif bo'yicha model tanlash: qimmatroq tarif = kuchliroq model
TARIFF_MODELS = {
    "standart": ["gemini-2.0-flash-lite", "gemini-2.5-flash-lite", "gemini-2.5-flash"],
    "premium": ["gemini-2.5-flash", "gemini-flash-latest", "gemini-2.5-flash-lite"],
    "premium_plus": ["gemini-2.5-pro", "gemini-2.5-flash", "gemini-flash-latest"],
}

# Tarif bo'yicha qo'shimcha sifat ko'rsatmalari
TARIFF_PROMPT_EXTRAS = {
    "standart": "",
    "premium": """
===== PREMIUM SIFAT TALABLARI =====
- Har bir slaydda kamida 2-3 ta ANIQ ilmiy termin ishlating va ularni qisqacha izohlang
- Statistik raqamlar, foizlar, yillarni ko'proq kiriting
- Matn akademik uslubda, ammo tushunarli bo'lsin""",
    "premium_plus": """
===== YUQORI SIFATLI PREMIUM TALABLARI =====
- Har bir slaydda 3-4 ta chuqur ilmiy termin, ta'rif va izohlar bilan
- Xalqaro tadqiqotlar, olimlar nomlari, manba yillarini kiriting
- Taqqoslash, tahlil va tanqidiy fikrlash elementlari bo'lsin
- Har slaydda kamida 80-150 so'z, mukammal akademik uslub
- Eng so'nggi tendensiyalar va kelajak prognozlarini qo'shing""",
}


def _generate(prompt: str, max_tokens: int = 4000, preferred_models: list = None) -> str:
    """AI orqali matn generatsiya qilish.
    Agar model limiti tugagan bo'lsa (429) — avtomatik boshqa modelni sinaydi."""
    if _client is None:
        raise ValueError("AI API kaliti sozlanmagan! Admin panel orqali API kalitini qo'shing.")

    # Tarif modellari > tanlangan model > zaxiralar
    if preferred_models:
        models_to_try = list(preferred_models) + [m for m in FALLBACK_MODELS if m not in preferred_models]
    else:
        models_to_try = [_model_name] + [m for m in FALLBACK_MODELS if m != _model_name]

    # O'tkazib yuboriladigan (keyingi modelni sinash kerak bo'lgan) xatolar:
    # 429 = limit tugadi, 404 = model topilmadi, 503 = model band,
    # 500 = server xatosi, UNAVAILABLE/overloaded = band
    RETRYABLE = ["429", "RESOURCE_EXHAUSTED", "404", "NOT_FOUND",
                 "503", "UNAVAILABLE", "overloaded", "high demand",
                 "500", "INTERNAL", "DEADLINE_EXCEEDED", "504"]

    last_error = None
    for model in models_to_try:
        try:
            response = _client.models.generate_content(
                model=model,
                contents=prompt,
                config={
                    "max_output_tokens": max_tokens,
                    "temperature": 0.7,
                }
            )
            if response.text:
                return response.text
            # Bo'sh javob — keyingi modelni sinaymiz
            last_error = RuntimeError(f"{model}: bo'sh javob")
            continue
        except Exception as e:
            err_str = str(e)
            last_error = e
            # Vaqtinchalik xatolar — keyingi modelni sinaymiz
            if any(marker in err_str for marker in RETRYABLE):
                continue
            # Jiddiy xato (masalan noto'g'ri API kalit) — to'xtatamiz
            raise

    raise RuntimeError(
        "Hozir barcha AI modellar band. Iltimos 1-2 daqiqadan keyin qaytadan urinib ko'ring."
    )


def generate_slide_content(topic: str, num_slides: int = 5, language: str = "uz",
                            tariff: str = "standart") -> dict:
    """Slayd uchun kontent generatsiya qilish - professional va uzun.
    AI o'zi mavzuga mos rang palitrasini va har bir slayd uchun rasm promptini tanlaydi.
    tariff: standart / premium / premium_plus — sifat darajasi"""
    prompt = f"""Siz professional taqdimot dizayneri va shu sohaning chuqur bilimli mutaxassisisiz. Quyidagi mavzu bo'yicha {num_slides} ta slayd uchun TO'LIQ va UZUN kontent yarating.

MAVZU: {topic}
TIL: {"O'zbek" if language == "uz" else "Rus" if language == "ru" else "English"}

===== 0. AVVAL MAVZUNI CHUQUR TAHLIL QILING (ichingizda, javobga yozmasdan) =====
Slaydlarni yozishdan OLDIN o'zingizga shu savollarni bering va javoblaringizni kontentda ishlating:
1. Bu mavzu ASLIDA nima? (fan sohasimi, mahsulotmi, o'yin, shaxs, hodisa, texnologiya, kasallik...?)
   MASALAN: "Garena" — bu o'yin platformasi/Free Fire noshiri, "Python" — dasturlash tili,
   "Amir Temur" — tarixiy shaxs. Mavzuni NOTO'G'RI tushunib olmang!
2. Bu mavzu qaysi sohaga tegishli va auditoriyasi kim?
3. Mavzuning eng muhim 5-7 jihati nima? (tarixi, tuzilishi, turlari, afzalliklari,
   muammolari, statistikasi, kelajagi, qiziqarli faktlar...)
4. Bu mavzu bo'yicha qanday ANIQ faktlar, raqamlar, sanalar, nomlar bor?
   Umumiy gaplar o'rniga aniq ma'lumotlar ishlating!
5. Rasm uchun eng mos vizual obrazlar qanday?

Tahlil natijasini bevosita slaydlar kontentiga singdiring — har bir slayd shu mavzuga
XOS bo'lsin, boshqa har qanday mavzuga ham to'g'ri keladigan umumiy gaplardan qoching!

===== 1. RANG PALITRASI =====
Mavzuga eng mos keladigan professional rang palitrasini tanlang (HEX formatda):
- Tibbiyot/salomatlik → yashil-moviy tonlar
- Texnologiya/IT → ko'k/binafsha tonlar
- Tarix/madaniyat → jigarrang/oltin tonlar
- Tabiat/ekologiya → yashil tonlar
- Biznes/iqtisod → to'q ko'k/kulrang tonlar
- Va hokazo — mavzudan kelib chiqib O'ZINGIZ tanlang!

===== 2. HAR BIR SLAYD UCHUN =====
- Slayd sarlavhasi (qisqa, ta'sirli, mavzuga XOS)
- Asosiy matn (KAMIDA 4-5 jumla, batafsil va uzun, 60-120 so'z):
  * ANIQ faktlar, raqamlar, sanalar, nomlar bo'lsin
  * "Bu juda muhim mavzu" kabi bo'sh gaplar TAQIQLANADI
  * Har jumla o'quvchiga YANGI ma'lumot bersin
- "image_prompt": shu slayd orqa foni uchun INGLIZ TILIDA rasm tavsifi.
  Rasm mavzuga va slayd mazmuniga aniq mos bo'lsin. Masalan slayd yurak
  kasalliklari haqida bo'lsa: "human heart medical illustration, cardiology,
  hospital". Har bir slaydda HAR XIL rasm prompti bo'lsin!
- Qo'shimcha fakt yoki statistika ("subtitle" maydonida)

MUHIM STRUKTURA — TAQDIMOT STANDARTI BO'YICHA ANIQ SHU TARTIBDA:
- 1-slayd: Sarlavha (titul) — mavzu nomi va qisqacha kirish
- 2-slayd: REJA — taqdimotda ko'riladigan bo'limlar ro'yxati ("bullet_points" bilan, type="bullet_list", title="Reja")
- 3-slayd: Kirish — mavzuga umumiy kirish, dolzarbligi
- O'rta slaydlar: mavzuning TURLI jihatlari MANTIQIY KETMA-KETLIKDA
  (ta'rif/tushuncha → tarixi → turlari/tuzilishi → ahamiyati → statistika → muammolari → kelajagi)
- Kamida 2-3 ta slaydda "bullet_points" massivi bo'lsin (5-6 ta punkt)
- Oxirgi slayd: Xulosa (type="conclusion")

QAT'IY TALAB: ANIQ {num_slides} TA slayd yozing — kam ham, ko'p ham emas!

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
{TARIFF_PROMPT_EXTRAS.get(tariff, "")}
"""

    # 25 tagacha slayd uchun katta limit
    models = TARIFF_MODELS.get(tariff)
    result = _generate(prompt, max_tokens=40000, preferred_models=models)
    data = _parse_json_safe(result)

    # ===== SLAYDLAR SONINI KAFOLATLASH =====
    # AI kam yozgan bo'lsa — yetishmagan slaydlarni qo'shimcha so'rov bilan to'ldiramiz
    slides = data.get('slides', [])
    attempts = 0
    while len(slides) < num_slides and attempts < 3:
        missing = num_slides - len(slides)
        existing_titles = "\n".join(f"- {s.get('title', '')}" for s in slides)
        extra_prompt = f"""Siz professional taqdimot yaratuvchisiz. "{topic}" mavzusi bo'yicha
taqdimotga YANA {missing} ta YANGI kontent slayd kerak.

Quyidagi mavzular ALLAQACHON yoritilgan (ularni TAKRORLAMANG):
{existing_titles}

Mavzuning BOSHQA jihatlarini yoriting (tarixi, statistikasi, texnologiyasi, muammolari,
taqqoslash, qiziqarli faktlar, mintaqaviy jihatlar, shaxslar, kelajak...).
TIL: {"O'zbek" if language == "uz" else "Rus" if language == "ru" else "English"}

Har slaydda: "type" ("content" yoki "bullet_list"), "title", "content" (60-120 so'z, aniq
faktlar bilan), "subtitle" (qiziq fakt), "bullet_points" (bullet_list uchun 5-6 ta),
"image_prompt" (ingliz tilida rasm tavsifi).

JAVOB — FAQAT JSON MASSIV:
[{{"type": "content", "title": "...", "content": "...", "subtitle": "...", "bullet_points": [], "image_prompt": "..."}}]"""

        extra_result = _generate(extra_prompt, max_tokens=30000, preferred_models=models)
        extra_slides = _parse_json_array_safe(extra_result)
        if extra_slides:
            # Xulosa slaydidan OLDIN qo'shamiz
            if slides and slides[-1].get('type') == 'conclusion':
                slides = slides[:-1] + extra_slides[:missing] + [slides[-1]]
            else:
                slides = slides + extra_slides[:missing]
        attempts += 1

    data['slides'] = slides
    return data


def generate_course_work_content(topic: str, language: str = "uz",
                                  tariff: str = "standart") -> dict:
    """Kurs ishi uchun kontent generatsiya qilish"""
    prompt = f"""Siz akademik kurs ishi yozuvchi professor va shu sohaning chuqur mutaxassisisiz. Quyidagi mavzu bo'yicha TO'LIQ kurs ishi kontentini yarating.

MAVZU: {topic}
TIL: {"O'zbek" if language == "uz" else "Rus" if language == "ru" else "English"}

===== AVVAL MAVZUNI CHUQUR TAHLIL QILING (ichingizda, javobga yozmasdan) =====
Yozishdan OLDIN aniqlang:
1. Bu mavzu ASLIDA nima va qaysi fanga tegishli? Mavzuni noto'g'ri tushunmang!
2. Bu mavzu bo'yicha qanday ANIQ ilmiy faktlar, nazariyalar, olimlar, sanalar, statistika bor?
3. Mavzuning nazariy asoslari, amaliy jihatlari va muammolari nima?
4. Qaysi real manbalar/adabiyotlar shu mavzuga mos keladi?

Tahlil natijasini matnga singdiring: har bir paragraf ANIQ ma'lumot, fakt va tahlil
bersin. "Bu mavzu juda dolzarb" kabi umumiy bo'sh gaplarni KO'P ishlatmang —
o'rniga aniq raqamlar, misollar, nomlar yozing.

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

    result = _generate(prompt, max_tokens=20000,
                       preferred_models=TARIFF_MODELS.get(tariff))
    return _parse_json_safe(result)


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
