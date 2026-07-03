"""
Asosiy Telegram Bot - Slayd & Kurs ishi yaratuvchi
Professional PPTX va DOCX generatsiya
"""
import asyncio
import io
import json
import logging
import os
import sys
from pathlib import Path

# Loglarni sozlash — Railway'da ko'rinadi
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    stream=sys.stdout,
)
log = logging.getLogger("slidebot")

from aiogram import Bot, Dispatcher, types, F
from aiogram.client.default import DefaultBotProperties
from aiogram.filters import Command, CommandStart
from aiogram.types import (
    ReplyKeyboardMarkup, KeyboardButton, WebAppInfo,
    InlineKeyboardMarkup, InlineKeyboardButton,
    BufferedInputFile, Message, CallbackQuery,
)
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.enums import ParseMode
from dotenv import load_dotenv

from database import (
    init_db, add_user, get_user, get_users_count,
    increment_user_stats, add_generation_record,
    get_generation_stats, get_user_today_count,
    get_setting, set_setting, get_all_users,
    block_user, set_admin, is_user_blocked,
)
from ai_service import init_ai, generate_slide_content, generate_course_work_content
from slides_generator import generate_professional_pptx
from docx_generator import generate_professional_docx

load_dotenv()

# ============ KONFIGURATSIYA ============
TOKEN = os.getenv("BOT_TOKEN", "")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
ADMIN_IDS = list(map(int, os.getenv("ADMIN_IDS", "").split(","))) if os.getenv("ADMIN_IDS") else []

# Netlify linklar
SLIDES_WEBAPP_URL = os.getenv("SLIDES_WEBAPP_URL", "https://genuine-fox-5d5d9a.netlify.app/slides.html")
KURS_WEBAPP_URL = os.getenv("KURS_WEBAPP_URL", "https://heroic-eclair-466aa4.netlify.app/kurs.html")
ADMIN_WEBAPP_URL = os.getenv("ADMIN_WEBAPP_URL", "https://vocal-licorice-109183.netlify.app/admin.html")

# ============ BOT & DISPATCHER ============
bot = Bot(token=TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
dp = Dispatcher(storage=MemoryStorage())


# ============ KLAVIATURALAR ============
def main_keyboard():
    """Asosiy reply klaviatura (pastki)"""
    return ReplyKeyboardMarkup(
        keyboard=[
            [
                KeyboardButton(text="📊 Slayd yaratish", web_app=WebAppInfo(url=SLIDES_WEBAPP_URL)),
                KeyboardButton(text="📝 Kurs ishi yaratish", web_app=WebAppInfo(url=KURS_WEBAPP_URL)),
            ],
            [
                KeyboardButton(text="📊 Statistikam"),
                KeyboardButton(text="ℹ️ Yordam"),
            ],
        ],
        resize_keyboard=True,
        persistent=True,
    )


async def admin_inline_keyboard():
    """Admin tugmalari — statistika URL orqali panelga uzatiladi"""
    # Statistikani yig'ish
    try:
        total_users = await get_users_count()
        stats = await get_generation_stats(30)
        model = await get_setting("ai_model") or "gemini-2.5-flash-lite"
        max_s = await get_setting("max_slides_per_day") or "10"
        max_c = await get_setting("max_courses_per_day") or "5"
        sep = "&" if "?" in ADMIN_WEBAPP_URL else "?"
        admin_url = (
            f"{ADMIN_WEBAPP_URL}{sep}"
            f"users={total_users}"
            f"&slides={stats.get('slide', 0)}"
            f"&courses={stats.get('course', 0)}"
            f"&total={stats.get('total', 0)}"
            f"&model={model}"
            f"&limit_s={max_s}"
            f"&limit_c={max_c}"
        )
    except Exception:
        admin_url = ADMIN_WEBAPP_URL

    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⚙️ Admin Panel", web_app=WebAppInfo(url=admin_url))],
        [
            InlineKeyboardButton(text="👥 Foydalanuvchilar", callback_data="admin_users"),
            InlineKeyboardButton(text="📊 Statistika", callback_data="admin_stats"),
        ],
        [
            InlineKeyboardButton(text="📨 Xabar yuborish", callback_data="admin_broadcast"),
            InlineKeyboardButton(text="⚙️ Sozlamalar", callback_data="admin_settings"),
        ],
    ])


# ============ STATE MACHINE ============
class BroadcastState(StatesGroup):
    waiting_for_message = State()


class SettingsState(StatesGroup):
    waiting_for_api_key = State()
    waiting_for_model = State()


# ============ /start ============
@dp.message(CommandStart())
async def cmd_start(message: Message):
    """Start komandasi"""
    if not message.from_user:
        return

    user_id = message.from_user.id

    # Bloklanganmi?
    if await is_user_blocked(user_id):
        await message.answer("❌ Siz bloklangansiz. Admin bilan bog'laning.")
        return

    # Bazaga qo'shish
    await add_user(
        user_id=user_id,
        username=message.from_user.username,
        first_name=message.from_user.first_name,
        last_name=message.from_user.last_name,
    )

    # Adminmi?
    if user_id in ADMIN_IDS:
        await set_admin(user_id, True)

    welcome_text = await get_setting("welcome_message")
    if not welcome_text:
        welcome_text = "Assalomu alaykum! Slayd yoki Kurs ishi yaratish uchun pastdagi tugmalarni bosing."

    await message.answer(
        welcome_text,
        reply_markup=main_keyboard(),
    )

    # Admin bo'lsa qo'shimcha admin panel tugmasi
    if user_id in ADMIN_IDS or await get_setting(f"admin_{user_id}") == "1":
        await message.answer(
            "🔧 <b>Admin panel:</b>",
            reply_markup=await admin_inline_keyboard(),
        )


# ============ WEB APP MA'LUMOTLARI ============
@dp.message(F.web_app_data)
async def handle_web_app_data(message: Message):
    """Mini Appdan kelgan ma'lumotlarni qayta ishlash"""
    if not message.from_user:
        return

    user_id = message.from_user.id

    if await is_user_blocked(user_id):
        await message.answer("❌ Siz bloklangansiz.")
        return

    try:
        data = json.loads(message.web_app_data.data)
    except json.JSONDecodeError:
        await message.answer("❌ Noto'g'ri ma'lumot formati.")
        return

    data_type = data.get("type", "")
    log.info(f"WebApp so'rov keldi: user={user_id}, type={data_type}, data={str(data)[:200]}")

    if data_type == "slide":
        await process_slide_request(message, user_id, data)
    elif data_type == "course":
        await process_course_request(message, user_id, data)
    elif data_type == "admin_action":
        await process_admin_action(message, user_id, data)
    elif data_type == "admin_setting":
        await process_admin_setting(message, user_id, data)
    elif data_type == "admin_broadcast":
        await process_admin_broadcast(message, user_id, data)
    else:
        await message.answer("❌ Noma'lum so'rov turi.")


async def process_slide_request(message: Message, user_id: int, data: dict):
    """Slayd so'rovini qayta ishlash — AI dizaynni mavzuga qarab o'zi tanlaydi"""
    topic = data.get("topic", "").strip()
    slide_count = int(data.get("slideCount", 5))
    language = data.get("language", "uz")

    if not topic:
        await message.answer("❌ Mavzu kiritilmagan. Iltimos qaytadan urinib ko'ring.")
        return

    # Bugungi limit tekshirish
    today_count = await get_user_today_count(user_id, "slide")
    max_per_day = int(await get_setting("max_slides_per_day") or "10")
    if today_count >= max_per_day:
        await message.answer(
            f"⚠️ Bugungi limit tugadi ({max_per_day} ta). Ertaga yana urinib ko'ring."
        )
        return

    # Kutish xabari
    wait_msg = await message.answer(
        "⏳ <b>Slayd tayyorlanmoqda...</b>\n\n"
        f"📌 Mavzu: {topic}\n"
        f"📑 Slaydlar soni: {slide_count}\n"
        f"🎨 Dizayn: AI mavzuga mos tanlaydi\n"
        f"🖼 Rasmlar: har bir slayd uchun AI yaratadi\n\n"
        "Bu jarayon 1-3 daqiqa vaqt olishi mumkin, iltimos kuting..."
    )

    try:
        log.info(f"SLAYD boshlandi: topic='{topic}', count={slide_count}")
        # AI kontent generatsiya (ranglar + rasm promptlari bilan)
        # Maksimal 3 daqiqa kutamiz
        slide_data = await asyncio.wait_for(
            asyncio.to_thread(generate_slide_content, topic, slide_count, language),
            timeout=180
        )
        log.info(f"SLAYD matn tayyor: {len(slide_data.get('slides', []))} ta slayd")

        await wait_msg.edit_text(
            "⏳ <b>Slayd tayyorlanmoqda...</b>\n\n"
            f"📌 Mavzu: {topic}\n"
            f"✅ Matn tayyor!\n"
            f"🖼 Endi rasmlar yaratilmoqda ({slide_count} ta)...\n\n"
            "Yana 1-2 daqiqa kuting..."
        )

        # Muallif ismini qo'shish
        author = data.get("author", "").strip()
        if author:
            slide_data["author"] = author

        # PPTX generatsiya (rasmlar yuklab olinadi — maksimal 8 daqiqa, 25 slayd uchun)
        pptx_bytes = await asyncio.wait_for(
            asyncio.to_thread(generate_professional_pptx, slide_data),
            timeout=480
        )
        log.info(f"SLAYD PPTX tayyor: {len(pptx_bytes)} bayt")

        # Statistikani yangilash
        await increment_user_stats(user_id, "slide")
        await add_generation_record(
            user_id, "slide", topic,
            {"slide_count": slide_count}
        )

        # Faylni jo'natish
        file = BufferedInputFile(pptx_bytes, filename=f"slayd_{topic[:30].replace(' ', '_')}.pptx")
        await message.answer_document(
            file,
            caption=f"✅ <b>Tayyor!</b>\n\n"
                    f"📌 Mavzu: {topic}\n"
                    f"📑 Slaydlar: {len(slide_data.get('slides', []))} ta\n"
                    f"🎨 Dizayn: AI tomonidan mavzuga moslab tanlandi\n"
                    f"🖼 Har bir varaq har xil dizaynda, rasmli fonlar bilan\n\n"
                    f"PowerPoint (.pptx) formatida. Yuklab oling!"
        )

        await wait_msg.delete()

    except asyncio.TimeoutError:
        log.error("SLAYD: vaqt tugadi (timeout)")
        try:
            await wait_msg.delete()
        except Exception:
            pass
        await message.answer(
            "⏱ Vaqt tugadi — server juda band. Iltimos 2-3 daqiqadan keyin qaytadan urinib ko'ring."
        )
    except Exception as e:
        log.exception(f"SLAYD xato: {e}")
        try:
            await wait_msg.delete()
        except Exception:
            pass
        await message.answer(
            f"❌ Xatolik yuz berdi: {str(e)[:200]}\n\n"
            "Iltimos qaytadan urinib ko'ring yoki admin bilan bog'laning.",
            disable_web_page_preview=True
        )


async def process_course_request(message: Message, user_id: int, data: dict):
    """Kurs ishi so'rovini qayta ishlash"""
    topic = data.get("topic", "").strip()
    language = data.get("language", "uz")

    if not topic:
        await message.answer("❌ Mavzu kiritilmagan.")
        return

    # Limit
    today_count = await get_user_today_count(user_id, "course")
    max_per_day = int(await get_setting("max_courses_per_day") or "5")
    if today_count >= max_per_day:
        await message.answer(
            f"⚠️ Bugungi limit tugadi ({max_per_day} ta)."
        )
        return

    wait_msg = await message.answer(
        f"⏳ <b>Kurs ishi tayyorlanmoqda...</b>\n\n"
        f"📌 Mavzu: {topic}\n\n"
        "Bu jarayon 1-2 daqiqa vaqt olishi mumkin..."
    )

    try:
        # AI kontent (maksimal 4 daqiqa)
        course_data = await asyncio.wait_for(
            asyncio.to_thread(generate_course_work_content, topic, language),
            timeout=240
        )

        # Titul sahifasi ma'lumotlari
        course_data["author"] = data.get("author", "").strip()
        course_data["teacher"] = data.get("teacher", "").strip()
        course_data["university"] = data.get("university", "").strip()
        course_data["group"] = data.get("group", "").strip()

        # DOCX generatsiya
        docx_bytes = await asyncio.to_thread(
            generate_professional_docx, course_data
        )

        # Statistika
        await increment_user_stats(user_id, "course")
        await add_generation_record(user_id, "course", topic, {"language": language})

        # Jo'natish
        file = BufferedInputFile(docx_bytes, filename=f"kurs_ishi_{topic[:30].replace(' ', '_')}.docx")
        await message.answer_document(
            file,
            caption=f"✅ <b>Kurs ishi tayyor!</b>\n\n"
                    f"📌 Mavzu: {topic}\n"
                    f"📄 Format: Microsoft Word (.docx)\n\n"
                    f"Yuklab oling va kerakli tahrirlarni kiriting!"
        )

        await wait_msg.delete()

    except asyncio.TimeoutError:
        try:
            await wait_msg.delete()
        except Exception:
            pass
        await message.answer("⏱ Vaqt tugadi. Iltimos 2-3 daqiqadan keyin qaytadan urinib ko'ring.")
    except Exception as e:
        try:
            await wait_msg.delete()
        except Exception:
            pass
        await message.answer(f"❌ Xatolik: {str(e)[:200]}", disable_web_page_preview=True)


async def process_admin_action(message: Message, user_id: int, data: dict):
    """Admin amallarini qayta ishlash"""
    if user_id not in ADMIN_IDS and await get_setting(f"admin_{user_id}") != "1":
        await message.answer("❌ Ruxsat yo'q.")
        return

    action = data.get("action", "")
    if action == "admin_stats":
        total_users = await get_users_count()
        stats = await get_generation_stats(30)
        await message.answer(
            f"📊 <b>Statistika (30 kunlik):</b>\n\n"
            f"👥 Foydalanuvchilar: {total_users}\n"
            f"📑 Slaydlar: {stats.get('slide', 0)}\n"
            f"📝 Kurs ishlari: {stats.get('course', 0)}\n"
            f"📦 Jami: {stats.get('total', 0)}"
        )
    elif action == "admin_users":
        users = await get_all_users()
        text = f"👥 <b>Foydalanuvchilar ({len(users)}):</b>\n\n"
        for u in users[:15]:
            name = u.get('first_name', '?') or '?'
            text += f"• <code>{u['user_id']}</code> — {name}\n"
        await message.answer(text)


async def process_admin_setting(message: Message, user_id: int, data: dict):
    """Admin sozlamalarini qayta ishlash"""
    if user_id not in ADMIN_IDS and await get_setting(f"admin_{user_id}") != "1":
        await message.answer("❌ Ruxsat yo'q.")
        return

    # Single setting
    setting = data.get("setting")
    value = data.get("value")
    if setting and value:
        await set_setting(setting, value)
        if setting == "api_key":
            model = await get_setting("ai_model") or "gemini-2.0-flash"
            init_ai(value, model)
        await message.answer(f"✅ Sozlama saqlandi: {setting}")
        return

    # Multiple settings
    settings = data.get("settings", {})
    if settings:
        for key, val in settings.items():
            await set_setting(key, str(val))
        await message.answer(f"✅ {len(settings)} ta sozlama saqlandi!")


async def process_admin_broadcast(message: Message, user_id: int, data: dict):
    """Admin xabar yuborish"""
    if user_id not in ADMIN_IDS and await get_setting(f"admin_{user_id}") != "1":
        await message.answer("❌ Ruxsat yo'q.")
        return

    msg = data.get("message", "")
    if not msg:
        await message.answer("❌ Xabar matni bo'sh.")
        return

    users = await get_all_users()
    success = 0
    failed = 0

    await message.answer(f"📨 {len(users)} ta foydalanuvchiga yuborilmoqda...")

    for user in users:
        try:
            await bot.send_message(user['user_id'], msg)
            success += 1
            await asyncio.sleep(0.05)
        except:
            failed += 1

    await message.answer(f"✅ Natija: {success} yuborildi, {failed} xato.")


# ============ ODDIY KOMANDALAR ============
@dp.message(F.text == "📊 Statistikam")
async def my_stats(message: Message):
    """Foydalanuvchi statistikasi"""
    if not message.from_user:
        return

    user = await get_user(message.from_user.id)
    if not user:
        await message.answer("Siz hali botdan foydalanmagansiz.")
        return

    await message.answer(
        f"📊 <b>Sizning statistikangiz:</b>\n\n"
        f"👤 Ism: {user.get('first_name', 'Noma\'lum')}\n"
        f"📑 Yaratilgan slaydlar: {user.get('total_slides', 0)} ta\n"
        f"📝 Yaratilgan kurs ishlari: {user.get('total_courses', 0)} ta\n"
        f"📅 Qo'shilgan sana: {user.get('joined_at', 'Noma\'lum')}\n"
    )


@dp.message(F.text == "ℹ️ Yordam")
async def help_command(message: Message):
    """Yordam"""
    await message.answer(
        "ℹ️ <b>Bot haqida:</b>\n\n"
        "Bu bot sun'iy intellekt yordamida professional slaydlar (.pptx) va "
        "kurs ishlari (.docx) yaratib beradi.\n\n"
        "<b>Qanday foydalaniladi:</b>\n"
        "1. Pastdagi <b>📊 Slayd yaratish</b> tugmasini bosing\n"
        "2. Mavzu, slaydlar soni va dizaynni tanlang\n"
        "3. \"Yaratish\" tugmasini bosing\n"
        "4. Tayyor faylni yuklab oling!\n\n"
        "<b>Kurs ishi uchun:</b>\n"
        "1. <b>📝 Kurs ishi yaratish</b> tugmasini bosing\n"
        "2. Mavzuni kiriting\n"
        "3. Kurs ishi DOCX formatda tayyor bo'ladi\n\n"
        "⚠️ <b>Cheklovlar:</b> Kuniga 10 ta slayd, 5 ta kurs ishi.",
        reply_markup=main_keyboard(),
    )


# ============ ADMIN KOMANDALAR ============
@dp.message(Command("admin"))
async def admin_command(message: Message):
    """Admin panelga kirish"""
    if not message.from_user:
        return

    user_id = message.from_user.id
    if user_id not in ADMIN_IDS and await get_setting(f"admin_{user_id}") != "1":
        await message.answer("❌ Ruxsat yo'q.")
        return

    await message.answer(
        "🔧 <b>Admin panel:</b>",
        reply_markup=await admin_inline_keyboard(),
    )


@dp.callback_query(F.data == "admin_stats")
async def admin_stats(callback: CallbackQuery):
    """Admin statistikasi"""
    if not callback.from_user:
        return

    user_id = callback.from_user.id
    if user_id not in ADMIN_IDS and await get_setting(f"admin_{user_id}") != "1":
        await callback.answer("❌ Ruxsat yo'q.", show_alert=True)
        return

    total_users = await get_users_count()
    stats = await get_generation_stats(30)
    await callback.message.edit_text(
        f"📊 <b>Bot statistikasi (30 kunlik):</b>\n\n"
        f"👥 Foydalanuvchilar: {total_users} ta\n"
        f"📑 Slaydlar: {stats.get('slide', 0)} ta\n"
        f"📝 Kurs ishlari: {stats.get('course', 0)} ta\n"
        f"📦 Jami generatsiyalar: {stats.get('total', 0)} ta\n",
        reply_markup=await admin_inline_keyboard(),
    )
    await callback.answer()


@dp.callback_query(F.data == "admin_users")
async def admin_users_list(callback: CallbackQuery):
    """Foydalanuvchilar ro'yxati"""
    if not callback.from_user:
        return

    user_id = callback.from_user.id
    if user_id not in ADMIN_IDS and await get_setting(f"admin_{user_id}") != "1":
        await callback.answer("❌ Ruxsat yo'q.", show_alert=True)
        return

    users = await get_all_users()
    if not users:
        await callback.message.edit_text("👥 Hali foydalanuvchilar yo'q.")
        return

    text = f"👥 <b>Foydalanuvchilar ({len(users)} ta):</b>\n\n"
    for u in users[:20]:
        name = u.get('first_name', 'Noma\'lum') or 'Noma\'lum'
        text += f"• <code>{u['user_id']}</code> — {name} | 📑{u.get('total_slides', 0)} 📝{u.get('total_courses', 0)}\n"

    await callback.message.edit_text(text, reply_markup=await admin_inline_keyboard())
    await callback.answer()


@dp.callback_query(F.data == "admin_broadcast")
async def admin_broadcast_start(callback: CallbackQuery, state: FSMContext):
    """Xabar yuborish"""
    if not callback.from_user:
        return

    user_id = callback.from_user.id
    if user_id not in ADMIN_IDS and await get_setting(f"admin_{user_id}") != "1":
        await callback.answer("❌ Ruxsat yo'q.", show_alert=True)
        return

    await callback.message.answer(
        "📨 <b>Barcha foydalanuvchilarga yuboriladigan xabarni kiriting:</b>\n"
        "(Bekor qilish uchun /cancel)"
    )
    await state.set_state(BroadcastState.waiting_for_message)
    await callback.answer()


@dp.message(BroadcastState.waiting_for_message)
async def broadcast_send(message: Message, state: FSMContext):
    """Xabarni yuborish"""
    if not message.from_user:
        return

    user_id = message.from_user.id
    if user_id not in ADMIN_IDS and await get_setting(f"admin_{user_id}") != "1":
        return

    if message.text == "/cancel":
        await state.clear()
        await message.answer("❌ Bekor qilindi.")
        return

    users = await get_all_users()
    success = 0
    failed = 0

    await message.answer(f"📨 Xabar {len(users)} ta foydalanuvchiga yuborilmoqda...")

    for user in users:
        try:
            await bot.send_message(user['user_id'], message.text or message.caption or "")
            success += 1
            await asyncio.sleep(0.05)  # Rate limit
        except:
            failed += 1

    await message.answer(
        f"✅ Xabar yuborildi:\n"
        f"• Yuborildi: {success} ta\n"
        f"• Xato: {failed} ta"
    )
    await state.clear()


@dp.callback_query(F.data == "admin_settings")
async def admin_settings_menu(callback: CallbackQuery):
    """Sozlamalar"""
    if not callback.from_user:
        return

    user_id = callback.from_user.id
    if user_id not in ADMIN_IDS and await get_setting(f"admin_{user_id}") != "1":
        await callback.answer("❌ Ruxsat yo'q.", show_alert=True)
        return

    api_key = await get_setting("ai_api_key") or "(sozlanmagan)"
    model = await get_setting("ai_model") or "gemini-2.0-flash"
    max_slides = await get_setting("max_slides_per_day") or "10"
    max_courses = await get_setting("max_courses_per_day") or "5"

    text = (
        f"⚙️ <b>Sozlamalar:</b>\n\n"
        f"🤖 AI Model: {model}\n"
        f"🔑 API Kalit: {api_key[:15]}...\n"
        f"📑 Kunlik slayd limiti: {max_slides}\n"
        f"📝 Kunlik kurs ishi limiti: {max_courses}\n\n"
        f"Sozlash uchun:\n"
        f"/set_api_key KALIT — API kalitini qo'shish\n"
        f"/set_model MODEL — Modelni tanlash\n"
        f"/set_limit_slides N — Slayd limiti\n"
        f"/set_limit_courses N — Kurs ishi limiti"
    )

    await callback.message.edit_text(text, reply_markup=await admin_inline_keyboard())
    await callback.answer()


@dp.message(Command("set_api_key"))
async def set_api_key_command(message: Message):
    """API kalitini sozlash"""
    if not message.from_user:
        return

    user_id = message.from_user.id
    if user_id not in ADMIN_IDS and await get_setting(f"admin_{user_id}") != "1":
        return

    parts = message.text.split(maxsplit=1)
    if len(parts) < 2:
        await message.answer("❌ Format: /set_api_key KALIT")
        return

    api_key = parts[1].strip()
    await set_setting("ai_api_key", api_key)
    init_ai(api_key, await get_setting("ai_model") or "gemini-2.0-flash")
    await message.answer("✅ API kaliti saqlandi!")


@dp.message(Command("set_model"))
async def set_model_command(message: Message):
    """AI modelini sozlash"""
    if not message.from_user:
        return

    user_id = message.from_user.id
    if user_id not in ADMIN_IDS and await get_setting(f"admin_{user_id}") != "1":
        return

    parts = message.text.split(maxsplit=1)
    if len(parts) < 2:
        await message.answer("❌ Format: /set_model gemini-2.0-flash")
        return

    model = parts[1].strip()
    await set_setting("ai_model", model)
    api_key = await get_setting("ai_api_key")
    if api_key:
        init_ai(api_key, model)
    await message.answer(f"✅ Model saqlandi: {model}")


@dp.message(Command("set_limit_slides"))
async def set_limit_slides(message: Message):
    """Slayd limitini sozlash"""
    if not message.from_user:
        return

    user_id = message.from_user.id
    if user_id not in ADMIN_IDS and await get_setting(f"admin_{user_id}") != "1":
        return

    parts = message.text.split(maxsplit=1)
    if len(parts) < 2 or not parts[1].isdigit():
        await message.answer("❌ Format: /set_limit_slides 20")
        return

    await set_setting("max_slides_per_day", parts[1])
    await message.answer(f"✅ Slayd limiti: {parts[1]}")


@dp.message(Command("set_limit_courses"))
async def set_limit_courses(message: Message):
    """Kurs ishi limitini sozlash"""
    if not message.from_user:
        return

    user_id = message.from_user.id
    if user_id not in ADMIN_IDS and await get_setting(f"admin_{user_id}") != "1":
        return

    parts = message.text.split(maxsplit=1)
    if len(parts) < 2 or not parts[1].isdigit():
        await message.answer("❌ Format: /set_limit_courses 10")
        return

    await set_setting("max_courses_per_day", parts[1])
    await message.answer(f"✅ Kurs ishi limiti: {parts[1]}")


@dp.message(Command("broadcast"))
async def broadcast_command(message: Message, state: FSMContext):
    """Broadcast komandasi"""
    if not message.from_user:
        return

    user_id = message.from_user.id
    if user_id not in ADMIN_IDS:
        return

    parts = message.text.split(maxsplit=1)
    if len(parts) < 2:
        await message.answer("❌ Format: /broadcast XABAR")
        return

    text = parts[1]
    users = await get_all_users()
    success = 0
    for user in users:
        try:
            await bot.send_message(user['user_id'], text)
            success += 1
            await asyncio.sleep(0.05)
        except:
            pass

    await message.answer(f"✅ {success}/{len(users)} ta foydalanuvchiga yuborildi.")


@dp.message(Command("cancel"))
async def cancel_command(message: Message, state: FSMContext):
    """Bekor qilish"""
    await state.clear()
    await message.answer("❌ Amal bekor qilindi.")


# ============ MENU ORQAGA QAYTISH ============
@dp.message(F.text == "⬅️ Orqaga")
async def go_back(message: Message):
    """Orqaga qaytish"""
    await message.answer("Asosiy menu:", reply_markup=main_keyboard())


# ============ ASOSIY ISHGA TUSHIRISH ============
async def on_startup():
    """Bot ishga tushganda"""
    await init_db()
    log.info("✅ Database initialized")

    # AI init
    api_key = GEMINI_API_KEY or await get_setting("ai_api_key")
    if api_key:
        model = await get_setting("ai_model") or "gemini-2.5-flash-lite"
        init_ai(api_key, model)
        log.info(f"✅ AI initialized with model: {model}")
    else:
        log.warning("⚠️ AI API key not set!")

    # Eski webhook bo'lsa o'chirish (polling to'g'ri ishlashi uchun)
    await bot.delete_webhook(drop_pending_updates=True)
    log.info("🚀 Bot started, polling boshlanadi...")


async def main():
    """Main funksiya"""
    await on_startup()
    await dp.start_polling(bot)


if __name__ == "__main__":
    if not TOKEN:
        print("❌ BOT_TOKEN environment variable not set!")
        print("Create .env file with: BOT_TOKEN=your_bot_token")
        sys.exit(1)

    asyncio.run(main())
