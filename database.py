"""
Ma'lumotlar bazasi - SQLite
Foydalanuvchilar, statistika va sozlamalar
"""
import aiosqlite
import json
from datetime import datetime
from pathlib import Path

DB_PATH = Path(__file__).parent / "bot_data.db"


async def init_db():
    """Bazani yaratish va jadvallarni tayyorlash"""
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("""
            CREATE TABLE IF NOT EXISTS users (
                user_id INTEGER PRIMARY KEY,
                username TEXT,
                first_name TEXT,
                last_name TEXT,
                joined_at TEXT DEFAULT (datetime('now')),
                last_active TEXT DEFAULT (datetime('now')),
                total_slides INTEGER DEFAULT 0,
                total_courses INTEGER DEFAULT 0,
                is_blocked INTEGER DEFAULT 0,
                is_admin INTEGER DEFAULT 0
            )
        """)

        await db.execute("""
            CREATE TABLE IF NOT EXISTS generation_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                type TEXT,  -- 'slide' yoki 'course'
                topic TEXT,
                details TEXT,
                created_at TEXT DEFAULT (datetime('now')),
                file_path TEXT,
                FOREIGN KEY (user_id) REFERENCES users(user_id)
            )
        """)

        await db.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)

        # Default sozlamalar
        default_settings = [
            ('ai_model', 'gemini-2.0-flash'),
            ('ai_api_key', ''),
            ('max_slides_per_day', '10'),
            ('max_courses_per_day', '5'),
            ('welcome_message', 'Assalomu alaykum! 👋\n\nMen professional <b>Slayd va Kurs ishi</b> yaratuvchi botman.\n\n📊 <b>Slayd yaratish</b> — mavzuni kiriting, AI sizga chiroyli PPTX prezentatsiya tayyorlab beradi.\n\n📝 <b>Kurs ishi</b> — mavzuni kiriting, professional DOCX hujjat tayyorlayman.\n\nBoshlash uchun pastdagi tugmalardan birini bosing 👇'),
            ('broadcast_enabled', '1'),
        ]
        await db.executemany(
            "INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)",
            default_settings
        )
        await db.commit()


async def get_setting(key: str) -> str | None:
    """Sozlamani olish"""
    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute("SELECT value FROM settings WHERE key = ?", (key,))
        row = await cursor.fetchone()
        return row[0] if row else None


async def set_setting(key: str, value: str):
    """Sozlamani yangilash"""
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)",
            (key, value)
        )
        await db.commit()


async def add_user(user_id: int, username: str = None, first_name: str = None, last_name: str = None):
    """Yangi foydalanuvchi qo'shish yoki yangilash"""
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("""
            INSERT INTO users (user_id, username, first_name, last_name, last_active)
            VALUES (?, ?, ?, ?, datetime('now'))
            ON CONFLICT(user_id) DO UPDATE SET
                username = COALESCE(?, username),
                first_name = COALESCE(?, first_name),
                last_name = COALESCE(?, last_name),
                last_active = datetime('now')
        """, (user_id, username, first_name, last_name, username, first_name, last_name))
        await db.commit()


async def get_user(user_id: int) -> dict | None:
    """Foydalanuvchi ma'lumotlarini olish"""
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        cursor = await db.execute("SELECT * FROM users WHERE user_id = ?", (user_id,))
        row = await cursor.fetchone()
        return dict(row) if row else None


async def get_all_users() -> list[dict]:
    """Barcha foydalanuvchilarni olish"""
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        cursor = await db.execute("SELECT * FROM users ORDER BY last_active DESC")
        rows = await cursor.fetchall()
        return [dict(row) for row in rows]


async def get_users_count() -> int:
    """Foydalanuvchilar soni"""
    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute("SELECT COUNT(*) FROM users")
        row = await cursor.fetchone()
        return row[0] if row else 0


async def increment_user_stats(user_id: int, stat_type: str):
    """Statistikani oshirish (slides yoki courses)"""
    column = "total_slides" if stat_type == "slide" else "total_courses"
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(f"UPDATE users SET {column} = {column} + 1 WHERE user_id = ?", (user_id,))
        await db.commit()


async def add_generation_record(user_id: int, gen_type: str, topic: str, details: str, file_path: str = ""):
    """Generatsiya tarixini qo'shish"""
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            "INSERT INTO generation_history (user_id, type, topic, details, file_path) VALUES (?, ?, ?, ?, ?)",
            (user_id, gen_type, topic, json.dumps(details, ensure_ascii=False), file_path)
        )
        await db.commit()


async def get_generation_stats(days: int = 7) -> dict:
    """Statistika"""
    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute(
            "SELECT type, COUNT(*) FROM generation_history WHERE created_at >= datetime('now', ?) GROUP BY type",
            (f'-{days} days',)
        )
        rows = await cursor.fetchall()
        stats = {"slide": 0, "course": 0}
        for row in rows:
            stats[row[0]] = row[1]

        cursor = await db.execute("SELECT COUNT(*) FROM generation_history")
        total = await cursor.fetchone()
        stats["total"] = total[0] if total else 0

        return stats


async def get_user_today_count(user_id: int, gen_type: str) -> int:
    """Bugungi generatsiyalar soni"""
    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute(
            """SELECT COUNT(*) FROM generation_history
               WHERE user_id = ? AND type = ? AND date(created_at) = date('now')""",
            (user_id, gen_type)
        )
        row = await cursor.fetchone()
        return row[0] if row else 0


async def block_user(user_id: int, blocked: bool = True):
    """Foydalanuvchini bloklash/ochish"""
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("UPDATE users SET is_blocked = ? WHERE user_id = ?", (1 if blocked else 0, user_id))
        await db.commit()


async def set_admin(user_id: int, is_admin: bool = True):
    """Admin qilish"""
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("UPDATE users SET is_admin = ? WHERE user_id = ?", (1 if is_admin else 0, user_id))
        await db.commit()


async def is_user_blocked(user_id: int) -> bool:
    """Bloklanganmi?"""
    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute("SELECT is_blocked FROM users WHERE user_id = ?", (user_id,))
        row = await cursor.fetchone()
        return bool(row and row[0])
