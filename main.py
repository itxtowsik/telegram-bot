import os, re, json, asyncio, logging, sqlite3, tempfile, shutil, html, secrets
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional, Any
from zoneinfo import ZoneInfo

import aiosqlite
from openpyxl import Workbook, load_workbook
from aiogram import Bot, Dispatcher, F
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.exceptions import TelegramBadRequest, TelegramForbiddenError, TelegramNetworkError
from aiogram.filters import Command, StateFilter
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import (
    Message, CallbackQuery, InlineKeyboardMarkup, InlineKeyboardButton,
    ReplyKeyboardMarkup, KeyboardButton, KeyboardButtonRequestChat,
    ReplyKeyboardRemove, BufferedInputFile
)

try:
    from aiohttp import web
except ImportError:
    web = None

# ============================================================
# 1. IMPORTS / CONFIG — SINGLE FILE SETUP
# ============================================================
# এই CONFIG-গুলো সরাসরি main.py-তে রাখা হয়েছে। আলাদা .env ফাইল প্রয়োজন নেই।
# Hosting-এর আগে শুধু BOT_TOKEN এবং OWNER_ID নিজের তথ্য দিয়ে বসিয়ে দিন।

BOT_TOKEN = "8735275840:AAGd-omfCoJz-ca2DHL0miAFPYmSMLHLQn0"
OWNER_ID = 8991905623

# Web Admin Panel login
WEB_HOST = "0.0.0.0"
WEB_PORT = int(os.environ.get("PORT", "8080"))
WEB_ADMIN_USER = "admin"
WEB_ADMIN_PASS = "CHANGE_THIS_WEB_PASSWORD"
WEB_ADMIN_SECRET = "CHANGE_THIS_TO_A_LONG_RANDOM_SECRET_9f7c2a1d6e4b8a3c"

DB_PATH = "id_submit_zone.db"
TIMEZONE = ZoneInfo("Asia/Dhaka")
CURRENCY = "৳"
DEFAULT_MIN_WITHDRAW = Decimal("20")
MAX_FILE_BYTES = 20 * 1024 * 1024
SENSITIVE_WORDS = {
    "password", "passwd", "passcode", "pwd", "cookie", "cookies",
    "session", "sessionid", "session_token", "sessiontoken", "token",
    "access_token", "accesstoken", "refresh_token", "refreshtoken",
    "recovery", "recovery_code", "recoverycode", "login", "credential",
    "credentials", "secret", "authorization", "auth_token", "authtoken",
    "api_key", "apikey", "private_key", "privatekey", "pin", "otp",
    "2fa", "two_factor", "verification_code", "security_code"
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s"
)
log = logging.getLogger("id_submit_zone")

if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN is missing. Open CONFIG section in main.py and set BOT_TOKEN.")
if OWNER_ID <= 0:
    raise RuntimeError("OWNER_ID is missing or invalid. Open CONFIG section in main.py and set OWNER_ID.")

# ============================================================
# 2. CONSTANTS
# ============================================================
MAIN_SUBMIT = "📤 SUBMIT FILE"
MAIN_PROFILE = "👤 MY PROFILE"
MAIN_WALLET = "💰 WALLET"
MAIN_RULES = "📋 RULES"
MAIN_HISTORY = "📜 HISTORY"
MAIN_SUPPORT = "💬 SUPPORT / HELP"
MAIN_ADMIN = "👑 ADMIN PANEL"
TELEGRAM_ADMIN_PANEL_ENABLED = True

class SubmitState(StatesGroup):
    waiting_category = State()
    waiting_file = State()
    awaiting_confirmation = State()

class SupportState(StatesGroup):
    active = State()

class UserInputState(StatesGroup):
    payment_method = State()
    payment_number = State()
    custom_withdraw = State()

class AdminState(StatesGroup):
    add_category_name = State()
    add_category_rate = State()
    rename_category = State()
    category_emoji = State()
    category_order = State()
    change_rate = State()
    add_format = State()
    edit_format = State()
    reject_reason = State()
    adjust_quantity = State()
    add_balance = State()
    deduct_balance = State()
    search_user = State()
    search_submission = State()
    search_payment = State()
    support_reply = State()
    broadcast_text = State()
    add_admin = State()
    remove_admin = State()
    edit_rules = State()
    set_min_withdraw = State()
    set_submission_schedule = State()
    payment_txid = State()
    payment_screenshot = State()

# ============================================================
# 3. GENERAL HELPERS
# ============================================================
def now_str() -> str:
    return datetime.now(TIMEZONE).strftime("%Y-%m-%d %H:%M:%S")

def money(value: Any) -> Decimal:
    try:
        return Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError, TypeError):
        return Decimal("0.00")

def fmt_money(value: Any) -> str:
    return f"{money(value):,.2f}"

def mask_number(number: str) -> str:
    number = number.strip()
    if len(number) <= 7:
        return "*" * max(0, len(number) - 2) + number[-2:]
    return number[:2] + "X" * max(1, len(number) - 5) + number[-3:]

def username_text(user) -> str:
    return f"@{user.username}" if getattr(user, "username", None) else "(no username)"

def display_name(user) -> str:
    name = " ".join(x for x in [getattr(user, "first_name", ""), getattr(user, "last_name", "")] if x)
    return name[:100] or "Unknown"

def safe_filename(name: str) -> str:
    name = Path(name).name
    return re.sub(r"[^A-Za-z0-9._-]", "_", name)[:150] or "file.xlsx"

def is_sensitive_column(name: str) -> bool:
    s = re.sub(r"[^a-z0-9]+", "_", name.lower()).strip("_")
    if not s:
        return False
    tokens = set(s.split("_"))
    if tokens & SENSITIVE_WORDS:
        return True
    compact = s.replace("_", "")
    return any(w.replace("_", "") in compact for w in SENSITIVE_WORDS if len(w) >= 5)

def parse_columns(text: str) -> list[str]:
    cols = [x.strip() for x in text.split("|") if x.strip()]
    return cols

def valid_format_columns(cols: list[str]) -> tuple[bool, str]:
    if not cols:
        return False, "কমপক্ষে ১টি column প্রয়োজন।"
    if len(cols) > 30:
        return False, "সর্বোচ্চ ৩০টি column রাখা যাবে।"
    lowered = [c.lower() for c in cols]
    if len(set(lowered)) != len(lowered):
        return False, "Duplicate column name রাখা যাবে না।"
    for c in cols:
        if is_sensitive_column(c):
            return False, f'নিরাপত্তার কারণে column "{c}" গ্রহণযোগ্য নয়।'
    if not any(c.lower() == "uid" for c in cols):
        return False, 'Required column "UID" থাকতে হবে।'
    return True, ""

def get_message_user(message: Message):
    return message.from_user

# ============================================================
# 4. DATABASE
# ============================================================
DB: Optional[aiosqlite.Connection] = None
DB_LOCK = asyncio.Lock()

SCHEMA = """
PRAGMA foreign_keys = ON;
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    telegram_id INTEGER UNIQUE NOT NULL,
    username TEXT,
    first_name TEXT,
    balance REAL NOT NULL DEFAULT 0,
    pending_balance REAL NOT NULL DEFAULT 0,
    total_earned REAL NOT NULL DEFAULT 0,
    total_paid REAL NOT NULL DEFAULT 0,
    banned INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS admins (
    user_id INTEGER PRIMARY KEY,
    role TEXT NOT NULL DEFAULT 'admin',
    added_by INTEGER,
    created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS categories (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL,
    rate REAL NOT NULL DEFAULT 0,
    enabled INTEGER NOT NULL DEFAULT 1,
    emoji TEXT NOT NULL DEFAULT '📦',
    button_style TEXT NOT NULL DEFAULT 'primary',
    sort_order INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS formats (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    category_id INTEGER UNIQUE NOT NULL,
    columns_json TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    FOREIGN KEY(category_id) REFERENCES categories(id) ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS submissions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    public_id TEXT UNIQUE,
    user_id INTEGER NOT NULL,
    category_id INTEGER NOT NULL,
    category_name TEXT NOT NULL,
    file_id TEXT NOT NULL,
    file_name TEXT NOT NULL,
    records_count INTEGER NOT NULL DEFAULT 0,
    valid_count INTEGER NOT NULL DEFAULT 0,
    invalid_count INTEGER NOT NULL DEFAULT 0,
    duplicate_count INTEGER NOT NULL DEFAULT 0,
    locked_rate REAL NOT NULL DEFAULT 0,
    estimated_amount REAL NOT NULL DEFAULT 0,
    approved_quantity INTEGER NOT NULL DEFAULT 0,
    final_amount REAL NOT NULL DEFAULT 0,
    validation_status TEXT NOT NULL DEFAULT 'valid',
    review_status TEXT NOT NULL DEFAULT 'pending',
    payment_status TEXT NOT NULL DEFAULT 'unpaid',
    rejection_reason TEXT,
    created_at TEXT NOT NULL,
    reviewed_at TEXT,
    credited_amount REAL NOT NULL DEFAULT 0,
    credited_by INTEGER,
    credited_at TEXT,
    FOREIGN KEY(user_id) REFERENCES users(telegram_id),
    FOREIGN KEY(category_id) REFERENCES categories(id)
);
CREATE TABLE IF NOT EXISTS submission_records (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    submission_id INTEGER NOT NULL,
    row_number INTEGER NOT NULL,
    uid TEXT,
    status TEXT,
    notes TEXT,
    data_json TEXT,
    valid INTEGER NOT NULL DEFAULT 1,
    duplicate INTEGER NOT NULL DEFAULT 0,
    FOREIGN KEY(submission_id) REFERENCES submissions(id) ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS withdrawals (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    public_id TEXT UNIQUE,
    user_id INTEGER NOT NULL,
    amount REAL NOT NULL,
    method TEXT NOT NULL,
    number TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'pending',
    proof_reference TEXT,
    admin_id INTEGER,
    related_submission_id INTEGER,
    created_at TEXT NOT NULL,
    processed_at TEXT,
    FOREIGN KEY(user_id) REFERENCES users(telegram_id)
);
CREATE TABLE IF NOT EXISTS transactions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    public_id TEXT UNIQUE,
    user_id INTEGER NOT NULL,
    txn_type TEXT NOT NULL,
    amount REAL NOT NULL,
    balance_after REAL NOT NULL,
    reason TEXT,
    source_id TEXT,
    admin_id INTEGER,
    created_at TEXT NOT NULL,
    FOREIGN KEY(user_id) REFERENCES users(telegram_id)
);
CREATE TABLE IF NOT EXISTS payment_methods (
    user_id INTEGER PRIMARY KEY,
    method TEXT,
    number TEXT,
    updated_at TEXT NOT NULL,
    FOREIGN KEY(user_id) REFERENCES users(telegram_id)
);
CREATE TABLE IF NOT EXISTS support_tickets (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    public_id TEXT UNIQUE,
    user_id INTEGER NOT NULL,
    topic TEXT,
    status TEXT NOT NULL DEFAULT 'closed',
    created_at TEXT NOT NULL,
    closed_at TEXT,
    FOREIGN KEY(user_id) REFERENCES users(telegram_id)
);
CREATE TABLE IF NOT EXISTS support_messages (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ticket_id INTEGER NOT NULL,
    sender_id INTEGER NOT NULL,
    sender_role TEXT NOT NULL,
    message_type TEXT NOT NULL DEFAULT 'text',
    body TEXT,
    created_at TEXT NOT NULL,
    FOREIGN KEY(ticket_id) REFERENCES support_tickets(id) ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS broadcasts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    admin_id INTEGER NOT NULL,
    body TEXT NOT NULL,
    recipients INTEGER NOT NULL DEFAULT 0,
    sent INTEGER NOT NULL DEFAULT 0,
    failed INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS admin_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    admin_id INTEGER NOT NULL,
    action TEXT NOT NULL,
    target TEXT,
    amount REAL,
    reference_id TEXT,
    created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_submissions_user ON submissions(user_id);
CREATE INDEX IF NOT EXISTS idx_submissions_status ON submissions(review_status);
CREATE INDEX IF NOT EXISTS idx_withdrawals_status ON withdrawals(status);
CREATE INDEX IF NOT EXISTS idx_transactions_user ON transactions(user_id);
CREATE INDEX IF NOT EXISTS idx_support_messages_ticket ON support_messages(ticket_id);
"""

async def db_init():
    global DB
    DB = await aiosqlite.connect(DB_PATH)
    DB.row_factory = aiosqlite.Row
    await DB.executescript(SCHEMA)
    # Safe migrations for existing installations.
    existing_cols = {r["name"] for r in await fetchall("PRAGMA table_info(submissions)")}
    wd_cols = {r["name"] for r in await fetchall("PRAGMA table_info(withdrawals)")}
    for col, definition in (("payment_screenshot_file_id", "TEXT"), ("payment_screenshot_type", "TEXT"), ("payment_proof_created_at", "TEXT"), ("payment_proof_channel_sent", "INTEGER NOT NULL DEFAULT 0"), ("payment_proof_user_sent", "INTEGER NOT NULL DEFAULT 0")):
        if col not in wd_cols:
            await DB.execute(f"ALTER TABLE withdrawals ADD COLUMN {col} {definition}")
    for col, definition in (("credited_amount", "REAL NOT NULL DEFAULT 0"), ("credited_by", "INTEGER"), ("credited_at", "TEXT")):
        if col not in existing_cols:
            await DB.execute(f"ALTER TABLE submissions ADD COLUMN {col} {definition}")
    cat_cols = {r["name"] for r in await fetchall("PRAGMA table_info(categories)")}
    for col, definition in (("emoji", "TEXT NOT NULL DEFAULT '📦'"), ("button_style", "TEXT NOT NULL DEFAULT 'primary'"), ("sort_order", "INTEGER NOT NULL DEFAULT 0")):
        if col not in cat_cols:
            await DB.execute(f"ALTER TABLE categories ADD COLUMN {col} {definition}")
    await DB.execute("UPDATE categories SET sort_order=id WHERE sort_order=0")
    await DB.commit()
    await set_setting_if_missing("min_withdraw", str(DEFAULT_MIN_WITHDRAW))
    await set_setting_if_missing("rules", "• Accepted file type: .xlsx\n• Required format: UID | Status | Notes\n• File review is required before approval.\n• Approved quantity × locked rate = approved amount.\n• Withdrawal minimum and payment rules are controlled by Admin.")
    await set_setting_if_missing("maintenance", "0")
    await set_setting_if_missing("submission_window_enabled", "1")
    await set_setting_if_missing("submission_start", "00:00")
    await set_setting_if_missing("submission_end", "20:00")
    await set_setting_if_missing("submission_alerts_enabled", "1")
    await set_setting_if_missing("bot_name", "Apex Valorem")
    await set_setting_if_missing("payment_proof_chat_id", "")
    await set_setting_if_missing("payment_proof_chat_title", "")
    await set_setting_if_missing("payment_proof_chat_username", "")
    await set_setting_if_missing("force_join_chat_id", "")
    await set_setting_if_missing("force_join_chat_title", "")
    await set_setting_if_missing("force_join_chat_username", "")
    await set_setting_if_missing("force_join_invite_link", "")
    await ensure_admin(OWNER_ID, "owner", OWNER_ID)
    defaults = [("Facebook", 4.80, "📘", "primary"), ("Instagram", 5.20, "📸", "success"), ("Gmail", 4.50, "✉️", "danger")]
    for idx, (name, rate, emoji, style) in enumerate(defaults, 1):
        cur = await DB.execute("SELECT id FROM categories WHERE name=?", (name,))
        row = await cur.fetchone()
        if not row:
            cur = await DB.execute("INSERT INTO categories(name,rate,enabled,emoji,button_style,sort_order,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?)", (name, rate, 1, emoji, style, idx, now_str(), now_str()))
            cid = cur.lastrowid
            await DB.execute("INSERT OR IGNORE INTO formats(category_id,columns_json,updated_at) VALUES(?,?,?)", (cid, json.dumps(["UID","Status","Notes"]), now_str()))
    await DB.commit()

async def fetchone(sql: str, params=()):
    async with DB.execute(sql, params) as cur:
        return await cur.fetchone()

async def fetchall(sql: str, params=()):
    async with DB.execute(sql, params) as cur:
        return await cur.fetchall()

async def execute(sql: str, params=(), commit=True):
    async with DB_LOCK:
        cur = await DB.execute(sql, params)
        if commit:
            await DB.commit()
        return cur

async def set_setting_if_missing(key: str, value: str):
    row = await fetchone("SELECT key FROM settings WHERE key=?", (key,))
    if not row:
        await DB.execute("INSERT INTO settings(key,value,updated_at) VALUES(?,?,?)", (key, value, now_str()))
        await DB.commit()

async def get_setting(key: str, default: str = "") -> str:
    row = await fetchone("SELECT value FROM settings WHERE key=?", (key,))
    return row["value"] if row else default

async def set_setting(key: str, value: str):
    await execute("INSERT INTO settings(key,value,updated_at) VALUES(?,?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value,updated_at=excluded.updated_at", (key, value, now_str()))

async def ensure_user(user):
    row = await fetchone("SELECT * FROM users WHERE telegram_id=?", (user.id,))
    if row:
        await execute("UPDATE users SET username=?,first_name=?,updated_at=? WHERE telegram_id=?", (user.username, display_name(user), now_str(), user.id))
        return row
    await execute("INSERT INTO users(telegram_id,username,first_name,created_at,updated_at) VALUES(?,?,?,?,?)", (user.id, user.username, display_name(user), now_str(), now_str()))
    return await fetchone("SELECT * FROM users WHERE telegram_id=?", (user.id,))

async def ensure_admin(user_id: int, role="admin", added_by=None):
    await DB.execute("INSERT OR IGNORE INTO admins(user_id,role,added_by,created_at) VALUES(?,?,?,?)", (user_id, role, added_by, now_str()))
    await DB.commit()

async def admin_role(user_id: int) -> Optional[str]:
    if user_id == OWNER_ID:
        return "owner"
    row = await fetchone("SELECT role FROM admins WHERE user_id=?", (user_id,))
    return row["role"] if row else None

async def is_admin(user_id: int) -> bool:
    return await admin_role(user_id) is not None

async def is_owner(user_id: int) -> bool:
    return user_id == OWNER_ID

async def log_admin(admin_id, action, target=None, amount=None, reference_id=None):
    await execute("INSERT INTO admin_logs(admin_id,action,target,amount,reference_id,created_at) VALUES(?,?,?,?,?,?)", (admin_id, action, target, float(amount) if amount is not None else None, reference_id, now_str()))

async def next_public_id(table: str, prefix: str) -> str:
    # IDs are based on the database row count and protected by UNIQUE constraints.
    row = await fetchone(f"SELECT COALESCE(MAX(id),0)+1 AS n FROM {table}")
    return f"{prefix}-{int(row['n']):06d}"

# ============================================================
# 5. KEYBOARDS
# ============================================================
def kb(rows):
    return InlineKeyboardMarkup(inline_keyboard=rows)

def btn(text, data, style=None):
    # Native Telegram button styles: primary=blue, success=green, danger=red.
    # Back/Cancel/Preview-like navigation is always red. Ordinary actions are
    # intentionally mixed between blue and green so the UI does not become one-color.
    t = (text or "").lower()
    danger_words = ("❌", "delete", "reject", "disable", "ban", "deduct", "cancel", "close", "remove", "◀️ back", "preview")
    success_words = ("✅", "approve", "confirm", "paid", "add balance", "enable", "send", "save", "submit", "on/off")
    if any(x in t for x in danger_words):
        style = "danger"
    elif style is None and any(x in t for x in success_words):
        style = "success"
    elif style is None:
        # Stable blue/green split based on the callback data, not randomness.
        style = "success" if sum(ord(ch) for ch in str(data)) % 2 else "primary"
    elif style == "primary" and not (str(data).startswith("u:cat:") or str(data).startswith("a:style:")):
        # Keep ordinary explicitly-primary actions mixed; category design selections
        # remain exactly as chosen by the admin.
        style = "success" if sum(ord(ch) for ch in str(data)) % 2 else "primary"
    kwargs = {"text": text, "callback_data": data}
    if style in {"primary", "success", "danger"}:
        kwargs["style"] = style
    return InlineKeyboardButton(**kwargs)

def rbtn(text, style=None):
    # Styled ReplyKeyboard button; supported by aiogram 3.31 / Bot API 10.3+.
    kwargs = {"text": text}
    if style in {"primary", "success", "danger"}:
        kwargs["style"] = style
    return KeyboardButton(**kwargs)

def main_kb(is_admin_user=False):
    """Premium Telegram bottom Reply Keyboard. Admin sees the admin entry only."""
    rows = [
        [rbtn(MAIN_SUBMIT, "primary")],
        [rbtn(MAIN_PROFILE, "primary"), rbtn(MAIN_WALLET, "success")],
        [rbtn(MAIN_RULES, "primary"), rbtn(MAIN_HISTORY, "primary")],
        [rbtn(MAIN_SUPPORT, "primary")],
    ]
    if is_admin_user and TELEGRAM_ADMIN_PANEL_ENABLED:
        rows.append([rbtn(MAIN_ADMIN, "danger")])
    return ReplyKeyboardMarkup(
        keyboard=rows,
        resize_keyboard=True,
        is_persistent=True,
        one_time_keyboard=False,
        input_field_placeholder="একটি অপশন নির্বাচন করুন…",
    )

def back_kb(data="u:home"):
    return kb([[btn("◀️ Back", data)]])

def confirm_kb(ok_data, cancel_data="u:home", ok="✅ Confirm", cancel="❌ Cancel"):
    return kb([[btn(ok, ok_data), btn(cancel, cancel_data)]])

# ============================================================
# 6. UI / NAVIGATION
# ============================================================
async def clear_transient_sessions(state: FSMContext):
    await state.clear()

async def leave_support_if_active(user_id: int, state: FSMContext):
    data = await state.get_data()
    if data.get("support_active"):
        await state.update_data(support_active=False)
        await state.set_state(None)

async def send_or_edit(target, text, reply_markup=None):
    try:
        if isinstance(target, CallbackQuery):
            try:
                await target.message.edit_text(text, reply_markup=reply_markup)
            except TelegramBadRequest as e:
                if "message is not modified" not in str(e).lower():
                    await target.message.answer(text, reply_markup=reply_markup)
            await target.answer()
        else:
            await target.answer(text, reply_markup=reply_markup)
    except TelegramBadRequest as e:
        log.warning("UI error: %s", e)

async def user_home_message(bot: Bot, chat_id: int, user_id: int):
    role = await admin_role(user_id)
    text = "🏠 প্রধান মেনু\n\nপ্রয়োজনীয় সেবা নির্বাচন করুন। 🤝"
    await bot.send_message(chat_id, text, reply_markup=main_kb(role is not None))

def parse_hhmm(value: str) -> Optional[tuple[int, int]]:
    try:
        h, m = value.strip().split(":", 1)
        h, m = int(h), int(m)
        if 0 <= h <= 23 and 0 <= m <= 59:
            return h, m
    except Exception:
        pass
    return None

def parse_12h_time(value: str) -> Optional[tuple[int, int]]:
    """Parse admin-entered 12-hour time, e.g. 1, 1:30, 12:05."""
    raw = value.strip().replace(".", ":")
    if not re.fullmatch(r"(?:[1-9]|1[0-2])(?::[0-5]\d)?", raw):
        return None
    parts = raw.split(":", 1)
    hour = int(parts[0])
    minute = int(parts[1]) if len(parts) > 1 else 0
    return hour, minute

def convert_12h_to_24(hour: int, minute: int, meridiem: str) -> str:
    meridiem = meridiem.upper()
    if meridiem not in {"AM", "PM"}:
        raise ValueError("Invalid AM/PM")
    if hour == 12:
        h24 = 0 if meridiem == "AM" else 12
    else:
        h24 = hour + (12 if meridiem == "PM" else 0)
    return f"{h24:02d}:{minute:02d}"

def format_12h(value: str) -> str:
    parsed = parse_hhmm(value)
    if not parsed:
        return value
    hour, minute = parsed
    meridiem = "AM" if hour < 12 else "PM"
    h12 = hour % 12 or 12
    return f"{h12}:{minute:02d} {meridiem}"

def hhmm_minutes(value: str) -> int:
    parsed = parse_hhmm(value)
    return parsed[0] * 60 + parsed[1] if parsed else 0

async def submission_window_status() -> tuple[bool, str]:
    enabled = await get_setting("submission_window_enabled", "1") == "1"
    start = await get_setting("submission_start", "00:00")
    end = await get_setting("submission_end", "20:00")
    if not enabled:
        return False, "⛔ বর্তমানে Admin নতুন submission নেওয়া বন্ধ রেখেছেন।"
    sp, ep = parse_hhmm(start), parse_hhmm(end)
    if not sp or not ep:
        return False, "⚠️ Submission schedule সঠিকভাবে সেট করা নেই। Admin-কে জানান।"
    now = datetime.now(TIMEZONE)
    cur = now.hour * 60 + now.minute
    sm, em = sp[0] * 60 + sp[1], ep[0] * 60 + ep[1]
    if sm == em:
        return True, ""
    if sm < em:
        opened = sm <= cur < em
    else:
        opened = cur >= sm or cur < em
    if not opened:
        return False, (f"⏰ <b>Submission Time এখন বন্ধ</b>\n\n📅 আজকের সময়: <b>{format_12h(start)} – {format_12h(end)}</b> (বাংলাদেশ সময়)\n"
                       f"\n🕐 বর্তমান সময়: <b>{format_12h(now.strftime('%H:%M'))}</b>\n\n"
                       "নির্ধারিত সময়ের মধ্যে আবার Submit করুন। 🤝")
    return True, ""

async def maintenance_on() -> bool:
    return (await get_setting("maintenance", "0")) == "1"

async def get_force_join_info() -> dict:
    chat_id = (await get_setting("force_join_chat_id", "")).strip()
    if not chat_id:
        return {}
    return {
        "chat_id": chat_id,
        "title": await get_setting("force_join_chat_title", "Force Join Channel"),
        "username": await get_setting("force_join_chat_username", ""),
        "invite_link": await get_setting("force_join_invite_link", ""),
    }

async def force_join_ok(bot: Bot, user_id: int) -> tuple[bool, dict]:
    if await is_admin(user_id):
        return True, {}
    info = await get_force_join_info()
    if not info:
        return True, {}
    try:
        member = await bot.get_chat_member(int(info["chat_id"]), user_id)
        status = getattr(member, "status", "")
        if status in {"creator", "administrator", "member"}:
            return True, info
        if status == "restricted" and bool(getattr(member, "is_member", False)):
            return True, info
        return False, info
    except Exception as exc:
        log.warning("Force join verification failed: %s", exc)
        # If Telegram cannot verify membership, keep the user blocked rather than
        # silently bypassing a configured Force Join requirement.
        return False, info

async def force_join_prompt(bot: Bot, chat_id: int, info: dict):
    title = html.escape(info.get("title") or "Required Channel")
    link = info.get("invite_link") or (f"https://t.me/{info['username']}" if info.get("username") else "")
    rows = []
    if link:
        rows.append([InlineKeyboardButton(text="📢 JOIN CHANNEL", url=link, style="primary")])
    rows.append([btn("🔄 I JOINED — CHECK", "u:forcecheck", "success")])
    await bot.send_message(
        chat_id,
        f"🔐 <b>CHANNEL JOIN REQUIRED</b>\n\n"
        f"এই Bot ব্যবহার করার আগে নিচের channel-এ join করতে হবে।\n\n"
        f"📢 Channel: <b>{title}</b>\n\n"
        "✅ Join করার পর <b>I JOINED — CHECK</b> button চাপুন।",
        reply_markup=kb(rows)
    )

async def guard_force_join_message(message: Message) -> bool:
    ok, info = await force_join_ok(message.bot, message.from_user.id)
    if ok:
        return True
    await force_join_prompt(message.bot, message.chat.id, info)
    return False

async def guard_user(message: Message) -> bool:
    await ensure_user(message.from_user)
    row = await fetchone("SELECT banned FROM users WHERE telegram_id=?", (message.from_user.id,))
    if row and row["banned"]:
        await message.answer("🚫 আপনার অ্যাকাউন্ট বর্তমানে সীমিত করা হয়েছে।")
        return False
    if not await is_admin(message.from_user.id):
        if not await guard_force_join_message(message):
            return False
    if await maintenance_on() and not await is_admin(message.from_user.id):
        await message.answer("🔧 Bot maintenance mode-এ আছে। অনুগ্রহ করে কিছুক্ষণ পরে চেষ্টা করুন।")
        return False
    return True

# ============================================================
# 7. VALIDATION / TEMPLATE
# ============================================================
async def get_category(category_id: int):
    return await fetchone("SELECT * FROM categories WHERE id=?", (category_id,))

async def get_format(category_id: int):
    row = await fetchone("SELECT columns_json FROM formats WHERE category_id=?", (category_id,))
    if not row:
        return ["UID", "Status", "Notes"]
    try:
        return json.loads(row["columns_json"])
    except Exception:
        return ["UID", "Status", "Notes"]

def submission_notice_text() -> str:
    today = datetime.now(TIMEZONE).day
    return (
        "⚠️ <b>Submission Notice</b>\n\n"
        "📂 শুধু <b>Excel (.xlsx) File</b> গ্রহণ করা হবে।\n"
        "❌ অন্য কোনো File গ্রহণ করা হবে না।\n\n"
        "🔐 <b>Password-এর শেষে অবশ্যই Submit করার দিনের তারিখ থাকতে হবে।</b>\n"
        f"📅 আজ: <b>{today}</b> → Password-এর শেষে <b>{today}</b> থাকতে হবে।\n\n"
        "❌ তারিখ না থাকলে File/Account গ্রহণ করা হবে না।\n\n"
        "📎 <b>Submit করার আগে File ও Password চেক করুন।</b>"
    )


async def create_template(category_id: int) -> bytes:
    cols = await get_format(category_id)
    wb = Workbook()
    ws = wb.active
    ws.title = "Submission"
    ws.append(cols)
    # Safe example row is intentionally omitted.
    import io
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()

async def analyze_xlsx(path: str, required_cols: list[str]) -> dict:
    result = {"ok": False, "reason": "", "total": 0, "valid": 0, "invalid": 0, "duplicate": 0, "records": []}
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        if not wb.sheetnames:
            result["reason"] = "Workbook-এ কোনো sheet নেই।"
            return result
        ws = wb[wb.sheetnames[0]]
        rows = ws.iter_rows(values_only=True)
        try:
            header_raw = next(rows)
        except StopIteration:
            result["reason"] = "Excel file খালি।"
            return result
        headers = [str(x).strip() if x is not None else "" for x in header_raw]
        if not headers or all(not h for h in headers):
            result["reason"] = "Header পাওয়া যায়নি।"
            return result
        for h in headers:
            if is_sensitive_column(h):
                result["reason"] = f'নিরাপত্তার কারণে column "{h}" গ্রহণযোগ্য নয়।'
                return result
        normalized = [h.lower() for h in headers]
        req_norm = [x.lower() for x in required_cols]
        for req in req_norm:
            if req not in normalized:
                result["reason"] = f'Required column "{required_cols[req_norm.index(req)]}" পাওয়া যায়নি।'
                return result
        idx = {h.lower(): i for i, h in enumerate(headers)}
        uid_i = idx.get("uid")
        status_i = idx.get("status")
        notes_i = idx.get("notes")
        seen = set()
        row_no = 1
        for raw in rows:
            row_no += 1
            if raw is None or all(v is None or str(v).strip() == "" for v in raw):
                continue
            result["total"] += 1
            data = {}
            for i, h in enumerate(headers):
                val = raw[i] if i < len(raw) else None
                data[h] = "" if val is None else str(val).strip()
            uid = data.get(headers[uid_i], "") if uid_i is not None else ""
            status = data.get(headers[status_i], "") if status_i is not None else ""
            notes = data.get(headers[notes_i], "") if notes_i is not None else ""
            invalid = not bool(uid)
            duplicate = False
            key = uid.casefold() if uid else f"__empty_{row_no}"
            if uid and key in seen:
                duplicate = True
            if uid:
                seen.add(key)
            if duplicate:
                result["duplicate"] += 1
            if invalid:
                result["invalid"] += 1
            else:
                result["valid"] += 1
            result["records"].append({"row_number": row_no, "uid": uid, "status": status, "notes": notes, "data": data, "valid": not invalid, "duplicate": duplicate})
        if result["total"] == 0:
            result["reason"] = "Excel file-এ কোনো data row নেই।"
            return result
        if result["invalid"] > 0:
            result["reason"] = f"Invalid records: {result['invalid']}"
            return result
        if result["duplicate"] > 0:
            result["reason"] = f"Duplicate records: {result['duplicate']}"
            return result
        result["ok"] = True
        return result
    except Exception as e:
        log.exception("xlsx validation failed")
        result["reason"] = "Excel file readable নয় বা corrupt।"
        return result

# ============================================================
# 8. USER START / MAIN
# ============================================================
router = None

def build_router():
    from aiogram import Router
    return Router(name="main")

router = build_router()

@router.message(Command("start"))
async def cmd_start(message: Message, state: FSMContext):
    if not await guard_user(message):
        return
    await clear_transient_sessions(state)
    role = await admin_role(message.from_user.id)
    await message.answer(
        "🇧🇩 স্বাগতম!\n\nApex Valorem-এ আপনাকে স্বাগতম। প্রয়োজনীয় সেবা নির্বাচন করুন। 🤝",
        reply_markup=main_kb(role is not None)
    )

@router.callback_query(F.data == "u:forcecheck")
async def force_join_check(call: CallbackQuery):
    if await is_admin(call.from_user.id):
        await call.answer("Admin-এর জন্য Force Join প্রযোজ্য নয়।")
        return
    ok, info = await force_join_ok(call.bot, call.from_user.id)
    if not ok:
        await call.answer("❌ এখনো channel-এ join করা হয়নি।", show_alert=True)
        return
    await call.answer("✅ Join verified!")
    try:
        await call.message.delete()
    except Exception:
        pass
    await user_home_message(call.bot, call.message.chat.id, call.from_user.id)

@router.callback_query(F.data == "u:home")
async def cb_home(call: CallbackQuery, state: FSMContext):
    if not await guard_user(call.message):
        return
    if not await is_admin(call.from_user.id) and await maintenance_on():
        await call.answer("Bot maintenance mode-এ আছে।", show_alert=True)
        return
    await clear_transient_sessions(state)
    try:
        await call.message.answer(
            "🏠 প্রধান মেনু\n\nপ্রয়োজনীয় সেবা নির্বাচন করুন। 🤝",
            reply_markup=main_kb(await is_admin(call.from_user.id))
        )
        await call.answer()
    except TelegramBadRequest:
        await call.answer()

@router.callback_query(F.data == "u:submit")
async def cb_submit(call: CallbackQuery, state: FSMContext):
    if not await guard_user(call.message):
        return
    allowed, reason = await submission_window_status()
    if not allowed:
        await call.answer("Submission এখন বন্ধ।", show_alert=True)
        await send_or_edit(call, reason, kb([[btn("🔄 Refresh", "u:submit", "primary")], [btn("🏠 Main Menu", "u:home")]]))
        return
    await clear_transient_sessions(state)
    await state.set_state(SubmitState.waiting_category)
    await state.update_data(submission_session=True)
    cats = await fetchall("SELECT id,name,rate FROM categories WHERE enabled=1 ORDER BY sort_order ASC, id ASC")
    if not cats:
        await send_or_edit(
            call,
            "📤 <b>ফাইল সাবমিট</b>\n\n"
            "বর্তমানে কোনো active category নেই। Admin-এর সাথে যোগাযোগ করুন। 🤝",
            kb([[btn("◀️ Back", "u:home")]])
        )
        return
    text = (
        "📤 <b>ফাইল সাবমিট</b>\n\n"
        "আপনার কাজের জন্য একটি category নির্বাচন করুন।\n"
        "Category অনুযায়ী rate ও submission rules দেখানো হবে। 🤝"
    )
    await send_or_edit(call, text, await category_keyboard(include_template=True))
    await call.answer()

@router.callback_query(F.data == "u:template_general")
async def cb_template_general(call: CallbackQuery, state: FSMContext):
    if not await guard_user(call.message):
        return
    # General safe template; do not depend on any category having database ID=1.
    wb = Workbook()
    ws = wb.active
    ws.title = "Submission"
    ws.append(["UID", "Status", "Notes"])
    import io
    bio = io.BytesIO()
    wb.save(bio)
    await call.message.answer_document(
        BufferedInputFile(bio.getvalue(), filename="Apex_Valorem_Template.xlsx"),
        caption="📄 Safe Excel Template\nFormat: UID | Status | Notes"
    )
    await call.answer()

@router.callback_query(F.data == "u:profile")
async def cb_profile(call: CallbackQuery, state: FSMContext):
    if not await guard_user(call.message):
        return
    await leave_support_if_active(call.from_user.id, state)
    await state.clear()
    row = await fetchone("SELECT * FROM users WHERE telegram_id=?", (call.from_user.id,))
    subs = await fetchone("SELECT COUNT(*) c FROM submissions WHERE user_id=?", (call.from_user.id,))
    approved = await fetchone("SELECT COUNT(*) c FROM submissions WHERE user_id=? AND review_status='approved'", (call.from_user.id,))
    text = (f"👤 MY PROFILE\n\n🆔 UID: {call.from_user.id}\n👤 Username: {username_text(call.from_user)}\n"
            f"💰 Balance: {CURRENCY}{fmt_money(row['balance'])}\n📦 Submissions: {subs['c']}\n✅ Approved: {approved['c']}")
    await send_or_edit(call, text, back_kb())

@router.callback_query(F.data == "u:rules")
async def cb_rules(call: CallbackQuery, state: FSMContext):
    if not await guard_user(call.message):
        return
    await leave_support_if_active(call.from_user.id, state); await state.clear()
    rules = await get_setting("rules")
    await send_or_edit(call, "📋 RULES\n\n" + rules, back_kb())

@router.callback_query(F.data == "u:history")
async def cb_history(call: CallbackQuery, state: FSMContext):
    if not await guard_user(call.message):
        return
    await leave_support_if_active(call.from_user.id, state); await state.clear()
    await send_or_edit(call, "📜 HISTORY\n\nআপনার history দেখার বিভাগ নির্বাচন করুন।", kb([
        [btn("📦 Submission History", "h:subs"), btn("💸 Payment History", "h:payments")],
        [btn("💰 Wallet Transactions", "h:tx")], [btn("◀️ Back", "u:home")]
    ]))

@router.callback_query(F.data == "h:subs")
async def history_subs(call: CallbackQuery):
    if not await guard_user(call.message):
        return
    rows = await fetchall("SELECT public_id,category_name,approved_quantity,final_amount,review_status FROM submissions WHERE user_id=? ORDER BY id DESC LIMIT 10", (call.from_user.id,))
    text = "📦 SUBMISSION HISTORY\n\n" + ("\n".join(f"{r['public_id']} • {r['category_name']} • {r['approved_quantity']} • {CURRENCY}{fmt_money(r['final_amount'])} • {r['review_status']}" for r in rows) if rows else "কোনো submission history নেই।")
    await send_or_edit(call, text, back_kb("u:history"))

@router.callback_query(F.data == "h:payments")
async def history_payments(call: CallbackQuery):
    if not await guard_user(call.message):
        return
    rows = await fetchall("SELECT public_id,amount,method,status FROM withdrawals WHERE user_id=? ORDER BY id DESC LIMIT 10", (call.from_user.id,))
    text = "💸 PAYMENT HISTORY\n\n" + ("\n".join(f"{r['public_id']} • {CURRENCY}{fmt_money(r['amount'])} • {r['method']} • {r['status']}" for r in rows) if rows else "কোনো payment history নেই।")
    await send_or_edit(call, text, back_kb("u:history"))

@router.callback_query(F.data == "h:tx")
async def history_tx(call: CallbackQuery):
    if not await guard_user(call.message):
        return
    rows = await fetchall("SELECT public_id,txn_type,amount,source_id,created_at FROM transactions WHERE user_id=? ORDER BY id DESC LIMIT 10", (call.from_user.id,))
    text = "💰 WALLET TRANSACTIONS\n\n" + ("\n".join(f"{r['public_id']} • {r['txn_type']} • {('+' if money(r['amount'])>=0 else '')}{CURRENCY}{fmt_money(abs(r['amount']))} • {r['source_id'] or '-'}" for r in rows) if rows else "কোনো transaction নেই।")
    await send_or_edit(call, text, back_kb("u:history"))

# ============================================================
# 9. CATEGORY / FILE SUBMISSION
# ============================================================
async def category_keyboard(include_template=False):
    """Dynamic user category UI. Only enabled categories are shown; each keeps its admin-selected emoji/style/order."""
    rows = []
    cats = await fetchall("SELECT id,name,rate,emoji,button_style FROM categories WHERE enabled=1 ORDER BY sort_order ASC, id ASC")
    for i in range(0, len(cats), 2):
        pair = cats[i:i+2]
        rows.append([
            btn(f"{c['emoji'] or '📦'} {c['name']} · {CURRENCY}{fmt_money(c['rate'])}", f"u:cat:{c['id']}", c['button_style'] if c['button_style'] in {'primary','success','danger'} else 'primary')
            for c in pair
        ])
    rows.append([btn("◀️ Back", "u:home")])
    return kb(rows)

@router.callback_query(F.data.startswith("u:cat:"))
async def cb_category(call: CallbackQuery, state: FSMContext):
    if not await guard_user(call.message):
        return
    try:
        cid = int(call.data.split(":")[-1])
    except ValueError:
        await call.answer("Invalid category", show_alert=True)
        return
    cat = await get_category(cid)
    if not cat or not cat["enabled"]:
        await call.answer("এই category এখন available নয়।", show_alert=True)
        return
    cols = await get_format(cid)
    await state.set_state(SubmitState.waiting_file)
    await state.update_data(
        submission_session=True, category_id=cid, analysis=None,
        file_id=None, file_name=None, file_size=0
    )
    detail_text = (
        "📦 <b>SUBMISSION DETAILS</b>\n\n"
        f"📌 Category: <b>{cat['name']}</b>\n"
        f"💰 Rate: <b>{CURRENCY}{fmt_money(cat['rate'])}</b> per valid record\n"
        "⏱ Report Time: <b>12–24 hours</b>\n\n"
        f"{submission_notice_text()}\n\n"
        f"📋 Required columns: <b>{' | '.join(cols)}</b>\n\n"
        "🔍 File পাওয়ার পর automatic format validation হবে।\n"
        "⚠️ Final payable amount admin approval-এর ওপর নির্ভর করবে।"
    )
    await send_or_edit(
        call, detail_text,
        kb([
            [btn("◀️ Back to Categories", "u:submit")]
        ])
    )
    await call.answer()

@router.callback_query(F.data.startswith("u:template:"))
async def cb_template(call: CallbackQuery):
    if not await guard_user(call.message):
        return
    try: cid = int(call.data.split(":")[-1])
    except: await call.answer("Invalid", show_alert=True); return
    cat = await get_category(cid)
    if not cat: await call.answer("Category পাওয়া যায়নি।", show_alert=True); return
    data = await create_template(cid)
    await call.message.answer_document(BufferedInputFile(data, filename=f"{safe_filename(cat['name'])}_Template.xlsx"), caption=f"📄 {cat['name']} Template\nRequired format: " + " | ".join(await get_format(cid)))
    await call.answer()

# ============================================================
# 6A. BOTTOM REPLY-KEYBOARD MAIN MENU
# ============================================================
@router.message(F.text == MAIN_SUBMIT)
async def reply_main_submit(message: Message, state: FSMContext):
    if not await guard_user(message):
        return
    # Check the admin-defined submission window immediately when the
    # user presses SUBMIT FILE, before opening the submission session.
    allowed, reason = await submission_window_status()
    if not allowed:
        await state.clear()
        await message.answer(reason, reply_markup=main_kb(await is_admin(message.from_user.id)))
        return
    await clear_transient_sessions(state)
    await state.set_state(SubmitState.waiting_category)
    await state.update_data(submission_session=True)
    cats = await fetchall("SELECT id,name,rate FROM categories WHERE enabled=1 ORDER BY sort_order ASC, id ASC")
    if not cats:
        await message.answer(
            "📤 <b>ফাইল সাবমিট</b>\n\nবর্তমানে কোনো active category নেই। Admin-এর সাথে যোগাযোগ করুন। 🤝",
            reply_markup=kb([[btn("◀️ Back", "u:home")]])
        )
        return
    await message.answer(
        "📤 <b>ফাইল সাবমিট</b>\n\n"
        "আপনার কাজের জন্য একটি category নির্বাচন করুন।\n"
        "Category অনুযায়ী rate ও submission rules দেখানো হবে। 🤝",
        reply_markup=await category_keyboard(include_template=True)
    )

@router.message(F.text == MAIN_PROFILE)
async def reply_main_profile(message: Message, state: FSMContext):
    if not await guard_user(message): return
    await leave_support_if_active(message.from_user.id, state); await state.clear()
    row = await fetchone("SELECT * FROM users WHERE telegram_id=?", (message.from_user.id,))
    subs = await fetchone("SELECT COUNT(*) c FROM submissions WHERE user_id=?", (message.from_user.id,))
    approved = await fetchone("SELECT COUNT(*) c FROM submissions WHERE user_id=? AND review_status='approved'", (message.from_user.id,))
    text = (f"👤 MY PROFILE\n\n🆔 UID: {message.from_user.id}\n👤 Username: {username_text(message.from_user)}\n"
            f"💰 Balance: {CURRENCY}{fmt_money(row['balance'])}\n📦 Submissions: {subs['c']}\n✅ Approved: {approved['c']}")
    await message.answer(text, reply_markup=main_kb(await is_admin(message.from_user.id)))

@router.message(F.text == MAIN_WALLET)
async def reply_main_wallet(message: Message, state: FSMContext):
    if not await guard_user(message): return
    await leave_support_if_active(message.from_user.id, state); await state.clear()
    u = await fetchone("SELECT * FROM users WHERE telegram_id=?", (message.from_user.id,))
    text = (f"💰 WALLET\n\n💵 Available Balance: {CURRENCY}{fmt_money(u['balance'])}\n"
            f"⏳ Pending: {CURRENCY}{fmt_money(u['pending_balance'])}\n📊 Total Earned: {CURRENCY}{fmt_money(u['total_earned'])}\n"
            f"💳 Total Paid: {CURRENCY}{fmt_money(u['total_paid'])}")
    await message.answer(text, reply_markup=kb([
        [btn("💸 Withdraw", "w:withdraw"), btn("💳 Payment Method", "w:method")],
        [btn("📜 Transactions", "h:tx")], [btn("◀️ Back", "u:home")]
    ]))

@router.message(F.text == MAIN_RULES)
async def reply_main_rules(message: Message, state: FSMContext):
    if not await guard_user(message): return
    await leave_support_if_active(message.from_user.id, state); await state.clear()
    await message.answer("📋 RULES\n\n" + await get_setting("rules"), reply_markup=main_kb(await is_admin(message.from_user.id)))

@router.message(F.text == MAIN_HISTORY)
async def reply_main_history(message: Message, state: FSMContext):
    if not await guard_user(message): return
    await leave_support_if_active(message.from_user.id, state); await state.clear()
    await message.answer("📜 HISTORY\n\nআপনার history দেখার বিভাগ নির্বাচন করুন।", reply_markup=kb([
        [btn("📦 Submission History", "h:subs"), btn("💸 Payment History", "h:payments")],
        [btn("💰 Wallet Transactions", "h:tx")], [btn("◀️ Back", "u:home")]
    ]))

@router.message(F.text == MAIN_SUPPORT)
async def reply_main_support(message: Message, state: FSMContext):
    if not await guard_user(message): return
    await leave_support_if_active(message.from_user.id, state); await state.clear()
    await message.answer("💬 SUPPORT CENTER\n\nআপনার সমস্যার ধরন নির্বাচন করুন 👇", reply_markup=kb([
        [btn("📤 File Problem", "s:topic:file"), btn("💳 Payment Problem", "s:topic:payment")],
        [btn("📁 Format Problem", "s:topic:format"), btn("💰 Wallet Problem", "s:topic:wallet")],
        [btn("👨‍💻 Contact Support", "s:start")], [btn("🏠 Main Menu", "u:home")]
    ]))

@router.message(F.document)
async def document_handler(message: Message, state: FSMContext):
    if not await guard_user(message): return
    allowed, reason = await submission_window_status()
    if not allowed:
        await state.clear()
        await message.answer(reason, reply_markup=main_kb(await is_admin(message.from_user.id)))
        return
    current = await state.get_state()
    data = await state.get_data()
    if current != SubmitState.waiting_file.state or not data.get("submission_session") or not data.get("category_id"):
        await message.answer("📌 বর্তমানে কোনো file submission session চালু নেই।\n\n📤 Submit File থেকে নতুন submission শুরু করুন। 🤝")
        return
    doc = message.document
    filename = doc.file_name or "file"
    if not filename.lower().endswith(".xlsx"):
        await message.answer("❌ ফাইল গ্রহণ করা যায়নি\n\nশুধুমাত্র .xlsx Excel ফাইল গ্রহণযোগ্য।\nঅনুগ্রহ করে সঠিক Excel file আবার পাঠান। 🤝", reply_markup=kb([[btn("🔄 আবার চেষ্টা করুন", "u:retryfile", "primary")]]))
        return
    if doc.file_size and doc.file_size > MAX_FILE_BYTES:
        await message.answer("❌ ফাইলটি অনেক বড়। অনুগ্রহ করে ছোট .xlsx file পাঠান।", reply_markup=kb([[btn("🔄 আবার চেষ্টা করুন", "u:retryfile", "primary")]]))
        return
    cid = data.get("category_id")
    cat = await get_category(cid) if cid else None
    if not cat or not cat["enabled"]:
        await state.clear(); await message.answer("📌 নির্বাচিত category আর available নয়। আবার Submit File শুরু করুন।", reply_markup=main_kb(await is_admin(message.from_user.id))); return
    await state.set_state(SubmitState.awaiting_confirmation)
    await state.update_data(file_id=doc.file_id, file_name=filename, file_size=doc.file_size or 0)
    await message.answer(f"📄 ফাইল পাওয়া গেছে\n\n📁 File: {filename}\n📦 Category: {cat['name']}\n\n🔍 ফাইল যাচাই করা হচ্ছে... ")
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp_path = tmp.name; tmp.close()
    try:
        bot = message.bot
        tg_file = await bot.get_file(doc.file_id)
        await bot.download(tg_file, destination=tmp_path)
        required = await get_format(cid)
        result = await analyze_xlsx(tmp_path, required)
        if not result["ok"]:
            reason = result["reason"] or "Format error"
            if "Required column" in reason:
                text = f"❌ ফাইল গ্রহণ করা যায়নি\n\n📌 কারণ: {reason}\n\nঅনুগ্রহ করে নির্ধারিত format অনুযায়ী file ঠিক করে আবার পাঠান। 🤝"
            elif result["duplicate"]:
                text = f"⚠️ ফাইল যাচাই সম্পন্ন\n\nDuplicate records পাওয়া গেছে।\n\n📌 Duplicate: {result['duplicate']}\n\nঅনুগ্রহ করে duplicate records সরিয়ে file আবার submit করুন। 🤝"
            else:
                text = f"❌ ফাইল গ্রহণ করা যায়নি\n\n📌 কারণ: {reason}\n\nঅনুগ্রহ করে file ঠিক করে আবার চেষ্টা করুন। 🤝"
            await state.set_state(SubmitState.waiting_file)
            await message.answer(text, reply_markup=kb([[btn("🔄 আবার চেষ্টা করুন", "u:retryfile", "primary")], [btn("❌ Cancel", "u:home", "danger")]]))
            return
        rate = money(cat["rate"])
        estimated = money(rate * result["valid"])
        await state.update_data(analysis=result, locked_rate=float(rate), estimated=float(estimated))
        text = (f"📊 FILE ANALYSIS\n\nTotal rows: {result['total']}\nValid: {result['valid']}\nInvalid: {result['invalid']}\nDuplicate: {result['duplicate']}\n\n"
                f"Rate: {CURRENCY}{fmt_money(rate)}\nEstimated: {result['valid']} × {CURRENCY}{fmt_money(rate)} = {CURRENCY}{fmt_money(estimated)}\n\n⚠️ Final amount will be based on admin approval.")
        await message.answer(text, reply_markup=kb([[btn("📤 Send - Submit File", "u:confirm_submit", "success")], [btn("❌ Cancel", "u:cancel_submit", "danger")]]))
    except Exception:
        log.exception("download/analysis error")
        await state.set_state(SubmitState.waiting_file)
        await message.answer("❌ ফাইল যাচাই করা যায়নি। Telegram file download বা Excel read-এ সমস্যা হয়েছে। আবার চেষ্টা করুন।", reply_markup=kb([[btn("🔄 আবার চেষ্টা করুন", "u:retryfile", "primary")]]))
    finally:
        try: os.unlink(tmp_path)
        except OSError: pass

@router.message(StateFilter(None), F.document)
async def no_session_document(message: Message):
    await message.answer("📌 বর্তমানে কোনো file submission session চালু নেই।\n\n📤 Submit File থেকে নতুন submission শুরু করুন। 🤝")

@router.callback_query(F.data == "u:retryfile")
async def retry_file(call: CallbackQuery, state: FSMContext):
    if not await guard_user(call.message):
        return
    data = await state.get_data(); cid = data.get("category_id")
    if not cid:
        await cb_submit(call, state); return
    await state.set_state(SubmitState.waiting_file)
    await state.update_data(submission_session=True)
    cat = await get_category(cid)
    await send_or_edit(call, f"📄 আবার আপনার Excel file পাঠান।\n\n📦 Category: {cat['name']}\n💰 Rate: {CURRENCY}{fmt_money(cat['rate'])}\n\n{submission_notice_text()}", kb([[btn("◀️ Back", "u:submit")]]))

@router.callback_query(F.data == "u:cancel_submit")
async def cancel_submit(call: CallbackQuery, state: FSMContext):
    if not await guard_user(call.message):
        return
    await state.clear()
    await send_or_edit(call, "❌ Submission বাতিল করা হয়েছে।", main_kb(await is_admin(call.from_user.id)))

@router.callback_query(F.data == "u:confirm_submit")
async def confirm_submit(call: CallbackQuery, state: FSMContext):
    if not await guard_user(call.message):
        return
    allowed, reason = await submission_window_status()
    if not allowed:
        await state.clear()
        await call.answer("Submission window বন্ধ।", show_alert=True)
        await send_or_edit(call, reason, main_kb(await is_admin(call.from_user.id)))
        return
    data = await state.get_data()
    if await state.get_state() != SubmitState.awaiting_confirmation.state or not data.get("analysis"):
        await call.answer("এই submission session আর valid নয়।", show_alert=True); return
    result = data["analysis"]
    cid = int(data["category_id"])
    cat = await get_category(cid)
    if not cat:
        await call.answer("Category পাওয়া যায়নি।", show_alert=True); return
    rate = money(data.get("locked_rate", cat["rate"]))
    estimated = money(data.get("estimated", 0))
    public_id = await next_public_id("submissions", "SUB")
    async with DB_LOCK:
        cur = await DB.execute("INSERT INTO submissions(public_id,user_id,category_id,category_name,file_id,file_name,records_count,valid_count,invalid_count,duplicate_count,locked_rate,estimated_amount,validation_status,review_status,payment_status,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (public_id, call.from_user.id, cid, cat["name"], data["file_id"], safe_filename(data["file_name"]), result["total"], result["valid"], result["invalid"], result["duplicate"], float(rate), float(estimated), "valid", "pending", "unpaid", now_str()))
        sid = cur.lastrowid
        for rec in result["records"]:
            await DB.execute("INSERT INTO submission_records(submission_id,row_number,uid,status,notes,data_json,valid,duplicate) VALUES(?,?,?,?,?,?,?,?)", (sid, rec["row_number"], rec["uid"], rec["status"], rec["notes"], json.dumps(rec["data"], ensure_ascii=False), int(rec["valid"]), int(rec["duplicate"])))
        await DB.commit()
    await state.clear()
    await call.message.edit_text(f"🟡 SUBMISSION RECEIVED\n\n📄 Submission: #{public_id}\n📦 Category: {cat['name']}\n📊 Records: {result['valid']}\n💰 Rate: {CURRENCY}{fmt_money(rate)}\n💵 Estimated: {CURRENCY}{fmt_money(estimated)}\n🕐 Status: Under Review\n\nআপনার file review-এর জন্য পাঠানো হয়েছে। 🤝", reply_markup=kb([[btn("📄 View Submission", f"u:viewsub:{sid}")], [btn("🏠 Main Menu", "u:home")]]))
    await notify_new_submission(call.bot, sid)
    await call.answer("Submission পাঠানো হয়েছে।")

async def notify_new_submission(bot: Bot, sid: int):
    row = await fetchone("SELECT s.*,u.username FROM submissions s LEFT JOIN users u ON u.telegram_id=s.user_id WHERE s.id=?", (sid,))
    if not row: return
    text = f"📥 নতুন ফাইল সাবমিশন\n\n👤 @{row['username'] or 'no_username'}\n🆔 UID: {row['user_id']}\n📦 Category: {row['category_name']}\n📄 {row['file_name']}\n📊 Records: {row['records_count']}\n💰 Rate: {CURRENCY}{fmt_money(row['locked_rate'])}\n\n#{row['public_id']}"
    await admin_broadcast_notification(bot, text, kb([[btn("🔍 Review", f"a:sub:{sid}", "primary")]]))

async def admin_broadcast_notification(bot: Bot, text: str, markup=None):
    admins = await fetchall("SELECT user_id FROM admins")
    for a in admins:
        try: await bot.send_message(a["user_id"], text, reply_markup=markup)
        except (TelegramForbiddenError, TelegramBadRequest): pass
        except Exception: log.exception("admin notification failed")

@router.callback_query(F.data.startswith("u:viewsub:"))
async def user_view_submission(call: CallbackQuery):
    if not await guard_user(call.message):
        return
    try: sid = int(call.data.split(":")[-1])
    except: await call.answer("Invalid", show_alert=True); return
    r = await fetchone("SELECT * FROM submissions WHERE id=? AND user_id=?", (sid, call.from_user.id))
    if not r: await call.answer("Submission পাওয়া যায়নি।", show_alert=True); return
    await send_or_edit(call, f"📄 SUBMISSION DETAILS\n\n#{r['public_id']}\n📦 {r['category_name']}\n📊 Records: {r['records_count']}\n💰 Rate: {CURRENCY}{fmt_money(r['locked_rate'])}\n💵 Estimated: {CURRENCY}{fmt_money(r['estimated_amount'])}\n🕐 Status: {r['review_status']}", back_kb())

# ============================================================
# 10. WALLET / PAYMENT METHOD / WITHDRAWAL
# ============================================================
@router.callback_query(F.data == "u:wallet")
async def cb_wallet(call: CallbackQuery, state: FSMContext):
    if not await guard_user(call.message):
        return
    await leave_support_if_active(call.from_user.id, state); await state.clear()
    u = await fetchone("SELECT * FROM users WHERE telegram_id=?", (call.from_user.id,))
    text = f"💰 WALLET\n\n💵 Available Balance: {CURRENCY}{fmt_money(u['balance'])}\n⏳ Pending: {CURRENCY}{fmt_money(u['pending_balance'])}\n📊 Total Earned: {CURRENCY}{fmt_money(u['total_earned'])}\n💳 Total Paid: {CURRENCY}{fmt_money(u['total_paid'])}"
    await send_or_edit(call, text, kb([[btn("💸 Withdraw", "w:withdraw"), btn("💳 Payment Method", "w:method")], [btn("📜 Transactions", "h:tx")], [btn("◀️ Back", "u:home")]]))

@router.callback_query(F.data == "w:method")
async def payment_method_screen(call: CallbackQuery, state: FSMContext):
    if not await guard_user(call.message):
        return
    await leave_support_if_active(call.from_user.id, state)
    pm = await fetchone("SELECT * FROM payment_methods WHERE user_id=?", (call.from_user.id,))
    if pm and pm["method"] and pm["number"]:
        text = f"💳 PAYMENT METHOD\n\nMethod: {pm['method']}\nNumber: {mask_number(pm['number'])}"
        markup = kb([[btn("✏️ Change Method", "w:change_method"), btn("✏️ Change Number", "w:change_number")], [btn("◀️ Back", "u:wallet")]])
    else:
        text = "💳 PAYMENT METHOD\n\nআপনার payment method নির্বাচন করুন।"
        markup = kb([[btn("🟣 bKash", "w:setmethod:bkash"), btn("🟠 Nagad", "w:setmethod:nagad")], [btn("◀️ Back", "u:wallet")]])
    await send_or_edit(call, text, markup)

@router.callback_query(F.data == "w:change_method")
async def change_method(call: CallbackQuery):
    if not await guard_user(call.message):
        return
    await send_or_edit(call, "💳 নতুন payment method নির্বাচন করুন।", kb([[btn("🟣 bKash", "w:setmethod:bkash"), btn("🟠 Nagad", "w:setmethod:nagad")], [btn("◀️ Back", "w:method")]]))

@router.callback_query(F.data.startswith("w:setmethod:"))
async def set_method(call: CallbackQuery, state: FSMContext):
    if not await guard_user(call.message):
        return
    method = call.data.split(":")[-1].lower()
    if method not in {"bkash", "nagad"}: await call.answer("Invalid method", show_alert=True); return
    pm = await fetchone("SELECT * FROM payment_methods WHERE user_id=?", (call.from_user.id,))
    await state.set_state(UserInputState.payment_number)
    await state.update_data(payment_method=method)
    prompt = f"📱 আপনার {method} number দিন।\n\nExample: 01XXXXXXXXX"
    await send_or_edit(call, prompt, kb([[btn("❌ Cancel", "w:method")]]))

@router.callback_query(F.data == "w:change_number")
async def change_number(call: CallbackQuery, state: FSMContext):
    if not await guard_user(call.message):
        return
    pm = await fetchone("SELECT method FROM payment_methods WHERE user_id=?", (call.from_user.id,))
    if not pm or not pm["method"]:
        await call.answer("আগে method নির্বাচন করুন।", show_alert=True); return
    await state.set_state(UserInputState.payment_number)
    await state.update_data(payment_method=pm["method"])
    await send_or_edit(call, f"📱 আপনার {pm['method']} number দিন।\n\nExample: 01XXXXXXXXX", kb([[btn("❌ Cancel", "w:method")]]))

@router.message(UserInputState.payment_number)
async def receive_payment_number(message: Message, state: FSMContext):
    number = (message.text or "").strip().replace(" ", "")
    if not re.fullmatch(r"01\d{9}", number):
        await message.answer("❌ সঠিক Bangladesh mobile number দিন।\nExample: 01XXXXXXXXX")
        return
    data = await state.get_data(); method = data.get("payment_method")
    await execute("INSERT INTO payment_methods(user_id,method,number,updated_at) VALUES(?,?,?,?) ON CONFLICT(user_id) DO UPDATE SET method=excluded.method,number=excluded.number,updated_at=excluded.updated_at", (message.from_user.id, method, number, now_str()))
    await state.clear()
    await message.answer(f"✅ Payment method save হয়েছে।\n\n💳 Method: {method}\n📱 Number: {mask_number(number)}", reply_markup=kb([[btn("◀️ Back", "u:wallet")]]))

@router.callback_query(F.data == "w:withdraw")
async def withdraw_screen(call: CallbackQuery, state: FSMContext):
    if not await guard_user(call.message):
        return
    await leave_support_if_active(call.from_user.id, state)
    u = await fetchone("SELECT balance FROM users WHERE telegram_id=?", (call.from_user.id,))
    pm = await fetchone("SELECT * FROM payment_methods WHERE user_id=?", (call.from_user.id,))
    minimum = money(await get_setting("min_withdraw", str(DEFAULT_MIN_WITHDRAW)))
    if not pm or not pm["method"] or not pm["number"]:
        await send_or_edit(call, "💳 Withdrawal-এর আগে payment method সেট করুন।", kb([[btn("💳 Payment Method", "w:method")], [btn("◀️ Back", "u:wallet")]])); return
    if money(u["balance"]) < minimum:
        await send_or_edit(call, f"💸 WITHDRAWAL\n\n💰 Available Balance: {CURRENCY}{fmt_money(u['balance'])}\n\n📌 Minimum withdrawal: {CURRENCY}{fmt_money(minimum)}", back_kb("u:wallet")); return
    await send_or_edit(call, f"💸 WITHDRAWAL\n\n💰 Available Balance: {CURRENCY}{fmt_money(u['balance'])}\n💳 Payment: {pm['method']}\n📱 Number: {mask_number(pm['number'])}\n\nAmount নির্বাচন করুন:", kb([[btn(f"{CURRENCY}500", "w:amt:500"), btn(f"{CURRENCY}1,000", "w:amt:1000")], [btn("Full Balance", "w:amt:full"), btn("Custom Amount", "w:amt:custom")], [btn("◀️ Back", "u:wallet")]]))

@router.callback_query(F.data.startswith("w:amt:"))
async def withdrawal_amount(call: CallbackQuery, state: FSMContext):
    if not await guard_user(call.message):
        return
    typ = call.data.split(":")[-1]
    if typ == "custom":
        await state.set_state(UserInputState.custom_withdraw)
        await send_or_edit(call, "💸 আপনার withdrawal amount লিখুন।", kb([[btn("❌ Cancel", "u:wallet")]])); return
    u = await fetchone("SELECT balance FROM users WHERE telegram_id=?", (call.from_user.id,))
    amount = money(u["balance"]) if typ == "full" else money(typ)
    await prepare_withdraw_confirmation(call, state, amount)

@router.message(UserInputState.custom_withdraw)
async def custom_withdraw(message: Message, state: FSMContext):
    try: amount = money(message.text)
    except: amount = Decimal("0")
    await state.clear()
    # Reuse a message-only confirmation helper.
    await prepare_withdraw_message(message, amount)

async def prepare_withdraw_message(message: Message, amount: Decimal):
    u = await fetchone("SELECT balance FROM users WHERE telegram_id=?", (message.from_user.id,))
    pm = await fetchone("SELECT * FROM payment_methods WHERE user_id=?", (message.from_user.id,))
    minimum = money(await get_setting("min_withdraw", str(DEFAULT_MIN_WITHDRAW)))
    if amount < minimum or amount > money(u["balance"]):
        await message.answer(f"❌ Withdrawal amount valid নয়।\n\nMinimum: {CURRENCY}{fmt_money(minimum)}\nAvailable: {CURRENCY}{fmt_money(u['balance'])}"); return
    await message.answer(f"💸 WITHDRAWAL CONFIRMATION\n\nAmount: {CURRENCY}{fmt_money(amount)}\nMethod: {pm['method']}\nNumber: {mask_number(pm['number'])}", reply_markup=confirm_kb(f"w:confirm:{amount}"))

async def prepare_withdraw_confirmation(call: CallbackQuery, state: FSMContext, amount: Decimal):
    await state.clear()
    u = await fetchone("SELECT balance FROM users WHERE telegram_id=?", (call.from_user.id,))
    pm = await fetchone("SELECT * FROM payment_methods WHERE user_id=?", (call.from_user.id,))
    minimum = money(await get_setting("min_withdraw", str(DEFAULT_MIN_WITHDRAW)))
    if not pm or not pm["method"]:
        await call.answer("Payment method সেট করুন।", show_alert=True); return
    if amount < minimum or amount > money(u["balance"]):
        await call.answer("Withdrawal amount valid নয়।", show_alert=True); return
    await send_or_edit(call, f"💸 WITHDRAWAL CONFIRMATION\n\nAmount: {CURRENCY}{fmt_money(amount)}\nMethod: {pm['method']}\nNumber: {mask_number(pm['number'])}", confirm_kb(f"w:confirm:{amount}"))

@router.callback_query(F.data.startswith("w:confirm:"))
async def confirm_withdraw(call: CallbackQuery):
    if not await guard_user(call.message):
        return
    try: amount = money(call.data.split(":", 2)[2])
    except: await call.answer("Invalid amount", show_alert=True); return
    minimum = money(await get_setting("min_withdraw", str(DEFAULT_MIN_WITHDRAW)))
    async with DB_LOCK:
        u = await fetchone("SELECT * FROM users WHERE telegram_id=?", (call.from_user.id,))
        pm = await fetchone("SELECT * FROM payment_methods WHERE user_id=?", (call.from_user.id,))
        if not u or not pm or amount < minimum or amount > money(u["balance"]):
            await call.answer("Withdrawal valid নয়।", show_alert=True); return
        public_id = await next_public_id("withdrawals", "WD")
        new_balance = money(u["balance"]) - amount
        await DB.execute("UPDATE users SET balance=?,updated_at=? WHERE telegram_id=?", (float(new_balance), now_str(), call.from_user.id))
        await DB.execute("INSERT INTO withdrawals(public_id,user_id,amount,method,number,status,created_at) VALUES(?,?,?,?,?,?,?)", (public_id, call.from_user.id, float(amount), pm["method"], pm["number"], "pending", now_str()))
        await DB.commit()
    await call.message.edit_text(f"💸 Withdrawal Request Sent\n\n💰 Amount: {CURRENCY}{fmt_money(amount)}\n💳 Method: {pm['method']}\n🧾 Request: #{public_id}\n\nআপনার অনুরোধটি সফলভাবে গ্রহণ করা হয়েছে। 🤝", reply_markup=kb([[btn("🏠 Main Menu", "u:home")]]))
    await admin_broadcast_notification(call.bot, f"💸 নতুন Withdrawal Request\n\n👤 {username_text(call.from_user)}\n🆔 UID: {call.from_user.id}\n💰 Amount: {CURRENCY}{fmt_money(amount)}\n💳 {pm['method']}\n📱 {mask_number(pm['number'])}\n🧾 #{public_id}", kb([[btn("🔍 View", f"a:wd:{public_id}")], [btn("💸 Process Payment", f"a:wd:{public_id}")]]))

# ============================================================
# 11. SUPPORT
# ============================================================
@router.callback_query(F.data == "u:support")
async def support_center(call: CallbackQuery, state: FSMContext):
    if not await guard_user(call.message):
        return
    await leave_support_if_active(call.from_user.id, state); await state.clear()
    await send_or_edit(call, "💬 SUPPORT CENTER\n\nআপনার সমস্যার ধরন নির্বাচন করুন 👇", kb([
        [btn("📤 File Problem", "s:topic:file"), btn("💳 Payment Problem", "s:topic:payment")],
        [btn("📁 Format Problem", "s:topic:format"), btn("💰 Wallet Problem", "s:topic:wallet")],
        [btn("👨‍💻 Contact Support", "s:start")], [btn("🏠 Main Menu", "u:home")]
    ]))

@router.callback_query(F.data.startswith("s:topic:"))
async def support_topic(call: CallbackQuery, state: FSMContext):
    if not await guard_user(call.message):
        return
    topic = call.data.split(":")[-1]
    await state.update_data(support_topic=topic)
    await send_or_edit(call, "💬 SUPPORT CENTER\n\nTopic selected. Contact Support চাপলে live chat শুরু হবে। 🤝", kb([[btn("👨‍💻 Contact Support", "s:start")], [btn("◀️ Back", "u:support")]]))

@router.callback_query(F.data == "s:start")
async def support_start(call: CallbackQuery, state: FSMContext):
    if not await guard_user(call.message):
        return
    data = await state.get_data(); topic = data.get("support_topic", "general")
    ticket = await fetchone("SELECT * FROM support_tickets WHERE user_id=? AND status='active' ORDER BY id DESC LIMIT 1", (call.from_user.id,))
    if ticket:
        await state.set_state(SupportState.active); await state.update_data(support_active=True, ticket_id=ticket["id"])
        await send_or_edit(call, f"💬 LIVE SUPPORT\n\n🧾 Ticket: #{ticket['public_id']}\n🟢 Live Chat: Active\n\nআপনার মেসেজ লিখুন। সাপোর্ট টিম আপনাকে এখানেই উত্তর দেবে। 🤝", kb([[btn("❌ Close Chat", "s:close")]])); return
    public_id = await next_public_id("support_tickets", "TKT")
    cur = await execute("INSERT INTO support_tickets(public_id,user_id,topic,status,created_at) VALUES(?,?,?,?,?)", (public_id, call.from_user.id, topic, "active", now_str()))
    tid = cur.lastrowid
    await state.set_state(SupportState.active); await state.update_data(support_active=True, ticket_id=tid)
    await send_or_edit(call, f"💬 LIVE SUPPORT\n\n🧾 Ticket: #{public_id}\n🟢 Live Chat: Active\n\nআপনার মেসেজ লিখুন। সাপোর্ট টিম আপনাকে এখানেই উত্তর দেবে। 🤝", kb([[btn("❌ Close Chat", "s:close")]]))

@router.callback_query(F.data == "s:close")
async def support_close(call: CallbackQuery, state: FSMContext):
    if not await guard_user(call.message):
        return
    data = await state.get_data(); tid = data.get("ticket_id")
    if tid:
        await execute("UPDATE support_tickets SET status='closed',closed_at=? WHERE id=?", (now_str(), tid))
    await state.clear()
    await send_or_edit(call, "🔒 Chat Closed\n\nআপনার support chat বন্ধ করা হয়েছে।\nপ্রয়োজনে আবার Support থেকে নতুন chat শুরু করতে পারেন। 🤝", kb([[btn("💬 Support", "u:support"), btn("🏠 Main Menu", "u:home")]]))

@router.message(SupportState.active, F.text)
async def support_user_message(message: Message, state: FSMContext):
    data = await state.get_data(); tid = data.get("ticket_id")
    if not tid:
        await state.clear(); await message.answer("Support session পাওয়া যায়নি। আবার Support থেকে শুরু করুন।"); return
    ticket = await fetchone("SELECT * FROM support_tickets WHERE id=? AND user_id=? AND status='active'", (tid, message.from_user.id))
    if not ticket:
        await state.clear(); await message.answer("🔒 Chat Closed\nআবার Support থেকে নতুন chat শুরু করুন।"); return
    await execute("INSERT INTO support_messages(ticket_id,sender_id,sender_role,message_type,body,created_at) VALUES(?,?,?,?,?,?)", (tid, message.from_user.id, "user", "text", message.text[:4000], now_str()))
    text = f"💬 নতুন সাপোর্ট মেসেজ\n\n👤 {username_text(message.from_user)}\n🆔 UID: {message.from_user.id}\n🧾 #{ticket['public_id']}\n\n💬 “{message.text[:1500]}”"
    await admin_broadcast_notification(message.bot, text, kb([[btn("💬 Reply", f"a:sreply:{tid}")]]))
    await message.answer("📨 আপনার মেসেজ সাপোর্ট টিমের কাছে পাঠানো হয়েছে। 🤝", reply_markup=kb([[btn("❌ Close Chat", "s:close")]]))

@router.message(SupportState.active)
async def support_non_text(message: Message):
    await message.answer("📌 Support chat-এ আপাতত text message পাঠান। 🤝")

@router.message(F.text == MAIN_ADMIN)
async def admin_entry(message: Message, state: FSMContext):
    if not await admin_guard(message):
        return
    await state.clear()
    await message.answer("👑 <b>ADMIN PANEL</b>\n\nAdmin tools নির্বাচন করুন।", reply_markup=ReplyKeyboardRemove())
    await message.answer("🛠️ নিচের অপশন থেকে প্রয়োজনীয় কাজ নির্বাচন করুন।", reply_markup=admin_home_kb())

@router.message(Command("admin"))
async def admin_command(message: Message, state: FSMContext):
    if not await admin_guard(message):
        return
    await state.clear()
    await message.answer("👑 <b>ADMIN PANEL</b>\n\nAdmin tools নির্বাচন করুন।", reply_markup=ReplyKeyboardRemove())
    await message.answer("🛠️ নিচের অপশন থেকে প্রয়োজনীয় কাজ নির্বাচন করুন।", reply_markup=admin_home_kb())

# ============================================================
# 12. ADMIN HOME / DASHBOARD
# ============================================================
async def admin_guard(call_or_message) -> bool:
    """Allow only configured Telegram admins/owner to use admin callbacks and states."""
    user_id = call_or_message.from_user.id
    if await is_admin(user_id):
        return True
    if isinstance(call_or_message, CallbackQuery):
        await call_or_message.answer("⛔ Admin access নেই।", show_alert=True)
    else:
        await call_or_message.answer("⛔ Admin access নেই।")
    return False

def admin_home_kb():
    return kb([
        [btn("📊 Dashboard", "a:dashboard"), btn("📥 Submissions", "a:subs")],
        [btn("💳 Payments", "a:payments"), btn("👥 Users", "a:users")],
        [btn("📦 Categories", "a:categories"), btn("💰 Balance Manager", "a:balance")],
        [btn("📢 Broadcast", "a:broadcast"), btn("⚙️ Settings", "a:settings")],
        [btn("📝 Admin Logs", "a:logs")], [btn("🏠 User Main Menu", "u:home")]
    ])

@router.callback_query(F.data == "a:home")
async def admin_home(call: CallbackQuery, state: FSMContext):
    if not await admin_guard(call): return
    await state.clear()
    try:
        await call.message.answer("🔒 User keyboard বন্ধ করা হয়েছে।", reply_markup=ReplyKeyboardRemove())
    except Exception:
        pass
    await send_or_edit(call, "👑 ADMIN PANEL\n\nAdmin tools নির্বাচন করুন।", admin_home_kb())

@router.callback_query(F.data == "a:dashboard")
async def admin_dashboard(call: CallbackQuery):
    if not await admin_guard(call): return
    users = await fetchone("SELECT COUNT(*) c FROM users")
    pending = await fetchone("SELECT COUNT(*) c FROM submissions WHERE review_status='pending'")
    today = datetime.now(TIMEZONE).strftime("%Y-%m-%d")
    approved_today = await fetchone("SELECT COUNT(*) c FROM submissions WHERE review_status='approved' AND substr(reviewed_at,1,10)=?", (today,))
    pp = await fetchone("SELECT COALESCE(SUM(amount),0) s FROM withdrawals WHERE status='pending'")
    paid = await fetchone("SELECT COALESCE(SUM(amount),0) s FROM withdrawals WHERE status='paid' AND substr(processed_at,1,10)=?", (today,))
    text = f"📊 DASHBOARD\n\n👥 Users: {users['c']}\n📥 Pending Files: {pending['c']}\n✅ Approved Today: {approved_today['c']}\n💳 Pending Payments: {CURRENCY}{fmt_money(pp['s'])}\n💸 Paid Today: {CURRENCY}{fmt_money(paid['s'])}"
    await send_or_edit(call, text, kb([[btn("📥 Submissions", "a:subs"), btn("💳 Payments", "a:payments")], [btn("🔄 Refresh", "a:dashboard")], [btn("◀️ Back", "a:home")]]))

# ============================================================
# 13. ADMIN SUBMISSIONS
# ============================================================
@router.callback_query(F.data == "a:subs")
async def admin_submissions(call: CallbackQuery):
    if not await admin_guard(call): return
    counts = {}
    for st in ("pending", "approved", "rejected"):
        r = await fetchone("SELECT COUNT(*) c FROM submissions WHERE review_status=?", (st,)); counts[st] = r["c"]
    await send_or_edit(call, f"📥 SUBMISSIONS\n\n🟡 Pending: {counts['pending']}\n✅ Approved: {counts['approved']}\n❌ Rejected: {counts['rejected']}", kb([[btn("🟡 Pending", "a:sublist:pending"), btn("✅ Approved", "a:sublist:approved")], [btn("❌ Rejected", "a:sublist:rejected")], [btn("🔎 Search", "a:searchsub")], [btn("◀️ Back", "a:home")]]))

@router.callback_query(F.data.startswith("a:sublist:"))
async def admin_sublist(call: CallbackQuery):
    if not await admin_guard(call): return
    st = call.data.split(":")[-1]
    rows = await fetchall("SELECT id,public_id,user_id,category_name,file_name,records_count,locked_rate,estimated_amount FROM submissions WHERE review_status=? ORDER BY id DESC LIMIT 15", (st,))
    text = f"📥 {st.upper()} SUBMISSIONS\n\n" + ("\n".join(f"#{r['public_id']} • UID {r['user_id']} • {r['category_name']} • {r['records_count']} • {CURRENCY}{fmt_money(r['estimated_amount'])}" for r in rows) if rows else "কোনো submission নেই।")
    rows_k = [[btn(f"📄 #{r['public_id']}", f"a:sub:{r['id']}")] for r in rows]
    rows_k.append([btn("◀️ Back", "a:subs")])
    await send_or_edit(call, text, kb(rows_k))

@router.callback_query(F.data.startswith("a:sub:"))
async def admin_submission_detail(call: CallbackQuery):
    if not await admin_guard(call): return
    try: sid = int(call.data.split(":")[-1])
    except: await call.answer("Invalid", show_alert=True); return
    r = await fetchone("SELECT s.*,u.username FROM submissions s LEFT JOIN users u ON u.telegram_id=s.user_id WHERE s.id=?", (sid,))
    if not r: await call.answer("Submission পাওয়া যায়নি।", show_alert=True); return
    u = await fetchone("SELECT * FROM users WHERE telegram_id=?", (r["user_id"],))
    daily = await fetchone("SELECT COUNT(*) c, SUM(CASE WHEN review_status='approved' THEN 1 ELSE 0 END) approved, SUM(CASE WHEN review_status='rejected' THEN 1 ELSE 0 END) rejected, SUM(CASE WHEN review_status='pending' THEN 1 ELSE 0 END) pending FROM submissions WHERE user_id=? AND substr(created_at,1,10)=?", (r["user_id"], datetime.now(TIMEZONE).strftime("%Y-%m-%d")))
    unpaid = await fetchone("SELECT COALESCE(SUM(CASE WHEN review_status='approved' AND payment_status!='paid' THEN final_amount ELSE 0 END),0) s FROM submissions WHERE user_id=?", (r["user_id"],))
    text = (f"📄 <b>#{r['public_id']}</b>\n\n👤 User: @{r['username'] or 'no_username'}\n🆔 UID: {r['user_id']}\n📦 Category: {r['category_name']}\n📄 File: {r['file_name']}\n📊 Submitted: {r['records_count']} | Valid: {r['valid_count']}\n✅ Approved Qty: {r['approved_quantity']}\n💰 Rate: {CURRENCY}{fmt_money(r['locked_rate'])}\n🧮 Calculated: {CURRENCY}{fmt_money(r['final_amount'])}\n💳 Credited: {CURRENCY}{fmt_money(r['credited_amount'])}\n🕐 Review: {r['review_status']}\n💳 Balance Status: {r['payment_status']}\n\n👤 <b>User Balance</b>: {CURRENCY}{fmt_money(u['balance'])}\n📆 <b>Today</b>: {daily['c']} files | ✅ {daily['approved']} | ❌ {daily['rejected']} | ⏳ {daily['pending']}\n💰 <b>Approved Unpaid</b>: {CURRENCY}{fmt_money(unpaid['s'])}")
    buttons = [[btn("📂 Open File", f"a:file:{sid}"), btn("📊 View Records", f"a:records:{sid}")], [btn("✏️ Adjust", f"a:adjust:{sid}")]]
    if r["review_status"] == "pending": buttons.append([btn("✅ Approve", f"a:approve:{sid}"), btn("❌ Reject", f"a:reject:{sid}")])
    elif r["review_status"] == "approved" and r["payment_status"] != "paid":
        buttons.append([btn(f"💰 Add {CURRENCY}{fmt_money(r['final_amount'])}", f"a:addsub:{sid}"), btn("✏️ Custom Amount", f"a:customsub:{sid}")])
        buttons.append([btn("📊 User Daily Stats", f"a:ustats:{r['user_id']}")])
    buttons.append([btn("◀️ Back", "a:subs")])
    await send_or_edit(call, text, kb(buttons))

@router.callback_query(F.data.startswith("a:file:"))
async def admin_open_file(call: CallbackQuery):
    if not await admin_guard(call): return
    sid = int(call.data.split(":")[-1]); r = await fetchone("SELECT * FROM submissions WHERE id=?", (sid,))
    if not r: await call.answer("File পাওয়া যায়নি।", show_alert=True); return
    try:
        await call.message.answer_document(r["file_id"], caption=f"📂 FILE\nSubmission: #{r['public_id']}\n📄 {r['file_name']}")
        await call.answer("File পাঠানো হয়েছে।")
    except Exception:
        await call.answer("Telegram file access failed.", show_alert=True)

@router.callback_query(F.data.startswith("a:records:"))
async def admin_records(call: CallbackQuery):
    if not await admin_guard(call): return
    sid = int(call.data.split(":")[-1]); r = await fetchone("SELECT public_id FROM submissions WHERE id=?", (sid,))
    if not r: await call.answer("Not found", show_alert=True); return
    recs = await fetchall("SELECT row_number,uid,status,notes,valid,duplicate FROM submission_records WHERE submission_id=? ORDER BY row_number LIMIT 50", (sid,))
    lines = [f"📊 VIEW RECORDS — #{r['public_id']}", "", "UID | Status | Notes"]
    for x in recs:
        lines.append(f"{x['uid'] or '-'} | {x['status'] or '-'} | {x['notes'] or '-'}")
    await send_or_edit(call, "\n".join(lines)[:3900], back_kb(f"a:sub:{sid}"))

# ============================================================
# 14. APPROVE / ADJUST / REJECT
# ============================================================
@router.callback_query(F.data.startswith("a:adjust:"))
async def admin_adjust(call: CallbackQuery, state: FSMContext):
    if not await admin_guard(call): return
    sid = int(call.data.split(":")[-1]); r = await fetchone("SELECT * FROM submissions WHERE id=?", (sid,))
    if not r or r["review_status"] != "pending": await call.answer("Submission আর pending নেই।", show_alert=True); return
    await state.set_state(AdminState.adjust_quantity); await state.update_data(sid=sid)
    await send_or_edit(call, f"✏️ APPROVED QUANTITY\n\nSubmitted: {r['records_count']}\n\nকতটি approve করবেন? 0–{r['valid_count']} এর মধ্যে সংখ্যা পাঠান।", kb([[btn("◀️ Back", f"a:sub:{sid}")]]))

@router.message(AdminState.adjust_quantity)
async def adjust_quantity_input(message: Message, state: FSMContext):
    data = await state.get_data(); sid = data.get("sid")
    try: q = int((message.text or "").strip())
    except: q = -1
    r = await fetchone("SELECT * FROM submissions WHERE id=?", (sid,))
    if not r or q < 0 or q > r["valid_count"]:
        await message.answer(f"❌ Valid quantity দিন: 0 থেকে {r['valid_count'] if r else 0}।"); return
    amount = money(q * money(r["locked_rate"]))
    await state.clear()
    await message.answer(f"✏️ APPROVED QUANTITY\n\nSubmitted: {r['records_count']}\nApproved: {q}\nRate: {CURRENCY}{fmt_money(r['locked_rate'])}\n\nAmount:\n{q} × {CURRENCY}{fmt_money(r['locked_rate'])} = {CURRENCY}{fmt_money(amount)}", reply_markup=confirm_kb(f"a:approveq:{sid}:{q}", f"a:sub:{sid}"))

@router.callback_query(F.data.startswith("a:approve:"))
async def admin_approve_prompt(call: CallbackQuery):
    if not await admin_guard(call): return
    sid = int(call.data.split(":")[-1]); r = await fetchone("SELECT * FROM submissions WHERE id=?", (sid,))
    if not r or r["review_status"] != "pending": await call.answer("Submission আর pending নেই।", show_alert=True); return
    amount = money(r["valid_count"] * money(r["locked_rate"]))
    await send_or_edit(call, f"⚠️ CONFIRM ACTION\n\nআপনি কি সত্যিই এই submission approve করতে চান?\n\n📄 #{r['public_id']}\n📊 Approved: {r['valid_count']}\n💰 Amount: {CURRENCY}{fmt_money(amount)}", confirm_kb(f"a:approveq:{sid}:{r['valid_count']}", f"a:sub:{sid}"))

@router.callback_query(F.data.startswith("a:approveq:"))
async def admin_approve_execute(call: CallbackQuery):
    if not await admin_guard(call): return
    parts = call.data.split(":"); sid = int(parts[2]); q = int(parts[3])
    async with DB_LOCK:
        r = await fetchone("SELECT * FROM submissions WHERE id=?", (sid,))
        if not r or r["review_status"] != "pending":
            await call.answer("এই submission আর pending নেই।", show_alert=True); return
        if q < 0 or q > r["valid_count"]:
            await call.answer("Invalid quantity.", show_alert=True); return
        amount = money(q * money(r["locked_rate"]))
        await DB.execute("UPDATE submissions SET review_status='approved',approved_quantity=?,final_amount=?,payment_status='unpaid',reviewed_at=? WHERE id=?", (q, float(amount), now_str(), sid))
        await DB.commit()
    await log_admin(call.from_user.id, "File approved - balance pending", str(r["user_id"]), amount, r["public_id"])
    await call.message.edit_text(f"✅ <b>File Approved</b>\n\n📄 #{r['public_id']}\n📊 Approved: {q}\n🧮 Calculated Amount: {CURRENCY}{fmt_money(amount)}\n\n💳 <b>Balance এখনো যোগ করা হয়নি।</b>\nAdmin পরে calculated বা custom amount add করতে পারবেন।")
    try:
        await call.bot.send_message(r["user_id"], f"✅ <b>আপনার ফাইল অনুমোদিত হয়েছে</b>\n\n📄 Submission: #{r['public_id']}\n📊 Approved: {q}\n🧮 হিসাবকৃত পরিমাণ: {CURRENCY}{fmt_money(amount)}\n\n💳 Balance এখনো যোগ করা হয়নি। Admin payment/balance processing সম্পন্ন করলে balance update হবে।\n⏱️ প্রয়োজন অনুযায়ী processing/report ১২–২৪ ঘণ্টার মধ্যে হতে পারে।", reply_markup=main_kb(await is_admin(r["user_id"])))
    except Exception: pass
    await call.answer("Approved — balance pending")

@router.callback_query(F.data.startswith("a:addsub:"))
async def add_submission_calculated(call: CallbackQuery):
    if not await admin_guard(call): return
    sid = int(call.data.split(":")[-1])
    async with DB_LOCK:
        r = await fetchone("SELECT * FROM submissions WHERE id=?", (sid,))
        if not r or r["review_status"] != "approved" or r["payment_status"] == "paid":
            await call.answer("এই submission-এর balance already processed বা eligible নয়।", show_alert=True); return
        amount = money(r["final_amount"])
        u = await fetchone("SELECT * FROM users WHERE telegram_id=?", (r["user_id"],))
        if not u: await call.answer("User পাওয়া যায়নি।", show_alert=True); return
        nb = money(u["balance"]) + amount
        txn = await next_public_id("transactions", "TXN")
        ts = now_str()
        await DB.execute("UPDATE users SET balance=?,total_earned=total_earned+?,updated_at=? WHERE telegram_id=?", (float(nb), float(amount), ts, r["user_id"]))
        await DB.execute("UPDATE submissions SET payment_status='paid',credited_amount=?,credited_by=?,credited_at=? WHERE id=? AND payment_status!='paid'", (float(amount), call.from_user.id, ts, sid))
        await DB.execute("INSERT INTO transactions(public_id,user_id,txn_type,amount,balance_after,reason,source_id,admin_id,created_at) VALUES(?,?,?,?,?,?,?,?,?)", (txn,r["user_id"],"SUBMISSION CREDIT",float(amount),float(nb),"Approved submission balance credit",r["public_id"],call.from_user.id,ts))
        await DB.commit()
    await log_admin(call.from_user.id,"Submission balance added",str(r["user_id"]),amount,r["public_id"])
    await call.message.edit_text(f"✅ <b>Balance Added</b>\n\n📄 #{r['public_id']}\n👤 UID: {r['user_id']}\n🧮 Calculated: {CURRENCY}{fmt_money(r['final_amount'])}\n💰 Credited: {CURRENCY}{fmt_money(amount)}\n🧾 TXN: {txn}")
    try: await call.bot.send_message(r["user_id"], f"💰 <b>আপনার balance update হয়েছে</b>\n\n📄 Submission: #{r['public_id']}\n➕ Added: {CURRENCY}{fmt_money(amount)}\n🧾 TXN: {txn}\n\nধন্যবাদ। 🤝")
    except Exception: pass
    await call.answer("Balance added")

@router.callback_query(F.data.startswith("a:customsub:"))
async def custom_submission_amount_start(call: CallbackQuery, state: FSMContext):
    if not await admin_guard(call): return
    sid = int(call.data.split(":")[-1]); r = await fetchone("SELECT * FROM submissions WHERE id=?", (sid,))
    if not r or r["review_status"] != "approved" or r["payment_status"] == "paid": await call.answer("এই submission আর payable নয়।", show_alert=True); return
    await state.set_state(AdminState.add_balance); await state.update_data(submission_credit_sid=sid)
    await send_or_edit(call, f"✏️ <b>CUSTOM BALANCE</b>\n\n📄 #{r['public_id']}\n👤 UID: {r['user_id']}\n🧮 Calculated: {CURRENCY}{fmt_money(r['final_amount'])}\n\nযে actual amount balance-এ যোগ করতে চান সেটি লিখুন।\nExample: 50", back_kb(f"a:sub:{sid}"))

@router.message(AdminState.add_balance)
async def add_balance_input(message: Message, state: FSMContext):
    data = await state.get_data()
    sid = data.get("submission_credit_sid")
    if sid:
        try: amount = money(message.text)
        except: amount = Decimal("0")
        r = await fetchone("SELECT * FROM submissions WHERE id=?", (sid,))
        if amount <= 0 or not r or r["review_status"] != "approved" or r["payment_status"] == "paid":
            await message.answer("❌ Amount valid নয় অথবা submission ইতোমধ্যে processed হয়েছে।"); return
        await state.clear()
        await set_setting(f"customsub_{message.from_user.id}_{sid}", str(amount))
        await message.answer(f"⚠️ <b>CONFIRM BALANCE</b>\n\n📄 #{r['public_id']}\n👤 UID: {r['user_id']}\n🧮 Calculated: {CURRENCY}{fmt_money(r['final_amount'])}\n💰 Actual Add: {CURRENCY}{fmt_money(amount)}\n\nCalculated amount-এর বদলে এই actual amount যোগ করবেন?", reply_markup=confirm_kb(f"a:customsubok:{sid}:{message.from_user.id}", f"a:sub:{sid}"))
        return
    # Existing manual user balance flow
    uid = data.get("target_uid")
    try: amount = money(message.text)
    except: amount = Decimal("0")
    if amount <= 0: await message.answer("Positive amount দিন।"); return
    u = await fetchone("SELECT balance FROM users WHERE telegram_id=?", (uid,))
    if not u: await message.answer("User not found"); await state.clear(); return
    await state.clear(); await set_setting(f"addbal_{message.from_user.id}_{uid}", str(amount))
    await message.answer(f"⚠️ CONFIRM ACTION\n\nUID: {uid}\nCurrent: {CURRENCY}{fmt_money(u['balance'])}\nAdd: +{CURRENCY}{fmt_money(amount)}\nReason: Manual adjustment", reply_markup=confirm_kb(f"a:addbalok:{uid}:{message.from_user.id}"))

@router.callback_query(F.data.startswith("a:customsubok:"))
async def custom_submission_amount_execute(call: CallbackQuery):
    if not await admin_guard(call): return
    _,_,sid_s,admin_s = call.data.split(":"); sid=int(sid_s); admin=int(admin_s)
    if admin != call.from_user.id: await call.answer("Invalid action.", show_alert=True); return
    amount=money(await get_setting(f"customsub_{admin}_{sid}","0"))
    if amount<=0: await call.answer("Action expired.",show_alert=True); return
    async with DB_LOCK:
        r=await fetchone("SELECT * FROM submissions WHERE id=?",(sid,))
        if not r or r["review_status"]!="approved" or r["payment_status"]=="paid": await call.answer("Already processed.",show_alert=True); return
        u=await fetchone("SELECT * FROM users WHERE telegram_id=?",(r["user_id"],))
        nb=money(u["balance"])+amount; txn=await next_public_id("transactions","TXN"); ts=now_str()
        await DB.execute("UPDATE users SET balance=?,total_earned=total_earned+?,updated_at=? WHERE telegram_id=?",(float(nb),float(amount),ts,r["user_id"]))
        await DB.execute("UPDATE submissions SET payment_status='paid',credited_amount=?,credited_by=?,credited_at=? WHERE id=? AND payment_status!='paid'",(float(amount),call.from_user.id,ts,sid))
        await DB.execute("INSERT INTO transactions(public_id,user_id,txn_type,amount,balance_after,reason,source_id,admin_id,created_at) VALUES(?,?,?,?,?,?,?,?,?)",(txn,r["user_id"],"SUBMISSION CUSTOM CREDIT",float(amount),float(nb),f"Custom credit; calculated={money(r['final_amount'])}",r["public_id"],call.from_user.id,ts)); await DB.commit()
    await log_admin(call.from_user.id,"Custom submission balance added",str(r["user_id"]),amount,r["public_id"])
    await call.message.edit_text(f"✅ <b>Custom Balance Added</b>\n\n📄 #{r['public_id']}\n🧮 Calculated: {CURRENCY}{fmt_money(r['final_amount'])}\n💰 Actual Credited: {CURRENCY}{fmt_money(amount)}\n🧾 TXN: {txn}")
    try: await call.bot.send_message(r["user_id"],f"💰 <b>Balance update হয়েছে</b>\n\n📄 Submission: #{r['public_id']}\n➕ Added: {CURRENCY}{fmt_money(amount)}\n🧾 TXN: {txn}\n\nAdmin-এর processing অনুযায়ী amount যোগ করা হয়েছে। 🤝")
    except Exception: pass
    await call.answer("Balance added")

@router.callback_query(F.data.startswith("a:ustats:"))
async def admin_user_daily_stats(call: CallbackQuery):
    if not await admin_guard(call): return
    uid=int(call.data.split(":")[-1]); today=datetime.now(TIMEZONE).strftime("%Y-%m-%d")
    u=await fetchone("SELECT * FROM users WHERE telegram_id=?",(uid,))
    d=await fetchone("SELECT COUNT(*) total, SUM(CASE WHEN review_status='approved' THEN 1 ELSE 0 END) approved, SUM(CASE WHEN review_status='rejected' THEN 1 ELSE 0 END) rejected, SUM(CASE WHEN review_status='pending' THEN 1 ELSE 0 END) pending, COALESCE(SUM(CASE WHEN review_status='approved' THEN final_amount ELSE 0 END),0) approved_amount, COALESCE(SUM(CASE WHEN review_status='approved' AND payment_status!='paid' THEN final_amount ELSE 0 END),0) unpaid_amount FROM submissions WHERE user_id=? AND substr(created_at,1,10)=?",(uid,today))
    await send_or_edit(call,f"📊 <b>USER DAILY STATS</b>\n\n👤 UID: {uid}\n📆 Date: {today}\n📥 Submitted: {d['total'] or 0}\n✅ Approved: {d['approved'] or 0}\n❌ Rejected: {d['rejected'] or 0}\n⏳ Pending: {d['pending'] or 0}\n🧮 Approved Total: {CURRENCY}{fmt_money(d['approved_amount'])}\n💳 Approved Unpaid: {CURRENCY}{fmt_money(d['unpaid_amount'])}\n💰 Current Balance: {CURRENCY}{fmt_money(u['balance'])}\n📌 Pending Balance: {CURRENCY}{fmt_money(u['pending_balance'])}",kb([[btn("📥 User Submissions",f"a:usersubs:{uid}")],[btn("◀️ User Profile",f"a:user:{uid}")]]))

@router.callback_query(F.data.startswith("a:reject:"))
async def admin_reject_prompt(call: CallbackQuery, state: FSMContext):
    if not await admin_guard(call): return
    sid = int(call.data.split(":")[-1]); r = await fetchone("SELECT * FROM submissions WHERE id=?", (sid,))
    if not r or r["review_status"] != "pending": await call.answer("Submission আর pending নেই।", show_alert=True); return
    await state.set_state(AdminState.reject_reason); await state.update_data(sid=sid)
    await send_or_edit(call, f"❌ REJECT SUBMISSION\n\n#{r['public_id']}\n\nReject করার কারণ লিখুন।", kb([[btn("❌ Cancel", f"a:sub:{sid}")]]))

@router.message(AdminState.reject_reason)
async def reject_reason(message: Message, state: FSMContext):
    data = await state.get_data(); sid = data.get("sid"); reason = (message.text or "").strip()[:1000]
    if not reason: await message.answer("Reject reason লিখুন।"); return
    r = await fetchone("SELECT * FROM submissions WHERE id=?", (sid,))
    await state.clear()
    if not r or r["review_status"] != "pending": await message.answer("Submission আর pending নেই।"); return
    await message.answer(f"⚠️ CONFIRM ACTION\n\nআপনি কি এই submission reject করতে চান?\n\n#{r['public_id']}\n📌 কারণ: {reason}", reply_markup=confirm_kb(f"a:rejectok:{sid}", f"a:sub:{sid}"))
    await state.update_data(reject_reason=reason)
    # Store reason temporarily in FSM is cleared above, so use a short database setting would be unsafe. Instead encode callback is impossible for long text; use per-admin temp setting.
    await set_setting(f"reject_reason_{message.from_user.id}_{sid}", reason)

@router.callback_query(F.data.startswith("a:rejectok:"))
async def reject_execute(call: CallbackQuery):
    if not await admin_guard(call): return
    sid = int(call.data.split(":")[-1]); reason = await get_setting(f"reject_reason_{call.from_user.id}_{sid}", "Admin rejection")
    async with DB_LOCK:
        r = await fetchone("SELECT * FROM submissions WHERE id=?", (sid,))
        if not r or r["review_status"] != "pending": await call.answer("Already processed.", show_alert=True); return
        await DB.execute("UPDATE submissions SET review_status='rejected',rejection_reason=?,reviewed_at=? WHERE id=?", (reason, now_str(), sid)); await DB.commit()
    await log_admin(call.from_user.id, "File rejected", str(r["user_id"]), None, r["public_id"])
    await call.message.edit_text(f"❌ Submission rejected\n\n#{r['public_id']}\n📌 কারণ: {reason}")
    try: await call.bot.send_message(r["user_id"], f"❌ ফাইল অনুমোদিত হয়নি\n\n📄 Submission: #{r['public_id']}\n📌 কারণ: {reason}\n\nপ্রয়োজনে সংশোধন করে নতুন file submit করুন। 🤝")
    except Exception: pass

# ============================================================
# 15. ADMIN PAYMENTS
# ============================================================
@router.callback_query(F.data == "a:payments")
async def admin_payments(call: CallbackQuery):
    if not await admin_guard(call): return
    p = await fetchone("SELECT COUNT(*) c,COALESCE(SUM(amount),0) s FROM withdrawals WHERE status='pending'")
    today = datetime.now(TIMEZONE).strftime("%Y-%m-%d")
    paid = await fetchone("SELECT COALESCE(SUM(amount),0) s FROM withdrawals WHERE status='paid' AND substr(processed_at,1,10)=?", (today,))
    await send_or_edit(call, f"💳 PAYMENT CENTER\n\n🟡 Pending: {p['c']}\n💰 Pending Amount: {CURRENCY}{fmt_money(p['s'])}\n✅ Paid Today: {CURRENCY}{fmt_money(paid['s'])}", kb([[btn("🟡 Pending Payments", "a:wdlist:pending")], [btn("✅ Payment History", "a:wdlist:paid")], [btn("🔎 Search Payment", "a:searchwd")], [btn("◀️ Back", "a:home")]]))

@router.callback_query(F.data.startswith("a:wdlist:"))
async def admin_wdlist(call: CallbackQuery):
    if not await admin_guard(call): return
    st = call.data.split(":")[-1]
    rows = await fetchall("SELECT id,public_id,user_id,amount,method,status FROM withdrawals WHERE status=? ORDER BY id DESC LIMIT 20", (st,))
    text = f"💳 {st.upper()} PAYMENTS\n\n" + ("\n".join(f"#{r['public_id']} • UID {r['user_id']} • {CURRENCY}{fmt_money(r['amount'])} • {r['method']}" for r in rows) if rows else "কোনো payment নেই।")
    buttons = [[btn(f"💳 #{r['public_id']}", f"a:wd:{r['public_id']}")] for r in rows]
    buttons.append([btn("◀️ Back", "a:payments")])
    await send_or_edit(call, text, kb(buttons))

@router.callback_query(F.data.startswith("a:wd:"))
async def admin_wd_detail(call: CallbackQuery):
    if not await admin_guard(call): return
    public_id = call.data.split(":", 2)[2]
    r = await fetchone("SELECT w.*,u.username FROM withdrawals w LEFT JOIN users u ON u.telegram_id=w.user_id WHERE w.public_id=?", (public_id,))
    if not r: await call.answer("Payment not found", show_alert=True); return
    text = f"💳 PAYMENT #{r['public_id']}\n\n👤 User: @{r['username'] or 'no_username'}\n🆔 UID: {r['user_id']}\n💰 Amount: {CURRENCY}{fmt_money(r['amount'])}\n💳 Method: {r['method']}\n📱 Number: {mask_number(r['number'])}\n🕐 Status: {r['status']}\n📌 Proof/Reference: {r['proof_reference'] or '-'}"
    buttons = [[btn("📄 Open Related Submission", "a:none")]]
    if r["status"] == "pending": buttons.append([btn("💸 Process Payment", f"a:paid:{r['id']}"), btn("❌ Reject", f"a:wdreject:{r['id']}")])
    elif r["status"] == "paid": buttons.append([btn("🔑 View TxID", "a:none")])
    buttons.append([btn("◀️ Back", "a:payments")])
    await send_or_edit(call, text, kb(buttons))

@router.callback_query(F.data.startswith("a:paid:"))
async def payment_paid_prompt(call: CallbackQuery, state: FSMContext):
    if not await admin_guard(call): return
    wid=int(call.data.split(":")[-1]); r=await fetchone("SELECT * FROM withdrawals WHERE id=?",(wid,))
    if not r or r["status"]!="pending": await call.answer("Already processed.",show_alert=True); return
    proof_chat=await get_setting("payment_proof_chat_id","")
    if not proof_chat:
        await call.answer("আগে Payment Proof Channel/Group select করুন।",show_alert=True); return
    await state.set_state(AdminState.payment_txid)
    await state.update_data(payment_wid=wid)
    await send_or_edit(call,f"💸 <b>PROCESS PAYMENT</b>\n\n#{r['public_id']}\n💰 Amount: {CURRENCY}{fmt_money(r['amount'])}\n💳 Method: {r['method']}\n📱 {mask_number(r['number'])}\n\n🔑 Admin যে <b>Transaction ID / TxID</b> দিয়ে payment করেছেন সেটি পাঠান।\n\n⚠️ TxID পুরোটা দিতে হবে।",kb([[btn("❌ Cancel","a:wd:"+r['public_id'])]]))

@router.message(AdminState.payment_txid)
async def payment_txid_input(message: Message, state: FSMContext):
    if not await admin_guard(message): return
    txid=(message.text or "").strip()
    if len(txid)<3 or len(txid)>200:
        await message.answer("❌ Valid Transaction ID পাঠান।"); return
    used=await fetchone("SELECT public_id FROM withdrawals WHERE proof_reference=? AND status='paid'",(txid,))
    if used:
        await message.answer(f"❌ এই TxID আগে ব্যবহার করা হয়েছে।\nWithdrawal: #{used['public_id']}\nঅন্য TxID দিন।"); return
    data=await state.get_data(); wid=data.get("payment_wid")
    r=await fetchone("SELECT * FROM withdrawals WHERE id=?",(wid,))
    if not r or r["status"]!="pending":
        await state.clear(); await message.answer("❌ এই withdrawal আর pending নেই।",reply_markup=back_kb("a:payments")); return
    await state.update_data(payment_txid=txid)
    await state.set_state(AdminState.payment_screenshot)
    await message.answer("📸 <b>Payment Screenshot</b> পাঠান।\n\nPhoto হিসেবে পাঠাতে পারেন।\nDocument হিসেবে image পাঠালেও গ্রহণ করা হবে।",reply_markup=kb([[btn("❌ Cancel",f"a:wd:{r['public_id']}")]]))

@router.message(AdminState.payment_screenshot)
async def payment_screenshot_input(message: Message, state: FSMContext):
    if not await admin_guard(message): return
    file_id=None; file_type=None
    if message.photo:
        file_id=message.photo[-1].file_id; file_type="photo"
    elif message.document and (message.document.mime_type or "").startswith("image/"):
        file_id=message.document.file_id; file_type="document"
    if not file_id:
        await message.answer("❌ শুধু payment screenshot-এর Photo বা Image Document পাঠান।"); return
    data=await state.get_data(); wid=data.get("payment_wid"); txid=data.get("payment_txid")
    r=await fetchone("SELECT w.*,u.username FROM withdrawals w LEFT JOIN users u ON u.telegram_id=w.user_id WHERE w.id=?",(wid,))
    if not r or r["status"]!="pending":
        await state.clear(); await message.answer("❌ এই withdrawal আর pending নেই।",reply_markup=back_kb("a:payments")); return
    await state.update_data(payment_screenshot_file_id=file_id,payment_screenshot_type=file_type)
    invoice=payment_invoice_text(r,txid,datetime.now(TIMEZONE),r["username"])
    preview=("👁️ <b>PAYMENT PROOF PREVIEW</b>\n\n"+invoice+"\n\n📸 Screenshot: <b>Attached</b>\n\n📢 Channel → Invoice only\n👤 User → Invoice + Screenshot")
    await message.answer(preview,reply_markup=kb([[btn("✅ Confirm & Send Proof",f"a:paidconfirm:{wid}","success")],[btn("❌ Cancel",f"a:wd:{r['public_id']}","danger")]]))

@router.callback_query(F.data.startswith("a:paidconfirm:"))
async def payment_paid_execute(call: CallbackQuery, state: FSMContext):
    if not await admin_guard(call): return
    wid=int(call.data.split(":")[-1]); data=await state.get_data()
    if int(data.get("payment_wid",-1))!=wid or not data.get("payment_txid") or not data.get("payment_screenshot_file_id"):
        await call.answer("TxID এবং Screenshot দুটোই দিন।",show_alert=True); return
    txid=data["payment_txid"]; screenshot_id=data["payment_screenshot_file_id"]; screenshot_type=data.get("payment_screenshot_type","photo")
    async with DB_LOCK:
        r=await fetchone("SELECT w.*,u.username FROM withdrawals w LEFT JOIN users u ON u.telegram_id=w.user_id WHERE w.id=?",(wid,))
        if not r or r["status"]!="pending": await call.answer("Already processed.",show_alert=True); return
        used=await fetchone("SELECT public_id FROM withdrawals WHERE proof_reference=? AND status='paid' AND id!=?",(txid,wid))
        if used: await call.answer("এই TxID আগে ব্যবহার করা হয়েছে।",show_alert=True); return
        processed=now_str()
        await DB.execute("UPDATE withdrawals SET status='paid',proof_reference=?,payment_screenshot_file_id=?,payment_screenshot_type=?,payment_proof_created_at=?,admin_id=?,processed_at=?,payment_proof_channel_sent=0,payment_proof_user_sent=0 WHERE id=?",(txid,screenshot_id,screenshot_type,processed,call.from_user.id,processed,wid))
        await DB.execute("UPDATE users SET total_paid=total_paid+?,updated_at=? WHERE telegram_id=?",(r["amount"],processed,r["user_id"]))
        await DB.commit()
    r=await fetchone("SELECT w.*,u.username FROM withdrawals w LEFT JOIN users u ON u.telegram_id=w.user_id WHERE w.id=?",(wid,))
    paid_dt=datetime.strptime(r["processed_at"],"%Y-%m-%d %H:%M:%S").replace(tzinfo=TIMEZONE)
    invoice=payment_invoice_text(r,txid,paid_dt,r["username"])
    channel_id=await get_setting("payment_proof_chat_id","")
    channel_ok=False; user_ok=False
    try:
        await call.bot.send_message(int(channel_id),invoice)
        channel_ok=True
    except Exception as exc:
        log.exception("Payment proof channel delivery failed: %s",exc)
    try:
        await call.bot.send_message(r["user_id"],invoice)
        if screenshot_type=="photo":
            await call.bot.send_photo(r["user_id"],screenshot_id,caption="📸 Payment Screenshot")
        else:
            await call.bot.send_document(r["user_id"],screenshot_id,caption="📸 Payment Screenshot")
        user_ok=True
    except Exception as exc:
        log.exception("Payment proof user delivery failed: %s",exc)
    await execute("UPDATE withdrawals SET payment_proof_channel_sent=?,payment_proof_user_sent=? WHERE id=?",(1 if channel_ok else 0,1 if user_ok else 0,wid))
    await log_admin(call.from_user.id,"Payment completed + proof sent",str(r["user_id"]),r["amount"],r["public_id"])
    status=("📢 Channel: ✅ Invoice sent" if channel_ok else "📢 Channel: ❌ Invoice failed")+"\n"+("👤 User: ✅ Invoice + Screenshot sent" if user_ok else "👤 User: ❌ Delivery failed")
    await state.clear()
    await call.message.edit_text(f"💳 <b>PAYMENT COMPLETED</b>\n\n🧾 #{r['public_id']}\n💰 {CURRENCY}{fmt_money(r['amount'])}\n🔑 TxID: <code>{html.escape(txid)}</code>\n\n{status}")

def payment_invoice_text(r, txid: str, paid_dt: datetime, username: str = None) -> str:
    uname=username or "no_username"
    return (
        "╭━━━〔 💳 <b>PAYMENT PROOF</b> 〕━━━╮\n\n"
        f"🧾 Withdrawal ID: <b>#{html.escape(str(r['public_id']))}</b>\n"
        f"👤 User: @{html.escape(uname)}\n"
        f"🆔 UID: <code>{r['user_id']}</code>\n\n"
        f"💰 Amount: <b>{CURRENCY}{fmt_money(r['amount'])}</b>\n"
        f"💳 Method: {html.escape(str(r['method']))}\n"
        f"📱 Number: <code>{html.escape(mask_number(r['number']))}</code>\n\n"
        f"🔑 Transaction ID: <code>{html.escape(txid)}</code>\n\n"
        f"🕐 Paid At: <b>{paid_dt.strftime('%d %b %Y • %I:%M %p')}</b> 🇧🇩\n"
        "👨‍💼 Paid By: <b>Admin</b>\n\n"
        "╰━━━━━━━━━━━━━━━━━━━━╯"
    )

@router.callback_query(F.data.startswith("a:wdreject:"))
async def wd_reject_prompt(call: CallbackQuery):
    if not await admin_guard(call): return
    wid = int(call.data.split(":")[-1]); r = await fetchone("SELECT * FROM withdrawals WHERE id=?", (wid,))
    if not r or r["status"] != "pending": await call.answer("Already processed.", show_alert=True); return
    await send_or_edit(call, f"⚠️ CONFIRM ACTION\n\nWithdrawal reject করলে {CURRENCY}{fmt_money(r['amount'])} user balance-এ ফেরত যাবে।\n\n#{r['public_id']}", confirm_kb(f"a:wdrejectok:{wid}"))

@router.callback_query(F.data.startswith("a:wdrejectok:"))
async def wd_reject_execute(call: CallbackQuery):
    if not await admin_guard(call): return
    wid = int(call.data.split(":")[-1])
    async with DB_LOCK:
        r = await fetchone("SELECT * FROM withdrawals WHERE id=?", (wid,))
        if not r or r["status"] != "pending": await call.answer("Already processed.", show_alert=True); return
        u = await fetchone("SELECT balance FROM users WHERE telegram_id=?", (r["user_id"],))
        new_balance = money(u["balance"]) + money(r["amount"])
        await DB.execute("UPDATE withdrawals SET status='rejected',admin_id=?,processed_at=? WHERE id=?", (call.from_user.id, now_str(), wid))
        await DB.execute("UPDATE users SET balance=?,updated_at=? WHERE telegram_id=?", (float(new_balance), now_str(), r["user_id"]))
        txn_public = await next_public_id("transactions", "TXN")
        await DB.execute("INSERT INTO transactions(public_id,user_id,txn_type,amount,balance_after,reason,source_id,admin_id,created_at) VALUES(?,?,?,?,?,?,?,?,?)", (txn_public,r["user_id"],"WITHDRAWAL REFUND",float(r["amount"]),float(new_balance),"Withdrawal rejected",r["public_id"],call.from_user.id,now_str()))
        await DB.commit()
    await log_admin(call.from_user.id, "Withdrawal rejected", str(r["user_id"]), r["amount"], r["public_id"])
    await call.message.edit_text(f"❌ Withdrawal rejected\n\n#{r['public_id']}\n💰 {CURRENCY}{fmt_money(r['amount'])} balance-এ ফেরত দেওয়া হয়েছে।")
    try: await call.bot.send_message(r["user_id"], f"❌ Withdrawal request reject হয়েছে।\n\n🧾 Request: #{r['public_id']}\n💰 {CURRENCY}{fmt_money(r['amount'])}\n\nAmount আপনার balance-এ ফেরত দেওয়া হয়েছে। 🤝")
    except Exception: pass

# ============================================================
# 16. BALANCE MANAGER
# ============================================================
@router.callback_query(F.data == "a:balance")
async def balance_manager(call: CallbackQuery):
    if not await admin_guard(call): return
    rows = await fetchall("SELECT s.id,s.public_id,s.user_id,s.category_name,s.final_amount,s.credited_amount,u.username FROM submissions s LEFT JOIN users u ON u.telegram_id=s.user_id WHERE s.review_status='approved' AND s.payment_status!='paid' ORDER BY s.id DESC LIMIT 30")
    lines = ["💰 <b>BALANCE MANAGER</b>", "", "Approved কিন্তু এখনো balance add হয়নি এমন files:"]
    buttons=[]
    if rows:
        for r in rows:
            lines.append(f"📄 #{r['public_id']} • UID {r['user_id']} • {CURRENCY}{fmt_money(r['final_amount'])}")
            buttons.append([btn(f"💰 #{r['public_id']} • {CURRENCY}{fmt_money(r['final_amount'])}",f"a:sub:{r['id']}")])
    else: lines.append("কোনো approved unpaid submission নেই।")
    buttons += [[btn("🔎 Search User", "a:searchuser")],[btn("◀️ Back", "a:home")]]
    await send_or_edit(call,"\n".join(lines),kb(buttons))

# ============================================================
# 16. USER MANAGER / BALANCE
# ============================================================
@router.callback_query(F.data == "a:users")
async def admin_users(call: CallbackQuery):
    if not await admin_guard(call): return
    total = await fetchone("SELECT COUNT(*) c FROM users"); active = await fetchone("SELECT COUNT(*) c FROM users WHERE banned=0"); banned = await fetchone("SELECT COUNT(*) c FROM users WHERE banned=1")
    await send_or_edit(call, f"👥 USER MANAGER\n\nTotal Users: {total['c']}\n🟢 Active: {active['c']}\n🔴 Banned: {banned['c']}", kb([[btn("🔎 Search User", "a:searchuser")], [btn("🟢 Active Users", "a:userlist:0"), btn("🔴 Banned Users", "a:userlist:1")], [btn("📊 User Statistics", "a:stats")], [btn("◀️ Back", "a:home")]]))

@router.callback_query(F.data.startswith("a:userlist:"))
async def user_list(call: CallbackQuery):
    if not await admin_guard(call): return
    banned = int(call.data.split(":")[-1]); rows = await fetchall("SELECT telegram_id,username,balance FROM users WHERE banned=? ORDER BY id DESC LIMIT 20", (banned,))
    buttons = [[btn(f"👤 {r['telegram_id']} • {CURRENCY}{fmt_money(r['balance'])}", f"a:user:{r['telegram_id']}")] for r in rows]
    buttons.append([btn("◀️ Back", "a:users")])
    await send_or_edit(call, "👥 USER LIST", kb(buttons))

@router.callback_query(F.data == "a:searchuser")
async def search_user_start(call: CallbackQuery, state: FSMContext):
    if not await admin_guard(call): return
    await state.set_state(AdminState.search_user); await send_or_edit(call, "🔎 User UID বা @username পাঠান।", back_kb("a:users"))

@router.message(AdminState.search_user)
async def search_user_input(message: Message, state: FSMContext):
    q = (message.text or "").strip().lstrip("@")
    row = None
    if q.isdigit(): row = await fetchone("SELECT * FROM users WHERE telegram_id=?", (int(q),))
    else: row = await fetchone("SELECT * FROM users WHERE lower(username)=lower(?)", (q,))
    await state.clear()
    if not row: await message.answer("❌ User পাওয়া যায়নি।", reply_markup=back_kb("a:users")); return
    await send_user_admin_profile(message, row)

async def send_user_admin_profile(message: Message, row):
    subs = await fetchone("SELECT COUNT(*) c FROM submissions WHERE user_id=?", (row["telegram_id"],)); appr = await fetchone("SELECT COUNT(*) c FROM submissions WHERE user_id=? AND review_status='approved'", (row["telegram_id"],)); paid = await fetchone("SELECT COALESCE(SUM(amount),0) s FROM withdrawals WHERE user_id=? AND status='paid'", (row["telegram_id"],))
    text = f"👤 USER PROFILE\n\nUID: {row['telegram_id']}\nUsername: @{row['username'] or 'none'}\n💰 Balance: {CURRENCY}{fmt_money(row['balance'])}\n📦 Submissions: {subs['c']}\n✅ Approved: {appr['c']}\n💳 Total Paid: {CURRENCY}{fmt_money(paid['s'])}"
    await message.answer(text, reply_markup=kb([[btn("💰 Add Balance", f"a:addbal:{row['telegram_id']}"), btn("➖ Deduct Balance", f"a:deduct:{row['telegram_id']}")], [btn("📥 Submissions", f"a:usersubs:{row['telegram_id']}"), btn("💳 Payments", f"a:userwd:{row['telegram_id']}")], [btn("🚫 Ban User", f"a:ban:{row['telegram_id']}")], [btn("◀️ Back", "a:users")]]))

@router.callback_query(F.data.startswith("a:user:"))
async def admin_user_detail(call: CallbackQuery):
    if not await admin_guard(call): return
    uid = int(call.data.split(":")[-1]); row = await fetchone("SELECT * FROM users WHERE telegram_id=?", (uid,))
    if not row: await call.answer("User not found", show_alert=True); return
    subs = await fetchone("SELECT COUNT(*) c FROM submissions WHERE user_id=?", (uid,)); appr = await fetchone("SELECT COUNT(*) c FROM submissions WHERE user_id=? AND review_status='approved'", (uid,)); paid = await fetchone("SELECT COALESCE(SUM(amount),0) s FROM withdrawals WHERE user_id=? AND status='paid'", (uid,))
    await send_or_edit(call, f"👤 USER PROFILE\n\nUID: {uid}\nUsername: @{row['username'] or 'none'}\n💰 Balance: {CURRENCY}{fmt_money(row['balance'])}\n📦 Submissions: {subs['c']}\n✅ Approved: {appr['c']}\n💳 Total Paid: {CURRENCY}{fmt_money(paid['s'])}", kb([[btn("💰 Add Balance", f"a:addbal:{uid}"), btn("➖ Deduct Balance", f"a:deduct:{uid}")], [btn("📥 Submissions", f"a:usersubs:{uid}"), btn("💳 Payments", f"a:userwd:{uid}")], [btn("🚫 Ban User", f"a:ban:{uid}"), btn("🟢 Unban User", f"a:unban:{uid}")], [btn("◀️ Back", "a:users")]]))

@router.callback_query(F.data.startswith("a:addbal:"))
async def add_balance_start(call: CallbackQuery, state: FSMContext):
    if not await admin_guard(call): return
    uid = int(call.data.split(":")[-1]); await state.set_state(AdminState.add_balance); await state.update_data(target_uid=uid)
    await send_or_edit(call, f"➕ ADD BALANCE\n\n👤 UID: {uid}\n\nAmount পাঠান। Example: 500", back_kb(f"a:user:{uid}"))

@router.callback_query(F.data.startswith("a:addbalok:"))
async def add_balance_execute(call: CallbackQuery):
    if not await admin_guard(call): return
    _,_,uid_s,admin_s = call.data.split(":"); uid=int(uid_s); admin=int(admin_s)
    if admin != call.from_user.id: await call.answer("Invalid action.", show_alert=True); return
    amount=money(await get_setting(f"addbal_{admin}_{uid}", "0"));
    if amount<=0: await call.answer("Action expired.", show_alert=True); return
    async with DB_LOCK:
        u=await fetchone("SELECT balance FROM users WHERE telegram_id=?",(uid,))
        if not u: await call.answer("User not found", show_alert=True); return
        nb=money(u["balance"])+amount; txn=await next_public_id("transactions","TXN")
        await DB.execute("UPDATE users SET balance=?,total_earned=total_earned+?,updated_at=? WHERE telegram_id=?",(float(nb),float(amount),now_str(),uid))
        await DB.execute("INSERT INTO transactions(public_id,user_id,txn_type,amount,balance_after,reason,admin_id,created_at) VALUES(?,?,?,?,?,?,?,?)",(txn,uid,"MANUAL CREDIT",float(amount),float(nb),"Manual adjustment",call.from_user.id,now_str())); await DB.commit()
    await log_admin(call.from_user.id,"Balance added",str(uid),amount,txn); await call.message.edit_text(f"✅ Balance added\n\nUID: {uid}\n+{CURRENCY}{fmt_money(amount)}\nTXN: {txn}")
    try: await call.bot.send_message(uid,f"💰 আপনার balance-এ {CURRENCY}{fmt_money(amount)} যোগ করা হয়েছে।\n🧾 {txn} 🤝")
    except: pass

@router.callback_query(F.data.startswith("a:deduct:"))
async def deduct_start(call: CallbackQuery, state: FSMContext):
    if not await admin_guard(call): return
    uid=int(call.data.split(":")[-1]); await state.set_state(AdminState.deduct_balance); await state.update_data(target_uid=uid)
    await send_or_edit(call, f"➖ DEDUCT BALANCE\n\nUID: {uid}\nAmount পাঠান।", back_kb(f"a:user:{uid}"))

@router.message(AdminState.deduct_balance)
async def deduct_input(message: Message, state: FSMContext):
    data=await state.get_data(); uid=data.get("target_uid")
    try: amount=money(message.text)
    except: amount=Decimal("0")
    u=await fetchone("SELECT balance FROM users WHERE telegram_id=?",(uid,))
    if amount<=0 or not u or amount>money(u["balance"]): await message.answer("Amount valid নয় বা balance-এর চেয়ে বেশি।"); return
    await state.clear(); await set_setting(f"deduct_{message.from_user.id}_{uid}",str(amount))
    await message.answer(f"⚠️ CONFIRM ACTION\n\nUID: {uid}\nCurrent: {CURRENCY}{fmt_money(u['balance'])}\nDeduct: -{CURRENCY}{fmt_money(amount)}", reply_markup=confirm_kb(f"a:deductok:{uid}:{message.from_user.id}"))

@router.callback_query(F.data.startswith("a:deductok:"))
async def deduct_execute(call: CallbackQuery):
    if not await admin_guard(call): return
    _,_,uid_s,admin_s=call.data.split(":"); uid=int(uid_s); admin=int(admin_s)
    amount=money(await get_setting(f"deduct_{admin}_{uid}","0"))
    if amount<=0: await call.answer("Action expired.",show_alert=True); return
    async with DB_LOCK:
        u=await fetchone("SELECT balance FROM users WHERE telegram_id=?",(uid,))
        if not u or amount>money(u["balance"]): await call.answer("Balance changed; action cancelled.",show_alert=True); return
        nb=money(u["balance"])-amount; txn=await next_public_id("transactions","TXN")
        await DB.execute("UPDATE users SET balance=?,updated_at=? WHERE telegram_id=?",(float(nb),now_str(),uid))
        await DB.execute("INSERT INTO transactions(public_id,user_id,txn_type,amount,balance_after,reason,admin_id,created_at) VALUES(?,?,?,?,?,?,?,?)",(txn,uid,"MANUAL DEBIT",float(-amount),float(nb),"Manual adjustment",call.from_user.id,now_str())); await DB.commit()
    await log_admin(call.from_user.id,"Balance deducted",str(uid),amount,txn); await call.message.edit_text(f"✅ Balance deducted\n\nUID: {uid}\n-{CURRENCY}{fmt_money(amount)}\nTXN: {txn}")

# ============================================================
# 17. BAN / USER SUBS/PAYMENTS
# ============================================================
# Unban is kept separate so banned users can be restored without touching balances/history.
@router.callback_query(F.data.startswith("a:unban:"))
async def unban_prompt(call: CallbackQuery):
    if not await admin_guard(call): return
    uid = int(call.data.split(":")[-1])
    await send_or_edit(call, f"⚠️ CONFIRM ACTION\n\nআপনি কি UID {uid}-কে unban করতে চান?", confirm_kb(f"a:unbanok:{uid}"))

@router.callback_query(F.data.startswith("a:unbanok:"))
async def unban_execute(call: CallbackQuery):
    if not await admin_guard(call): return
    uid = int(call.data.split(":")[-1])
    await execute("UPDATE users SET banned=0,updated_at=? WHERE telegram_id=?", (now_str(), uid))
    await log_admin(call.from_user.id, "User unbanned", str(uid))
    await call.message.edit_text(f"🟢 User {uid} unbanned.")

@router.callback_query(F.data.startswith("a:ban:"))
async def ban_prompt(call: CallbackQuery):
    if not await admin_guard(call): return
    uid=int(call.data.split(":")[-1])
    if uid==OWNER_ID: await call.answer("Owner protected.",show_alert=True); return
    await send_or_edit(call,f"⚠️ CONFIRM ACTION\n\nআপনি কি UID {uid}-কে ban করতে চান?",confirm_kb(f"a:banok:{uid}"))

@router.callback_query(F.data.startswith("a:banok:"))
async def ban_execute(call: CallbackQuery):
    if not await admin_guard(call): return
    uid=int(call.data.split(":")[-1])
    if uid==OWNER_ID: await call.answer("Owner protected.",show_alert=True); return
    await execute("UPDATE users SET banned=1,updated_at=? WHERE telegram_id=?",(now_str(),uid)); await log_admin(call.from_user.id,"User banned",str(uid)); await call.message.edit_text(f"🚫 User {uid} banned.")

@router.callback_query(F.data.startswith("a:usersubs:"))
async def user_subs_admin(call: CallbackQuery):
    if not await admin_guard(call): return
    uid=int(call.data.split(":")[-1]); rows=await fetchall("SELECT id,public_id,category_name,review_status,final_amount FROM submissions WHERE user_id=? ORDER BY id DESC LIMIT 15",(uid,))
    await send_or_edit(call,"📥 USER SUBMISSIONS\n\n"+(("\n".join(f"#{r['public_id']} • {r['category_name']} • {r['review_status']} • {CURRENCY}{fmt_money(r['final_amount'])}" for r in rows)) if rows else "কোনো submission নেই।"),back_kb(f"a:user:{uid}"))

@router.callback_query(F.data.startswith("a:userwd:"))
async def user_wd_admin(call: CallbackQuery):
    if not await admin_guard(call): return
    uid=int(call.data.split(":")[-1]); rows=await fetchall("SELECT public_id,amount,method,status FROM withdrawals WHERE user_id=? ORDER BY id DESC LIMIT 15",(uid,))
    await send_or_edit(call,"💳 USER PAYMENTS\n\n"+(("\n".join(f"#{r['public_id']} • {CURRENCY}{fmt_money(r['amount'])} • {r['method']} • {r['status']}" for r in rows)) if rows else "কোনো payment নেই।"),back_kb(f"a:user:{uid}"))

# ============================================================
# 18. CATEGORY MANAGER
# ============================================================
@router.callback_query(F.data == "a:categories")
async def category_manager(call: CallbackQuery):
    if not await admin_guard(call): return
    active=await fetchone("SELECT COUNT(*) c FROM categories WHERE enabled=1"); dis=await fetchone("SELECT COUNT(*) c FROM categories WHERE enabled=0")
    await send_or_edit(call,f"📦 CATEGORY MANAGER\n\n🟢 Active: {active['c']}\n🔴 Disabled: {dis['c']}",kb([[btn("➕ Add Category","a:addcat")],[btn("⚙️ Manage Categories","a:catlist"),btn("🔴 Disabled","a:catlist:disabled")],[btn("◀️ Back","a:home")]]))

@router.callback_query(F.data == "a:addcat")
async def add_category_start(call: CallbackQuery,state:FSMContext):
    if not await admin_guard(call): return
    await state.set_state(AdminState.add_category_name); await send_or_edit(call,"➕ ADD CATEGORY\n\nCategory name লিখুন।",back_kb("a:categories"))

@router.message(AdminState.add_category_name)
async def add_category_name(message:Message,state:FSMContext):
    name=(message.text or "").strip()[:80]
    if not name: await message.answer("Category name দিন।"); return
    if await fetchone("SELECT id FROM categories WHERE lower(name)=lower(?)",(name,)): await message.answer("এই category already exists।"); return
    await state.update_data(cat_name=name); await state.set_state(AdminState.add_category_rate); await message.answer("💵 Rate লিখুন। Example: 4.80")

@router.message(AdminState.add_category_rate)
async def add_category_rate(message:Message,state:FSMContext):
    try: rate=money(message.text)
    except: rate=Decimal("-1")
    data=await state.get_data(); name=data.get("cat_name")
    if rate<0: await message.answer("Valid rate দিন।"); return
    next_order = await fetchone("SELECT COALESCE(MAX(sort_order),0)+1 n FROM categories")
    cur=await execute("INSERT INTO categories(name,rate,enabled,emoji,button_style,sort_order,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?)",(name,float(rate),1,"📦","primary",int(next_order["n"]),now_str(),now_str())); cid=cur.lastrowid
    await execute("INSERT INTO formats(category_id,columns_json,updated_at) VALUES(?,?,?)",(cid,json.dumps(["UID","Status","Notes"]),now_str()))
    await log_admin(message.from_user.id,"Category added",name,rate); await state.clear(); await message.answer(f"✅ Category তৈরি হয়েছে।\n\n📦 {name}\n💰 Rate: {CURRENCY}{fmt_money(rate)}",reply_markup=back_kb("a:categories"))

@router.callback_query(F.data == "a:catlist")
async def catlist(call:CallbackQuery): await show_catlist(call,True)
@router.callback_query(F.data == "a:catlist:disabled")
async def catlist_disabled(call:CallbackQuery): await show_catlist(call,False)
async def show_catlist(call,enabled):
    if not await admin_guard(call): return
    rows=await fetchall("SELECT id,name,rate,enabled,emoji,button_style,sort_order FROM categories WHERE enabled=? ORDER BY sort_order ASC,id ASC",(1 if enabled else 0,))
    buttons=[[btn(f"{r['emoji'] or '📦'} {r['name']} • {CURRENCY}{fmt_money(r['rate'])} • {'🟢 ON' if r['enabled'] else '⚪ OFF'}",f"a:cat:{r['id']}",r['button_style'] if r['enabled'] and r['button_style'] in {'primary','success','danger'} else 'primary')] for r in rows]
    buttons.append([btn("◀️ Back","a:categories")]); await send_or_edit(call,"📦 CATEGORIES\n\n🟢 ON = colored button shown to users\n⚪ OFF = hidden from user Submit File menu",kb(buttons))

@router.callback_query(F.data.startswith("a:cat:"))
async def cat_detail(call:CallbackQuery):
    if not await admin_guard(call): return
    cid=int(call.data.split(":")[-1]); r=await get_category(cid)
    if not r: await call.answer("Category not found",show_alert=True); return
    action="🔴 Disable" if r["enabled"] else "🟢 Enable"
    style=r["button_style"] if r["button_style"] in {"primary","success","danger"} else "primary"
    await send_or_edit(call,f"📦 {r['emoji'] or '📦'} {r['name']}\n\n💰 Rate: {CURRENCY}{fmt_money(r['rate'])}\n🎨 Button Style: {style.title()}\n🔢 Position: {r['sort_order']}\n🟢 Status: {'Enabled' if r['enabled'] else 'Disabled'}\n\n<b>User view:</b> {'রঙিন button হিসেবে দেখাবে' if r['enabled'] else 'Submit File menu-তে দেখাবে না'}",kb([[btn("💵 Change Rate",f"a:rate:{cid}"),btn(action,f"a:togglecat:{cid}")],[btn("🎨 Design",f"a:design:{cid}","success"),btn("✏️ Rename",f"a:rename:{cid}")],[btn("🔢 Position",f"a:order:{cid}"),btn("🗑 Delete",f"a:delcat:{cid}","danger")],[btn("◀️ Back","a:categories")]]))

@router.callback_query(F.data.startswith("a:rate:"))
async def change_rate_start(call:CallbackQuery,state:FSMContext):
    if not await admin_guard(call): return
    cid=int(call.data.split(":")[-1]); await state.set_state(AdminState.change_rate); await state.update_data(cid=cid); await send_or_edit(call,"💵 নতুন rate লিখুন।",back_kb(f"a:cat:{cid}"))
@router.message(AdminState.change_rate)
async def change_rate_input(message:Message,state:FSMContext):
    try: rate=money(message.text)
    except: rate=Decimal("-1")
    data=await state.get_data(); cid=data.get("cid")
    if rate<0: await message.answer("Valid rate দিন।"); return
    await execute("UPDATE categories SET rate=?,updated_at=? WHERE id=?",(float(rate),now_str(),cid)); await log_admin(message.from_user.id,"Category rate changed",str(cid),rate); await state.clear(); await message.answer(f"✅ Rate updated: {CURRENCY}{fmt_money(rate)}",reply_markup=back_kb("a:categories"))

@router.callback_query(F.data.startswith("a:togglecat:"))
async def toggle_category_prompt(call:CallbackQuery):
    if not await admin_guard(call): return
    cid=int(call.data.split(":")[-1]); r=await get_category(cid); target="enable" if not r["enabled"] else "disable"; await send_or_edit(call,f"⚠️ CONFIRM ACTION\n\nCategory {target} করবেন?\n\n📦 {r['name']}",confirm_kb(f"a:togglecatok:{cid}"))
@router.callback_query(F.data.startswith("a:togglecatok:"))
async def toggle_category_execute(call:CallbackQuery):
    if not await admin_guard(call): return
    cid=int(call.data.split(":")[-1]); r=await get_category(cid); new=0 if r["enabled"] else 1; await execute("UPDATE categories SET enabled=?,updated_at=? WHERE id=?",(new,now_str(),cid)); await log_admin(call.from_user.id,"Category enabled" if new else "Category disabled",r["name"]); await call.message.edit_text(f"✅ Category {'enabled' if new else 'disabled'}: {r['name']}")

@router.callback_query(F.data.startswith("a:rename:"))
async def rename_start(call:CallbackQuery,state:FSMContext):
    if not await admin_guard(call): return
    cid=int(call.data.split(":")[-1]); await state.set_state(AdminState.rename_category); await state.update_data(cid=cid); await send_or_edit(call,"✏️ নতুন category name লিখুন।",back_kb(f"a:cat:{cid}"))
@router.message(AdminState.rename_category)
async def rename_input(message:Message,state:FSMContext):
    name=(message.text or "").strip()[:80]; data=await state.get_data(); cid=data.get("cid")
    if not name: await message.answer("Name দিন。"); return
    try:
        await execute("UPDATE categories SET name=?,updated_at=? WHERE id=?",(name,now_str(),cid)); await log_admin(message.from_user.id,"Category renamed",name); await state.clear(); await message.answer("✅ Category renamed.",reply_markup=back_kb("a:categories"))
    except Exception: await message.answer("❌ এই name already used হতে পারে।")

@router.callback_query(F.data.startswith("a:delcat:"))
async def delete_cat_prompt(call:CallbackQuery):
    if not await admin_guard(call): return
    cid=int(call.data.split(":")[-1]); r=await get_category(cid); await send_or_edit(call,f"⚠️ CONFIRM ACTION\n\nআপনি কি category delete করতে চান?\n\n📦 {r['name']}",confirm_kb(f"a:delcatok:{cid}"))
@router.callback_query(F.data.startswith("a:delcatok:"))
async def delete_cat_execute(call:CallbackQuery):
    if not await admin_guard(call): return
    cid=int(call.data.split(":")[-1]); r=await get_category(cid)
    used=await fetchone("SELECT COUNT(*) c FROM submissions WHERE category_id=?",(cid,))
    if used and used["c"]>0: await call.answer("এই category-এর historical submissions আছে; delete করা যাবে না। Disable করুন।",show_alert=True); return
    await execute("DELETE FROM categories WHERE id=?",(cid,)); await log_admin(call.from_user.id,"Category deleted",r["name"]); await call.message.edit_text("🗑 Category deleted.")

# ============================================================
# 19. CATEGORY DESIGN / ORDER
# ============================================================
@router.callback_query(F.data.startswith("a:design:"))
async def category_design(call: CallbackQuery):
    if not await admin_guard(call): return
    cid=int(call.data.split(":")[-1]); r=await get_category(cid)
    if not r: await call.answer("Category not found",show_alert=True); return
    await send_or_edit(call, f"🎨 CATEGORY DESIGN\n\n{r['emoji'] or '📦'} <b>{r['name']}</b>\n\nCurrent style: <b>{r['button_style']}</b>\nCurrent emoji: <b>{r['emoji'] or '📦'}</b>\n\n🟦 Blue = Primary\n🟩 Green = Success\n🟥 Red = Danger", kb([[btn("🟦 Blue",f"a:style:{cid}:primary","primary"),btn("🟩 Green",f"a:style:{cid}:success","success")],[btn("🟥 Red",f"a:style:{cid}:danger","danger"),btn("🙂 Change Emoji",f"a:emoji:{cid}")],[btn("◀️ Back",f"a:cat:{cid}")]]))

@router.callback_query(F.data.startswith("a:style:"))
async def category_style(call: CallbackQuery):
    if not await admin_guard(call): return
    parts=call.data.split(":"); cid=int(parts[2]); style=parts[3]
    if style not in {"primary","success","danger"}: return
    r=await get_category(cid)
    await execute("UPDATE categories SET button_style=?,updated_at=? WHERE id=?",(style,now_str(),cid))
    await log_admin(call.from_user.id,"Category button style changed",r["name"],None,style)
    await call.answer("Design updated")
    await category_design(call)

@router.callback_query(F.data.startswith("a:emoji:"))
async def category_emoji_start(call: CallbackQuery,state:FSMContext):
    if not await admin_guard(call): return
    cid=int(call.data.split(":")[-1]); await state.set_state(AdminState.category_emoji); await state.update_data(cid=cid)
    await send_or_edit(call,"🙂 CATEGORY EMOJI\n\nএকটি emoji পাঠান। Example: 📘 / 📸 / ✉️ / 🔥",back_kb(f"a:design:{cid}"))

@router.message(AdminState.category_emoji)
async def category_emoji_input(message: Message,state:FSMContext):
    emoji=(message.text or "").strip()
    data=await state.get_data(); cid=data.get("cid")
    if not emoji or len(emoji)>8: await message.answer("❌ একটি ছোট emoji দিন।"); return
    await execute("UPDATE categories SET emoji=?,updated_at=? WHERE id=?",(emoji,now_str(),cid)); await log_admin(message.from_user.id,"Category emoji changed",str(cid)); await state.clear()
    await message.answer("✅ Category emoji update হয়েছে।",reply_markup=back_kb(f"a:cat:{cid}"))

@router.callback_query(F.data.startswith("a:order:"))
async def category_order_start(call: CallbackQuery,state:FSMContext):
    if not await admin_guard(call): return
    cid=int(call.data.split(":")[-1]); await state.set_state(AdminState.category_order); await state.update_data(cid=cid)
    await send_or_edit(call,"🔢 CATEGORY POSITION\n\n1, 2, 3... এর মতো position number পাঠান।\nছোট number আগে দেখাবে।",back_kb(f"a:cat:{cid}"))

@router.message(AdminState.category_order)
async def category_order_input(message: Message,state:FSMContext):
    try: order=int((message.text or "").strip())
    except: order=0
    data=await state.get_data(); cid=data.get("cid")
    if order<1 or order>999: await message.answer("❌ 1 থেকে 999-এর মধ্যে position দিন।"); return
    await execute("UPDATE categories SET sort_order=?,updated_at=? WHERE id=?",(order,now_str(),cid)); await log_admin(message.from_user.id,"Category position changed",str(cid),order); await state.clear()
    await message.answer(f"✅ Position set: {order}",reply_markup=back_kb(f"a:cat:{cid}"))

# ============================================================
# 19. FORMAT MANAGER
# ============================================================
@router.callback_query(F.data == "a:formats")
async def format_manager(call:CallbackQuery):
    if not await admin_guard(call): return
    rows=await fetchall("SELECT id,name FROM categories ORDER BY id")
    buttons=[[btn(f"📦 {r['name']}",f"a:format:{r['id']}")] for r in rows]; buttons.append([btn("➕ New Format","a:newformat")]); buttons.append([btn("◀️ Back","a:home")])
    await send_or_edit(call,"📋 FORMAT MANAGER\n\nকোন Category-এর format পরিবর্তন করবেন?",kb(buttons))

@router.callback_query(F.data.startswith("a:format:"))
async def format_settings(call:CallbackQuery):
    if not await admin_guard(call): return
    cid=int(call.data.split(":")[-1]); c=await get_category(cid); cols=await get_format(cid)
    await send_or_edit(call,f"📋 FORMAT SETTINGS\n\nCategory: {c['name']}\n\nCurrent Format:\n{' | '.join(cols)}",kb([[btn("✏️ Edit Format",f"a:editformat:{cid}"),btn("📄 Preview",f"a:previewformat:{cid}")],[btn("📥 Create Template",f"u:template:{cid}"),btn("🔄 Reset",f"a:resetformat:{cid}")],[btn("◀️ Back","a:formats")]]))

@router.callback_query(F.data.startswith("a:editformat:"))
async def edit_format_start(call:CallbackQuery,state:FSMContext):
    if not await admin_guard(call): return
    cid=int(call.data.split(":")[-1]); await state.set_state(AdminState.edit_format); await state.update_data(cid=cid); await send_or_edit(call,"✏️ FORMAT EDITOR\n\nনতুন format লিখুন।\nExample: UID | Status | Notes",back_kb(f"a:format:{cid}"))
@router.message(AdminState.edit_format)
async def edit_format_input(message:Message,state:FSMContext):
    cols=parse_columns(message.text or ""); ok,reason=valid_format_columns(cols)
    if not ok: await message.answer("❌ "+reason); return
    data=await state.get_data(); await state.update_data(cols=cols); await state.set_state(AdminState.edit_format); await set_setting(f"formatdraft_{message.from_user.id}",json.dumps({"cid":data['cid'],"cols":cols}))
    await message.answer("🔎 FORMAT PREVIEW\n\n"+"\n".join(f"{i+1}️⃣ {c}" for i,c in enumerate(cols)),reply_markup=kb([[btn("✅ Save Format",f"a:saveformat:{data['cid']}"),btn("✏️ Edit Again",f"a:editformat:{data['cid']}")],[btn("❌ Cancel",f"a:format:{data['cid']}")]]))
@router.callback_query(F.data.startswith("a:saveformat:"))
async def save_format(call:CallbackQuery,state:FSMContext):
    if not await admin_guard(call): return
    cid=int(call.data.split(":")[-1]); raw=await get_setting(f"formatdraft_{call.from_user.id}","")
    try: cols=json.loads(raw)["cols"]
    except: await call.answer("Draft expired.",show_alert=True); return
    await execute("INSERT INTO formats(category_id,columns_json,updated_at) VALUES(?,?,?) ON CONFLICT(category_id) DO UPDATE SET columns_json=excluded.columns_json,updated_at=excluded.updated_at",(cid,json.dumps(cols,ensure_ascii=False),now_str())); await log_admin(call.from_user.id,"Format updated",str(cid)); await state.clear(); await call.message.edit_text("✅ Format saved.",reply_markup=back_kb(f"a:format:{cid}"))
@router.callback_query(F.data.startswith("a:previewformat:"))
async def preview_format(call:CallbackQuery):
    if not await admin_guard(call): return
    cid=int(call.data.split(":")[-1]); cols=await get_format(cid); await send_or_edit(call,"📄 REQUIRED EXCEL FORMAT\n\n"+"\n".join(f"Column {i+1} → {c}" for i,c in enumerate(cols)),kb([[btn("📥 Create Template",f"u:template:{cid}")],[btn("◀️ Back",f"a:format:{cid}")]]))
@router.callback_query(F.data.startswith("a:resetformat:"))
async def reset_format_prompt(call:CallbackQuery):
    if not await admin_guard(call): return
    cid=int(call.data.split(":")[-1]); await send_or_edit(call,"⚠️ CONFIRM ACTION\n\nFormat UID | Status | Notes-এ reset করবেন?",confirm_kb(f"a:resetformatok:{cid}"))
@router.callback_query(F.data.startswith("a:resetformatok:"))
async def reset_format_execute(call:CallbackQuery):
    if not await admin_guard(call): return
    cid=int(call.data.split(":")[-1]); await execute("INSERT INTO formats(category_id,columns_json,updated_at) VALUES(?,?,?) ON CONFLICT(category_id) DO UPDATE SET columns_json=excluded.columns_json,updated_at=excluded.updated_at",(cid,json.dumps(["UID","Status","Notes"]),now_str())); await log_admin(call.from_user.id,"Format reset",str(cid)); await call.message.edit_text("✅ Format reset হয়েছে।")
@router.callback_query(F.data == "a:newformat")
async def new_format_info(call:CallbackQuery):
    if not await admin_guard(call): return
    await send_or_edit(call,"📋 New Format\n\nপ্রথমে একটি category তৈরি করুন, তারপর সেই category-এর Format Manager থেকে format edit করুন।",back_kb("a:formats"))

# ============================================================
# 20. STATISTICS / LOGS
# ============================================================
@router.callback_query(F.data == "a:stats")
async def admin_stats(call:CallbackQuery):
    if not await admin_guard(call): return
    await show_stats(call,"all")
@router.callback_query(F.data.startswith("a:stats:"))
async def admin_stats_period(call:CallbackQuery):
    if not await admin_guard(call): return
    await show_stats(call,call.data.split(":")[-1])
async def show_stats(call,period):
    where=""; params=()
    if period=="today": where=" AND substr(created_at,1,10)=?"; params=(datetime.now(TIMEZONE).strftime("%Y-%m-%d"),)
    elif period=="week": where=" AND datetime(created_at)>=datetime(?)"; params=((datetime.now(TIMEZONE)-timedelta(days=7)).strftime("%Y-%m-%d %H:%M:%S"),)
    elif period=="month": where=" AND datetime(created_at)>=datetime(?)"; params=((datetime.now(TIMEZONE)-timedelta(days=30)).strftime("%Y-%m-%d %H:%M:%S"),)
    users=await fetchone("SELECT COUNT(*) c FROM users")
    subs=await fetchone("SELECT COUNT(*) c FROM submissions WHERE 1=1"+where,params); appr=await fetchone("SELECT COUNT(*) c FROM submissions WHERE review_status='approved'"+where,params); rej=await fetchone("SELECT COUNT(*) c FROM submissions WHERE review_status='rejected'"+where,params); amt=await fetchone("SELECT COALESCE(SUM(final_amount),0) s FROM submissions WHERE review_status='approved'"+where,params); paid=await fetchone("SELECT COALESCE(SUM(amount),0) s FROM withdrawals WHERE status='paid'"+where,params)
    text=f"📈 STATISTICS\n\n👥 Total Users: {users['c']}\n\n📦 Total Submissions: {subs['c']}\n✅ Approved: {appr['c']}\n❌ Rejected: {rej['c']}\n\n💰 Total Approved Amount: {CURRENCY}{fmt_money(amt['s'])}\n💳 Total Paid: {CURRENCY}{fmt_money(paid['s'])}"
    await send_or_edit(call,text,kb([[btn("📅 Today","a:stats:today"),btn("📅 This Week","a:stats:week")],[btn("📅 This Month","a:stats:month")],[btn("◀️ Back","a:home")]]))

@router.callback_query(F.data == "a:logs")
async def admin_logs(call:CallbackQuery):
    if not await admin_guard(call): return
    rows=await fetchall("SELECT admin_id,action,target,amount,reference_id,created_at FROM admin_logs ORDER BY id DESC LIMIT 30")
    text="📝 ADMIN LOGS\n\n"+(("\n".join(f"{r['created_at'][11:16]} — {r['action']} — {r['target'] or '-'} — {r['reference_id'] or '-'}" for r in rows)) if rows else "কোনো log নেই।")
    await send_or_edit(call,text,kb([[btn("◀️ Back","a:home")]]))

# ============================================================
# 21. BROADCAST
# ============================================================
@router.callback_query(F.data == "a:broadcast")
async def broadcast_home(call:CallbackQuery):
    if not await admin_guard(call): return
    await send_or_edit(call,"📢 BROADCAST",kb([[btn("✉️ Create Broadcast","a:createbroadcast"),btn("📜 Broadcast History","a:bchistory")],[btn("◀️ Back","a:home")]]))
@router.callback_query(F.data == "a:createbroadcast")
async def create_broadcast(call:CallbackQuery,state:FSMContext):
    if not await admin_guard(call): return
    await state.set_state(AdminState.broadcast_text); await send_or_edit(call,"📢 BROADCAST\n\nযে message পাঠাতে চান সেটি লিখুন।",back_kb("a:broadcast"))
@router.message(AdminState.broadcast_text)
async def broadcast_input(message:Message,state:FSMContext):
    body=(message.text or "").strip()[:4000]
    if not body: await message.answer("Message লিখুন।"); return
    users=await fetchone("SELECT COUNT(*) c FROM users WHERE banned=0"); await state.clear(); await set_setting(f"broadcast_{message.from_user.id}",body)
    await message.answer(f"📢 BROADCAST PREVIEW\n\nYour message:\n\n“{body}”\n\nRecipients: {users['c']} users",reply_markup=kb([[btn("👀 Preview",f"a:bcpreview:{message.from_user.id}")],[btn("📤 Send",f"a:bcsend:{message.from_user.id}"),btn("❌ Cancel","a:broadcast")]]))
@router.callback_query(F.data.startswith("a:bcpreview:"))
async def bc_preview(call:CallbackQuery):
    if not await admin_guard(call): return
    body=await get_setting(f"broadcast_{call.from_user.id}",""); await send_or_edit(call,f"👀 BROADCAST PREVIEW\n\n{body}",kb([[btn("📤 Send",f"a:bcsend:{call.from_user.id}"),btn("❌ Cancel","a:broadcast")]]))
@router.callback_query(F.data.startswith("a:bcsend:"))
async def bc_send_prompt(call:CallbackQuery):
    if not await admin_guard(call): return
    await send_or_edit(call,"⚠️ CONFIRM ACTION\n\nএই broadcast সব active user-কে পাঠাতে চান?",confirm_kb(f"a:bcsendok:{call.from_user.id}","a:broadcast","📤 Send","❌ Cancel"))
@router.callback_query(F.data.startswith("a:bcsendok:"))
async def bc_send_execute(call:CallbackQuery):
    if not await admin_guard(call): return
    body=await get_setting(f"broadcast_{call.from_user.id}", "")
    if not body:
        await call.answer("Broadcast expired.", show_alert=True); return
    rows=await fetchall("SELECT telegram_id FROM users WHERE banned=0")
    sent=failed=0
    for r in rows:
        try:
            await call.bot.send_message(r["telegram_id"], "📢 BROADCAST\n\n" + body)
            sent += 1
        except (TelegramForbiddenError, TelegramBadRequest):
            failed += 1
        except Exception:
            failed += 1
        await asyncio.sleep(0.04)
    await execute("INSERT INTO broadcasts(admin_id,body,recipients,sent,failed,created_at) VALUES(?,?,?,?,?,?)",(call.from_user.id,body,len(rows),sent,failed,now_str()))
    await log_admin(call.from_user.id,"Broadcast sent",None,None,None)
    await call.message.edit_text(f"✅ Broadcast complete\n\nRecipients: {len(rows)}\nSent: {sent}\nFailed: {failed}")

@router.callback_query(F.data == "a:bchistory")
async def bc_history(call:CallbackQuery):
    if not await admin_guard(call): return
    rows=await fetchall("SELECT created_at,recipients,sent,failed,body FROM broadcasts ORDER BY id DESC LIMIT 10")
    text="📜 BROADCAST HISTORY\n\n"+(("\n".join(f"{r['created_at'][0:16]} • {r['sent']}/{r['recipients']} • {r['body'][:60]}" for r in rows)) if rows else "কোনো broadcast history নেই।")
    await send_or_edit(call,text,back_kb("a:broadcast"))

# ============================================================
# 22. SETTINGS / ADMIN MANAGEMENT
# ============================================================
@router.callback_query(F.data == "a:settings")
async def admin_settings(call:CallbackQuery):
    if not await admin_guard(call): return
    maintenance=await get_setting("maintenance","0"); minimum=await get_setting("min_withdraw",str(DEFAULT_MIN_WITHDRAW))
    sw=await get_setting("submission_window_enabled","1"); ss=await get_setting("submission_start","00:00"); se=await get_setting("submission_end","20:00"); alerts=await get_setting("submission_alerts_enabled","1")
    schedule_status = "🟢 ON" if sw=="1" else "🔴 OFF"
    alert_status = "🟢 ON" if alerts=="1" else "🔴 OFF"
    proof_chat_id=await get_setting("payment_proof_chat_id","")
    proof_title=await get_setting("payment_proof_chat_title","")
    proof_status=(f"🟢 {proof_title or proof_chat_id}" if proof_chat_id else "🔴 Not selected")
    fj_title=await get_setting("force_join_chat_title","")
    fj_id=await get_setting("force_join_chat_id","")
    fj_status=(f"🟢 {fj_title or fj_id}" if fj_id else "🔴 Not selected")
    await send_or_edit(call,f"⚙️ BOT SETTINGS\n\n💵 Withdrawal Minimum: {CURRENCY}{fmt_money(minimum)}\n🔧 Maintenance: {'ON' if maintenance=='1' else 'OFF'}\n📢 Payment Proof Destination: <b>{proof_status}</b>\n🔐 Force Join Channel: <b>{fj_status}</b>\n\n🕐 Submission Window: <b>{schedule_status}</b>\n⏰ Time: <b>{format_12h(ss)} – {format_12h(se)}</b> 🇧🇩\n🔔 Deadline Alerts: <b>{alert_status}</b>",kb([[btn("💵 Rates","a:categories"),btn("⏱ Payment Settings","a:paysettings")],[btn("🕐 Submission Schedule","a:subschedule","primary")],[btn("📢 Payment Proof Channel / Group","a:proofchannel","success")],[btn("🔐 Force Join Channel","a:forcejoin","success")],[btn("🗑️ Clear Proof Destination","a:proofclear","danger"),btn("🗑️ Clear Force Join","a:forcejoinclear","danger")],[btn("👑 Admin Management","a:admins"),btn("🔧 Maintenance","a:maintenance")],[btn("📝 Rules","a:editrules")],[btn("◀️ Back","a:home")]]))
@router.message(F.text == "❌ Cancel")
async def cancel_proof_channel_selector(message: Message, state: FSMContext):
    if not await admin_guard(message): return
    await state.clear()
    await message.answer("❌ Channel selection cancelled.", reply_markup=ReplyKeyboardRemove())

@router.callback_query(F.data == "a:proofchannel")
async def proof_channel_selector(call: CallbackQuery):
    if not await admin_guard(call): return
    current=await get_setting("payment_proof_chat_title","")
    text=(
        "📢 <b>PAYMENT PROOF DESTINATION</b>\n\n"
        "নিচের <b>Select Channel</b> বা <b>Select Group</b> button চাপুন এবং Telegram-এর chat picker থেকে সরাসরি payment proof destination নির্বাচন করুন।\n\n"
        "✅ কোনো Channel ID লিখতে হবে না\n"
        "✅ কোনো link দিতে হবে না\n"
        "📄 Channel-এ শুধু Invoice যাবে\n"
        "👤 User-এর কাছে Invoice + Screenshot যাবে"
    )
    if current: text += f"\n\n📌 Current: <b>{current}</b>"
    picker=ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📢 Select Channel", request_chat=KeyboardButtonRequestChat(request_id=91001, chat_is_channel=True, bot_is_member=True))],
            [KeyboardButton(text="👥 Select Group", request_chat=KeyboardButtonRequestChat(request_id=91003, chat_is_channel=False, bot_is_member=True))],
            [KeyboardButton(text="❌ Cancel")]
        ], resize_keyboard=True, one_time_keyboard=True
    )
    await call.message.answer(text, reply_markup=picker)
    await call.answer()

@router.message(F.chat_shared)
async def chat_picker_selected(message: Message):
    if not await admin_guard(message): return
    shared=message.chat_shared
    chat_id=int(shared.chat_id)
    request_id=int(getattr(shared, "request_id", 0) or 0)
    try:
        chat=await message.bot.get_chat(chat_id)
    except Exception as exc:
        log.exception("Selected chat could not be resolved: %s", exc)
        await message.answer("❌ এই chat access করা যাচ্ছে না। Bot-কে ওই channel-এ add করে আবার Select করুন।", reply_markup=ReplyKeyboardRemove())
        return
    title=chat.title or shared.title or "Selected Chat"
    username=getattr(chat,"username",None) or getattr(shared,"username",None) or ""

    if request_id == 91002:
        # Force Join must be a channel where the bot is an administrator.
        try:
            me = await message.bot.get_me()
            bot_member = await message.bot.get_chat_member(chat_id, me.id)
            bot_status = getattr(bot_member, "status", "")
            if bot_status not in {"creator", "administrator"}:
                await message.answer("❌ Force Join set করা যায়নি।\n\nএই channel-এ Bot-কে আগে <b>Administrator</b> করতে হবে।\nতারপর আবার Select Force Join Channel চাপুন।", reply_markup=ReplyKeyboardRemove())
                return
        except Exception as exc:
            log.exception("Force Join bot permission check failed: %s", exc)
            await message.answer("❌ Channel permission verify করা যায়নি। Bot-কে channel-এর Administrator করে আবার চেষ্টা করুন।", reply_markup=ReplyKeyboardRemove())
            return
        invite_link = f"https://t.me/{username}" if username else ""
        if not invite_link:
            try:
                invite_link = await message.bot.export_chat_invite_link(chat_id)
            except Exception as exc:
                log.warning("Could not create/export Force Join invite link: %s", exc)
        if not invite_link:
            await message.answer("❌ Force Join set করা যায়নি।\n\nPrivate channel-এর জন্য Bot একটি invite link তৈরি করতে পারেনি। Bot-কে Administrator করে আবার চেষ্টা করুন।", reply_markup=ReplyKeyboardRemove())
            return
        await set_setting("force_join_chat_id",str(chat_id))
        await set_setting("force_join_chat_title",title)
        await set_setting("force_join_chat_username",username)
        await set_setting("force_join_invite_link",invite_link)
        await log_admin(message.from_user.id,"Force Join channel selected",str(chat_id))
        await message.answer(
            f"✅ <b>Force Join Channel Set</b>\n\n📢 {html.escape(title)}\n🆔 Chat ID: <code>{chat_id}</code>\n\n"
            "🔐 এখন থেকে user-কে এই channel-এ join করে bot ব্যবহার করতে হবে।\n"
            f"🔗 Join link: <code>{html.escape(invite_link)}</code>",
            reply_markup=ReplyKeyboardRemove()
        )
        return

    if request_id in {91001, 91003}:
        # Payment proof destination: channel needs admin rights; group only needs membership.
        try:
            me = await message.bot.get_me()
            bot_member = await message.bot.get_chat_member(chat_id, me.id)
            bot_status = getattr(bot_member, "status", "")
            if bot_status in {"left", "kicked"}:
                await message.answer("❌ Bot এই chat-এর member নয়। আগে Bot-কে add করুন এবং আবার select করুন।", reply_markup=ReplyKeyboardRemove())
                return
            if request_id == 91001 and bot_status not in {"creator", "administrator"}:
                await message.answer("❌ Payment Proof Channel set করা যায়নি। Channel-এ Bot-কে Administrator করুন এবং আবার select করুন।", reply_markup=ReplyKeyboardRemove())
                return
        except Exception as exc:
            log.exception("Payment proof destination permission check failed: %s", exc)
            await message.answer("❌ Selected chat-এর permission verify করা যায়নি। Bot-কে ওই chat-এ add করুন এবং আবার চেষ্টা করুন।", reply_markup=ReplyKeyboardRemove())
            return
        await set_setting("payment_proof_chat_id",str(chat_id))
        await set_setting("payment_proof_chat_title",title)
        await set_setting("payment_proof_chat_username",username)
        await log_admin(message.from_user.id,"Payment proof destination selected",str(chat_id))
        await message.answer(f"✅ <b>Payment Proof Destination Set</b>\n\n📢 {html.escape(title)}\n🆔 Chat ID: <code>{chat_id}</code>\n\n📄 Destination-এ শুধু invoice যাবে।\n📸 User-এর কাছে invoice + screenshot যাবে।", reply_markup=ReplyKeyboardRemove())
        return

    await message.answer("❌ Unknown chat selection request। আবার Settings থেকে নির্বাচন করুন।", reply_markup=ReplyKeyboardRemove())

@router.callback_query(F.data == "a:forcejoin")
async def force_join_selector(call: CallbackQuery):
    if not await admin_guard(call): return
    current=await get_setting("force_join_chat_title","")
    text=(
        "🔐 <b>FORCE JOIN CHANNEL</b>\n\n"
        "নিচের <b>Select Channel</b> button চাপুন এবং Telegram-এর chat picker থেকে সরাসরি Force Join channel নির্বাচন করুন।\n\n"
        "✅ কোনো Channel ID লাগবে না\n"
        "✅ কোনো link দিতে হবে না\n"
        "👤 User আগে channel-এ join করবে, তারপর bot ব্যবহার করতে পারবে।"
    )
    if current: text += f"\n\n📌 Current: <b>{html.escape(current)}</b>"
    picker=ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="🔐 Select Force Join Channel", request_chat=KeyboardButtonRequestChat(request_id=91002, chat_is_channel=True, bot_is_member=True))],
            [KeyboardButton(text="❌ Cancel")]
        ], resize_keyboard=True, one_time_keyboard=True
    )
    await call.message.answer(text, reply_markup=picker)
    await call.answer()

@router.callback_query(F.data == "a:forcejoinclear")
async def force_join_clear(call: CallbackQuery):
    if not await admin_guard(call): return
    await set_setting("force_join_chat_id","")
    await set_setting("force_join_chat_title","")
    await set_setting("force_join_chat_username","")
    await set_setting("force_join_invite_link","")
    await log_admin(call.from_user.id,"Force Join channel cleared")
    await call.answer("Force Join channel cleared", show_alert=True)
    await admin_settings(call)
@router.callback_query(F.data == "a:proofclear")
async def proof_channel_clear(call: CallbackQuery):
    if not await admin_guard(call): return
    await set_setting("payment_proof_chat_id","")
    await set_setting("payment_proof_chat_title","")
    await set_setting("payment_proof_chat_username","")
    await log_admin(call.from_user.id,"Payment proof destination cleared")
    await call.answer("Payment proof destination cleared", show_alert=True)
    await admin_settings(call)

@router.callback_query(F.data == "a:paysettings")
async def pay_settings(call:CallbackQuery):
    if not await admin_guard(call): return
    await send_or_edit(call,f"⏱ PAYMENT SETTINGS\n\nMinimum withdrawal: {CURRENCY}{fmt_money(await get_setting('min_withdraw',str(DEFAULT_MIN_WITHDRAW)))}",kb([[btn("✏️ Change Minimum","a:minwd")],[btn("◀️ Back","a:settings")]]))
@router.callback_query(F.data == "a:minwd")
async def minwd_start(call:CallbackQuery,state:FSMContext):
    if not await admin_guard(call): return
    await state.set_state(AdminState.set_min_withdraw); await send_or_edit(call,"💵 নতুন minimum withdrawal amount লিখুন।",back_kb("a:paysettings"))
@router.message(AdminState.set_min_withdraw)
async def minwd_input(message:Message,state:FSMContext):
    try: amount=money(message.text)
    except: amount=Decimal("-1")
    if amount<0: await message.answer("Valid amount দিন।"); return
    await set_setting("min_withdraw",str(amount)); await log_admin(message.from_user.id,"Minimum withdrawal changed",None,amount); await state.clear(); await message.answer(f"✅ Minimum withdrawal updated: {CURRENCY}{fmt_money(amount)}",reply_markup=back_kb("a:settings"))

@router.callback_query(F.data == "a:editrules")
async def edit_rules_start(call:CallbackQuery,state:FSMContext):
    if not await admin_guard(call): return
    await state.set_state(AdminState.edit_rules); await send_or_edit(call,"📝 RULES EDITOR\n\nবর্তমান rules:\n\n"+(await get_setting("rules")),kb([[btn("❌ Cancel","a:settings")]]))
@router.message(AdminState.edit_rules)
async def edit_rules_input(message:Message,state:FSMContext):
    text=(message.text or "").strip()[:7000]
    if not text: await message.answer("Rules লিখুন।"); return
    await set_setting("rules",text); await log_admin(message.from_user.id,"Rules updated"); await state.clear(); await message.answer("✅ Rules updated.",reply_markup=back_kb("a:settings"))

@router.callback_query(F.data == "a:subschedule")
async def submission_schedule_settings(call: CallbackQuery):
    if not await admin_guard(call): return
    enabled=await get_setting("submission_window_enabled","1")
    start_time=await get_setting("submission_start","00:00")
    end_time=await get_setting("submission_end","20:00")
    alerts=await get_setting("submission_alerts_enabled","1")
    await send_or_edit(
        call,
        f"🕐 <b>SUBMISSION SCHEDULE</b> 🇧🇩\n\n"
        f"📌 Status: <b>{'🟢 ON' if enabled=='1' else '🔴 OFF'}</b>\n"
        f"⏰ Submission Time: <b>{format_12h(start_time)} – {format_12h(end_time)}</b>\n"
        f"🔔 Alerts: <b>{'🟢 ON' if alerts=='1' else '🔴 OFF'}</b>\n\n"
        "💡 ON থাকলে শুধু নির্ধারিত সময়ের মধ্যে নতুন file submit করা যাবে।\n"
        "🛑 OFF করলে পুরো submission window বন্ধ থাকবে।\n\n"
        "🕰️ সময় এখন 12-hour format: <b>1–12 + AM/PM</b>\n"
        "✏️ উদাহরণ: <b>10:30 AM – 8:00 PM</b>",
        kb([
            [btn("🟢/🔴 Submission ON/OFF","a:subscheduletoggle","success")],
            [btn("⏰ Change Time","a:subscheduletime","primary"),btn("🔔 Alerts ON/OFF","a:submissionalerts","success")],
            [btn("◀️ Back","a:settings")]
        ])
    )

@router.callback_query(F.data == "a:subscheduletoggle")
async def submission_schedule_toggle(call: CallbackQuery):
    if not await admin_guard(call): return
    cur=await get_setting("submission_window_enabled","1"); new="0" if cur=="1" else "1"
    await set_setting("submission_window_enabled",new); await log_admin(call.from_user.id,"Submission window changed",new)
    await call.answer("Submission window ON" if new=="1" else "Submission window OFF", show_alert=True)
    await submission_schedule_settings(call)

@router.callback_query(F.data == "a:subscheduletime")
async def submission_schedule_time_start(call: CallbackQuery,state:FSMContext):
    if not await admin_guard(call): return
    await state.set_state(AdminState.set_submission_schedule)
    await state.update_data(schedule_step="start", schedule_meridiem=None, schedule_hour=None, schedule_minute=None)
    await send_or_edit(
        call,
        "⏰ <b>SET SUBMISSION TIME</b> 🇧🇩\n\n"
        "<b>Step 1/2 — START TIME</b>\n\n"
        "✏️ আগে শুধু সময় লিখুন: <b>1–12</b> অথবা <b>1:30</b>\n"
        "তারপর নিচের <b>AM / PM</b> button চাপুন।\n\n"
        "উদাহরণ: <code>10:30</code> → তারপর <b>AM</b> চাপুন।",
        kb([[btn("🌅 AM","a:schedampm:start:AM","success"),btn("🌙 PM","a:schedampm:start:PM","primary")],[btn("❌ Cancel","a:subschedule","danger")]])
    )

@router.callback_query(F.data.startswith("a:schedampm:"))
async def submission_schedule_ampm(call: CallbackQuery,state:FSMContext):
    if not await admin_guard(call): return
    parts=call.data.split(":")
    step=parts[2]
    meridiem=parts[3].upper()
    if meridiem not in {"AM","PM"}:
        await call.answer("AM/PM নির্বাচন সঠিক নয়।",show_alert=True); return
    data=await state.get_data()
    if data.get("schedule_step") not in {"start","end"} or data.get("schedule_step") != step:
        await call.answer("এই AM/PM button এখন active নেই।",show_alert=True); return
    raw=data.get("schedule_raw_time")
    if raw:
        parsed=parse_12h_time(raw)
        if not parsed:
            await call.answer("সময়টি আগে সঠিকভাবে লিখুন।",show_alert=True); return
        hour,minute=parsed
        converted=convert_12h_to_24(hour,minute,meridiem)
        if step=="start":
            await state.update_data(start_24=converted,schedule_step="end",schedule_raw_time=None,schedule_meridiem=None)
            await call.answer(f"Start time: {format_12h(converted)}")
            await send_or_edit(call,"⏰ <b>Step 2/2 — END TIME</b>\n\n✏️ End time লিখুন: <b>1–12</b> অথবা <b>1:30</b>\nতারপর নিচের <b>AM / PM</b> button চাপুন।",kb([[btn("🌅 AM","a:schedampm:end:AM","success"),btn("🌙 PM","a:schedampm:end:PM","primary")],[btn("❌ Cancel","a:subschedule","danger")]]))
        else:
            start_24=data.get("start_24")
            if not start_24:
                await call.answer("Start time পাওয়া যায়নি। আবার শুরু করুন।",show_alert=True); return
            await set_setting("submission_start",start_24); await set_setting("submission_end",converted)
            await log_admin(call.from_user.id,"Submission schedule changed",f"{start_24}-{converted}")
            await state.clear()
            await call.answer("Submission time updated")
            await send_or_edit(call,f"✅ <b>Submission time updated</b>\n\n⏰ <b>{format_12h(start_24)} – {format_12h(converted)}</b> 🇧🇩",back_kb("a:subschedule"))
    else:
        await state.update_data(schedule_meridiem=meridiem)
        await call.answer(f"{meridiem} selected")
        await send_or_edit(call,f"✏️ <b>{step.upper()} TIME</b>\n\nএখন সময়টি লিখুন: <b>1–12</b> অথবা <b>1:30</b>\n\nSelected: <b>{meridiem}</b>",kb([[btn("🌅 AM","a:schedampm:"+step+":AM","success"),btn("🌙 PM","a:schedampm:"+step+":PM","primary")],[btn("❌ Cancel","a:subschedule","danger")]]))

@router.message(AdminState.set_submission_schedule)
async def submission_schedule_time_input(message: Message,state:FSMContext):
    raw=(message.text or "").strip()
    parsed=parse_12h_time(raw)
    if not parsed:
        await message.answer("❌ সময় ভুল। শুধু 1–12 বা 1:30 format দিন। Example: 10:30"); return
    data=await state.get_data()
    step=data.get("schedule_step")
    if step not in {"start","end"}:
        await state.clear(); await message.answer("❌ Schedule session expired। আবার Change Time চাপুন।",reply_markup=back_kb("a:subschedule")); return
    meridiem=data.get("schedule_meridiem")
    if meridiem:
        hour,minute=parsed
        converted=convert_12h_to_24(hour,minute,meridiem)
        if step=="start":
            await state.update_data(start_24=converted,schedule_step="end",schedule_raw_time=None,schedule_meridiem=None)
            await message.answer(f"✅ Start time: <b>{format_12h(converted)}</b>\n\n⏰ <b>Step 2/2 — END TIME</b>\n\n✏️ End time লিখুন: <b>1–12</b> অথবা <b>1:30</b>\nতারপর AM/PM select করুন।",reply_markup=kb([[btn("🌅 AM","a:schedampm:end:AM","success"),btn("🌙 PM","a:schedampm:end:PM","primary")],[btn("❌ Cancel","a:subschedule","danger")]]))
        else:
            start_24=data.get("start_24")
            if not start_24:
                await state.clear(); await message.answer("❌ Start time পাওয়া যায়নি। আবার Change Time চাপুন।",reply_markup=back_kb("a:subschedule")); return
            await set_setting("submission_start",start_24); await set_setting("submission_end",converted)
            await log_admin(message.from_user.id,"Submission schedule changed",f"{start_24}-{converted}")
            await state.clear()
            await message.answer(f"✅ <b>Submission time updated</b>\n\n⏰ <b>{format_12h(start_24)} – {format_12h(converted)}</b> 🇧🇩",reply_markup=back_kb("a:subschedule"))
    else:
        await state.update_data(schedule_raw_time=raw)
        await message.answer(f"🕐 <b>{step.upper()} TIME</b>: <b>{raw}</b>\n\nএখন নিচের <b>AM / PM</b> button থেকে একটি select করুন।",reply_markup=kb([[btn("🌅 AM","a:schedampm:"+step+":AM","success"),btn("🌙 PM","a:schedampm:"+step+":PM","primary")],[btn("❌ Cancel","a:subschedule","danger")]]))

@router.callback_query(F.data == "a:submissionalerts")
async def submission_alerts_toggle(call: CallbackQuery):
    if not await admin_guard(call): return
    cur=await get_setting("submission_alerts_enabled","1"); new="0" if cur=="1" else "1"
    await set_setting("submission_alerts_enabled",new); await log_admin(call.from_user.id,"Submission alerts changed",new)
    await call.answer("Alerts ON" if new=="1" else "Alerts OFF", show_alert=True)
    await submission_schedule_settings(call)

@router.callback_query(F.data == "a:maintenance")
async def maintenance_prompt(call:CallbackQuery):
    if not await admin_guard(call): return
    cur=await get_setting("maintenance","0"); new="0" if cur=="1" else "1"
    await send_or_edit(call,f"⚠️ CONFIRM ACTION\n\nMaintenance {'ON' if new=='1' else 'OFF'} করবেন?",confirm_kb(f"a:maintenanceok:{new}","a:settings"))
@router.callback_query(F.data.startswith("a:maintenanceok:"))
async def maintenance_execute(call:CallbackQuery):
    if not await admin_guard(call): return
    new=call.data.split(":")[-1]; await set_setting("maintenance",new); await log_admin(call.from_user.id,"Maintenance changed",new); await call.message.edit_text(f"🔧 Maintenance {'ON' if new=='1' else 'OFF'}")

@router.callback_query(F.data == "a:admins")
async def admin_management(call:CallbackQuery):
    if not await admin_guard(call): return
    if not await is_owner(call.from_user.id):
        await call.answer("শুধু Owner এই section manage করতে পারবেন।",show_alert=True); return
    rows=await fetchall("SELECT user_id,role FROM admins ORDER BY user_id")
    listing="\n".join(f"• {r['user_id']} — {r['role']}" for r in rows)
    await send_or_edit(call,"👑 ADMIN MANAGEMENT\n\n"+listing,kb([[btn("➕ Add Second Admin","a:addadmin")],[btn("➖ Remove Admin","a:removeadmin")],[btn("◀️ Back","a:settings")]]))
@router.callback_query(F.data == "a:addadmin")
async def add_admin_start(call:CallbackQuery,state:FSMContext):
    if not await is_owner(call.from_user.id): await call.answer("Owner only.",show_alert=True); return
    await state.set_state(AdminState.add_admin); await send_or_edit(call,"➕ Admin-এর Telegram numeric UID পাঠান।",back_kb("a:admins"))
@router.message(AdminState.add_admin)
async def add_admin_input(message:Message,state:FSMContext):
    if not await is_owner(message.from_user.id): await state.clear(); return
    try: uid=int(message.text.strip())
    except: await message.answer("Numeric Telegram UID দিন।"); return
    if uid==OWNER_ID: await message.answer("Owner already protected."); return
    await ensure_admin(uid,"admin",message.from_user.id); await log_admin(message.from_user.id,"Second admin added",str(uid)); await state.clear(); await message.answer(f"✅ Admin added: {uid}",reply_markup=back_kb("a:admins"))
@router.callback_query(F.data == "a:removeadmin")
async def remove_admin_start(call:CallbackQuery,state:FSMContext):
    if not await is_owner(call.from_user.id): await call.answer("Owner only.",show_alert=True); return
    await state.set_state(AdminState.remove_admin); await send_or_edit(call,"➖ যে admin remove করবেন তার UID পাঠান।",back_kb("a:admins"))
@router.message(AdminState.remove_admin)
async def remove_admin_input(message:Message,state:FSMContext):
    if not await is_owner(message.from_user.id): await state.clear(); return
    try: uid=int(message.text.strip())
    except: await message.answer("Numeric UID দিন।"); return
    if uid==OWNER_ID: await message.answer("Owner remove করা যাবে না।"); return
    await execute("DELETE FROM admins WHERE user_id=?",(uid,)); await log_admin(message.from_user.id,"Admin removed",str(uid)); await state.clear(); await message.answer(f"✅ Admin removed: {uid}",reply_markup=back_kb("a:admins"))

# ============================================================
# 23. ADMIN SUPPORT REPLY
# ============================================================
@router.callback_query(F.data.startswith("a:sreply:"))
async def admin_support_reply_start(call:CallbackQuery,state:FSMContext):
    if not await admin_guard(call): return
    tid=int(call.data.split(":")[-1]); ticket=await fetchone("SELECT * FROM support_tickets WHERE id=?",(tid,))
    if not ticket or ticket["status"]!="active": await call.answer("Ticket closed.",show_alert=True); return
    await state.set_state(AdminState.support_reply); await state.update_data(ticket_id=tid)
    await send_or_edit(call,f"💬 SUPPORT REPLY\n\nTicket: #{ticket['public_id']}\nReply message লিখুন।",kb([[btn("❌ Cancel","a:home")]]))
@router.message(AdminState.support_reply)
async def admin_support_reply_message(message:Message,state:FSMContext):
    if not await admin_guard(message): return
    data=await state.get_data(); tid=data.get("ticket_id"); ticket=await fetchone("SELECT * FROM support_tickets WHERE id=?",(tid,))
    if not ticket or ticket["status"]!="active": await state.clear(); await message.answer("Ticket closed."); return
    body=(message.text or "").strip()[:4000]
    if not body: await message.answer("Reply text লিখুন।"); return
    await execute("INSERT INTO support_messages(ticket_id,sender_id,sender_role,message_type,body,created_at) VALUES(?,?,?,?,?,?)",(tid,message.from_user.id,"admin","text",body,now_str()))
    await state.clear()
    try: await message.bot.send_message(ticket["user_id"],f"👨‍💻 Support\n\n{body}",reply_markup=kb([[btn("❌ Close Chat","s:close")]]))
    except Exception: pass
    await message.answer("✅ Reply পাঠানো হয়েছে।",reply_markup=back_kb("a:home"))

# ============================================================
# 24. SEARCH PLACEHOLDERS / SAFE SEARCH
# ============================================================
@router.callback_query(F.data == "a:searchsub")
async def search_sub_start(call:CallbackQuery,state:FSMContext):
    if not await admin_guard(call): return
    await state.set_state(AdminState.search_submission); await send_or_edit(call,"🔎 Submission ID (SUB-000001) পাঠান।",back_kb("a:subs"))
@router.message(AdminState.search_submission)
async def search_sub_input(message:Message,state:FSMContext):
    q=(message.text or "").strip().upper(); r=await fetchone("SELECT id FROM submissions WHERE upper(public_id)=?",(q,)); await state.clear()
    if not r: await message.answer("❌ Submission পাওয়া যায়নি।",reply_markup=back_kb("a:subs")); return
    await message.answer("🔎 Submission found.",reply_markup=kb([[btn("📄 Open",f"a:sub:{r['id']}")],[btn("◀️ Back","a:subs")]]))

@router.callback_query(F.data == "a:searchwd")
async def search_wd_start(call:CallbackQuery,state:FSMContext):
    if not await admin_guard(call): return
    await state.set_state(AdminState.search_payment); await send_or_edit(call,"🔎 Withdrawal ID (WD-000001) পাঠান।",back_kb("a:payments"))
@router.message(AdminState.search_payment)
async def search_wd_input(message:Message,state:FSMContext):
    q=(message.text or "").strip().upper(); r=await fetchone("SELECT public_id FROM withdrawals WHERE upper(public_id)=?",(q,)); await state.clear()
    if not r: await message.answer("❌ Payment পাওয়া যায়নি।",reply_markup=back_kb("a:payments")); return
    await message.answer("🔎 Payment found.",reply_markup=kb([[btn("💳 Open",f"a:wd:{r['public_id']}")],[btn("◀️ Back","a:payments")]]))

# ============================================================
# 25. TEXT ROUTING / CANCEL
# ============================================================
# The visible UI is inline, so ordinary text matching below only handles graceful unexpected messages.
@router.message(F.text)
async def generic_text(message:Message,state:FSMContext):
    if not await guard_user(message): return
    current=await state.get_state()
    if current in {
        AdminState.add_category_name.state,AdminState.add_category_rate.state,AdminState.rename_category.state,
        AdminState.change_rate.state,AdminState.add_format.state,AdminState.edit_format.state,AdminState.reject_reason.state,
        AdminState.adjust_quantity.state,AdminState.add_balance.state,AdminState.deduct_balance.state,AdminState.search_user.state,
        AdminState.search_submission.state,AdminState.search_payment.state,AdminState.support_reply.state,AdminState.broadcast_text.state,
        AdminState.add_admin.state,AdminState.remove_admin.state,AdminState.edit_rules.state,AdminState.set_min_withdraw.state,AdminState.payment_txid.state,AdminState.payment_screenshot.state,
        UserInputState.payment_number.state,UserInputState.custom_withdraw.state,SupportState.active.state
    }:
        return
    await message.answer("📌 প্রয়োজনীয় কাজটি Main Menu থেকে নির্বাচন করুন। 🤝",reply_markup=main_kb(await is_admin(message.from_user.id)))

@router.message(Command("cancel"))
async def cmd_cancel(message: Message, state: FSMContext):
    await state.clear()
    await message.answer("❌ Current action বাতিল করা হয়েছে।", reply_markup=main_kb(await is_admin(message.from_user.id)))

# ============================================================
# 26. CALLBACK SAFETY / UNKNOWN CALLBACK
# ============================================================
@router.callback_query()
async def unknown_callback(call:CallbackQuery):
    if call.data == "a:none":
        await call.answer("এই action এখনো প্রয়োজন নেই।",show_alert=True); return
    await call.answer("⚠️ এই button আর active নেই। নতুন করে screen খুলুন।",show_alert=True)

# ============================================================
# 27. ERROR HANDLING
# ============================================================
async def on_error(event):
    log.exception("Unhandled bot error: %r", event.exception)

# ============================================================
# 28. DAILY SUBMISSION DEADLINE ALERTS
# ============================================================
# Alerts are automatically calculated from the Admin-configured submission_end time.
# Default: 20:00 => 19:30 (30 min) and 19:50 (10 min), Bangladesh time.

def deadline_alert_text(minutes_left: int, end_time: str) -> str:
    if minutes_left == 30:
        return (
            "╭━━━〔 🔔 <b>SUBMISSION ALERT</b> 〕━━━╮\n"
            "\n"
            "🌙 <b>আজকের Submission Window বন্ধ হতে আর মাত্র ৩০ মিনিট!</b>\n\n"
            "📥 আপনার account/file এখনো জমা না দিলে এখনই Submit করে দিন।\n"
            "⚡ শেষ মুহূর্তে অপেক্ষা না করে ফাইল পাঠিয়ে দিন।\n\n"
            f"⏰ <b>শেষ সময়: {end_time} (বাংলাদেশ সময়)</b> 🇧🇩\n"
            "🟢 Submission এখনো গ্রহণ করা হচ্ছে।\n\n"
            "╰━━━〔 🚀 <b>SUBMIT NOW</b> 〕━━━╯\n"
        )
    return (
        "╭━━━〔 🚨 <b>FINAL ALERT</b> 〕━━━╮\n"
        "\n"
        "⏳ <b>শেষ ১০ মিনিট শুরু হয়ে গেছে!</b> 🔥\n\n"
        "📤 আজকের submission এখনো খোলা আছে—যাদের file বাকি, দ্রুত Submit করুন।\n"
        f"🕗 <b>Deadline: {end_time}</b> 🇧🇩\n"
        "🔒 সময় শেষ হলে নতুন submission নেওয়া বন্ধ হয়ে যাবে।\n\n"
        "🏃‍♂️ দেরি করবেন না!\n"
        "╰━━━〔 ⚡ <b>10 MINUTES LEFT</b> 〕━━━╯\n"
    )

async def send_daily_submission_alert(bot: Bot, text: str, slot_key: str):
    setting_key = f"daily_alert_sent_{slot_key}_{datetime.now(TIMEZONE).strftime('%Y-%m-%d')}"
    if await get_setting(setting_key, "0") == "1":
        return
    users = await fetchall("SELECT telegram_id FROM users WHERE COALESCE(banned,0)=0")
    sent=failed=0
    markup=kb([[btn("📤 এখনই Submit করুন", "u:submit", "success")]])
    for row in users:
        uid=int(row["telegram_id"])
        try:
            await bot.send_message(uid,text,reply_markup=markup); sent+=1; await asyncio.sleep(0.05)
        except Exception as exc:
            failed+=1; log.warning("Daily alert failed for %s: %s",uid,exc)
    await set_setting(setting_key,"1")
    log.info("Daily submission alert sent | slot=%s | sent=%s | failed=%s",slot_key,sent,failed)

async def daily_submission_alert_scheduler(bot: Bot):
    while True:
        try:
            if await get_setting("submission_alerts_enabled","1") == "1" and await get_setting("submission_window_enabled","1") == "1":
                end=await get_setting("submission_end","20:00")
                ep=parse_hhmm(end)
                if ep:
                    now=datetime.now(TIMEZONE); now_min=now.hour*60+now.minute; end_min=ep[0]*60+ep[1]
                    for left in (30,10):
                        target=(end_min-left)%1440
                        if now_min==target:
                            await send_daily_submission_alert(bot,deadline_alert_text(left,format_12h(end)),f"{end.replace(':','')}_{left}")
            await asyncio.sleep(20)
        except asyncio.CancelledError:
            raise
        except Exception as exc:
            log.exception("Daily alert scheduler error: %s",exc); await asyncio.sleep(20)

# ============================================================
# 29. WEB ADMIN PANEL — Telegram Admin + Web Admin together
# ============================================================
WEB_SESSIONS = {}
WEB_CSS = """
*{box-sizing:border-box}body{margin:0;background:#0b1220;color:#e5e7eb;font-family:Arial,sans-serif}a{text-decoration:none;color:inherit}.wrap{display:flex;min-height:100vh}.side{width:245px;background:#111827;padding:18px;position:fixed;inset:0 auto 0 0;overflow:auto}.brand{font-size:20px;font-weight:800;margin-bottom:18px}.nav a{display:block;padding:11px 12px;margin:5px 0;border-radius:10px;color:#cbd5e1}.nav a:hover{background:#1f2937;color:#fff}.main{margin-left:245px;padding:24px;width:calc(100% - 245px)}.top{display:flex;justify-content:space-between;align-items:center;margin-bottom:22px}.title{font-size:25px;font-weight:800}.muted{color:#94a3b8}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:14px}.card{background:#111827;border:1px solid #263244;border-radius:14px;padding:17px}.num{font-size:27px;font-weight:800;margin-top:8px}.table{width:100%;border-collapse:collapse;background:#111827}.table th,.table td{padding:11px;border-bottom:1px solid #263244;text-align:left;font-size:14px}.table th{color:#93c5fd}.btn{display:inline-block;border:0;border-radius:9px;padding:9px 13px;background:#2563eb;color:#fff;cursor:pointer;margin:3px}.green{background:#059669}.red{background:#dc2626}.gray{background:#374151}.input,.select,.textarea{width:100%;padding:11px;border-radius:9px;border:1px solid #334155;background:#0f172a;color:#fff;margin:5px 0 10px}.form{max-width:700px}.alert{padding:11px;background:#172554;border-radius:10px;margin-bottom:14px}@media(max-width:800px){.side{position:static;width:100%;min-height:auto}.wrap{display:block}.main{margin:0;width:100%;padding:15px}.nav{display:grid;grid-template-columns:1fr 1fr;gap:4px}}
"""

def web_html(title, body, active="dashboard"):
    nav=[("dashboard","📊 Dashboard","/"),("submissions","📥 Submissions","/submissions"),("payments","💳 Payments","/payments"),("users","👥 Users","/users"),("categories","📦 Categories","/categories"),("balance","💰 Balance Management","/balance"),("statistics","📈 Statistics","/statistics"),("broadcast","📢 Broadcast","/broadcast"),("settings","⚙️ Settings","/settings"),("admins","👑 Admin Management","/admins"),("logs","📝 Admin Logs","/logs")]
    links=''.join(f'<a href="{u}" style="{("background:#1e293b;color:#fff" if k==active else "")}">{n}</a>' for k,n,u in nav)
    return f'''<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><title>{html.escape(title)}</title><style>{WEB_CSS}</style></head><body><div class="wrap"><aside class="side"><div class="brand">👑 Apex Valorem</div><div class="muted">Web Admin Panel</div><nav class="nav">{links}<a href="/logout">🚪 Logout</a></nav></aside><main class="main"><div class="top"><div class="title">{html.escape(title)}</div><div class="muted">🇧🇩 Asia/Dhaka</div></div>{body}</main></div></body></html>'''

def web_login_page(msg=""):
    alert=f'<div class="alert">{html.escape(msg)}</div>' if msg else ''
    return f'''<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><title>Web Admin Login</title><style>{WEB_CSS}</style></head><body><div class="login card"><h2>👑 Admin Web Panel</h2>{alert}<form method="post" action="/login"><label>Username</label><input class="input" name="username" required><label>Password</label><input class="input" type="password" name="password" required><button class="btn green" style="width:100%">🔐 Login</button></form></div></body></html>'''

def web_authed(request):
    token=request.cookies.get("admin_session","")
    return bool(token and WEB_SESSIONS.get(token)==WEB_ADMIN_USER)

async def web_auth(request):
    if not web_authed(request): raise web.HTTPFound('/login')

async def web_login(request):
    if request.method=="GET": return web.Response(text=web_login_page(),content_type="text/html")
    data=await request.post()
    if data.get("username")==WEB_ADMIN_USER and data.get("password")==WEB_ADMIN_PASS:
        token=secrets.token_urlsafe(32); WEB_SESSIONS[token]=WEB_ADMIN_USER
        resp=web.HTTPFound('/'); resp.set_cookie("admin_session",token,httponly=True,samesite="Lax",max_age=86400); raise resp
    return web.Response(text=web_login_page("❌ Username or password ভুল।"),content_type="text/html",status=401)

async def web_logout(request):
    WEB_SESSIONS.pop(request.cookies.get("admin_session",""),None); resp=web.HTTPFound('/login'); resp.del_cookie("admin_session"); raise resp

async def web_dashboard(request):
    await web_auth(request)
    users=(await fetchone("SELECT COUNT(*) c FROM users"))["c"]; subs=(await fetchone("SELECT COUNT(*) c FROM submissions"))["c"]; pending=(await fetchone("SELECT COUNT(*) c FROM submissions WHERE review_status='pending'"))["c"]; wd=(await fetchone("SELECT COUNT(*) c FROM withdrawals WHERE status='pending'"))["c"]; paid=(await fetchone("SELECT COALESCE(SUM(amount),0) s FROM withdrawals WHERE status='paid'"))["s"]; bal=(await fetchone("SELECT COALESCE(SUM(balance),0) s FROM users"))["s"]
    body=f'''<div class="grid"><div class="card">👥 Users<div class="num">{users}</div></div><div class="card">📥 Submissions<div class="num">{subs}</div></div><div class="card">⏳ Pending Submissions<div class="num">{pending}</div></div><div class="card">💳 Pending Payments<div class="num">{wd}</div></div><div class="card">💰 User Balance<div class="num">{CURRENCY}{fmt_money(bal)}</div></div><div class="card">💸 Total Paid<div class="num">{CURRENCY}{fmt_money(paid)}</div></div></div><br><div class="card"><b>Quick Actions</b><br><a class="btn" href="/submissions?status=pending">📥 Review Submissions</a><a class="btn green" href="/payments?status=pending">💳 Process Payments</a><a class="btn gray" href="/users">👥 Manage Users</a></div>'''
    return web.Response(text=web_html("Dashboard",body),content_type="text/html")

async def web_submissions(request):
    await web_auth(request); status=request.query.get("status",""); q=request.query.get("q","")
    sql="SELECT s.*,u.username,u.first_name FROM submissions s LEFT JOIN users u ON u.telegram_id=s.user_id WHERE 1=1"; params=[]
    if status: sql+=" AND s.review_status=?"; params.append(status)
    if q: sql+=" AND (s.public_id LIKE ? OR s.file_name LIKE ? OR CAST(s.user_id AS TEXT) LIKE ?)"; params += [f"%{q}%"]*3
    sql+=" ORDER BY s.id DESC LIMIT 100"; rows=await fetchall(sql,params)
    trs=''.join(f'<tr><td>{html.escape(r["public_id"] or str(r["id"]))}</td><td>{r["user_id"]}</td><td>{html.escape(r["username"] or r["first_name"] or "-")}</td><td>{html.escape(r["file_name"])}</td><td>{html.escape(r["category_name"])}</td><td>{r["review_status"]}</td><td>{CURRENCY}{fmt_money(r["final_amount"])}</td><td><a class="btn" href="/submission/{r["id"]}">Open</a></td></tr>' for r in rows)
    body=f'''<form><input class="input" name="q" value="{html.escape(q)}" placeholder="Search Submission ID / File / UID"><select class="select" name="status"><option value="">All Status</option><option value="pending">Pending</option><option value="approved">Approved</option><option value="rejected">Rejected</option></select><button class="btn">🔎 Search</button></form><div style="overflow:auto"><table class="table"><tr><th>ID</th><th>UID</th><th>User</th><th>File</th><th>Category</th><th>Status</th><th>Amount</th><th></th></tr>{trs or '<tr><td colspan="8">No submissions found.</td></tr>'}</table></div>'''
    return web.Response(text=web_html("Submissions",body,"submissions"),content_type="text/html")

async def web_submission_detail(request):
    await web_auth(request); sid=int(request.match_info["id"]); r=await fetchone("SELECT s.*,u.username,u.first_name,u.balance FROM submissions s LEFT JOIN users u ON u.telegram_id=s.user_id WHERE s.id=?",(sid,))
    if not r: raise web.HTTPNotFound()
    body=f'''<div class="card"><p><b>{html.escape(r["public_id"] or "-")}</b></p><p>👤 UID: {r["user_id"]} | @{html.escape(r["username"] or "none")}</p><p>📄 File: {html.escape(r["file_name"])}</p><p>📦 Category: {html.escape(r["category_name"])}</p><p>📊 Records: {r["records_count"]} | Valid: {r["valid_count"]}</p><p>💰 Rate: {CURRENCY}{fmt_money(r["locked_rate"])} | Calculated: {CURRENCY}{fmt_money(r["final_amount"])}</p><p>📌 Review: <b>{r["review_status"]}</b> | Payment: <b>{r["payment_status"]}</b></p><p>💳 Credited: {CURRENCY}{fmt_money(r["credited_amount"])}</p><p>🕐 Submitted: {html.escape(r["created_at"])}</p><a class="btn" href="/download-submission/{sid}">📥 Download File</a></div>'''
    if r["review_status"]=="pending": body += f'''<div class="card"><h3>Review</h3><form method="post" action="/submission/{sid}/approve"><input class="input" name="amount" placeholder="Credit amount (optional)"><button class="btn green">✅ Approve</button></form><form method="post" action="/submission/{sid}/reject"><input class="input" name="reason" placeholder="Reject reason" required><button class="btn red">❌ Reject</button></form></div>'''
    return web.Response(text=web_html("Submission Review",body,"submissions"),content_type="text/html")

async def web_download_submission(request):
    await web_auth(request); sid=int(request.match_info["id"]); r=await fetchone("SELECT file_id,file_name FROM submissions WHERE id=?",(sid,))
    if not r: raise web.HTTPNotFound()
    bot=Bot(BOT_TOKEN,default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    try:
        f=await bot.get_file(r["file_id"]); path=Path(tempfile.gettempdir())/f"sub_{sid}_{safe_filename(r['file_name'])}"; await bot.download_file(f.file_path,path)
        return web.FileResponse(path,headers={"Content-Disposition":f'attachment; filename="{safe_filename(r["file_name"])}"'})
    finally: await bot.session.close()

async def web_submission_approve(request):
    await web_auth(request); sid=int(request.match_info["id"]); data=await request.post(); r=await fetchone("SELECT * FROM submissions WHERE id=?",(sid,))
    if not r or r["review_status"]!="pending": raise web.HTTPBadRequest(text="Submission already processed")
    amount=money(data.get("amount") or r["final_amount"])
    await execute("UPDATE submissions SET review_status='approved',approved_quantity=?,final_amount=?,payment_status='unpaid',reviewed_at=? WHERE id=?",(r['valid_count'],float(amount),now_str(),sid)); await log_admin(OWNER_ID,"Web submission approved",r["public_id"],amount,r["public_id"])
    try:
        b=Bot(BOT_TOKEN); await b.send_message(r["user_id"],f"✅ <b>Submission Approved</b>\n\n📄 {r['public_id']}\n💰 Approved Amount: {CURRENCY}{fmt_money(amount)}"); await b.session.close()
    except Exception: pass
    raise web.HTTPFound(f'/submission/{sid}')

async def web_submission_reject(request):
    await web_auth(request); sid=int(request.match_info["id"]); data=await request.post(); reason=(data.get("reason") or "Rejected by Admin")[:500]; r=await fetchone("SELECT * FROM submissions WHERE id=?",(sid,))
    if not r or r["review_status"]!="pending": raise web.HTTPBadRequest(text="Submission already processed")
    await execute("UPDATE submissions SET review_status='rejected',payment_status='unpaid',rejection_reason=?,reviewed_at=? WHERE id=?",(reason,now_str(),sid)); await log_admin(OWNER_ID,"Web submission rejected",r["public_id"],None,r["public_id"]); raise web.HTTPFound(f'/submission/{sid}')

async def web_payments(request):
    await web_auth(request); status=request.query.get("status",""); rows=await fetchall("SELECT w.*,u.username,u.first_name FROM withdrawals w LEFT JOIN users u ON u.telegram_id=w.user_id WHERE (?='' OR w.status=?) ORDER BY w.id DESC LIMIT 100",(status,status))
    trs=''.join(f'<tr><td>{html.escape(r["public_id"] or "-")}</td><td>{r["user_id"]}</td><td>{html.escape(r["username"] or r["first_name"] or "-")}</td><td>{CURRENCY}{fmt_money(r["amount"])}</td><td>{html.escape(r["method"])}</td><td>{html.escape(r["number"][:2]+"XXXXXXXX"+r["number"][-3:] if len(r["number"])>5 else r["number"])}</td><td>{r["status"]}</td><td><a class="btn" href="/payment/{r["id"]}">Open</a></td></tr>' for r in rows)
    body=f'''<form><select class="select" name="status"><option value="">All</option><option value="pending">Pending</option><option value="paid">Paid</option><option value="rejected">Rejected</option></select><button class="btn">🔎 Filter</button></form><div style="overflow:auto"><table class="table"><tr><th>ID</th><th>UID</th><th>User</th><th>Amount</th><th>Method</th><th>Number</th><th>Status</th><th></th></tr>{trs or '<tr><td colspan="8">No payments found.</td></tr>'}</table></div>'''
    return web.Response(text=web_html("Payments",body,"payments"),content_type="text/html")

async def web_payment_detail(request):
    await web_auth(request); wid=int(request.match_info["id"]); r=await fetchone("SELECT w.*,u.username,u.first_name FROM withdrawals w LEFT JOIN users u ON u.telegram_id=w.user_id WHERE w.id=?",(wid,))
    if not r: raise web.HTTPNotFound()
    body=f'''<div class="card"><p><b>{html.escape(r["public_id"] or "-")}</b></p><p>👤 UID: {r["user_id"]} | @{html.escape(r["username"] or "none")}</p><p>💰 Amount: {CURRENCY}{fmt_money(r["amount"])}</p><p>💳 Method: {html.escape(r["method"])} | Number: {html.escape(r["number"])}</p><p>Status: <b>{r["status"]}</b></p><p>Created: {html.escape(r["created_at"])}</p></div>'''
    if r["status"]=="pending": body += f'''<div class="card"><form method="post" action="/payment/{wid}/reject"><button class="btn red">❌ Reject & Refund</button></form><p class="muted">Paid + TxID + Screenshot complete flow Telegram Admin Panel-এও available.</p></div>'''
    return web.Response(text=web_html("Payment Details",body,"payments"),content_type="text/html")

async def web_payment_reject(request):
    await web_auth(request); wid=int(request.match_info["id"]); r=await fetchone("SELECT * FROM withdrawals WHERE id=?",(wid,))
    if not r or r["status"]!="pending": raise web.HTTPBadRequest(text="Already processed")
    u=await fetchone("SELECT balance FROM users WHERE telegram_id=?",(r["user_id"],)); newbal=money(u["balance"])+money(r["amount"])
    await execute("UPDATE users SET balance=? WHERE telegram_id=?",(float(newbal),r["user_id"])); await execute("UPDATE withdrawals SET status='rejected',admin_id=?,processed_at=? WHERE id=?",(OWNER_ID,now_str(),wid)); await log_admin(OWNER_ID,"Web payment rejected/refunded",r["public_id"],r["amount"],r["public_id"]); raise web.HTTPFound(f'/payment/{wid}')

async def web_users(request):
    await web_auth(request); q=request.query.get("q",""); rows=await fetchall("SELECT * FROM users WHERE telegram_id LIKE ? OR username LIKE ? OR first_name LIKE ? ORDER BY id DESC LIMIT 100",(f"%{q}%",f"%{q}%",f"%{q}%")); trs=''.join(f'<tr><td>{r["telegram_id"]}</td><td>@{html.escape(r["username"] or "none")}</td><td>{html.escape(r["first_name"] or "-")}</td><td>{CURRENCY}{fmt_money(r["balance"])}</td><td>{"🚫 Banned" if r["banned"] else "🟢 Active"}</td><td><a class="btn" href="/user/{r["telegram_id"]}">Open</a></td></tr>' for r in rows); body=f'''<form><input class="input" name="q" value="{html.escape(q)}" placeholder="Search UID / username / name"><button class="btn">🔎 Search</button></form><div style="overflow:auto"><table class="table"><tr><th>UID</th><th>Username</th><th>Name</th><th>Balance</th><th>Status</th><th></th></tr>{trs or '<tr><td colspan="6">No users found.</td></tr>'}</table></div>'''; return web.Response(text=web_html("Users",body,"users"),content_type="text/html")

async def web_user_detail(request):
    await web_auth(request); uid=int(request.match_info["id"]); r=await fetchone("SELECT * FROM users WHERE telegram_id=?",(uid,))
    if not r: raise web.HTTPNotFound()
    body=f'''<div class="card"><p>🆔 <b>{uid}</b> | @{html.escape(r["username"] or "none")}</p><p>👤 {html.escape(r["first_name"] or "-")}</p><p>💰 Balance: <b>{CURRENCY}{fmt_money(r["balance"])}</b></p><p>⏳ Pending: {CURRENCY}{fmt_money(r["pending_balance"])}</p><p>💵 Total Earned: {CURRENCY}{fmt_money(r["total_earned"])}</p><p>💸 Total Paid: {CURRENCY}{fmt_money(r["total_paid"])}</p></div><div class="card form"><form method="post" action="/user/{uid}/balance"><input class="input" name="amount" placeholder="Amount" required><select class="select" name="action"><option value="add">Add Balance</option><option value="deduct">Deduct Balance</option></select><input class="input" name="reason" placeholder="Reason"><button class="btn green">💰 Apply</button></form><form method="post" action="/user/{uid}/ban"><button class="btn {'green' if r['banned'] else 'red'}">{'🟢 Unban' if r['banned'] else '🚫 Ban'} User</button></form></div>'''; return web.Response(text=web_html("User Management",body,"users"),content_type="text/html")

async def web_user_balance(request):
    await web_auth(request); uid=int(request.match_info["id"]); d=await request.post(); amount=money(d.get("amount")); action=d.get("action"); reason=(d.get("reason") or "Web Admin adjustment")[:200]; r=await fetchone("SELECT balance FROM users WHERE telegram_id=?",(uid,))
    if not r or amount<=0: raise web.HTTPBadRequest(text="Invalid amount")
    newbal=money(r["balance"])+(amount if action=="add" else -amount)
    if newbal<0: raise web.HTTPBadRequest(text="Insufficient balance")
    await execute("UPDATE users SET balance=?,total_earned=CASE WHEN ?='add' THEN total_earned+? ELSE total_earned END,updated_at=? WHERE telegram_id=?",(float(newbal),action,float(amount),now_str(),uid)); await execute("INSERT INTO transactions(public_id,user_id,txn_type,amount,balance_after,reason,source_id,admin_id,created_at) VALUES(?,?,?,?,?,?,?,?,?)",(f"TX-WEB-{secrets.token_hex(4).upper()}",uid,"admin_add" if action=="add" else "admin_deduct",float(amount),float(newbal),reason,None,OWNER_ID,now_str())); await log_admin(OWNER_ID,f"Web balance {action}",str(uid),amount,None); raise web.HTTPFound(f'/user/{uid}')

async def web_user_ban(request):
    await web_auth(request); uid=int(request.match_info["id"]); r=await fetchone("SELECT banned FROM users WHERE telegram_id=?",(uid,));
    if not r: raise web.HTTPNotFound()
    new=0 if r["banned"] else 1; await execute("UPDATE users SET banned=?,updated_at=? WHERE telegram_id=?",(new,now_str(),uid)); await log_admin(OWNER_ID,"Web user ban toggle",str(uid),None,None); raise web.HTTPFound(f'/user/{uid}')

async def web_categories(request):
    await web_auth(request); rows=await fetchall("SELECT * FROM categories ORDER BY sort_order,id"); trs=''.join(f'<tr><td>{r["id"]}</td><td>{r["emoji"]} {html.escape(r["name"])}</td><td>{CURRENCY}{fmt_money(r["rate"])}</td><td>{"🟢 Enabled" if r["enabled"] else "🔴 Disabled"}</td><td><a class="btn" href="/category/{r["id"]}">Edit</a></td></tr>' for r in rows); body=f'''<a class="btn green" href="/category/new">➕ Add Category</a><div style="overflow:auto"><table class="table"><tr><th>ID</th><th>Name</th><th>Rate</th><th>Status</th><th></th></tr>{trs}</table></div>'''; return web.Response(text=web_html("Categories",body,"categories"),content_type="text/html")

async def web_category_form(request):
    await web_auth(request); cid=request.match_info["id"]; r=None if cid=="new" else await fetchone("SELECT * FROM categories WHERE id=?",(int(cid),))
    if cid!="new" and not r: raise web.HTTPNotFound()
    name=r["name"] if r else ''; rate=r["rate"] if r else ''; emoji=r["emoji"] if r else '📦'; enabled=r["enabled"] if r else 1; order=r["sort_order"] if r else 0
    body=f'''<div class="card form"><form method="post" action="/category/{cid}/save"><label>Name</label><input class="input" name="name" value="{html.escape(str(name))}" required><label>Rate</label><input class="input" name="rate" value="{rate}" required><label>Emoji</label><input class="input" name="emoji" value="{html.escape(emoji)}"><label>Position</label><input class="input" name="sort_order" value="{order}"><label>Status</label><select class="select" name="enabled"><option value="1" {'selected' if enabled else ''}>Enabled</option><option value="0" {'selected' if not enabled else ''}>Disabled</option></select><button class="btn green">💾 Save</button></form></div>'''; return web.Response(text=web_html("Category Editor",body,"categories"),content_type="text/html")

async def web_category_save(request):
    await web_auth(request); cid=request.match_info["id"]; d=await request.post(); name=(d.get("name") or "").strip()[:80]; rate=money(d.get("rate")); emoji=(d.get("emoji") or "📦")[:8]; order=int(d.get("sort_order") or 0); enabled=1 if d.get("enabled")=="1" else 0
    if not name: raise web.HTTPBadRequest(text="Name required")
    if cid=="new": await execute("INSERT INTO categories(name,rate,enabled,emoji,button_style,sort_order,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?)",(name,float(rate),enabled,emoji,"primary",order,now_str(),now_str()))
    else: await execute("UPDATE categories SET name=?,rate=?,enabled=?,emoji=?,sort_order=?,updated_at=? WHERE id=?",(name,float(rate),enabled,emoji,order,now_str(),int(cid)))
    await log_admin(OWNER_ID,"Web category saved",name,rate,None); raise web.HTTPFound('/categories')

async def web_balance(request):
    await web_auth(request); body='<div class="card"><p>💰 Balance Management: individual user profile থেকে Add/Deduct করা যাবে।</p><a class="btn" href="/users">👥 Open Users</a></div>'; return web.Response(text=web_html("Balance Management",body,"balance"),content_type="text/html")

async def web_statistics(request):
    await web_auth(request); today=datetime.now(TIMEZONE).strftime("%Y-%m-%d"); u=(await fetchone("SELECT COUNT(*) c FROM users"))["c"]; s=(await fetchone("SELECT COUNT(*) c FROM submissions"))["c"]; sa=(await fetchone("SELECT COUNT(*) c FROM submissions WHERE review_status='approved'"))["c"]; sr=(await fetchone("SELECT COUNT(*) c FROM submissions WHERE review_status='rejected'"))["c"]; st=(await fetchone("SELECT COUNT(*) c FROM submissions WHERE substr(created_at,1,10)=?",(today,)))["c"]; paid=(await fetchone("SELECT COALESCE(SUM(amount),0) s FROM withdrawals WHERE status='paid'"))["s"]; body=f'''<div class="grid"><div class="card">👥 Users<div class="num">{u}</div></div><div class="card">📥 Submissions<div class="num">{s}</div></div><div class="card">📅 Today<div class="num">{st}</div></div><div class="card">✅ Approved<div class="num">{sa}</div></div><div class="card">❌ Rejected<div class="num">{sr}</div></div><div class="card">💸 Paid<div class="num">{CURRENCY}{fmt_money(paid)}</div></div></div>'''; return web.Response(text=web_html("Statistics",body,"statistics"),content_type="text/html")

async def web_broadcast(request):
    await web_auth(request)
    if request.method=="POST":
        d=await request.post(); body=(d.get("body") or "").strip()[:4000]
        if body:
            rows=await fetchall("SELECT telegram_id FROM users WHERE banned=0"); sent=failed=0; b=Bot(BOT_TOKEN,default=DefaultBotProperties(parse_mode=ParseMode.HTML))
            try:
                for r in rows:
                    try: await b.send_message(r["telegram_id"],"📢 BROADCAST\n\n"+body); sent+=1
                    except Exception: failed+=1
                    await asyncio.sleep(0.04)
            finally: await b.session.close()
            await execute("INSERT INTO broadcasts(admin_id,body,recipients,sent,failed,created_at) VALUES(?,?,?,?,?,?)",(OWNER_ID,body,len(rows),sent,failed,now_str())); await log_admin(OWNER_ID,"Web broadcast sent",None,None,None); return web.Response(text=web_html("Broadcast",f'<div class="alert">✅ Sent: {sent} | Failed: {failed}</div><a class="btn" href="/broadcast">Back</a>',"broadcast"),content_type="text/html")
    body='''<div class="card form"><form method="post"><label>Broadcast Message</label><textarea class="textarea" name="body" rows="8" required></textarea><button class="btn green">📢 Send Broadcast</button></form></div>'''; return web.Response(text=web_html("Broadcast",body,"broadcast"),content_type="text/html")

async def web_settings(request):
    await web_auth(request)
    if request.method=="POST":
        d=await request.post()
        for key in ("min_withdraw","maintenance","submission_window_enabled","submission_alerts_enabled","submission_start","submission_end"):
            if key in d: await set_setting(key,str(d[key]))
        await log_admin(OWNER_ID,"Web settings updated",None,None,None)
    vals={k:await get_setting(k,"") for k in ("min_withdraw","maintenance","submission_window_enabled","submission_alerts_enabled","submission_start","submission_end","payment_proof_chat_title","force_join_chat_title")}
    body=f'''<div class="card form"><form method="post"><label>Minimum Withdrawal</label><input class="input" name="min_withdraw" value="{vals['min_withdraw']}"><label>Maintenance</label><select class="select" name="maintenance"><option value="0" {'selected' if vals['maintenance']!='1' else ''}>OFF</option><option value="1" {'selected' if vals['maintenance']=='1' else ''}>ON</option></select><label>Submission Window</label><select class="select" name="submission_window_enabled"><option value="1" {'selected' if vals['submission_window_enabled']=='1' else ''}>ON</option><option value="0" {'selected' if vals['submission_window_enabled']!='1' else ''}>OFF</option></select><label>Start</label><input class="input" name="submission_start" value="{vals['submission_start']}"><label>End</label><input class="input" name="submission_end" value="{vals['submission_end']}"><label>Deadline Alerts</label><select class="select" name="submission_alerts_enabled"><option value="1" {'selected' if vals['submission_alerts_enabled']=='1' else ''}>ON</option><option value="0" {'selected' if vals['submission_alerts_enabled']!='1' else ''}>OFF</option></select><button class="btn green">💾 Save Settings</button></form></div><div class="card"><p>📢 Payment Proof: <b>{html.escape(vals['payment_proof_chat_title'] or 'Not selected')}</b></p><p>🔐 Force Join: <b>{html.escape(vals['force_join_chat_title'] or 'Not selected')}</b></p><p class="muted">Channel/Group picker এবং Force Join selection Telegram Admin Panel থেকে করা যাবে।</p></div>'''; return web.Response(text=web_html("Settings",body,"settings"),content_type="text/html")

async def web_admins(request):
    await web_auth(request); rows=await fetchall("SELECT * FROM admins ORDER BY created_at DESC"); trs=''.join(f'<tr><td>{r["user_id"]}</td><td>{html.escape(r["role"])}</td><td>{r["created_at"]}</td></tr>' for r in rows); body=f'''<div class="card"><p>Owner: <b>{OWNER_ID}</b></p></div><table class="table"><tr><th>User ID</th><th>Role</th><th>Added</th></tr>{trs}</table>'''; return web.Response(text=web_html("Admin Management",body,"admins"),content_type="text/html")

async def web_logs(request):
    await web_auth(request); rows=await fetchall("SELECT * FROM admin_logs ORDER BY id DESC LIMIT 100"); trs=''.join(f'<tr><td>{r["created_at"]}</td><td>{r["admin_id"]}</td><td>{html.escape(r["action"])}</td><td>{html.escape(str(r["target"] or "-"))}</td><td>{r["amount"] if r["amount"] is not None else "-"}</td></tr>' for r in rows); body=f'''<div style="overflow:auto"><table class="table"><tr><th>Time</th><th>Admin</th><th>Action</th><th>Target</th><th>Amount</th></tr>{trs or '<tr><td colspan="5">No logs.</td></tr>'}</table></div>'''; return web.Response(text=web_html("Admin Logs",body,"logs"),content_type="text/html")

async def start_web_admin():
    if web is None: log.warning("aiohttp is not installed; Web Admin Panel disabled."); return None
    app=web.Application(client_max_size=30*1024*1024)
    app.router.add_get('/login',web_login); app.router.add_post('/login',web_login); app.router.add_get('/logout',web_logout)
    app.router.add_get('/',web_dashboard); app.router.add_get('/submissions',web_submissions); app.router.add_get('/submission/{id}',web_submission_detail); app.router.add_get('/download-submission/{id}',web_download_submission); app.router.add_post('/submission/{id}/approve',web_submission_approve); app.router.add_post('/submission/{id}/reject',web_submission_reject)
    app.router.add_get('/payments',web_payments); app.router.add_get('/payment/{id}',web_payment_detail); app.router.add_post('/payment/{id}/reject',web_payment_reject)
    app.router.add_get('/users',web_users); app.router.add_get('/user/{id}',web_user_detail); app.router.add_post('/user/{id}/balance',web_user_balance); app.router.add_post('/user/{id}/ban',web_user_ban)
    app.router.add_get('/categories',web_categories); app.router.add_get('/category/{id}',web_category_form); app.router.add_post('/category/{id}/save',web_category_save); app.router.add_get('/balance',web_balance); app.router.add_get('/statistics',web_statistics)
    app.router.add_get('/broadcast',web_broadcast); app.router.add_post('/broadcast',web_broadcast); app.router.add_get('/settings',web_settings); app.router.add_post('/settings',web_settings); app.router.add_get('/admins',web_admins); app.router.add_get('/logs',web_logs)
    runner=web.AppRunner(app); await runner.setup(); await web.TCPSite(runner,WEB_HOST,WEB_PORT).start(); log.info("Web Admin Panel started at http://%s:%s",WEB_HOST,WEB_PORT); return runner

# ============================================================
# 29. STARTUP / MAIN
# ============================================================
async def main():
    await db_init()
    bot = Bot(BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dp = Dispatcher()
    dp.include_router(router)
    alert_task = None
    web_runner = None
    log.info("Apex Valorem starting | Telegram Admin Panel + Web Admin Panel enabled | timezone=Asia/Dhaka | owner=%s", OWNER_ID)
    try:
        web_runner = await start_web_admin()
        alert_task = asyncio.create_task(daily_submission_alert_scheduler(bot))
        await bot.delete_webhook(drop_pending_updates=False)
        await dp.start_polling(bot, allowed_updates=dp.resolve_used_update_types())
    finally:
        if web_runner:
            await web_runner.cleanup()
        if alert_task:
            alert_task.cancel()
            try:
                await alert_task
            except asyncio.CancelledError:
                pass
        await bot.session.close()
        if DB:
            await DB.close()

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        log.info("Bot stopped by user.")
