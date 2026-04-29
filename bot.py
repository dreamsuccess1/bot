

import asyncio
import logging
import re
import time
import io
from datetime import datetime

import openpyxl
from telegram import Update, Poll, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, PollAnswerHandler,
    ConversationHandler, filters, ContextTypes
)
from telegram.constants import ParseMode
from telegram.error import TelegramError

import database as db
from pdf_generator import generate_result_pdf
from config import BOT_TOKEN, ADMIN_IDS, BOT_NAME, BOT_USER, TARGET_TXT, TIMERS

logging.basicConfig(format="%(asctime)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# ── Conversation States ────────────────────────────────────────────────────────
(
    MANUAL_QUESTION, MANUAL_OPTION_A, MANUAL_OPTION_B,
    MANUAL_OPTION_C, MANUAL_OPTION_D, MANUAL_CORRECT,
    MANUAL_EXPLANATION, MANUAL_TIMER, SET_NAME,
    PARSED_SET_SELECT, PHOTO_CORRECT, PHOTO_SET_SELECT,
    CREATE_WAITING_NAME,
    CREATE_COLLECTING,
    # Marking setup
    MARKING_POS, MARKING_NEG,
    # Section setup
    SECTION_NAME, SECTION_TIMER, SECTION_CONFIRM,
    # /create with per-question timer prompt
    CREATE_Q_TIMER,
) = range(20)

POLL_TO_CHAT: dict = {}

# ── Helpers ───────────────────────────────────────────────────────────────────

def is_admin(uid: int) -> bool:
    return int(uid) in [int(a) for a in ADMIN_IDS]

def timer_kb(extra_options=None):
    base = TIMERS if TIMERS else [10, 15, 20, 25, 30, 45, 60]
    if extra_options:
        base = extra_options
    row = [InlineKeyboardButton(f"⏱ {t}s", callback_data=f"timer_{t}") for t in base]
    # split into rows of 4
    rows = [row[i:i+4] for i in range(0, len(row), 4)]
    return InlineKeyboardMarkup(rows)

def section_timer_kb():
    mins = [1, 2, 3, 5, 10, 15, 20, 30, 45, 60]
    rows = []
    row = []
    for m in mins:
        row.append(InlineKeyboardButton(f"⏱ {m}m", callback_data=f"sectimer_{m*60}"))
        if len(row) == 4:
            rows.append(row)
            row = []
    if row:
        rows.append(row)
    return InlineKeyboardMarkup(rows)

def option_kb(options: list, prefix="correct"):
    labels = ["A","B","C","D"]
    return InlineKeyboardMarkup([[
        InlineKeyboardButton(f"{labels[i]}: {str(o)[:20]}", callback_data=f"{prefix}_{i}")
    ] for i, o in enumerate(options)])

def fmt_time(seconds: float) -> str:
    m, s = divmod(int(seconds), 60)
    return f"{m}m {s}s"

def fmt_score(score: float) -> str:
    if score == int(score):
        return str(int(score))
    return f"{score:.2f}"

def calc_acc(correct: int, total: int) -> int:
    return round((correct / total) * 100) if total > 0 else 0

# ── Format Parsers ─────────────────────────────────────────────────────────────

def parse_create_format(text: str):
    lines = [l.strip() for l in text.strip().splitlines() if l.strip()]
    if len(lines) < 3:
        return None
    CORRECT_MARKERS = ["✅", "[✅]", "(✅)", "(correct)", "(ans)", "(answer)"]
    question_lines = []
    option_lines   = []
    for line in lines:
        is_opt = bool(re.match(r'^(\d[\.\)]\s*|[A-Da-d][\.\)]\s*)', line))
        if is_opt or option_lines:
            option_lines.append(line)
        else:
            question_lines.append(line)
    if not option_lines and len(lines) >= 5:
        question_lines = [lines[0]]
        option_lines   = lines[1:]
    if len(option_lines) < 2:
        return None
    question = " ".join(question_lines).strip()
    options  = []
    correct  = -1
    for i, opt in enumerate(option_lines[:4]):
        is_correct = False
        clean_opt  = opt
        for marker in CORRECT_MARKERS:
            if marker in opt:
                is_correct = True
                clean_opt  = opt.replace(marker, "").strip()
                break
        clean_opt = re.sub(r'^(\d[\.\)]\s*|[A-Da-d][\.\)]\s*)', '', clean_opt).strip()
        clean_opt = clean_opt.rstrip(".")
        options.append(clean_opt)
        if is_correct:
            correct = i
    if len(options) < 2 or correct == -1:
        return None
    return question, options, correct


def parse_text_question(text: str):
    text  = text.strip()
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    if len(lines) < 5:
        return None
    q_match   = re.search(r'^Q[\.\:][\s]*(.*)', text, re.MULTILINE | re.IGNORECASE)
    a_match   = re.search(r'^[Aa][\.\:][\s]*(.*)', text, re.MULTILINE)
    b_match   = re.search(r'^[Bb][\.\:][\s]*(.*)', text, re.MULTILINE)
    c_match   = re.search(r'^[Cc][\.\:][\s]*(.*)', text, re.MULTILINE)
    d_match   = re.search(r'^[Dd][\.\:][\s]*(.*)', text, re.MULTILINE)
    ans_match = re.search(
        r'^(?:Ans|Answer|Correct|Ans\.|उत्तर|सही जवाब)[\s\.\:]+([A-D1-4])',
        text, re.MULTILINE | re.IGNORECASE)
    exp_match = re.search(
        r'^(?:Exp|Explanation|Expl|व्याख्या)[\s\.\:]+(.+)',
        text, re.MULTILINE | re.IGNORECASE)
    if q_match and a_match and b_match and c_match and d_match and ans_match:
        question = q_match.group(1).strip()
        options  = [a_match.group(1).strip(), b_match.group(1).strip(),
                    c_match.group(1).strip(), d_match.group(1).strip()]
        ans_raw  = ans_match.group(1).strip().upper()
        ans_map  = {"A":0,"B":1,"C":2,"D":3,"1":0,"2":1,"3":2,"4":3}
        correct  = ans_map.get(ans_raw, -1)
        expl     = exp_match.group(1).strip() if exp_match else ""
        if correct >= 0 and all(options):
            return {"question":question,"options":options,"correct":correct,"explanation":expl}
    num_opts = re.findall(r'^[1-4][\.\)]\s*(.+)', text, re.MULTILINE)
    ans2     = re.search(r'^(?:Ans|Answer|Correct|उत्तर)[\s\.\:]+([1-4A-D])',
                         text, re.MULTILINE | re.IGNORECASE)
    if len(num_opts) >= 4 and ans2:
        question = lines[0]
        options  = num_opts[:4]
        ans_raw  = ans2.group(1).strip().upper()
        ans_map  = {"1":0,"2":1,"3":2,"4":3,"A":0,"B":1,"C":2,"D":3}
        correct  = ans_map.get(ans_raw, -1)
        expl     = exp_match.group(1).strip() if exp_match else ""
        if correct >= 0 and all(options):
            return {"question":question,"options":options,"correct":correct,"explanation":expl}
    return None


# ── /start & /help ─────────────────────────────────────────────────────────────

async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    db.register_user(user.id, user.full_name, user.username)

    # Student message
    student_text = (
        f"🎯 *{BOT_NAME} में आपका स्वागत है!*\n\n"
        f"📌 {TARGET_TXT}\n\n"
        "Quiz में participate करने के लिए तैयार रहें।\n"
        "_(PDF result पाने के लिए bot को एक बार /start करना ज़रूरी है)_"
    )

    # Admin full command list
    admin_text = (
        f"🎯 *{BOT_NAME} — Admin Panel*\n"
        f"📌 {TARGET_TXT}\n"
        + "━"*30 + "\n\n"

        "📝 *QUIZ BANANE KE COMMANDS:*\n"
        "/create — नई Quiz बनाएं\n"
        "   ↳ ✅ format, marking scheme, sections, per-Q timer\n"
        "/newquiz — Step-by-step एक सवाल बनाएं\n"
        "/bulkupload — Excel (.xlsx) से bulk upload\n\n"

        "📂 *QUIZ MANAGE KARNE KE COMMANDS:*\n"
        "/sets — सभी Quiz sets देखें\n"
        "/startquiz — Quiz शुरू करें (set चुनें)\n"
        "/stopquiz — चल रही Quiz रोकें\n\n"

        "🏆 *LEADERBOARD COMMANDS:*\n"
        "/leaderboard — Current leaderboard देखें\n"
        "/resetscores — Leaderboard reset करें\n\n"

        "━"*30 + "\n"
        "📌 *QUESTION FORMAT (✅ wala correct):*\n"
        "`सवाल यहाँ?`\n"
        "`A. Option A`\n"
        "`B. Option B ✅`\n"
        "`C. Option C`\n"
        "`D. Option D`\n\n"

        "📌 *BULK FORMAT (.txt):*\n"
        "`Q. सवाल?`\n"
        "`A. Option A`\n"
        "`B. Option B`\n"
        "`C. Option C`\n"
        "`D. Option D ✅`\n\n"

        "📌 *EXCEL FORMAT (.xlsx):*\n"
        "`Question|A|B|C|D|Correct(0-3)|Explanation|Timer`\n\n"

        "📌 *FORWARDED POLL:* Group se poll forward karo, save ho jaega\n\n"

        "━"*30 + "\n"
        "⚙️ *FEATURES:*\n"
        "✅ Forward karke quiz save\n"
        "✅ Negative marking (+/- custom)\n"
        "✅ Auto leaderboard + PDF\n"
        "✅ Sectional quiz (alag-alag timers)\n"
        "✅ Per-question timer\n"
        "✅ Questions forward-proof\n"
        "✅ Unlimited questions\n"
        "✅ Bulk upload (.txt/.xlsx)"
    )

    if is_admin(user.id):
        await update.message.reply_text(admin_text, parse_mode=ParseMode.MARKDOWN)
    else:
        await update.message.reply_text(student_text, parse_mode=ParseMode.MARKDOWN)


async def help_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await start(update, ctx)


# ── /create Flow (with per-question timer + sections + marking) ───────────────

async def create_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        return ConversationHandler.END
    ctx.user_data.clear()
    await update.message.reply_text(
        "📝 *नई Quiz बनाएं*\n\nQuiz का नाम भेजें:",
        parse_mode=ParseMode.MARKDOWN
    )
    return CREATE_WAITING_NAME


async def create_recv_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    name = update.message.text.strip()
    ctx.user_data["create_name"] = name
    # Ask marking scheme
    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ +1 / ❌ 0 (No Negative)", callback_data="mark_1_0")],
        [InlineKeyboardButton("✅ +2 / ❌ -0.5", callback_data="mark_2_0.5")],
        [InlineKeyboardButton("✅ +4 / ❌ -1", callback_data="mark_4_1")],
        [InlineKeyboardButton("✅ +1 / ❌ -0.25", callback_data="mark_1_0.25")],
        [InlineKeyboardButton("🔢 Custom marks set करें", callback_data="mark_custom")],
    ])
    await update.message.reply_text(
        f"✅ Name: *{name}*\n\n📊 *Marking Scheme चुनें:*",
        parse_mode=ParseMode.MARKDOWN,
        reply_markup=kb
    )
    return MARKING_POS


async def create_marking_choice(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    if data == "mark_custom":
        await query.message.reply_text(
            "✅ *Correct answer के marks* टाइप करें (जैसे: 2 या 1.5):"
            , parse_mode=ParseMode.MARKDOWN)
        return MARKING_POS
    # parse preset: mark_POS_NEG
    parts = data.split("_")
    pos = float(parts[1])
    neg = float(parts[2])
    return await _finish_marking(query.message, ctx, pos, neg)


async def create_recv_pos_marks(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    try:
        pos = float(update.message.text.strip())
    except ValueError:
        await update.message.reply_text("❌ सिर्फ number डालें जैसे: 1 या 2")
        return MARKING_POS
    ctx.user_data["pos_marks"] = pos
    await update.message.reply_text(
        f"✅ Correct = +{pos}\n\n❌ *Wrong answer पर कितने marks काटें?*\n"
        f"_(0 लिखें अगर कोई negative marking नहीं चाहिए)_",
        parse_mode=ParseMode.MARKDOWN
    )
    return MARKING_NEG


async def create_recv_neg_marks(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    try:
        neg = float(update.message.text.strip())
    except ValueError:
        await update.message.reply_text("❌ सिर्फ number डालें जैसे: 0 या 0.25")
        return MARKING_NEG
    return await _finish_marking(update.message, ctx, ctx.user_data.get("pos_marks", 1.0), neg)


async def _finish_marking(msg, ctx, pos: float, neg: float):
    ctx.user_data["pos_marks"] = pos
    ctx.user_data["neg_marks"] = neg
    name   = ctx.user_data["create_name"]
    # Safe call — purana database.py bhi support karta hai
    try:
        set_id = db.create_set(name, neg_marks=neg, pos_marks=pos)
    except TypeError:
        # Purana database.py installed hai — bina marks ke create karo
        set_id = db.create_set(name)
        # Baad mein update karo agar function exist karta hai
        try:
            db.update_set_marks(set_id, pos, neg)
        except Exception:
            pass
    ctx.user_data["create_set_id"]   = set_id
    ctx.user_data["create_set_name"] = name
    ctx.user_data["create_count"]    = 0
    ctx.user_data["current_section"] = None

    neg_str = f"❌ -{neg}" if neg > 0 else "❌ No Negative"
    # Ask about sections
    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("📂 Sections बनाएं (Sectional Quiz)", callback_data="use_sections_yes")],
        [InlineKeyboardButton("📋 Normal Quiz (No Sections)", callback_data="use_sections_no")],
    ])
    await msg.reply_text(
        f"✅ *Quiz: {name}*\n"
        f"📊 Marking: +{pos} / {neg_str}\n\n"
        f"*क्या इस Quiz में Sections होंगे?*",
        parse_mode=ParseMode.MARKDOWN,
        reply_markup=kb
    )
    return SECTION_CONFIRM


async def create_section_confirm(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "use_sections_yes":
        set_id = ctx.user_data["create_set_id"]
        db.set_has_sections(set_id, True)
        ctx.user_data["section_count"] = 0
        await query.message.reply_text(
            "📂 *Section का नाम भेजें* (जैसे: General Knowledge, Science, Math):",
            parse_mode=ParseMode.MARKDOWN
        )
        return SECTION_NAME
    else:
        return await _start_collecting(query.message, ctx)


async def create_recv_section_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    name = update.message.text.strip()
    ctx.user_data["new_section_name"] = name
    await update.message.reply_text(
        f"⏱ *'{name}' section का time कितना होगा?*\n\n"
        f"_(यह पूरे section के लिए total time है)_",
        parse_mode=ParseMode.MARKDOWN,
        reply_markup=section_timer_kb()
    )
    return SECTION_TIMER


async def create_recv_section_timer(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    sec_timer = int(query.data.split("_")[1])
    set_id    = ctx.user_data["create_set_id"]
    sec_name  = ctx.user_data["new_section_name"]
    count     = ctx.user_data.get("section_count", 0)
    sec_id    = db.create_section(set_id, sec_name, timer=sec_timer, position=count)
    ctx.user_data["current_section"] = sec_id
    ctx.user_data["current_section_name"] = sec_name
    ctx.user_data["section_count"] = count + 1
    m, s = divmod(sec_timer, 60)
    await query.message.reply_text(
        f"✅ *Section: {sec_name}* ({m}m {s}s)\n\n"
        f"अब इस section के questions भेजें।\n"
        f"_नया section — /newsection_\nQuiz पूरी — /done_",
        parse_mode=ParseMode.MARKDOWN
    )
    return CREATE_COLLECTING


async def _start_collecting(msg, ctx):
    pos = ctx.user_data.get("pos_marks", 1.0)
    neg = ctx.user_data.get("neg_marks", 0.0)
    neg_str = f"❌ -{neg}" if neg > 0 else "❌ No Negative"
    await msg.reply_text(
        f"✅ *Quiz Ready — Questions भेजें!*\n\n"
        f"📊 Marking: ✅ +{pos} / {neg_str}\n\n"
        f"*Format:*\n"
        f"`सवाल?\nOption A.\nOption B. ✅\nOption C.\nOption D.`\n\n"
        f"✅ wala = correct answer\n"
        f"Poll/Forwarded poll/.txt/.xlsx सब चलेगा\n\n"
        f"/done — Quiz पूरी करें",
        parse_mode=ParseMode.MARKDOWN
    )
    return CREATE_COLLECTING


async def create_new_section(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Admin sends /newsection during collecting to add next section"""
    set_id = ctx.user_data.get("create_set_id")
    if not set_id:
        return CREATE_COLLECTING
    set_info = db.get_set(set_id)
    if not set_info or not set_info.get("has_sections"):
        await update.message.reply_text("❌ यह Quiz sectional नहीं है।")
        return CREATE_COLLECTING
    await update.message.reply_text(
        "📂 *नए Section का नाम भेजें:*",
        parse_mode=ParseMode.MARKDOWN
    )
    return SECTION_NAME


async def create_recv_question(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msg    = update.message
    set_id = ctx.user_data.get("create_set_id")
    count  = ctx.user_data.get("create_count", 0)
    section_id = ctx.user_data.get("current_section")

    if not set_id:
        await msg.reply_text("❌ Pehle /create se shuru karein।")
        return ConversationHandler.END

    # /done
    if msg.text and msg.text.strip().lower() in ["/done", "done"]:
        name = ctx.user_data.get("create_set_name", "Quiz")
        pos  = ctx.user_data.get("pos_marks", 1.0)
        neg  = ctx.user_data.get("neg_marks", 0.0)
        neg_str = f"❌ -{neg}" if neg > 0 else "No Negative"
        await msg.reply_text(
            f"✅ *Quiz Ready!*\n"
            f"📚 {name}\n"
            f"❓ {count} questions saved\n"
            f"📊 +{pos} / {neg_str}\n\n"
            f"/startquiz se shuru karein!",
            parse_mode=ParseMode.MARKDOWN
        )
        ctx.user_data.clear()
        return ConversationHandler.END

    # /newsection redirect
    if msg.text and msg.text.strip().lower() == "/newsection":
        return await create_new_section(update, ctx)

    # Poll (forwarded ya nahi) — Feature 1 & 4
    if msg.poll:
        poll = msg.poll
        if poll.type != Poll.QUIZ or poll.correct_option_id is None:
            await msg.reply_text("⚠️ Sirf Quiz type polls (with correct answer) bhejein।")
            return CREATE_COLLECTING
        options  = [o.text for o in poll.options]
        correct  = poll.correct_option_id
        question = poll.question
        expl     = poll.explanation or ""
        # Ask timer for this question
        ctx.user_data["pending_q"] = {
            "question": question, "options": options, "correct": correct,
            "explanation": expl, "photo_id": None, "section_id": section_id
        }
        labels = ["A","B","C","D"]
        await msg.reply_text(
            f"✅ Poll received!\n❓ {question}\n✔️ Correct: *{labels[correct]}*\n\n"
            f"⏱ *इस question का timer चुनें:*",
            parse_mode=ParseMode.MARKDOWN,
            reply_markup=timer_kb()
        )
        return CREATE_Q_TIMER

    # .txt file — bulk Feature 5
    if msg.document and msg.document.file_name and msg.document.file_name.endswith(".txt"):
        file = await ctx.bot.get_file(msg.document.file_id)
        buf  = io.BytesIO()
        await file.download_to_memory(buf)
        raw    = buf.getvalue().decode("utf-8", errors="ignore")
        blocks = [b.strip() for b in re.split(r'\n\s*\n', raw) if b.strip()]
        saved, failed = 0, 0
        for block in blocks:
            result = parse_create_format(block)
            if result:
                q, opts, corr = result
                db.add_question(set_id=set_id, question=q, options=opts,
                                correct=corr, explanation="", timer=20, section_id=section_id)
                saved += 1
                continue
            result_q = parse_text_question(block)
            if result_q:
                db.add_question(set_id=set_id, question=result_q["question"],
                                options=result_q["options"], correct=result_q["correct"],
                                explanation=result_q.get("explanation",""), timer=20,
                                section_id=section_id)
                saved += 1
            else:
                failed += 1
        count += saved
        ctx.user_data["create_count"] = count
        sec_name = ctx.user_data.get("current_section_name", "")
        sec_info = f"\n📂 Section: {sec_name}" if sec_name else ""
        await msg.reply_text(
            f"✅ *{saved} questions saved from .txt!*{sec_info}\n"
            f"❌ {failed} skip हुए\n\nTotal: {count} | Send more or /done",
            parse_mode=ParseMode.MARKDOWN
        )
        return CREATE_COLLECTING

    # Photo + caption
    if msg.photo:
        caption  = (msg.caption or "").strip()
        photo_id = msg.photo[-1].file_id
        result   = parse_create_format(caption) if caption else None
        if result:
            q, opts, corr = result
            ctx.user_data["pending_q"] = {
                "question": q, "options": opts, "correct": corr,
                "explanation": "", "photo_id": photo_id, "section_id": section_id
            }
            labels = ["A","B","C","D"]
            await msg.reply_text(
                f"✅ Photo question!\n✔️ Correct: *{labels[corr]}*\n\n⏱ *Timer चुनें:*",
                parse_mode=ParseMode.MARKDOWN,
                reply_markup=timer_kb()
            )
            return CREATE_Q_TIMER
        else:
            await msg.reply_text(
                "⚠️ Photo caption mein ✅ format nahi mila।\n\n"
                "Format:\n`सवाल?\nOption A.\nOption B. ✅\nOption C.\nOption D.`",
                parse_mode=ParseMode.MARKDOWN
            )
        return CREATE_COLLECTING

    # Plain text
    text = (msg.text or "").strip()
    if not text:
        return CREATE_COLLECTING

    # ✅ format
    result = parse_create_format(text)
    if result:
        q, opts, corr = result
        ctx.user_data["pending_q"] = {
            "question": q, "options": opts, "correct": corr,
            "explanation": "", "photo_id": None, "section_id": section_id
        }
        labels = ["A","B","C","D"]
        await msg.reply_text(
            f"✅ Question parsed!\n✔️ Correct: *{labels[corr]}* — {opts[corr]}\n\n"
            f"⏱ *इस question का timer चुनें:*",
            parse_mode=ParseMode.MARKDOWN,
            reply_markup=timer_kb()
        )
        return CREATE_Q_TIMER

    # Q:/A: format
    parsed = parse_text_question(text)
    if parsed:
        ctx.user_data["pending_q"] = {
            "question": parsed["question"], "options": parsed["options"],
            "correct": parsed["correct"], "explanation": parsed.get("explanation",""),
            "photo_id": None, "section_id": section_id
        }
        labels = ["A","B","C","D"]
        await msg.reply_text(
            f"✅ Question parsed!\n✔️ Correct: *{labels[parsed['correct']]}*\n\n"
            f"⏱ *इस question का timer चुनें:*",
            parse_mode=ParseMode.MARKDOWN,
            reply_markup=timer_kb()
        )
        return CREATE_Q_TIMER

    await msg.reply_text(
        "⚠️ Format samajh nahi aaya।\n\n"
        "*Sahi format:*\n"
        "```\nसवाल यहाँ?\nOption A.\nOption B. ✅\nOption C.\nOption D.\n```\n"
        "✅ wala option correct maana jaega।\n\nSend more or /done",
        parse_mode=ParseMode.MARKDOWN
    )
    return CREATE_COLLECTING


async def create_recv_q_timer(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Receive timer for pending question from /create flow"""
    query = update.callback_query
    await query.answer()
    timer  = int(query.data.split("_")[1])
    pq     = ctx.user_data.get("pending_q", {})
    set_id = ctx.user_data.get("create_set_id")
    if not pq or not set_id:
        await query.message.reply_text("❌ Error — /create se dobara shuru karein।")
        return ConversationHandler.END
    db.add_question(
        set_id=set_id, question=pq["question"], options=pq["options"],
        correct=pq["correct"], explanation=pq.get("explanation",""),
        timer=timer, photo_id=pq.get("photo_id"),
        section_id=pq.get("section_id")
    )
    count = ctx.user_data.get("create_count", 0) + 1
    ctx.user_data["create_count"] = count
    ctx.user_data.pop("pending_q", None)
    labels = ["A","B","C","D"]
    sec_name = ctx.user_data.get("current_section_name", "")
    sec_info = f"\n📂 {sec_name}" if sec_name else ""
    await query.message.reply_text(
        f"✅ *Q{count} saved!* ⏱ {timer}s{sec_info}\n\nSend more or /done",
        parse_mode=ParseMode.MARKDOWN
    )
    return CREATE_COLLECTING


async def create_cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    count = ctx.user_data.get("create_count", 0)
    name  = ctx.user_data.get("create_set_name","")
    ctx.user_data.clear()
    if count > 0:
        await update.message.reply_text(
            f"❌ Cancelled। {count} questions already save hain set '{name}' mein।"
        )
    else:
        await update.message.reply_text("❌ Cancelled।")
    return ConversationHandler.END


# ── Manual /newquiz (one by one) ──────────────────────────────────────────────

async def newquiz_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        return ConversationHandler.END
    ctx.user_data.clear()
    await update.message.reply_text(
        "📝 *नया सवाल बनाएं*\n\nसवाल टाइप करें\n"
        "_(Photo के साथ — photo भेजें, caption में सवाल)_\n\n/cancel — रद्द करें",
        parse_mode=ParseMode.MARKDOWN
    )
    return MANUAL_QUESTION


async def recv_question(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    if msg.photo:
        ctx.user_data["photo_id"] = msg.photo[-1].file_id
        ctx.user_data["question"] = msg.caption or ""
    else:
        ctx.user_data["question"] = msg.text.strip()
    await msg.reply_text("✅ सवाल मिला!\n\n*Option A* टाइप करें:", parse_mode=ParseMode.MARKDOWN)
    return MANUAL_OPTION_A

async def recv_option_a(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["options"] = [update.message.text.strip()]
    await update.message.reply_text("*Option B* टाइप करें:", parse_mode=ParseMode.MARKDOWN)
    return MANUAL_OPTION_B

async def recv_option_b(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["options"].append(update.message.text.strip())
    await update.message.reply_text("*Option C* टाइप करें:", parse_mode=ParseMode.MARKDOWN)
    return MANUAL_OPTION_C

async def recv_option_c(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["options"].append(update.message.text.strip())
    await update.message.reply_text("*Option D* टाइप करें:", parse_mode=ParseMode.MARKDOWN)
    return MANUAL_OPTION_D

async def recv_option_d(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["options"].append(update.message.text.strip())
    await update.message.reply_text(
        "✅ चारों options मिले!\n\nसही जवाब चुनें:",
        reply_markup=option_kb(ctx.user_data["options"]),
        parse_mode=ParseMode.MARKDOWN
    )
    return MANUAL_CORRECT

async def recv_correct(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    ctx.user_data["correct"] = int(query.data.split("_")[1])
    await query.message.reply_text(
        "📖 Explanation लिखें:\n_(नहीं चाहिए तो /skip करें)_",
        parse_mode=ParseMode.MARKDOWN
    )
    return MANUAL_EXPLANATION

async def recv_explanation(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text.strip()
    ctx.user_data["explanation"] = "" if txt == "/skip" else txt
    await update.message.reply_text("⏱ Timer चुनें:", reply_markup=timer_kb())
    return MANUAL_TIMER

async def recv_timer(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    ctx.user_data["timer"] = int(query.data.split("_")[1])
    sets = db.get_all_sets()
    if sets:
        btns = [[InlineKeyboardButton(s["name"], callback_data=f"addtoset_{s['id']}")] for s in sets]
        btns.append([InlineKeyboardButton("➕ नया Set", callback_data="newset")])
        await query.message.reply_text("किस Set में जोड़ें?", reply_markup=InlineKeyboardMarkup(btns))
    else:
        await query.message.reply_text("नए Set का नाम टाइप करें:")
    return SET_NAME

async def recv_set_choice(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "newset":
        await query.message.reply_text("नए Set का नाम टाइप करें:")
        return SET_NAME
    set_id = int(query.data.split("_")[1])
    return await _save_question(query.message, ctx, set_id)

async def recv_set_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    name   = update.message.text.strip()
    set_id = db.create_set(name)
    return await _save_question(update.message, ctx, set_id)

async def _save_question(msg, ctx, set_id: int):
    d = ctx.user_data
    db.add_question(
        set_id=set_id, question=d.get("question",""),
        options=d.get("options",[]), correct=d.get("correct",0),
        explanation=d.get("explanation",""), timer=d.get("timer",20),
        photo_id=d.get("photo_id"),
    )
    await msg.reply_text("✅ *सवाल save हो गया!*", parse_mode=ParseMode.MARKDOWN)
    ctx.user_data.clear()
    return ConversationHandler.END

async def cancel_conv(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("❌ रद्द कर दिया।")
    ctx.user_data.clear()
    return ConversationHandler.END


# ── Forwarded Poll Handler (Feature 1) ────────────────────────────────────────

async def handle_forwarded_poll(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """
    Forwarded quiz poll ko save karta hai.
    Bug fix: filters.Poll removed, use filters.POLL or filters.ALL
    Also: msg.poll None-safe check added
    """
    if not is_admin(update.effective_user.id):
        return
    msg = update.message
    if not msg or not msg.poll:
        return
    poll = msg.poll
    # Sirf Quiz type polls accept karo jisme correct_option_id ho
    if poll.type != Poll.QUIZ or poll.correct_option_id is None:
        await msg.reply_text(
            "⚠️ Yeh Quiz poll nahi hai ya correct answer set nahi hai.\n"
            "Sirf Quiz type poll forward karein (jisme ✅ correct option ho)।"
        )
        return
    options  = [o.text for o in poll.options]
    correct  = poll.correct_option_id
    question = poll.question
    expl     = poll.explanation or ""
    sets     = db.get_all_sets()
    set_id   = sets[0]["id"] if sets else db.create_set("Forwarded Polls")
    db.add_question(set_id=set_id, question=question, options=options,
                    correct=correct, explanation=expl, timer=20)
    labels   = ["A","B","C","D"]
    set_info = db.get_set(set_id)
    set_name = set_info["name"] if set_info else "Unknown"
    await msg.reply_text(
        f"✅ *Forwarded Poll save हो गया!*\n\n"
        f"❓ {question}\n"
        f"✔️ सही: *{labels[correct]}*. {options[correct]}\n"
        f"📂 Set: {set_name}",
        parse_mode=ParseMode.MARKDOWN
    )


# ── Bulk Excel Upload (Feature 5) ─────────────────────────────────────────────

async def bulk_upload_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        return
    await update.message.reply_text(
        "📊 *Excel Bulk Upload*\n\n"
        "Format: `Question|A|B|C|D|Correct(0-3)|Explanation|Timer`\n\n"
        "अब .xlsx file भेजें:",
        parse_mode=ParseMode.MARKDOWN
    )

async def handle_excel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        return
    doc = update.message.document
    if not doc or not doc.file_name.endswith(".xlsx"):
        await update.message.reply_text("❌ .xlsx file भेजें।")
        return
    await update.message.reply_text("⏳ Process हो रही है...")
    file = await ctx.bot.get_file(doc.file_id)
    buf  = io.BytesIO()
    await file.download_to_memory(buf)
    buf.seek(0)
    wb       = openpyxl.load_workbook(buf)
    ws       = wb.active
    set_name = doc.file_name.replace(".xlsx","")
    set_id   = db.create_set(set_name)
    count, errors = 0, 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        try:
            vals = list(row) + [None]*8
            q,a,b,c,d,correct,expl,t = vals[:8]
            if not q:
                continue
            db.add_question(set_id=set_id, question=str(q),
                            options=[str(a),str(b),str(c),str(d)],
                            correct=int(correct), explanation=str(expl or ""),
                            timer=int(t or 20))
            count += 1
        except Exception as e:
            logger.warning(f"Row error: {e}")
            errors += 1
    await update.message.reply_text(
        f"✅ *Upload पूरा!*\n📂 {set_name}\n✔️ {count} सवाल | ❌ {errors} errors",
        parse_mode=ParseMode.MARKDOWN
    )


# ── Quiz Engine ────────────────────────────────────────────────────────────────

async def list_sets(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        return
    sets = db.get_all_sets()
    if not sets:
        await update.message.reply_text("कोई Set नहीं। /create से बनाएं।")
        return
    btns = []
    for s in sets:
        neg = s.get("neg_marks", 0)
        pos = s.get("pos_marks", 1)
        neg_str = f"-{neg}" if neg > 0 else "0"
        label = f"▶️ {s['name']} ({s['count']}Q | +{pos}/{neg_str})"
        btns.append([InlineKeyboardButton(label, callback_data=f"startset_{s['id']}")])
    await update.message.reply_text(
        "📚 *सभी Quiz Sets:*", reply_markup=InlineKeyboardMarkup(btns),
        parse_mode=ParseMode.MARKDOWN
    )

async def startquiz_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await list_sets(update, ctx)

async def start_quiz_callback(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if not is_admin(query.from_user.id):
        return
    set_id    = int(query.data.split("_")[1])
    chat_id   = query.message.chat_id
    set_info  = db.get_set(set_id)
    if not set_info:
        await query.message.reply_text("❌ Set नहीं मिला।")
        return

    pos_marks = set_info.get("pos_marks", 1.0)
    neg_marks = set_info.get("neg_marks", 0.0)
    has_sec   = set_info.get("has_sections", 0)

    if has_sec:
        sections  = db.get_sections(set_id)
        if not sections:
            await query.message.reply_text("❌ कोई section नहीं मिला।")
            return
        # Build sectional quiz
        all_questions = []
        section_breaks = []  # list of (section_name, section_timer, start_idx, end_idx)
        for sec in sections:
            qs = db.get_questions(set_id, section_id=sec["id"])
            if qs:
                start = len(all_questions)
                all_questions.extend(qs)
                end = len(all_questions)
                section_breaks.append({
                    "name": sec["name"],
                    "timer": sec["timer"],
                    "start": start,
                    "end": end
                })
    else:
        all_questions = db.get_questions(set_id)
        section_breaks = None

    if not all_questions:
        await query.message.reply_text("❌ Set में कोई सवाल नहीं।")
        return

    now_str = datetime.now().strftime("%d %b %Y, %I:%M %p IST")
    neg_str = f"❌ -{neg_marks}" if neg_marks > 0 else "No Negative"
    quiz = {
        "questions": all_questions,
        "scores": {}, "active": True, "finished": False,
        "poll_map": {}, "start_times": {}, "student_answers": {},
        "set_name": set_info["name"],
        "quiz_date": now_str,
        "total_q": len(all_questions),
        "chat_id": chat_id,
        "pos_marks": pos_marks,
        "neg_marks": neg_marks,
        "section_breaks": section_breaks,
    }
    ctx.chat_data["quiz"] = quiz

    sec_info = f"\n📂 {len(section_breaks)} Sections" if section_breaks else ""
    await query.message.reply_text(
        f"🚀 *Quiz शुरू!*\n"
        f"📚 {set_info['name']}\n"
        f"❓ {len(all_questions)} सवाल{sec_info}\n"
        f"📊 +{pos_marks} / {neg_str}\n\n"
        f"_सभी students पहले bot को /start करें!_",
        parse_mode=ParseMode.MARKDOWN
    )
    asyncio.create_task(run_quiz(ctx.bot, chat_id, quiz))


async def run_quiz(bot, chat_id: int, quiz: dict):
    questions     = quiz["questions"]
    section_breaks = quiz.get("section_breaks")

    if section_breaks:
        # Sectional mode — Feature 6
        for sec in section_breaks:
            if not quiz.get("active"):
                break
            sec_name  = sec["name"]
            sec_timer = sec["timer"]
            sec_qs    = questions[sec["start"]:sec["end"]]
            m, s = divmod(sec_timer, 60)
            await bot.send_message(
                chat_id,
                f"📂 *Section: {sec_name}*\n"
                f"❓ {len(sec_qs)} Questions | ⏱ {m}m {s}s total\n\n"
                f"_Section शुरू हो रहा है..._",
                parse_mode=ParseMode.MARKDOWN, protect_content=True
            )
            await asyncio.sleep(3)
            sec_start_time = time.time()
            for local_idx, q in enumerate(sec_qs):
                if not quiz.get("active"):
                    break
                global_idx = sec["start"] + local_idx
                elapsed = time.time() - sec_start_time
                remaining = sec_timer - elapsed
                if remaining <= 0:
                    await bot.send_message(
                        chat_id, f"⏰ *Section '{sec_name}' का time खत्म!*",
                        parse_mode=ParseMode.MARKDOWN
                    )
                    break
                q_timer = min(q.get("timer", 20), int(remaining))
                if q_timer <= 0:
                    break
                await _send_question(bot, chat_id, quiz, global_idx, q, q_timer)
                try:
                    await asyncio.sleep(q_timer + 3)
                except asyncio.CancelledError:
                    break
            await bot.send_message(
                chat_id, f"✅ *Section '{sec_name}' खत्म!*",
                parse_mode=ParseMode.MARKDOWN
            )
            await asyncio.sleep(3)
    else:
        # Normal mode — Feature 8 (no limit, per-Q timer)
        for idx, q in enumerate(questions):
            if not quiz.get("active"):
                break
            timer = q.get("timer", 20)
            await _send_question(bot, chat_id, quiz, idx, q, timer)
            try:
                await asyncio.sleep(timer + 3)
            except asyncio.CancelledError:
                break

    if quiz.get("active") and not quiz.get("finished"):
        await finish_quiz(bot, chat_id, quiz)


async def _send_question(bot, chat_id, quiz, idx, q, timer):
    if q.get("photo_id"):
        try:
            await bot.send_photo(
                chat_id=chat_id, photo=q["photo_id"],
                caption=f"❓ *Q{idx+1}:* {q['question']}",
                parse_mode=ParseMode.MARKDOWN,
                protect_content=True  # Feature 7
            )
        except TelegramError as e:
            logger.warning(f"Photo failed Q{idx+1}: {e}")
    try:
        sent = await bot.send_poll(
            chat_id=chat_id,
            question=f"Q{idx+1}: {q['question'][:255]}",
            options=q["options"],
            type=Poll.QUIZ,
            correct_option_id=q["correct"],
            explanation=(q.get("explanation","") or "")[:200] or None,
            open_period=max(5, min(timer, 600)),
            is_anonymous=False,
            protect_content=True,  # Feature 7 — Forward disable
        )
        poll_id = sent.poll.id
        quiz["poll_map"][poll_id]    = idx
        quiz["start_times"][poll_id] = time.time()
        POLL_TO_CHAT[poll_id]        = chat_id
    except TelegramError as e:
        logger.error(f"Poll failed Q{idx+1}: {e}")


async def handle_poll_answer(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    answer  = update.poll_answer
    poll_id = answer.poll_id
    chat_id = POLL_TO_CHAT.get(poll_id)
    if not chat_id:
        return
    quiz = ctx.application.chat_data.get(chat_id, {}).get("quiz")
    if not quiz or poll_id not in quiz.get("poll_map", {}):
        return
    uid    = answer.user.id
    name   = answer.user.full_name
    idx    = quiz["poll_map"][poll_id]
    q      = quiz["questions"][idx]
    taken  = round(time.time() - quiz["start_times"].get(poll_id, time.time()), 1)
    chosen = answer.option_ids[0] if answer.option_ids else -1
    correct_idx = q["correct"]

    pos_marks = quiz.get("pos_marks", 1.0)
    neg_marks = quiz.get("neg_marks", 0.0)

    if uid not in quiz["scores"]:
        quiz["scores"][uid] = {
            "name": name, "score": 0.0,
            "correct": 0, "wrong": 0, "skipped": 0,
            "time": 0.0, "answered": 0
        }
    e = quiz["scores"][uid]

    if chosen == correct_idx:
        e["score"]   += pos_marks
        e["correct"] += 1
    elif chosen == -1:
        e["skipped"] += 1
    else:
        e["score"]   -= neg_marks   # Feature 2 — Negative Marking (can go to 0 or negative)
        e["wrong"]   += 1

    e["time"]     += taken
    e["answered"] += 1

    if uid not in quiz["student_answers"]:
        quiz["student_answers"][uid] = {}
    quiz["student_answers"][uid][idx] = chosen
    db.record_answer(uid, name, poll_id, chosen, correct_idx, taken)


async def finish_quiz(bot, chat_id: int, quiz: dict):
    if quiz.get("finished"):
        return
    quiz["finished"] = True
    quiz["active"]   = False
    scores = quiz["scores"]
    if not scores:
        await bot.send_message(chat_id, "⚠️ Quiz खत्म — कोई जवाब नहीं मिला।")
        return

    total_q       = quiz.get("total_q", len(quiz["questions"]))
    pos_marks     = quiz.get("pos_marks", 1.0)
    neg_marks     = quiz.get("neg_marks", 0.0)
    max_score     = total_q * pos_marks
    # Sort by score desc, then time asc — Feature 2 (score can be 0 or negative)
    sorted_scores = sorted(scores.items(), key=lambda x: (-x[1]["score"], x[1]["time"]))
    total_students= len(sorted_scores)
    medals = ["🥇","🥈","🥉"]

    neg_str = f"❌ -{neg_marks}" if neg_marks > 0 else "No Negative"
    text   = (f"🏆 *Final Leaderboard*\n"
              f"📊 +{pos_marks} / {neg_str} | Max: {fmt_score(max_score)}\n"
              + "─"*30 + "\n")

    for rank, (uid, s) in enumerate(sorted_scores, 1):
        medal = medals[rank-1] if rank <= 3 else f"#{rank}"
        acc   = calc_acc(s["correct"], s["answered"])
        score_disp = fmt_score(s["score"])
        text += (f"{medal} *{s['name']}*\n"
                 f"   💯 {score_disp}/{fmt_score(max_score)} | ✅ {s['correct']} | "
                 f"❌ {s['wrong']} | 🎯 {acc}% | ⏱ {fmt_time(s['time'])}\n\n")

    # Feature 3 — Automatic leaderboard in group (protect_content so can't forward)
    await bot.send_message(
        chat_id, text, parse_mode=ParseMode.MARKDOWN, protect_content=True
    )

    questions  = quiz["questions"]
    now_str    = quiz.get("quiz_date", datetime.now().strftime("%d %b %Y, %I:%M %p IST"))
    set_name   = quiz.get("set_name","Quiz")
    lb_for_pdf = []
    for rank, (uid, s) in enumerate(sorted_scores, 1):
        acc = calc_acc(s["correct"], s["answered"])
        lb_for_pdf.append({
            "rank": rank, "name": s["name"], "score": s["score"],
            "correct": s["correct"], "wrong": s["wrong"],
            "acc": acc, "time": fmt_time(s["time"])
        })

    scoring_str = f"+{pos_marks} / -{neg_marks}" if neg_marks > 0 else f"+{pos_marks} / 0"
    await bot.send_message(chat_id, "📄 *PDF भेजी जा रही है...*", parse_mode=ParseMode.MARKDOWN)

    # Feature 3 — PDF automatic send
    sent, failed = 0, []
    for rank, (uid, s) in enumerate(sorted_scores, 1):
        try:
            acc     = calc_acc(s["correct"], s["answered"])
            std_ans = quiz["student_answers"].get(uid, {})
            pdf_buf = generate_result_pdf(
                quiz_title=set_name, quiz_day=BOT_USER, quiz_date=now_str,
                total_questions=total_q, scoring=scoring_str,
                leaderboard=lb_for_pdf, questions=questions,
                student_answers=std_ans, student_name=s["name"],
                pos_marks=pos_marks, neg_marks=neg_marks,
            )
            score_disp = fmt_score(s["score"])
            await bot.send_document(
                chat_id=uid, document=pdf_buf,
                filename=f"Result_{s['name'].replace(' ','_')}.pdf",
                caption=(f"🎯 *आपका Result*\n\n🏆 Rank: #{rank}/{total_students}\n"
                         f"💯 Score: {score_disp}/{fmt_score(max_score)}\n"
                         f"✅ Correct: {s['correct']} | ❌ Wrong: {s['wrong']}\n"
                         f"🎯 Accuracy: {acc}%\n⏱ Time: {fmt_time(s['time'])}"),
                parse_mode=ParseMode.MARKDOWN,
                protect_content=True,
            )
            sent += 1
            await asyncio.sleep(0.05)
        except TelegramError as e:
            logger.warning(f"PDF failed {s['name']}: {e}")
            failed.append(s["name"])

    msg = f"✅ *{sent}/{total_students} students को PDF मिली!*"
    if failed:
        msg += f"\n\n⚠️ *इन्हें नहीं मिली* (bot को /start करें):\n"
        msg += "\n".join(f"• {n}" for n in failed[:15])
    await bot.send_message(chat_id, msg, parse_mode=ParseMode.MARKDOWN)

    db.save_leaderboard(chat_id, sorted_scores)
    db.cleanup_old_answers()
    for pid in list(POLL_TO_CHAT.keys()):
        if POLL_TO_CHAT[pid] == chat_id:
            del POLL_TO_CHAT[pid]


async def stop_quiz(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        return
    quiz = ctx.chat_data.get("quiz")
    if quiz and quiz.get("active") and not quiz.get("finished"):
        await finish_quiz(ctx.bot, update.effective_chat.id, quiz)
    else:
        await update.message.reply_text("कोई Quiz नहीं चल रही।")


async def leaderboard_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    rows = db.get_leaderboard(update.effective_chat.id)
    if not rows:
        await update.message.reply_text("अभी कोई score नहीं है।")
        return
    medals = ["🥇","🥈","🥉"]
    text   = "🏆 *Overall Leaderboard*\n\n"
    for i, r in enumerate(rows[:20], 1):
        medal = medals[i-1] if i <= 3 else f"#{i}"
        text += f"{medal} {r['name']} — {fmt_score(r['score'])} pts\n"
    await update.message.reply_text(text, parse_mode=ParseMode.MARKDOWN, protect_content=True)


async def reset_scores(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        return
    db.reset_leaderboard(update.effective_chat.id)
    await update.message.reply_text("✅ Scores reset हो गए।")


# ── App Build ──────────────────────────────────────────────────────────────────

def build_app():
    app = Application.builder().token(BOT_TOKEN).build()

    # /create conversation — with marking + sections + per-Q timer
    create_conv = ConversationHandler(
        entry_points=[CommandHandler("create", create_start)],
        states={
            CREATE_WAITING_NAME: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, create_recv_name)
            ],
            MARKING_POS: [
                CallbackQueryHandler(create_marking_choice, pattern=r"^mark_"),
                MessageHandler(filters.TEXT & ~filters.COMMAND, create_recv_pos_marks),
            ],
            MARKING_NEG: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, create_recv_neg_marks),
            ],
            SECTION_CONFIRM: [
                CallbackQueryHandler(create_section_confirm, pattern=r"^use_sections_"),
            ],
            SECTION_NAME: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, create_recv_section_name),
            ],
            SECTION_TIMER: [
                CallbackQueryHandler(create_recv_section_timer, pattern=r"^sectimer_"),
            ],
            CREATE_COLLECTING: [
                CommandHandler("newsection", create_new_section),
                MessageHandler(
                    filters.TEXT | filters.PHOTO | filters.ALL | filters.Document.ALL,
                    create_recv_question
                ),
            ],
            CREATE_Q_TIMER: [
                CallbackQueryHandler(create_recv_q_timer, pattern=r"^timer_"),
            ],
        },
        fallbacks=[CommandHandler("cancel", create_cancel)],
        per_chat=False,
        allow_reentry=True,
    )

    # /newquiz conversation
    newquiz_conv = ConversationHandler(
        entry_points=[CommandHandler("newquiz", newquiz_start)],
        states={
            MANUAL_QUESTION   : [MessageHandler(filters.TEXT | filters.PHOTO, recv_question)],
            MANUAL_OPTION_A   : [MessageHandler(filters.TEXT, recv_option_a)],
            MANUAL_OPTION_B   : [MessageHandler(filters.TEXT, recv_option_b)],
            MANUAL_OPTION_C   : [MessageHandler(filters.TEXT, recv_option_c)],
            MANUAL_OPTION_D   : [MessageHandler(filters.TEXT, recv_option_d)],
            MANUAL_CORRECT    : [CallbackQueryHandler(recv_correct, pattern=r"^correct_")],
            MANUAL_EXPLANATION: [MessageHandler(filters.TEXT, recv_explanation)],
            MANUAL_TIMER      : [CallbackQueryHandler(recv_timer, pattern=r"^timer_")],
            SET_NAME          : [
                MessageHandler(filters.TEXT, recv_set_name),
                CallbackQueryHandler(recv_set_choice, pattern=r"^(addtoset_|newset)"),
            ],
        },
        fallbacks=[CommandHandler("cancel", cancel_conv)],
        per_chat=False,
    )

    app.add_handler(CommandHandler("start",       start))
    app.add_handler(CommandHandler("help",        help_cmd))
    app.add_handler(CommandHandler("sets",        list_sets))
    app.add_handler(CommandHandler("startquiz",   startquiz_cmd))
    app.add_handler(CommandHandler("stopquiz",    stop_quiz))
    app.add_handler(CommandHandler("leaderboard", leaderboard_cmd))
    app.add_handler(CommandHandler("resetscores", reset_scores))
    app.add_handler(CommandHandler("bulkupload",  bulk_upload_start))
    app.add_handler(create_conv)
    app.add_handler(newquiz_conv)
    app.add_handler(MessageHandler(filters.FORWARDED & filters.POLL, handle_forwarded_poll))
    app.add_handler(MessageHandler(filters.Document.FileExtension("xlsx"), handle_excel))
    app.add_handler(PollAnswerHandler(handle_poll_answer))
    app.add_handler(CallbackQueryHandler(start_quiz_callback, pattern=r"^startset_"))

    return app


if __name__ == "__main__":
    if not BOT_TOKEN or BOT_TOKEN == "YOUR_NEW_TOKEN_HERE":
        print("[ERROR] config.py mein BOT_TOKEN set nahi hai!")
        exit(1)
    if not ADMIN_IDS or ADMIN_IDS == [123456789]:
        print("[WARNING] config.py mein apna ADMIN_IDS set karo!")
    app = build_app()
    logger.info(f"{BOT_NAME} starting up...")
    app.run_polling(allowed_updates=Update.ALL_TYPES, drop_pending_updates=True)
