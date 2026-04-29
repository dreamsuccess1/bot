"""
Quiz Bot — All Bugs Fixed
Fixed:
1. BOT_TOKEN — direct string, no os.getenv, no comma
2. Forwarded filter — sirf Poll.QUIZ wale messages
3. answer.option_ids safe access
4. POLL_TO_CHAT cleanup on restart
5. DB cleanup old answers
6. asyncio.sleep exception handling
7. Admin check reliable
"""

import asyncio
import logging
import os
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

logging.basicConfig(
    format="%(asctime)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# ── Conversation States ──────────────────────────────────────────────────────
(
    MANUAL_QUESTION, MANUAL_OPTION_A, MANUAL_OPTION_B,
    MANUAL_OPTION_C, MANUAL_OPTION_D, MANUAL_CORRECT,
    MANUAL_EXPLANATION, MANUAL_TIMER, SET_NAME,
) = range(9)

# BUG #9 FIX — restart pe clean slate
POLL_TO_CHAT: dict = {}

# ── Helpers ──────────────────────────────────────────────────────────────────

def is_admin(uid: int) -> bool:
    # BUG #6 FIX — reliable check
    return int(uid) in [int(a) for a in ADMIN_IDS]

def timer_kb():
    return InlineKeyboardMarkup([[
        InlineKeyboardButton(f"⏱ {t}s", callback_data=f"timer_{t}")
        for t in TIMERS
    ]])

def option_kb(options: list, prefix="correct"):
    labels = ["A","B","C","D"]
    return InlineKeyboardMarkup([[
        InlineKeyboardButton(
            f"{labels[i]}: {str(o)[:20]}",
            callback_data=f"{prefix}_{i}"
        )
    ] for i, o in enumerate(options)])

def fmt_time(seconds: float) -> str:
    m, s = divmod(int(seconds), 60)
    return f"{m}m {s}s"

def calc_acc(correct: int, total: int) -> int:
    return round((correct / total) * 100) if total > 0 else 0

# ── /start ───────────────────────────────────────────────────────────────────

async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    db.register_user(user.id, user.full_name, user.username)
    text = (
        f"🎯 *{BOT_NAME} में आपका स्वागत है!*\n\n"
        "Quiz में participate करने के लिए तैयार रहें।\n"
        "_(PDF result पाने के लिए bot को /start करना ज़रूरी है)_"
    )
    if is_admin(user.id):
        text += (
            "\n\n🔧 *Admin Commands:*\n"
            "/newquiz — नया सवाल बनाएं\n"
            "/bulkupload — Excel से upload\n"
            "/sets — सभी sets\n"
            "/startquiz — Quiz शुरू करें\n"
            "/stopquiz — Quiz रोकें\n"
            "/leaderboard — Rankings\n"
            "/resetscores — Reset scores"
        )
    await update.message.reply_text(text, parse_mode=ParseMode.MARKDOWN)

async def help_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await start(update, ctx)

# ── Manual Question Creation ─────────────────────────────────────────────────

async def newquiz_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        return ConversationHandler.END
    ctx.user_data.clear()
    await update.message.reply_text(
        "📝 *नया सवाल बनाएं*\n\n"
        "सवाल टाइप करें\n"
        "_(Photo के साथ — photo भेजें, caption में सवाल)_\n\n"
        "/cancel — रद्द करें",
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
        set_id=set_id,
        question=d.get("question",""),
        options=d.get("options",[]),
        correct=d.get("correct",0),
        explanation=d.get("explanation",""),
        timer=d.get("timer",20),
        photo_id=d.get("photo_id"),
    )
    await msg.reply_text("✅ *सवाल save हो गया!*", parse_mode=ParseMode.MARKDOWN)
    ctx.user_data.clear()
    return ConversationHandler.END

async def cancel_conv(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("❌ रद्द कर दिया।")
    ctx.user_data.clear()
    return ConversationHandler.END

# ── Poll Forwarding ──────────────────────────────────────────────────────────
# BUG #4 FIX — sirf forwarded poll messages handle karo

async def handle_forwarded_poll(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        return

    msg  = update.message
    poll = msg.poll

    # Agar poll nahi hai to silently return
    if not poll:
        return

    # Sirf QUIZ type polls
    if poll.type != Poll.QUIZ:
        await msg.reply_text("⚠️ Sirf Quiz polls forward karein.")
        return

    # Correct answer check
    if poll.correct_option_id is None:
        await msg.reply_text("⚠️ Is poll mein sahi jawab nahi hai.")
        return

    options  = [o.text for o in poll.options]
    correct  = poll.correct_option_id
    question = poll.question
    expl     = poll.explanation or ""

    sets   = db.get_all_sets()
    set_id = sets[0]["id"] if sets else db.create_set("Forwarded Polls")

    db.add_question(
        set_id=set_id, question=question,
        options=options, correct=correct,
        explanation=expl, timer=20
    )

    labels = ["A","B","C","D"]
    await msg.reply_text(
        f"✅ *Poll save हो गया!*\n\n"
        f"❓ {question}\n"
        f"✔️ सही: {labels[correct]}. {options[correct]}\n"
        f"📂 Set: {db.get_set(set_id)['name']}",
        parse_mode=ParseMode.MARKDOWN
    )

# ── Bulk Excel Upload ─────────────────────────────────────────────────────────

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
            q,a,b,c,d,correct,expl,timer = vals[:8]
            if not q:
                continue
            db.add_question(
                set_id=set_id, question=str(q),
                options=[str(a),str(b),str(c),str(d)],
                correct=int(correct),
                explanation=str(expl or ""),
                timer=int(timer or 20)
            )
            count += 1
        except Exception as e:
            logger.warning(f"Row error: {e}")
            errors += 1

    await update.message.reply_text(
        f"✅ *Upload पूरा!*\n📂 {set_name}\n✔️ {count} सवाल | ❌ {errors} errors",
        parse_mode=ParseMode.MARKDOWN
    )

# ── Quiz Engine ───────────────────────────────────────────────────────────────

async def list_sets(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        return
    sets = db.get_all_sets()
    if not sets:
        await update.message.reply_text("कोई Set नहीं। /newquiz से बनाएं।")
        return
    btns = [[InlineKeyboardButton(
        f"▶️ {s['name']} ({s['count']} सवाल)",
        callback_data=f"startset_{s['id']}"
    )] for s in sets]
    await update.message.reply_text(
        "📚 *सभी Quiz Sets:*",
        reply_markup=InlineKeyboardMarkup(btns),
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
    questions = db.get_questions(set_id)

    if not questions:
        await query.message.reply_text("❌ Set में कोई सवाल नहीं।")
        return

    set_info = db.get_set(set_id)
    now_str  = datetime.now().strftime("%d %b %Y, %I:%M %p IST")

    quiz = {
        "questions"      : questions,
        "scores"         : {},
        "active"         : True,
        "finished"       : False,
        "poll_map"       : {},
        "start_times"    : {},
        "student_answers": {},
        "set_name"       : set_info["name"] if set_info else "Quiz",
        "quiz_date"      : now_str,
        "total_q"        : len(questions),
        "chat_id"        : chat_id,
    }

    ctx.chat_data["quiz"] = quiz

    await query.message.reply_text(
        f"🚀 *Quiz शुरू!*\n"
        f"📚 {set_info['name']}\n"
        f"❓ {len(questions)} सवाल\n\n"
        f"_सभी students पहले bot को /start करें!_",
        parse_mode=ParseMode.MARKDOWN
    )
    asyncio.create_task(run_quiz(ctx.bot, chat_id, quiz))

async def run_quiz(bot, chat_id: int, quiz: dict):
    for idx, q in enumerate(quiz["questions"]):
        if not quiz.get("active"):
            break

        timer = q.get("timer", 20)

        if q.get("photo_id"):
            try:
                await bot.send_photo(
                    chat_id=chat_id,
                    photo=q["photo_id"],
                    caption=f"❓ *Q{idx+1}:* {q['question']}",
                    parse_mode=ParseMode.MARKDOWN,
                    protect_content=True,
                )
            except TelegramError as e:
                logger.warning(f"Photo failed Q{idx+1}: {e}")

        try:
            sent = await bot.send_poll(
                chat_id           = chat_id,
                question          = f"Q{idx+1}: {q['question'][:255]}",
                options           = q["options"],
                type              = Poll.QUIZ,
                correct_option_id = q["correct"],
                explanation       = (q.get("explanation","") or "")[:200] or None,
                open_period       = timer,
                is_anonymous      = False,
                protect_content   = True,
            )
            poll_id = sent.poll.id
            quiz["poll_map"][poll_id]    = idx
            quiz["start_times"][poll_id] = time.time()
            POLL_TO_CHAT[poll_id]        = chat_id

        except TelegramError as e:
            logger.error(f"Poll failed Q{idx+1}: {e}")
            continue

        # BUG #10 FIX — exception safe sleep
        try:
            await asyncio.sleep(timer + 3)
        except asyncio.CancelledError:
            break

    if quiz.get("active") and not quiz.get("finished"):
        await finish_quiz(bot, chat_id, quiz)

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

    # BUG #5 FIX — safe access
    chosen = answer.option_ids[0] if answer.option_ids else -1
    correct= q["correct"]

    if uid not in quiz["scores"]:
        quiz["scores"][uid] = {
            "name":name, "score":0,
            "correct":0, "wrong":0,
            "time":0.0, "answered":0
        }

    e = quiz["scores"][uid]
    if chosen == correct:
        e["score"]   += 1
        e["correct"] += 1
    else:
        e["wrong"] += 1
    e["time"]     += taken
    e["answered"] += 1

    if uid not in quiz["student_answers"]:
        quiz["student_answers"][uid] = {}
    quiz["student_answers"][uid][idx] = chosen

    db.record_answer(uid, name, poll_id, chosen, correct, taken)

async def finish_quiz(bot, chat_id: int, quiz: dict):
    if quiz.get("finished"):
        return
    quiz["finished"] = True
    quiz["active"]   = False

    scores = quiz["scores"]
    if not scores:
        await bot.send_message(chat_id, "⚠️ Quiz खत्म — कोई जवाब नहीं मिला।")
        return

    total_q = quiz.get("total_q", len(quiz["questions"]))
    sorted_scores = sorted(
        scores.items(),
        key=lambda x: (-x[1]["score"], x[1]["time"])
    )
    total_students = len(sorted_scores)

    medals = ["🥇","🥈","🥉"]
    text   = "🏆 *Final Leaderboard*\n" + "─"*30 + "\n"
    for rank, (uid, s) in enumerate(sorted_scores, 1):
        medal = medals[rank-1] if rank <= 3 else f"#{rank}"
        acc   = calc_acc(s["correct"], s["answered"])
        text += (
            f"{medal} *{s['name']}*\n"
            f"   💯 {s['score']}/{total_q} | ✅ {s['correct']} | "
            f"❌ {s['wrong']} | 🎯 {acc}% | ⏱ {fmt_time(s['time'])}\n\n"
        )

    await bot.send_message(
        chat_id, text,
        parse_mode=ParseMode.MARKDOWN,
        protect_content=True
    )

    questions = quiz["questions"]
    now_str   = quiz.get("quiz_date", datetime.now().strftime("%d %b %Y, %I:%M %p IST"))
    set_name  = quiz.get("set_name","Quiz")

    lb_for_pdf = []
    for rank, (uid, s) in enumerate(sorted_scores, 1):
        acc = calc_acc(s["correct"], s["answered"])
        lb_for_pdf.append({
            "rank":rank, "name":s["name"],
            "score":s["score"], "wrong":s["wrong"],
            "acc":acc, "time":fmt_time(s["time"]),
        })

    await bot.send_message(
        chat_id, "📄 *PDF भेजी जा रही है...*",
        parse_mode=ParseMode.MARKDOWN
    )

    sent, failed = 0, []
    for rank, (uid, s) in enumerate(sorted_scores, 1):
        try:
            acc     = calc_acc(s["correct"], s["answered"])
            std_ans = quiz["student_answers"].get(uid, {})

            pdf_buf = generate_result_pdf(
                quiz_title      = set_name,
                quiz_day        = BOT_USER,
                quiz_date       = now_str,
                total_questions = total_q,
                scoring         = "+1 / -0",
                leaderboard     = lb_for_pdf,
                questions       = questions,
                student_answers = std_ans,
                student_name    = s["name"],
            )

            await bot.send_document(
                chat_id        = uid,
                document       = pdf_buf,
                filename       = f"Result_{s['name'].replace(' ','_')}.pdf",
                caption        = (
                    f"🎯 *आपका Result*\n\n"
                    f"🏆 Rank: #{rank}/{total_students}\n"
                    f"💯 Score: {s['score']}/{total_q}\n"
                    f"✅ Correct: {s['correct']} | ❌ Wrong: {s['wrong']}\n"
                    f"🎯 Accuracy: {acc}%\n"
                    f"⏱ Time: {fmt_time(s['time'])}"
                ),
                parse_mode     = ParseMode.MARKDOWN,
                protect_content= True,
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

    # BUG #8 FIX — old answers cleanup
    db.cleanup_old_answers()

    # BUG #9 FIX — POLL_TO_CHAT cleanup
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
        text += f"{medal} {r['name']} — {r['score']} pts\n"
    await update.message.reply_text(
        text, parse_mode=ParseMode.MARKDOWN, protect_content=True
    )

async def reset_scores(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        return
    db.reset_leaderboard(update.effective_chat.id)
    await update.message.reply_text("✅ Scores reset हो गए।")

# ── App Build ─────────────────────────────────────────────────────────────────

def build_app():
    app = Application.builder().token(BOT_TOKEN).build()

    conv = ConversationHandler(
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
    app.add_handler(conv)

    # BUG #4 FIX — FORWARDED & Poll sirf poll messages
    app.add_handler(MessageHandler(
        filters.FORWARDED & filters.Poll(None), handle_forwarded_poll
    ))
    app.add_handler(MessageHandler(
        filters.Document.FileExtension("xlsx"), handle_excel
    ))
    app.add_handler(PollAnswerHandler(handle_poll_answer))
    app.add_handler(CallbackQueryHandler(start_quiz_callback, pattern=r"^startset_"))

    return app

if __name__ == "__main__":
    # BUG #1 FIX — Token verify karo start pe
    if not BOT_TOKEN or BOT_TOKEN == "YOUR_NEW_TOKEN_HERE":
        print("❌ ERROR: config.py mein BOT_TOKEN set nahi hai!")
        exit(1)

    if not ADMIN_IDS or ADMIN_IDS == [123456789]:
        print("⚠️ WARNING: config.py mein apna ADMIN_IDS set karo!")

    app = build_app()
    logger.info(f"🚀 {BOT_NAME} चालू हो रहा है...")
    app.run_polling(allowed_updates=Update.ALL_TYPES, drop_pending_updates=True)
