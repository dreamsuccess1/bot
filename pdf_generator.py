"""
pdf_generator.py — Dream Success style Quiz PDF
Config bot.py ki config.py se import hoti hai automatically
"""

import io
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table,
    TableStyle, HRFlowable
)

# Config ek jagah se — automatically match karega
try:
    from config import BOT_NAME, BOT_USER, TARGET_TXT
except ImportError:
    BOT_NAME   = "Quiz Bot"
    BOT_USER   = "@quizbot"
    TARGET_TXT = "Target Exams"

CLR_BLUE       = colors.HexColor("#1565C0")
CLR_BLUE_LIGHT = colors.HexColor("#1976D2")
CLR_GREEN_TXT  = colors.HexColor("#2E7D32")
CLR_GREY_ROW   = colors.HexColor("#F5F5F5")
CLR_WHITE      = colors.white
CLR_BLACK      = colors.HexColor("#212121")
CLR_BORDER     = colors.HexColor("#BDBDBD")
CLR_RED_TXT    = colors.HexColor("#C62828")
CLR_GREEN_BG   = colors.HexColor("#E8F5E9")
CLR_RED_BG     = colors.HexColor("#FFEBEE")
LABELS         = ["A","B","C","D"]

def _s(name, **kw):
    return ParagraphStyle(name, **kw)

def generate_result_pdf(
    quiz_title: str,
    quiz_day: str,
    quiz_date: str,
    total_questions: int,
    scoring: str,
    leaderboard: list,
    questions: list,
    student_answers: dict,
    student_name: str = None,
) -> io.BytesIO:

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
        rightMargin=1.2*cm, leftMargin=1.2*cm,
        topMargin=1.2*cm,  bottomMargin=1.2*cm)
    story = []

    # ── HEADER ──────────────────────────────────────────────────────────
    story.append(Paragraph(f"{BOT_NAME} — {quiz_title}",
        _s("T1", fontName="Helvetica-Bold", fontSize=17,
           textColor=CLR_BLUE, alignment=1, spaceAfter=2)))
    story.append(Paragraph(quiz_day,
        _s("T2", fontName="Helvetica-Bold", fontSize=13,
           textColor=CLR_BLUE, alignment=1, spaceAfter=3)))
    story.append(Paragraph(
        f"{TARGET_TXT} &nbsp;|&nbsp; {quiz_date} &nbsp;|&nbsp; "
        f"{total_questions} Questions &nbsp;|&nbsp; {scoring}",
        _s("M", fontName="Helvetica", fontSize=9,
           textColor=CLR_BLACK, alignment=1, spaceAfter=5)))
    story.append(HRFlowable(width="100%", thickness=2, color=CLR_BLUE, spaceAfter=8))

    # ── LEADERBOARD ─────────────────────────────────────────────────────
    def _blue_hdr(txt):
        t = Table([[Paragraph(f"  {txt}",
            _s(f"H{txt}", fontName="Helvetica-Bold", fontSize=11, textColor=CLR_WHITE))]],)
        t.setStyle(TableStyle([
            ("BACKGROUND",    (0,0),(-1,-1), CLR_BLUE_LIGHT),
            ("TOPPADDING",    (0,0),(-1,-1), 7),
            ("BOTTOMPADDING", (0,0),(-1,-1), 7),
        ]))
        return t

    story.append(_blue_hdr("LEADERBOARD"))
    story.append(Spacer(1,3))

    cs = _s("CS", fontName="Helvetica-Bold", fontSize=9, textColor=CLR_WHITE, alignment=1)
    ce = _s("CE", fontName="Helvetica",      fontSize=8, textColor=CLR_BLACK)
    cb = _s("CB", fontName="Helvetica-Bold", fontSize=9, textColor=CLR_BLUE)

    lb_data = [[
        Paragraph("Rank",cs), Paragraph("Participant",cs),
        Paragraph("Score",cs), Paragraph("Wrong",cs),
        Paragraph("Acc%",cs), Paragraph("Time",cs)
    ]]
    for row in leaderboard[:60]:
        lb_data.append([
            Paragraph(str(row.get("rank","")), ce),
            Paragraph(str(row.get("name",""))[:32], ce),
            Paragraph(str(row.get("score","")), cb),
            Paragraph(str(row.get("wrong","")), ce),
            Paragraph(f"{row.get('acc','')}%", ce),
            Paragraph(str(row.get("time","")), ce),
        ])

    lb = Table(lb_data, colWidths=[1.2*cm,7.5*cm,2*cm,1.5*cm,1.5*cm,2.3*cm], repeatRows=1)
    lb.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,0), CLR_BLUE_LIGHT),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[CLR_WHITE,CLR_GREY_ROW]),
        ("GRID",          (0,0),(-1,-1), 0.4, CLR_BORDER),
        ("TOPPADDING",    (0,0),(-1,-1), 4),
        ("BOTTOMPADDING", (0,0),(-1,-1), 4),
        ("LEFTPADDING",   (0,0),(-1,-1), 5),
        ("VALIGN",        (0,0),(-1,-1),"MIDDLE"),
    ]))
    story.append(lb)
    story.append(Spacer(1,12))

    # ── QUESTIONS & ANSWERS ─────────────────────────────────────────────
    story.append(_blue_hdr("QUESTIONS & ANSWERS"))
    story.append(Spacer(1,6))

    qn_s  = _s("QN",  fontName="Helvetica-Bold", fontSize=8.5, textColor=CLR_WHITE)
    qt_s  = _s("QT",  fontName="Helvetica-Bold", fontSize=8.5, textColor=CLR_BLACK, leading=12)
    oc_s  = _s("OC",  fontName="Helvetica-Bold", fontSize=8,   textColor=CLR_GREEN_TXT, leading=11)
    ow_s  = _s("OW",  fontName="Helvetica",      fontSize=8,   textColor=CLR_BLACK,    leading=11)
    oww_s = _s("OWW", fontName="Helvetica",      fontSize=8,   textColor=CLR_RED_TXT,  leading=11)

    def make_q(idx, q):
        chosen  = student_answers.get(idx, -1) if student_answers else -1
        correct = q.get("correct", 0)
        opts    = q.get("options", [])

        qh = Table([[
            Paragraph(f" Q{idx+1}", qn_s),
            Paragraph(str(q.get("question",""))[:200], qt_s),
        ]], colWidths=[1*cm,7.8*cm])
        qh.setStyle(TableStyle([
            ("BACKGROUND",    (0,0),(0,0), CLR_BLUE),
            ("BACKGROUND",    (1,0),(1,0), CLR_GREY_ROW),
            ("TOPPADDING",    (0,0),(-1,-1),5),
            ("BOTTOMPADDING", (0,0),(-1,-1),5),
            ("LEFTPADDING",   (0,0),(-1,-1),4),
            ("VALIGN",        (0,0),(-1,-1),"TOP"),
        ]))

        rows = [[qh]]
        for j, opt in enumerate(opts[:4]):  # max 4 options
            lbl = f"{LABELS[j]}) {opt}"
            if j == correct:
                st, bg = oc_s, CLR_GREEN_BG
            elif j == chosen:
                st, bg = oww_s, CLR_RED_BG
            else:
                st, bg = ow_s, CLR_WHITE

            or_ = Table([[Paragraph(f"  {lbl}", st)]], colWidths=[8.8*cm])
            or_.setStyle(TableStyle([
                ("BACKGROUND",    (0,0),(-1,-1), bg),
                ("TOPPADDING",    (0,0),(-1,-1), 3),
                ("BOTTOMPADDING", (0,0),(-1,-1), 3),
                ("LEFTPADDING",   (0,0),(-1,-1), 6),
                ("LINEBELOW",     (0,0),(-1,-1), 0.3, CLR_BORDER),
            ]))
            rows.append([or_])

        ct = Table(rows, colWidths=[8.8*cm])
        ct.setStyle(TableStyle([
            ("BOX",           (0,0),(-1,-1), 0.5, CLR_BORDER),
            ("TOPPADDING",    (0,0),(-1,-1), 0),
            ("BOTTOMPADDING", (0,0),(-1,-1), 0),
            ("LEFTPADDING",   (0,0),(-1,-1), 0),
            ("RIGHTPADDING",  (0,0),(-1,-1), 0),
        ]))
        return ct

    pairs = []
    for i in range(0, len(questions), 2):
        left  = make_q(i, questions[i])
        right = make_q(i+1, questions[i+1]) if i+1 < len(questions) else ""
        pairs.append([left, right])

    if pairs:
        grid = Table(pairs, colWidths=[9*cm,9*cm], hAlign="LEFT")
        grid.setStyle(TableStyle([
            ("VALIGN",        (0,0),(-1,-1),"TOP"),
            ("TOPPADDING",    (0,0),(-1,-1), 4),
            ("BOTTOMPADDING", (0,0),(-1,-1), 4),
            ("LEFTPADDING",   (0,0),(-1,-1), 2),
            ("RIGHTPADDING",  (0,0),(-1,-1), 2),
        ]))
        story.append(grid)

    # ── FOOTER ──────────────────────────────────────────────────────────
    story.append(Spacer(1,10))
    story.append(HRFlowable(width="100%", thickness=1, color=CLR_BORDER))
    now    = datetime.now().strftime("%d %b %Y, %I:%M %p")
    footer = f"Generated by {BOT_USER} • Pro Report Edition • {now}"
    if student_name:
        footer = f"Result for: {student_name} • " + footer
    story.append(Paragraph(footer,
        _s("F", fontName="Helvetica", fontSize=8,
           textColor=colors.grey, alignment=1, spaceBefore=4)))

    doc.build(story)
    buf.seek(0)
    return buf
