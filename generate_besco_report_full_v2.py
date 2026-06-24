#!/usr/bin/env python3
"""BESCO New Members Report Generator.

Usage:
    python3 generate_besco_report.py --start 2026-05-16 --end 2026-06-17 \
        --excel path/to/database.xlsx
"""

from __future__ import annotations

import argparse
import re
import os
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Iterable, Tuple

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle, PageBreak

# -----------------------------------------------------------------------------
# Font setup
# -----------------------------------------------------------------------------
# The original script hard-coded a Linux font path. That breaks on macOS/Windows.
# We try a small set of common locations and fall back gracefully.


def _first_existing(paths: Iterable[str]) -> str | None:
    for p in paths:
        if p and os.path.exists(p):
            return p
    return None


def register_fonts() -> Tuple[str, str, str]:
    """Register a Unicode font family and return (regular, bold, italic).

    We prefer fonts that support Cyrillic. If nothing is found, we fall back to
    the built-in Helvetica family, which will still allow the PDF to be created.
    """

    font_families = [
        (
            "DejaVu",
            [
                "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                "/Library/Fonts/DejaVu Sans.ttf",
                "/Library/Fonts/DejaVuSans.ttf",
                "/System/Library/Fonts/Supplemental/DejaVuSans.ttf",
                str(Path.home() / "Library/Fonts/DejaVu Sans.ttf"),
            ],
            [
                "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
                "/Library/Fonts/DejaVu Sans Bold.ttf",
                "/Library/Fonts/DejaVuSans-Bold.ttf",
                "/System/Library/Fonts/Supplemental/DejaVuSans-Bold.ttf",
                str(Path.home() / "Library/Fonts/DejaVu Sans Bold.ttf"),
            ],
            [
                "/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf",
                "/Library/Fonts/DejaVu Sans Oblique.ttf",
                "/Library/Fonts/DejaVuSans-Oblique.ttf",
                "/System/Library/Fonts/Supplemental/DejaVuSans-Oblique.ttf",
                str(Path.home() / "Library/Fonts/DejaVu Sans Oblique.ttf"),
            ],
        ),
        (
            "ArialUnicode",
            [
                "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
                "/System/Library/Fonts/Supplemental/Arial Unicode MS.ttf",
                "/Library/Fonts/Arial Unicode.ttf",
                "/Library/Fonts/Arial Unicode MS.ttf",
                r"C:\\Windows\\Fonts\\arialuni.ttf",
                r"C:\\Windows\\Fonts\\arial.ttf",
            ],
            [
                "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
                "/Library/Fonts/Arial Bold.ttf",
                r"C:\\Windows\\Fonts\\arialbd.ttf",
            ],
            [
                "/System/Library/Fonts/Supplemental/Arial Italic.ttf",
                "/Library/Fonts/Arial Italic.ttf",
                r"C:\\Windows\\Fonts\\ariali.ttf",
            ],
        ),
    ]

    for family_name, regular_paths, bold_paths, italic_paths in font_families:
        regular = _first_existing(regular_paths)
        if not regular:
            continue

        bold = _first_existing(bold_paths) or regular
        italic = _first_existing(italic_paths) or regular

        try:
            pdfmetrics.registerFont(TTFont(family_name, regular))
            pdfmetrics.registerFont(TTFont(f"{family_name}-Bold", bold))
            pdfmetrics.registerFont(TTFont(f"{family_name}-Italic", italic))
            return family_name, f"{family_name}-Bold", f"{family_name}-Italic"
        except Exception:
            # Try the next family if one of the font files is broken.
            continue

    return "Helvetica", "Helvetica-Bold", "Helvetica-Oblique"


FONT_NAME, FONT_BOLD, FONT_ITALIC = register_fonts()

# -----------------------------------------------------------------------------
# Colours
# -----------------------------------------------------------------------------
BLACK = colors.HexColor("#1A1A1A")
BESCO_DARK = colors.HexColor("#1A1A1A")
HEADER_BG = colors.HexColor("#1A1A1A")
ROW_ALT = colors.HexColor("#F5F5F5")
ACCENT = colors.HexColor("#2D2D2D")
WHITE = colors.white
LIGHT_GRAY = colors.HexColor("#E0E0E0")


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
MONTHS_BG = {
    1: "януари",
    2: "февруари",
    3: "март",
    4: "април",
    5: "май",
    6: "юни",
    7: "юли",
    8: "август",
    9: "септември",
    10: "октомври",
    11: "ноември",
    12: "декември",
}


def format_bg_date(dt: date) -> str:
    return f"{dt.day} {MONTHS_BG[dt.month]} {dt.year}"


def parse_date_col(series: pd.Series) -> pd.Series:
    """Try multiple date formats."""
    return pd.to_datetime(series, errors="coerce", dayfirst=True)


def normalize_url(url: object) -> str:
    """Return a usable URL string or an empty string."""
    if url is None:
        return ""
    s = str(url).strip()
    if not s or s.lower() == "nan":
        return ""
    if s.startswith("www."):
        s = "https://" + s
    if not re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", s):
        # If it looks like a domain, make it clickable.
        if "." in s and " " not in s:
            s = "https://" + s
    return s


def make_company_link(name: str, website: object, style: ParagraphStyle) -> Paragraph:
    """Create a company name paragraph that links to the website when present."""
    safe_name = str(name).strip() or ""
    safe_name = safe_name.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    url = normalize_url(website)
    if url:
        return Paragraph(f'<link href="{url}">{safe_name}</link>', style)
    return Paragraph(safe_name, style)


# -----------------------------------------------------------------------------
# Styles
# -----------------------------------------------------------------------------

def make_styles():
    return {
        "title": ParagraphStyle(
            "title",
            fontName=FONT_BOLD,
            fontSize=22,
            textColor=BESCO_DARK,
            spaceAfter=4,
            leading=26,
        ),
        "subtitle": ParagraphStyle(
            "subtitle",
            fontName=FONT_NAME,
            fontSize=10,
            textColor=colors.HexColor("#555555"),
            spaceAfter=2,
            leading=14,
        ),
        "address": ParagraphStyle(
            "address",
            fontName=FONT_NAME,
            fontSize=8.5,
            textColor=colors.HexColor("#555555"),
            spaceAfter=0,
            leading=13,
        ),
        "body": ParagraphStyle(
            "body",
            fontName=FONT_NAME,
            fontSize=9.5,
            textColor=BLACK,
            leading=15,
            spaceAfter=6,
        ),
        "body_bold": ParagraphStyle(
            "body_bold",
            fontName=FONT_BOLD,
            fontSize=9.5,
            textColor=BLACK,
            leading=15,
        ),
        "section": ParagraphStyle(
            "section",
            fontName=FONT_BOLD,
            fontSize=11,
            textColor=BESCO_DARK,
            spaceBefore=14,
            spaceAfter=8,
            leading=16,
        ),
        "footer": ParagraphStyle(
            "footer",
            fontName=FONT_NAME,
            fontSize=8,
            textColor=colors.HexColor("#888888"),
            leading=12,
        ),
        "cell": ParagraphStyle(
            "cell",
            fontName=FONT_NAME,
            fontSize=8.5,
            textColor=BLACK,
            leading=12,
        ),
        "cell_bold": ParagraphStyle(
            "cell_bold",
            fontName=FONT_BOLD,
            fontSize=8.5,
            textColor=BLACK,
            leading=12,
        ),
    }


# -----------------------------------------------------------------------------
# Data loading
# -----------------------------------------------------------------------------

def detect_website_column(columns) -> str | None:
    candidates = [
        "Website",
        "website",
        "Web Site",
        "Web site",
        "Company Website",
        "Company website",
        "URL",
        "Url",
        "Web",
        "Homepage",
        "Site",
    ]
    col_map = {str(c).strip().lower(): str(c) for c in columns}
    for candidate in candidates:
        key = candidate.strip().lower()
        if key in col_map:
            return col_map[key]
    return None


def load_members(excel_path, start_dt, end_dt):
    """Load and filter new members from Active and Fundraising sheets."""

    excel_path = str(excel_path)

    # Active sheet – "Joined 2026" new members
    df_active = pd.read_excel(excel_path, sheet_name="Active", header=0)
    df_active.columns = df_active.columns.astype(str).str.strip()

    joined_col = " "  # brand name column
    company_col = "Legal Company Name (Cyrillic)"
    type_col = "Type"
    date_col = "Date of payment in 2026"
    uic_col = "UIC"
    fee_col = "Fee"

    df_active[type_col] = df_active[type_col].astype(str).str.strip()
    df_new = df_active[df_active[type_col] == "Joined 2026"].copy()
    df_new[date_col] = parse_date_col(df_new[date_col])
    df_new = df_new[
        (df_new[date_col] >= pd.Timestamp(start_dt))
        & (df_new[date_col] <= pd.Timestamp(end_dt))
    ].copy()

    def tier(fee):
        try:
            f = float(fee)
            if f >= 800:
                return 'The Big "A" Play'
            elif f >= 360:
                return "Scaleup"
            return "Startup"
        except Exception:
            return "Startup"

    records = []
    for _, row in df_new.iterrows():
        name = str(row.get(joined_col, "")).strip() or str(row.get(company_col, "")).strip()
        legal = str(row.get(company_col, "")).strip()
        uic = str(row.get(uic_col, "")).strip()
        fee_v = row.get(fee_col, "")
        t = tier(fee_v)
        dt = row.get(date_col)
        date_str = dt.strftime("%d.%m.%Y") if not pd.isnull(dt) else ""
        website_col = detect_website_column(df_new.columns)
        website = row.get(website_col, "") if website_col else ""
        records.append(
            {
                "name": name,
                "legal": legal,
                "website": website,
                "date": date_str,
                "uic": uic,
                "tier": t,
                "source": "Active",
            }
        )

    # Fundraising sheet – "Fundraising New"
    df_fund = pd.read_excel(excel_path, sheet_name="Fundraising", header=0)
    df_fund.columns = df_fund.columns.astype(str).str.strip()
    df_fund["Type"] = df_fund["Type"].astype(str).str.strip()
    fund_new = df_fund[df_fund["Type"] == "Fundraising New"].copy()
    fund_new["Date of payment in 2026"] = parse_date_col(fund_new["Date of payment in 2026"])
    fund_new = fund_new[
        (fund_new["Date of payment in 2026"] >= pd.Timestamp(start_dt))
        & (fund_new["Date of payment in 2026"] <= pd.Timestamp(end_dt))
    ].copy()

    for _, row in fund_new.iterrows():
        name = str(row.get("Name", "")).strip()
        legal = str(row.get("Legal Company Name for Payment", "")).strip()
        uic = str(row.get("UIC/ ЕГН", "")).strip()
        dt = row.get("Date of payment in 2026")
        date_str = dt.strftime("%d.%m.%Y") if not pd.isnull(dt) else ""
        website_col = detect_website_column(df_fund.columns)
        website = row.get(website_col, "") if website_col else ""
        records.append(
            {
                "name": name,
                "legal": legal,
                "website": website,
                "date": date_str,
                "uic": uic,
                "tier": "Fundraising",
                "source": "Fundraising",
            }
        )

    def sort_key(r):
        try:
            return datetime.strptime(r["date"], "%d.%m.%Y")
        except Exception:
            return datetime.min

    records.sort(key=sort_key)
    return records


# -----------------------------------------------------------------------------
# PDF generation
# -----------------------------------------------------------------------------

def build_pdf(records, start_dt, end_dt, output_path):
    styles = make_styles()

    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=2.5 * cm,
        rightMargin=2.5 * cm,
        topMargin=2.2 * cm,
        bottomMargin=2.5 * cm,
    )

    W = A4[0] - 5 * cm
    story = []

    # Logo / header
    logo_text = (
        f'<font name="{FONT_BOLD}" size="28">B</font>'
        f'<font name="{FONT_BOLD}" size="20"><super>E</super></font>'
        f'<font name="{FONT_BOLD}" size="28">SCO</font>'
    )
    story.append(
        Paragraph(
            logo_text,
            ParagraphStyle(
                "logo",
                fontName=FONT_BOLD,
                fontSize=28,
                textColor=BESCO_DARK,
                spaceAfter=20,
                leading=34,
            ),
        )
    )

    addr_style = ParagraphStyle(
        "addr", fontName=FONT_NAME, fontSize=9, textColor=BLACK, leading=14, alignment=2, spaceAfter=2
    )
    story.append(Paragraph("До Управителния съвет", addr_style))
    story.append(Paragraph('на Сдружение "Българската предприемаческа асоциация" (BESCO)', addr_style))
    story.append(Spacer(1, 14))

    story.append(
        Paragraph(
            "Мотивирано предложение за прием на нови членове",
            ParagraphStyle(
                "title2",
                fontName=FONT_BOLD,
                fontSize=12,
                textColor=BLACK,
                alignment=1,
                spaceAfter=12,
                leading=16,
            ),
        )
    )

    intro = (
        f"За периода {format_bg_date(start_dt)} - {format_bg_date(end_dt)} следните кандидати изразиха желание за "
        "встъпване в членство в BESCO, съгласно условията и по реда на чл. 14 от Устава, а именно:"
    )
    story.append(
        Paragraph(
            intro,
            ParagraphStyle(
                "intro", fontName=FONT_NAME, fontSize=9.5, textColor=BLACK, leading=15, alignment=4, spaceAfter=14
            ),
        )
    )

    col_widths = [1.0 * cm, 6.2 * cm, 2.8 * cm, 2.8 * cm, 3.2 * cm, 2.5 * cm]

    header_style = ParagraphStyle(
        "th", fontName=FONT_BOLD, fontSize=8.5, textColor=WHITE, leading=12, alignment=1
    )

    headers = [
        Paragraph("#", header_style),
        Paragraph("Наименование на кандидата", header_style),
        Paragraph("Дата на подаване\nна заявление", header_style),
        Paragraph("ЕИК", header_style),
        Paragraph("Членски внос", header_style),
        Paragraph("Коментари", header_style),
    ]

    table_data = [headers]

    TIER_COLORS = {
        "Startup": colors.HexColor("#FFF3CD"),
        "Scaleup": colors.HexColor("#D4EDDA"),
        'The Big "A" Play': colors.HexColor("#D4EDDA"),
        "Fundraising": colors.HexColor("#CCE5FF"),
    }

    for i, rec in enumerate(records, 1):
        row_style = ParagraphStyle(
            "td", fontName=FONT_NAME, fontSize=8.5, textColor=BLACK, leading=12, alignment=1
        )
        row_left = ParagraphStyle(
            "tdl", fontName=FONT_NAME, fontSize=8.5, textColor=BLACK, leading=12, alignment=0
        )
        table_data.append(
            [
                Paragraph(str(i), row_style),
                make_company_link(rec["name"], rec.get("website", ""), row_left),
                Paragraph(rec["date"], row_style),
                Paragraph(rec["uic"], row_style),
                Paragraph(rec["tier"], row_style),
                Paragraph("", row_style),
            ]
        )

    table = Table(table_data, colWidths=col_widths, repeatRows=1)
    ts = TableStyle(
        [
            ("BACKGROUND", (0, 0), (-1, 0), HEADER_BG),
            ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
            ("FONTNAME", (0, 0), (-1, 0), FONT_BOLD),
            ("FONTSIZE", (0, 0), (-1, 0), 8.5),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, ROW_ALT]),
            ("GRID", (0, 0), (-1, -1), 0.5, LIGHT_GRAY),
            ("ALIGN", (0, 0), (0, -1), "CENTER"),
            ("VALIGN", (0, 1), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ]
    )

    for i, rec in enumerate(records, 1):
        ts.add("BACKGROUND", (4, i), (4, i), TIER_COLORS.get(rec["tier"], WHITE))

    table.setStyle(ts)
    story.append(table)

    story.append(Spacer(1, 18))
    footer_para = ParagraphStyle(
        "fp", fontName=FONT_NAME, fontSize=9.5, textColor=BLACK, leading=15, alignment=4, spaceAfter=8
    )
    story.append(
        Paragraph(
            "Кандидатите са заплатили съотвения членски внос, запознали са се с Устава, "
            "другите вътрешни актове на Сдружение, вкл. и Моралния кодекс на BESCO (Startup "
            "Bushido) и декларират, че ще ги спазват, в случай че бъдат приети за членове.",
            footer_para,
        )
    )

    story.append(Spacer(1, 30))

    today_str = date.today().strftime("%d.%m.%Y")
    sig_data = [
        [
            Paragraph(
                f"Дата: {today_str}",
                ParagraphStyle("sig", fontName=FONT_NAME, fontSize=9.5, textColor=BLACK, leading=14),
            ),
            Paragraph(
                "Подпис: ……………………<br/>(Изпълнителен директор на BESCO)",
                ParagraphStyle("sig2", fontName=FONT_NAME, fontSize=9.5, textColor=BLACK, leading=14, alignment=2),
            ),
        ]
    ]
    sig_table = Table(sig_data, colWidths=[W / 2, W / 2])
    sig_table.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("TOPPADDING", (0, 0), (-1, -1), 0)]))
    story.append(sig_table)

    # ------------------------------------------------------------------
    # Second section: protocol / decision pages
    # ------------------------------------------------------------------
    story.append(PageBreak())
    story.append(
        Paragraph(
            logo_text,
            ParagraphStyle(
                "logo2",
                fontName=FONT_BOLD,
                fontSize=28,
                textColor=BESCO_DARK,
                spaceAfter=20,
                leading=34,
            ),
        )
    )
    story.append(
        Paragraph(
            'Протокол от решение на Управителния съвет на Сдружение “Българската предприемаческа асоциация” (BESCO)',
            ParagraphStyle(
                "protocol_title",
                fontName=FONT_BOLD,
                fontSize=11.5,
                textColor=BLACK,
                alignment=1,
                spaceAfter=12,
                leading=15,
            ),
        )
    )
    protocol_date = end_dt.strftime("%d.%m.%Y")
    story.append(
        Paragraph(
            f'Днес, {protocol_date}, се проведе заседание на Управителния съвет на Сдружение “Българската предприемаческа асоциация” (BESCO), вписано в ТРРЮЛНЦ с ЕИК 177239971, със седалище и адрес на управление: гр. София, бул. Христо Ботев 117, ет. 2 (наричано по-долу за краткост “Сдружението”).',
            ParagraphStyle(
                "protocol_intro",
                fontName=FONT_NAME,
                fontSize=9.5,
                textColor=BLACK,
                leading=15,
                alignment=4,
                spaceAfter=10,
            ),
        )
    )
    story.append(Paragraph('Членовете на Управителния съвет:', ParagraphStyle('board_lbl', fontName=FONT_BOLD, fontSize=10.5, textColor=BLACK, leading=14, spaceAfter=6)))
    board_members = [
        'Росен Иванов',
        'Светозар Георгиев',
        'Пресиян Каракостов',
        'Йордан Матеев',
        'Велизар Величков',
        'Тенко Николов',
        'Милена Драгийска-Денчева',
        'Светла Костадинова',
        'Андрей Бъчваров',
        'Таня Бузева',
        'Донка Димитрова',
    ]
    board_text = '<br/>'.join(f'{i}. {m}' for i, m in enumerate(board_members, 1))
    story.append(Paragraph(board_text, ParagraphStyle('board_list', fontName=FONT_NAME, fontSize=9.3, textColor=BLACK, leading=14, leftIndent=14, spaceAfter=10)))
    story.append(Paragraph('КАТО ВЗЕХА ПРЕДВИД, ЧЕ:', ParagraphStyle('consider_lbl', fontName=FONT_BOLD, fontSize=10.5, textColor=BLACK, alignment=1, spaceAfter=6, leading=14)))
    considerations = [
        'Кандидатът за член подава писмена молба или попълва онлайн молба за членство до Управителния съвет, с която заявява, че желае да стане член на Сдружението',
        'Молбите за членство се разглеждат от Изпълнителния директор, който прави мотивирано предложение до Управителния съвет за приемане или неприемане на дадено лице за член на Сдружението',
        'Съгласно Устава на Сдружението кандидатът за нов член на Сдружението се приема с решение на Управителния съвет, взето с мнозинство от присъстващите.',
        'Членството възниква от датата на решението на Управителния съвет за приемане на новия член.',
    ]
    for c in considerations:
        story.append(Paragraph(f'• {c}', ParagraphStyle('consider_item', fontName=FONT_NAME, fontSize=9.2, textColor=BLACK, leading=14, leftIndent=8, spaceAfter=4)))
    story.append(Paragraph('Членовете на Управителния съвет приеха единодушно следния ДНЕВЕН РЕД:', ParagraphStyle('agenda_lbl', fontName=FONT_NAME, fontSize=9.5, textColor=BLACK, leading=14, spaceAfter=10)))
    story.append(Paragraph('1. Вземане на решение за приемане на нови членове на Сдружението, съгласно мотивирано предложение на Изпълнителния директор за периода 16 май - 15 юни 2026.', ParagraphStyle('agenda1', fontName=FONT_NAME, fontSize=9.2, textColor=BLACK, leading=14, leftIndent=14, spaceAfter=8)))
    story.append(Paragraph('2. Разни.', ParagraphStyle('agenda2', fontName=FONT_NAME, fontSize=9.2, textColor=BLACK, leading=14, leftIndent=14, spaceAfter=10)))
    story.append(Paragraph(f'След проведено обсъждане Управителният съвет прие единодушно следните РЕШЕНИЯ:', ParagraphStyle('decisions_lbl', fontName=FONT_BOLD, fontSize=10.5, textColor=BLACK, alignment=1, spaceAfter=10)))
    story.append(Paragraph('По точка 1', ParagraphStyle('pt1', fontName=FONT_BOLD, fontSize=9.8, textColor=BLACK, underlineWidth=0.5, spaceAfter=8)))
    story.append(Paragraph('След разглеждане на постъпилите кандидатури съгласно представеното мотивирано предложение на Изпълнителния директор, Управителният съвет приема за нови членове на Сдружението следните кандидати:', ParagraphStyle('pt1txt', fontName=FONT_NAME, fontSize=9.2, textColor=BLACK, leading=14, alignment=4, spaceAfter=10)))
    story.append(Table(table_data, colWidths=col_widths, repeatRows=1))
    story.append(Spacer(1, 12))
    story.append(Paragraph('По точка 2', ParagraphStyle('pt2', fontName=FONT_BOLD, fontSize=9.8, textColor=BLACK, underlineWidth=0.5, spaceAfter=8)))
    story.append(Paragraph('Други предложения и обсъждания не бяха правени.', ParagraphStyle('pt2txt', fontName=FONT_NAME, fontSize=9.2, textColor=BLACK, leading=14, alignment=4, spaceAfter=18)))
    story.append(Paragraph('Председателят на Управителния съвет закри заседанието поради изчерпване на Дневния ред.', ParagraphStyle('close1', fontName=FONT_NAME, fontSize=9.2, textColor=BLACK, leading=14, alignment=4, spaceAfter=16)))
    story.append(Paragraph('УПРАВИТЕЛЕН СЪВЕТ:', ParagraphStyle('board_hdr', fontName=FONT_BOLD, fontSize=10.5, textColor=BLACK, spaceAfter=8)))
    sig_rows = [[Paragraph(name, ParagraphStyle('sig_name', fontName=FONT_NAME, fontSize=9.2, textColor=BLACK, leading=12)), Paragraph('_________________________', ParagraphStyle('sig_line', fontName=FONT_NAME, fontSize=9.2, textColor=BLACK, alignment=0, leading=12))] for name in board_members]
    sig_tbl = Table(sig_rows, colWidths=[7.0 * cm, W - 7.0 * cm])
    sig_tbl.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'MIDDLE'),('TOPPADDING',(0,0),(-1,-1),2),('BOTTOMPADDING',(0,0),(-1,-1),2)]))
    story.append(sig_tbl)
    story.append(Spacer(1, 12))
    story.append(Paragraph(f'Дата: 17.06.2026', ParagraphStyle('date2', fontName=FONT_NAME, fontSize=9.2, textColor=BLACK, leading=14, alignment=0, spaceAfter=2)))
    story.append(Paragraph('Подпис: ……………………<br/>(Изпълнителен директор на BESCO)', ParagraphStyle('sig_exec', fontName=FONT_NAME, fontSize=9.2, textColor=BLACK, leading=14, alignment=2, spaceAfter=0)))

    def add_page_number(canvas, doc):
        canvas.saveState()
        canvas.setFont(FONT_NAME, 8)
        canvas.setFillColor(colors.HexColor("#888888"))
        canvas.drawRightString(A4[0] - 2.5 * cm, 1.5 * cm, str(doc.page))
        canvas.restoreState()

    doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
    print(f"✓  Report saved → {output_path}  ({len(records)} members)")


# -----------------------------------------------------------------------------
# CLI entry point
# -----------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Generate BESCO new members PDF report")
    parser.add_argument("--excel", required=True, help="Path to User_Membership_Database.xlsx")
    parser.add_argument("--start", required=True, help="Start date YYYY-MM-DD (inclusive)")
    parser.add_argument("--end", required=True, help="End date YYYY-MM-DD (inclusive)")
    parser.add_argument("--output", default="", help="Output PDF path (optional)")
    args = parser.parse_args()

    start_dt = datetime.strptime(args.start, "%Y-%m-%d").date()
    end_dt = datetime.strptime(args.end, "%Y-%m-%d").date()

    if not args.output:
        s = start_dt.strftime("%d_%m_%y")
        e = end_dt.strftime("%d_%m_%y")
        args.output = f"BESCO_New_Members_{s}-{e}.pdf"

    records = load_members(args.excel, start_dt, end_dt)

    if not records:
        print("⚠  No new members found for the specified period.")
        sys.exit(0)

    build_pdf(records, start_dt, end_dt, args.output)


if __name__ == "__main__":
    main()
