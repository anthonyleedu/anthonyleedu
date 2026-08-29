#!/usr/bin/env python3
"""Build 51Careers.AI investor presentation v.2.

Starts from the existing institutional deck and updates the round to a
RMB 200 million valuation with a 10% raise (RMB 20 million / ≈ US$2.97 million),
adding bilingual (RMB + US$) use-of-funds slides for the marketing budget.
"""

from __future__ import annotations

import shutil
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

SRC = Path("/workspace/51Careers_AI_Investor_Presentation.pptx")
OUT = Path("/workspace/51Careers_AI_Investor_Presentation v.2.pptx")

# Existing deck palette / type
NAVY = RGBColor(0x13, 0x2A, 0x4F)
GOLD = RGBColor(0x9A, 0x7B, 0x3F)
INK = RGBColor(0x1C, 0x27, 0x33)
MUTED = RGBColor(0x5A, 0x6B, 0x7E)
PANEL = RGBColor(0xF3, 0xF6, 0xF9)
SOFT = RGBColor(0xEE, 0xF3, 0xF9)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LINE = RGBColor(0xD5, 0xDC, 0xE4)
FONT_D = "Century Schoolbook"
FONT_B = "Calibri"

ML = Inches(0.62)
CONTENT_W = Inches(12.093)

# Illustrative FX used for the USD translations supplied by management:
# US$2,971,500 / RMB 20,000,000 = US$0.148575 per RMB  (≈ RMB 6.73 / US$1)
FX_NOTE = "USD equivalents at an illustrative rate of RMB 6.73 per US$1."


def set_run(run, text, font=FONT_B, size=11, bold=False, color=INK, italic=False):
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color


def add_text(shape, paragraphs, valign=MSO_ANCHOR.TOP):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    try:
        shape.text_frame._txBody.bodyPr.set(
            "anchor",
            {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}.get(
                valign, "t"
            ),
        )
    except Exception:
        pass
    first = True
    for spec in paragraphs:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = spec.get("align", PP_ALIGN.LEFT)
        if "space_after" in spec:
            p.space_after = Pt(spec["space_after"])
        if "space_before" in spec:
            p.space_before = Pt(spec["space_before"])
        run = p.add_run()
        set_run(
            run,
            spec["text"],
            font=spec.get("font", FONT_B),
            size=spec.get("size", 11),
            bold=spec.get("bold", False),
            color=spec.get("color", INK),
            italic=spec.get("italic", False),
        )


def fill_solid(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    try:
        shape.line.fill.background()
    except (AttributeError, TypeError, ValueError):
        pass


def rect(slide, l, t, w, h, fill=None):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    if fill is None:
        s.fill.background()
        s.line.fill.background()
    else:
        fill_solid(s, fill)
    return s


def textbox(slide, l, t, w, h, paragraphs, valign=MSO_ANCHOR.TOP):
    s = slide.shapes.add_textbox(l, t, w, h)
    add_text(s, paragraphs, valign=valign)
    return s


def add_header(slide, eyebrow: str, title: str):
    textbox(
        slide,
        ML,
        Inches(0.42),
        CONTENT_W,
        Inches(0.28),
        [{"text": eyebrow, "size": 10, "bold": True, "color": GOLD}],
    )
    textbox(
        slide,
        ML,
        Inches(0.72),
        CONTENT_W,
        Inches(0.50),
        [{"text": title, "size": 26, "bold": True, "color": NAVY, "font": FONT_D}],
    )


def add_footer(slide, page: int):
    textbox(
        slide,
        ML,
        Inches(7.13),
        Inches(5.50),
        Inches(0.25),
        [{"text": "51CAREERS.AI   ·   CONFIDENTIAL", "size": 7.5, "color": MUTED}],
    )
    textbox(
        slide,
        Inches(12.35),
        Inches(7.13),
        Inches(0.36),
        Inches(0.25),
        [{"text": str(page), "size": 8, "color": MUTED, "align": PP_ALIGN.RIGHT}],
    )


def replace_first_run(shape, new_text: str):
    """Replace text while preserving the first run's formatting."""
    tf = shape.text_frame
    first = True
    for p in tf.paragraphs:
        if first:
            if p.runs:
                p.runs[0].text = new_text
                for extra in p.runs[1:]:
                    extra.text = ""
            else:
                p.add_run().text = new_text
            first = False
        else:
            for r in p.runs:
                r.text = ""


def shape_by_name(slide, name: str):
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    raise KeyError(name)


def move_slide(prs: Presentation, old_index: int, new_index: int) -> None:
    xml_slides = prs.slides._sldIdLst  # type: ignore[attr-defined]
    slides = list(xml_slides)
    el = slides[old_index]
    xml_slides.remove(el)
    xml_slides.insert(new_index, el)


def update_existing_copy(prs: Presentation) -> None:
    # --- Slide 1: cover ---
    s1 = prs.slides[0]
    replace_first_run(
        shape_by_name(s1, "Text 11"),
        "Growth Financing  ·  RMB 200 Million Valuation",
    )
    replace_first_run(
        shape_by_name(s1, "Text 12"),
        "RMB 20 Million Raise (10%)   ·   ≈ US$2.97 Million   ·   July 2026",
    )

    # --- Slide 3: investment highlights ---
    s3 = prs.slides[2]
    replace_first_run(
        shape_by_name(s3, "Text 19"),
        "RMB 20 million raise (≈ US$3.0M) at a RMB 200 million valuation (≈ US$29.7M) to fund the product roadmap and global go-to-market.",
    )

    # --- Slide 13: the offering ---
    s13 = prs.slides[12]
    replace_first_run(
        shape_by_name(s13, "Text 1"),
        "RMB 20 Million to Accelerate Global Leadership",
    )
    replace_first_run(
        shape_by_name(s13, "Text 5"),
        "RMB 20.0M · 10% · ≈ US$2.97M",
    )
    replace_first_run(
        shape_by_name(s13, "Text 7"),
        "RMB 200M pre-money · ≈ US$29.7M",
    )
    replace_first_run(
        shape_by_name(s13, "Text 11"),
        "Personnel, marketing & product",
    )

    # Use-of-proceeds doughnut → actual budget mix of the RMB 20M raise
    for shape in s13.shapes:
        if not shape.has_chart:
            continue
        data = CategoryChartData()
        data.categories = [
            "Advertising & marketing",
            "Talent & personnel",
            "Product, expansion & reserves",
        ]
        data.add_series("Use of Proceeds", (40.0, 25.5, 34.5))
        shape.chart.replace_data(data)


def add_role_row(slide, x, y, w, name, count, monthly, annual):
    """One compact role line: title + count, then RMB/US$ monthly and annual."""
    textbox(
        slide,
        x,
        y,
        w - Inches(0.70),
        Inches(0.26),
        [{"text": name, "size": 11, "bold": True, "color": INK}],
    )
    textbox(
        slide,
        x + w - Inches(0.62),
        y,
        Inches(0.50),
        Inches(0.26),
        [{"text": count, "size": 11, "bold": True, "color": GOLD, "align": PP_ALIGN.RIGHT}],
    )
    textbox(
        slide,
        x,
        y + Inches(0.24),
        w,
        Inches(0.36),
        [
            {"text": monthly, "size": 10, "color": MUTED, "space_after": 0},
            {"text": annual, "size": 10, "color": MUTED},
        ],
    )


def add_section_label(slide, x, y, w, title):
    textbox(
        slide,
        x,
        y,
        w,
        Inches(0.28),
        [{"text": title, "size": 9, "bold": True, "color": GOLD}],
    )


def add_personnel_slide(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    add_header(slide, "USE OF FUNDS", "Marketing & Promotion Budget")
    textbox(
        slide,
        ML,
        Inches(1.20),
        CONTENT_W,
        Inches(0.28),
        [
            {
                "text": "Annual personnel costs — RMB with English-dollar (US$) translations",
                "size": 12,
                "color": MUTED,
            }
        ],
    )

    left_w = Inches(6.05)
    right_w = Inches(5.82)
    left_x = ML
    right_x = ML + left_w + Inches(0.22)
    panel_top = Inches(1.50)
    panel_h = Inches(4.80)

    rect(slide, left_x, panel_top, left_w, panel_h, fill=PANEL)
    rect(slide, right_x, panel_top, right_w, panel_h, fill=PANEL)

    # --- Left: China technical ---
    add_section_label(
        slide, left_x + Inches(0.22), panel_top + Inches(0.14), left_w - Inches(0.40),
        "CHINA-BASED TECHNICAL TEAM",
    )
    tech = [
        ("Mid-Level Product Manager", "1",
         "RMB 25,000 / month   ·   US$3,715",
         "RMB 300,000 / year    ·   US$44,573"),
        ("Designer", "1",
         "RMB 20,000 / month   ·   US$2,972",
         "RMB 240,000 / year    ·   US$35,658"),
        ("QA / Test Engineer", "1",
         "RMB 15,000 / month   ·   US$2,229",
         "RMB 180,000 / year    ·   US$26,744"),
        ("Front-End Engineer", "1",
         "RMB 25,000 / month   ·   US$3,715",
         "RMB 300,000 / year    ·   US$44,573"),
        ("Back-End Engineers", "2",
         "RMB 30,000 / month each   ·   US$4,458 each",
         "RMB 720,000 / year total   ·   US$107,334"),
        ("AI Engineer", "1",
         "RMB 40,000 / month   ·   US$5,943",
         "RMB 480,000 / year    ·   US$71,317"),
    ]
    row_h = Inches(0.70)
    for i, (name, n, monthly, annual) in enumerate(tech):
        add_role_row(
            slide,
            left_x + Inches(0.22),
            panel_top + Inches(0.48) + i * row_h,
            left_w - Inches(0.44),
            name, n, monthly, annual,
        )

    # --- Right: GTM + founders ---
    pad = Inches(0.22)
    y = panel_top + Inches(0.14)
    add_section_label(slide, right_x + pad, y, right_w - Inches(0.40),
                      "CHINA-BASED MARKETING & BUSINESS DEVELOPMENT")
    y += Inches(0.32)
    china_gtm = [
        ("Marketing / Advertising Specialist", "1",
         "RMB 20,000 / month   ·   US$2,972",
         "RMB 240,000 / year    ·   US$35,658"),
        ("B2B Business Development Specialist", "1",
         "RMB 20,000 / month   ·   US$2,972",
         "RMB 240,000 / year    ·   US$35,658"),
    ]
    for name, n, monthly, annual in china_gtm:
        add_role_row(slide, right_x + pad, y, right_w - Inches(0.44), name, n, monthly, annual)
        y += Inches(0.70)

    y += Inches(0.06)
    add_section_label(slide, right_x + pad, y, right_w - Inches(0.40),
                      "U.S.-BASED MARKETING & PROMOTION TEAM")
    y += Inches(0.32)
    us_gtm = [
        ("Advertising / Media Buying Specialist", "1",
         "RMB 80,000 / month   ·   US$11,886",
         "RMB 960,000 / year    ·   US$142,634"),
        ("Field Marketing Representatives", "2",
         "RMB 40,000 / month each   ·   US$5,943 each",
         "RMB 960,000 / year total   ·   US$142,634"),
    ]
    for name, n, monthly, annual in us_gtm:
        add_role_row(slide, right_x + pad, y, right_w - Inches(0.44), name, n, monthly, annual)
        y += Inches(0.70)

    y += Inches(0.06)
    add_section_label(slide, right_x + pad, y, right_w - Inches(0.40), "FOUNDING TEAM")
    y += Inches(0.32)
    add_role_row(
        slide, right_x + pad, y, right_w - Inches(0.44),
        "Founders — base salary", "4",
        "RMB 10,000 / month per person   ·   US$1,486",
        "RMB 480,000 / year total   ·   US$71,317",
    )

    # Total strip
    bar_y = Inches(6.38)
    rect(slide, ML, bar_y, CONTENT_W, Inches(0.40), fill=NAVY)
    textbox(
        slide,
        ML + Inches(0.22),
        bar_y + Inches(0.06),
        Inches(4.2),
        Inches(0.30),
        [{"text": "PERSONNEL SUBTOTAL  ·  16 PEOPLE", "size": 11, "bold": True, "color": WHITE}],
        valign=MSO_ANCHOR.MIDDLE,
    )
    textbox(
        slide,
        Inches(6.4),
        bar_y + Inches(0.06),
        Inches(6.1),
        Inches(0.30),
        [{"text": "RMB 5,100,000     ·     ≈ US$758,100 per year", "size": 12, "bold": True, "color": WHITE, "align": PP_ALIGN.RIGHT}],
        valign=MSO_ANCHOR.MIDDLE,
    )

    textbox(
        slide,
        ML,
        Inches(6.84),
        CONTENT_W,
        Inches(0.22),
        [{"text": f"Additional hires may be added as needed.  {FX_NOTE}", "size": 8, "color": MUTED}],
    )
    add_footer(slide, 14)


def add_budget_summary_slide(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    add_header(slide, "USE OF FUNDS", "Advertising Spend, Total Budget & Raise")

    cards = [
        (
            "CHINA ADVERTISING",
            "RMB 1.0M",
            "≈ US$148,575 per year",
            "Annual advertising & marketing spend in China.",
        ),
        (
            "UNITED STATES ADVERTISING",
            "RMB 7.0M",
            "≈ US$1,040,025 per year",
            "Annual advertising & marketing spend in the United States.",
        ),
        (
            "TOTAL ANNUAL BUDGET",
            "RMB 13.1M",
            "≈ US$1,946,333",
            "Personnel plus advertising — the operating plan this round funds.",
        ),
    ]
    card_w = Inches(3.85)
    gap = Inches(0.27)
    y = Inches(1.38)
    for i, (label, big, usd, desc) in enumerate(cards):
        x = ML + i * (card_w + gap)
        rect(slide, x, y, card_w, Inches(2.28), fill=PANEL)
        textbox(
            slide,
            x + Inches(0.22),
            y + Inches(0.16),
            card_w - Inches(0.40),
            Inches(0.26),
            [{"text": label, "size": 9, "bold": True, "color": GOLD}],
        )
        textbox(
            slide,
            x + Inches(0.22),
            y + Inches(0.44),
            card_w - Inches(0.40),
            Inches(0.58),
            [{"text": big, "size": 28, "bold": True, "color": NAVY, "font": FONT_D}],
        )
        textbox(
            slide,
            x + Inches(0.22),
            y + Inches(1.04),
            card_w - Inches(0.40),
            Inches(0.28),
            [{"text": usd, "size": 12, "bold": True, "color": INK}],
        )
        textbox(
            slide,
            x + Inches(0.22),
            y + Inches(1.36),
            card_w - Inches(0.40),
            Inches(0.72),
            [{"text": desc, "size": 11, "color": MUTED}],
        )

    # Two light panels: fundraising target + allocation
    bottom_y = Inches(3.84)
    bottom_h = Inches(3.00)
    left_w = Inches(6.05)
    right_w = Inches(5.82)
    right_x = ML + left_w + Inches(0.22)

    rect(slide, ML, bottom_y, left_w, bottom_h, fill=PANEL)
    textbox(
        slide,
        ML + Inches(0.28),
        bottom_y + Inches(0.18),
        left_w - Inches(0.50),
        Inches(0.26),
        [{"text": "FUNDRAISING TARGET", "size": 9, "bold": True, "color": GOLD}],
    )
    textbox(
        slide,
        ML + Inches(0.28),
        bottom_y + Inches(0.48),
        left_w - Inches(0.50),
        Inches(0.58),
        [{"text": "RMB 20,000,000", "size": 28, "bold": True, "color": NAVY, "font": FONT_D}],
    )
    textbox(
        slide,
        ML + Inches(0.28),
        bottom_y + Inches(1.10),
        left_w - Inches(0.50),
        Inches(0.32),
        [{"text": "≈ US$2,971,500", "size": 16, "bold": True, "color": INK}],
    )
    textbox(
        slide,
        ML + Inches(0.28),
        bottom_y + Inches(1.48),
        left_w - Inches(0.50),
        Inches(1.30),
        [
            {
                "text": "10% of a RMB 200,000,000 pre-money valuation (≈ US$29.7 million).",
                "size": 12,
                "color": MUTED,
                "space_after": 8,
            },
            {
                "text": "The target fundraising amount can be raised in two stages. Preferred equity; indicative close 2H 2026.",
                "size": 12,
                "color": MUTED,
            },
        ],
    )

    rect(slide, right_x, bottom_y, right_w, bottom_h, fill=PANEL)
    textbox(
        slide,
        right_x + Inches(0.28),
        bottom_y + Inches(0.18),
        right_w - Inches(0.50),
        Inches(0.26),
        [{"text": "ALLOCATION OF THE RMB 20 MILLION RAISE", "size": 9, "bold": True, "color": GOLD}],
    )
    allocations = [
        ("Advertising & marketing", "RMB 8.0M", "40%", "≈ US$1,188,600"),
        ("Talent & personnel", "RMB 5.1M", "25.5%", "≈ US$758,100"),
        ("Product, expansion & reserves", "RMB 6.9M", "34.5%", "≈ US$1,025,167"),
    ]
    for i, (name, amt, pct, usd) in enumerate(allocations):
        yy = bottom_y + Inches(0.56) + i * Inches(0.76)
        textbox(
            slide,
            right_x + Inches(0.28),
            yy,
            Inches(2.95),
            Inches(0.26),
            [{"text": name, "size": 12, "bold": True, "color": INK}],
        )
        textbox(
            slide,
            right_x + Inches(3.30),
            yy,
            Inches(2.20),
            Inches(0.26),
            [{"text": f"{amt}   {pct}", "size": 12, "bold": True, "color": NAVY, "align": PP_ALIGN.RIGHT}],
        )
        textbox(
            slide,
            right_x + Inches(0.28),
            yy + Inches(0.26),
            right_w - Inches(0.50),
            Inches(0.24),
            [{"text": usd, "size": 11, "color": MUTED}],
        )

    textbox(
        slide,
        ML,
        Inches(6.90),
        CONTENT_W,
        Inches(0.20),
        [{"text": FX_NOTE, "size": 8, "color": MUTED}],
    )
    add_footer(slide, 15)


def renumber_footers(prs: Presentation) -> None:
    for i, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            # Page-number boxes sit at the far right of the footer.
            if abs(shape.left.inches - 12.35) > 0.15:
                continue
            for p in shape.text_frame.paragraphs:
                for run in p.runs:
                    if run.text.strip().isdigit():
                        run.text = str(i)


def build() -> Path:
    shutil.copyfile(SRC, OUT)
    prs = Presentation(str(OUT))
    update_existing_copy(prs)
    add_personnel_slide(prs)
    add_budget_summary_slide(prs)
    # New slides were appended; place them after The Offering (index 12)
    # After two appends they sit at indices 15 and 16 (0-based).
    # Move last slide (budget summary, 16) to 13, then personnel (now 16) to 13.
    n = len(prs.slides)
    move_slide(prs, n - 2, 13)  # personnel → after offering
    move_slide(prs, n - 1, 14)  # budget summary → after personnel
    renumber_footers(prs)
    prs.save(str(OUT))
    return OUT


if __name__ == "__main__":
    path = build()
    print(f"Wrote {path}")
