#!/usr/bin/env python3
"""Update 51Careers_Addition.pptx with RMB 200M valuation / 10% raise.

Keeps the existing teal/Georgia deck. Updates round terms in place, fills
spare space on the offering/cover, and adds bilingual use-of-funds slides
in the same visual system.
"""

from __future__ import annotations

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

PATH = "/workspace/51Careers_Addition.pptx"

# This deck's visual system
NAVY = RGBColor(0x0B, 0x1F, 0x33)
TEAL = RGBColor(0x0A, 0x6B, 0x63)
TEAL_DK = RGBColor(0x07, 0x52, 0x4C)
INK = RGBColor(0x14, 0x28, 0x3C)
BODY = RGBColor(0x2C, 0x3A, 0x4B)
MUTED = RGBColor(0x5E, 0x6E, 0x80)
LINE = RGBColor(0xD4, 0xDE, 0xE8)
SOFT = RGBColor(0xEE, 0xF3, 0xF8)
PANEL = RGBColor(0xF5, 0xF8, 0xFB)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TEAL_SOFT = RGBColor(0xD7, 0xEB, 0xE8)
COVER_PANEL = RGBColor(0x14, 0x28, 0x3C)
FONT_D = "Georgia"
FONT_B = "Segoe UI"

SW, SH = Inches(13.333), Inches(7.5)
ML = Inches(0.55)
CONTENT_W = Inches(12.20)
FX_NOTE = "USD equivalents at an illustrative rate of RMB 6.73 per US$1."


def set_run(run, text, font=FONT_B, size=11, bold=False, color=BODY, italic=False):
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
            color=spec.get("color", BODY),
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


def round_rect(slide, l, t, w, h, fill=WHITE):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    fill_solid(s, fill)
    try:
        s.adjustments[0] = 0.08
    except Exception:
        pass
    return s


def textbox(slide, l, t, w, h, paragraphs, valign=MSO_ANCHOR.TOP):
    s = slide.shapes.add_textbox(l, t, w, h)
    add_text(s, paragraphs, valign=valign)
    return s


def set_slide_bg(slide, hex_color: str) -> None:
    cSld = slide._element.find(qn("p:cSld"))
    existing = cSld.find(qn("p:bg"))
    if existing is not None:
        cSld.remove(existing)
    bg = etree.Element(qn("p:bg"))
    bgPr = etree.SubElement(bg, qn("p:bgPr"))
    solid = etree.SubElement(bgPr, qn("a:solidFill"))
    srgb = etree.SubElement(solid, qn("a:srgbClr"))
    srgb.set("val", hex_color)
    etree.SubElement(bgPr, qn("a:effectLst"))
    spTree = cSld.find(qn("p:spTree"))
    spTree.addprevious(bg)


def send_to_index(slide, shape, index: int) -> None:
    """Place shape at z-order index among spTree children (after nvGrpSpPr/grpSpPr)."""
    spTree = slide.shapes._spTree
    el = shape._element
    spTree.remove(el)
    # children 0 = nvGrpSpPr, 1 = grpSpPr, then shapes
    spTree.insert(index, el)


def chrome(slide, page: int) -> None:
    rect(slide, Inches(0), Inches(0), SW, Inches(0.08), fill=TEAL)
    rect(slide, Inches(0), Inches(7.08), SW, Inches(0.42), fill=PANEL)
    rect(slide, Inches(0), Inches(7.08), SW, Inches(0.01), fill=LINE)
    textbox(
        slide,
        ML,
        Inches(7.14),
        Inches(6.00),
        Inches(0.28),
        [{"text": "51CAREERS.AI   ·   CONFIDENTIAL", "size": 8, "bold": True, "color": MUTED}],
    )
    textbox(
        slide,
        Inches(12.23),
        Inches(7.14),
        Inches(0.70),
        Inches(0.28),
        [{"text": str(page), "size": 9, "color": MUTED, "align": PP_ALIGN.RIGHT}],
    )


def header(slide, eyebrow: str, title: str) -> None:
    textbox(
        slide,
        ML,
        Inches(0.28),
        Inches(12.00),
        Inches(0.28),
        [{"text": eyebrow, "size": 10, "bold": True, "color": TEAL}],
    )
    textbox(
        slide,
        ML,
        Inches(0.52),
        CONTENT_W,
        Inches(0.55),
        [{"text": title, "size": 26, "bold": True, "color": NAVY, "font": FONT_D}],
    )


def replace_first_run(shape, new_text: str) -> None:
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


def set_two_line(shape, line1: str, line2: str) -> None:
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    p0 = tf.paragraphs[0]
    r0 = p0.add_run()
    set_run(r0, line1, font=FONT_B, size=13, bold=True, color=NAVY)
    p1 = tf.add_paragraph()
    r1 = p1.add_run()
    set_run(r1, line2, font=FONT_B, size=11, bold=False, color=MUTED)


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


def update_existing(prs: Presentation) -> None:
    # --- Cover ---
    s1 = prs.slides[0]
    banner = round_rect(s1, Inches(0.55), Inches(3.08), Inches(9.20), Inches(0.95), fill=COVER_PANEL)
    accent = rect(s1, Inches(0.55), Inches(3.08), Inches(0.10), Inches(0.95), fill=TEAL)
    send_to_index(s1, banner, 5)
    send_to_index(s1, accent, 6)

    t8 = shape_by_name(s1, "TextBox 8")
    t8.left, t8.top, t8.width, t8.height = Inches(0.85), Inches(3.16), Inches(8.70), Inches(0.32)
    replace_first_run(t8, "Growth Financing  ·  RMB 200 Million Valuation (≈ US$29.7M)")

    t9 = shape_by_name(s1, "TextBox 9")
    t9.left, t9.top, t9.width, t9.height = Inches(0.85), Inches(3.50), Inches(8.70), Inches(0.40)
    replace_first_run(t9, "RMB 20 Million Raise (10%)   ·   ≈ US$2.97 Million   ·   July 2026")

    # --- Highlights 06 ---
    s3 = prs.slides[2]
    replace_first_run(
        shape_by_name(s3, "TextBox 33"),
        "RMB 20 million raise (≈ US$3.0M) at a RMB 200 million valuation (≈ US$29.7M) to fund product, AI features, and global user growth.",
    )

    # --- Offering ---
    s13 = prs.slides[12]
    replace_first_run(
        shape_by_name(s13, "TextBox 3"),
        "RMB 20 Million to Accelerate Global Leadership",
    )

    # Taller bilingual value cells to fill the terms card
    for name, line1, line2, top in [
        ("TextBox 8", "RMB 20.0 million  ·  10%", "≈ US$2,971,500 (primary)", Inches(1.92)),
        ("TextBox 11", "RMB 200 million pre-money", "≈ US$29.7 million", Inches(2.58)),
        ("TextBox 17", "Personnel, marketing & product", "Team · ads · platform build", Inches(3.88)),
    ]:
        box = shape_by_name(s13, name)
        box.top = top
        box.height = Inches(0.58)
        box.width = Inches(3.00)
        set_two_line(box, line1, line2)

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
    textbox(
        slide, x, y, w - Inches(0.55), Inches(0.28),
        [{"text": name, "size": 13, "bold": True, "color": NAVY}],
    )
    textbox(
        slide, x + w - Inches(0.50), y, Inches(0.48), Inches(0.28),
        [{"text": count, "size": 13, "bold": True, "color": TEAL, "align": PP_ALIGN.RIGHT}],
    )
    textbox(
        slide, x, y + Inches(0.26), w, Inches(0.44),
        [
            {"text": monthly, "size": 11, "color": MUTED, "space_after": 0},
            {"text": annual, "size": 11, "color": MUTED},
        ],
    )


def add_personnel_slide(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, "F5F8FB")
    header(slide, "USE OF FUNDS", "Marketing & Promotion Budget")
    textbox(
        slide, ML, Inches(1.08), CONTENT_W, Inches(0.28),
        [{"text": "Annual personnel costs — RMB with English-dollar (US$) translations", "size": 13, "color": MUTED}],
    )

    left_w, right_w = Inches(6.15), Inches(5.90)
    left_x, right_x = ML, ML + left_w + Inches(0.18)
    top, h = Inches(1.40), Inches(4.88)

    round_rect(slide, left_x, top, left_w, h, fill=WHITE)
    rect(slide, left_x, top, Inches(0.08), h, fill=TEAL)
    round_rect(slide, right_x, top, right_w, h, fill=WHITE)
    rect(slide, right_x, top, Inches(0.08), h, fill=TEAL)

    # Left: China technical
    lx = left_x + Inches(0.28)
    lw = left_w - Inches(0.48)
    textbox(slide, lx, top + Inches(0.16), lw, Inches(0.28),
            [{"text": "CHINA-BASED TECHNICAL TEAM", "size": 11, "bold": True, "color": TEAL}])
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
    for i, row in enumerate(tech):
        add_role_row(slide, lx, top + Inches(0.48) + i * Inches(0.70), lw, *row)

    # Right: GTM + founders
    rx = right_x + Inches(0.28)
    rw = right_w - Inches(0.48)
    y = top + Inches(0.16)
    textbox(slide, rx, y, rw, Inches(0.28),
            [{"text": "CHINA-BASED MARKETING & BUSINESS DEVELOPMENT", "size": 11, "bold": True, "color": TEAL}])
    y += Inches(0.30)
    for row in [
        ("Marketing / Advertising Specialist", "1",
         "RMB 20,000 / month   ·   US$2,972",
         "RMB 240,000 / year    ·   US$35,658"),
        ("B2B Business Development Specialist", "1",
         "RMB 20,000 / month   ·   US$2,972",
         "RMB 240,000 / year    ·   US$35,658"),
    ]:
        add_role_row(slide, rx, y, rw, *row)
        y += Inches(0.70)

    y += Inches(0.04)
    textbox(slide, rx, y, rw, Inches(0.28),
            [{"text": "U.S.-BASED MARKETING & PROMOTION TEAM", "size": 11, "bold": True, "color": TEAL}])
    y += Inches(0.30)
    for row in [
        ("Advertising / Media Buying Specialist", "1",
         "RMB 80,000 / month   ·   US$11,886",
         "RMB 960,000 / year    ·   US$142,634"),
        ("Field Marketing Representatives", "2",
         "RMB 40,000 / month each   ·   US$5,943 each",
         "RMB 960,000 / year total   ·   US$142,634"),
    ]:
        add_role_row(slide, rx, y, rw, *row)
        y += Inches(0.70)

    y += Inches(0.04)
    textbox(slide, rx, y, rw, Inches(0.28),
            [{"text": "FOUNDING TEAM", "size": 11, "bold": True, "color": TEAL}])
    y += Inches(0.30)
    add_role_row(
        slide, rx, y, rw,
        "Founders — base salary", "4",
        "RMB 10,000 / month per person   ·   US$1,486",
        "RMB 480,000 / year total   ·   US$71,317",
    )

    bar = round_rect(slide, ML, Inches(6.36), CONTENT_W, Inches(0.44), fill=TEAL)
    textbox(
        slide, ML + Inches(0.22), Inches(6.40), Inches(5.4), Inches(0.36),
        [{"text": "PERSONNEL SUBTOTAL  ·  16 PEOPLE", "size": 13, "bold": True, "color": WHITE}],
        valign=MSO_ANCHOR.MIDDLE,
    )
    textbox(
        slide, Inches(6.3), Inches(6.40), Inches(6.2), Inches(0.36),
        [{"text": "RMB 5,100,000     ·     ≈ US$758,100 / year", "size": 14, "bold": True, "color": WHITE, "align": PP_ALIGN.RIGHT}],
        valign=MSO_ANCHOR.MIDDLE,
    )
    textbox(
        slide, ML, Inches(6.84), CONTENT_W, Inches(0.22),
        [{"text": f"Additional hires may be added as needed.  {FX_NOTE}", "size": 8, "color": MUTED}],
    )
    chrome(slide, 14)


def add_budget_slide(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, "F5F8FB")
    header(slide, "USE OF FUNDS", "Advertising Spend, Total Budget & Raise")

    cards = [
        ("CHINA ADVERTISING", "RMB 1.0M", "≈ US$148,575 / year",
         "Annual advertising & marketing spend in China."),
        ("UNITED STATES ADVERTISING", "RMB 7.0M", "≈ US$1,040,025 / year",
         "Annual advertising & marketing spend in the United States."),
        ("TOTAL ANNUAL BUDGET", "RMB 13.1M", "≈ US$1,946,333",
         "Personnel plus advertising — the operating plan this round funds."),
    ]
    card_w, gap = Inches(3.95), Inches(0.18)
    y = Inches(1.22)
    for i, (label, big, usd, desc) in enumerate(cards):
        x = ML + i * (card_w + gap)
        round_rect(slide, x, y, card_w, Inches(2.35), fill=WHITE)
        rect(slide, x, y, card_w, Inches(0.10), fill=TEAL)
        textbox(slide, x + Inches(0.22), y + Inches(0.22), card_w - Inches(0.40), Inches(0.26),
                [{"text": label, "size": 10, "bold": True, "color": TEAL}])
        textbox(slide, x + Inches(0.22), y + Inches(0.50), card_w - Inches(0.40), Inches(0.58),
                [{"text": big, "size": 28, "bold": True, "color": NAVY, "font": FONT_D}])
        textbox(slide, x + Inches(0.22), y + Inches(1.12), card_w - Inches(0.40), Inches(0.28),
                [{"text": usd, "size": 13, "bold": True, "color": INK}])
        textbox(slide, x + Inches(0.22), y + Inches(1.46), card_w - Inches(0.40), Inches(0.70),
                [{"text": desc, "size": 12, "color": MUTED}])

    bottom_y, bottom_h = Inches(3.70), Inches(3.08)
    left_w, right_w = Inches(6.15), Inches(5.90)
    right_x = ML + left_w + Inches(0.18)

    round_rect(slide, ML, bottom_y, left_w, bottom_h, fill=WHITE)
    rect(slide, ML, bottom_y, Inches(0.08), bottom_h, fill=TEAL)
    textbox(slide, ML + Inches(0.30), bottom_y + Inches(0.20), left_w - Inches(0.50), Inches(0.26),
            [{"text": "FUNDRAISING TARGET", "size": 11, "bold": True, "color": TEAL}])
    textbox(slide, ML + Inches(0.30), bottom_y + Inches(0.52), left_w - Inches(0.50), Inches(0.62),
            [{"text": "RMB 20,000,000", "size": 32, "bold": True, "color": NAVY, "font": FONT_D}])
    textbox(slide, ML + Inches(0.30), bottom_y + Inches(1.18), left_w - Inches(0.50), Inches(0.36),
            [{"text": "≈ US$2,971,500", "size": 18, "bold": True, "color": INK}])
    textbox(
        slide, ML + Inches(0.30), bottom_y + Inches(1.62), left_w - Inches(0.50), Inches(1.35),
        [
            {"text": "10% of a RMB 200,000,000 pre-money valuation (≈ US$29.7 million).",
             "size": 13, "color": BODY, "space_after": 8},
            {"text": "The target fundraising amount can be raised in two stages. Preferred equity; indicative close 2H 2026.",
             "size": 13, "color": MUTED},
        ],
    )

    round_rect(slide, right_x, bottom_y, right_w, bottom_h, fill=WHITE)
    rect(slide, right_x, bottom_y, Inches(0.08), bottom_h, fill=TEAL)
    textbox(slide, right_x + Inches(0.30), bottom_y + Inches(0.20), right_w - Inches(0.50), Inches(0.26),
            [{"text": "ALLOCATION OF THE RMB 20 MILLION RAISE", "size": 11, "bold": True, "color": TEAL}])
    allocations = [
        ("Advertising & marketing", "RMB 8.0M", "40%", "≈ US$1,188,600"),
        ("Talent & personnel", "RMB 5.1M", "25.5%", "≈ US$758,100"),
        ("Product, expansion & reserves", "RMB 6.9M", "34.5%", "≈ US$1,025,167"),
    ]
    for i, (name, amt, pct, usd) in enumerate(allocations):
        yy = bottom_y + Inches(0.58) + i * Inches(0.82)
        textbox(slide, right_x + Inches(0.30), yy, Inches(2.85), Inches(0.28),
                [{"text": name, "size": 13, "bold": True, "color": NAVY}])
        textbox(slide, right_x + Inches(3.20), yy, Inches(2.40), Inches(0.28),
                [{"text": f"{amt}   {pct}", "size": 13, "bold": True, "color": TEAL, "align": PP_ALIGN.RIGHT}])
        textbox(slide, right_x + Inches(0.30), yy + Inches(0.28), right_w - Inches(0.50), Inches(0.26),
                [{"text": usd, "size": 12, "color": MUTED}])

    textbox(slide, ML, Inches(6.84), CONTENT_W, Inches(0.22),
            [{"text": FX_NOTE, "size": 8, "color": MUTED}])
    chrome(slide, 15)


def renumber(prs: Presentation) -> None:
    for i, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            # Standard footer page box
            if abs(shape.left.inches - 12.23) < 0.15 and abs(shape.top.inches - 7.14) < 0.15:
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        if run.text.strip().isdigit():
                            run.text = str(i)
            # Closing slide embeds the page number in the confidential line
            full = "".join(p.text for p in shape.text_frame.paragraphs)
            if "CONFIDENTIAL   ·   " in full and "51CAREERS.AI" in full:
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        run.text = run.text.replace("·   15", f"·   {i}").replace("·  15", f"·  {i}")


def main() -> None:
    prs = Presentation(PATH)
    update_existing(prs)
    add_personnel_slide(prs)
    add_budget_slide(prs)
    n = len(prs.slides)
    move_slide(prs, n - 2, 13)
    move_slide(prs, n - 1, 14)
    renumber(prs)
    prs.save(PATH)
    print(f"Updated {PATH} ({len(prs.slides)} slides)")


if __name__ == "__main__":
    main()
