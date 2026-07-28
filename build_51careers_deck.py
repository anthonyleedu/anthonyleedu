#!/usr/bin/env python3
"""Rebuild 51Careers.AI investor presentation — institutional visual system."""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import nsmap, qn
from pptx.util import Emu, Inches, Pt
from lxml import etree

OUT = Path("/workspace/51Careers_AI_Investor_Presentation.pptx")
SW, SH = Inches(13.333), Inches(7.5)

# Typography — different from Century Schoolbook / Calibri
FONT_DISPLAY = "Georgia"
FONT_BODY = "Segoe UI"

# Palette — institutional navy + deep teal (no purple / cream clichés)
NAVY = RGBColor(0x0B, 0x1F, 0x33)
INK = RGBColor(0x14, 0x28, 0x3C)
BODY = RGBColor(0x2C, 0x3A, 0x4B)
MUTED = RGBColor(0x5E, 0x6E, 0x80)
LINE = RGBColor(0xD4, 0xDE, 0xE8)
SOFT = RGBColor(0xEE, 0xF3, 0xF8)
PANEL = RGBColor(0xF5, 0xF8, 0xFB)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TEAL = RGBColor(0x0A, 0x6B, 0x63)
TEAL_DK = RGBColor(0x07, 0x52, 0x4C)
TEAL_SOFT = RGBColor(0xD7, 0xEB, 0xE8)
BLUE = RGBColor(0x2F, 0x5D, 0x8A)
BRONZE = RGBColor(0x8A, 0x6D, 0x3B)
SLATE = RGBColor(0x6B, 0x7C, 0x8F)
CHART_COLORS = [TEAL, BLUE, BRONZE, SLATE, RGBColor(0xA8, 0xB8, 0xC6)]


def rgb_hex(c: RGBColor) -> str:
    return f"{c[0]:02X}{c[1]:02X}{c[2]:02X}"


def set_run(run, text, font=FONT_BODY, size=11, bold=False, color=BODY, italic=False):
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color


def clear_tf(tf):
    tf.clear()
    p = tf.paragraphs[0]
    p.clear()
    return p


def add_text(shape, paragraphs, valign=MSO_ANCHOR.TOP):
    """paragraphs: list of dicts with text/font/size/bold/color/align/space_after/italic"""
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    try:
        shape.text_frame._txBody.bodyPr.set("anchor", {
            MSO_ANCHOR.TOP: "t",
            MSO_ANCHOR.MIDDLE: "ctr",
            MSO_ANCHOR.BOTTOM: "b",
        }.get(valign, "t"))
    except Exception:
        pass
    first = True
    for spec in paragraphs:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.alignment = spec.get("align", PP_ALIGN.LEFT)
        if "space_after" in spec:
            p.space_after = Pt(spec["space_after"])
        if "space_before" in spec:
            p.space_before = Pt(spec["space_before"])
        run = p.add_run()
        set_run(
            run,
            spec["text"],
            font=spec.get("font", FONT_BODY),
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


def stroke(shape, color=LINE, width_pt=1.0):
    shape.line.color.rgb = color
    shape.line.width = Pt(width_pt)


def rect(slide, l, t, w, h, fill=None, line=None, line_w=1.0):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    if fill is None:
        s.fill.background()
    else:
        fill_solid(s, fill)
    if line is None:
        s.line.fill.background()
    else:
        stroke(s, line, line_w)
    return s


def round_rect(slide, l, t, w, h, fill=WHITE, line=None, line_w=1.0):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    fill_solid(s, fill)
    # tighten radius
    try:
        s.adjustments[0] = 0.08
    except Exception:
        pass
    if line is None:
        s.line.fill.background()
    else:
        stroke(s, line, line_w)
    return s


def textbox(slide, l, t, w, h, paragraphs, valign=MSO_ANCHOR.TOP):
    s = slide.shapes.add_textbox(l, t, w, h)
    add_text(s, paragraphs, valign=valign)
    return s


def accent_bar(slide, l, t, w=Inches(0.08), h=Inches(0.35), color=TEAL):
    return rect(slide, l, t, w, h, fill=color)


def footer(slide, page: int, total: int = 15):
    rect(slide, Inches(0), SH - Inches(0.42), SW, Inches(0.42), fill=PANEL)
    rect(slide, Inches(0), SH - Inches(0.42), SW, Pt(1), fill=LINE)
    textbox(
        slide,
        Inches(0.55),
        SH - Inches(0.36),
        Inches(6),
        Inches(0.28),
        [{"text": "51CAREERS.AI   ·   CONFIDENTIAL", "size": 8, "color": MUTED, "bold": True}],
    )
    textbox(
        slide,
        SW - Inches(1.1),
        SH - Inches(0.36),
        Inches(0.7),
        Inches(0.28),
        [{"text": str(page), "size": 9, "color": MUTED, "align": PP_ALIGN.RIGHT}],
    )


def section_header(slide, eyebrow: str, title: str, subtitle: str | None = None):
    rect(slide, Inches(0), Inches(0), SW, Inches(0.08), fill=TEAL)
    textbox(
        slide,
        Inches(0.55),
        Inches(0.28),
        Inches(12),
        Inches(0.28),
        [{"text": eyebrow.upper(), "size": 10, "bold": True, "color": TEAL, "font": FONT_BODY}],
    )
    textbox(
        slide,
        Inches(0.55),
        Inches(0.52),
        Inches(12.2),
        Inches(0.55),
        [{"text": title, "size": 28, "bold": True, "color": NAVY, "font": FONT_DISPLAY}],
    )
    if subtitle:
        textbox(
            slide,
            Inches(0.55),
            Inches(1.05),
            Inches(12.2),
            Inches(0.35),
            [{"text": subtitle, "size": 12, "color": MUTED}],
        )


def style_chart(chart, has_legend=True, point_colors=None):
    chart.has_legend = has_legend
    chart.has_title = False
    if has_legend:
        chart.legend.include_in_layout = False
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.font.size = Pt(9)
        chart.legend.font.name = FONT_BODY
        chart.legend.font.color.rgb = MUTED
    try:
        plot = chart.plots[0]
        plot.has_data_labels = False
    except Exception:
        pass
    # color series
    try:
        for i, series in enumerate(chart.series):
            color = CHART_COLORS[i % len(CHART_COLORS)]
            solid = series.format.fill
            solid.solid()
            solid.fore_color.rgb = color
            try:
                series.format.line.color.rgb = color
            except Exception:
                pass
            if point_colors:
                for j, pt in enumerate(series.points):
                    c = point_colors[j % len(point_colors)]
                    pt.format.fill.solid()
                    pt.format.fill.fore_color.rgb = c
    except Exception:
        pass
    # axes
    try:
        cat = chart.category_axis
        cat.has_major_gridlines = False
        cat.tick_labels.font.size = Pt(9)
        cat.tick_labels.font.name = FONT_BODY
        cat.tick_labels.font.color.rgb = MUTED
        cat.format.line.color.rgb = LINE
    except Exception:
        pass
    try:
        val = chart.value_axis
        val.has_major_gridlines = True
        val.major_gridlines.format.line.color.rgb = LINE
        val.tick_labels.font.size = Pt(9)
        val.tick_labels.font.name = FONT_BODY
        val.tick_labels.font.color.rgb = MUTED
        val.format.line.color.rgb = LINE
    except Exception:
        pass


def add_chart(slide, chart_type, left, top, width, height, categories, series_data, legend=True, point_colors=None, value_max=None):
    """series_data: list of (name, values)"""
    data = CategoryChartData()
    data.categories = categories
    for name, values in series_data:
        data.add_series(name, values)
    chart_shape = slide.shapes.add_chart(chart_type, left, top, width, height, data)
    style_chart(chart_shape.chart, has_legend=legend, point_colors=point_colors)
    if value_max is not None:
        try:
            chart_shape.chart.value_axis.maximum_scale = value_max
            chart_shape.chart.value_axis.minimum_scale = 0
        except Exception:
            pass
    return chart_shape


def blank_slide(prs):
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    fill_solid(slide.background, PANEL)
    return slide


# ───────────────────────────── SLIDES ─────────────────────────────


def slide_cover(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fill_solid(slide.background, NAVY)
    rect(slide, Inches(0), Inches(0), SW, Inches(0.12), fill=TEAL)
    rect(slide, Inches(0), SH - Inches(1.35), SW, Inches(1.35), fill=RGBColor(0x08, 0x18, 0x28))
    rect(slide, Inches(0), Inches(0), Inches(0.18), SH, fill=TEAL)

    textbox(
        slide, Inches(0.7), Inches(0.5), Inches(10), Inches(0.3),
        [{"text": "CONFIDENTIAL INVESTOR PRESENTATION", "size": 11, "bold": True, "color": TEAL_SOFT}],
    )
    textbox(
        slide, Inches(0.7), Inches(1.15), Inches(11), Inches(0.75),
        [{"text": "51CAREERS.AI", "size": 46, "bold": True, "color": WHITE, "font": FONT_DISPLAY}],
    )
    textbox(
        slide, Inches(0.7), Inches(1.95), Inches(11), Inches(0.4),
        [{"text": "The AI-Powered Global Career Platform", "size": 18, "color": RGBColor(0xC5, 0xD5, 0xE0)}],
    )
    textbox(
        slide, Inches(0.7), Inches(2.45), Inches(11.5), Inches(0.55),
        [{"text": "Who we are — An AI career platform from the 51Careers team: mock interviews, a 100K+ question bank, and a new AI résumé builder — free today, built for global professionals.",
          "size": 13, "color": RGBColor(0xA8, 0xBE, 0xCC)}],
    )
    textbox(
        slide, Inches(0.7), Inches(3.15), Inches(8), Inches(0.3),
        [{"text": "Growth Financing", "size": 13, "bold": True, "color": TEAL}],
    )
    textbox(
        slide, Inches(0.7), Inches(3.45), Inches(9), Inches(0.3),
        [{"text": "US$5.0 Million Private Placement   ·   July 2026", "size": 14, "color": WHITE}],
    )

    kpis = [
        ("2025", "AI PLATFORM LAUNCHED"),
        ("10,000+", "CANDIDATES HELPED · 8 YRS"),
        ("100K+", "MOCK INTERVIEW QUESTIONS"),
        ("US$0", "PLATFORM REVENUE TODAY"),
    ]
    x = Inches(0.55)
    for val, label in kpis:
        round_rect(slide, x, Inches(4.15), Inches(2.95), Inches(1.35), fill=RGBColor(0x12, 0x2A, 0x40))
        textbox(
            slide, x + Inches(0.18), Inches(4.3), Inches(2.6), Inches(0.55),
            [{"text": val, "size": 26, "bold": True, "color": WHITE, "font": FONT_DISPLAY}],
        )
        textbox(
            slide, x + Inches(0.18), Inches(4.95), Inches(2.6), Inches(0.4),
            [{"text": label, "size": 9, "bold": True, "color": TEAL_SOFT}],
        )
        x += Inches(3.15)

    textbox(
        slide, Inches(0.7), Inches(5.7), Inches(11), Inches(0.25),
        [{"text": "Currently free to grow active users · Monetization via membership, consulting & education fees", "size": 11, "color": SLATE}],
    )
    textbox(
        slide, Inches(0.7), SH - Inches(0.85), Inches(12), Inches(0.35),
        [{
            "text": "NEW YORK  ·  SHANGHAI  ·  LONDON  ·  SINGAPORE  ·  SEOUL  ·  TORONTO  ·  DELHI",
            "size": 10, "bold": True, "color": RGBColor(0x9A, 0xB0, 0xC0),
        }],
    )


def slide_disclaimer(prs):
    slide = blank_slide(prs)
    section_header(slide, "Important Notice", "Disclaimer")
    body = (
        "This presentation has been prepared by 51 Careers Inc. and its affiliates (together, “51Careers.AI” or the “Company”) "
        "solely for informational purposes in connection with a proposed private financing. It is confidential and intended "
        "exclusively for the named recipient; it may not be reproduced, distributed, or disclosed, in whole or in part, without "
        "the Company’s prior written consent.\n\n"
        "This document does not constitute an offer to sell, or a solicitation of an offer to buy, any securities in any "
        "jurisdiction. Any such offer will be made only pursuant to definitive documentation and only to qualified investors in "
        "compliance with applicable securities laws.\n\n"
        "Certain statements herein are forward-looking and reflect current expectations of management. Such statements involve "
        "known and unknown risks and uncertainties, and actual results may differ materially. Certain market, operating, and "
        "financial figures are estimates — including third-party estimates — presented for illustration only and remain subject "
        "to confirmation in due diligence. No representation or warranty, express or implied, is made as to the accuracy or "
        "completeness of the information contained herein, and nothing herein constitutes legal, tax, financial, or investment "
        "advice. Recipients should conduct their own independent investigation and analysis."
    )
    panel = round_rect(slide, Inches(0.55), Inches(1.35), Inches(12.2), Inches(4.8), fill=WHITE, line=LINE)
    textbox(
        slide,
        Inches(0.85),
        Inches(1.6),
        Inches(11.6),
        Inches(4.2),
        [{"text": body, "size": 12, "color": BODY}],
    )
    textbox(
        slide,
        Inches(0.55),
        Inches(6.35),
        Inches(8),
        Inches(0.3),
        [{"text": "Prepared by 51Careers.AI — July 2026", "size": 10, "color": MUTED}],
    )
    footer(slide, 2)


def slide_highlights(prs):
    slide = blank_slide(prs)
    section_header(slide, "Executive Summary", "Who We Are & Why Now")
    items = [
        ("01", "Who We Are",
         "51Careers.AI is the AI-powered career platform from the 51Careers team — helping professionals worldwide prepare, apply, and compete for great jobs."),
        ("02", "Our Product",
         "AI mock interviews, a proprietary 100K+ interview-question database, and a new AI résumé builder — plus coaching pathways as users grow."),
        ("03", "Proven Demand Base",
         "Over 10,000 candidates helped across eight years of career-development operations; the AI platform launched in 2025 and is free today."),
        ("04", "Global Market",
         "A multi-hundred-billion-dollar global recruitment and career-development market shifting rapidly toward AI-native tools."),
        ("05", "Clear Monetization Path",
         "Grow active users on a free product, then introduce membership fees, consulting fees, and education fees as engagement compounds."),
        ("06", "Capital to Scale",
         "US$5.0M primary round at a US$50.0M valuation to fund product, AI features, and global user growth."),
    ]
    positions = [
        (0.55, 1.35), (4.55, 1.35), (8.55, 1.35),
        (0.55, 4.05), (4.55, 4.05), (8.55, 4.05),
    ]
    for (num, title, desc), (x, y) in zip(items, positions):
        round_rect(slide, Inches(x), Inches(y), Inches(3.75), Inches(2.4), fill=WHITE, line=LINE)
        accent_bar(slide, Inches(x), Inches(y), w=Inches(0.08), h=Inches(2.4))
        textbox(slide, Inches(x + 0.25), Inches(y + 0.2), Inches(3.2), Inches(0.4),
                [{"text": num, "size": 18, "bold": True, "color": TEAL, "font": FONT_DISPLAY}])
        textbox(slide, Inches(x + 0.25), Inches(y + 0.7), Inches(3.2), Inches(0.45),
                [{"text": title, "size": 14, "bold": True, "color": NAVY}])
        textbox(slide, Inches(x + 0.25), Inches(y + 1.2), Inches(3.2), Inches(1.0),
                [{"text": desc, "size": 11, "color": BODY}])
    footer(slide, 3)


def slide_problem(prs):
    slide = blank_slide(prs)
    section_header(slide, "Market Context", "Hiring Is Broken on Both Sides of the Table")

    narratives = [
        ("For candidates",
         "Hiring processes are opaque and intensely competitive worldwide. Interview preparation remains fragmented, and quality coaching is expensive and does not scale."),
        ("For global talent",
         "Language, cultural, and cross-border complexity compound an already difficult search — hundreds of millions of professionals are underserved by generic AI chat tools."),
        ("For employers",
         "Application volumes overwhelm recruiting teams while offering limited signal. Screening is slow, costly, and inconsistent, and strong candidates are lost in the noise."),
    ]
    y = 1.35
    for title, body in narratives:
        round_rect(slide, Inches(0.55), Inches(y), Inches(6.3), Inches(1.5), fill=WHITE, line=LINE)
        accent_bar(slide, Inches(0.55), Inches(y), h=Inches(1.5))
        textbox(slide, Inches(0.85), Inches(y + 0.18), Inches(5.7), Inches(0.35),
                [{"text": title, "size": 14, "bold": True, "color": NAVY}])
        textbox(slide, Inches(0.85), Inches(y + 0.55), Inches(5.7), Inches(0.85),
                [{"text": body, "size": 11, "color": BODY}])
        y += 1.65

    round_rect(slide, Inches(7.1), Inches(1.35), Inches(5.65), Inches(5.0), fill=WHITE, line=LINE)
    textbox(slide, Inches(7.35), Inches(1.5), Inches(5.2), Inches(0.35),
            [{"text": "Global friction indicators", "size": 13, "bold": True, "color": NAVY}])
    metrics = [
        ("250+", "Applications per corporate job opening (global)", 0.72, TEAL),
        ("6M+", "International students worldwide (annual)", 0.85, BLUE),
        ("75%", "Of résumés screened out before human review", 0.75, BRONZE),
    ]
    y = 2.0
    for value, label, frac, color in metrics:
        textbox(slide, Inches(7.45), Inches(y), Inches(5.1), Inches(0.45),
                [{"text": value, "size": 28, "bold": True, "color": NAVY, "font": FONT_DISPLAY}])
        textbox(slide, Inches(7.45), Inches(y + 0.48), Inches(5.1), Inches(0.3),
                [{"text": label.upper(), "size": 9, "bold": True, "color": MUTED}])
        rect(slide, Inches(7.45), Inches(y + 0.85), Inches(5.0), Inches(0.12), fill=SOFT)
        rect(slide, Inches(7.45), Inches(y + 0.85), Inches(5.0 * frac), Inches(0.12), fill=color)
        y += 1.25
    textbox(slide, Inches(7.35), Inches(5.85), Inches(5.2), Inches(0.35),
            [{"text": "Sources: UNESCO / OECD mobility data; Glassdoor; industry estimates.", "size": 8, "color": MUTED}])
    footer(slide, 4)


def slide_market(prs):
    slide = blank_slide(prs)
    section_header(slide, "Market Opportunity", "A Global Career Market in Structural Shift")

    markets = [
        ("TOTAL ADDRESSABLE MARKET", "US$200B+", "Global recruitment, talent-acquisition, and career-development spend worldwide.", TEAL, 1.0),
        ("SERVICEABLE MARKET", "US$30B+", "Global AI-enabled hiring technology and digital career preparation.", BLUE, 0.58),
        ("NEAR-TERM FOCUS", "US$5B+", "Global professionals seeking AI prep, résumés, and interview readiness.", BRONZE, 0.34),
    ]
    y = 1.32
    for label, value, desc, color, width_frac in markets:
        round_rect(slide, Inches(0.55), Inches(y), Inches(6.4), Inches(1.48), fill=WHITE, line=LINE)
        textbox(slide, Inches(0.75), Inches(y + 0.12), Inches(5.9), Inches(0.22),
                [{"text": label, "size": 9, "bold": True, "color": MUTED}])
        textbox(slide, Inches(0.75), Inches(y + 0.36), Inches(5.9), Inches(0.4),
                [{"text": value, "size": 26, "bold": True, "color": NAVY, "font": FONT_DISPLAY}])
        textbox(slide, Inches(0.75), Inches(y + 0.82), Inches(5.9), Inches(0.35),
                [{"text": desc, "size": 11, "color": BODY}])
        bar_w = Inches(5.9 * width_frac)
        rect(slide, Inches(0.75), Inches(y + 1.22), bar_w, Inches(0.14), fill=color)
        y += 1.58

    round_rect(slide, Inches(7.15), Inches(1.32), Inches(5.6), Inches(4.72), fill=WHITE, line=LINE)
    textbox(slide, Inches(7.35), Inches(1.45), Inches(5.2), Inches(0.3),
            [{"text": "Global AI career-tech spend (illustrative, US$B)", "size": 12, "bold": True, "color": NAVY}])
    add_chart(
        slide,
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(7.25), Inches(1.85), Inches(5.35), Inches(3.95),
        ["2024", "2025E", "2026E", "2027E", "2028E"],
        [("Spend", (18, 23, 29, 37, 46))],
        legend=False,
    )
    textbox(slide, Inches(0.55), Inches(6.35), Inches(12), Inches(0.3),
            [{"text": "Management estimates informed by public global industry research; presented for illustration.", "size": 8, "color": MUTED}])
    footer(slide, 5)


def slide_platform(prs):
    slide = blank_slide(prs)
    section_header(slide, "Our Product", "What 51Careers.AI Delivers Today")

    # Intro strip
    round_rect(slide, Inches(0.55), Inches(1.28), Inches(12.25), Inches(0.7), fill=TEAL_SOFT)
    textbox(
        slide, Inches(0.75), Inches(1.4), Inches(11.85), Inches(0.5),
        [{"text": "51Careers.AI is a free AI career platform — mock interviews, a 100K+ question database, and a new AI résumé builder — designed to help professionals worldwide compete for better jobs.",
          "size": 13, "bold": True, "color": TEAL_DK}],
    )

    products = [
        ("AI Mock Interviews", "Role- and company-specific practice with personalized feedback powered by career-domain intelligence.", TEAL),
        ("100K+ Question Bank", "Proprietary library of real interview questions built from years of coaching and candidate preparation.", BLUE),
        ("AI Résumé Builder", "New AI feature that drafts and refines résumés tailored to roles, industries, and hiring expectations.", BRONZE),
        ("Career Pathways", "Coaching and education pathways ready to monetize via membership, consulting, and education fees.", SLATE),
    ]
    x = 0.55
    for title, desc, color in products:
        round_rect(slide, Inches(x), Inches(2.2), Inches(3.0), Inches(3.5), fill=WHITE, line=LINE)
        rect(slide, Inches(x), Inches(2.2), Inches(3.0), Inches(0.12), fill=color)
        textbox(slide, Inches(x + 0.18), Inches(2.5), Inches(2.65), Inches(0.7),
                [{"text": title, "size": 15, "bold": True, "color": NAVY, "font": FONT_DISPLAY}])
        textbox(slide, Inches(x + 0.18), Inches(3.3), Inches(2.65), Inches(2.0),
                [{"text": desc, "size": 12, "color": BODY}])
        x += 3.15

    round_rect(slide, Inches(0.55), Inches(5.9), Inches(12.25), Inches(0.7), fill=WHITE, line=LINE)
    textbox(
        slide, Inches(0.75), Inches(6.05), Inches(11.85), Inches(0.45),
        [{"text": "Pricing today: Free  ·  Next: membership tiers · consulting fees · education fees as active users scale globally",
          "size": 13, "bold": True, "color": NAVY}],
    )
    footer(slide, 6)


def slide_tech(prs):
    slide = blank_slide(prs)
    section_header(slide, "Technology", "Enterprise-Grade AI, Built with Alibaba Cloud")

    round_rect(slide, Inches(0.55), Inches(1.28), Inches(7.4), Inches(2.7), fill=WHITE, line=LINE)
    textbox(slide, Inches(0.8), Inches(1.4), Inches(7.0), Inches(0.32),
            [{"text": "Alibaba Cloud AI partnership — July 2025", "size": 14, "bold": True, "color": TEAL}])
    textbox(slide, Inches(0.8), Inches(1.8), Inches(6.9), Inches(0.95),
            [{"text": "A deep collaboration integrating frontier AI across the platform — Tongyi LLM, NLP, inference, recommendation, and semantic search — on enterprise-grade compute and security infrastructure.",
              "size": 12, "color": BODY}])
    bullets = [
        "AI mock interviews and career advisor with personalized recommendations",
        "AI résumé builder tuned for real hiring workflows",
        "Enterprise-grade data security protecting user privacy",
    ]
    y = 2.85
    for b in bullets:
        textbox(slide, Inches(0.8), Inches(y), Inches(6.9), Inches(0.3),
                [{"text": "▸  " + b, "size": 11, "color": INK}])
        y += 0.32

    round_rect(slide, Inches(8.15), Inches(1.28), Inches(4.65), Inches(2.7), fill=NAVY)
    textbox(slide, Inches(8.4), Inches(1.5), Inches(4.2), Inches(0.35),
            [{"text": "Proprietary data moat", "size": 15, "bold": True, "color": WHITE, "font": FONT_DISPLAY}])
    textbox(slide, Inches(8.4), Inches(2.05), Inches(4.2), Inches(1.6),
            [{"text": "100K+ proprietary interview questions and eight years of candidate-prep insight — assets that general-purpose models like ChatGPT or Gemini cannot readily replicate.",
              "size": 12, "color": RGBColor(0xC5, 0xD5, 0xE0)}])

    textbox(slide, Inches(0.55), Inches(4.15), Inches(8), Inches(0.28),
            [{"text": "Platform architecture (simplified)", "size": 12, "bold": True, "color": NAVY}])
    layers = [
        ("APPLICATIONS", "Mock Interviews · AI Résumé Builder · Career Advisor · Question Bank", TEAL),
        ("INTELLIGENCE", "Tongyi LLM · NLP · Semantic Search · Recommendation", BLUE),
        ("DATA & INFRASTRUCTURE", "100K+ question corpus · 8-year coaching data · Alibaba Cloud", BRONZE),
    ]
    y = 4.5
    for label, detail, color in layers:
        round_rect(slide, Inches(0.55), Inches(y), Inches(12.25), Inches(0.58), fill=WHITE, line=LINE)
        rect(slide, Inches(0.55), Inches(y), Inches(0.12), Inches(0.58), fill=color)
        textbox(slide, Inches(0.9), Inches(y + 0.14), Inches(3.5), Inches(0.3),
                [{"text": label, "size": 11, "bold": True, "color": color}])
        textbox(slide, Inches(4.5), Inches(y + 0.14), Inches(8.0), Inches(0.3),
                [{"text": detail, "size": 12, "color": BODY}])
        y += 0.65
    footer(slide, 7)


def slide_model(prs):
    slide = blank_slide(prs)
    section_header(slide, "Business Model", "Free Today — Paid as Users Scale")

    pillars = [
        ("01 · TODAY", "Free Platform",
         "51Careers.AI is free to maximize active users: AI mock interviews, the 100K+ question bank, and the new AI résumé builder.",
         "Acquire · Engage · Build habit", TEAL),
        ("02 · NEXT", "Membership Fees",
         "Tiered memberships unlock advanced practice, premium résumé tools, deeper question packs, and progress analytics.",
         "Recurring · Scalable · Global", BLUE),
        ("03 · EXPAND", "Consulting & Education",
         "Higher-ARPU consulting, coaching, and education programs for users who want human guidance and structured career curricula.",
         "High ARPU · Proven demand", BRONZE),
    ]
    x = 0.55
    for eyebrow, title, desc, tag, color in pillars:
        round_rect(slide, Inches(x), Inches(1.35), Inches(3.95), Inches(2.9), fill=WHITE, line=LINE)
        rect(slide, Inches(x), Inches(1.35), Inches(3.95), Inches(0.12), fill=color)
        textbox(slide, Inches(x + 0.2), Inches(1.6), Inches(3.55), Inches(0.3),
                [{"text": eyebrow, "size": 10, "bold": True, "color": color}])
        textbox(slide, Inches(x + 0.2), Inches(1.95), Inches(3.55), Inches(0.4),
                [{"text": title, "size": 16, "bold": True, "color": NAVY, "font": FONT_DISPLAY}])
        textbox(slide, Inches(x + 0.2), Inches(2.45), Inches(3.55), Inches(1.1),
                [{"text": desc, "size": 11, "color": BODY}])
        textbox(slide, Inches(x + 0.2), Inches(3.7), Inches(3.55), Inches(0.35),
                [{"text": tag, "size": 10, "bold": True, "color": MUTED}])
        x += 4.15

    round_rect(slide, Inches(0.55), Inches(4.45), Inches(12.25), Inches(2.15), fill=WHITE, line=LINE)
    textbox(slide, Inches(0.75), Inches(4.55), Inches(7.5), Inches(0.3),
            [{"text": "Illustrative monetization mix once paid conversion begins (%)", "size": 12, "bold": True, "color": NAVY}])
    add_chart(
        slide,
        XL_CHART_TYPE.COLUMN_STACKED_100,
        Inches(0.7), Inches(4.85), Inches(7.6), Inches(1.6),
        ["2027E", "2029E"],
        [
            ("Membership fees", (55, 50)),
            ("Consulting fees", (30, 30)),
            ("Education fees", (15, 20)),
        ],
        legend=True,
    )
    textbox(slide, Inches(8.6), Inches(5.05), Inches(3.9), Inches(1.3),
            [{"text": "Path to revenue", "size": 13, "bold": True, "color": NAVY, "space_after": 8},
             {"text": "Platform revenue is US$0 today. Grow active users on free access, then introduce membership, consulting, and education fees.",
              "size": 12, "color": BODY}])
    footer(slide, 8)


def slide_traction(prs):
    slide = blank_slide(prs)
    section_header(slide, "Traction", "Built on Real Candidate Outcomes")

    kpis = [
        ("10,000+", "CANDIDATES HELPED · 8 YEARS"),
        ("100K+", "MOCK INTERVIEW QUESTIONS"),
        ("2025", "AI PLATFORM LAUNCH"),
        ("US$0", "PLATFORM REVENUE (FREE)"),
    ]
    x = 0.55
    for val, label in kpis:
        round_rect(slide, Inches(x), Inches(1.28), Inches(2.95), Inches(1.15), fill=WHITE, line=LINE)
        textbox(slide, Inches(x + 0.15), Inches(1.38), Inches(2.65), Inches(0.5),
                [{"text": val, "size": 24, "bold": True, "color": NAVY, "font": FONT_DISPLAY}])
        textbox(slide, Inches(x + 0.15), Inches(1.95), Inches(2.65), Inches(0.35),
                [{"text": label, "size": 9, "bold": True, "color": MUTED}])
        x += 3.15

    # Equal-height panels
    panel_top, panel_h = 2.6, 3.55
    round_rect(slide, Inches(0.55), Inches(panel_top), Inches(7.5), Inches(panel_h), fill=WHITE, line=LINE)
    textbox(slide, Inches(0.75), Inches(panel_top + 0.12), Inches(7), Inches(0.28),
            [{"text": "Selected milestones", "size": 13, "bold": True, "color": NAVY}])
    rect(slide, Inches(1.15), Inches(panel_top + 0.6), Inches(0.04), Inches(2.7), fill=TEAL)
    milestones = [
        ("2016", "51Careers founded — premium career development for global professionals"),
        ("2016–24", "Helped 10,000+ candidates; built deep interview and résumé expertise"),
        ("2025", "Launched 51Careers.AI — free AI platform with mock interviews and 100K+ questions"),
        ("2025–26", "Developing AI résumé builder; growing active users globally"),
        ("Next", "Introduce membership, consulting, and education fees as usage scales"),
    ]
    y = panel_top + 0.55
    for date, desc in milestones:
        oval = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.05), Inches(y + 0.06), Inches(0.24), Inches(0.24))
        fill_solid(oval, TEAL)
        textbox(slide, Inches(1.5), Inches(y), Inches(1.15), Inches(0.28),
                [{"text": date, "size": 11, "bold": True, "color": TEAL}])
        textbox(slide, Inches(2.7), Inches(y), Inches(5.1), Inches(0.5),
                [{"text": desc, "size": 11, "color": BODY}])
        y += 0.55

    round_rect(slide, Inches(8.25), Inches(panel_top), Inches(4.55), Inches(panel_h), fill=WHITE, line=LINE)
    textbox(slide, Inches(8.45), Inches(panel_top + 0.12), Inches(4.2), Inches(0.28),
            [{"text": "Why it matters", "size": 13, "bold": True, "color": NAVY}])
    reasons = [
        "Eight years of real candidate-prep insight behind the AI product",
        "100K+ proprietary questions create a durable advantage vs. generic LLMs",
        "Free launch maximizes global user acquisition before monetization",
        "Clear path: membership · consulting · education fees",
    ]
    y = panel_top + 0.55
    for r in reasons:
        textbox(slide, Inches(8.45), Inches(y), Inches(4.15), Inches(0.65),
                [{"text": "▸  " + r, "size": 12, "color": BODY}])
        y += 0.7

    textbox(slide, Inches(0.55), Inches(6.3), Inches(12), Inches(0.28),
            [{"text": "Platform revenue is currently US$0 by design (free access). Candidate totals reflect 51Careers operating history since 2016.",
              "size": 8, "color": MUTED}])
    footer(slide, 9)


def slide_competition(prs):
    slide = blank_slide(prs)
    section_header(slide, "Competitive Positioning", "Purpose-Built Career AI vs. General LLMs")

    # Score chart vs general AI
    round_rect(slide, Inches(0.55), Inches(1.3), Inches(5.7), Inches(3.55), fill=WHITE, line=LINE)
    textbox(slide, Inches(0.75), Inches(1.4), Inches(5.3), Inches(0.3),
            [{"text": "Career-domain capability score (illustrative, 0–6)", "size": 12, "bold": True, "color": NAVY}])
    add_chart(
        slide,
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(0.65), Inches(1.75), Inches(5.45), Inches(2.9),
        ["51Careers.AI", "ChatGPT / OpenAI", "Google Gemini", "Claude", "Generic prep apps"],
        [("Score", (6, 2, 2, 2, 3))],
        legend=False,
        value_max=6,
    )

    headers = ["Capability", "51Careers.AI", "ChatGPT", "Gemini", "Claude"]
    rows = [
        ["Career-specific mock interviews", "●", "○", "○", "○"],
        ["100K+ proprietary question bank", "●", "—", "—", "—"],
        ["AI résumé builder (hiring-tuned)", "●", "○", "○", "○"],
        ["8-year coaching / candidate data", "●", "—", "—", "—"],
        ["Membership / consulting path", "●", "—", "—", "—"],
        ["General conversational AI", "○", "●", "●", "●"],
    ]
    table_shape = slide.shapes.add_table(7, 5, Inches(6.45), Inches(1.3), Inches(6.35), Inches(3.55))
    table = table_shape.table
    for c, h in enumerate(headers):
        cell = table.cell(0, c)
        cell.text = h
        for p in cell.text_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER if c else PP_ALIGN.LEFT
            for run in p.runs:
                run.font.name = FONT_BODY
                run.font.size = Pt(9)
                run.font.bold = True
                run.font.color.rgb = WHITE
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
    for r, row in enumerate(rows, 1):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            cell.text = val
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER if c else PP_ALIGN.LEFT
                for run in p.runs:
                    run.font.name = FONT_BODY
                    run.font.size = Pt(10)
                    run.font.bold = (c == 1 and val == "●")
                    run.font.color.rgb = TEAL if val == "●" else (BODY if c == 0 else MUTED)
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if r % 2 else SOFT

    textbox(slide, Inches(0.55), Inches(5.05), Inches(12.2), Inches(0.3),
            [{"text": "●  Full / native capability      ○  Partial / generic      —  Not offered", "size": 10, "color": MUTED}])
    round_rect(slide, Inches(0.55), Inches(5.4), Inches(12.25), Inches(0.95), fill=TEAL_SOFT)
    textbox(slide, Inches(0.75), Inches(5.55), Inches(11.85), Inches(0.7),
            [{"text": "General models are powerful — but not career platforms. 51Careers.AI combines domain AI, a proprietary 100K+ question corpus, and an AI résumé builder purpose-built for global hiring outcomes.",
              "size": 13, "bold": True, "color": TEAL_DK}])
    textbox(slide, Inches(0.55), Inches(6.4), Inches(12), Inches(0.25),
            [{"text": "Illustrative comparison based on management assessment of publicly available general-purpose AI products.", "size": 8, "color": MUTED}])
    footer(slide, 10)


def slide_growth(prs):
    slide = blank_slide(prs)
    section_header(slide, "Growth Strategy", "Users First, Then Monetize Globally")

    phases = [
        ("2026", "Grow free usage", TEAL, [
            "Scale active users on free mock interviews, question bank, and AI résumé builder",
            "Ship résumé-builder enhancements and expand question coverage globally",
            "Instrument conversion funnels for membership pilots",
        ]),
        ("2027", "Introduce paid tiers", BLUE, [
            "Launch membership fees for premium AI practice and résumé tools",
            "Add consulting and education fee packages for high-intent users",
            "Expand language and market coverage worldwide",
        ]),
        ("2028+", "Compound globally", BRONZE, [
            "Optimize paid conversion across membership, consulting, and education",
            "Deepen verticals and seniority segments globally",
            "Selective partnerships with universities and employers",
        ]),
    ]
    x = 0.55
    for year, title, color, bullets in phases:
        round_rect(slide, Inches(x), Inches(1.35), Inches(3.95), Inches(3.7), fill=WHITE, line=LINE)
        rect(slide, Inches(x), Inches(1.35), Inches(3.95), Inches(0.9), fill=color)
        textbox(slide, Inches(x + 0.25), Inches(1.45), Inches(3.45), Inches(0.4),
                [{"text": year, "size": 26, "bold": True, "color": WHITE, "font": FONT_DISPLAY}])
        textbox(slide, Inches(x + 0.25), Inches(1.9), Inches(3.45), Inches(0.3),
                [{"text": title, "size": 13, "bold": True, "color": WHITE}])
        y = 2.5
        for b in bullets:
            textbox(slide, Inches(x + 0.25), Inches(y), Inches(3.45), Inches(0.7),
                    [{"text": "▸  " + b, "size": 12, "color": BODY}])
            y += 0.75
        x += 4.15

    round_rect(slide, Inches(0.55), Inches(5.2), Inches(12.25), Inches(1.2), fill=WHITE, line=LINE)
    textbox(slide, Inches(0.75), Inches(5.28), Inches(11.8), Inches(0.25),
            [{"text": "Illustrative active-user growth (index)", "size": 11, "bold": True, "color": NAVY}])
    add_chart(
        slide,
        XL_CHART_TYPE.LINE_MARKERS,
        Inches(0.7), Inches(5.4), Inches(11.9), Inches(0.95),
        ["2025", "2026", "2027", "2028", "2029"],
        [("Index", (10, 35, 80, 140, 220))],
        legend=False,
    )
    textbox(slide, Inches(0.55), Inches(6.45), Inches(12), Inches(0.25),
            [{"text": "Strategy: win global users with a free career AI product, then monetize engagement through membership, consulting, and education fees.",
              "size": 11, "color": MUTED}])
    footer(slide, 11)


def slide_financials(prs):
    slide = blank_slide(prs)
    section_header(slide, "Financial Outlook", "From Free Platform to Paid Scale")

    round_rect(slide, Inches(0.55), Inches(1.3), Inches(7.9), Inches(3.15), fill=WHITE, line=LINE)
    textbox(slide, Inches(0.75), Inches(1.4), Inches(6), Inches(0.28),
            [{"text": "Illustrative platform revenue, US$ millions", "size": 12, "bold": True, "color": NAVY}])
    add_chart(
        slide,
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(0.65), Inches(1.65), Inches(7.65), Inches(2.65),
        ["2025A", "2026E", "2027E", "2028E", "2029E"],
        [("Revenue", (0, 0.8, 3.5, 9.0, 18.0))],
        legend=False,
    )

    round_rect(slide, Inches(0.55), Inches(4.6), Inches(7.9), Inches(1.7), fill=WHITE, line=LINE)
    textbox(slide, Inches(0.75), Inches(4.68), Inches(6), Inches(0.25),
            [{"text": "Illustrative paid conversion of active users (%)", "size": 11, "bold": True, "color": NAVY}])
    add_chart(
        slide,
        XL_CHART_TYPE.LINE_MARKERS,
        Inches(0.65), Inches(4.85), Inches(7.65), Inches(1.4),
        ["2025A", "2026E", "2027E", "2028E", "2029E"],
        [("Conversion", (0, 2, 5, 8, 12))],
        legend=False,
    )

    round_rect(slide, Inches(8.65), Inches(1.3), Inches(4.15), Inches(1.55), fill=NAVY)
    textbox(slide, Inches(8.85), Inches(1.45), Inches(3.75), Inches(0.55),
            [{"text": "US$0", "size": 34, "bold": True, "color": WHITE, "font": FONT_DISPLAY}])
    textbox(slide, Inches(8.85), Inches(2.15), Inches(3.75), Inches(0.4),
            [{"text": "PLATFORM REVENUE TODAY (FREE)", "size": 10, "bold": True, "color": TEAL_SOFT}])

    round_rect(slide, Inches(8.65), Inches(3.05), Inches(4.15), Inches(3.25), fill=WHITE, line=LINE)
    textbox(slide, Inches(8.85), Inches(3.2), Inches(3.75), Inches(0.3),
            [{"text": "Revenue drivers", "size": 13, "bold": True, "color": NAVY}])
    drivers = [
        "Grow free active users globally",
        "Introduce membership fees for premium AI tools",
        "Add consulting fees for high-touch guidance",
        "Layer education fees for structured programs",
    ]
    y = 3.6
    for d in drivers:
        textbox(slide, Inches(8.85), Inches(y), Inches(3.75), Inches(0.55),
                [{"text": "▸  " + d, "size": 11, "color": BODY}])
        y += 0.55

    textbox(slide, Inches(0.55), Inches(6.4), Inches(12.2), Inches(0.3),
            [{"text": "Illustrative management projections only — not a forecast. 2025A platform revenue is US$0 (free product). Actual results will differ.",
              "size": 8, "color": MUTED}])
    footer(slide, 12)


def slide_offering(prs):
    slide = blank_slide(prs)
    section_header(slide, "The Offering", "US$5.0 Million to Accelerate Global Leadership")

    # Terms panel
    round_rect(slide, Inches(0.55), Inches(1.3), Inches(5.9), Inches(4.55), fill=WHITE, line=LINE)
    textbox(slide, Inches(0.8), Inches(1.5), Inches(5.4), Inches(0.3),
            [{"text": "INDICATIVE TERMS", "size": 11, "bold": True, "color": TEAL}])
    terms = [
        ("Round size", "US$5.0 million (primary)"),
        ("Valuation", "US$50.0 million (pre-money)"),
        ("Structure", "Preferred equity"),
        ("Use of funds", "Product & AI, expansion, team"),
        ("Target close", "2H 2026"),
    ]
    y = 2.0
    for label, value in terms:
        rect(slide, Inches(0.8), Inches(y + 0.55), Inches(5.4), Pt(1), fill=LINE)
        textbox(slide, Inches(0.8), Inches(y), Inches(2.4), Inches(0.4),
                [{"text": label, "size": 12, "color": MUTED}])
        textbox(slide, Inches(3.2), Inches(y), Inches(3.0), Inches(0.4),
                [{"text": value, "size": 13, "bold": True, "color": NAVY}])
        y += 0.65
    textbox(slide, Inches(0.8), Inches(5.4), Inches(5.4), Inches(0.3),
            [{"text": "Indicative only; subject to definitive documentation.", "size": 9, "italic": True, "color": MUTED}])

    # Doughnut
    round_rect(slide, Inches(6.7), Inches(1.3), Inches(6.1), Inches(4.55), fill=WHITE, line=LINE)
    textbox(slide, Inches(6.95), Inches(1.5), Inches(5.6), Inches(0.3),
            [{"text": "Use of proceeds", "size": 13, "bold": True, "color": NAVY}])
    add_chart(
        slide,
        XL_CHART_TYPE.DOUGHNUT,
        Inches(6.9), Inches(1.85), Inches(5.7), Inches(3.5),
        [
            "Product & AI development (35%)",
            "International expansion (30%)",
            "Talent & team (20%)",
            "Brand & growth marketing (10%)",
            "Working capital (5%)",
        ],
        [("Allocation", (35, 30, 20, 10, 5))],
        legend=True,
        point_colors=CHART_COLORS,
    )

    round_rect(slide, Inches(0.55), Inches(6.0), Inches(12.25), Inches(0.55), fill=TEAL_SOFT)
    textbox(slide, Inches(0.75), Inches(6.1), Inches(11.85), Inches(0.4),
            [{"text": "What this capital achieves — AI résumé builder & mock-interview scale · global active-user growth · paid membership / consulting / education launch.",
              "size": 12, "bold": True, "color": TEAL_DK}])
    footer(slide, 13)


def slide_leadership(prs):
    """Executive leadership — six condensed investor-ready profiles on one slide."""
    slide = blank_slide(prs)
    section_header(slide, "Leadership", "Executive Leadership Team")

    leaders = [
        (
            "Rocky Chen",
            "FOUNDER & CEO",
            [
                "Founder of 51Careers and 51Careers.AI; investor & Managing Director at publicly listed Helio",
                "Decade of international business, capital markets, and global resource integration",
            ],
        ),
        (
            "Stephanie Li",
            "CO-FOUNDER & CFO",
            [
                "Co-Founder since 2016; U.S. CPA — M.S. Accounting (Pace University)",
                "Oversees financial strategy, corporate governance, and sustainable growth",
            ],
        ),
        (
            "Gavin Ding",
            "CO-FOUNDER, CTO & COO",
            [
                "Serial entrepreneur (10+ yrs) across SaaS, AI, and digital business; B.S. CS, ECUST",
                "Prior CTO roles and multiple co-founded ventures; leads technology, operations, and AI",
            ],
        ),
        (
            "Robin Zhu",
            "HEAD OF PRODUCT",
            [
                "Leads product strategy, architecture, and development of the AI-powered platform",
                "Former CTO & Director of Product, Ci Finance; senior roles at CPIC, Allinpay, Noah",
            ],
        ),
        (
            "Chris Lin",
            "NORTH AMERICAN PARTNER",
            [
                "Full-stack engineer; previously at Amazon Web Services (AWS)",
                "B.S. Northwestern, M.S. Robotics; extensive hiring-panel experience and mentorship",
            ],
        ),
        (
            "Jon Serbin",
            "SENIOR ADVISOR",
            [
                "Harvard & MIT; former senior executive at Morgan Stanley; founder of Cedar",
                "40+ years in tech M&A and capital raising; advises on growth and expansion",
            ],
        ),
    ]

    # 2 × 3 grid
    card_w, card_h = Inches(4.0), Inches(2.35)
    gap_x, gap_y = Inches(0.2), Inches(0.18)
    origin_x, origin_y = Inches(0.55), Inches(1.28)
    accents = [TEAL, BLUE, BRONZE, TEAL, BLUE, BRONZE]

    for i, (name, title, bullets) in enumerate(leaders):
        col, row = i % 3, i // 3
        x = origin_x + col * (card_w + gap_x)
        y = origin_y + row * (card_h + gap_y)
        round_rect(slide, x, y, card_w, card_h, fill=WHITE, line=LINE)
        accent_bar(slide, x, y, w=Inches(0.08), h=card_h, color=accents[i])
        textbox(
            slide,
            x + Inches(0.22),
            y + Inches(0.14),
            card_w - Inches(0.35),
            Inches(0.35),
            [{"text": name, "size": 15, "bold": True, "color": NAVY, "font": FONT_DISPLAY}],
        )
        textbox(
            slide,
            x + Inches(0.22),
            y + Inches(0.48),
            card_w - Inches(0.35),
            Inches(0.28),
            [{"text": title, "size": 10, "bold": True, "color": TEAL}],
        )
        by = y + Inches(0.85)
        for bullet in bullets:
            textbox(
                slide,
                x + Inches(0.22),
                by,
                card_w - Inches(0.38),
                Inches(0.65),
                [{"text": "▸  " + bullet, "size": 10, "color": BODY}],
            )
            by += Inches(0.68)

    textbox(
        slide,
        Inches(0.55),
        Inches(6.3),
        Inches(12.2),
        Inches(0.3),
        [{
            "text": "Supported by ≈50 professionals across New York, Shanghai, London, Singapore, Seoul, Toronto, and Delhi. Full biographies available in the data room.",
            "size": 9,
            "color": MUTED,
        }],
    )
    footer(slide, 14)


def slide_close(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fill_solid(slide.background, NAVY)
    rect(slide, Inches(0), Inches(0), SW, Inches(0.12), fill=TEAL)
    rect(slide, Inches(0), Inches(0), Inches(0.18), SH, fill=TEAL)

    textbox(
        slide,
        Inches(1.0),
        Inches(1.8),
        Inches(11.3),
        Inches(1.5),
        [{
            "text": "“Our mission is to make great jobs accessible to everyone.”",
            "size": 28,
            "bold": True,
            "color": WHITE,
            "font": FONT_DISPLAY,
            "align": PP_ALIGN.CENTER,
        }],
        valign=MSO_ANCHOR.MIDDLE,
    )
    textbox(
        slide,
        Inches(1.0),
        Inches(3.5),
        Inches(11.3),
        Inches(0.5),
        [{"text": "51CAREERS.AI", "size": 22, "bold": True, "color": TEAL, "font": FONT_DISPLAY, "align": PP_ALIGN.CENTER}],
    )
    textbox(
        slide,
        Inches(1.0),
        Inches(4.2),
        Inches(11.3),
        Inches(0.35),
        [{"text": "Rocky Chen  ·  Founder & Chief Executive Officer", "size": 13, "color": RGBColor(0xC5, 0xD5, 0xE0), "align": PP_ALIGN.CENTER}],
    )
    textbox(
        slide,
        Inches(1.0),
        Inches(5.0),
        Inches(11.3),
        Inches(0.35),
        [{"text": "info@51careers.com      ·      www.51careers.ai", "size": 12, "color": WHITE, "align": PP_ALIGN.CENTER}],
    )
    textbox(
        slide,
        Inches(1.0),
        Inches(5.4),
        Inches(11.3),
        Inches(0.35),
        [{"text": "48 Wall Street, 11th Floor, New York, NY 10005", "size": 12, "color": SLATE, "align": PP_ALIGN.CENTER}],
    )
    textbox(
        slide,
        Inches(1.0),
        SH - Inches(0.7),
        Inches(11.3),
        Inches(0.3),
        [{"text": "51CAREERS.AI   ·   CONFIDENTIAL   ·   15", "size": 9, "bold": True, "color": SLATE, "align": PP_ALIGN.CENTER}],
    )


def build():
    prs = Presentation()
    prs.slide_width = SW
    prs.slide_height = SH

    slide_cover(prs)
    slide_disclaimer(prs)
    slide_highlights(prs)
    slide_problem(prs)
    slide_market(prs)
    slide_platform(prs)
    slide_tech(prs)
    slide_model(prs)
    slide_traction(prs)
    slide_competition(prs)
    slide_growth(prs)
    slide_financials(prs)
    slide_offering(prs)
    slide_leadership(prs)
    slide_close(prs)

    prs.save(OUT)
    print(f"Wrote {OUT} ({OUT.stat().st_size:,} bytes) with {len(prs.slides)} slides")
    # count charts
    charts = 0
    for s in prs.slides:
        for sh in s.shapes:
            if sh.has_chart:
                charts += 1
    print(f"Native charts: {charts}")


if __name__ == "__main__":
    build()
