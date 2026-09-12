#!/usr/bin/env python3
"""Build 51Careers_Addition.pptx — full investor narrative.

Cover + sections 01–24 + three institutional slides (TAM/SAM/SOM,
3-year model, RMB 20M → investor outcomes). Teal / Georgia visual system.
Financing: RMB 20M at RMB 180M pre / RMB 200M post (10% equity).
"""

from __future__ import annotations

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

OUT = "/workspace/51Careers_Addition.pptx"

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
COVER = RGBColor(0x08, 0x18, 0x28)
BLUE = RGBColor(0x2F, 0x5D, 0x8A)
BRONZE = RGBColor(0x8A, 0x6D, 0x3B)
FONT_D = "Georgia"
FONT_B = "Segoe UI"

SW, SH = Inches(13.333), Inches(7.5)
ML = Inches(0.55)
CW = Inches(12.23)
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
FX = "USD at an illustrative RMB 6.73 / US$1. Figures labeled illustrative are management projections, not forecasts."


def apply_typeface(run, name):
    """Force Latin / EA / CS typeface so digits do not fall back to Calibri."""
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", name)


def set_run(run, text, font=FONT_B, size=11, bold=False, color=BODY, italic=False):
    run.text = text
    apply_typeface(run, font)
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
            {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}.get(valign, "t"),
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
            run, spec["text"], font=spec.get("font", FONT_B), size=spec.get("size", 11),
            bold=spec.get("bold", False), color=spec.get("color", BODY), italic=spec.get("italic", False),
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
        s.adjustments[0] = 0.06
    except Exception:
        pass
    return s


def textbox(slide, l, t, w, h, paragraphs, valign=MSO_ANCHOR.TOP):
    s = slide.shapes.add_textbox(l, t, w, h)
    add_text(s, paragraphs, valign=valign)
    return s


def set_bg(slide, hex_color="F5F8FB"):
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
    cSld.find(qn("p:spTree")).addprevious(bg)


def chrome(slide, page: int):
    rect(slide, Inches(0), Inches(0), SW, Inches(0.08), fill=TEAL)
    rect(slide, Inches(0), Inches(7.08), SW, Inches(0.42), fill=PANEL)
    rect(slide, Inches(0), Inches(7.08), SW, Inches(0.01), fill=LINE)
    textbox(slide, ML, Inches(7.14), Inches(8.5), Inches(0.28),
            [{"text": "51 CAREERS   ·   CONFIDENTIAL INVESTOR PRESENTATION", "size": 8, "bold": True, "color": MUTED}])
    textbox(slide, Inches(12.20), Inches(7.14), Inches(0.70), Inches(0.28),
            [{"text": str(page), "size": 9, "color": MUTED, "align": PP_ALIGN.RIGHT}])


def header(slide, eyebrow: str, title: str, subtitle: str | None = None):
    textbox(slide, ML, Inches(0.20), CW, Inches(0.22),
            [{"text": eyebrow.upper(), "size": 10, "bold": True, "color": TEAL}])
    title_size = 20 if len(title) > 52 else 24
    textbox(slide, ML, Inches(0.40), CW, Inches(0.50),
            [{"text": title, "size": title_size, "bold": True, "color": NAVY, "font": FONT_D}])
    if subtitle:
        textbox(slide, ML, Inches(0.90), CW, Inches(0.26),
                [{"text": subtitle, "size": 13, "color": MUTED}])


def blank(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide)
    return slide


def accent_card(slide, l, t, w, h, fill=WHITE, bar=TEAL):
    round_rect(slide, l, t, w, h, fill=fill)
    rect(slide, l, t, Inches(0.08), h, fill=bar)
    return l + Inches(0.22), t + Inches(0.14), w - Inches(0.36)


def style_cell(cell, text, size=10, bold=False, color=BODY, fill=None, align=PP_ALIGN.LEFT):
    cell.text = text
    tf = cell.text_frame
    tf.word_wrap = True
    for p in tf.paragraphs:
        p.alignment = align
        for run in p.runs:
            set_run(run, run.text, size=size, bold=bold, color=color)
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcPr.set("anchor", "ctr")
    for child in list(tcPr):
        if child.tag == f"{{{NS_A}}}solidFill":
            tcPr.remove(child)
    if fill:
        solid = etree.SubElement(tcPr, f"{{{NS_A}}}solidFill")
        srgb = etree.SubElement(solid, f"{{{NS_A}}}srgbClr")
        srgb.set("val", fill)


# ---------------------------------------------------------------------------
# Slides
# ---------------------------------------------------------------------------

def s_cover(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, "0B1F33")
    rect(s, Inches(0), Inches(0), SW, Inches(0.12), fill=TEAL)
    rect(s, Inches(0), Inches(0), Inches(0.16), SH, fill=TEAL)
    rect(s, Inches(0), Inches(6.55), SW, Inches(0.95), fill=COVER)
    textbox(s, Inches(0.70), Inches(0.42), Inches(12), Inches(0.28),
            [{"text": "CONFIDENTIAL INVESTOR PRESENTATION", "size": 11, "bold": True, "color": TEAL_SOFT}])
    textbox(s, Inches(0.70), Inches(0.90), Inches(12), Inches(0.70),
            [{"text": "51 CAREERS", "size": 42, "bold": True, "color": WHITE, "font": FONT_D}])
    textbox(s, Inches(0.70), Inches(1.58), Inches(12), Inches(0.40),
            [{"text": "Building the Global AI Career Platform", "size": 20, "color": TEAL_SOFT}])
    textbox(s, Inches(0.70), Inches(2.10), Inches(12), Inches(0.32),
            [{"text": "AI Resume  →  AI Career  →  Global Talent Ecosystem", "size": 15, "bold": True, "color": TEAL}])
    textbox(s, Inches(0.70), Inches(2.55), Inches(12), Inches(0.36),
            [{"text": "Seeking RMB 20 Million in Growth Capital   ·   ≈ US$2.97 Million", "size": 16, "bold": True, "color": WHITE}])

    cards = [
        ("PRE-MONEY", "RMB 180M", "≈ US$26.7 million"),
        ("RAISE  ·  10% EQUITY", "RMB 20M", "≈ US$2.97 million"),
        ("POST-MONEY", "RMB 200M", "≈ US$29.7 million"),
    ]
    for i, (lab, big, usd) in enumerate(cards):
        x = Inches(0.70) + i * Inches(4.05)
        round_rect(s, x, Inches(3.05), Inches(3.85), Inches(2.15), fill=COVER)
        rect(s, x, Inches(3.05), Inches(3.85), Inches(0.08), fill=TEAL)
        textbox(s, x + Inches(0.22), Inches(3.28), Inches(3.45), Inches(0.28),
                [{"text": lab, "size": 12, "bold": True, "color": TEAL}])
        textbox(s, x + Inches(0.22), Inches(3.62), Inches(3.45), Inches(0.70),
                [{"text": big, "size": 32, "bold": True, "color": WHITE, "font": FONT_D}])
        textbox(s, x + Inches(0.22), Inches(4.40), Inches(3.45), Inches(0.50),
                [{"text": usd, "size": 15, "color": TEAL_SOFT}])

    textbox(s, Inches(0.70), Inches(5.38), Inches(12), Inches(0.85),
            [{"text": "The resume is the entry point. The ambition is the global AI career platform — connecting job seekers, employers, career services, education, and the talent ecosystem.",
              "size": 16, "color": RGBColor(0xC5, 0xD5, 0xE0)}])
    textbox(s, Inches(0.70), Inches(6.75), Inches(12), Inches(0.32),
            [{"text": "CHINA R&D   ·   U.S. COMMERCIALIZATION   ·   GLOBAL DISTRIBUTION", "size": 11, "bold": True, "color": RGBColor(0x9A, 0xB0, 0xC0)}])


def s_exec(prs, n):
    s = blank(prs)
    header(s, "01  ·  Executive Summary", "Building the AI-Powered Future of Career Services",
           "Career Services  →  AI Resume  →  AI Career Platform  →  Global Career Marketplace")
    round_rect(s, ML, Inches(1.28), Inches(5.70), Inches(5.55), fill=WHITE)
    rect(s, ML, Inches(1.28), Inches(0.08), Inches(5.55), fill=TEAL)
    textbox(s, Inches(0.85), Inches(1.48), Inches(5.15), Inches(5.15), [
        {"text": "Who we are", "size": 14, "bold": True, "color": NAVY, "space_after": 6},
        {"text": "51 Careers is building a global AI Career Platform to help job seekers create better resumes, discover opportunities, improve job-search outcomes, and access professional career services.", "size": 13, "color": BODY, "space_after": 12},
        {"text": "Entry point", "size": 14, "bold": True, "color": NAVY, "space_after": 6},
        {"text": "An AI-powered Resume Platform that turns unstructured career experience into professional, quantified, job-specific resumes.", "size": 13, "color": BODY, "space_after": 12},
        {"text": "This round", "size": 14, "bold": True, "color": NAVY, "space_after": 6},
        {"text": "RMB 20 million (10% of RMB 200 million post-money) to accelerate product development, user acquisition, and global commercialization. Pre-money RMB 180 million.", "size": 13, "color": BODY},
    ])
    highlights = [
        ("01", "AI Resume demo already developed"),
        ("02", "Initial testing completed with promising results"),
        ("03", "Stronger professional positioning & quantification vs. general-purpose models in internal tests"),
        ("04", "Product can serve users globally from day one"),
        ("05", "Existing career-services business: domain expertise + monetization"),
        ("06", "U.S. market as the first international growth beachhead"),
        ("07", "Multiple revenue streams beyond AI Resume subscriptions"),
        ("08", "RMB 20M to fund product, acquisition, and global go-to-market"),
    ]
    for i, (num, txt) in enumerate(highlights):
        col, row = i % 2, i // 2
        x = Inches(6.50) + col * Inches(3.15)
        y = Inches(1.28) + row * Inches(1.38)
        round_rect(s, x, y, Inches(3.00), Inches(1.26), fill=WHITE)
        textbox(s, x + Inches(0.14), y + Inches(0.10), Inches(2.72), Inches(0.28),
                [{"text": num, "size": 12, "bold": True, "color": TEAL, "font": FONT_D}])
        textbox(s, x + Inches(0.14), y + Inches(0.40), Inches(2.72), Inches(0.74),
                [{"text": txt, "size": 12, "color": BODY}])
    chrome(s, n)


def s_problem(prs, n):
    s = blank(prs)
    header(s, "02  ·  The Problem", "Job Seeking Is Still Highly Inefficient",
           "For millions of candidates the process remains fragmented, expensive, and low-signal.")
    items = [
        ("01", "Resume Quality", "Most candidates struggle to communicate experience in a professional, results-oriented way."),
        ("02", "Quantification", "People describe responsibilities rather than measurable achievements."),
        ("03", "Personalization", "One generic resume rarely works across jobs and employers."),
        ("04", "Job Matching", "Candidates apply without a clear view of actual fit."),
        ("05", "Career Guidance", "Professional advice remains expensive and consultant-dependent."),
    ]
    for i, (num, title, body) in enumerate(items):
        x = ML + i * Inches(2.46)
        round_rect(s, x, Inches(1.22), Inches(2.34), Inches(4.35), fill=WHITE)
        rect(s, x, Inches(1.22), Inches(2.34), Inches(0.10), fill=TEAL)
        textbox(s, x + Inches(0.14), Inches(1.48), Inches(2.06), Inches(0.50),
                [{"text": num, "size": 22, "bold": True, "color": TEAL, "font": FONT_D}])
        textbox(s, x + Inches(0.14), Inches(2.05), Inches(2.06), Inches(0.85),
                [{"text": title, "size": 16, "bold": True, "color": NAVY}])
        textbox(s, x + Inches(0.14), Inches(2.95), Inches(2.06), Inches(2.35),
                [{"text": body, "size": 14, "color": BODY}])
    round_rect(s, ML, Inches(5.70), CW, Inches(1.18), fill=NAVY)
    textbox(s, Inches(0.80), Inches(5.82), Inches(11.7), Inches(0.26),
            [{"text": "THE RESULT", "size": 12, "bold": True, "color": TEAL}])
    textbox(s, Inches(0.80), Inches(6.14), Inches(11.7), Inches(0.52),
            [{"text": "Low-quality applications  →  Low interview rates  →  Long job searches  →  High frustration",
              "size": 17, "bold": True, "color": WHITE}])
    chrome(s, n)


def s_opportunity(prs, n):
    s = blank(prs)
    header(s, "03  ·  The Opportunity", "AI Is Reshaping the Career Services Industry",
           "Generative AI has changed how people create, communicate, and consume professional information.")
    xl, xr, w, top, h = ML, Inches(6.90), Inches(5.90), Inches(1.22), Inches(4.35)
    round_rect(s, xl, top, w, h, fill=WHITE)
    rect(s, xl, top, w, Inches(0.10), fill=BRONZE)
    textbox(s, xl + Inches(0.28), top + Inches(0.28), w - Inches(0.50), Inches(0.32),
            [{"text": "TRADITIONAL MODEL  ·  HUMAN CONSULTANT", "size": 12, "bold": True, "color": BRONZE}])
    for i, line in enumerate(["Manual analysis", "Manual resume writing", "Manual job search", "High cost", "Limited scalability"]):
        textbox(s, xl + Inches(0.28), top + Inches(0.80) + i * Inches(0.64), w - Inches(0.50), Inches(0.56),
                [{"text": f"→   {line}", "size": 17, "color": BODY}])
    round_rect(s, xr, top, w, h, fill=WHITE)
    rect(s, xr, top, w, Inches(0.10), fill=TEAL)
    textbox(s, xr + Inches(0.28), top + Inches(0.28), w - Inches(0.50), Inches(0.32),
            [{"text": "AI-ENABLED MODEL  ·  CAREER PLATFORM", "size": 12, "bold": True, "color": TEAL}])
    for i, line in enumerate(["Automated career analysis", "AI resume generation", "Job matching", "Continuous optimization", "Global scalability"]):
        textbox(s, xr + Inches(0.28), top + Inches(0.80) + i * Inches(0.64), w - Inches(0.50), Inches(0.56),
                [{"text": f"→   {line}", "size": 17, "bold": True, "color": NAVY}])
    round_rect(s, ML, Inches(5.70), CW, Inches(1.18), fill=TEAL)
    textbox(s, Inches(0.80), Inches(5.84), Inches(11.7), Inches(0.26),
            [{"text": "OUR OPPORTUNITY", "size": 12, "bold": True, "color": WHITE}])
    textbox(s, Inches(0.80), Inches(6.16), Inches(11.7), Inches(0.52),
            [{"text": "Transform high-cost, labor-intensive career services into a scalable AI-powered platform.",
              "size": 18, "bold": True, "color": WHITE, "font": FONT_D}])
    chrome(s, n)


def s_why_now(prs, n):
    s = blank(prs)
    header(s, "04  ·  Why Now", "Three Structural Trends Are Converging",
           "AI adoption, data-driven hiring, and fragmented career services are moving at the same time.")
    trends = [
        ("01", "Generative AI has reached mass adoption",
         "Consumers are increasingly comfortable using AI for professional and personal tasks. That adoption is the distribution opening for an AI career product."),
        ("02", "Recruiting is becoming more data-driven",
         "Employers rely on structured information, keywords, measurable achievements, and automated screening. Candidates need a product that writes to that standard."),
        ("03", "Career services remain fragmented",
         "The industry is still dependent on manual services and fragmented providers. That gap is the opening for a scalable AI + human platform."),
    ]
    for i, (num, title, body) in enumerate(trends):
        x = ML + i * Inches(4.08)
        round_rect(s, x, Inches(1.22), Inches(3.92), Inches(4.35), fill=WHITE)
        rect(s, x, Inches(1.22), Inches(3.92), Inches(0.10), fill=TEAL)
        textbox(s, x + Inches(0.24), Inches(1.46), Inches(3.44), Inches(0.42),
                [{"text": num, "size": 20, "bold": True, "color": TEAL, "font": FONT_D}])
        textbox(s, x + Inches(0.24), Inches(1.92), Inches(3.44), Inches(1.10),
                [{"text": title, "size": 18, "bold": True, "color": NAVY, "font": FONT_D}])
        textbox(s, x + Inches(0.24), Inches(3.10), Inches(3.44), Inches(2.20),
                [{"text": body, "size": 15, "color": BODY}])
    round_rect(s, ML, Inches(5.70), CW, Inches(1.18), fill=NAVY)
    textbox(s, Inches(0.80), Inches(5.84), Inches(11.7), Inches(0.26),
            [{"text": "51 CAREERS SITS AT THE INTERSECTION", "size": 12, "bold": True, "color": TEAL}])
    textbox(s, Inches(0.80), Inches(6.16), Inches(11.7), Inches(0.52),
            [{"text": "AI   ×   Employment   ×   Career Services", "size": 22, "bold": True, "color": WHITE, "font": FONT_D}])
    chrome(s, n)


def s_tam(prs, n):
    s = blank(prs)
    header(s, "Market Size", "TAM  →  SAM  →  SOM  →  Three-Year Revenue",
           "From global job seekers to paid AI career users. Illustrative, for discussion.")
    cards = [
        ("01", "TAM", "US$ 200B+", "RMB 1.35T+",
         "Global recruitment, talent-acquisition, and career-development spend."),
        ("02", "SAM", "US$ 30B+", "RMB 200B+",
         "Digital / AI-enabled hiring tools, resume products, and online career preparation."),
        ("03", "SOM", "US$ 1.5B+", "RMB 10B+",
         "AI resume and career-assistant demand in initial markets: China, the U.S., and English-speaking Asia."),
    ]
    for i, (num, lab, usd, rmb, desc) in enumerate(cards):
        x = ML + i * Inches(4.08)
        round_rect(s, x, Inches(1.22), Inches(3.92), Inches(3.28), fill=WHITE)
        rect(s, x, Inches(1.22), Inches(3.92), Inches(0.10), fill=TEAL)
        textbox(s, x + Inches(0.22), Inches(1.40), Inches(3.48), Inches(0.28),
                [{"text": f"{num}   ·   {lab}", "size": 13, "bold": True, "color": TEAL, "font": FONT_D}])
        textbox(s, x + Inches(0.22), Inches(1.74), Inches(3.48), Inches(0.58),
                [{"text": usd, "size": 28, "bold": True, "color": NAVY, "font": FONT_D}])
        textbox(s, x + Inches(0.22), Inches(2.36), Inches(3.48), Inches(0.30),
                [{"text": rmb, "size": 15, "bold": True, "color": INK, "font": FONT_D}])
        textbox(s, x + Inches(0.22), Inches(2.74), Inches(3.48), Inches(1.50),
                [{"text": desc, "size": 14, "color": BODY}])
    round_rect(s, ML, Inches(4.62), CW, Inches(2.26), fill=WHITE)
    rect(s, ML, Inches(4.62), Inches(0.08), Inches(2.26), fill=TEAL)
    textbox(s, Inches(0.85), Inches(4.74), Inches(11.6), Inches(0.26),
            [{"text": "51 CAREERS  ·  2027E–2029E REVENUE", "size": 12, "bold": True, "color": TEAL}])
    kpis = [
        ("2027E", "US$ 2.0M", "≈ RMB 13.5M"),
        ("2028E", "US$ 8.0M", "≈ RMB 53.8M"),
        ("2029E", "US$ 28.0M", "≈ RMB 188.4M"),
    ]
    for i, (y, a, b) in enumerate(kpis):
        x = Inches(0.78) + i * Inches(4.08)
        textbox(s, x, Inches(5.10), Inches(3.70), Inches(0.24),
                [{"text": y, "size": 13, "bold": True, "color": TEAL, "font": FONT_D}])
        textbox(s, x, Inches(5.36), Inches(3.70), Inches(0.50),
                [{"text": a, "size": 26, "bold": True, "color": NAVY, "font": FONT_D}])
        textbox(s, x, Inches(5.90), Inches(3.70), Inches(0.28),
                [{"text": b, "size": 14, "color": MUTED, "font": FONT_D}])
    textbox(s, Inches(0.78), Inches(6.32), Inches(11.7), Inches(0.36),
            [{"text": "Logic: global job seekers  →  addressable AI Resume users  →  paid conversion  →  stacked revenue (AI + services + commission + B2B + ads).",
              "size": 12, "color": MUTED}])
    chrome(s, n)


def s_solution(prs, n):
    s = blank(prs)
    header(s, "05  ·  Our Solution", "51 Careers AI Career Platform",
           "Not another AI writing tool — an AI-powered career infrastructure layer.")
    steps = [
        "Understand professional experience",
        "Build a stronger professional profile",
        "Generate optimized resumes",
        "Match with relevant opportunities",
        "Prepare for interviews",
        "Make better career decisions",
        "Access professional services",
    ]
    for i, step in enumerate(steps):
        y = Inches(1.22) + i * Inches(0.78)
        round_rect(s, ML, y, Inches(7.60), Inches(0.72), fill=WHITE)
        rect(s, ML, y, Inches(0.08), Inches(0.72), fill=TEAL)
        textbox(s, Inches(0.85), y + Inches(0.16), Inches(0.50), Inches(0.42),
                [{"text": f"{i+1:02d}", "size": 16, "bold": True, "color": TEAL, "font": FONT_D}])
        textbox(s, Inches(1.45), y + Inches(0.16), Inches(6.40), Inches(0.42),
                [{"text": step, "size": 17, "bold": True, "color": NAVY}])
    round_rect(s, Inches(8.40), Inches(1.22), Inches(4.38), Inches(5.66), fill=NAVY)
    textbox(s, Inches(8.65), Inches(1.50), Inches(3.95), Inches(0.28),
            [{"text": "STARTING POINT", "size": 12, "bold": True, "color": TEAL}])
    textbox(s, Inches(8.65), Inches(1.90), Inches(3.95), Inches(0.90),
            [{"text": "AI Resume", "size": 32, "bold": True, "color": WHITE, "font": FONT_D}])
    textbox(s, Inches(8.65), Inches(2.90), Inches(3.95), Inches(3.50),
            [{"text": "The resume is the first touchpoint between the user and the 51 Careers ecosystem — the on-ramp to AI Career tools, services, matching, and the marketplace.\n\nOne product. One profile. Multiple monetization paths across the career journey.",
              "size": 16, "color": TEAL_SOFT}])
    chrome(s, n)


def s_product(prs, n):
    s = blank(prs)
    header(s, "06  ·  Product", "From Career Experience to Quantified Professional Value")
    round_rect(s, ML, Inches(1.18), Inches(4.05), Inches(5.62), fill=WHITE)
    rect(s, ML, Inches(1.18), Inches(4.05), Inches(0.08), fill=TEAL)
    textbox(s, Inches(0.80), Inches(1.36), Inches(3.55), Inches(0.30),
            [{"text": "USERS PROVIDE", "size": 12, "bold": True, "color": TEAL}])
    for i, item in enumerate(["Education", "Work experience", "Projects", "Skills", "Achievements", "Career goals", "Target job descriptions"]):
        textbox(s, Inches(0.80), Inches(1.78) + i * Inches(0.66), Inches(3.55), Inches(0.58),
                [{"text": f"▸  {item}", "size": 15, "color": NAVY}])
    round_rect(s, Inches(4.80), Inches(1.18), Inches(3.55), Inches(5.62), fill=NAVY)
    textbox(s, Inches(5.05), Inches(1.50), Inches(3.10), Inches(0.30),
            [{"text": "OUR AI GENERATES", "size": 12, "bold": True, "color": TEAL}])
    textbox(s, Inches(5.05), Inches(2.00), Inches(3.10), Inches(3.10), [
        {"text": "Professional", "size": 24, "bold": True, "color": WHITE, "font": FONT_D, "space_after": 10},
        {"text": "+  Quantified", "size": 24, "bold": True, "color": WHITE, "font": FONT_D, "space_after": 10},
        {"text": "+  Job-specific", "size": 24, "bold": True, "color": WHITE, "font": FONT_D, "space_after": 10},
        {"text": "resume content", "size": 16, "color": TEAL_SOFT},
    ])
    textbox(s, Inches(5.05), Inches(5.20), Inches(3.10), Inches(1.35),
            [{"text": "Optimized for ATS, role, employer, and geography — so one career profile can serve many applications.", "size": 14, "color": TEAL_SOFT}])
    caps = [
        "AI Resume Generation", "Resume Optimization", "Job-Specific Customization", "Achievement Quantification",
        "ATS Optimization", "Multiple Resume Versions", "Cover Letter Generation", "Career Profile Creation",
    ]
    for i, cap in enumerate(caps):
        col, row = i % 2, i // 2
        x = Inches(8.55) + col * Inches(2.22)
        y = Inches(1.18) + row * Inches(1.40)
        round_rect(s, x, y, Inches(2.10), Inches(1.26), fill=WHITE)
        textbox(s, x + Inches(0.12), y + Inches(0.28), Inches(1.86), Inches(0.72),
                [{"text": cap, "size": 13, "bold": True, "color": NAVY}])
    chrome(s, n)


def s_validation(prs, n):
    s = blank(prs)
    header(s, "07  ·  Product Validation", "The Product Has Already Been Built and Tested")
    checks = [
        ("Demo developed", "A working AI Resume product is available for demonstration."),
        ("Initial testing completed", "Early users and internal reviews show promising quality."),
        ("Internal benchmark testing", "Compared against direct outputs from leading general-purpose AI models."),
    ]
    for i, (t, b) in enumerate(checks):
        x = ML + i * Inches(4.08)
        round_rect(s, x, Inches(1.22), Inches(3.92), Inches(2.05), fill=WHITE)
        textbox(s, x + Inches(0.22), Inches(1.42), Inches(3.48), Inches(0.40),
                [{"text": f"✓   {t}", "size": 16, "bold": True, "color": TEAL}])
        textbox(s, x + Inches(0.22), Inches(1.90), Inches(3.48), Inches(1.15),
                [{"text": b, "size": 14, "color": BODY}])
    textbox(s, ML, Inches(3.40), CW, Inches(0.32),
            [{"text": "INITIAL TESTING INDICATES STRONGER PERFORMANCE IN", "size": 12, "bold": True, "color": TEAL}])
    dims = ["Professional positioning", "Achievement quantification", "Job relevance",
            "Career-experience extraction", "Resume structure"]
    for i, d in enumerate(dims):
        x = ML + i * Inches(2.46)
        round_rect(s, x, Inches(3.76), Inches(2.34), Inches(1.48), fill=WHITE)
        textbox(s, x + Inches(0.12), Inches(4.10), Inches(2.10), Inches(0.95),
                [{"text": d, "size": 15, "bold": True, "color": NAVY}])
    round_rect(s, ML, Inches(5.38), CW, Inches(1.50), fill=SOFT)
    textbox(s, Inches(0.80), Inches(5.48), Inches(11.7), Inches(0.28),
            [{"text": "IMPORTANT NOTE", "size": 12, "bold": True, "color": TEAL}])
    textbox(s, Inches(0.80), Inches(5.82), Inches(11.7), Inches(0.78),
            [{"text": "Current results are based on internal testing. The next stage is larger-scale, structured, blind A/B testing to establish statistically meaningful benchmarks versus general-purpose models.",
              "size": 14, "color": BODY}])
    chrome(s, n)


def s_tech(prs, n):
    s = blank(prs)
    header(s, "08  ·  Technology Advantage", "We Are Not Simply Building an LLM Wrapper")
    round_rect(s, ML, Inches(1.18), CW, Inches(1.05), fill=NAVY)
    textbox(s, Inches(0.80), Inches(1.42), Inches(11.7), Inches(0.58),
            [{"text": "LLM  +  Career Knowledge  +  Resume Framework  +  Job Intelligence  +  User Data",
              "size": 18, "bold": True, "color": WHITE, "font": FONT_D}])
    chips = [
        "Career-experience extraction", "Achievement quantification", "Job-description analysis", "Resume optimization",
        "Industry terminology", "Role-specific requirements", "Geographic resume conventions", "ATS optimization",
    ]
    for i, c in enumerate(chips):
        col, row = i % 4, i // 4
        x = ML + col * Inches(3.08)
        y = Inches(2.38) + row * Inches(0.95)
        round_rect(s, x, y, Inches(2.95), Inches(0.84), fill=WHITE)
        textbox(s, x + Inches(0.14), y + Inches(0.22), Inches(2.67), Inches(0.44),
                [{"text": c, "size": 14, "bold": True, "color": NAVY}])
    textbox(s, ML, Inches(4.38), CW, Inches(0.28),
            [{"text": "DATA FEEDBACK LOOP", "size": 12, "bold": True, "color": TEAL}])
    loop = ["User input", "AI analysis", "Resume gen.", "Application", "Feedback", "Model upgrade", "Better outcomes", "More users"]
    for i, step in enumerate(loop):
        x = ML + i * Inches(1.53)
        round_rect(s, x, Inches(4.70), Inches(1.42), Inches(1.38), fill=WHITE)
        textbox(s, x + Inches(0.06), Inches(4.86), Inches(1.30), Inches(0.32),
                [{"text": f"{i+1:02d}", "size": 12, "bold": True, "color": TEAL}])
        textbox(s, x + Inches(0.06), Inches(5.22), Inches(1.30), Inches(0.70),
                [{"text": step, "size": 13, "bold": True, "color": NAVY}])
    textbox(s, ML, Inches(6.20), CW, Inches(0.68),
            [{"text": "Each cycle compounds proprietary career intelligence that generic chat models do not accumulate.",
              "size": 14, "color": MUTED}])
    chrome(s, n)


def s_global_market(prs, n):
    s = blank(prs)
    header(s, "09  ·  Global Market Opportunity", "AI Resume Is a Naturally Global Product",
           "A resume is a universal component of the global employment market.")
    markets = ["United States", "Canada", "Central America", "United Kingdom", "Australia",
               "Singapore", "Hong Kong", "China", "Europe", "Middle East"]
    for i, m in enumerate(markets):
        col, row = i % 5, i // 5
        x = ML + col * Inches(2.46)
        y = Inches(1.22) + row * Inches(2.15)
        round_rect(s, x, y, Inches(2.34), Inches(2.00), fill=WHITE)
        textbox(s, x + Inches(0.14), y + Inches(0.65), Inches(2.06), Inches(0.70),
                [{"text": m, "size": 16, "bold": True, "color": NAVY}])
    round_rect(s, ML, Inches(5.58), CW, Inches(1.30), fill=NAVY)
    textbox(s, Inches(0.85), Inches(5.70), Inches(11.6), Inches(0.24),
            [{"text": "THE UNIVERSAL QUESTION", "size": 12, "bold": True, "color": TEAL}])
    textbox(s, Inches(0.85), Inches(5.98), Inches(11.6), Inches(0.36),
            [{"text": "“How do I present my professional value to employers?”",
              "size": 20, "bold": True, "color": WHITE, "font": FONT_D}])
    textbox(s, Inches(0.85), Inches(6.38), Inches(11.6), Inches(0.38),
            [{"text": "That shared problem gives AI Resume a fundamentally different scalability profile from traditional, local career-consulting businesses.",
              "size": 13, "color": TEAL_SOFT}])
    chrome(s, n)


def s_strategy(prs, n):
    s = blank(prs)
    header(s, "10  ·  Global Product Strategy", "One Platform. Multiple Labor Markets.")
    dims = [
        ("Geography",
         "U.S.  ·  Canada  ·  Central America  ·  U.K.  ·  Australia  ·  Europe  ·  Asia  ·  Middle East"),
        ("Industry",
         "Applicable across industries — including Technology, Fintech, Financial Services, Advanced Manufacturing, Consulting, Healthcare, Marketing, Education, and more."),
        ("Career level", "Student / Graduate / Professional / Manager / Executive"),
        ("Target role", "Software Engineer / PM / Data Scientist / Marketing / Finance / more"),
        ("Target employer", "Startup / SME / Enterprise / Multinational"),
    ]
    for i, (t, b) in enumerate(dims):
        y = Inches(1.20) + i * Inches(1.14)
        round_rect(s, ML, y, CW, Inches(1.06), fill=WHITE)
        rect(s, ML, y, Inches(0.08), Inches(1.06), fill=TEAL)
        textbox(s, Inches(0.90), y + Inches(0.16), Inches(3.20), Inches(0.74),
                [{"text": t, "size": 20, "bold": True, "color": NAVY, "font": FONT_D}], valign=MSO_ANCHOR.MIDDLE)
        textbox(s, Inches(4.30), y + Inches(0.16), Inches(8.20), Inches(0.74),
                [{"text": b, "size": 15, "color": BODY}], valign=MSO_ANCHOR.MIDDLE)
    chrome(s, n)


def s_phases(prs, n):
    s = blank(prs)
    header(s, "11  ·  From AI Resume to AI Career Agent", "Resume Is the Entry Point — Not the Destination")
    phases = [
        ("PHASE 1", "AI Resume", "“Help me create a better resume.”"),
        ("PHASE 2", "AI Career Assistant", "“What jobs should I apply for?”"),
        ("PHASE 3", "AI Career Agent", "“Help me manage my entire job search.”"),
        ("PHASE 4", "Global Platform", "Job seekers × employers × services × education × talent ecosystem"),
    ]
    for i, (p, t, b) in enumerate(phases):
        x = ML + i * Inches(3.08)
        round_rect(s, x, Inches(1.22), Inches(2.95), Inches(3.70), fill=WHITE)
        rect(s, x, Inches(1.22), Inches(2.95), Inches(0.10), fill=TEAL if i < 3 else NAVY)
        textbox(s, x + Inches(0.16), Inches(1.48), Inches(2.63), Inches(0.28),
                [{"text": p, "size": 12, "bold": True, "color": TEAL}])
        textbox(s, x + Inches(0.16), Inches(1.88), Inches(2.63), Inches(1.00),
                [{"text": t, "size": 20, "bold": True, "color": NAVY, "font": FONT_D}])
        textbox(s, x + Inches(0.16), Inches(2.95), Inches(2.63), Inches(1.70),
                [{"text": b, "size": 15, "color": BODY}])
    caps = ["Resume creation", "Job discovery", "Job matching", "Application optimization",
            "Cover letters", "Interview prep", "Career planning"]
    for i, c in enumerate(caps):
        x = ML + i * Inches(1.75)
        round_rect(s, x, Inches(5.08), Inches(1.65), Inches(1.78), fill=SOFT)
        textbox(s, x + Inches(0.08), Inches(5.48), Inches(1.49), Inches(1.10),
                [{"text": c, "size": 14, "bold": True, "color": NAVY}])
    chrome(s, n)


def s_model(prs, n):
    s = blank(prs)
    header(s, "12  ·  Business Model", "Multiple Monetization Layers Across the Career Journey")
    streams = [
        ("01  AI PRODUCT", "Subscription, premium AI features, resume packages, AI Career tools."),
        ("02  CAREER SERVICES", "Resume consulting, career consulting, interview coaching, premium services."),
        ("03  REFERRAL & COMMISSION", "CPA / CPS / revenue share with education, training, and recruitment partners."),
        ("04  B2B / EMPLOYERS", "Talent acquisition, employer branding, recruitment campaigns, job promotion."),
        ("05  ADVERTISING", "As traffic scales: relevant brands → targeted career audience."),
        ("06  SPONSORSHIP", "Career fairs, conferences, recruitment events, AI career events."),
    ]
    for i, (t, b) in enumerate(streams):
        col, row = i % 3, i // 3
        x = ML + col * Inches(4.08)
        y = Inches(1.22) + row * Inches(2.80)
        round_rect(s, x, y, Inches(3.92), Inches(2.68), fill=WHITE)
        rect(s, x, y, Inches(3.92), Inches(0.10), fill=TEAL)
        textbox(s, x + Inches(0.22), y + Inches(0.32), Inches(3.48), Inches(0.75),
                [{"text": t, "size": 16, "bold": True, "color": NAVY}])
        textbox(s, x + Inches(0.22), y + Inches(1.15), Inches(3.48), Inches(1.30),
                [{"text": b, "size": 15, "color": BODY}])
    chrome(s, n)


def s_flywheel(prs, n):
    s = blank(prs)
    header(s, "13  ·  The Career Monetization Flywheel", "One User Can Generate Multiple Revenue Opportunities")
    steps = ["Free AI Resume", "User acquisition", "Career profile", "AI Career tools", "Premium subscription",
             "Career services", "Job matching", "Recruitment / referral", "Advertising", "Employer services"]
    for i, st in enumerate(steps):
        col, row = i % 5, i // 5
        x = ML + col * Inches(2.46)
        y = Inches(1.22) + row * Inches(1.95)
        round_rect(s, x, y, Inches(2.34), Inches(1.82), fill=WHITE)
        textbox(s, x + Inches(0.14), y + Inches(0.22), Inches(2.06), Inches(0.36),
                [{"text": f"{i+1:02d}", "size": 13, "bold": True, "color": TEAL}])
        textbox(s, x + Inches(0.14), y + Inches(0.65), Inches(2.06), Inches(0.95),
                [{"text": st, "size": 16, "bold": True, "color": NAVY}])
    round_rect(s, ML, Inches(5.22), CW, Inches(1.66), fill=NAVY)
    textbox(s, Inches(0.80), Inches(5.38), Inches(11.7), Inches(0.28),
            [{"text": "THEREFORE", "size": 12, "bold": True, "color": TEAL}])
    textbox(s, Inches(0.80), Inches(5.72), Inches(11.7), Inches(0.95),
            [{"text": "Revenue is not limited to the first AI Resume transaction. The long-term objective is to maximize lifetime value per career user.",
              "size": 18, "bold": True, "color": WHITE, "font": FONT_D}])
    chrome(s, n)


def s_acquisition(prs, n):
    s = blank(prs)
    header(s, "14  ·  User Acquisition", "Digital  +  Offline  +  Organic")
    cols = [
        ("DIGITAL", TEAL, ["Google", "Meta", "TikTok", "YouTube", "LinkedIn", "Search / social", "Content marketing"]),
        ("ORGANIC", BLUE, ["SEO", "AI-search optimization", "Career content", "Educational content", "Referral programs", "Affiliate partnerships"]),
        ("OFFLINE", BRONZE, ["Universities", "Career fairs", "Student organizations", "Professional orgs", "Community events", "Employer events"]),
    ]
    for i, (t, color, items) in enumerate(cols):
        x = ML + i * Inches(4.08)
        round_rect(s, x, Inches(1.22), Inches(3.92), Inches(4.35), fill=WHITE)
        rect(s, x, Inches(1.22), Inches(3.92), Inches(0.10), fill=color)
        textbox(s, x + Inches(0.22), Inches(1.46), Inches(3.48), Inches(0.40),
                [{"text": t, "size": 16, "bold": True, "color": color}])
        for j, it in enumerate(items):
            textbox(s, x + Inches(0.22), Inches(1.98) + j * Inches(0.48), Inches(3.48), Inches(0.44),
                    [{"text": f"▸  {it}", "size": 15, "color": BODY}])
    round_rect(s, ML, Inches(5.70), CW, Inches(1.18), fill=TEAL)
    textbox(s, Inches(0.80), Inches(5.84), Inches(11.7), Inches(0.26),
            [{"text": "CORE GROWTH STRATEGY", "size": 12, "bold": True, "color": WHITE}])
    textbox(s, Inches(0.80), Inches(6.16), Inches(11.7), Inches(0.52),
            [{"text": "Free AI product  →  Low-friction acquisition  →  Registration  →  Engagement  →  Monetization",
              "size": 16, "bold": True, "color": WHITE}])
    chrome(s, n)


def s_geo(prs, n):
    s = blank(prs)
    header(s, "15  ·  China + U.S. + Global", "China as R&D Base  ·  U.S. as Commercialization Market")
    geos = [
        ("CHINA  ·  R&D AND PRODUCT",
         "Product, design, engineering, AI, testing, marketing operations, B2B business development."),
        ("UNITED STATES  ·  INTERNATIONAL GTM",
         "User acquisition, local partnerships, employer and university relationships, offline events, brand building."),
        ("GLOBAL EXPANSION",
         "Once product-market fit is established: North America → UK → Australia → Europe → Asia → other markets."),
    ]
    for i, (t, b) in enumerate(geos):
        y = Inches(1.22) + i * Inches(1.86)
        round_rect(s, ML, y, CW, Inches(1.76), fill=WHITE)
        rect(s, ML, y, Inches(0.08), Inches(1.76), fill=TEAL)
        textbox(s, Inches(0.90), y + Inches(0.22), Inches(11.5), Inches(0.42),
                [{"text": t, "size": 18, "bold": True, "color": NAVY}])
        textbox(s, Inches(0.90), y + Inches(0.72), Inches(11.5), Inches(0.80),
                [{"text": b, "size": 16, "color": BODY}])
    chrome(s, n)


def s_compete(prs, n):
    s = blank(prs)
    header(s, "16  ·  Competitive Landscape", "Career Infrastructure — Not Another Resume Builder")
    quads = [
        (ML, Inches(1.20), "JOB PLATFORMS", "Large user base  ·  Limited personalized career intelligence"),
        (Inches(6.90), Inches(1.20), "GENERAL-PURPOSE AI", "High AI capability  ·  Limited career specialization"),
        (ML, Inches(3.28), "TRADITIONAL CAREER SERVICES", "High human service  ·  Low scalability"),
        (Inches(6.90), Inches(3.28), "AI RESUME TOOLS", "Resume-focused  ·  Limited ecosystem"),
    ]
    for x, y, t, b in quads:
        round_rect(s, x, y, Inches(5.90), Inches(1.95), fill=WHITE)
        textbox(s, x + Inches(0.24), y + Inches(0.28), Inches(5.42), Inches(0.40),
                [{"text": t, "size": 15, "bold": True, "color": MUTED}])
        textbox(s, x + Inches(0.24), y + Inches(0.80), Inches(5.42), Inches(0.90),
                [{"text": b, "size": 17, "color": BODY}])
    round_rect(s, ML, Inches(5.38), CW, Inches(1.50), fill=NAVY)
    textbox(s, Inches(0.80), Inches(5.52), Inches(11.7), Inches(0.28),
            [{"text": "51 CAREERS", "size": 12, "bold": True, "color": TEAL}])
    textbox(s, Inches(0.80), Inches(5.86), Inches(11.7), Inches(0.80),
            [{"text": "AI  +  Career expertise  +  Services  +  Marketplace. We are not competing only with resume builders — we are building career infrastructure.",
              "size": 17, "bold": True, "color": WHITE}])
    chrome(s, n)


def s_advantages(prs, n):
    s = blank(prs)
    header(s, "17  ·  Competitive Advantages", "Why 51 Careers")
    items = [
        ("01", "Existing career-services foundation", "We understand the industry from an operating perspective."),
        ("02", "AI product already in development", "Not a concept or a slide-deck company. Demo built and tested."),
        ("03", "Global scalability", "AI Resume can be distributed worldwide through digital channels."),
        ("04", "Multiple monetization paths", "Subscription + services + commission + B2B + ads + sponsorship."),
        ("05", "AI + human hybrid", "AI for scalable work; humans for high-value, complex services."),
        ("06", "Platform potential", "The long-term opportunity extends beyond resumes into the career ecosystem."),
    ]
    for i, (num, t, b) in enumerate(items):
        col, row = i % 3, i // 3
        x = ML + col * Inches(4.08)
        y = Inches(1.22) + row * Inches(2.80)
        round_rect(s, x, y, Inches(3.92), Inches(2.68), fill=WHITE)
        textbox(s, x + Inches(0.22), y + Inches(0.22), Inches(3.48), Inches(0.42),
                [{"text": num, "size": 20, "bold": True, "color": TEAL, "font": FONT_D}])
        textbox(s, x + Inches(0.22), y + Inches(0.70), Inches(3.48), Inches(0.75),
                [{"text": t, "size": 17, "bold": True, "color": NAVY}])
        textbox(s, x + Inches(0.22), y + Inches(1.50), Inches(3.48), Inches(0.95),
                [{"text": b, "size": 14, "color": BODY}])
    chrome(s, n)


def s_team(prs, n):
    s = blank(prs)
    header(s, "18  ·  Team & Operating Structure", "Core founding team in place  ·  Planned hiring funded by this round")
    round_rect(s, ML, Inches(1.20), Inches(3.92), Inches(5.00), fill=WHITE)
    rect(s, ML, Inches(1.20), Inches(3.92), Inches(0.10), fill=TEAL)
    textbox(s, Inches(0.80), Inches(1.42), Inches(3.50), Inches(0.26),
            [{"text": "EXISTING", "size": 11, "bold": True, "color": TEAL}])
    textbox(s, Inches(0.80), Inches(1.72), Inches(3.50), Inches(0.40),
            [{"text": "FOUNDING TEAM", "size": 16, "bold": True, "color": NAVY, "font": FONT_D}])
    textbox(s, Inches(0.80), Inches(2.18), Inches(3.50), Inches(0.70),
            [{"text": "4", "size": 44, "bold": True, "color": NAVY, "font": FONT_D}])
    textbox(s, Inches(0.80), Inches(2.95), Inches(3.50), Inches(0.36),
            [{"text": "Core founding team members", "size": 14, "bold": True, "color": BODY}])
    for i, line in enumerate([
        "In place today",
        "Product, capital, and go-to-market leadership",
        "Additional roles are not yet hired",
        "This round funds the planned operating team",
    ]):
        textbox(s, Inches(0.80), Inches(3.40) + i * Inches(0.52), Inches(3.50), Inches(0.48),
                [{"text": f"▸  {line}", "size": 14, "color": BODY}])

    planned = [
        ("CHINA TECHNOLOGY", "7", "Product, design, engineering, and AI.",
         [("Product Manager", "1"), ("Designer", "1"), ("QA / Testing", "1"),
          ("Front-End Engineer", "1"), ("Back-End Engineers", "2"), ("AI Engineer", "1")]),
        ("CHINA GROWTH", "2+", "Performance marketing and B2B development.",
         [("Marketing / Performance", "1"), ("B2B Business Development", "1+")]),
        ("U.S. GROWTH", "3", "Paid digital and local field marketing.",
         [("Performance Marketing", "1"), ("Local / Field Marketing", "2")]),
    ]
    for i, (title, count, purpose, roles) in enumerate(planned):
        x = Inches(4.70) + i * Inches(2.82)
        round_rect(s, x, Inches(1.20), Inches(2.70), Inches(5.00), fill=WHITE)
        rect(s, x, Inches(1.20), Inches(2.70), Inches(0.10), fill=LINE)
        textbox(s, x + Inches(0.14), Inches(1.40), Inches(2.42), Inches(0.22),
                [{"text": "PLANNED", "size": 11, "bold": True, "color": MUTED}])
        textbox(s, x + Inches(0.14), Inches(1.62), Inches(2.42), Inches(0.50),
                [{"text": title, "size": 13, "bold": True, "color": NAVY}])
        textbox(s, x + Inches(0.14), Inches(2.12), Inches(2.42), Inches(0.42),
                [{"text": count, "size": 26, "bold": True, "color": NAVY, "font": FONT_D}])
        textbox(s, x + Inches(0.14), Inches(2.56), Inches(2.42), Inches(0.70),
                [{"text": purpose, "size": 12, "color": MUTED}])
        for j, (role, nhead) in enumerate(roles):
            textbox(s, x + Inches(0.14), Inches(3.30) + j * Inches(0.42), Inches(1.85), Inches(0.40),
                    [{"text": role, "size": 11, "color": BODY}])
            textbox(s, x + Inches(1.95), Inches(3.30) + j * Inches(0.42), Inches(0.58), Inches(0.40),
                    [{"text": nhead, "size": 11, "bold": True, "color": TEAL, "align": PP_ALIGN.RIGHT}])
    round_rect(s, ML, Inches(6.32), CW, Inches(0.58), fill=SOFT)
    textbox(s, Inches(0.80), Inches(6.40), Inches(11.7), Inches(0.42),
            [{"text": "Only the core founding team is in place today. China technology, China growth, and U.S. growth roles are planned hires — not current headcount.",
              "size": 13, "color": BODY}])
    chrome(s, n)


def s_funds(prs, n):
    s = blank(prs)
    header(s, "19  ·  Use of Funds", "RMB 20 Million Growth Financing",
           "RMB 13.1M is a planned personnel + advertising budget. Remaining capital funds infrastructure, AI, operations, events, expansion and working capital.")
    alloc = [
        ("Marketing & user acquisition", "40%", "RMB 8.0M"),
        ("Product & AI development", "22.5%", "RMB 4.5M"),
        ("U.S. market expansion", "12.5%", "RMB 2.5M"),
        ("China market & B2B", "7.5%", "RMB 1.5M"),
        ("Brand & events", "7.5%", "RMB 1.5M"),
        ("Operations / legal / finance", "5%", "RMB 1.0M"),
        ("Cash reserve", "5%", "RMB 1.0M"),
    ]
    data = CategoryChartData()
    data.categories = [a[0] for a in alloc]
    data.add_series("Use of Funds", (40, 22.5, 12.5, 7.5, 7.5, 5, 5))
    chart = s.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(0.40), Inches(1.32), Inches(5.40), Inches(3.55), data).chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    for i, (name, pct, amt) in enumerate(alloc):
        y = Inches(1.38) + i * Inches(0.50)
        textbox(s, Inches(5.85), y, Inches(4.40), Inches(0.48),
                [{"text": name, "size": 13, "color": BODY}])
        textbox(s, Inches(10.30), y, Inches(2.45), Inches(0.48),
                [{"text": f"{pct}   ·   {amt}", "size": 13, "bold": True, "color": NAVY, "align": PP_ALIGN.RIGHT}])
    kpis = [
        ("PERSONNEL", "RMB 5.1M", "Planned hiring budget — not current payroll"),
        ("CHINA ADS", "RMB 1.0M", "Planned annual spend"),
        ("U.S. ADS", "RMB 7.0M", "Planned annual spend"),
        ("P+A BUDGET", "RMB 13.1M", "Planned personnel + advertising"),
    ]
    for i, (lab, big, sub) in enumerate(kpis):
        x = ML + i * Inches(3.08)
        round_rect(s, x, Inches(4.98), Inches(2.95), Inches(1.90), fill=WHITE)
        textbox(s, x + Inches(0.16), Inches(5.10), Inches(2.63), Inches(0.26),
                [{"text": lab, "size": 11, "bold": True, "color": TEAL}])
        textbox(s, x + Inches(0.16), Inches(5.40), Inches(2.63), Inches(0.52),
                [{"text": big, "size": 22, "bold": True, "color": NAVY, "font": FONT_D}])
        textbox(s, x + Inches(0.16), Inches(5.98), Inches(2.63), Inches(0.70),
                [{"text": sub, "size": 13, "color": MUTED}])
    chrome(s, n)


def s_roadmap(prs, n):
    s = blank(prs)
    header(s, "20  ·  Growth Roadmap", "Validation  →  Growth  →  Global Expansion")
    phases = [
        ("PHASE 1  ·  0–6 MONTHS", "Product validation",
         ["Launch AI Resume", "Improve AI quality", "Structured A/B testing", "Initial CAC benchmarks", "Conversion funnel", "Validate pricing"]),
        ("PHASE 2  ·  6–12 MONTHS", "Growth",
         ["Scale paid acquisition", "Build SEO engine", "Expand U.S. acquisition", "B2B partnerships", "More AI Career features", "Expand monetization"]),
        ("PHASE 3  ·  12–24 MONTHS", "Global expansion",
         ["More English-speaking markets", "Localized products", "Global partnerships", "Employer ecosystem", "Marketplace capabilities", "Central America expansion"]),
    ]
    for i, (lab, title, items) in enumerate(phases):
        x = ML + i * Inches(4.08)
        round_rect(s, x, Inches(1.22), Inches(3.92), Inches(5.55), fill=WHITE)
        rect(s, x, Inches(1.22), Inches(3.92), Inches(0.90), fill=TEAL if i < 2 else NAVY)
        textbox(s, x + Inches(0.22), Inches(1.32), Inches(3.48), Inches(0.28),
                [{"text": lab, "size": 11, "bold": True, "color": WHITE}])
        textbox(s, x + Inches(0.22), Inches(1.62), Inches(3.48), Inches(0.38),
                [{"text": title, "size": 18, "bold": True, "color": WHITE, "font": FONT_D}])
        for j, it in enumerate(items):
            textbox(s, x + Inches(0.22), Inches(2.35) + j * Inches(0.65), Inches(3.48), Inches(0.58),
                    [{"text": f"▸  {it}", "size": 14, "color": BODY}])
    chrome(s, n)


def s_model_3yr(prs, n):
    s = blank(prs)
    header(
        s,
        "21  ·  Three-Year Financial Model",
        "Users, Revenue, and Profitability",
        "Illustrative 2027E–2029E operating plan. Figures in US$.",
    )

    panels = [
        (
            "USERS",
            [
                ("Users", "1.0M", "3.0M", "7.0M"),
                ("Registered users", "300K", "1.0M", "2.5M"),
                ("Monthly active users", "150K", "500K", "1.25M"),
                ("Paid users", "15K", "60K", "180K"),
                ("Paid conversion", "5.0%", "6.0%", "7.2%"),
                ("MAU / registered", "50%", "50%", "50%"),
            ],
        ),
        (
            "REVENUE",
            [
                ("AI revenue", "$0.70M", "$2.40M", "$8.40M"),
                ("Career services", "$0.50M", "$1.60M", "$5.60M"),
                ("Commission", "$0.30M", "$1.20M", "$4.20M"),
                ("Advertising", "$0.20M", "$0.80M", "$2.80M"),
                ("B2B revenue", "$0.30M", "$2.00M", "$7.00M"),
                ("Total revenue", "$2.00M", "$8.00M", "$28.00M"),
            ],
        ),
        (
            "PROFITABILITY",
            [
                ("Gross margin", "65%", "70%", "75%"),
                ("Gross profit", "$1.30M", "$5.60M", "$21.00M"),
                ("EBITDA margin", "–50%", "–15%", "15%"),
                ("EBITDA", "–$1.00M", "–$1.20M", "$4.20M"),
                ("Profitability", "Loss", "Loss", "Profitable"),
            ],
        ),
    ]

    years = ["2027E", "2028E", "2029E"]
    for i, (title, rows) in enumerate(panels):
        x = ML + i * Inches(4.08)
        round_rect(s, x, Inches(1.22), Inches(3.92), Inches(5.48), fill=WHITE)
        rect(s, x, Inches(1.22), Inches(3.92), Inches(0.08), fill=TEAL)
        textbox(
            s, x + Inches(0.16), Inches(1.38), Inches(3.60), Inches(0.30),
            [{"text": title, "size": 13, "bold": True, "color": TEAL}],
        )
        col_w = Inches(1.16)
        col0 = x + Inches(0.28)
        for yi, yr in enumerate(years):
            textbox(
                s, col0 + yi * col_w, Inches(1.70), col_w, Inches(0.24),
                [{"text": yr, "size": 11, "bold": True, "color": MUTED, "align": PP_ALIGN.CENTER}],
            )
        row_h = 4.50 / max(len(rows), 1)
        for ri, (metric, *vals) in enumerate(rows):
            y = Inches(2.00) + ri * Inches(row_h)
            strong = metric in {"Total revenue", "Profitability", "Users"}
            if strong:
                rect(s, x + Inches(0.10), y, Inches(3.72), Inches(row_h - 0.04), fill=SOFT)
            textbox(
                s, x + Inches(0.16), y + Inches(0.04), Inches(3.60), Inches(0.24),
                [{"text": metric, "size": 12, "bold": True, "color": MUTED}],
            )
            for vi, val in enumerate(vals):
                accent = TEAL if (metric == "Profitability" and val == "Profitable") else NAVY
                textbox(
                    s, col0 + vi * col_w, y + Inches(0.30), col_w, Inches(0.42),
                    [{"text": val, "size": 15, "bold": True, "color": accent, "font": FONT_D,
                      "align": PP_ALIGN.CENTER}],
                )
    textbox(
        s, ML, Inches(6.74), CW, Inches(0.28),
        [{"text": "Management projections for discussion. Not a forecast. Paid conversion is paid users / registered users.",
          "size": 11, "color": MUTED}],
    )
    chrome(s, n)


def s_investment(prs, n):
    s = blank(prs)
    header(s, "22  ·  Investment Opportunity", "The Next Generation of Career Services Will Be AI + Human + Platform")
    axes = [
        ("Generative AI", "Professional content creation is now cheap, fast, and widely adopted."),
        ("Global employment", "Talent is mobile. English-language careers travel across markets."),
        ("Career services", "Advice is still manual, expensive, and hard to scale."),
        ("Digital user acquisition", "A free AI Resume is a low-friction, global top-of-funnel."),
    ]
    for i, (a, b) in enumerate(axes):
        col, row = i % 2, i // 2
        x = ML + col * Inches(6.16)
        y = Inches(1.22) + row * Inches(2.15)
        round_rect(s, x, y, Inches(6.00), Inches(2.02), fill=WHITE)
        textbox(s, x + Inches(0.24), y + Inches(0.22), Inches(5.52), Inches(0.50),
                [{"text": a, "size": 20, "bold": True, "color": NAVY, "font": FONT_D}])
        textbox(s, x + Inches(0.24), y + Inches(0.82), Inches(5.52), Inches(0.95),
                [{"text": b, "size": 16, "color": BODY}])
    round_rect(s, ML, Inches(5.62), CW, Inches(1.26), fill=NAVY)
    textbox(s, Inches(0.85), Inches(5.74), Inches(11.6), Inches(0.24),
            [{"text": "THE OPPORTUNITY", "size": 12, "bold": True, "color": TEAL}])
    textbox(s, Inches(0.85), Inches(6.02), Inches(11.6), Inches(0.70), [
        {"text": "The next generation of career services will be AI + human + platform. The AI Resume is the entry point; the ambition is the global AI career platform.",
         "size": 16, "bold": True, "color": WHITE, "font": FONT_D},
    ])
    chrome(s, n)


def s_financing(prs, n):
    s = blank(prs)
    header(s, "23  ·  Financing", "Current Financing Round")
    terms = [
        ("Raise", "RMB 20 million", "≈ US$2.97 million"),
        ("Pre-money", "RMB 180 million", "≈ US$26.7 million"),
        ("Post-money", "RMB 200 million", "≈ US$29.7 million"),
        ("Equity offered", "10%", "Primary preferred"),
        ("Structure", "Preferred equity", "Indicative; subject to docs"),
        ("Use of capital", "Product, users, global GTM", "Monetization + ecosystem"),
    ]
    for i, (lab, a, b) in enumerate(terms):
        col, row = i % 3, i // 3
        x = ML + col * Inches(4.08)
        y = Inches(1.22) + row * Inches(2.35)
        round_rect(s, x, y, Inches(3.92), Inches(2.22), fill=WHITE)
        rect(s, x, y, Inches(3.92), Inches(0.10), fill=TEAL)
        textbox(s, x + Inches(0.22), y + Inches(0.24), Inches(3.48), Inches(0.30),
                [{"text": lab.upper(), "size": 12, "bold": True, "color": TEAL}])
        textbox(s, x + Inches(0.22), y + Inches(0.60), Inches(3.48), Inches(0.65),
                [{"text": a, "size": 22, "bold": True, "color": NAVY, "font": FONT_D}])
        textbox(s, x + Inches(0.22), y + Inches(1.32), Inches(3.48), Inches(0.65),
                [{"text": b, "size": 15, "color": MUTED}])
    round_rect(s, ML, Inches(6.00), CW, Inches(0.88), fill=SOFT)
    textbox(s, Inches(0.80), Inches(6.16), Inches(11.7), Inches(0.58),
            [{"text": "Capital will be used to build the product, acquire users, expand globally, build monetization, and develop the career ecosystem.",
              "size": 16, "color": BODY}])
    chrome(s, n)


def s_what_20m(prs, n):
    s = blank(prs)
    header(s, "RMB 20M  →  What Investors Get", "12–18 month targets this round is designed to fund")
    outcomes = [
        ("01", "Public launch", "AI Resume live for users in China and the U.S."),
        ("02", "50–100k users", "Registered users; funnel instrumented end-to-end."),
        ("03", "Blind A/B tests", "Statistically meaningful quality vs. general-purpose models."),
        ("04", "Unit economics", "CAC, conversion, ARPU, and payback benchmarks."),
        ("05", "U.S. acquisition", "Paid digital + university/offline channels operating."),
        ("06", "First paid cohort", "AI subscriptions and career-services attach live."),
        ("07", "B2B pipeline", "Employer and education-partner conversations in market."),
        ("08", "Next-round pack", "Traction, model, and data room ready for the following raise."),
    ]
    for i, (num, t, b) in enumerate(outcomes):
        col, row = i % 4, i // 4
        x = ML + col * Inches(3.08)
        y = Inches(1.20) + row * Inches(2.45)
        round_rect(s, x, y, Inches(2.95), Inches(2.32), fill=WHITE)
        textbox(s, x + Inches(0.16), y + Inches(0.16), Inches(2.63), Inches(0.34),
                [{"text": num, "size": 16, "bold": True, "color": TEAL, "font": FONT_D}])
        textbox(s, x + Inches(0.16), y + Inches(0.54), Inches(2.63), Inches(0.58),
                [{"text": t, "size": 17, "bold": True, "color": NAVY}])
        textbox(s, x + Inches(0.16), y + Inches(1.18), Inches(2.63), Inches(0.95),
                [{"text": b, "size": 14, "color": BODY}])
    round_rect(s, ML, Inches(6.18), CW, Inches(0.70), fill=TEAL)
    textbox(s, Inches(0.80), Inches(6.30), Inches(11.7), Inches(0.48),
            [{"text": "RMB 20M  →  product + users + benchmarks  →  revenue  →  next round  →  valuation step-up.",
              "size": 16, "bold": True, "color": WHITE}])
    chrome(s, n)


def s_vision(prs, n):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, "0B1F33")
    rect(s, Inches(0), Inches(0), SW, Inches(0.12), fill=TEAL)
    rect(s, Inches(0), Inches(0), Inches(0.16), SH, fill=TEAL)
    textbox(s, Inches(0.90), Inches(1.15), Inches(11.5), Inches(0.40),
            [{"text": "24  ·  VISION", "size": 14, "bold": True, "color": TEAL}])
    textbox(s, Inches(0.90), Inches(1.80), Inches(11.5), Inches(1.30),
            [{"text": "We are not building another AI resume tool.",
              "size": 32, "bold": True, "color": WHITE, "font": FONT_D}])
    textbox(s, Inches(0.90), Inches(3.25), Inches(11.5), Inches(1.30),
            [{"text": "We are building the Global AI Career Platform.",
              "size": 32, "bold": True, "color": TEAL, "font": FONT_D}])
    textbox(s, Inches(0.90), Inches(4.75), Inches(11.5), Inches(0.50),
            [{"text": "AI   ×   Career   ×   Talent   ×   Opportunity", "size": 20, "color": TEAL_SOFT}])
    textbox(s, Inches(0.90), Inches(5.45), Inches(11.5), Inches(0.50),
            [{"text": "51 CAREERS", "size": 22, "bold": True, "color": WHITE, "font": FONT_D}])
    rect(s, Inches(0), Inches(6.35), SW, Inches(1.15), fill=COVER)
    textbox(s, Inches(0.90), Inches(6.58), Inches(11.5), Inches(0.32),
            [{"text": "CHINA R&D   ·   U.S. COMMERCIALIZATION   ·   GLOBAL DISTRIBUTION",
              "size": 14, "bold": True, "color": TEAL_SOFT}])
    textbox(s, Inches(0.90), Inches(6.95), Inches(11.5), Inches(0.32),
            [{"text": f"51 CAREERS   ·   CONFIDENTIAL   ·   {n}", "size": 10, "bold": True, "color": RGBColor(0x6B, 0x7C, 0x8F)}])


def build():
    prs = Presentation()
    prs.slide_width = SW
    prs.slide_height = SH
    s_cover(prs)
    builders = [
        s_exec, s_problem, s_opportunity, s_why_now, s_tam, s_solution, s_product,
        s_validation, s_tech, s_global_market, s_strategy, s_phases, s_model,
        s_flywheel, s_acquisition, s_geo, s_compete, s_advantages, s_team,
        s_funds, s_roadmap, s_model_3yr, s_investment, s_financing,
        s_what_20m, s_vision,
    ]
    for i, fn in enumerate(builders, start=2):
        fn(prs, i)
    prs.save(OUT)
    print(f"Wrote {OUT} ({len(prs.slides)} slides)")


if __name__ == "__main__":
    build()
