#!/usr/bin/env python3
"""Apply visual enhancements to the 51 Careers.AI investor presentation."""
from __future__ import annotations

import shutil
from copy import deepcopy

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

SRC = "/home/ubuntu/.cursor/projects/workspace/uploads/51careers.ai_Presentation_updated_ab7b.pptx"
DST = "/workspace/51careers.ai Presentation_updated.pptx"
GEN = "/workspace/deck_assets/generated"

NAVY = RGBColor(0x10, 0x22, 0x2F)
BODY = RGBColor(0x1E, 0x2A, 0x33)
BLUE = RGBColor(0x00, 0x7B, 0xCB)
MUTED = RGBColor(0x5C, 0x6B, 0x76)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
CARD = RGBColor(0xF3, 0xF7, 0xFA)


def rgb(hex_str: str) -> RGBColor:
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def by_id(slide, sid):
    for sh in slide.shapes:
        if sh.shape_id == sid:
            return sh
    raise KeyError(sid)


def delete_ids(slide, ids):
    want = set(ids)
    for sh in list(slide.shapes):
        if sh.shape_id in want:
            el = sh._element
            el.getparent().remove(el)


def pos(sh, l=None, t=None, w=None, h=None):
    if l is not None:
        sh.left = Inches(l)
    if t is not None:
        sh.top = Inches(t)
    if w is not None:
        sh.width = Inches(w)
    if h is not None:
        sh.height = Inches(h)


def pic(slide, path, l, t, w, h):
    return slide.shapes.add_picture(path, Inches(l), Inches(t), Inches(w), Inches(h))


def zero_margins(tf):
    tf.margin_left = Inches(0.04)
    tf.margin_right = Inches(0.04)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)


def add_text(slide, l, t, w, h, text, size, bold=False, color=NAVY, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font="Garamond"):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    try:
        tf._txBody.bodyPr.set("anchor", {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}.get(anchor, "t"))
    except Exception:
        pass
    zero_margins(tf)
    p = tf.paragraphs[0]
    p.alignment = align
    p.space_before = Pt(0)
    p.space_after = Pt(0)
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = font
    run.font.color.rgb = color
    return box


def add_round_rect(slide, l, t, w, h, fill=CARD):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    try:
        sh.adjustments[0] = 0.06
    except Exception:
        pass
    return sh


def add_oval(slide, l, t, w, h, fill):
    sh = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(l), Inches(t), Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    return sh


def set_run_size(shape, pt, wrap=None):
    if not shape.has_text_frame:
        return
    if wrap is not None:
        shape.text_frame.word_wrap = wrap
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            r.font.size = Pt(pt)


def edit_slide_3(slide):
    """Widen the three problem boxes to the standard 3-column grid."""
    cols = [
        # box, circle, number, title, body, icon
        (146, 147, 148, 149, 150, "icon_fragmented.png"),
        (151, 152, 153, 154, 155, "icon_expensive.png"),
        (156, 157, 158, 159, 160, "icon_lowsignal.png"),
    ]
    xs = [0.90, 4.81, 8.72]
    for x, (box, circ, num, title, body, icon) in zip(xs, cols):
        pos(by_id(slide, box), l=x, w=3.71)
        pos(by_id(slide, circ), l=x + 0.22)
        pos(by_id(slide, num), l=x + 0.22)
        pos(by_id(slide, title), l=x + 0.22, w=2.70)
        pos(by_id(slide, body), l=x + 0.22, w=3.27)
        pic(slide, f"{GEN}/{icon}", x + 3.12, 1.58, 0.36, 0.36)


def edit_slide_8(slide):
    """Full-width before/after résumé rewrite — spans the content area."""
    delete_ids(slide, list(range(312, 346)))
    slide.shapes.add_picture(
        f"{GEN}/resume_before_after.png",
        Inches(0.90),
        Inches(1.38),
        width=Inches(11.53),
        height=Inches(5.62),
    )


def edit_slide_10(slide):
    """Small technology icons on the eight capability tiles."""
    tiles = [
        (0.90, 1.42, "icon_extract.png"),
        (3.81, 1.42, "icon_scoring.png"),
        (6.72, 1.42, "icon_jd.png"),
        (9.63, 1.42, "icon_resume.png"),
        (0.90, 2.60, "icon_industry.png"),
        (3.81, 2.60, "icon_role.png"),
        (6.72, 2.60, "icon_geo.png"),
        (9.63, 2.60, "icon_ats.png"),
    ]
    for x, y, name in tiles:
        pic(slide, f"{GEN}/{name}", x + 2.42, y + 0.14, 0.26, 0.26)
    # keep the data-loop labels on one line (Feedback / Outcomes were wrapping)
    for sid in (412, 413, 415, 416, 418, 419, 421, 422, 424, 425, 427, 428, 430, 431, 433, 434):
        set_run_size(by_id(slide, sid), 13, wrap=False)


def edit_slide_11(slide):
    """Full-width world map with a compact chip legend underneath."""
    delete_ids(slide, range(448, 478))
    pic(slide, f"{GEN}/world_map.png", 0.90, 1.38, 11.53, 3.90)

    chips = [
        ("#007BCB", "United States"),
        ("#4FA3D4", "Canada"),
        ("#7FB7D4", "Central America"),
        ("#4FA3D4", "United Kingdom"),
        ("#4FA3D4", "Australia"),
        ("#007BCB", "Singapore"),
        ("#007BCB", "Hong Kong"),
        ("#0A5F9A", "China"),
        ("#7FB7D4", "Europe"),
        ("#7FB7D4", "Middle East"),
    ]
    x = 0.90
    y = 5.36
    w = 2.22
    gap = 0.08
    for i, (color, name) in enumerate(chips):
        if i == 5:
            x = 0.90
            y = 5.76
        add_round_rect(slide, x, y, w, 0.34, CARD)
        add_oval(slide, x + 0.10, y + 0.10, 0.14, 0.14, rgb(color))
        add_text(slide, x + 0.30, y, w - 0.38, 0.34, name, 12, True, NAVY, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        x += w + gap


def edit_slide_12(slide):
    """Row icons for the five localization dimensions."""
    icons = [
        (1.42, "icon_globe.png"),
        (2.36, "icon_industry.png"),
        (3.30, "icon_levels.png"),
        (4.24, "icon_role.png"),
        (5.18, "icon_employer.png"),
    ]
    labels = [492, 495, 498, 501, 504]
    for (y, icon), lid in zip(icons, labels):
        pic(slide, f"{GEN}/{icon}", 1.02, y + 0.26, 0.34, 0.34)
        pos(by_id(slide, lid), l=1.42, w=2.10)


def edit_slide_14(slide):
    """Business-model icons in each revenue card."""
    items = [
        (0.90, 1.42, "icon_ai.png"),
        (4.81, 1.42, "icon_services.png"),
        (8.72, 1.42, "icon_referral.png"),
        (0.90, 3.74, "icon_b2b.png"),
        (4.81, 3.74, "icon_ads.png"),
        (8.72, 3.74, "icon_sponsor.png"),
    ]
    for x, y, name in items:
        pic(slide, f"{GEN}/{name}", x + 3.16, y + 0.16, 0.38, 0.38)


def edit_slide_16(slide):
    """Official channel logos on Digital; icons on every Organic and Offline row."""
    logos = [
        (1.96, "logo_google.png"),
        (2.52, "logo_meta.png"),
        (3.08, "logo_tiktok.png"),
        (3.64, "logo_youtube.png"),
        (4.20, "logo_linkedin.png"),
    ]
    texts = [668, 669, 670, 671, 672]
    for (y, logo), tid in zip(logos, texts):
        pic(slide, f"{GEN}/{logo}", 1.16, y + 0.10, 0.30, 0.30)
        pos(by_id(slide, tid), l=1.54, w=2.82)

    digital_more = [
        (4.76, 673, "icon_search.png"),
        (5.32, 674, "icon_content.png"),
    ]
    for y, tid, icon in digital_more:
        pic(slide, f"{GEN}/{icon}", 1.16, y + 0.10, 0.28, 0.28)
        pos(by_id(slide, tid), l=1.54, w=2.82)

    organic = [
        (1.96, 677, "icon_search.png"),
        (2.52, 678, "icon_ai.png"),
        (3.08, 679, "icon_content.png"),
        (3.64, 680, "icon_jd.png"),
        (4.20, 681, "icon_referral.png"),
        (4.76, 682, "icon_link.png"),
        (5.32, 683, "icon_services.png"),
    ]
    for y, tid, icon in organic:
        pic(slide, f"{GEN}/{icon}", 5.07, y + 0.10, 0.28, 0.28)
        pos(by_id(slide, tid), l=5.42, w=2.85)

    offline = [
        (1.96, 686, "icon_university.png"),
        (2.52, 687, "icon_calendar.png"),
        (3.08, 688, "icon_role.png"),
        (3.64, 689, "icon_b2b.png"),
        (4.20, 690, "icon_sponsor.png"),
        (4.76, 691, "icon_employer.png"),
        (5.32, 692, "icon_services.png"),
    ]
    for y, tid, icon in offline:
        pic(slide, f"{GEN}/{icon}", 8.98, y + 0.10, 0.28, 0.28)
        pos(by_id(slide, tid), l=9.34, w=2.85)


def edit_slide_17(slide):
    """Banner map of the Shanghai → New York operating route."""
    # shrink and move the three operating-model bars under the map
    groups = [
        # card, circle, number, title, body, orig_top
        (706, 707, 708, 709, 710, 1.42),
        (711, 712, 713, 714, 715, 2.99),
        (716, 717, 718, 719, 720, 4.57),
    ]
    new_tops = [3.18, 4.18, 5.18]
    new_h = 0.96
    for (card, circ, num, title, body, orig), new_top in zip(groups, new_tops):
        dy = new_top - orig
        pos(by_id(slide, card), t=new_top, h=new_h)
        pos(by_id(slide, circ), t=new_top + 0.30)
        pos(by_id(slide, num), t=new_top + 0.30)
        pos(by_id(slide, title), t=new_top + 0.08, h=0.28)
        b = by_id(slide, body)
        pos(b, t=new_top + 0.38, h=0.54)
        set_run_size(b, 13)
        set_run_size(by_id(slide, title), 16)

    pic(slide, f"{GEN}/shanghai_ny_map.png", 0.90, 1.36, 11.53, 1.74)


def edit_slide_20(slide):
    """Circular monogram portraits on each leadership card."""
    people = [
        (802, 803, 804, 805, "headshot_rc.png"),
        (806, 807, 808, 809, "headshot_sl.png"),
        (810, 811, 812, 813, "headshot_gd.png"),
        (814, 815, 816, 817, "headshot_rz.png"),
        (818, 819, 820, 821, "headshot_cl.png"),
        (822, 823, 824, 825, "headshot_js.png"),
    ]
    for card_id, name_id, title_id, bio_id, img in people:
        card = by_id(slide, card_id)
        x = card.left.inches
        y = card.top.inches
        pic(slide, f"{GEN}/{img}", x + 0.16, y + 0.14, 0.72, 0.72)
        pos(by_id(slide, name_id), l=x + 0.98, t=y + 0.18, w=2.55)
        pos(by_id(slide, title_id), l=x + 0.98, t=y + 0.50, w=2.55)
        pos(by_id(slide, bio_id), t=y + 0.96, h=1.50)


def edit_slide_24(slide):
    """Stop metric values wrapping by reducing type size."""
    number_ids = (
        list(range(982, 985))
        + list(range(986, 989))
        + list(range(990, 993))
        + list(range(994, 997))
        + list(range(998, 1001))
        + list(range(1007, 1010))
        + list(range(1011, 1014))
        + list(range(1015, 1018))
        + list(range(1019, 1022))
        + list(range(1023, 1026))
        + list(range(1027, 1030))
        + list(range(1036, 1039))
        + list(range(1040, 1043))
        + list(range(1044, 1047))
        + list(range(1048, 1051))
        + list(range(1052, 1054))
    )
    for sid in number_ids:
        set_run_size(by_id(slide, sid), 11, wrap=False)
    set_run_size(by_id(slide, 1054), 9, wrap=False)  # Profitable
    for sid in (
        981, 985, 989, 993, 997,
        1006, 1010, 1014, 1018, 1022, 1026,
        1035, 1039, 1043, 1047, 1051,
    ):
        set_run_size(by_id(slide, sid), 11)
    set_run_size(by_id(slide, 1051), 10)  # Profitability label
    for sid in (997, 1010, 1035, 1043):
        set_run_size(by_id(slide, sid), 10, wrap=True)


def main():
    shutil.copyfile(SRC, DST)
    prs = Presentation(DST)
    edit_slide_3(prs.slides[2])
    edit_slide_8(prs.slides[7])
    edit_slide_10(prs.slides[9])
    edit_slide_11(prs.slides[10])
    edit_slide_12(prs.slides[11])
    edit_slide_14(prs.slides[13])
    edit_slide_16(prs.slides[15])
    edit_slide_17(prs.slides[16])
    edit_slide_20(prs.slides[19])
    edit_slide_24(prs.slides[23])
    prs.save(DST)
    print("saved", DST)


if __name__ == "__main__":
    main()
