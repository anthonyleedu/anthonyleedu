#!/usr/bin/env python3
"""Generate visual assets for the 51 Careers.AI investor deck."""
from __future__ import annotations

import json
import math
import os
from io import BytesIO

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.collections import PatchCollection
from matplotlib.patches import Polygon as MplPolygon, FancyBboxPatch, Circle, FancyArrowPatch, Rectangle
from matplotlib.path import Path
from matplotlib.patches import PathPatch
import numpy as np
from PIL import Image, ImageDraw, ImageFilter, ImageFont

ROOT = "/workspace/deck_assets"
OUT = os.path.join(ROOT, "generated")
os.makedirs(OUT, exist_ok=True)

NAVY = "#10222F"
BODY = "#1E2A33"
BLUE = "#007BCB"
BLUE_DEEP = "#0A5F9A"
BLUE_MID = "#4FA3D4"
BLUE_PALE = "#A9CDE0"
CARD = "#F3F7FA"
OCEAN = "#D5E8F3"
LAND = "#A8C0CF"
GLOW = "#8EC8E8"
MUTED = "#5C6B76"
WHITE = "#FFFFFF"

ISO_NAME = {
    "USA": "United States",
    "CAN": "Canada",
    "MEX": "Mexico",
    "GTM": "Guatemala",
    "HND": "Honduras",
    "SLV": "El Salvador",
    "NIC": "Nicaragua",
    "CRI": "Costa Rica",
    "PAN": "Panama",
    "BLZ": "Belize",
    "GBR": "United Kingdom",
    "AUS": "Australia",
    "CHN": "China",
    "FRA": "France",
    "DEU": "Germany",
    "ITA": "Italy",
    "ESP": "Spain",
    "NLD": "Netherlands",
    "BEL": "Belgium",
    "SWE": "Sweden",
    "NOR": "Norway",
    "CHE": "Switzerland",
    "AUT": "Austria",
    "POL": "Poland",
    "IRL": "Ireland",
    "PRT": "Portugal",
    "DNK": "Denmark",
    "FIN": "Finland",
    "CZE": "Czechia",
    "GRC": "Greece",
    "ROU": "Romania",
    "HUN": "Hungary",
    "ARE": "UAE",
    "SAU": "Saudi Arabia",
    "QAT": "Qatar",
    "BHR": "Bahrain",
    "KWT": "Kuwait",
    "OMN": "Oman",
    "ISR": "Israel",
    "JOR": "Jordan",
}

HIGHLIGHT = {
    "USA": BLUE,
    "CAN": BLUE_MID,
    "MEX": BLUE_PALE,
    "GTM": BLUE_PALE,
    "HND": BLUE_PALE,
    "SLV": BLUE_PALE,
    "NIC": BLUE_PALE,
    "CRI": BLUE_PALE,
    "PAN": BLUE_PALE,
    "BLZ": BLUE_PALE,
    "GBR": BLUE_MID,
    "AUS": BLUE_MID,
    "CHN": BLUE_DEEP,
    "FRA": BLUE_PALE,
    "DEU": BLUE_PALE,
    "ITA": BLUE_PALE,
    "ESP": BLUE_PALE,
    "NLD": BLUE_PALE,
    "BEL": BLUE_PALE,
    "SWE": BLUE_PALE,
    "NOR": BLUE_PALE,
    "CHE": BLUE_PALE,
    "AUT": BLUE_PALE,
    "POL": BLUE_PALE,
    "IRL": BLUE_PALE,
    "PRT": BLUE_PALE,
    "DNK": BLUE_PALE,
    "FIN": BLUE_PALE,
    "CZE": BLUE_PALE,
    "GRC": BLUE_PALE,
    "ROU": BLUE_PALE,
    "HUN": BLUE_PALE,
    "ARE": BLUE_PALE,
    "SAU": BLUE_PALE,
    "QAT": BLUE_PALE,
    "BHR": BLUE_PALE,
    "KWT": BLUE_PALE,
    "OMN": BLUE_PALE,
    "ISR": BLUE_PALE,
    "JOR": BLUE_PALE,
}


def font(size, bold=False):
    path = (
        "/usr/share/fonts/truetype/liberation/LiberationSerif-Bold.ttf"
        if bold
        else "/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf"
    )
    sans = (
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"
        if bold
        else "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"
    )
    try:
        return ImageFont.truetype(path, size)
    except OSError:
        return ImageFont.truetype(sans, size)


def sans(size, bold=False):
    path = (
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"
        if bold
        else "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"
    )
    return ImageFont.truetype(path, size)


def load_geo():
    with open(os.path.join(ROOT, "maps/countries.geo.json")) as f:
        return json.load(f)


def iter_rings(geom):
    if geom["type"] == "Polygon":
        yield geom["coordinates"][0]
    elif geom["type"] == "MultiPolygon":
        for poly in geom["coordinates"]:
            yield poly[0]


def shift_lon(lon, center=0):
    x = lon - center
    while x < -180:
        x += 360
    while x > 180:
        x -= 360
    return x


def split_ring(ring, center=0):
    pts = [(shift_lon(lon, center), lat) for lon, lat in ring]
    segs, cur = [], [pts[0]]
    for i in range(1, len(pts)):
        if abs(pts[i][0] - pts[i - 1][0]) > 180:
            if len(cur) >= 3:
                segs.append(cur)
            cur = [pts[i]]
        else:
            cur.append(pts[i])
    if len(cur) >= 3:
        segs.append(cur)
    return segs


def great_circle(lat1, lon1, lat2, lon2, n=80):
    p1 = np.array(
        [
            math.cos(math.radians(lat1)) * math.cos(math.radians(lon1)),
            math.cos(math.radians(lat1)) * math.sin(math.radians(lon1)),
            math.sin(math.radians(lat1)),
        ]
    )
    p2 = np.array(
        [
            math.cos(math.radians(lat2)) * math.cos(math.radians(lon2)),
            math.cos(math.radians(lat2)) * math.sin(math.radians(lon2)),
            math.sin(math.radians(lat2)),
        ]
    )
    omega = math.acos(max(-1.0, min(1.0, float(np.dot(p1, p2)))))
    if omega < 1e-6:
        return [(lon1, lat1), (lon2, lat2)]
    pts = []
    for i in range(n + 1):
        t = i / n
        a = math.sin((1 - t) * omega) / math.sin(omega)
        b = math.sin(t * omega) / math.sin(omega)
        p = a * p1 + b * p2
        nrm = np.linalg.norm(p)
        p = p / nrm
        lat = math.degrees(math.asin(max(-1.0, min(1.0, p[2]))))
        lon = math.degrees(math.atan2(p[1], p[0]))
        pts.append((lon, lat))
    return pts


def collect_segments(geo, center=0, lon_lim=None, lat_lim=None, highlighted=False):
    segs = []
    fills = []
    for feat in geo["features"]:
        iso = feat.get("id")
        if iso in ("ATA", "ATF", "GRL"):
            continue
        is_hi = iso in HIGHLIGHT
        if highlighted and not is_hi:
            continue
        if not highlighted and is_hi:
            continue
        fill = HIGHLIGHT.get(iso, LAND)
        for ring in iter_rings(feat["geometry"]):
            for seg in split_ring(ring, center):
                xs = [p[0] for p in seg]
                ys = [p[1] for p in seg]
                if lon_lim and (max(xs) < lon_lim[0] - 8 or min(xs) > lon_lim[1] + 8):
                    continue
                if lat_lim and (max(ys) < lat_lim[0] - 8 or min(ys) > lat_lim[1] + 8):
                    continue
                segs.append(seg)
                fills.append(fill)
    return segs, fills


def ocean_backdrop(ax, xlim, ylim):
    """North–south ocean wash in data coordinates (avoids imshow + equal-aspect banding)."""
    n = 48
    y0, y1 = ylim
    c_north = np.array([236, 246, 251], dtype=float)
    c_south = np.array([186, 214, 230], dtype=float)
    span = xlim[1] - xlim[0]
    step = (y1 - y0) / n
    for i in range(n):
        t = (i + 0.5) / n
        rgb = (c_north * (1 - t) + c_south * t) / 255.0
        ax.add_patch(
            Rectangle(
                (xlim[0] - 2, y0 + i * step - 0.05),
                span + 4,
                step + 0.12,
                facecolor=rgb,
                edgecolor="none",
                linewidth=0,
                zorder=0,
            )
        )


def draw_world(ax, geo, center=0, lon_lim=None, lat_lim=None, lw=0.35, shadow=True):
    base, base_c = collect_segments(geo, center, lon_lim, lat_lim, highlighted=False)
    hi, hi_c = collect_segments(geo, center, lon_lim, lat_lim, highlighted=True)
    land = base + hi
    if shadow and land:
        ax.add_collection(
            PatchCollection(
                [MplPolygon([(p[0] + 0.85, p[1] - 0.65) for p in s], closed=True) for s in land],
                facecolor="#6E93A8",
                edgecolor="none",
                alpha=0.18,
                zorder=1,
            )
        )
    if base:
        ax.add_collection(
            PatchCollection(
                [MplPolygon(s, closed=True) for s in base],
                facecolor=base_c,
                edgecolor="#F7FBFD",
                linewidths=lw,
                antialiased=True,
                zorder=2,
            )
        )
    if hi:
        ax.add_collection(
            PatchCollection(
                [MplPolygon(s, closed=True) for s in hi],
                facecolor="none",
                edgecolor=GLOW,
                linewidths=lw + 3.8,
                alpha=0.48,
                antialiased=True,
                zorder=3,
            )
        )
        ax.add_collection(
            PatchCollection(
                [MplPolygon(s, closed=True) for s in hi],
                facecolor=hi_c,
                edgecolor="white",
                linewidths=lw + 0.25,
                antialiased=True,
                zorder=4,
            )
        )


def pill(ax, x, y, text, sub=None, ha="center", va="center"):
    label = text if not sub else f"{text}\n{sub}"
    ax.annotate(
        label,
        xy=(x, y),
        xytext=(0, 0),
        textcoords="offset points",
        ha=ha,
        va=va,
        fontsize=8.2 if not sub else 7.6,
        color=NAVY,
        fontweight="bold",
        fontfamily="serif",
        zorder=12,
        bbox=dict(boxstyle="round,pad=0.35", fc="white", ec=BLUE, lw=0.9, alpha=0.96),
        linespacing=1.25,
    )


def pin(ax, lon, lat, center=0):
    x, y = shift_lon(lon, center), lat
    ax.plot(x, y, "o", color=WHITE, markersize=8.5, zorder=10)
    ax.plot(x, y, "o", color=BLUE, markersize=5.4, zorder=11)
    ax.plot(x, y, "o", color=WHITE, markersize=2.2, zorder=12)
    return x, y


def world_map():
    geo = load_geo()
    xlim, ylim = (-168, 178), (-48, 74)
    fig, ax = plt.subplots(figsize=(14.6, 5.05), dpi=200)
    fig.patch.set_facecolor(OCEAN)
    ax.set_facecolor(OCEAN)
    ocean_backdrop(ax, xlim, ylim)
    for lon in range(-150, 181, 30):
        ax.plot([lon, lon], [ylim[0], ylim[1]], color="white", lw=0.45, alpha=0.18, zorder=0.5)
    for lat in range(-40, 71, 20):
        ax.plot([xlim[0], xlim[1]], [lat, lat], color="white", lw=0.45, alpha=0.18, zorder=0.5)
    draw_world(ax, geo, center=0, lw=0.42)

    pins = [
        (-74.0, 40.7),
        (-79.38, 43.65),
        (-0.13, 51.5),
        (103.8, 1.35),
        (114.17, 22.32),
        (121.47, 31.23),
        (151.2, -33.87),
        (55.27, 25.2),
    ]
    for lon, lat in pins:
        pin(ax, lon, lat)

    pill(ax, -97, 37, "United States", "First beachhead")
    pill(ax, -105, 59, "Canada")
    pill(ax, -89, 13, "Central America")
    pill(ax, -9, 62, "United Kingdom")
    pill(ax, 21, 45, "Europe", "Localized later")
    pill(ax, 47, 23, "Middle East")
    pill(ax, 103, 35, "China", "R&D and users")
    pill(ax, 108, 5, "Singapore")
    pill(ax, 129, 22, "Hong Kong")
    pill(ax, 134, -24, "Australia")

    ax.set_xlim(*xlim)
    ax.set_ylim(*ylim)
    ax.set_aspect("equal", adjustable="box")
    ax.axis("off")
    fig.subplots_adjust(0, 0, 1, 1)
    path = os.path.join(OUT, "world_map.png")
    fig.savefig(path, dpi=200, facecolor=fig.get_facecolor())
    plt.close(fig)
    return path


def shanghai_ny_map():
    geo = load_geo()
    center = 170
    xlim, ylim = (-68, 124), (10, 56)
    fig, ax = plt.subplots(figsize=(14.6, 4.15), dpi=200)
    fig.patch.set_facecolor(OCEAN)
    ax.set_facecolor(OCEAN)
    ocean_backdrop(ax, xlim, ylim)
    for lon in range(-90, 91, 30):
        ax.plot([lon, lon], [ylim[0], ylim[1]], color="white", lw=0.55, alpha=0.35, zorder=0.5)
    draw_world(ax, geo, center=center, lon_lim=(-90, 90), lat_lim=(-10, 72), lw=0.55)

    sh_lon, sh_lat = 121.47, 31.23
    ny_lon, ny_lat = -74.0, 40.71
    x0, y0 = shift_lon(sh_lon, center), sh_lat
    x1, y1 = shift_lon(ny_lon, center), ny_lat
    xc, yc = shift_lon(178.0, center), 44.0
    ts = np.linspace(0, 1, 90)
    xs = (1 - ts) ** 2 * x0 + 2 * (1 - ts) * ts * xc + ts ** 2 * x1
    ys = (1 - ts) ** 2 * y0 + 2 * (1 - ts) * ts * yc + ts ** 2 * y1
    ax.plot(xs, ys, color="#E7F4FB", linewidth=13.0, zorder=6, solid_capstyle="round")
    ax.plot(xs, ys, color=GLOW, linewidth=6.2, zorder=7, solid_capstyle="round")
    ax.plot(xs, ys, color=BLUE, linewidth=2.5, zorder=8, solid_capstyle="round")
    for i in (22, 45, 68):
        ax.annotate(
            "",
            xy=(xs[i + 3], ys[i + 3]),
            xytext=(xs[i], ys[i]),
            arrowprops=dict(arrowstyle="-|>", color=BLUE_DEEP, lw=1.25),
            zorder=9,
        )

    pin(ax, sh_lon, sh_lat, center)
    pin(ax, ny_lon, ny_lat, center)
    ax.annotate(
        "Shanghai\nR&D  ·  Product",
        xy=(x0, y0),
        xytext=(16, 14),
        textcoords="offset points",
        ha="left",
        va="bottom",
        fontsize=10,
        color=NAVY,
        fontweight="bold",
        fontfamily="serif",
        zorder=13,
        bbox=dict(boxstyle="round,pad=0.5", fc="white", ec=BLUE, lw=1.1),
        linespacing=1.35,
    )
    ax.annotate(
        "New York\nU.S. GTM",
        xy=(x1, y1),
        xytext=(-16, 14),
        textcoords="offset points",
        ha="right",
        va="bottom",
        fontsize=10,
        color=NAVY,
        fontweight="bold",
        fontfamily="serif",
        zorder=13,
        bbox=dict(boxstyle="round,pad=0.5", fc="white", ec=BLUE, lw=1.1),
        linespacing=1.35,
    )
    mid = 44
    ax.text(
        (x0 + x1) / 2,
        ys[mid] + 7.2,
        "China builds   →   the U.S. commercializes",
        ha="center",
        va="center",
        fontsize=10,
        color=BLUE_DEEP,
        fontfamily="serif",
        fontstyle="italic",
        zorder=12,
        bbox=dict(boxstyle="round,pad=0.45", fc="white", ec=BLUE, lw=0.7, alpha=0.97),
    )
    ax.text(
        (x0 + x1) / 2,
        18.4,
        "P A C I F I C    O C E A N",
        ha="center",
        va="center",
        fontsize=9.5,
        color=BLUE_DEEP,
        alpha=0.32,
        fontfamily="serif",
        fontstyle="italic",
        zorder=5,
    )
    ax.text(
        xs[mid],
        ys[mid] - 4.4,
        "Pacific route  ·  7,350 mi",
        ha="center",
        va="center",
        fontsize=8.4,
        color=BLUE_DEEP,
        fontfamily="serif",
        zorder=12,
        bbox=dict(boxstyle="round,pad=0.32", fc="white", ec=GLOW, lw=0.7, alpha=0.94),
    )

    ax.set_xlim(*xlim)
    ax.set_ylim(*ylim)
    ax.set_aspect("equal", adjustable="box")
    ax.axis("off")
    fig.subplots_adjust(0.0, 0.02, 1.0, 0.98)
    path = os.path.join(OUT, "shanghai_ny_map.png")
    fig.savefig(path, dpi=200, facecolor=fig.get_facecolor())
    plt.close(fig)
    return path


def rounded_rect(draw, box, radius, fill, outline=None, width=1):
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def round_png(path, radius=48):
    im = Image.open(path).convert("RGBA")
    mask = Image.new("L", im.size, 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, im.size[0] - 1, im.size[1] - 1), radius=radius, fill=255)
    im.putalpha(mask)
    im.save(path)


def resume_rewrite():
    """Full-bleed product screenshot — aspect matches the 11.53×5.62 slide frame."""
    W, H = 2400, 1170
    im = Image.new("RGB", (W, H), (243, 247, 250))
    d = ImageDraw.Draw(im)

    rounded_rect(d, (8, 8, W - 8, H - 8), 26, (255, 255, 255), (214, 224, 232), 2)
    d.rectangle((8, 8, W - 8, 72), fill=(243, 247, 250))
    for i, col in enumerate([(255, 95, 86), (255, 189, 46), (39, 201, 63)]):
        d.ellipse((28 + i * 24, 28, 46 + i * 24, 46), fill=col)
    d.text((112, 26), "51careers.ai  /  resume", font=sans(20), fill=(92, 107, 118))
    d.text((500, 22), "AI Resume Rewrite  ·  Software Engineer, New York", font=font(26, True), fill=(16, 34, 47))
    d.rounded_rectangle((1988, 22, 2360, 56), 10, fill=(0, 123, 203))
    d.text((2020, 28), "Illustrative output", font=sans(16, True), fill=(255, 255, 255))

    def chips(x, y, max_x, items):
        xx = x
        yy = y
        for t in items:
            bbox = d.textbbox((0, 0), t, font=sans(14, True))
            tw = bbox[2] - bbox[0]
            w = tw + 22
            if xx + w > max_x:
                xx = x
                yy += 34
            d.rounded_rectangle((xx, yy, xx + w, yy + 26), 8, fill=(232, 244, 251), outline=(169, 205, 224))
            d.text((xx + 11, yy + 4), t, font=sans(14, True), fill=(0, 123, 203))
            xx += w + 8
        return yy + 34

    def panel(x0, x1, y0, y1, tag, tag_fill, title, blocks):
        rounded_rect(d, (x0, y0, x1, y1), 16, (252, 253, 254), (226, 233, 238), 1)
        d.rounded_rectangle((x0 + 22, y0 + 16, x0 + 22 + 108, y0 + 46), 8, fill=tag_fill)
        d.text((x0 + 40, y0 + 21), tag, font=sans(15, True), fill=(255, 255, 255))
        d.text((x0 + 148, y0 + 18), title, font=font(24, True), fill=(16, 34, 47))
        yy = y0 + 62
        for kind, text, extra in blocks:
            if kind == "rule":
                d.line((x0 + 22, yy, x1 - 22, yy), fill=(230, 236, 241), width=1)
                yy += 12
            elif kind == "head":
                d.text((x0 + 22, yy), text, font=sans(26, True), fill=(16, 34, 47))
                yy += 32
            elif kind == "sub":
                d.text((x0 + 22, yy), text, font=sans(16), fill=(92, 107, 118))
                yy += 24
            elif kind == "sec":
                d.text((x0 + 22, yy), text, font=sans(13, True), fill=(0, 123, 203) if extra else (140, 151, 160))
                yy += 24
            elif kind == "job":
                d.text((x0 + 22, yy), text, font=sans(16, True), fill=(16, 34, 47))
                yy += 24
            elif kind == "chips":
                yy = chips(x0 + 22, yy, x1 - 22, extra)
            elif kind == "quote":
                d.rounded_rectangle((x0 + 22, yy, x1 - 22, yy + 70), 10, fill=(243, 247, 250), outline=(226, 233, 238))
                d.text((x0 + 36, yy + 10), text, font=sans(15), fill=(30, 42, 51))
                d.text((x0 + 36, yy + 36), extra, font=sans(15), fill=(0, 123, 203))
                yy += 82
            elif kind == "ghost":
                d.rounded_rectangle((x0 + 22, yy + 2, x1 - 70, yy + 16), 4, fill=(232, 237, 241))
                yy += 22
            elif kind == "dash":
                d.rounded_rectangle((x0 + 22, yy, x1 - 22, yy + 86), 12, outline=(214, 224, 232), width=2)
                d.text((x0 + 40, yy + 32), text, font=sans(14), fill=(140, 151, 160))
                yy += 100
            else:
                fill = (0, 123, 203) if extra else (30, 42, 51)
                prefix = "•  " if kind == "bullet" else ""
                d.text((x0 + 22, yy), prefix + text, font=sans(16, bold=bool(extra)), fill=fill)
                yy += 24
        return yy

    before = [
        ("head", "Alex Chen", False),
        ("sub", "Looking for software job in NY", False),
        ("rule", "", False),
        ("sec", "NOTES  ·  UNSTRUCTURED INPUT", False),
        ("body", "Startup — software person", False),
        ("body", "Worked at a startup doing coding", False),
        ("body", "Helped with the company website", False),
        ("body", "Worked with the team on projects", False),
        ("body", "Used python and javascript", False),
        ("body", "Did some cloud stuff on AWS", False),
        ("body", "Made the site faster I think", False),
        ("body", "Some intern work before that", False),
        ("rule", "", False),
        ("sec", "EDUCATION / SKILLS", False),
        ("body", "Graduated from a university", False),
        ("body", "Skills: coding, teamwork, python", False),
        ("body", "Want a software engineer role", False),
        ("rule", "", False),
        ("dash", "Paste more career facts — education, dates, metrics…", False),
        ("ghost", "", False),
        ("ghost", "", False),
        ("ghost", "", False),
        ("ghost", "", False),
        ("ghost", "", False),
        ("ghost", "", False),
        ("ghost", "", False),
        ("ghost", "", False),
    ]
    after = [
        ("head", "Alex Chen", False),
        ("sub", "Software Engineer  ·  New York, NY  ·  alex@email.com", False),
        ("rule", "", False),
        ("sec", "PROFESSIONAL SUMMARY", True),
        ("body", "Software engineer who ships customer-facing product and", False),
        ("body", "owns delivery metrics — 40K MAU, 35% faster loads, 28% fewer defects.", True),
        ("rule", "", False),
        ("sec", "EXPERIENCE", True),
        ("job", "Software Engineer, Northstar Labs  ·  2022–Present", False),
        ("bullet", "Engineered web platform serving 40,000 MAU; cut p95 load 35%.", True),
        ("bullet", "Shipped 12 releases with PM and design; reduced defect escape 28%.", True),
        ("bullet", "Owned AWS infrastructure for 4 production services.", True),
        ("job", "Engineering Intern, Harbor Labs  ·  2021", False),
        ("bullet", "Built internal tools that cut QA cycle time from 4 days to 1.5.", True),
        ("rule", "", False),
        ("sec", "EDUCATION / SKILLS", True),
        ("body", "B.S. Computer Science  ·  Python, TypeScript, AWS, PostgreSQL", False),
        ("chips", "", ["Python", "TypeScript", "AWS", "PostgreSQL", "ATS-ready", "New York market"]),
        ("rule", "", False),
        ("sec", "SELECTED PROJECT", True),
        ("job", "Customer metrics dashboard  ·  Northstar Labs", False),
        ("bullet", "Shipped a weekly dashboard used by 8 PMs to track 40K MAU.", True),
        ("rule", "", False),
        ("sec", "TARGET ROLE FIT", True),
        ("job", "Software Engineer  ·  New York  ·  ATS 94  ·  JD match 91", False),
        ("bullet", "Impact metrics, AWS ownership, and shipping cadence on the page.", True),
        ("bullet", "Language and structure matched to U.S. software-engineer JDs.", True),
        ("rule", "", False),
        ("sec", "COVER LETTER EXCERPT", True),
        (
            "quote",
            "I ship quantified product work — 40K MAU, 35% faster p95 —",
            "and I am targeting software-engineer roles in New York.",
        ),
        ("chips", "", ["Software Engineer · NY", "Backend", "Full-stack", "Cover letter"]),
    ]

    panel_top, panel_bot = 88, 1076
    panel(28, 1124, panel_top, panel_bot, "BEFORE", (140, 151, 160), "Candidate input", before)
    d.rounded_rectangle((1144, 548, 1256, 660), 56, fill=(0, 123, 203))
    d.polygon([(1178, 578), (1230, 604), (1178, 630)], fill=(255, 255, 255))
    y_after = panel(1276, 2372, panel_top, panel_bot, "AFTER", (0, 123, 203), "51 Careers.AI rewrite", after)

    stats = [("94", "ATS score"), ("91", "JD match"), ("12", "Metrics on page"), ("3", "Versions")]
    x0, x1 = 1276, 2372
    yy = min(max(y_after + 8, panel_bot - 78), panel_bot - 70)
    sw = (x1 - x0 - 44 - 24) / 4
    for i, (val, lab) in enumerate(stats):
        x = x0 + 22 + i * (sw + 8)
        d.rounded_rectangle((x, yy, x + sw, yy + 56), 10, fill=(232, 244, 251), outline=(169, 205, 224))
        d.text((x + 14, yy + 6), val, font=sans(20, True), fill=(0, 123, 203))
        d.text((x + 14, yy + 32), lab, font=sans(12), fill=(92, 107, 118))

    d.rectangle((8, 1090, W - 8, H - 8), fill=(243, 247, 250))
    d.text(
        (72, 1108),
        "AI Resume   ·   Optimization   ·   Job-specific versions   ·   Quantification   ·   ATS   ·   Cover letters   ·   Career profile",
        font=sans(16),
        fill=(92, 107, 118),
    )

    path = os.path.join(OUT, "resume_before_after.png")
    im.save(path, "PNG")
    return path


def circle_crop(im, size):
    im = im.convert("RGBA").resize((size, size), Image.Resampling.LANCZOS)
    mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(mask).ellipse((1, 1, size - 2, size - 2), fill=255)
    out = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    out.paste(im, (0, 0))
    out.putalpha(mask)
    return out


def headshots():
    people = [
        ("RC", "Rocky Chen", (10, 95, 155)),
        ("SL", "Stephanie Li", (14, 70, 120)),
        ("GD", "Gavin Ding", (0, 123, 203)),
        ("RZ", "Robin Zhu", (20, 88, 140)),
        ("CL", "Chris Lin", (8, 108, 168)),
        ("JS", "Jon Serbin", (16, 52, 86)),
    ]
    paths = {}
    size = 512
    for initials, name, rgb in people:
        im = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        d = ImageDraw.Draw(im)
        # outer ring
        d.ellipse((0, 0, size - 1, size - 1), fill=(255, 255, 255, 255))
        d.ellipse((22, 22, size - 23, size - 23), fill=rgb + (255,))
        overlay = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        od = ImageDraw.Draw(overlay)
        od.pieslice((22, 22, size - 23, size - 23), 200, 340, fill=(255, 255, 255, 32))
        im = Image.alpha_composite(im, overlay)
        d = ImageDraw.Draw(im)
        f = font(168, True)
        bbox = d.textbbox((0, 0), initials, font=f)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        d.text(((size - tw) / 2, (size - th) / 2 - 18), initials, font=f, fill=(255, 255, 255, 255))
        path = os.path.join(OUT, f"headshot_{initials.lower()}.png")
        im.save(path, "PNG")
        paths[name] = path
    return paths


def icon_base():
    im = Image.new("RGBA", (256, 256), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    return im, d


def stroke():
    return (0, 123, 203, 255)


def wline(d, pts, width=14):
    d.line(pts, fill=stroke(), width=width, joint="curve")
    # round caps
    r = width // 2
    for x, y in (pts[0:2], pts[-2:]):
        d.ellipse((x - r, y - r, x + r, y + r), fill=stroke())


def save_icon(im, name):
    path = os.path.join(OUT, f"icon_{name}.png")
    im.save(path, "PNG")
    return path


def make_icons():
    paths = {}
    # fragmented — three disconnected squares
    im, d = icon_base()
    for box in [(40, 48, 118, 126), (138, 40, 216, 118), (88, 138, 166, 216)]:
        d.rounded_rectangle(box, 16, outline=stroke(), width=12)
    paths["fragmented"] = save_icon(im, "fragmented")

    # expensive — clock + coin
    im, d = icon_base()
    d.ellipse((36, 36, 220, 220), outline=stroke(), width=12)
    d.line((128, 128, 128, 70), fill=stroke(), width=12)
    d.line((128, 128, 176, 148), fill=stroke(), width=12)
    paths["expensive"] = save_icon(im, "expensive")

    # low-signal — weak bars
    im, d = icon_base()
    bars = [(48, 168, 88, 208), (104, 128, 144, 208), (160, 88, 200, 208)]
    for i, box in enumerate(bars):
        fill = stroke() if i == 0 else (0, 123, 203, 70)
        d.rounded_rectangle(box, 8, fill=fill)
    d.line((52, 60, 204, 60), fill=stroke(), width=10)
    d.line((188, 44, 212, 68), fill=stroke(), width=10)
    d.line((188, 76, 212, 52), fill=stroke(), width=10)
    paths["lowsignal"] = save_icon(im, "lowsignal")

    # AI / spark
    im, d = icon_base()
    star = [(128, 24), (148, 108), (232, 128), (148, 148), (128, 232), (108, 148), (24, 128), (108, 108)]
    d.polygon(star, outline=stroke(), width=12)
    d.line((128, 24, 128, 232), fill=stroke(), width=12)
    d.line((24, 128, 232, 128), fill=stroke(), width=12)
    paths["ai"] = save_icon(im, "ai")

    # handshake / services — two users
    im, d = icon_base()
    d.ellipse((56, 40, 120, 104), outline=stroke(), width=12)
    d.arc((28, 112, 148, 232), 200, 340, fill=stroke(), width=12)
    d.ellipse((136, 40, 200, 104), outline=stroke(), width=12)
    d.arc((108, 112, 228, 232), 200, 340, fill=stroke(), width=12)
    paths["services"] = save_icon(im, "services")

    # referral / share
    im, d = icon_base()
    for c in [(64, 64, 108, 108), (148, 40, 192, 84), (148, 172, 192, 216)]:
        d.ellipse(c, outline=stroke(), width=12)
    d.line((108, 86, 148, 62), fill=stroke(), width=12)
    d.line((108, 96, 148, 194), fill=stroke(), width=12)
    paths["referral"] = save_icon(im, "referral")

    # building / B2B
    im, d = icon_base()
    d.rounded_rectangle((56, 48, 200, 220), 8, outline=stroke(), width=12)
    for x in (88, 128, 168):
        for y in (84, 124, 164):
            d.rounded_rectangle((x - 10, y - 10, x + 10, y + 10), 3, fill=stroke())
    paths["b2b"] = save_icon(im, "b2b")

    # megaphone
    im, d = icon_base()
    d.polygon([(40, 100), (120, 80), (120, 176), (40, 156)], outline=stroke())
    d.line(list(d.polygon([(40, 100), (120, 80), (120, 176), (40, 156)]) or []), fill=stroke(), width=1)
    d.polygon([(40, 100), (120, 80), (120, 176), (40, 156)], fill=(0, 123, 203, 40), outline=stroke(), width=12)
    d.polygon([(120, 80), (208, 48), (208, 208), (120, 176)], fill=(0, 123, 203, 40), outline=stroke(), width=12)
    d.arc((150, 88, 210, 168), -50, 50, fill=stroke(), width=12)
    paths["ads"] = save_icon(im, "ads")

    # sponsorship / badge
    im, d = icon_base()
    d.regular_polygon((128, 120, 88), 6, rotation=30, outline=stroke())
    try:
        d.regular_polygon((128, 120, 88), 6, rotation=30, outline=stroke())
    except Exception:
        pass
    d.ellipse((128 - 88, 120 - 88, 128 + 88, 120 + 88), outline=stroke(), width=12)
    d.polygon([(128, 64), (148, 108), (196, 112), (160, 144), (172, 190), (128, 166), (84, 190), (96, 144), (60, 112), (108, 108)], outline=stroke(), width=10)
    paths["sponsor"] = save_icon(im, "sponsor")

    # extract
    im, d = icon_base()
    d.rounded_rectangle((48, 40, 176, 216), 16, outline=stroke(), width=12)
    d.line((76, 88, 148, 88), fill=stroke(), width=12)
    d.line((76, 128, 148, 128), fill=stroke(), width=12)
    d.line((76, 168, 120, 168), fill=stroke(), width=12)
    d.polygon([(168, 148), (228, 128), (228, 168)], fill=stroke())
    paths["extract"] = save_icon(im, "extract")

    # globe
    im, d = icon_base()
    d.ellipse((36, 36, 220, 220), outline=stroke(), width=12)
    d.ellipse((86, 36, 170, 220), outline=stroke(), width=10)
    d.arc((36, 86, 220, 130), 0, 180, fill=stroke(), width=10)
    d.arc((36, 126, 220, 170), 180, 360, fill=stroke(), width=10)
    paths["globe"] = save_icon(im, "globe")

    # industry / grid
    im, d = icon_base()
    for i in range(2):
        for j in range(2):
            x, y = 48 + i * 92, 48 + j * 92
            d.rounded_rectangle((x, y, x + 68, y + 68), 12, outline=stroke(), width=12)
    paths["industry"] = save_icon(im, "industry")

    # levels
    im, d = icon_base()
    d.polygon([(128, 40), (208, 200), (48, 200)], outline=stroke(), width=12)
    d.line((80, 148, 176, 148), fill=stroke(), width=10)
    paths["levels"] = save_icon(im, "levels")

    # role / person
    im, d = icon_base()
    d.ellipse((88, 40, 168, 120), outline=stroke(), width=12)
    d.arc((48, 128, 208, 248), 200, 340, fill=stroke(), width=12)
    paths["role"] = save_icon(im, "role")

    # employer / office
    im, d = icon_base()
    d.rectangle((48, 72, 208, 220), outline=stroke(), width=12)
    d.polygon([(48, 72), (128, 28), (208, 72)], outline=stroke(), width=12)
    d.rectangle((112, 160, 144, 220), outline=stroke(), width=10)
    paths["employer"] = save_icon(im, "employer")

    # search
    im, d = icon_base()
    d.ellipse((44, 44, 176, 176), outline=stroke(), width=12)
    d.line((160, 160, 216, 216), fill=stroke(), width=16)
    paths["search"] = save_icon(im, "search")

    # university
    im, d = icon_base()
    d.polygon([(40, 100), (128, 52), (216, 100)], outline=stroke(), width=12)
    d.line((128, 52, 128, 200), fill=stroke(), width=10)
    d.line((64, 108, 64, 176), fill=stroke(), width=10)
    d.line((192, 108, 192, 176), fill=stroke(), width=10)
    d.line((40, 176, 216, 176), fill=stroke(), width=12)
    paths["university"] = save_icon(im, "university")

    # scoring
    im, d = icon_base()
    d.line((56, 200, 56, 80), fill=stroke(), width=12)
    d.line((56, 200, 208, 200), fill=stroke(), width=12)
    d.rectangle((80, 140, 112, 200), fill=stroke())
    d.rectangle((124, 100, 156, 200), fill=stroke())
    d.rectangle((168, 64, 200, 200), fill=stroke())
    paths["scoring"] = save_icon(im, "scoring")

    # jd / document
    im, d = icon_base()
    d.rounded_rectangle((64, 32, 192, 224), 16, outline=stroke(), width=12)
    d.line((88, 80, 168, 80), fill=stroke(), width=10)
    d.line((88, 116, 168, 116), fill=stroke(), width=10)
    d.line((88, 152, 140, 152), fill=stroke(), width=10)
    paths["jd"] = save_icon(im, "jd")

    # resume page
    im, d = icon_base()
    d.rounded_rectangle((60, 28, 196, 228), 12, outline=stroke(), width=12)
    d.ellipse((96, 52, 160, 116), outline=stroke(), width=10)
    d.line((84, 140, 172, 140), fill=stroke(), width=10)
    d.line((84, 168, 172, 168), fill=stroke(), width=10)
    d.line((84, 196, 140, 196), fill=stroke(), width=10)
    paths["resume"] = save_icon(im, "resume")

    # terms
    im, d = icon_base()
    d.rounded_rectangle((40, 72, 216, 184), 20, outline=stroke(), width=12)
    d.text((78, 100), "ATS", font=sans(48, True), fill=stroke())
    paths["terms"] = save_icon(im, "terms")

    # geo pin
    im, d = icon_base()
    d.ellipse((72, 40, 184, 152), outline=stroke(), width=12)
    d.polygon([(92, 130), (128, 220), (164, 130)], outline=stroke(), width=12)
    d.ellipse((108, 76, 148, 116), fill=stroke())
    paths["geo"] = save_icon(im, "geo")

    # ats check
    im, d = icon_base()
    d.rounded_rectangle((48, 48, 208, 208), 24, outline=stroke(), width=12)
    d.line((84, 132, 118, 168), fill=stroke(), width=16)
    d.line((118, 168, 176, 96), fill=stroke(), width=16)
    paths["ats"] = save_icon(im, "ats")

    return paths


def tiktok_app_icon():
    """Official-style TikTok mark: white note on black rounded square."""
    src = Image.open(os.path.join(ROOT, "logos/png/tiktok.png")).convert("RGBA")
    size = 256
    im = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    d.rounded_rectangle((0, 0, size - 1, size - 1), 56, fill=(17, 17, 17, 255))
    note = src.resize((170, 170), Image.Resampling.LANCZOS)
    im.alpha_composite(note, (43, 43))
    path = os.path.join(OUT, "logo_tiktok.png")
    im.save(path, "PNG")
    return path


def copy_logos():
    mapping = {
        "google": os.path.join(ROOT, "logos/png/google_g.png"),
        "linkedin": os.path.join(ROOT, "logos/png/linkedin.png"),
        "youtube": os.path.join(ROOT, "logos/png/youtube.png"),
        "meta": os.path.join(ROOT, "logos/png/meta_mark.png"),
    }
    out = {}
    for name, src in mapping.items():
        im = Image.open(src).convert("RGBA")
        path = os.path.join(OUT, f"logo_{name}.png")
        im.save(path, "PNG")
        out[name] = path
    out["tiktok"] = tiktok_app_icon()
    return out


def main():
    print("maps...")
    wm = world_map()
    sm = shanghai_ny_map()
    round_png(wm, 36)
    round_png(sm, 28)
    print(" ", wm)
    print(" ", sm)
    print("resume...")
    print(" ", resume_rewrite())
    print("headshots...")
    for k, v in headshots().items():
        print(" ", k, v)
    print("icons...")
    icons = make_icons()
    print(" ", len(icons), "icons")
    print("logos...")
    logos = copy_logos()
    print(" ", logos)
    print("done")


if __name__ == "__main__":
    main()
