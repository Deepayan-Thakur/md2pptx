"""
icons.py — Programmatically generated SVG icons as PNG bytes.
All icons are 100% original — no external images, no copyright issues.
"""
import io
import math
from typing import Tuple
from PIL import Image, ImageDraw, ImageFont


# Accenture palette
RED    = (239, 68,  68)
BLUE   = (15,  158, 213)
GREEN  = (25,  107, 36)
ORANGE = (233, 113, 50)
PURPLE = (160, 43,  147)
DARK   = (44,  44,  44)
WHITE  = (255, 255, 255)
LIGHT  = (232, 232, 232)

ACCENT_COLORS = [RED, BLUE, GREEN, ORANGE, PURPLE]


def _new_canvas(size: int = 120, bg: Tuple = None) -> Tuple[Image.Image, ImageDraw.ImageDraw]:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0) if bg is None else (*bg, 255))
    draw = ImageDraw.Draw(img)
    return img, draw


def _to_bytes(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf.read()


def circle_icon(symbol: str, color: Tuple = RED, size: int = 120) -> bytes:
    """A colored circle with a letter/number inside."""
    img, draw = _new_canvas(size)
    r = size // 2 - 4
    cx, cy = size // 2, size // 2
    draw.ellipse([cx - r, cy - r, cx + r, cy + r], fill=(*color, 255))
    try:
        font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
                                   int(size * 0.4))
    except Exception:
        font = ImageFont.load_default()
    bbox = draw.textbbox((0, 0), symbol, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    draw.text((cx - tw // 2, cy - th // 2 - 2), symbol, fill=WHITE, font=font)
    return _to_bytes(img)


def bar_icon(size: int = 120, color: Tuple = RED) -> bytes:
    """Upward bar chart icon."""
    img, draw = _new_canvas(size)
    bars = [0.4, 0.65, 0.9, 0.75]
    w = size // (len(bars) * 2 + 1)
    gap = w
    for i, h in enumerate(bars):
        x0 = gap + i * (w + gap)
        y1 = size - 8
        y0 = int(size - 8 - h * (size - 16))
        c = tuple(min(255, int(c * (0.7 + i * 0.1))) for c in color) + (255,)
        draw.rectangle([x0, y0, x0 + w, y1], fill=c, outline=None)
    return _to_bytes(img)


def pie_icon(size: int = 120) -> bytes:
    """Pie chart icon."""
    img, draw = _new_canvas(size)
    cx, cy, r = size // 2, size // 2, size // 2 - 6
    slices = [(45, RED), (90, BLUE), (110, GREEN), (115, ORANGE)]
    start = -90
    for degrees, color in slices:
        draw.pieslice([cx - r, cy - r, cx + r, cy + r],
                      start=start, end=start + degrees,
                      fill=(*color, 255), outline=WHITE)
        start += degrees
    return _to_bytes(img)


def arrow_up_icon(size: int = 120, color: Tuple = GREEN) -> bytes:
    """Upward arrow — growth indicator."""
    img, draw = _new_canvas(size)
    cx = size // 2
    points = [
        (cx, 8),
        (size - 12, size // 2),
        (cx + 12, size // 2),
        (cx + 12, size - 8),
        (cx - 12, size - 8),
        (cx - 12, size // 2),
        (12, size // 2),
    ]
    draw.polygon(points, fill=(*color, 255))
    return _to_bytes(img)


def shield_icon(size: int = 120, color: Tuple = BLUE) -> bytes:
    """Shield — security/protection."""
    img, draw = _new_canvas(size)
    cx, s = size // 2, size
    pts = [
        (cx, 6), (s - 10, 20), (s - 10, s * 0.6), (cx, s - 8), (10, s * 0.6), (10, 20)
    ]
    draw.polygon(pts, fill=(*color, 200))
    # checkmark
    ck = [(cx - 16, s * 0.5), (cx - 4, s * 0.62), (cx + 18, s * 0.32)]
    draw.line(ck, fill=WHITE, width=max(4, size // 20))
    return _to_bytes(img)


def gear_icon(size: int = 120, color: Tuple = DARK) -> bytes:
    """Gear — technology/operations."""
    img, draw = _new_canvas(size)
    cx, cy = size // 2, size // 2
    R, r_in, teeth = size // 2 - 6, size // 4, 8
    pts = []
    for i in range(teeth * 4):
        angle = math.pi * 2 * i / (teeth * 4) - math.pi / 2
        rad = R if (i % 4 < 2) else R - size // 8
        pts.append((cx + rad * math.cos(angle), cy + rad * math.sin(angle)))
    draw.polygon(pts, fill=(*color, 220))
    draw.ellipse([cx - r_in, cy - r_in, cx + r_in, cy + r_in], fill=(0, 0, 0, 0))
    return _to_bytes(img)


def globe_icon(size: int = 120, color: Tuple = BLUE) -> bytes:
    """Globe — global/geography."""
    img, draw = _new_canvas(size)
    cx, cy, r = size // 2, size // 2, size // 2 - 6
    draw.ellipse([cx - r, cy - r, cx + r, cy + r], outline=(*color, 255), width=3)
    # meridians
    draw.arc([cx - r // 2, cy - r, cx + r // 2, cy + r], 0, 360, fill=(*color, 180), width=2)
    draw.line([(cx - r, cy), (cx + r, cy)], fill=(*color, 180), width=2)
    draw.line([(cx, cy - r), (cx, cy + r)], fill=(*color, 180), width=2)
    return _to_bytes(img)


def people_icon(size: int = 120, color: Tuple = PURPLE) -> bytes:
    """People icon — workforce/talent."""
    img, draw = _new_canvas(size)
    for ox, oy, sc in [(size * 0.28, 0, 0.75), (size * 0.65, 0, 0.75), (size * 0.5, -4, 1.0)]:
        head_r = int(size * 0.13 * sc)
        hx, hy = int(ox), int(oy + size * 0.22 * sc)
        draw.ellipse([hx - head_r, hy - head_r, hx + head_r, hy + head_r],
                     fill=(*color, 200))
        body_w = int(size * 0.18 * sc)
        body_top = hy + head_r + 2
        body_bot = int(oy + size * 0.75 * sc)
        draw.ellipse([hx - body_w, body_top, hx + body_w, body_bot],
                     fill=(*color, 180))
    return _to_bytes(img)


def lightbulb_icon(size: int = 120, color: Tuple = ORANGE) -> bytes:
    """Lightbulb — innovation/ideas."""
    img, draw = _new_canvas(size)
    cx, top = size // 2, 8
    r = int(size * 0.35)
    draw.ellipse([cx - r, top, cx + r, top + 2 * r], fill=(*color, 230))
    # base
    bw, bh = int(size * 0.18), int(size * 0.22)
    bx, by = cx - bw, top + 2 * r - 4
    draw.rectangle([bx, by, bx + 2 * bw, by + bh], fill=(*color, 200))
    # rays
    for angle in range(0, 360, 45):
        rad = math.radians(angle)
        x1 = cx + (r + 4) * math.cos(rad)
        y1 = top + r + (r + 4) * math.sin(rad)
        x2 = cx + (r + 14) * math.cos(rad)
        y2 = top + r + (r + 14) * math.sin(rad)
        draw.line([(x1, y1), (x2, y2)], fill=(*color, 180), width=2)
    return _to_bytes(img)


def scale_icon(size: int = 120, color: Tuple = BLUE) -> bytes:
    """Expansion/Scaling icon."""
    img, draw = _new_canvas(size)
    cx, cy = size // 2, size // 2
    # Multiple overlapping rectangles growing
    for i in range(3):
        s = int(size * (0.3 + i * 0.2))
        x0, y0 = cx - s // 2, cy - s // 2
        draw.rectangle([x0, y0, x0 + s, y0 + s], outline=(*color, 200 - i * 50), width=2)
    # arrow up right
    draw.line([(cx, cy), (size - 10, 10)], fill=(*color, 255), width=4)
    draw.polygon([(size - 10, 10), (size - 25, 10), (size - 10, 25)], fill=(*color, 255))
    return _to_bytes(img)


def roi_icon(size: int = 120, color: Tuple = GREEN) -> bytes:
    """Return on Investment / Dollar growth."""
    img, draw = _new_canvas(size)
    cx, cy = size // 2, size // 2
    # Dollar sign
    try:
        font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", int(size * 0.6))
    except:
        font = ImageFont.load_default()
    draw.text((cx - size // 4, cy - size // 3), "$", fill=(*color, 255), font=font)
    # Circular arrow around it
    draw.arc([10, 10, size - 10, size - 10], start=0, end=270, fill=(*color, 180), width=4)
    draw.polygon([(size - 10, cy), (size - 25, cy - 10), (size - 25, cy + 10)], fill=(*color, 180))
    return _to_bytes(img)


# Icon registry by keyword
_ICON_MAP = {
    "ai": lambda: gear_icon(color=BLUE),
    "cyber": lambda: shield_icon(color=BLUE),
    "security": lambda: shield_icon(color=BLUE),
    "cloud": lambda: globe_icon(color=BLUE),
    "global": lambda: globe_icon(color=BLUE),
    "geographic": lambda: globe_icon(color=BLUE),
    "growth": lambda: arrow_up_icon(color=GREEN),
    "revenue": lambda: bar_icon(color=GREEN),
    "financial": lambda: bar_icon(color=GREEN),
    "talent": lambda: people_icon(color=PURPLE),
    "learning": lambda: people_icon(color=PURPLE),
    "workforce": lambda: people_icon(color=PURPLE),
    "innovation": lambda: lightbulb_icon(color=ORANGE),
    "strategy": lambda: lightbulb_icon(color=RED),
    "data": lambda: bar_icon(color=BLUE),
    "chart": lambda: bar_icon(color=RED),
    "acquisition": lambda: arrow_up_icon(color=RED),
    "roi": lambda: roi_icon(color=GREEN),
    "scaling": lambda: scale_icon(color=BLUE),
    "efficiency": lambda: gear_icon(color=ORANGE),
    "conclusion": lambda: circle_icon("✓", GREEN),
}


def get_icon_for_title(title: str, index: int = 0) -> bytes:
    """Return an appropriate icon based on slide title keywords."""
    title_lower = title.lower()
    for keyword, factory in _ICON_MAP.items():
        if keyword in title_lower:
            return factory()
    # Fallback: numbered circle
    return circle_icon(str((index % 9) + 1), ACCENT_COLORS[index % len(ACCENT_COLORS)])


def get_numbered_icon(num: int, color: Tuple = None) -> bytes:
    c = color or ACCENT_COLORS[(num - 1) % len(ACCENT_COLORS)]
    return circle_icon(str(num), c)