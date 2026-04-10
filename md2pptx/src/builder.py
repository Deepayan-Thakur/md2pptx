"""
builder.py — PPTX construction engine.
Uses the provided Accenture Slide Master template as the base for every slide.
All layout/styling is driven by the template's theme (fonts, colors, backgrounds).
"""
import io
import os
from typing import Dict, Any, List, Optional, Tuple
from copy import deepcopy
from PIL import Image

from pptx import Presentation
from pptx.util import Emu, Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
import lxml.etree as etree

from .parser import clean_text, Table
from .charts import render_chart, stat_card_image, bar_chart
from .icons import get_icon_for_title, get_numbered_icon, ACCENT_COLORS
from .image_gen import generate_slide_asset

# ── Template path ─────────────────────────────────────────────────────────────
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "..", "assets", "template.pptx")

# ── Slide geometry (Template: 12192000 × 6858000 EMU - 16:9 ratio) ──────────────
SW = 12192000   # slide width (from template)
SH = 6858000    # slide height (from template)
M  = 350000     # left/right margin (~0.35 in)
TM = 600000     # top margin
TH = 530000     # title height
CT = 1350000    # content area top
CH = SH - CT - 320000  # content area height

# ── Brand colors (from theme1.xml) ────────────────────────────────────────────
C = {
    "red":    RGBColor(0xEF, 0x44, 0x44),
    "blue":   RGBColor(0x0F, 0x9E, 0xD5),
    "green":  RGBColor(0x19, 0x6B, 0x24),
    "orange": RGBColor(0xE9, 0x71, 0x32),
    "purple": RGBColor(0xA0, 0x2B, 0x93),
    "dark":   RGBColor(0x2C, 0x2C, 0x2C),
    "mid":    RGBColor(0x66, 0x66, 0x66),
    "light":  RGBColor(0xE8, 0xE8, 0xE8),
    "white":  RGBColor(0xFF, 0xFF, 0xFF),
}
ACCENT_RGB = [C["red"], C["blue"], C["green"], C["orange"], C["purple"]]

FONT_H = "Libre Baskerville"   # header
FONT_B = "Inter"               # body


# ─────────────────────────────────────────────────────────────────────────────
#  Low-level helpers
# ─────────────────────────────────────────────────────────────────────────────

def _rgb(r, g, b) -> RGBColor:
    return RGBColor(r, g, b)


def _para_props(p, space_before_pt: int = 0, space_after_pt: int = 0,
                line_spacing_pt: int = 0):
    """Set paragraph spacing on a paragraph element."""
    pPr = p._pPr
    if pPr is None:
        pPr = p._p.get_or_add_pPr()
    if space_before_pt:
        sb = etree.SubElement(pPr, qn("a:spcBef"))
        etree.SubElement(sb, qn("a:spcPts")).set("val", str(space_before_pt * 100))
    if space_after_pt:
        sa = etree.SubElement(pPr, qn("a:spcAft"))
        etree.SubElement(sa, qn("a:spcPts")).set("val", str(space_after_pt * 100))


def _add_run(p, text: str, font_name: str, size_pt: float, bold: bool = False,
             italic: bool = False, color: Optional[RGBColor] = None):
    run = p.add_run()
    run.text = text
    rf = run.font
    rf.name = font_name
    rf.size = Pt(size_pt)
    rf.bold = bold
    rf.italic = italic
    if color:
        rf.color.rgb = color
    return run


def _add_textbox(slide, text: str, left, top, width, height,
                 font_name: str = FONT_B, size_pt: float = 14,
                 bold: bool = False, italic: bool = False,
                 color: RGBColor = None, align: PP_ALIGN = PP_ALIGN.LEFT,
                 word_wrap: bool = True, space_before: int = 0) -> Any:
    color = color or C["dark"]
    txb = slide.shapes.add_textbox(Emu(left), Emu(top), Emu(width), Emu(height))
    tf = txb.text_frame
    tf.word_wrap = word_wrap
    p = tf.paragraphs[0]
    p.alignment = align
    if space_before:
        _para_props(p, space_before_pt=space_before)
    _add_run(p, clean_text(text), font_name, size_pt, bold, italic, color)
    return txb


def _add_rect(slide, left, top, width, height, fill_rgb: RGBColor,
              no_line: bool = True) -> Any:
    shape = slide.shapes.add_shape(1, Emu(left), Emu(top), Emu(width), Emu(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_rgb
    if no_line:
        shape.line.fill.background()
    return shape


def _add_image(slide, img_bytes: bytes, left, top, width, height) -> Any:
    """Add image while preserving aspect ratio and fitting within bounds."""
    if not img_bytes:
        return None
    stream = io.BytesIO(img_bytes)
    with Image.open(stream) as img:
        w, h = img.size
    ratio = min(width / w, height / h)
    final_w = int(w * ratio)
    final_h = int(h * ratio)
    final_left = left + (width - final_w) // 2
    final_top = top + (height - final_h) // 2
    stream.seek(0)
    return slide.shapes.add_picture(stream, Emu(final_left), Emu(final_top), width=Emu(final_w), height=Emu(final_h))


def _add_title(slide, title_text: str, left=M, top=TM, width=SW - 2*M, height=TH,
               color: RGBColor = None, size_pt: float = 32) -> Any:
    title_text = clean_text(title_text)  # Allow full text with word wrapping
    
    # Try placeholder first
    if slide.shapes.title:
        slide.shapes.title.text = title_text
        return slide.shapes.title

    color = color or C["dark"]
    txb = slide.shapes.add_textbox(Emu(left), Emu(top), Emu(width), Emu(height))
    tf = txb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT  # Left-align for professional look
    _add_run(p, title_text, FONT_H, size_pt, bold=True, color=color)
    # Red underbar - prominent width for visual consistency
    bar_h = int(SH * 0.008)
    bar_top = top + height + int(SH * 0.006)
    _add_rect(slide, left, bar_top, int(SW * 0.18), bar_h, C["red"])  # 18% width, 8% thickness
    return txb


def _add_bullets(slide, points: List[str], left, top, width, height,
                 size_pt: float = 14, bold: bool = False,
                 color: RGBColor = None, numbered: bool = False,
                 icon_bytes: Optional[List[bytes]] = None) -> None:
    color = color or C["dark"]
    if not points:
        return
        
    n = min(len(points), 6)  # Support up to 6 custom boxes
    box_gap = int(SH * 0.018)
    box_h = (height - box_gap * max(0, n - 1)) // max(n, 1)

    for i, point in enumerate(points[:n]):
        by = top + i * (box_h + box_gap)
        accent = ACCENT_RGB[i % len(ACCENT_RGB)]
        
        # Soft background box
        bg_rgb = _rgb(245, 245, 247)
        _add_rect(slide, left, by, width, box_h, bg_rgb)
        
        # Left colored accent stripe
        accent_w = int(SW * 0.008)
        _add_rect(slide, left, by, accent_w, box_h, accent)
        
        # Text inside the creative box
        tx = left + accent_w + 50000
        tw = width - accent_w - 70000
        _add_textbox(slide, clean_text(point),
                     left=tx, top=by, width=tw, height=box_h,
                     font_name=FONT_B, size_pt=size_pt, color=color,
                     align=PP_ALIGN.LEFT)


def _add_icon(slide, img_bytes: bytes, cx: int, cy: int, size: int = 80000) -> None:
    """Place a square icon centered at (cx, cy)."""
    if not img_bytes:
        return
    _add_image(slide, img_bytes, cx - size // 2, cy - size // 2, size, size)


def _add_table(slide, tbl: Table, left, top, width, height) -> None:
    if not tbl or not tbl.rows:
        return
    headers = tbl.headers[:6]
    rows = tbl.rows[:10]
    col_count = len(headers)
    row_count = len(rows) + 1
    row_h = max(int(height / row_count), 400000)

    pptx_tbl = slide.shapes.add_table(
        row_count, col_count,
        Emu(left), Emu(top), Emu(width), Emu(row_h * row_count)
    ).table

    for ci, header in enumerate(headers):
        cell = pptx_tbl.cell(0, ci)
        cell.text = clean_text(header)[:40]
        cell.fill.solid()
        cell.fill.fore_color.rgb = C["red"]
        for p in cell.text_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER
            for run in p.runs:
                run.font.name = FONT_B
                run.font.size = Pt(10)
                run.font.bold = True
                run.font.color.rgb = C["white"]

    for ri, row in enumerate(rows):
        bg = C["light"] if ri % 2 == 0 else C["white"]
        for ci, val in enumerate(row[:col_count]):
            cell = pptx_tbl.cell(ri + 1, ci)
            cell.text = clean_text(str(val))[:50]
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER
                for run in p.runs:
                    run.font.name = FONT_B
                    run.font.size = Pt(10)
                    run.font.color.rgb = C["dark"]


# ─────────────────────────────────────────────────────────────────────────────
#  Slide builders
# ─────────────────────────────────────────────────────────────────────────────

def _build_title_slide(slide, sp: Dict, parsed: Dict) -> None:
    title = clean_text(parsed.get("title", sp.get("title", "")))
    subtitle = clean_text(parsed.get("subtitle", sp.get("subtitle", "")))

    # Try to use slide's title placeholder, fallback to custom textbox
    try:
        if slide.shapes.title:
            slide.shapes.title.text = title
        else:
            raise AttributeError("No title placeholder")
    except:
        # Fallback: custom textbox
        _add_textbox(slide, title,
                     left=M, top=int(SH * 0.30), width=SW - 2 * M, height=int(SH * 0.18),
                     font_name=FONT_H, size_pt=44, bold=True, color=C["dark"],
                     align=PP_ALIGN.CENTER)
    
    # Add subtitle
    if subtitle:
        try:
            if len(slide.placeholders) > 1:
                try:
                    subtitle_shape = slide.placeholders[1]
                    if subtitle_shape and hasattr(subtitle_shape, "text_frame"):
                        subtitle_shape.text = subtitle
                    else:
                        raise AttributeError("No text_frame")
                except (IndexError, AttributeError, KeyError):
                    raise AttributeError("No subtitle placeholder")
            else:
                raise AttributeError("No subtitle placeholder")
        except:
            # Fallback: custom textbox
            _add_textbox(slide, subtitle,
                         left=M, top=int(SH * 0.50), width=SW - 2 * M, height=int(SH * 0.12),
                         font_name=FONT_B, size_pt=18, color=C["mid"],
                         align=PP_ALIGN.CENTER)


def _build_agenda_slide(slide, sp: Dict) -> None:
    _add_title(slide, sp.get("title", "Agenda"))
    bullets = sp.get("bullets", [])
    if not bullets:
        return
    n = len(bullets)
    col_count = 2 if n >= 4 else 1
    col_w = (SW - 2 * M - (col_count - 1) * int(M * 1.2)) // col_count

    for ci in range(col_count):
        chunk = bullets[ci::col_count]
        lx = M + ci * (col_w + int(M * 1.2))
        item_h_each = (CH - int(max(len(chunk) - 1, 0) * 45000)) // max(len(chunk), 1)

        for ri, item in enumerate(chunk):
            item_top = CT + ri * item_h_each
            item_h = item_h_each - 45000
            num_size = min(int(SH * 0.058), int(item_h * 0.82))

            # Colored number box
            bg_rgb = ACCENT_RGB[(ci * len(bullets) // col_count + ri) % len(ACCENT_RGB)]
            _add_rect(slide, lx, item_top, num_size, item_h, bg_rgb)

            # Centered number
            num_txb = slide.shapes.add_textbox(
                Emu(lx), Emu(item_top), Emu(num_size), Emu(item_h))
            num_txb.text_frame.word_wrap = False
            p = num_txb.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            _add_run(p, str(ci * len(bullets) // col_count + ri + 1),
                     FONT_H, 18, bold=True, color=C["white"])

            # Item text (right of number box)
            txt_x = lx + num_size + 70000
            txt_w = col_w - num_size - 90000
            _add_textbox(slide, clean_text(item),
                         txt_x, item_top, txt_w, item_h,
                         font_name=FONT_B, size_pt=14, color=C["dark"])


def _build_exec_summary(slide, sp: Dict, parsed: Dict) -> None:
    _add_title(slide, "Executive Summary")
    text = clean_text(parsed.get("executive_summary", " ".join(sp.get("bullets", []))))

    # Split into sentences for bullets
    import re
    sentences = re.split(r"(?<=[.!?])\s+", text)
    bullets = [s for s in sentences if len(s) > 15][:5]

    icon_size = int(SH * 0.068)
    icon_gap = 30000
    item_h = (CH - icon_gap * max(0, len(bullets) - 1)) // max(len(bullets), 1)
    
    for i, bullet in enumerate(bullets):
        by = CT + i * (item_h + icon_gap)
        accent = ACCENT_RGB[i % len(ACCENT_RGB)]
        
        # Creative Box Background
        _add_rect(slide, M, by, SW - 2 * M, item_h, _rgb(245, 245, 247))
        _add_rect(slide, M, by, int(SW * 0.008), item_h, accent)
        
        icon_bytes = get_numbered_icon(i + 1)
        _add_icon(slide, icon_bytes,
                  cx=M + icon_size // 2 + 100000,
                  cy=by + item_h // 2,
                  size=icon_size)
        _add_textbox(slide, bullet,
                     left=M + icon_size + 180000, top=by,
                     width=SW - 2 * M - icon_size - 220000,
                     height=item_h,
                     font_name=FONT_B, size_pt=14, color=C["dark"])


def _build_section_divider(slide, sp: Dict) -> None:
    title = clean_text(sp.get("title", ""))
    # Colored background strip (centered vertically) - smaller to not overwhelm
    divider_h = int(SH * 0.26)
    divider_top = int((SH - divider_h) / 2) + int(SH * 0.02)
    _add_rect(slide, 0, divider_top, SW, divider_h, C["red"])
    
    # Centered title - balanced size
    _add_textbox(slide, title,
                 left=M, top=divider_top, width=SW - 2 * M, height=divider_h,
                 font_name=FONT_H, size_pt=34, bold=True, color=C["white"],
                 align=PP_ALIGN.CENTER)
    
    subtitle = clean_text(sp.get("subtitle", ""))
    if subtitle:
        subtitle_top = divider_top + divider_h + int(SH * 0.04)
        _add_textbox(slide, subtitle,
                     left=M, top=subtitle_top,
                     width=SW - 2 * M, height=int(SH * 0.10),
                     font_name=FONT_B, size_pt=14, color=C["dark"],
                     align=PP_ALIGN.CENTER)


def _build_content_slide(slide, sp: Dict, parsed: Dict) -> None:
    _add_title(slide, sp.get("title", ""))
    bullets = sp.get("bullets", [])
    if not bullets:
        return
    
    # Icon beside title
    icon_bytes = get_icon_for_title(sp.get("title", ""),
                                    index=sp.get("slide_number", 1) - 1)
    icon_size = int(SH * 0.072)
    _add_icon(slide, icon_bytes,
              cx=SW - M - icon_size // 2 - 100000,
              cy=TM + TH // 2,
              size=icon_size)

    # Place bullets on left ~58%
    text_width = int((SW - 2 * M) * 0.58)
    _add_bullets(slide, bullets[:6],
                 left=M, top=CT, width=text_width, height=CH,
                 size_pt=15, color=C["dark"])
                 
    # Fill right-side gap with dynamic AI visual asset!
    title = sp.get("title", "Corporate Slide")
    doc_title = parsed.get("title", "Corporate Strategy")
    
    bg_bytes = generate_slide_asset(title, doc_title)
    
    if not bg_bytes:
        # Fallback to static assets
        title_lower = title.lower()
        asset_name = "corporate_strategy.png"
        if any(keyword in title_lower for keyword in ["tech", "ai", "digital", "data", "cloud"]):
            asset_name = "corporate_tech.png"
        elif any(keyword in title_lower for keyword in ["finance", "revenue", "growth", "market", "acquisition"]):
            asset_name = "corporate_finance.png"
            
        img_path = os.path.join(os.path.dirname(__file__), "..", "assets", asset_name)
        if os.path.exists(img_path):
            with open(img_path, "rb") as f:
                bg_bytes = f.read()
                
    if bg_bytes:
        # Draw image on right gap
        img_left = M + text_width + int(M * 0.5)
        img_width = int((SW - 2 * M) * 0.40)
        _add_image(slide, bg_bytes, img_left, CT, img_width, CH - int(SH * 0.05))


def _build_two_column(slide, sp: Dict) -> None:
    _add_title(slide, sp.get("title", ""))
    left_b = sp.get("left_bullets", [])
    right_b = sp.get("right_bullets", [])
    
    # Better column sizing with centered divider
    col_w = (SW - 2 * M - int(M * 1.2)) // 2
    divider_x = M + col_w + int(M * 0.6)

    # Left column accent line
    _add_rect(slide, M, CT - 70000, col_w - 60000, 35000, C["red"])
    _add_bullets(slide, left_b[:6],
                 M, CT, col_w - 80000, CH, size_pt=14)

    # Centered divider (subtle)
    _add_rect(slide, divider_x - 18000, CT - 50000, 36000, CH + 50000, C["light"])

    # Right column accent line
    right_x = divider_x + int(M * 0.6)
    _add_rect(slide, right_x, CT - 70000, col_w - 60000, 35000, C["blue"])
    _add_bullets(slide, right_b[:6],
                 right_x, CT, col_w - 80000, CH, size_pt=14)


def _build_data_chart(slide, sp: Dict, parsed: Dict) -> None:
    _add_title(slide, sp.get("title", ""))
    chart_spec = sp.get("chart")
    img_bytes = b""

    if chart_spec and chart_spec.get("labels") and chart_spec.get("values"):
        img_bytes = render_chart(chart_spec)

    if not img_bytes:
        # Try to auto-extract from a table in parsed data
        ti = sp.get("table_index", -1)
        all_tables = parsed.get("all_tables", [])
        if 0 <= ti < len(all_tables):
            cd = all_tables[ti].chart_data()
            if cd:
                img_bytes = render_chart(cd)
        if not img_bytes and all_tables:
            for tbl in all_tables:
                cd = tbl.chart_data()
                if cd:
                    img_bytes = render_chart(cd)
                    break

    if img_bytes:
        # Properly size chart to fit within content area
        # Use 85% of available content width for good margins
        chart_w = int((SW - 2 * M) * 0.85)
        chart_left = M + int(((SW - 2 * M) - chart_w) / 2)  # Center horizontally
        chart_top = CT + int(SH * 0.04)  # Small top gap
        chart_h = CH - int(SH * 0.06)  # Use most of content height
        _add_image(slide, img_bytes, chart_left, chart_top, chart_w, chart_h)
    else:
        # Fallback: content bullets
        _add_bullets(slide, sp.get("bullets", []),
                     M, CT, SW - 2 * M, CH, size_pt=14)


def _build_data_table(slide, sp: Dict, parsed: Dict) -> None:
    _add_title(slide, sp.get("title", ""))
    all_tables = parsed.get("all_tables", [])
    ti = sp.get("table_index", -1)
    tbl = None
    if 0 <= ti < len(all_tables):
        tbl = all_tables[ti]
    elif all_tables:
        tbl = all_tables[0]

    if tbl:
        table_top = CT + int(SH * 0.02)
        table_h = int(CH * 0.60)  # Use 60% of content height for table
        table_w = SW - 2 * M
        _add_table(slide, tbl, M, table_top, table_w, table_h)

        # If table has chart data, show chart below
        cd = tbl.chart_data()
        if cd and len(tbl.rows) >= 2:
            chart_img = render_chart(cd)
            if chart_img:
                chart_top = table_top + table_h + int(SH * 0.03)
                remaining_h = SH - chart_top - int(SH * 0.08)
                if remaining_h > int(SH * 0.15):
                    chart_w = int((SW - 2 * M) * 0.80)
                    chart_left = M + int(((SW - 2 * M) - chart_w) / 2)
                    _add_image(slide, chart_img, chart_left, chart_top, chart_w, remaining_h)
    else:
        _add_bullets(slide, sp.get("bullets", []),
                     M, CT, SW - 2 * M, CH, size_pt=15)


def _build_conclusion(slide, sp: Dict) -> None:
    _add_title(slide, sp.get("title", "Key Takeaways"))
    bullets = sp.get("bullets", [])
    if not bullets:
        return
    
    n = min(len(bullets), 6)  # Support up to 6 takeaways
    if n <= 0:
        return
        
    # Dynamic spacing
    box_gap = int(SH * 0.015)
    box_h = (CH - box_gap * max(0, n - 1)) // n

    for i, point in enumerate(bullets[:n]):
        by = CT + i * (box_h + box_gap)
        accent = ACCENT_RGB[i % len(ACCENT_RGB)]

        # Creative Box Background
        _add_rect(slide, M, by, SW - 2 * M, box_h, _rgb(245, 245, 247))
        
        # Left accent stripe
        accent_w = int(SW * 0.008)
        _add_rect(slide, left=M, top=by, width=accent_w, height=box_h, fill_rgb=accent)

        # Icon (using index for auto-color selection)
        icon_bytes = get_numbered_icon(i + 1)
        icon_s = int(box_h * 0.60)
        _add_icon(slide, icon_bytes,
                  cx=M + accent_w + 70000 + icon_s // 2,
                  cy=by + box_h // 2,
                  size=icon_s)

        # Text with proper wrapping
        tx = M + accent_w + 70000 + icon_s + 70000
        tw = SW - tx - M - 30000
        _add_textbox(slide, clean_text(point),
                     left=tx, top=by, width=tw, height=box_h,
                     font_name=FONT_B, size_pt=13, color=C["dark"])


def _build_thankyou_slide(slide, sp: Dict) -> None:
    # Try to use template's title placeholder, fallback to custom textbox
    try:
        if slide.shapes.title:
            slide.shapes.title.text = "Thank You"
        else:
            raise AttributeError("No title placeholder")
    except:
        _add_textbox(slide, "Thank You",
                     left=M, top=int(SH * 0.30), width=SW - 2 * M, height=int(SH * 0.20),
                     font_name=FONT_H, size_pt=44, bold=True, color=C["dark"],
                     align=PP_ALIGN.CENTER)
    
    subtitle = clean_text(sp.get("subtitle", "Questions & Discussion"))
    
    # Add subtitle
    if subtitle:
        try:
            if len(slide.placeholders) > 1:
                try:
                    subtitle_shape = slide.placeholders[1]
                    if subtitle_shape and hasattr(subtitle_shape, "text_frame"):
                        subtitle_shape.text = subtitle
                    else:
                        raise AttributeError("No text_frame")
                except (IndexError, AttributeError, KeyError):
                    raise AttributeError("No subtitle placeholder")
            else:
                raise AttributeError("No subtitle placeholder")
        except:
            _add_textbox(slide, subtitle,
                         left=M, top=int(SH * 0.50), width=SW - 2 * M, height=int(SH * 0.12),
                         font_name=FONT_B, size_pt=18, color=C["mid"],
                         align=PP_ALIGN.CENTER)


# ─────────────────────────────────────────────────────────────────────────────
#  Layout resolver
# ─────────────────────────────────────────────────────────────────────────────

def _get_layout(prs: Presentation, name: str):
    """Get layout from template by exact name match."""
    layout_map = {
        "cover":      "1_Cover",
        "2_cover":    "2_Cover",
        "divider":    "Divider",
        "blank":      "Blank",
        "title_only": "Title only",
        "thankyou":   "Thank You",
    }
    target = layout_map.get(name.lower(), "Blank")
    
    # Exact match first
    for layout in prs.slide_layouts:
        if layout.name == target:
            return layout
    
    # Fallback: return Blank layout
    for layout in prs.slide_layouts:
        if layout.name == "Blank":
            return layout
    
    return prs.slide_layouts[-1]  # Last resort


# ─────────────────────────────────────────────────────────────────────────────
#  Main public function
# ─────────────────────────────────────────────────────────────────────────────

def generate_pptx(parsed: Dict[str, Any], slide_plan: List[Dict],
                  output_path: str) -> None:
    """Build the full PPTX from parsed markdown + AI slide plan using template."""
    prs = Presentation(TEMPLATE_PATH)
    
    # DO NOT override slide dimensions - use template's dimensions!
    # Template is 12192000 × 6858000 (16:9 ratio)
    
    # Remove any existing slides from the template (keep only layouts)
    while len(prs.slides) > 0:
        rId = prs.slides._sldIdLst[0].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[0]

    BUILDER_MAP = {
        "title":           lambda sl, sp: _build_title_slide(sl, sp, parsed),
        "agenda":          lambda sl, sp: _build_agenda_slide(sl, sp),
        "exec_summary":    lambda sl, sp: _build_exec_summary(sl, sp, parsed),
        "section_divider": lambda sl, sp: _build_section_divider(sl, sp),
        "content":         lambda sl, sp: _build_content_slide(sl, sp, parsed),
        "two_column":      lambda sl, sp: _build_two_column(sl, sp),
        "data_chart":      lambda sl, sp: _build_data_chart(sl, sp, parsed),
        "data_table":      lambda sl, sp: _build_data_table(sl, sp, parsed),
        "conclusion":      lambda sl, sp: _build_conclusion(sl, sp),
        "thankyou":        lambda sl, sp: _build_thankyou_slide(sl, sp),
    }

    for sp in slide_plan:
        slide_type = sp.get("type", "content")
        layout_name = sp.get("layout", "blank")

        # Map slide types to appropriate layouts
        if slide_type == "title":
            layout_name = "cover"
        elif slide_type == "thankyou":
            layout_name = "thankyou"
        elif slide_type == "section_divider":
            layout_name = "divider"
        else:
            layout_name = "title_only"

        layout = _get_layout(prs, layout_name)
        slide = prs.slides.add_slide(layout)

        builder = BUILDER_MAP.get(slide_type, BUILDER_MAP["content"])
        try:
            builder(slide, sp)
        except Exception as exc:
            print(f"   [WARN] Slide {sp.get('slide_number', '?')} ({slide_type}): {exc}")
            # Best-effort fallback
            try:
                _build_content_slide(slide, sp, parsed)
            except Exception:
                pass

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    prs.save(output_path)
    print(f"   [OK] Saved -> {output_path}  ({len(slide_plan)} slides)")