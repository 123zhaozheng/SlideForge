#!/usr/bin/env python3
"""Master-grade native python-pptx component toolkit for SlideForge.

Eliminates classic python-pptx visual flaws across any brand palette or slide template:
  - Enforces East Asian font (<a:ea>) so CJK text never falls back to SimSun
  - Explicitly clears heavy default PowerPoint theme shadows (<a:effectLst/>) on flat shapes
  - Provides multi-theme palette builders (navy, crimson, emerald, violet, slate, or custom brand hex)
  - Provides soft consulting shadows, OOXML linear gradients, alpha solid fills,
    precision percentage donut rings (handling OOXML angle scaling automatically),
    1px dashed Bento frames, pill badges, aspect-ratio-safe image placement,
    and consulting-grade native editable tables.
"""

from __future__ import annotations

import io
from pathlib import Path
from typing import Dict, Optional, Sequence, Tuple, Union

from PIL import Image, ImageDraw, ImageOps
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

DEFAULT_FONT = "Microsoft YaHei"


def rgb(hex_str: str) -> RGBColor:
    h = hex_str.strip().lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _tint_with_white(hex_str: str, white_ratio: float = 0.93) -> RGBColor:
    h = hex_str.strip().lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    nr = int(round(r * (1.0 - white_ratio) + 255 * white_ratio))
    ng = int(round(g * (1.0 - white_ratio) + 255 * white_ratio))
    nb = int(round(b * (1.0 - white_ratio) + 255 * white_ratio))
    return RGBColor(nr, ng, nb)


def make_palette(
    primary: str = "#1B4FD8",
    secondary: str = "#0284C7",
    tertiary: str = "#0D9488",
    ink: str = "#1F2329",
) -> Dict[str, RGBColor]:
    """Build a balanced 4-stop + neutral consulting palette from any brand hex colors."""
    return {
        "primary": rgb(primary),
        "secondary": rgb(secondary),
        "tertiary": rgb(tertiary),
        "tint_1": _tint_with_white(primary, 0.93),
        "tint_2": _tint_with_white(secondary, 0.93),
        "tint_3": _tint_with_white(tertiary, 0.93),
        "ink": rgb(ink),
        "body": rgb("#4E5969"),
        "muted": rgb("#7A8594"),
        "dash": rgb("#CDD3DB"),
        "line": rgb("#E6E9EF"),
        "white": rgb("#FFFFFF"),
    }


THEMES: Dict[str, Dict[str, RGBColor]] = {
    "navy": make_palette("#0F2A56", "#0284C7", "#0D9488"),
    "crimson": make_palette("#C4122E", "#E85D22", "#EE9B2F"),
    "indigo": make_palette("#4F46E5", "#10B981", "#F59E0B"),
    "emerald": make_palette("#047857", "#0284C7", "#D97706"),
    "slate": make_palette("#1E293B", "#3B82F6", "#64748B"),
}

PALETTE_DEFAULT = THEMES["navy"]
PALETTE_CRIMSON = THEMES["crimson"]


def set_default_font(font_name: str) -> None:
    global DEFAULT_FONT
    DEFAULT_FONT = font_name


def create_deck(width_in: float = 13.333, height_in: float = 7.5) -> Presentation:
    prs = Presentation()
    prs.slide_width = Inches(width_in)
    prs.slide_height = Inches(height_in)
    return prs


def add_blank_slide(prs: Presentation):
    return prs.slides.add_slide(prs.slide_layouts[6])


def add_baseboard(slide, baseboard_png: Union[str, Path], width_in: float = 13.333, height_in: float = 7.5):
    """Place any user-supplied background / corporate template image edge-to-edge on Layer 1."""
    return slide.shapes.add_picture(
        str(baseboard_png), Inches(0), Inches(0), width=Inches(width_in), height=Inches(height_in)
    )


def clear_shadow(shape) -> None:
    """Strip PowerPoint's default heavy theme shadow so flat vector shapes stay ultra-clean."""
    spPr = shape._element.spPr
    for old in spPr.findall(qn("a:effectLst")):
        spPr.remove(old)
    spPr.append(OxmlElement("a:effectLst"))


def set_soft_shadow(shape, blur_pt: float = 7, dist_pt: float = 2, alpha_pct: int = 7) -> None:
    spPr = shape._element.spPr
    for old in spPr.findall(qn("a:effectLst")):
        spPr.remove(old)
    effectLst = OxmlElement("a:effectLst")
    outer = OxmlElement("a:outerShdw")
    outer.set("blurRad", str(int(blur_pt * 12700)))
    outer.set("dist", str(int(dist_pt * 12700)))
    outer.set("dir", "5400000")
    outer.set("algn", "ctr")
    clr = OxmlElement("a:srgbClr")
    clr.set("val", "1F2329")
    a = OxmlElement("a:alpha")
    a.set("val", str(alpha_pct * 1000))
    clr.append(a)
    outer.append(clr)
    effectLst.append(outer)
    spPr.append(effectLst)


def set_run_font(
    run,
    font_name: Optional[str] = None,
    size: float = 12,
    color: RGBColor = PALETTE_DEFAULT["ink"],
    bold: bool = False,
) -> None:
    """Set Latin + East Asian + Complex Script font tags so CJK renders cleanly in PowerPoint/WPS."""
    fname = font_name or DEFAULT_FONT
    run.font.name = fname
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        for old in rPr.findall(qn(tag)):
            rPr.remove(old)
        el = OxmlElement(tag)
        el.set("typeface", fname)
        rPr.append(el)


def set_linear_gradient(
    shape,
    stops: Sequence[Tuple[float, str]],
    angle_deg: float = 0.0,
) -> None:
    """Apply crisp native OOXML linear gradient to any PPTX shape.

    stops: list of (position_0_to_1, hex_color) e.g. [(0.0, '#1B4FD8'), (1.0, '#0F2A56')]
    angle_deg: 0 = left->right, 90 = top->bottom, 135 = top-left->bottom-right.
    """
    spPr = shape._element.spPr
    for tag in ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill"):
        for old in spPr.findall(qn(tag)):
            spPr.remove(old)
    gradFill = OxmlElement("a:gradFill")
    gradFill.set("rotWithShape", "1")
    gsLst = OxmlElement("a:gsLst")
    for pos, hex_c in stops:
        gs = OxmlElement("a:gs")
        gs.set("pos", str(int(round(pos * 100000))))
        srgb = OxmlElement("a:srgbClr")
        srgb.set("val", hex_c.strip().lstrip("#").upper())
        gs.append(srgb)
        gsLst.append(gs)
    gradFill.append(gsLst)
    lin = OxmlElement("a:lin")
    lin.set("ang", str(int(round((angle_deg % 360) * 60000))))
    lin.set("scaled", "1")
    gradFill.append(lin)
    ln = spPr.find(qn("a:ln"))
    eff = spPr.find(qn("a:effectLst"))
    ref = ln if ln is not None else eff
    if ref is not None:
        spPr.insert(list(spPr).index(ref), gradFill)
    else:
        spPr.append(gradFill)


def set_solid_alpha(shape, hex_color: str, alpha_pct: int) -> None:
    """Apply solid fill with explicit OOXML alpha transparency (0-100)."""
    spPr = shape._element.spPr
    for tag in ("a:noFill", "a:solidFill", "a:gradFill"):
        for old in spPr.findall(qn(tag)):
            spPr.remove(old)
    sf = OxmlElement("a:solidFill")
    clr = OxmlElement("a:srgbClr")
    clr.set("val", hex_color.strip().lstrip("#").upper())
    a = OxmlElement("a:alpha")
    a.set("val", str(int(alpha_pct * 1000)))
    clr.append(a)
    sf.append(clr)
    ln = spPr.find(qn("a:ln"))
    eff = spPr.find(qn("a:effectLst"))
    ref = ln if ln is not None else eff
    if ref is not None:
        spPr.insert(list(spPr).index(ref), sf)
    else:
        spPr.append(sf)


def add_box(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    fill: Optional[RGBColor] = None,
    line_color: Optional[RGBColor] = None,
    line_pt: float = 1.0,
    dashed: bool = False,
    radius: float = 0.05,
    shadow: bool = False,
):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    if sh.adjustments:
        sh.adjustments[0] = radius
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    if line_color is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line_color
        sh.line.width = Pt(line_pt)
        if dashed:
            sh.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    if shadow:
        set_soft_shadow(sh)
    else:
        clear_shadow(sh)
    return sh


def add_bento_dashed(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    fill: Optional[RGBColor] = None,
    dash_color: RGBColor = PALETTE_DEFAULT["dash"],
    radius: float = 0.04,
):
    """Consulting-grade 1px dashed modular Bento container."""
    return add_box(slide, x, y, w, h, fill=fill, line_color=dash_color, line_pt=1.0, dashed=True, radius=radius, shadow=False)


def add_card(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    accent_color: Optional[RGBColor] = None,
    accent_side: str = "left",
    fill: RGBColor = PALETTE_DEFAULT["white"],
    border_color: RGBColor = PALETTE_DEFAULT["line"],
    radius: float = 0.08,
    shadow: bool = True,
):
    """Clean white executive card with optional left or top colored accent strip."""
    card = add_box(slide, x, y, w, h, fill=fill, line_color=border_color, line_pt=0.8, radius=radius, shadow=shadow)
    if accent_color is not None:
        if accent_side == "left":
            add_box(slide, x, y, 0.06, h, fill=accent_color, radius=0.02, shadow=False)
        elif accent_side == "top":
            add_box(slide, x, y, w, 0.06, fill=accent_color, radius=0.02, shadow=False)
    return card


def add_text(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    text: str,
    size: float = 12,
    color: RGBColor = PALETTE_DEFAULT["ink"],
    bold: bool = False,
    align: PP_ALIGN = PP_ALIGN.LEFT,
    valign: MSO_ANCHOR = MSO_ANCHOR.TOP,
    line_sp: float = 1.25,
    font_name: Optional[str] = None,
):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = valign
    tf.margin_left = tf.margin_right = Inches(0.02)
    tf.margin_top = tf.margin_bottom = Inches(0.01)
    for idx, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_sp
        r = p.add_run()
        r.text = line
        set_run_font(r, font_name=font_name, size=size, color=color, bold=bold)
    return tb


def add_rich_text(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    paragraphs: Sequence[Sequence[Tuple[str, float, RGBColor, bool]]],
    align: PP_ALIGN = PP_ALIGN.LEFT,
    valign: MSO_ANCHOR = MSO_ANCHOR.TOP,
    line_sp: float = 1.25,
    font_name: Optional[str] = None,
):
    """Add multi-run rich text where each paragraph is a list of (text, size_pt, color, bold)."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = valign
    tf.margin_left = tf.margin_right = Inches(0.02)
    tf.margin_top = tf.margin_bottom = Inches(0.01)
    for p_idx, runs in enumerate(paragraphs):
        p = tf.paragraphs[0] if p_idx == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_sp
        for txt, sz, clr, bld in runs:
            r = p.add_run()
            r.text = txt
            set_run_font(r, font_name=font_name, size=sz, color=clr, bold=bld)
    return tb


def add_pill(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    text: str,
    fill: Optional[RGBColor],
    text_color: RGBColor,
    size: float = 10.0,
    bold: bool = True,
    border: Optional[RGBColor] = None,
    font_name: Optional[str] = None,
):
    sh = add_box(slide, x, y, w, h, fill=fill, line_color=border, line_pt=0.9, radius=0.5, shadow=False)
    tf = sh.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = Inches(0.04)
    tf.margin_top = tf.margin_bottom = Inches(0.01)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    set_run_font(r, font_name=font_name, size=size, color=text_color, bold=bold)
    return sh


def add_circle_badge(
    slide,
    cx: float,
    cy: float,
    d: float,
    text: str = "",
    fill: Optional[RGBColor] = PALETTE_DEFAULT["primary"],
    text_color: RGBColor = PALETTE_DEFAULT["white"],
    size: float = 11.0,
    bold: bool = True,
    border: Optional[RGBColor] = None,
    border_pt: float = 1.0,
    shadow: bool = False,
):
    """Add a centered circle shape (cx, cy = center in inches, d = diameter in inches)."""
    sh = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - d / 2), Inches(cy - d / 2), Inches(d), Inches(d))
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    if border is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = border
        sh.line.width = Pt(border_pt)
    if shadow:
        set_soft_shadow(sh)
    else:
        clear_shadow(sh)
    if text:
        tf = sh.text_frame
        tf.word_wrap = False
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(0)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = text
        set_run_font(r, size=size, color=text_color, bold=bold)
    return sh


def add_slide_title(
    slide,
    title: str,
    x: float = 0.60,
    y: float = 0.28,
    w: float = 9.0,
    h: float = 0.52,
    size: float = 24.0,
    color: RGBColor = PALETTE_DEFAULT["ink"],
):
    """Clean top-left slide header title positioned inside the template header safe zone."""
    return add_text(slide, x, y, w, h, title, size=size, color=color, bold=True, valign=MSO_ANCHOR.MIDDLE)


def add_section_header(
    slide,
    x: float,
    y: float,
    w: float,
    number: str,
    title: str,
    subtitle: str = "",
    badge_color: RGBColor = PALETTE_DEFAULT["primary"],
):
    """Numbered colored block + bold section title + optional right-aligned subtitle."""
    add_pill(slide, x, y + 0.03, 0.42, 0.28, number, fill=badge_color, text_color=PALETTE_DEFAULT["white"], size=11, bold=True)
    add_text(slide, x + 0.50, y, w * 0.58, 0.34, title, size=14, color=PALETTE_DEFAULT["ink"], bold=True, valign=MSO_ANCHOR.MIDDLE)
    if subtitle:
        add_text(
            slide, x + w * 0.50, y, w * 0.50, 0.34, subtitle,
            size=10.5, color=badge_color, bold=True, align=PP_ALIGN.RIGHT, valign=MSO_ANCHOR.MIDDLE,
        )


def add_line_seg(
    slide,
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    color: RGBColor = PALETTE_DEFAULT["line"],
    width_pt: float = 1.0,
    dashed: bool = False,
):
    """Add a straight connector line between (x1, y1) and (x2, y2) in inches."""
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    conn.line.color.rgb = color
    conn.line.width = Pt(width_pt)
    if dashed:
        conn.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    clear_shadow(conn)
    return conn


def add_donut_ring(
    slide,
    cx: float,
    cy: float,
    d: float,
    pct: float,
    thickness: float = 0.16,
    track_color: RGBColor = RGBColor(0xE6, 0xE8, 0xEC),
    arc_stops: Sequence[Tuple[float, str]] = ((0.0, "#1B4FD8"), (0.55, "#0284C7"), (1.0, "#0D9488")),
    shadow: bool = True,
):
    """Draw a precision 12-o'clock clockwise native editable percentage donut ring.

    Automatically handles OOXML angle scaling (60000ths/deg vs python-pptx *100000 => deg * 0.6).
    pct: 0.0 to 1.0 (e.g. 0.48 for 48%).
    """
    pct_clamped = max(0.01, min(0.999, float(pct)))
    base = slide.shapes.add_shape(MSO_SHAPE.DONUT, Inches(cx - d / 2), Inches(cy - d / 2), Inches(d), Inches(d))
    if base.adjustments:
        base.adjustments[0] = thickness
    base.fill.solid()
    base.fill.fore_color.rgb = track_color
    base.line.fill.background()
    if shadow:
        set_soft_shadow(base, blur_pt=10, dist_pt=3, alpha_pct=9)
    else:
        clear_shadow(base)

    start_deg = 270.0  # 12 o'clock
    end_deg = (start_deg + 360.0 * pct_clamped) % 360.0
    arc = slide.shapes.add_shape(MSO_SHAPE.BLOCK_ARC, Inches(cx - d / 2), Inches(cy - d / 2), Inches(d), Inches(d))
    if arc.adjustments and len(arc.adjustments) >= 3:
        arc.adjustments[0] = start_deg * 0.6
        arc.adjustments[1] = end_deg * 0.6
        arc.adjustments[2] = thickness
    arc.line.fill.background()
    set_linear_gradient(arc, arc_stops, angle_deg=90.0)
    clear_shadow(arc)
    return base, arc


def add_picture_fit(
    slide,
    img_path: Union[str, Path],
    x: float,
    y: float,
    w: float,
    h: float,
    fit: str = "contain",
):
    """Insert transparent AI component PNG preserving exact aspect ratio (contain) or filling box (stretch)."""
    p = Path(img_path)
    if fit == "stretch":
        return slide.shapes.add_picture(str(p), Inches(x), Inches(y), width=Inches(w), height=Inches(h))
    with Image.open(p) as im:
        iw, ih = im.size
    if not (iw and ih):
        return slide.shapes.add_picture(str(p), Inches(x), Inches(y), width=Inches(w), height=Inches(h))
    scale = min(w / iw, h / ih)
    nw, nh = iw * scale, ih * scale
    nx = x + (w - nw) / 2.0
    ny = y + (h - nh) / 2.0
    return slide.shapes.add_picture(str(p), Inches(nx), Inches(ny), width=Inches(nw), height=Inches(nh))


def _render_slot_image_bytes(
    img_path: Path,
    target_ratio: float,
    fit: str = "cover",
    shape: str = "rounded",
    radius_ratio: float = 0.055,
) -> Tuple[io.BytesIO, Tuple[int, int]]:
    """Center-crop (cover) and apply optional supersampled rounded/circle alpha mask in memory."""
    with Image.open(img_path) as raw:
        im = ImageOps.exif_transpose(raw).convert("RGBA")

    iw, ih = im.size
    if fit == "cover" and iw > 0 and ih > 0 and target_ratio > 0:
        src_ratio = iw / ih
        if src_ratio > target_ratio:
            new_w = max(1, int(round(ih * target_ratio)))
            left = (iw - new_w) // 2
            im = im.crop((left, 0, left + new_w, ih))
        elif src_ratio < target_ratio:
            new_h = max(1, int(round(iw / target_ratio)))
            top = (ih - new_h) // 2
            im = im.crop((0, top, iw, top + new_h))

    cw, ch = im.size
    if shape == "circle":
        side = min(cw, ch)
        left = (cw - side) // 2
        top = (ch - side) // 2
        im = im.crop((left, top, left + side, top + side))
        cw = ch = side
        ss = 2
        mask = Image.new("L", (cw * ss, ch * ss), 0)
        ImageDraw.Draw(mask).ellipse((0, 0, cw * ss - 1, ch * ss - 1), fill=255)
        mask = mask.resize((cw, ch), Image.Resampling.LANCZOS)
        r, g, b, a = im.split()
        im.putalpha(Image.composite(a, Image.new("L", (cw, ch), 0), mask))
    elif shape == "rounded":
        radius = max(4, int(round(min(cw, ch) * radius_ratio)))
        ss = 2
        mask = Image.new("L", (cw * ss, ch * ss), 0)
        ImageDraw.Draw(mask).rounded_rectangle(
            (0, 0, cw * ss - 1, ch * ss - 1),
            radius=radius * ss,
            fill=255,
        )
        mask = mask.resize((cw, ch), Image.Resampling.LANCZOS)
        r, g, b, a = im.split()
        im.putalpha(Image.composite(a, Image.new("L", (cw, ch), 0), mask))

    buf = io.BytesIO()
    im.save(buf, format="PNG")
    buf.seek(0)
    return buf, (cw, ch)


def add_image_slot(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    img_path: Optional[Union[str, Path]] = None,
    fit: str = "cover",
    shape: str = "rounded",
    radius_ratio: float = 0.055,
    border_color: Optional[RGBColor] = PALETTE_DEFAULT["line"],
    border_pt: float = 0.75,
    shadow: bool = True,
    caption: str = "",
    placeholder_label: str = "图片 / 插图槽位（可在 PowerPoint 中替换为真实图片）",
    placeholder_fill: RGBColor = RGBColor(0xF6, 0xF8, 0xFA),
    placeholder_border: RGBColor = PALETTE_DEFAULT["dash"],
):
    """Place a real photo/illustration into a clean visual slot, or render a native dashed placeholder slot.

    - Filled mode (img_path provided): supports `fit="cover"` (smart center crop, zero distortion)
      or `fit="contain"`, with `shape="rect" | "rounded" | "circle"`, subtle frame/shadow, and optional caption.
    - Placeholder mode (img_path is None): draws a crisp native editable dashed slot box with centered guidance
      text so users can drop in their own screenshot/photo later in PowerPoint without breaking layout.
    """
    cap_h = 0.28 if caption else 0.0
    img_h = max(0.3, h - cap_h)

    if img_path is not None and str(img_path).strip():
        p = Path(img_path)
        if not p.is_file():
            raise FileNotFoundError(f"Image slot file not found: {p}")

        target_ratio = (w / img_h) if img_h > 0 else 1.0
        buf, (cw, ch) = _render_slot_image_bytes(
            p,
            target_ratio=target_ratio,
            fit=fit,
            shape=shape,
            radius_ratio=radius_ratio,
        )

        if shape == "circle":
            side = min(w, img_h)
            px = x + (w - side) / 2.0
            py = y + (img_h - side) / 2.0
            pw = ph = side
        elif fit == "contain" and cw > 0 and ch > 0:
            scale = min(w / cw, img_h / ch)
            pw, ph = cw * scale, ch * scale
            px = x + (w - pw) / 2.0
            py = y + (img_h - ph) / 2.0
        else:
            px, py, pw, ph = x, y, w, img_h

        frame_shape = None
        if border_color is not None or shadow:
            if shape == "circle":
                frame_shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(px), Inches(py), Inches(pw), Inches(ph))
            elif shape == "rounded":
                frame_shape = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(px), Inches(py), Inches(pw), Inches(ph)
                )
                if frame_shape.adjustments:
                    frame_shape.adjustments[0] = radius_ratio
            else:
                frame_shape = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE, Inches(px), Inches(py), Inches(pw), Inches(ph)
                )
            frame_shape.fill.solid()
            frame_shape.fill.fore_color.rgb = PALETTE_DEFAULT["white"]
            if border_color is not None:
                frame_shape.line.color.rgb = border_color
                frame_shape.line.width = Pt(border_pt)
            else:
                frame_shape.line.fill.background()
            if shadow:
                set_soft_shadow(frame_shape, blur_pt=10, dist_pt=3, alpha_pct=8)
            else:
                clear_shadow(frame_shape)

        pic = slide.shapes.add_picture(buf, Inches(px), Inches(py), width=Inches(pw), height=Inches(ph))
        primary_obj = pic
    else:
        mso = MSO_SHAPE.OVAL if shape == "circle" else (
            MSO_SHAPE.ROUNDED_RECTANGLE if shape == "rounded" else MSO_SHAPE.RECTANGLE
        )
        box = slide.shapes.add_shape(mso, Inches(x), Inches(y), Inches(w), Inches(img_h))
        if shape == "rounded" and box.adjustments:
            box.adjustments[0] = radius_ratio
        box.fill.solid()
        box.fill.fore_color.rgb = placeholder_fill
        box.line.color.rgb = placeholder_border
        box.line.width = Pt(1.25)
        box.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        clear_shadow(box)
        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = Inches(0.25)
        tf.margin_top = tf.margin_bottom = Inches(0.15)
        p0 = tf.paragraphs[0]
        p0.alignment = PP_ALIGN.CENTER
        r0 = p0.add_run()
        r0.text = placeholder_label
        set_run_font(r0, size=11, color=PALETTE_DEFAULT["muted"], bold=False)
        primary_obj = box

    if caption:
        add_text(
            slide,
            x,
            y + img_h + 0.04,
            w,
            0.24,
            caption,
            size=9.5,
            color=PALETTE_DEFAULT["muted"],
            align=PP_ALIGN.CENTER,
        )
    return primary_obj


def _set_cell_borders(cell, bottom_hex: str = "E6E9EF", bottom_pt: float = 0.75) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    for side in ("lnL", "lnR", "lnT", "lnB"):
        for old in tcPr.findall(qn(f"a:{side}")):
            tcPr.remove(old)
        ln = OxmlElement(f"a:{side}")
        if side == "lnB" and bottom_hex:
            ln.set("w", str(int(bottom_pt * 12700)))
            solid = OxmlElement("a:solidFill")
            clr = OxmlElement("a:srgbClr")
            clr.set("val", bottom_hex)
            solid.append(clr)
            ln.append(solid)
        else:
            ln.append(OxmlElement("a:noFill"))
        tcPr.append(ln)


def add_table_matrix(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    headers: Sequence[str],
    rows: Sequence[Sequence[str]],
    col_ratios: Optional[Sequence[float]] = None,
    header_fill: RGBColor = PALETTE_DEFAULT["primary"],
    stripe_fill: RGBColor = RGBColor(0xFA, 0xFB, 0xFC),
    font_size: float = 10.5,
    header_size: float = 11.5,
):
    """Create a 100% native editable executive consulting table (no ugly blue theme or black gridlines)."""
    n_rows = 1 + len(rows)
    n_cols = len(headers)
    shape = slide.shapes.add_table(n_rows, n_cols, Inches(x), Inches(y), Inches(w), Inches(h))
    tbl = shape.table

    tblPr = tbl._tbl.tblPr
    for attr in ("bandRow", "firstRow", "bandCol", "firstCol", "lastRow", "lastCol"):
        if attr in tblPr.attrib:
            del tblPr.attrib[attr]
    for st in tblPr.findall(qn("a:tableStyleId")):
        tblPr.remove(st)

    ratios = list(col_ratios) if col_ratios and len(col_ratios) == n_cols else [1.0] * n_cols
    total_r = sum(ratios)
    for c_i, r_val in enumerate(ratios):
        tbl.columns[c_i].width = Inches(w * (r_val / total_r))

    for c_i, text in enumerate(headers):
        cell = tbl.cell(0, c_i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_fill
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = cell.margin_right = Inches(0.10)
        cell.margin_top = cell.margin_bottom = Inches(0.05)
        _set_cell_borders(cell, bottom_hex="", bottom_pt=0)
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER if c_i == 0 else PP_ALIGN.LEFT
        r = p.add_run()
        r.text = str(text)
        set_run_font(r, size=header_size, color=PALETTE_DEFAULT["white"], bold=True)

    for r_i, row_vals in enumerate(rows, start=1):
        bg = PALETTE_DEFAULT["white"] if r_i % 2 == 1 else stripe_fill
        for c_i, val in enumerate(row_vals):
            cell = tbl.cell(r_i, c_i)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = Inches(0.10)
            cell.margin_top = cell.margin_bottom = Inches(0.05)
            _set_cell_borders(cell, bottom_hex="E6E9EF", bottom_pt=0.75)
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if c_i == 0 else PP_ALIGN.LEFT
            p.line_spacing = 1.22
            r = p.add_run()
            r.text = str(val)
            set_run_font(
                r,
                size=font_size,
                color=PALETTE_DEFAULT["ink"] if c_i == 0 else PALETTE_DEFAULT["body"],
                bold=(c_i == 0),
            )
    return shape
