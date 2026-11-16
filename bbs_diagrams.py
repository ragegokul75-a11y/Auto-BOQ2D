"""
Generates crisp, high-DPI (2x supersampled) engineering rebar diagrams (PNG) using Pillow (PIL):
  - Row thumbnails for every (Shape Code, Bend/Crank Angle) combination:
      * SC-00 (0°)
      * SC-01 (45°, 90°, 135°, 180°)
      * SC-02 (45°, 90°, 135°)
      * SC-41 (45°, 90°, 135°)
      * SC-51 (90°, 135°)
  - Large detailed engineering figures for Sheet 6:
      1. Cranked Bar Arrangement (45° / 90° / 135° bent-up)
      2. Closed Stirrup & 135° Seismic Hook Detail
      3. Extra Support Bars Arrangement (Top Extra at Support & Bottom Curtailment)
"""

from __future__ import annotations

import os
from PIL import Image, ImageDraw, ImageFont

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "assets", "bbs_shapes")


def _get_font(size: int = 13, bold: bool = False):
    candidates = [
        r"C:\Windows\Fonts\calibrib.ttf" if bold else r"C:\Windows\Fonts\calibri.ttf",
        r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\segoeui.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                pass
    return ImageFont.load_default()


def _draw_arrow_dim(draw: ImageDraw.ImageDraw, p1, p2, label: str, font, color="#1E293B", offset_y=0, offset_x=0):
    x1, y1 = p1[0] + offset_x, p1[1] + offset_y
    x2, y2 = p2[0] + offset_x, p2[1] + offset_y
    draw.line([(x1, y1), (x2, y2)], fill=color, width=2)
    if abs(y2 - y1) < 3:
        draw.line([(x1, y1 - 5), (x1, y1 + 5)], fill=color, width=2)
        draw.line([(x2, y2 - 5), (x2, y2 + 5)], fill=color, width=2)
        tx, ty = (x1 + x2) // 2, y1 - 16
    else:
        draw.line([(x1 - 5, y1), (x1 + 5, y1)], fill=color, width=2)
        draw.line([(x2 - 5, y2), (x2 + 5, y2)], fill=color, width=2)
        tx, ty = x1 + 6, (y1 + y2) // 2 - 7
    bbox = draw.textbbox((0, 0), label, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    draw.rectangle([tx - tw // 2 - 3, ty - 2, tx + tw // 2 + 3, ty + th + 2], fill="#FFFFFF")
    draw.text((tx - tw // 2, ty), label, fill=color, font=font)


def make_thumbnail(shape_code: str, angle: int, out_path: str) -> str:
    """Renders a crisp 2x supersampled (256x76 -> 128x38) rebar shape figure with angle label."""
    w, h = 256, 76
    img = Image.new("RGB", (w, h), "#FFFFFF")
    d = ImageDraw.Draw(img)
    f_lbl = _get_font(18, bold=True)
    f_ang = _get_font(17, bold=True)

    rebar = "#DC2626"   # crisp engineering red
    stirrup = "#1D4ED8" # crisp blue for closed stirrup
    ang_col = "#1D4ED8"
    txt_col = "#0F172A"
    rw = 6  # rebar stroke width at 2x

    if shape_code == "SC-00":
        d.line([(20, 44), (236, 44)], fill=rebar, width=rw)
        # small end ticks
        d.line([(20, 37), (20, 51)], fill=rebar, width=3)
        d.line([(236, 37), (236, 51)], fill=rebar, width=3)
        d.text((120, 14), "L (0°)", fill=txt_col, font=f_lbl)

    elif shape_code == "SC-01":
        if angle == 45:
            pts = [(36, 16), (62, 54), (232, 54)]
            d.line(pts, fill=rebar, width=rw, joint="curve")
            d.text((66, 22), "45°", fill=ang_col, font=f_ang)
            d.text((140, 28), "B", fill=txt_col, font=f_lbl)
        elif angle == 135:
            pts = [(74, 16), (40, 54), (232, 54)]
            d.line(pts, fill=rebar, width=rw, joint="curve")
            d.text((66, 28), "135°", fill=ang_col, font=f_ang)
            d.text((150, 28), "B", fill=txt_col, font=f_lbl)
        elif angle == 180:
            d.arc([24, 24, 62, 54], start=90, end=270, fill=rebar, width=rw)
            d.line([(43, 24), (72, 24)], fill=rebar, width=rw)
            d.line([(43, 54), (232, 54)], fill=rebar, width=rw)
            d.text((82, 16), "180°", fill=ang_col, font=f_ang)
            d.text((150, 28), "B", fill=txt_col, font=f_lbl)
        else:  # 90°
            pts = [(36, 14), (36, 54), (232, 54)]
            d.line(pts, fill=rebar, width=rw, joint="curve")
            d.rectangle([36, 42, 48, 54], outline=ang_col, width=2)
            d.text((54, 26), "90°", fill=ang_col, font=f_ang)
            d.text((142, 28), "B", fill=txt_col, font=f_lbl)

    elif shape_code == "SC-02":
        if angle == 45:
            pts = [(24, 16), (48, 54), (208, 54), (232, 16)]
            d.line(pts, fill=rebar, width=rw, joint="curve")
            d.text((108, 26), "45°", fill=ang_col, font=f_ang)
        elif angle == 135:
            pts = [(56, 16), (26, 54), (230, 54), (200, 16)]
            d.line(pts, fill=rebar, width=rw, joint="curve")
            d.text((104, 26), "135°", fill=ang_col, font=f_ang)
        else:  # 90° U-bar
            pts = [(32, 14), (32, 54), (224, 54), (224, 14)]
            d.line(pts, fill=rebar, width=rw, joint="curve")
            d.rectangle([32, 42, 44, 54], outline=ang_col, width=2)
            d.rectangle([212, 42, 224, 54], outline=ang_col, width=2)
            d.text((112, 26), "90°", fill=ang_col, font=f_ang)
        d.text((10, 16), "A", fill=txt_col, font=f_lbl)
        d.text((236, 16), "C", fill=txt_col, font=f_lbl)

    elif shape_code == "SC-41":
        if angle == 90:
            pts = [(12, 18), (68, 18), (68, 56), (188, 56), (188, 18), (244, 18)]
            d.line(pts, fill=rebar, width=rw, joint="curve")
            d.text((110, 26), "90°", fill=ang_col, font=f_ang)
        elif angle == 135:
            pts = [(12, 18), (82, 18), (56, 56), (200, 56), (174, 18), (244, 18)]
            d.line(pts, fill=rebar, width=rw, joint="curve")
            d.text((106, 26), "135°", fill=ang_col, font=f_ang)
        else:  # 45° standard crank
            pts = [(12, 18), (54, 18), (84, 56), (172, 56), (202, 18), (244, 18)]
            d.line(pts, fill=rebar, width=rw, joint="curve")
            d.text((110, 26), "45°", fill=ang_col, font=f_ang)
        # dashed Hc helper
        d.line([(84, 18), (84, 56)], fill="#94A3B8", width=2)

    elif shape_code == "SC-51":
        d.rectangle([62, 12, 194, 62], outline=stirrup, width=rw)
        if angle == 90:
            d.line([(62, 12), (90, 12)], fill=rebar, width=rw)
            d.line([(62, 12), (62, 36)], fill=rebar, width=rw)
            d.text((105, 28), "90°", fill=ang_col, font=f_ang)
        else:  # 135° seismic hooks
            d.line([(62, 12), (92, 40)], fill=rebar, width=rw)
            d.line([(62, 12), (82, 44)], fill=rebar, width=rw)
            d.text((102, 28), "135°", fill=ang_col, font=f_ang)
        # corner bar dots
        for cx, cy in ((70, 20), (186, 20), (70, 54), (186, 54)):
            d.ellipse([cx - 4, cy - 4, cx + 4, cy + 4], fill=txt_col)

    img.save(out_path, "PNG")
    return out_path


def make_cranked_bar_detail(out_path: str) -> str:
    w, h = 920, 290
    img = Image.new("RGB", (w, h), "#FFFFFF")
    d = ImageDraw.Draw(img)
    f_title = _get_font(15, bold=True)
    f_lbl = _get_font(12, bold=True)
    f_sm = _get_font(11, bold=False)

    d.rectangle([0, 0, w - 1, h - 1], outline="#CBD5E1", width=2)
    d.rectangle([0, 0, w - 1, 32], fill="#0F172A")
    d.text((14, 7), "1. CRANKED BAR ARRANGEMENT (SC-41) — SLAB & BEAM BENT-UP REINFORCEMENT (45° / 90° / 135°)", fill="#FFFFFF", font=f_title)

    d.rectangle([70, 70, 850, 195], fill="#F1F5F9", outline="#475569", width=2)
    d.rectangle([45, 70, 95, 245], fill="#CBD5E1", outline="#334155", width=2)
    d.rectangle([825, 70, 875, 245], fill="#CBD5E1", outline="#334155", width=2)
    d.text((53, 210), "Support", fill="#334155", font=f_sm)
    d.text((831, 210), "Support", fill="#334155", font=f_sm)

    d.line([(55, 86), (865, 86)], fill="#64748B", width=2)
    d.line([(55, 182), (865, 182)], fill="#94A3B8", width=2)

    pts = [(55, 92), (210, 92), (285, 174), (635, 174), (710, 92), (865, 92)]
    d.line(pts, fill="#DC2626", width=5)

    d.line([(285, 92), (285, 174)], fill="#2563EB", width=2)
    d.text((292, 122), "Hc = D - 2c - d", fill="#1D4ED8", font=f_lbl)
    d.text((222, 142), "45°", fill="#DC2626", font=f_lbl)
    d.text((682, 142), "45°", fill="#DC2626", font=f_lbl)

    _draw_arrow_dim(d, (95, 50), (825, 50), "Clear Span L", f_lbl, color="#0F172A")
    _draw_arrow_dim(d, (95, 228), (285, 228), "Crank Zone (L/4 or 0.15L)", f_sm, color="#1D4ED8")
    _draw_arrow_dim(d, (635, 228), (825, 228), "Crank Zone (L/4 or 0.15L)", f_sm, color="#1D4ED8")

    d.text((240, 256), "Formula: Cut Len = L + 2×Bearing + n×(k×Hc) - n×Ded   (45°: k=0.42 | 90°: k=1.00 | 135°: k=1.30)", fill="#0F172A", font=f_lbl)
    img.save(out_path, "PNG")
    return out_path


def make_closed_stirrup_detail(out_path: str) -> str:
    w, h = 920, 290
    img = Image.new("RGB", (w, h), "#FFFFFF")
    d = ImageDraw.Draw(img)
    f_title = _get_font(15, bold=True)
    f_lbl = _get_font(12, bold=True)
    f_sm = _get_font(11, bold=False)

    d.rectangle([0, 0, w - 1, h - 1], outline="#CBD5E1", width=2)
    d.rectangle([0, 0, w - 1, 32], fill="#0F172A")
    d.text((14, 7), "2. CLOSED STIRRUP / TIE HOOK DETAIL (SC-51) — 135° SEISMIC ANCHORAGE (IS 13920 / IS 2502)", fill="#FFFFFF", font=f_title)

    d.rectangle([60, 58, 280, 248], fill="#F1F5F9", outline="#475569", width=2)
    d.rectangle([86, 84, 254, 222], outline="#1D4ED8", width=4)
    d.line([(86, 84), (126, 124)], fill="#DC2626", width=4)
    d.line([(86, 84), (120, 128)], fill="#DC2626", width=4)
    for cx, cy in ((94, 92), (246, 92), (94, 214), (246, 214)):
        d.ellipse([cx - 6, cy - 6, cx + 6, cy + 6], fill="#0F172A")

    d.text((135, 118), "135° Hooks", fill="#DC2626", font=f_lbl)
    d.text((128, 134), "(10d each)", fill="#DC2626", font=f_sm)
    _draw_arrow_dim(d, (60, 44), (280, 44), "Width b", f_sm, color="#0F172A")
    _draw_arrow_dim(d, (295, 58), (295, 248), "Depth D", f_sm, color="#0F172A")
    d.text((62, 258), "Cover c = 25mm (Beam) / 40mm (Col)", fill="#475569", font=f_sm)

    d.rectangle([380, 52, 895, 268], fill="#F8FAFC", outline="#94A3B8", width=1)
    d.text((395, 62), "ZOOM: 135° SEISMIC HOOK DETAIL & CUTTING LENGTH CALCULATION", fill="#1E293B", font=f_lbl)

    d.line([(500, 105), (610, 105)], fill="#1D4ED8", width=5)
    d.line([(500, 105), (500, 195)], fill="#1D4ED8", width=5)
    d.line([(500, 105), (555, 160)], fill="#DC2626", width=5)
    d.line([(500, 105), (546, 166)], fill="#DC2626", width=5)
    d.ellipse([496, 101, 514, 119], fill="#0F172A")
    d.text((560, 146), "Hook Length = 10d (min 75mm)", fill="#DC2626", font=f_lbl)
    d.text((516, 114), "135°", fill="#DC2626", font=f_lbl)

    d.text((395, 205), "• Core Sides:  a = b - 2c   |   b' = D - 2c", fill="#0F172A", font=f_sm)
    d.text((395, 223), "• 2 Seismic Hooks (135°):  + 2 × 10d = +20d", fill="#0F172A", font=f_sm)
    d.text((395, 241), "• Bend Deductions (3×90° + 2×135°):  -(6d + 6d) = -12d", fill="#0F172A", font=f_sm)
    img.save(out_path, "PNG")
    return out_path


def make_extra_support_detail(out_path: str) -> str:
    w, h = 920, 290
    img = Image.new("RGB", (w, h), "#FFFFFF")
    d = ImageDraw.Draw(img)
    f_title = _get_font(15, bold=True)
    f_lbl = _get_font(12, bold=True)
    f_sm = _get_font(11, bold=False)

    d.rectangle([0, 0, w - 1, h - 1], outline="#CBD5E1", width=2)
    d.rectangle([0, 0, w - 1, 32], fill="#0F172A")
    d.text((14, 7), "3. EXTRA SUPPORT BARS ARRANGEMENT (TOP EXTRA AT COLUMN SUPPORT & BOTTOM CURTAILMENT)", fill="#FFFFFF", font=f_title)

    d.rectangle([50, 75, 110, 245], fill="#CBD5E1", outline="#334155", width=2)
    d.rectangle([520, 75, 580, 245], fill="#CBD5E1", outline="#334155", width=2)
    d.rectangle([110, 75, 520, 175], fill="#F1F5F9", outline="#475569", width=2)
    d.rectangle([580, 75, 860, 175], fill="#F1F5F9", outline="#475569", width=2)

    d.line([(65, 88), (845, 88)], fill="#64748B", width=3)
    d.line([(65, 162), (845, 162)], fill="#2563EB", width=3)

    d.line([(65, 145), (65, 96), (255, 96)], fill="#DC2626", width=5)
    d.text((125, 103), "Top Extra Support (0.3L + Ld)", fill="#DC2626", font=f_lbl)

    d.line([(385, 96), (715, 96)], fill="#DC2626", width=5)
    d.text((445, 103), "Continuous Top Extra (0.3L1 + Col + 0.3L2)", fill="#DC2626", font=f_lbl)

    d.line([(175, 152), (455, 152)], fill="#16A34A", width=4)
    d.text((225, 134), "Bottom Extra Mid-Span (0.7L)", fill="#15803D", font=f_lbl)

    _draw_arrow_dim(d, (110, 54), (520, 54), "Clear Span L1", f_lbl, color="#0F172A")
    _draw_arrow_dim(d, (110, 200), (255, 200), "L/3 or 0.3L1", f_sm, color="#DC2626")
    _draw_arrow_dim(d, (385, 200), (520, 200), "0.3L1", f_sm, color="#DC2626")
    _draw_arrow_dim(d, (580, 200), (715, 200), "0.3L2", f_sm, color="#DC2626")

    d.text((65, 256), "End Support Extra Cut Len = Ld(col) + 0.3×L - 2d (90° bend)   |   Interior Extra Cut Len = 0.3×L1 + Col Width + 0.3×L2", fill="#0F172A", font=f_lbl)
    img.save(out_path, "PNG")
    return out_path


def ensure_bbs_diagrams() -> dict[str, str]:
    os.makedirs(ASSETS_DIR, exist_ok=True)
    paths: dict[str, str] = {}
    combos = [
        ("SC-00", 0),
        ("SC-01", 45), ("SC-01", 90), ("SC-01", 135), ("SC-01", 180),
        ("SC-02", 45), ("SC-02", 90), ("SC-02", 135),
        ("SC-41", 45), ("SC-41", 90), ("SC-41", 135),
        ("SC-51", 90), ("SC-51", 135),
    ]
    for sc, ang in combos:
        key = f"{sc}_{ang}"
        paths[key] = make_thumbnail(sc, ang, os.path.join(ASSETS_DIR, f"thumb_{sc}_{ang}.png"))
        if sc not in paths:
            paths[sc] = paths[key]
    paths["DETAIL_CRANK"] = make_cranked_bar_detail(os.path.join(ASSETS_DIR, "detail_1_cranked_bar.png"))
    paths["DETAIL_STIRRUP"] = make_closed_stirrup_detail(os.path.join(ASSETS_DIR, "detail_2_closed_stirrup_hook.png"))
    paths["DETAIL_EXTRA"] = make_extra_support_detail(os.path.join(ASSETS_DIR, "detail_3_extra_support_bars.png"))
    return paths


if __name__ == "__main__":
    out = ensure_bbs_diagrams()
    print(f"Generated {len(out)} crisp BBS diagram assets.")
