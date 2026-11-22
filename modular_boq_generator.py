"""
SmartCivilTools-Style Modular AutoCAD-to-Excel BOQ Generator (9 Separate Tabs)
Generates:
  1. Commercial Summary (Abstract + Contingency + OHP + GST + Cost/Sqft)
  2. Foundation & Earthwork
  3. RCC Structure & Steel (Columns, Beams, Slabs, Staircase, Shuttering, TMT)
  4. Masonry & Plastering (with live Opening Deductions linked from Tab 5)
  5. Doors, Windows & Openings Schedule
  6. Flooring, Tiles & Painting (Room-by-Room Schedule)
  7. Electrical, Plumbing & Additional Works
  8. Concrete & Mortar Material Mix Yield (Cement Bags, Sand, Aggregate, Bricks, Steel)
  9. AutoCAD Raw Takeoff (Polylines imported from .DXF or AutoBOQ_SmartTools.lsp .CSV)

Also supports single-part standalone Excel exports and live DXF/CSV ingestion.
"""

from __future__ import annotations

import csv
import io
import math
import os
import shutil
import subprocess
import tempfile
from typing import Any, Dict, List

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXPORTS_DIR = os.path.join(BASE_DIR, "exports")
MASTER_FILENAME = "Modular_BOQ_Master.xlsx"

# Styling
THIN = Side(style="thin", color="CBD5E1")
MED = Side(style="medium", color="162A4A")
BORDER_GRID = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BORDER_HDR = Border(left=THIN, right=THIN, top=MED, bottom=MED)
BORDER_TOT = Border(left=THIN, right=THIN, top=MED, bottom=MED)

FILL_NAVY = PatternFill("solid", fgColor="162A4A")
FILL_SEC = PatternFill("solid", fgColor="DBEAFE")
FILL_INPUT = PatternFill("solid", fgColor="FEF9C3")
FILL_CALC = PatternFill("solid", fgColor="F8FAFC")
FILL_TOTAL = PatternFill("solid", fgColor="E2E8F0")
FILL_GREEN = PatternFill("solid", fgColor="DCFCE7")
FILL_AMBER = PatternFill("solid", fgColor="FEF3C7")

FONT_TITLE = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
FONT_SUB = Font(name="Calibri", size=9, italic=True, color="DBEAFE")
FONT_HDR = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
FONT_SEC = Font(name="Calibri", size=10, bold=True, color="162A4A")
FONT_DATA = Font(name="Calibri", size=10, color="0F172A")
FONT_INPUT = Font(name="Calibri", size=10, bold=True, color="1E3A8A")
FONT_BOLD = Font(name="Calibri", size=10, bold=True, color="0F172A")

ALIGN_L = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_R = Alignment(horizontal="right", vertical="center", wrap_text=True)


# ---------------------------------------------------------------------------
# AutoCAD (.DXF or .CSV) Takeoff Parser
# ---------------------------------------------------------------------------
def _classify_layer(layer: str) -> str:
    u = (layer or "").upper()
    if any(k in u for k in ("FOOT", "FTG", "FDN")):
        return "FOOTING"
    if any(k in u for k in ("COL", "PIL")):
        return "COLUMN"
    if any(k in u for k in ("BEAM", "BM", "PLINTH")):
        return "BEAM"
    if any(k in u for k in ("SLAB", "ROOF")):
        return "SLAB"
    if any(k in u for k in ("WALL", "MAS", "BRICK")):
        return "MASONRY"
    if any(k in u for k in ("DOOR", "WIN", "OPEN")):
        return "OPENING"
    return "ROOM"


def _poly_area_perim(pts: list[tuple[float, float]]) -> tuple[float, float, float, float]:
    n = len(pts)
    if n < 3:
        return 0.0, 0.0, 0.0, 0.0
    area = 0.0
    perim = 0.0
    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    for i in range(n):
        x1, y1 = pts[i]
        x2, y2 = pts[(i + 1) % n]
        area += x1 * y2 - x2 * y1
        perim += math.hypot(x2 - x1, y2 - y1)
    return abs(area) * 0.5, perim, max(xs) - min(xs), max(ys) - min(ys)


def parse_autocad_takeoff(file_bytes: bytes, filename: str) -> list[dict[str, Any]]:
    """Parse AutoCAD .CSV (from AutoBOQ_SmartTools.lsp) or .DXF into normalized takeoff rows."""
    ext = os.path.splitext(filename)[1].lower()
    rows: list[dict[str, Any]] = []

    if ext == ".csv":
        text = file_bytes.decode("utf-8-sig", errors="ignore")
        reader = csv.DictReader(io.StringIO(text))
        for idx, rec in enumerate(reader, start=1):
            rows.append({
                "part": (rec.get("Part_Category") or "ROOM").strip().upper(),
                "layer": (rec.get("Layer") or "0").strip(),
                "mark": (rec.get("Mark") or f"E{idx}").strip(),
                "nos": int(float(rec.get("Nos") or 1)),
                "area_m2": round(float(rec.get("Area_m2") or 0.0), 3),
                "perim_m": round(float(rec.get("Perimeter_m") or 0.0), 3),
                "len_m": round(float(rec.get("Length_m") or 0.0), 3),
                "wid_m": round(float(rec.get("Width_m") or 0.0), 3),
                "ht_m": round(float(rec.get("Depth_or_Height_m") or 3.0), 3),
            })
        return rows

    if ext == ".dxf":
        import ezdxf

        fd, tmp_p = tempfile.mkstemp(suffix=".dxf")
        try:
            with os.fdopen(fd, "wb") as fh:
                fh.write(file_bytes)
            doc = ezdxf.readfile(tmp_p)
            msp = doc.modelspace()
            ins = int(doc.header.get("$INSUNITS", 4) or 4)
            scale = {1: 0.0254, 2: 0.3048, 4: 0.001, 5: 0.01, 6: 1.0}.get(ins, 0.001)
            default_ht_map = {"FOOTING": 0.525, "COLUMN": 3.3, "BEAM": 0.45, "SLAB": 0.125, "MASONRY": 3.0, "OPENING": 2.1, "ROOM": 3.0}
            cnt = 0
            for ent in msp.query("LWPOLYLINE"):
                if not ent.closed:
                    continue
                pts = [(p[0] * scale, p[1] * scale) for p in ent.get_points("xy")]
                area_m2, perim_m, lx, ly = _poly_area_perim(pts)
                if area_m2 < 0.04:
                    continue
                cnt += 1
                lay = ent.dxf.layer or "0"
                cat = _classify_layer(lay)
                rows.append({
                    "part": cat,
                    "layer": lay,
                    "mark": f"{cat[:1]}{cnt}",
                    "nos": 1,
                    "area_m2": round(area_m2, 3),
                    "perim_m": round(perim_m, 3),
                    "len_m": round(lx, 3),
                    "wid_m": round(ly, 3),
                    "ht_m": default_ht_map.get(cat, 3.0),
                })
        finally:
            if os.path.exists(tmp_p):
                os.remove(tmp_p)
        return rows

    return rows


def default_autocad_sample_rows() -> list[dict[str, Any]]:
    return [
        {"part": "FOOTING", "layer": "S-FOOTING", "mark": "F1", "nos": 6, "area_m2": 3.24, "perim_m": 7.20, "len_m": 1.80, "wid_m": 1.80, "ht_m": 0.525},
        {"part": "FOOTING", "layer": "S-FOOTING", "mark": "F2", "nos": 4, "area_m2": 4.41, "perim_m": 8.40, "len_m": 2.10, "wid_m": 2.10, "ht_m": 0.600},
        {"part": "COLUMN",  "layer": "S-COLUMN",  "mark": "C1", "nos": 10, "area_m2": 0.135, "perim_m": 1.50, "len_m": 0.30, "wid_m": 0.45, "ht_m": 3.30},
        {"part": "SLAB",    "layer": "S-SLAB",    "mark": "S1", "nos": 1, "area_m2": 112.00, "perim_m": 44.00, "len_m": 14.00, "wid_m": 8.00, "ht_m": 0.125},
        {"part": "ROOM",    "layer": "A-ROOM",    "mark": "Living Hall", "nos": 1, "area_m2": 26.40, "perim_m": 21.20, "len_m": 6.00, "wid_m": 4.40, "ht_m": 3.00},
        {"part": "ROOM",    "layer": "A-ROOM",    "mark": "Kitchen",     "nos": 1, "area_m2": 12.60, "perim_m": 14.40, "len_m": 4.20, "wid_m": 3.00, "ht_m": 3.00},
        {"part": "ROOM",    "layer": "A-ROOM",    "mark": "Master Bed",  "nos": 1, "area_m2": 18.00, "perim_m": 17.00, "len_m": 4.50, "wid_m": 4.00, "ht_m": 3.00},
        {"part": "ROOM",    "layer": "A-ROOM",    "mark": "Bedroom 2",   "nos": 1, "area_m2": 14.40, "perim_m": 15.20, "len_m": 4.00, "wid_m": 3.60, "ht_m": 3.00},
        {"part": "ROOM",    "layer": "A-ROOM",    "mark": "Bath 1",      "nos": 1, "area_m2": 4.80,  "perim_m": 8.80,  "len_m": 2.40, "wid_m": 2.00, "ht_m": 3.00},
        {"part": "ROOM",    "layer": "A-ROOM",    "mark": "Bath 2",      "nos": 1, "area_m2": 4.20,  "perim_m": 8.20,  "len_m": 2.10, "wid_m": 2.00, "ht_m": 3.00},
    ]


# ---------------------------------------------------------------------------
# Sheet Helpers
# ---------------------------------------------------------------------------
def _banner(ws, max_col: int, title: str, subtitle: str) -> None:
    ws.views.sheetView[0].showGridLines = True
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    c1 = ws.cell(1, 1, title)
    c1.font = FONT_TITLE
    c1.fill = FILL_NAVY
    c1.alignment = ALIGN_L
    ws.row_dimensions[1].height = 26

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
    c2 = ws.cell(2, 1, subtitle)
    c2.font = FONT_SUB
    c2.fill = FILL_NAVY
    c2.alignment = ALIGN_L
    ws.row_dimensions[2].height = 19


def _write_table_headers(ws, row: int, headers: list[tuple[str, int]]) -> None:
    ws.row_dimensions[row].height = 24
    for ci, (txt, width) in enumerate(headers, start=1):
        c = ws.cell(row, ci, txt)
        c.font = FONT_HDR
        c.fill = FILL_NAVY
        c.alignment = ALIGN_C
        c.border = BORDER_HDR
        ws.column_dimensions[get_column_letter(ci)].width = width


def _configure_print(ws, landscape: bool = True) -> None:
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE if landscape else ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def _build_standard_trade_sheet(
    wb: Workbook,
    sheet_name: str,
    title: str,
    subtitle: str,
    rows_data: list[tuple],
) -> dict[str, Any]:
    """Build a standard 11-column measurement & BOQ trade sheet (A..K)."""
    ws = wb.create_sheet(sheet_name)
    _banner(ws, 11, title, subtitle)
    _write_table_headers(ws, 4, [
        ("Code", 9),
        ("Description of Work (IS 1200 / CPWD DSR)", 48),
        ("Unit", 8),
        ("Nos", 8),
        ("L / Area", 10),
        ("B", 8),
        ("D / H", 8),
        ("Quantity", 11),
        ("Rate (₹)", 12),
        ("Amount (₹)", 15),
        ("Remarks / Basis", 22),
    ])

    start_r = 5
    code_rows: dict[str, int] = {}
    for idx, (code, desc, unit, nos, lv, bv, dv, rate, rem) in enumerate(rows_data):
        r = start_r + idx
        code_rows[code] = r
        ws.row_dimensions[r].height = 30 if len(desc) > 80 else 23
        ws.cell(r, 1, code).alignment = ALIGN_C
        ws.cell(r, 2, desc).alignment = ALIGN_L
        ws.cell(r, 3, unit).alignment = ALIGN_C
        ws.cell(r, 4, nos).alignment = ALIGN_C
        ws.cell(r, 5, lv).alignment = ALIGN_R
        ws.cell(r, 6, bv).alignment = ALIGN_R
        ws.cell(r, 7, dv).alignment = ALIGN_R
        ws.cell(r, 8, f"=ROUND(D{r}*E{r}*F{r}*G{r}, 2)").alignment = ALIGN_R
        ws.cell(r, 9, rate).alignment = ALIGN_R
        ws.cell(r, 10, f"=ROUND(H{r}*I{r}, 2)").alignment = ALIGN_R
        ws.cell(r, 11, rem).alignment = ALIGN_L

        for ci in range(1, 12):
            c = ws.cell(r, ci)
            c.border = BORDER_GRID
            is_formula_input = isinstance(c.value, str) and c.value.startswith("=")
            if ci in (4, 5, 6, 7, 9) and not is_formula_input:
                c.fill = FILL_INPUT
                c.font = FONT_INPUT
            else:
                c.fill = FILL_CALC
                c.font = FONT_BOLD if ci == 10 else FONT_DATA
        for ci in (5, 6, 7, 8):
            ws.cell(r, ci).number_format = "#,##0.00"
        ws.cell(r, 9).number_format = "₹#,##0.00"
        ws.cell(r, 10).number_format = "₹#,##0.00"

    end_r = start_r + len(rows_data) - 1
    tot_r = end_r + 1
    ws.row_dimensions[tot_r].height = 24
    ws.merge_cells(start_row=tot_r, start_column=1, end_row=tot_r, end_column=9)
    ws.cell(tot_r, 1, f"SUB-TOTAL — {sheet_name.upper()}").alignment = ALIGN_R
    ws.cell(tot_r, 10, f"=SUM(J{start_r}:J{end_r})").alignment = ALIGN_R
    ws.cell(tot_r, 10).number_format = "₹#,##0.00"
    for ci in range(1, 12):
        ws.cell(tot_r, ci).fill = FILL_GREEN if ci == 10 else FILL_TOTAL
        ws.cell(tot_r, ci).font = FONT_BOLD
        ws.cell(tot_r, ci).border = BORDER_TOT

    ws.freeze_panes = "A5"
    _configure_print(ws, landscape=True)
    return {"sheet_name": sheet_name, "total_row": tot_r, "code_rows": code_rows}


# ---------------------------------------------------------------------------
# Individual Modular Sheet Builders
# ---------------------------------------------------------------------------
def build_autocad_raw_sheet(wb: Workbook, raw_rows: list[dict[str, Any]]) -> dict[str, Any]:
    ws = wb.create_sheet("9. AutoCAD Raw Takeoff")
    _banner(
        ws,
        10,
        "AUTOCAD RAW POLYLINE TAKEOFF (IMPORTED FROM .DXF OR AutoBOQ_SmartTools.lsp .CSV)",
        "Auto-extracted element areas, perimeters, and bounding dimensions by building part.",
    )
    _write_table_headers(ws, 4, [
        ("S.No", 7),
        ("Part Category", 14),
        ("AutoCAD Layer", 16),
        ("Mark / Room", 18),
        ("Nos", 8),
        ("Area (m²)", 11),
        ("Perimeter (m)", 12),
        ("Length (m)", 11),
        ("Width (m)", 11),
        ("Depth/Ht (m)", 12),
    ])
    for idx, rdata in enumerate(raw_rows, start=1):
        r = 4 + idx
        ws.row_dimensions[r].height = 21
        vals = [
            idx, rdata["part"], rdata["layer"], rdata["mark"], rdata["nos"],
            rdata["area_m2"], rdata["perim_m"], rdata["len_m"], rdata["wid_m"], rdata["ht_m"],
        ]
        for ci, v in enumerate(vals, start=1):
            c = ws.cell(r, ci, v)
            c.border = BORDER_GRID
            c.fill = FILL_INPUT if ci >= 5 else FILL_CALC
            c.font = FONT_INPUT if ci >= 5 else FONT_DATA
            c.alignment = ALIGN_C if ci in (1, 2, 5) else (ALIGN_L if ci in (3, 4) else ALIGN_R)
            if ci >= 6:
                c.number_format = "#,##0.000"
    ws.freeze_panes = "A5"
    _configure_print(ws, landscape=True)
    return {"sheet_name": ws.title, "row_count": len(raw_rows)}


def build_doors_windows_sheet(wb: Workbook, params: dict[str, Any]) -> dict[str, Any]:
    ws = wb.create_sheet("5. Doors & Windows")
    _banner(
        ws,
        9,
        "DOORS, WINDOWS & OPENINGS SCHEDULE (LINKED TO MASONRY & PLASTER DEDUCTIONS)",
        "Total Opening Area in Col F automatically feeds Masonry & Plaster deduction formulas on Tab 4.",
    )
    _write_table_headers(ws, 4, [
        ("Mark", 9),
        ("Opening Description (Frame, Shutter & Hardware)", 44),
        ("Nos", 8),
        ("Width (m)", 10),
        ("Height (m)", 10),
        ("Total Area (m²)", 13),
        ("Rate / No (₹)", 13),
        ("Amount (₹)", 15),
        ("Remarks", 20),
    ])
    openings = [
        ("MD", "Main Entrance Teakwood Panelled Door (1.20×2.10 m) with SS brass hardware", 1, 1.20, 2.10, 28500.0, "Teak frame + polish"),
        ("D1", "Bedroom Flush Door (1.00×2.10 m) with laminate finish & mortise lock", 3, 1.00, 2.10, 14500.0, "WPC / Hardwood frame"),
        ("D2", "Toilet / Utility Waterproof WPC Door (0.80×2.10 m) with SS fittings", 2, 0.80, 2.10, 9800.0, "Waterproof WPC"),
        ("W1", "UPVC 3-Track Sliding Window (1.50×1.35 m) with 5 mm clear glass & mosquito mesh", 5, 1.50, 1.35, 13200.0, "Living & Bedrooms"),
        ("W2", "UPVC 2-Track Sliding Window (1.20×1.20 m) with MS safety grill", 2, 1.20, 1.20, 10500.0, "Kitchen & Lobby"),
        ("V1", "UPVC Louvered Ventilator (0.60×0.60 m) with exhaust fan cutout", 2, 0.60, 0.60, 3600.0, "Bathrooms"),
    ]
    start_r = 5
    for idx, (mk, desc, nos, w, h, rate, rem) in enumerate(openings):
        r = start_r + idx
        ws.row_dimensions[r].height = 24
        ws.cell(r, 1, mk).alignment = ALIGN_C
        ws.cell(r, 2, desc).alignment = ALIGN_L
        ws.cell(r, 3, nos).alignment = ALIGN_C
        ws.cell(r, 4, w).alignment = ALIGN_R
        ws.cell(r, 5, h).alignment = ALIGN_R
        ws.cell(r, 6, f"=ROUND(C{r}*D{r}*E{r}, 2)").alignment = ALIGN_R
        ws.cell(r, 7, rate).alignment = ALIGN_R
        ws.cell(r, 8, f"=ROUND(C{r}*G{r}, 2)").alignment = ALIGN_R
        ws.cell(r, 9, rem).alignment = ALIGN_L
        for ci in range(1, 10):
            c = ws.cell(r, ci)
            c.border = BORDER_GRID
            if ci in (3, 4, 5, 7):
                c.fill = FILL_INPUT
                c.font = FONT_INPUT
            else:
                c.fill = FILL_CALC
                c.font = FONT_BOLD if ci in (6, 8) else FONT_DATA
        for ci in (4, 5, 6):
            ws.cell(r, ci).number_format = "#,##0.00"
        ws.cell(r, 7).number_format = "₹#,##0.00"
        ws.cell(r, 8).number_format = "₹#,##0.00"

    end_r = start_r + len(openings) - 1
    tot_r = end_r + 1
    ws.row_dimensions[tot_r].height = 24
    ws.merge_cells(start_row=tot_r, start_column=1, end_row=tot_r, end_column=5)
    ws.cell(tot_r, 1, "TOTAL OPENINGS AREA (m²) & DOORS/WINDOWS COST").alignment = ALIGN_R
    ws.cell(tot_r, 6, f"=SUM(F{start_r}:F{end_r})").alignment = ALIGN_R
    ws.cell(tot_r, 6).number_format = "#,##0.00"
    ws.cell(tot_r, 8, f"=SUM(H{start_r}:H{end_r})").alignment = ALIGN_R
    ws.cell(tot_r, 8).number_format = "₹#,##0.00"
    for ci in range(1, 10):
        ws.cell(tot_r, ci).fill = FILL_GREEN if ci == 8 else FILL_TOTAL
        ws.cell(tot_r, ci).font = FONT_BOLD
        ws.cell(tot_r, ci).border = BORDER_TOT

    ws.freeze_panes = "A5"
    _configure_print(ws, landscape=True)
    return {"sheet_name": ws.title, "total_row": tot_r, "area_col": "F", "amount_col": "H"}


def build_material_mix_yield_sheet(
    wb: Workbook,
    f_ref: dict[str, Any],
    r_ref: dict[str, Any],
    m_ref: dict[str, Any],
    static_qty: dict[str, float] | None = None,
) -> dict[str, Any]:
    """Tab 8: Concrete, Mortar, Brick & Steel Procurement Mix Yield."""
    ws = wb.create_sheet("8. Material Mix Yield")
    _banner(
        ws,
        8,
        "CONCRETE, MORTAR & STEEL MATERIAL MIX YIELD (PROCUREMENT ABSTRACT)",
        "Auto-calculates Cement (50 kg bags), M-Sand (Cum/CFT), Aggregate (Cum/CFT), Bricks & TMT Steel from Tabs 2–4.",
    )
    _write_table_headers(ws, 4, [
        ("S.No", 7),
        ("Material / Mix Component", 42),
        ("Base Work Qty", 13),
        ("Base Unit", 9),
        ("Yield Factor", 13),
        ("Procurement Qty", 14),
        ("Proc. Unit", 10),
        ("Engineering Basis", 28),
    ])

    if static_qty:
        pcc_expr = static_qty["pcc_cum"]
        rcc_expr = static_qty["rcc_cum"]
        mas_expr = static_qty["mas_cum"]
        pls_expr = static_qty["pls_sqm"]
        stl_expr = static_qty["steel_kg"]
    else:
        fs = f"'{f_ref['sheet_name']}'"
        rs = f"'{r_ref['sheet_name']}'"
        ms = f"'{m_ref['sheet_name']}'"
        pcc_expr = f"={fs}!H{f_ref['code_rows']['FD-03']}"
        rcc_expr = (
            f"={fs}!H{f_ref['code_rows']['FD-04']}"
            f"+{rs}!H{r_ref['code_rows']['RC-01']}"
            f"+{rs}!H{r_ref['code_rows']['RC-02']}"
            f"+{rs}!H{r_ref['code_rows']['RC-03']}"
            f"+{rs}!H{r_ref['code_rows']['RC-04']}"
        )
        mas_expr = f"={ms}!H{m_ref['code_rows']['MS-01']}+{ms}!H{m_ref['code_rows']['MS-02']}+{ms}!H{m_ref['code_rows']['MS-03']}"
        pls_expr = f"={ms}!H{m_ref['code_rows']['PL-01']}+{ms}!H{m_ref['code_rows']['PL-02']}+{ms}!H{m_ref['code_rows']['PL-03']}"
        stl_expr = f"={rs}!H{r_ref['code_rows']['RC-06']}"

    # Material yield rows: (Label, BaseFormula, BaseUnit, Factor, ProcUnit, Basis)
    yield_rows = [
        ("Cement — M25 RCC Concrete (1:1:2 / Design ~8.2 bags/m³)", rcc_expr, "Cum", 8.20, "Bags (50kg)", "410 kg/m³ nominal M25 yield"),
        ("M-Sand — M25 RCC Concrete (0.44 m³/m³)", rcc_expr, "Cum", 0.44, "Cum", "× 35.315 for CFT"),
        ("20 mm Aggregate — M25 RCC Concrete (0.86 m³/m³)", rcc_expr, "Cum", 0.86, "Cum", "× 35.315 for CFT"),
        ("Cement — PCC 1:4:8 Levelling Bed (3.4 bags/m³)", pcc_expr, "Cum", 3.40, "Bags (50kg)", "170 kg/m³ dry vol 1.54×"),
        ("Sand — PCC 1:4:8 (0.47 m³/m³)", pcc_expr, "Cum", 0.47, "Cum", "IS 456 blinding bed"),
        ("40 mm Aggregate — PCC 1:4:8 (0.94 m³/m³)", pcc_expr, "Cum", 0.94, "Cum", "IS 456 blinding bed"),
        ("Clay Bricks / Solid Blocks — Masonry (500 nos/m³)", mas_expr, "Cum", 500.0, "Nos", "Standard 190×90×90 mm + mortar"),
        ("Cement — Masonry Mortar CM 1:6 (1.25 bags/m³)", mas_expr, "Cum", 1.25, "Bags (50kg)", "0.25 m³ wet mortar / m³ wall"),
        ("Sand — Masonry Mortar CM 1:6 (0.27 m³/m³)", mas_expr, "Cum", 0.27, "Cum", "Includes 30% dry bulking"),
        ("Cement — Plastering CM 1:5 avg 15 mm (0.12 bags/m²)", pls_expr, "Sqm", 0.12, "Bags (50kg)", "Internal + External + Ceiling"),
        ("Plaster Sand — CM 1:5 avg 15 mm (0.021 m³/m²)", pls_expr, "Sqm", 0.021, "Cum", "Screened fine river/M-sand"),
        ("TMT Fe-500D Reinforcement Steel (+3% Wastage)", stl_expr, "kg", 1.03, "kg", "Gross procurement weight"),
        ("TMT Fe-500D Steel in Metric Tonnes (MT)", stl_expr, "kg", 0.00103, "MT", "Gross steel ÷ 1000"),
    ]

    for idx, (lbl, b_expr, b_unit, fac, p_unit, basis) in enumerate(yield_rows, start=1):
        r = 4 + idx
        ws.row_dimensions[r].height = 23
        ws.cell(r, 1, idx).alignment = ALIGN_C
        ws.cell(r, 2, lbl).alignment = ALIGN_L
        ws.cell(r, 3, b_expr).alignment = ALIGN_R
        ws.cell(r, 4, b_unit).alignment = ALIGN_C
        ws.cell(r, 5, fac).alignment = ALIGN_R
        ws.cell(r, 6, f"=ROUND(C{r}*E{r}, 2)").alignment = ALIGN_R
        ws.cell(r, 7, p_unit).alignment = ALIGN_C
        ws.cell(r, 8, basis).alignment = ALIGN_L
        for ci in range(1, 9):
            c = ws.cell(r, ci)
            c.border = BORDER_GRID
            if ci == 5:
                c.fill = FILL_INPUT
                c.font = FONT_INPUT
            else:
                c.fill = FILL_CALC
                c.font = FONT_BOLD if ci == 6 else FONT_DATA
        ws.cell(r, 3).number_format = "#,##0.00"
        ws.cell(r, 5).number_format = "#,##0.000"
        ws.cell(r, 6).number_format = "#,##0.00"

    # Summary Procurement Totals (Cement Bags, Total Sand Cum & CFT, Total Agg Cum)
    sr = 4 + len(yield_rows) + 2
    summary_lines = [
        ("TOTAL CEMENT REQUIRED (50 kg BAGS)", "=ROUND(F5+F8+F12+F14, 0)", "Bags", "RCC + PCC + Masonry + Plaster"),
        ("TOTAL SAND REQUIRED (M-SAND / RIVER SAND)", "=ROUND(F6+F9+F13+F15, 2)", "Cum", "Multiply by 35.315 for CFT"),
        ("TOTAL COARSE AGGREGATE (20 mm + 40 mm)", "=ROUND(F7+F10, 2)", "Cum", "Multiply by 35.315 for CFT"),
        ("TOTAL GROSS TMT STEEL (Fe-500D)", "=F17", "MT", "Includes 3% rolling & cutting wastage"),
    ]
    for off, (lbl, fml, unit, note) in enumerate(summary_lines):
        r = sr + off
        ws.row_dimensions[r].height = 23
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        ws.cell(r, 1, lbl).alignment = ALIGN_R
        ws.cell(r, 6, fml).alignment = ALIGN_R
        ws.cell(r, 6).number_format = "#,##0.00"
        ws.cell(r, 7, unit).alignment = ALIGN_C
        ws.cell(r, 8, note).alignment = ALIGN_L
        for ci in range(1, 9):
            ws.cell(r, ci).fill = FILL_GREEN if ci == 6 else FILL_AMBER
            ws.cell(r, ci).font = FONT_BOLD
            ws.cell(r, ci).border = BORDER_TOT

    ws.freeze_panes = "A5"
    _configure_print(ws, landscape=True)
    return {"sheet_name": ws.title}


def build_commercial_summary_sheet(
    wb: Workbook,
    trade_refs: list[tuple[str, str, str, int]],
    builtup_area_sqm: float,
    params: dict[str, Any],
) -> None:
    """Tab 1: Commercial Summary & Abstract of Cost."""
    ws = wb.create_sheet("1. Commercial Summary", 0)
    _banner(
        ws,
        6,
        "MODULAR BUILDING BOQ — COMMERCIAL SUMMARY & ABSTRACT OF COST (IS 1200 / CPWD DSR)",
        "Every building part is separated into its own tab (Tabs 2–9). All totals update live automatically.",
    )
    _write_table_headers(ws, 4, [
        ("S.No", 7),
        ("Building Part / Trade Module", 44),
        ("Source Tab", 24),
        ("Amount (₹)", 16),
        ("% Share", 11),
        ("Remarks", 22),
    ])

    first_r = 5
    last_r = first_r + len(trade_refs) - 1
    sub_r = last_r + 1

    for idx, (label, sname, amt_col, trow) in enumerate(trade_refs, start=1):
        r = first_r + idx - 1
        ws.row_dimensions[r].height = 23
        ws.cell(r, 1, idx).alignment = ALIGN_C
        ws.cell(r, 2, label).alignment = ALIGN_L
        ws.cell(r, 3, sname).alignment = ALIGN_C
        ws.cell(r, 4, f"='{sname}'!{amt_col}{trow}").alignment = ALIGN_R
        ws.cell(r, 4).number_format = "₹#,##0.00"
        ws.cell(r, 5, f"=IF($D${sub_r}>0, D{r}/$D${sub_r}, 0)").alignment = ALIGN_R
        ws.cell(r, 5).number_format = "0.0%"
        ws.cell(r, 6, "Live Tab Link").alignment = ALIGN_L
        for ci in range(1, 7):
            ws.cell(r, ci).fill = FILL_CALC
            ws.cell(r, ci).font = FONT_BOLD if ci == 4 else FONT_DATA
            ws.cell(r, ci).border = BORDER_GRID

    # Direct Cost Sub-Total
    ws.row_dimensions[sub_r].height = 24
    ws.merge_cells(start_row=sub_r, start_column=1, end_row=sub_r, end_column=3)
    ws.cell(sub_r, 1, "DIRECT CONSTRUCTION COST SUB-TOTAL").alignment = ALIGN_R
    ws.cell(sub_r, 4, f"=SUM(D{first_r}:D{last_r})").alignment = ALIGN_R
    ws.cell(sub_r, 4).number_format = "₹#,##0.00"
    ws.cell(sub_r, 5, f"=SUM(E{first_r}:E{last_r})").alignment = ALIGN_R
    ws.cell(sub_r, 5).number_format = "0.0%"
    for ci in range(1, 7):
        ws.cell(sub_r, ci).fill = FILL_TOTAL
        ws.cell(sub_r, ci).font = FONT_BOLD
        ws.cell(sub_r, ci).border = BORDER_TOT

    # Contingency, OHP, Pre-GST, GST, Grand Total
    cont_pct = float(params.get("contingency_pct", 0.05))
    ohp_pct = float(params.get("ohp_pct", 0.10))
    gst_pct = float(params.get("gst_pct", 0.18))

    cont_r = sub_r + 1
    ohp_r = sub_r + 2
    pregst_r = sub_r + 3
    gst_r = sub_r + 4
    grand_r = sub_r + 5

    commercial_rows = [
        (cont_r, "A", "Add: Contingencies & Unforeseen Items (%)", f"=ROUND(D{sub_r}*E{cont_r}, 2)", cont_pct, "Editable %", True),
        (ohp_r,  "B", "Add: Contractor Overheads & Profit (%)", f"=ROUND(D{sub_r}*E{ohp_r}, 2)", ohp_pct, "Editable %", True),
        (pregst_r, "", "SUB-TOTAL BEFORE GST", f"=ROUND(D{sub_r}+D{cont_r}+D{ohp_r}, 2)", "", "Taxable Value", False),
        (gst_r,  "C", "Add: Goods & Services Tax — GST (%)", f"=ROUND(D{pregst_r}*E{gst_r}, 2)", gst_pct, "Works Contract GST", True),
    ]
    for rr, code, lbl, fml, pct, rem, is_inp in commercial_rows:
        ws.row_dimensions[rr].height = 22
        ws.cell(rr, 1, code).alignment = ALIGN_C
        ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=3)
        ws.cell(rr, 2, lbl).alignment = ALIGN_L if code else ALIGN_R
        ws.cell(rr, 4, fml).alignment = ALIGN_R
        ws.cell(rr, 4).number_format = "₹#,##0.00"
        ws.cell(rr, 5, pct).alignment = ALIGN_C
        if is_inp:
            ws.cell(rr, 5).number_format = "0.0%"
        ws.cell(rr, 6, rem).alignment = ALIGN_L
        for ci in range(1, 7):
            ws.cell(rr, ci).fill = FILL_INPUT if (is_inp and ci == 5) else (FILL_TOTAL if not code else FILL_CALC)
            ws.cell(rr, ci).font = FONT_INPUT if (is_inp and ci == 5) else (FONT_BOLD if not code else FONT_DATA)
            ws.cell(rr, ci).border = BORDER_TOT if not code else BORDER_GRID

    ws.row_dimensions[grand_r].height = 26
    ws.merge_cells(start_row=grand_r, start_column=1, end_row=grand_r, end_column=3)
    ws.cell(grand_r, 1, "GRAND TOTAL CONTRACT VALUE (INR)").alignment = ALIGN_R
    ws.cell(grand_r, 4, f"=ROUND(D{pregst_r}+D{gst_r}, 2)").alignment = ALIGN_R
    ws.cell(grand_r, 4).number_format = "₹#,##0.00"
    ws.cell(grand_r, 5, "INR").alignment = ALIGN_C
    ws.cell(grand_r, 6, "All-Inclusive Estimate").alignment = ALIGN_L
    for ci in range(1, 7):
        ws.cell(grand_r, ci).fill = FILL_GREEN
        ws.cell(grand_r, ci).font = FONT_BOLD
        ws.cell(grand_r, ci).border = BORDER_TOT

    # Unit Cost KPIs (per Sqm & per Sqft)
    kpi_r = grand_r + 2
    ws.row_dimensions[kpi_r].height = 22
    ws.cell(kpi_r, 1, "KPI").alignment = ALIGN_C
    ws.cell(kpi_r, 2, "Total Built-up Area (Sq.m) — Editable").alignment = ALIGN_L
    ws.cell(kpi_r, 3, round(builtup_area_sqm, 2)).alignment = ALIGN_R
    ws.cell(kpi_r, 3).number_format = "#,##0.00"
    ws.cell(kpi_r, 4, f"=IF(C{kpi_r}>0, ROUND(D{grand_r}/C{kpi_r}, 2), 0)").alignment = ALIGN_R
    ws.cell(kpi_r, 4).number_format = "₹#,##0.00"
    ws.cell(kpi_r, 5, "₹ / Sq.m").alignment = ALIGN_C
    ws.cell(kpi_r, 6, f"=IF(C{kpi_r}>0, ROUND(D{grand_r}/(C{kpi_r}*10.7639), 2), 0)").alignment = ALIGN_R
    ws.cell(kpi_r, 6).number_format = '"₹"#,##0.00" / Sq.ft"'
    for ci in range(1, 7):
        ws.cell(kpi_r, ci).fill = FILL_INPUT if ci == 3 else FILL_AMBER
        ws.cell(kpi_r, ci).font = FONT_INPUT if ci == 3 else FONT_BOLD
        ws.cell(kpi_r, ci).border = BORDER_GRID

    _configure_print(ws, landscape=True)


# ---------------------------------------------------------------------------
# Master Assembly & Standalone Trade Exports
# ---------------------------------------------------------------------------
def _derive_trade_datasets(
    raw_rows: list[dict[str, Any]],
    params: dict[str, Any],
) -> dict[str, Any]:
    """Convert raw AutoCAD polyline takeoff rows into structured rows for Tabs 2–7."""
    story_ht = float(params.get("story_height_m", 3.30))
    slab_thk = float(params.get("slab_thickness_m", 0.125))
    steel_rate = float(params.get("steel_rate_inr", 82.0))
    conc_rate = float(params.get("conc_rate_inr", 7600.0))

    ftgs = [r for r in raw_rows if r["part"] == "FOOTING"]
    cols = [r for r in raw_rows if r["part"] == "COLUMN"]
    pbeams = [r for r in raw_rows if r["part"] == "PLINTH_BEAM"]
    beams = [r for r in raw_rows if r["part"] == "BEAM"]
    slabs = [r for r in raw_rows if r["part"] == "SLAB"]
    walls = [r for r in raw_rows if r["part"] in ("WALL", "MASONRY")]
    rooms = [r for r in raw_rows if r["part"] == "ROOM"]

    ftg_area = sum(r["nos"] * r["area_m2"] for r in ftgs) or 37.08
    ftg_vol = sum(r["nos"] * r["area_m2"] * r["ht_m"] for r in ftgs) or 20.80
    ftg_shut = sum(r["nos"] * r["perim_m"] * r["ht_m"] for r in ftgs) or 42.84

    col_vol = sum(r["nos"] * r["area_m2"] * (r["ht_m"] or story_ht) for r in cols) or 4.46
    col_shut = sum(r["nos"] * r["perim_m"] * (r["ht_m"] or story_ht) for r in cols) or 49.50

    slab_area = sum(r["nos"] * r["area_m2"] for r in slabs) or 112.00
    slab_vol = round(slab_area * slab_thk, 2)
    beam_vol = round(sum(r["nos"] * r["area_m2"] * (r["ht_m"] or 0.45) for r in beams), 2) or round(slab_area * 0.065, 2)
    plinth_beam_vol = round(sum(r["nos"] * r["area_m2"] * (r["ht_m"] or 0.38) for r in pbeams), 2) or round(slab_area * 0.055, 2)
    stair_vol = 2.65
    total_rcc_vol = round(ftg_vol + col_vol + plinth_beam_vol + beam_vol + slab_vol + stair_vol, 2)
    total_steel_kg = round(total_rcc_vol * 120.0, 2)

    room_area = sum(r["nos"] * r["area_m2"] for r in rooms) or 80.40
    room_perim = sum(r["nos"] * r["perim_m"] for r in rooms) or 84.80
    outer_wall_len = sum(r["nos"] * r["len_m"] for r in walls if r["wid_m"] >= 0.18)
    part_wall_len = sum(r["nos"] * r["len_m"] for r in walls if r["wid_m"] < 0.18)
    outer_perim = round(outer_wall_len, 2) if outer_wall_len > 0 else max(sum(r["perim_m"] for r in slabs), 44.0)
    part_len = round(part_wall_len, 2) if part_wall_len > 0 else round(room_perim * 0.45, 2)

    fdn_rows = [
        ("FD-01", "Earthwork in excavation in foundation trenches & footing pits up to 2.25 m depth", "Cum", 1, round(ftg_area * 1.15, 2), 1.0, 2.00, 320.0, "Includes working space"),
        ("FD-02", "River / M-sand filling & consolidation (150 mm thick) under all footings", "Cum", 1, round(ftg_area * 1.10, 2), 1.0, 0.15, 1650.0, "150 mm sand cushion"),
        ("FD-03", "Plain Cement Concrete (PCC 1:4:8, 100 mm thick) blinding under footings", "Cum", 1, round(ftg_area * 1.10, 2), 1.0, 0.10, 5400.0, "IS 456 levelling bed"),
        ("FD-04", "RCC M25 in Isolated & Combined Footings (from AutoCAD Footing Polylines)", "Cum", 1, round(ftg_vol, 2), 1.0, 1.0, conc_rate, "Auto-summed from Tab 9"),
        ("FD-05", "Centering & shuttering for footing sides including strutting & propping", "Sqm", 1, round(ftg_shut, 2), 1.0, 1.0, 520.0, "Perimeter × Df"),
        ("FD-06", "Backfilling around foundations with excavated earth & watering/compaction", "Cum", 1, round(ftg_area * 1.15 * 1.25, 2), 1.0, 1.0, 240.0, "Net compacted fill"),
    ]

    rcc_rows = [
        ("RC-01", "RCC M25 in Columns (Footing Neck + Superstructure) from AutoCAD Column Polylines", "Cum", 1, round(col_vol, 2), 1.0, 1.0, conc_rate + 600, "Auto-summed from Tab 9"),
        ("RC-02", "RCC M25 in Plinth Beams & Tie Beams (230×380 mm typical)", "Cum", 1, plinth_beam_vol, 1.0, 1.0, conc_rate + 400, "Sub-structure tie grid"),
        ("RC-03", "RCC M25 in Roof / Floor Beams (230×450 mm typical)", "Cum", 1, beam_vol, 1.0, 1.0, conc_rate + 600, "Superstructure beams"),
        ("RC-04", "RCC M25 in Roof Slab & Staircase Waist Slab (from AutoCAD Slab Polylines)", "Cum", 1, round(slab_vol + stair_vol, 2), 1.0, 1.0, conc_rate + 500, f"Slab thk = {slab_thk*1000:.0f} mm"),
        ("RC-05", "Centering, shuttering & staging for Columns, Beams, Slab & Staircase", "Sqm", 1, round(col_shut + slab_area * 1.35, 2), 1.0, 1.0, 650.0, "Plywood shuttering"),
        ("RC-06", "TMT Fe-500D Reinforcement Steel — cutting, bending & binding complete (~120 kg/m³)", "kg", 1, total_steel_kg, 1.0, 1.0, steel_rate, "Linked to Material Yield"),
    ]

    mas_rows = [
        ("MS-01", "230 mm Outer Brick / AAC Block Masonry in CM 1:6 (Perimeter × Story Ht)", "Cum", 1, round(outer_perim, 2), 0.23, story_ht, 6900.0, "External walls"),
        ("MS-02", "115 mm Internal Partition Masonry in CM 1:4", "Cum", 1, part_len, 0.115, story_ht, 7200.0, "Internal partitions"),
        ("MS-03", "Less: Deduction for Doors, Windows & Ventilators (Auto-linked from Tab 5)", "Cum", -1, "='5. Doors & Windows'!F11", 0.23, 1.0, 6900.0, "Live link to Tab 5"),
        ("PL-01", "12 mm Internal Cement Plastering in CM 1:6 on walls (net of openings)", "Sqm", 1, round(room_perim * story_ht * 0.85, 2), 1.0, 1.0, 310.0, "Smooth trowel finish"),
        ("PL-02", "18 mm External Two-Coat Waterproof Plastering in CM 1:4", "Sqm", 1, round(outer_perim * story_ht * 0.90, 2), 1.0, 1.0, 390.0, "Weatherproof coat"),
        ("PL-03", "10 mm Ceiling Plastering in CM 1:4 (from AutoCAD Slab/Room Area)", "Sqm", 1, round(slab_area * 0.92, 2), 1.0, 1.0, 330.0, "Soffit of slab"),
    ]

    flr_rows = []
    for i, rm in enumerate(rooms[:8], start=1):
        flr_rows.append((
            f"FL-{i:02d}",
            f"Flooring — {rm['mark']} ({rm['len_m']}×{rm['wid_m']} m) Vitrified / Anti-Skid Tiles over CM 1:4 bed",
            "Sqm", rm["nos"], rm["area_m2"], 1.0, 1.0, 1450.0, "AutoCAD Room Polyline",
        ))
    flr_rows.extend([
        ("SK-01", "100 mm High Matching Tile / Granite Skirting along room perimeters", "Rmt", 1, round(room_perim * 0.88, 2), 1.0, 1.0, 240.0, "Net of door widths"),
        ("DD-01", "Glazed Ceramic Wall Dado Tiles (up to 2.1 m ht) in Bathrooms & Kitchen", "Sqm", 1, 46.00, 1.0, 1.0, 1380.0, "Wet areas"),
        ("PT-01", "Interior Painting — 2 coats Wall Putty + 1 coat Primer + 2 coats Acrylic Emulsion", "Sqm", 1, round(room_perim * story_ht * 0.85 + slab_area * 0.92, 2), 1.0, 1.0, 195.0, "Walls + Ceilings"),
        ("PT-02", "Exterior Weather-Shield Emulsion Paint (1 coat Primer + 2 coats Apex/WeatherCoat)", "Sqm", 1, round(outer_perim * story_ht * 0.90, 2), 1.0, 1.0, 215.0, "External facade"),
    ])

    mep_rows = [
        ("EL-01", "Concealed PVC Conduit Wiring — Light, Fan & 6A Socket Points (FRLS Copper)", "Pts", 48, 1.0, 1.0, 1.0, 1150.0, "IS 694 FRLS wires"),
        ("EL-02", "16A Power Points (AC, Geyser, Kitchen Appliances) with modular switches", "Pts", 10, 1.0, 1.0, 1.0, 1850.0, "4 sq.mm copper"),
        ("EL-03", "Main Distribution Board (8-Way TPN MCB/RCCB) & Copper Plate Earthing", "Lot", 1, 1.0, 1.0, 1.0, 24500.0, "Safety & protection"),
        ("PL-01", "CPVC / UPVC Concealed Water Supply & SWR Drainage Piping (Bathrooms + Kitchen)", "Lot", 1, 1.0, 1.0, 1.0, 68000.0, "Pressure tested"),
        ("PL-02", "Sanitaryware & CP Fittings (EWC, Wash Basin, Diverter, Mixers, Sink)", "Set", 2, 1.0, 1.0, 1.0, 32000.0, "2 Bathrooms + Kitchen"),
        ("WT-01", "Overhead PVC Water Tank (2000 Ltr) & Underground RCC Sump (6000 Ltr) with Motor", "Lot", 1, 1.0, 1.0, 1.0, 95000.0, "Complete plumbing"),
    ]

    return {
        "fdn_rows": fdn_rows,
        "rcc_rows": rcc_rows,
        "mas_rows": mas_rows,
        "flr_rows": flr_rows,
        "mep_rows": mep_rows,
        "builtup_area_sqm": slab_area,
    }


def _wb_bytes(wb: Workbook) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def compile_modular_boq_bytes(
    raw_rows: list[dict[str, Any]] | None = None,
    params: dict[str, Any] | None = None,
) -> dict[str, bytes]:
    """Compile Master 9-Tab Modular BOQ + Standalone Individual Part Workbooks in memory."""
    p = params or {}
    raw = raw_rows if raw_rows else default_autocad_sample_rows()
    ds = _derive_trade_datasets(raw, p)
    out: dict[str, bytes] = {}

    # 1. Master 9-Tab Workbook
    wb = Workbook()
    d_ws = wb.active
    f_ref = _build_standard_trade_sheet(wb, "2. Foundation & Earthwork", "MODULE 1: FOUNDATION & EARTHWORK BOQ", "Excavation, Sand Fill, PCC 1:4:8, Footing M25 Concrete, Shuttering & Backfill.", ds["fdn_rows"])
    r_ref = _build_standard_trade_sheet(wb, "3. RCC Structure & Steel", "MODULE 2: RCC STRUCTURE & STEEL BOQ", "Columns, Plinth Beams, Floor Beams, Roof Slabs, Staircase, Shuttering & Fe-500D Steel.", ds["rcc_rows"])
    m_ref = _build_standard_trade_sheet(wb, "4. Masonry & Plastering", "MODULE 3: MASONRY & PLASTERING BOQ", "Outer 230 mm & Partition 115 mm Masonry with Live Opening Deductions from Tab 5.", ds["mas_rows"])
    dw_ref = build_doors_windows_sheet(wb, p)
    fl_ref = _build_standard_trade_sheet(wb, "6. Flooring, Tiles & Paint", "MODULE 5: FLOORING, TILES & PAINTING BOQ", "Room-by-Room AutoCAD Takeoff Schedule — Vitrified Tiles, Skirting, Dado & Paint.", ds["flr_rows"])
    ep_ref = _build_standard_trade_sheet(wb, "7. Electrical & Plumbing", "MODULE 6: ELECTRICAL, PLUMBING & WATER TANK BOQ", "Concealed Wiring, DB, Earthing, CPVC/SWR Piping, Sanitaryware & Water Storage.", ds["mep_rows"])
    build_material_mix_yield_sheet(wb, f_ref, r_ref, m_ref)
    build_autocad_raw_sheet(wb, raw)
    trade_refs = [
        ("Foundation & Earthwork", f_ref["sheet_name"], "J", f_ref["total_row"]),
        ("RCC Structure & Steel", r_ref["sheet_name"], "J", r_ref["total_row"]),
        ("Masonry & Plastering", m_ref["sheet_name"], "J", m_ref["total_row"]),
        ("Doors, Windows & Openings", dw_ref["sheet_name"], dw_ref["amount_col"], dw_ref["total_row"]),
        ("Flooring, Tiles & Painting", fl_ref["sheet_name"], "J", fl_ref["total_row"]),
        ("Electrical, Plumbing & Misc", ep_ref["sheet_name"], "J", ep_ref["total_row"]),
    ]
    build_commercial_summary_sheet(wb, trade_refs, ds["builtup_area_sqm"], p)
    wb.remove(d_ws)
    out[MASTER_FILENAME] = _wb_bytes(wb)

    # 2. Standalone Part Exports (zero broken links)
    standalone_specs = [
        ("BOQ_Foundation_Earthwork.xlsx", "2. Foundation & Earthwork", "MODULE 1: FOUNDATION & EARTHWORK BOQ", ds["fdn_rows"]),
        ("BOQ_RCC_Structure.xlsx", "3. RCC Structure & Steel", "MODULE 2: RCC STRUCTURE & STEEL BOQ", ds["rcc_rows"]),
        ("BOQ_Flooring_Paint.xlsx", "6. Flooring, Tiles & Paint", "MODULE 5: FLOORING, TILES & PAINTING BOQ", ds["flr_rows"]),
        ("BOQ_Electrical_Plumbing.xlsx", "7. Electrical & Plumbing", "MODULE 6: ELECTRICAL & PLUMBING BOQ", ds["mep_rows"]),
    ]
    for fname, sname, title, rdata in standalone_specs:
        wbs = Workbook()
        dw = wbs.active
        _build_standard_trade_sheet(wbs, sname, title, "Standalone Trade Export — AutoBOQ 2D Modular System.", rdata)
        wbs.remove(dw)
        out[fname] = _wb_bytes(wbs)

    # Doors & Windows standalone
    wb_dw = Workbook()
    dw = wb_dw.active
    build_doors_windows_sheet(wb_dw, p)
    wb_dw.remove(dw)
    out["BOQ_Doors_Windows.xlsx"] = _wb_bytes(wb_dw)

    # Masonry & Plaster standalone (includes Doors & Windows tab so opening deduction link never #REF!s)
    wb_mp = Workbook()
    dw = wb_mp.active
    _build_standard_trade_sheet(wb_mp, "4. Masonry & Plastering", "MODULE 3: MASONRY & PLASTERING BOQ", "Includes Doors & Windows tab for live opening deduction.", ds["mas_rows"])
    build_doors_windows_sheet(wb_mp, p)
    wb_mp.remove(dw)
    wb_mp.active = 0
    out["BOQ_Masonry_Plaster.xlsx"] = _wb_bytes(wb_mp)

    return out


def generate_modular_boq_files() -> list[str]:
    """Generate and save Master 9-Tab Modular BOQ (.xlsx + .pdf) and standalone part files to disk."""
    os.makedirs(EXPORTS_DIR, exist_ok=True)
    files_map = compile_modular_boq_bytes()
    saved: list[str] = []

    root_xlsx = os.path.join(BASE_DIR, MASTER_FILENAME)
    root_pdf = os.path.join(BASE_DIR, os.path.splitext(MASTER_FILENAME)[0] + ".pdf")
    exp_pdf = os.path.join(EXPORTS_DIR, os.path.splitext(MASTER_FILENAME)[0] + ".pdf")

    with open(root_xlsx, "wb") as fh:
        fh.write(files_map[MASTER_FILENAME])
    saved.append(root_xlsx)

    for fname, data in files_map.items():
        ep = os.path.join(EXPORTS_DIR, fname)
        with open(ep, "wb") as fh:
            fh.write(data)
        saved.append(ep)

    pdf_ps1 = os.path.join(BASE_DIR, "export_xlsx_to_pdf.ps1")
    if os.name == "nt" and os.path.exists(pdf_ps1):
        try:
            subprocess.run(
                ["powershell", "-ExecutionPolicy", "Bypass", "-File", pdf_ps1, "-XlsxPath", root_xlsx, "-PdfPath", root_pdf],
                check=True,
                capture_output=True,
                text=True,
            )
            if os.path.exists(root_pdf):
                shutil.copy2(root_pdf, exp_pdf)
                saved.extend([root_pdf, exp_pdf])
        except Exception as exc:
            print("Warning: Could not export Modular BOQ PDF:", exc)

    return saved


if __name__ == "__main__":
    for p in generate_modular_boq_files():
        print("Generated:", p)
