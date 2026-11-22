"""
SmartCivilTools-Style Complete 12-Tab AutoCAD-to-Excel Modular BOQ Generator
Covers 100% of Quick Building Estimator PRO & AutoCAD Civil LISP items:
  1.  Commercial Summary (Materials / Labour / Joinery / Additional Works + Cont/OHP/GST + Rs/Sqft)
  2.  Client BOQ & Quotation (Client Unit Rates incl. Contingency & OHP)
  3.  Material Takeoff (Supply) — All 22 Supply Materials (Theoretical vs Wastage) + Thumb Checks
  4.  Foundation & Basement (Excavation, Sand, PCC, Footing, Basement Fill, Anti-Termite, DPC)
  5.  RCC Structure & Steel (Col, Plinth Beam, Beam, Slab, Staircase, Shuttering, TMT, Binding Wire, Cover Blocks)
  6.  Masonry & Plastering (230/115 Walls, Live Opening Deduction from Tab 7, Plaster, Weathering Course, Curing)
  7.  Doors & Windows (MD1, BD1, Bath, W1, W2, W3, V1 + Opening Area + Wood Volume CFT)
  8.  Flooring, Tiles & Paint (Rooms, Bath, Parking Pavers, Kitchen/Bath Dado, Granite Counter/Stair, Handrail, Paints)
  9.  Electrical & Plumbing (Wiring, DB, Toilet Sanitary, Taps/Bib Cocks, Kitchen SS Sink/Swan Neck, Civil Labour)
  10. Additional Works (UG Sump, Septic Tank, OHT+Stand, Compound Wall, Gate, Balcony Railing, Ramp, RWH Pit)
  11. Bank Loan Estimate (L x B x D Abstract + Stage-Wise Payment Drawdown Schedule)
  12. AutoCAD Raw Takeoff (Polylines/Lines from .DXF or AutoBOQ_SmartTools.lsp command AD)
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
    if "PLINTH" in u:
        return "PLINTH_BEAM"
    if any(k in u for k in ("BEAM", "BM")):
        return "BEAM"
    if any(k in u for k in ("SLAB", "ROOF")):
        return "SLAB"
    if "STAIR" in u:
        return "STAIRCASE"
    if any(k in u for k in ("TANK", "SUMP")):
        return "WATER_TANK"
    if any(k in u for k in ("WALL", "MAS", "BRICK")):
        return "WALL"
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
    """Parse AutoCAD .CSV (from AutoBOQ_SmartTools.lsp AD command) or .DXF into normalized rows."""
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
            default_ht_map = {
                "FOOTING": 0.525, "COLUMN": 3.3, "PLINTH_BEAM": 0.38, "BEAM": 0.45,
                "SLAB": 0.125, "STAIRCASE": 0.175, "WATER_TANK": 2.0, "WALL": 3.0, "OPENING": 2.1, "ROOM": 3.0,
            }
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


def _build_trade_sheet(
    wb: Workbook,
    sheet_name: str,
    title: str,
    subtitle: str,
    rows_data: list[tuple],
) -> dict[str, Any]:
    """Standard 12-col BOQ sheet (A..L): Code, Trade, Description & First-Principles Formula, Unit, Nos, L, B, D, Qty, Rate, Amount, Basis."""
    ws = wb.create_sheet(sheet_name)
    _banner(ws, 12, title, subtitle)
    _write_table_headers(ws, 4, [
        ("Code", 9),
        ("Trade Filter", 15),
        ("Work Description & First-Principles Formula", 48),
        ("Unit", 8),
        ("Nos", 7),
        ("L / Area", 9),
        ("B", 7),
        ("D / H", 7),
        ("Quantity", 11),
        ("Rate (₹)", 12),
        ("Amount (₹)", 15),
        ("Scope Bucket", 16),
    ])

    start_r = 5
    code_rows: dict[str, int] = {}
    for idx, (code, trade, desc, unit, nos, lv, bv, dv, rate, scope) in enumerate(rows_data):
        r = start_r + idx
        code_rows[code] = r
        ws.row_dimensions[r].height = 30 if len(desc) > 80 else 23
        ws.cell(r, 1, code).alignment = ALIGN_C
        ws.cell(r, 2, trade).alignment = ALIGN_C
        ws.cell(r, 3, desc).alignment = ALIGN_L
        ws.cell(r, 4, unit).alignment = ALIGN_C
        ws.cell(r, 5, nos).alignment = ALIGN_C
        ws.cell(r, 6, lv).alignment = ALIGN_R
        ws.cell(r, 7, bv).alignment = ALIGN_R
        ws.cell(r, 8, dv).alignment = ALIGN_R
        ws.cell(r, 9, f"=ROUND(E{r}*F{r}*G{r}*H{r}, 2)").alignment = ALIGN_R
        ws.cell(r, 10, rate).alignment = ALIGN_R
        ws.cell(r, 11, f"=ROUND(I{r}*J{r}, 2)").alignment = ALIGN_R
        ws.cell(r, 12, scope).alignment = ALIGN_C

        for ci in range(1, 13):
            c = ws.cell(r, ci)
            c.border = BORDER_GRID
            is_fml = isinstance(c.value, str) and c.value.startswith("=")
            if ci in (5, 6, 7, 8, 10) and not is_fml:
                c.fill = FILL_INPUT
                c.font = FONT_INPUT
            else:
                c.fill = FILL_CALC
                c.font = FONT_BOLD if ci == 11 else FONT_DATA
        for ci in (6, 7, 8, 9):
            ws.cell(r, ci).number_format = "#,##0.00"
        ws.cell(r, 10).number_format = "₹#,##0.00"
        ws.cell(r, 11).number_format = "₹#,##0.00"

    end_r = start_r + len(rows_data) - 1
    tot_r = end_r + 1
    ws.row_dimensions[tot_r].height = 24
    ws.merge_cells(start_row=tot_r, start_column=1, end_row=tot_r, end_column=10)
    ws.cell(tot_r, 1, f"SUB-TOTAL — {sheet_name.upper()}").alignment = ALIGN_R
    ws.cell(tot_r, 11, f"=SUM(K{start_r}:K{end_r})").alignment = ALIGN_R
    ws.cell(tot_r, 11).number_format = "₹#,##0.00"
    for ci in range(1, 13):
        ws.cell(tot_r, ci).fill = FILL_GREEN if ci == 11 else FILL_TOTAL
        ws.cell(tot_r, ci).font = FONT_BOLD
        ws.cell(tot_r, ci).border = BORDER_TOT

    ws.freeze_panes = "A5"
    _configure_print(ws, landscape=True)
    return {"sheet_name": sheet_name, "total_row": tot_r, "code_rows": code_rows, "amount_col": "K", "qty_col": "I"}


# ---------------------------------------------------------------------------
# Specialized Sheet Builders
# ---------------------------------------------------------------------------
def build_autocad_raw_sheet(wb: Workbook, raw_rows: list[dict[str, Any]]) -> dict[str, Any]:
    ws = wb.create_sheet("12. AutoCAD Raw Takeoff")
    _banner(
        ws,
        10,
        "AUTOCAD RAW POLYLINE TAKEOFF & ROOM DIMENSIONS SCHEDULE",
        "Imported from .DXF or AutoBOQ_SmartTools.lsp (Command: AD — Footing, Column, Beam, Slab, Staircase, Water Tank, Wall, Room).",
    )
    _write_table_headers(ws, 4, [
        ("S.No", 7), ("Part Category", 14), ("AutoCAD Layer", 15), ("Mark / Room", 18),
        ("Nos", 7), ("Area (m²)", 11), ("Perimeter (m)", 12), ("Length (m)", 10),
        ("Width (m)", 10), ("Depth/Ht (m)", 11),
    ])
    for idx, rd in enumerate(raw_rows, start=1):
        r = 4 + idx
        ws.row_dimensions[r].height = 21
        vals = [idx, rd["part"], rd["layer"], rd["mark"], rd["nos"], rd["area_m2"], rd["perim_m"], rd["len_m"], rd["wid_m"], rd["ht_m"]]
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
    return {"sheet_name": ws.title}


def build_doors_windows_sheet(wb: Workbook) -> dict[str, Any]:
    ws = wb.create_sheet("7. Doors & Windows")
    _banner(
        ws,
        10,
        "MODULE 4: DOORS, WINDOWS & OPENINGS SCHEDULE (JOINERY + WOOD CFT)",
        "Opening Area (Col F) automatically feeds Masonry Deductions on Tab 6; Wood Volume (Col G) feeds Material Takeoff on Tab 3.",
    )
    _write_table_headers(ws, 4, [
        ("Mark", 8), ("Opening Description (Frame, Shutter, Grill & Hardware)", 44),
        ("Nos", 7), ("Width (m)", 9), ("Height (m)", 9), ("Area (m²)", 11),
        ("Wood (CFT)", 11), ("Rate / No (₹)", 13), ("Amount (₹)", 15), ("Trade", 14),
    ])
    openings = [
        ("MD1", "Main Entrance Teakwood Panelled Door (1.20×2.10 m) with brass fittings & polish", 1, 1.20, 2.10, 8.50, 32000.0, "Carpentry / Doors"),
        ("BD1", "Bedroom Flush Door (1.00×2.10 m) with Sal/WPC frame, laminate & mortise lock", 3, 1.00, 2.10, 4.80, 14800.0, "Carpentry / Doors"),
        ("D2",  "Bathroom Waterproof WPC Door (0.80×2.10 m) with SS hinges & tower bolts", 2, 0.80, 2.10, 0.00, 9800.0, "Carpentry / Doors"),
        ("W1",  "UPVC 3-Track Sliding Window (1.50×1.35 m) with 5 mm glass, mesh & MS grill", 4, 1.50, 1.35, 0.00, 13500.0, "Carpentry / Doors"),
        ("W2",  "UPVC 2-Track Sliding Window (1.20×1.20 m) with MS safety grill (Kitchen)", 2, 1.20, 1.20, 0.00, 10600.0, "Carpentry / Doors"),
        ("W3",  "UPVC Staircase / Lobby Fixed + Casement Window (1.00×1.50 m) with grill", 1, 1.00, 1.50, 0.00, 9500.0, "Carpentry / Doors"),
        ("V1",  "UPVC Louvered Bathroom Ventilator (0.60×0.60 m) with exhaust fan opening", 2, 0.60, 0.60, 0.00, 3800.0, "Carpentry / Doors"),
    ]
    start_r = 5
    tot_wood_cft = 0.0
    tot_open_area = 0.0
    tot_amt = 0.0
    for idx, (mk, desc, nos, w, h, cft_per, rate, trade) in enumerate(openings):
        r = start_r + idx
        tot_open_area += round(nos * w * h, 2)
        tot_wood_cft += round(nos * cft_per, 2)
        tot_amt += round(nos * rate, 2)
        ws.row_dimensions[r].height = 24
        ws.cell(r, 1, mk).alignment = ALIGN_C
        ws.cell(r, 2, desc).alignment = ALIGN_L
        ws.cell(r, 3, nos).alignment = ALIGN_C
        ws.cell(r, 4, w).alignment = ALIGN_R
        ws.cell(r, 5, h).alignment = ALIGN_R
        ws.cell(r, 6, f"=ROUND(C{r}*D{r}*E{r}, 2)").alignment = ALIGN_R
        ws.cell(r, 7, f"=ROUND(C{r}*{cft_per}, 2)").alignment = ALIGN_R
        ws.cell(r, 8, rate).alignment = ALIGN_R
        ws.cell(r, 9, f"=ROUND(C{r}*H{r}, 2)").alignment = ALIGN_R
        ws.cell(r, 10, trade).alignment = ALIGN_C
        for ci in range(1, 11):
            c = ws.cell(r, ci)
            c.border = BORDER_GRID
            if ci in (3, 4, 5, 8):
                c.fill = FILL_INPUT
                c.font = FONT_INPUT
            else:
                c.fill = FILL_CALC
                c.font = FONT_BOLD if ci in (6, 9) else FONT_DATA
        for ci in (4, 5, 6, 7):
            ws.cell(r, ci).number_format = "#,##0.00"
        ws.cell(r, 8).number_format = "₹#,##0.00"
        ws.cell(r, 9).number_format = "₹#,##0.00"

    end_r = start_r + len(openings) - 1
    tot_r = end_r + 1
    ws.row_dimensions[tot_r].height = 24
    ws.merge_cells(start_row=tot_r, start_column=1, end_row=tot_r, end_column=5)
    ws.cell(tot_r, 1, "TOTAL OPENINGS AREA (m²), WOOD (CFT) & JOINERY COST").alignment = ALIGN_R
    ws.cell(tot_r, 6, f"=SUM(F{start_r}:F{end_r})").alignment = ALIGN_R
    ws.cell(tot_r, 6).number_format = "#,##0.00"
    ws.cell(tot_r, 7, f"=SUM(G{start_r}:G{end_r})").alignment = ALIGN_R
    ws.cell(tot_r, 7).number_format = "#,##0.00"
    ws.cell(tot_r, 9, f"=SUM(I{start_r}:I{end_r})").alignment = ALIGN_R
    ws.cell(tot_r, 9).number_format = "₹#,##0.00"
    for ci in range(1, 11):
        ws.cell(tot_r, ci).fill = FILL_GREEN if ci == 9 else FILL_TOTAL
        ws.cell(tot_r, ci).font = FONT_BOLD
        ws.cell(tot_r, ci).border = BORDER_TOT

    ws.freeze_panes = "A5"
    _configure_print(ws, landscape=True)
    return {
        "sheet_name": ws.title,
        "total_row": tot_r,
        "area_col": "F",
        "wood_col": "G",
        "amount_col": "I",
        "tot_open_area": round(tot_open_area, 2),
        "tot_wood_cft": round(tot_wood_cft, 2),
        "tot_amt": round(tot_amt, 2),
    }


def build_material_takeoff_sheet(
    wb: Workbook,
    metrics: dict[str, float],
    builtup_sqft: float,
    steel_rate: float = 82.0,
) -> dict[str, Any]:
    """Tab 3: Complete 22-Item Material Take-Off Statement (Theoretical vs Procurement with Wastage) + Thumb Checks."""
    ws = wb.create_sheet("3. Material Takeoff (Supply)")
    _banner(
        ws,
        8,
        "MATERIAL TAKE-OFF STATEMENT (SUPPLY ONLY) — THEORETICAL vs. PROCUREMENT WITH WASTAGE",
        "Covers all 22 core supply materials (Cement, M-Sand, P-Sand, 20mm/40mm Agg, Steel, Bricks, Timber, Putty, Paints, Binding Wire, Cover Blocks).",
    )
    _write_table_headers(ws, 4, [
        ("S.No", 6),
        ("Material Name (Supply Only)", 42),
        ("Theoretical Qty", 14),
        ("Wastage %", 10),
        ("Qty w/ Wastage", 14),
        ("Unit", 11),
        ("Unit Rate (₹)", 12),
        ("Total Supply Cost (₹)", 16),
    ])

    rcc_v = metrics["rcc_cum"]
    pcc_v = metrics["pcc_cum"]
    mas_v = metrics["mas_cum"]
    pls_a = metrics["pls_sqm"]
    int_p = metrics["int_paint_sqm"]
    ext_p = metrics["ext_paint_sqm"]
    enm_p = metrics["enamel_sqm"]
    stl_k = metrics["steel_kg"]
    shut_a = metrics["shuttering_sqm"]
    grvl_v = metrics["gravel_cum"]
    wood_c = metrics["wood_cft"]

    cement_th = round(rcc_v * 8.2 + pcc_v * 3.4 + mas_v * 1.25 + pls_a * 0.12, 1)
    msand_th = round(rcc_v * 0.44 + pcc_v * 0.47 + mas_v * 0.27, 2)
    psand_th = round(pls_a * 0.021, 2)
    agg20_th = round(rcc_v * 0.86, 2)
    agg40_th = round(pcc_v * 0.94, 2)
    bricks_th = round(mas_v * 500.0, 0)
    w_cem_th = round((int_p + ext_p) * 0.008, 1)
    putty_th = round(int_p * 0.045, 1)       # 40 kg bags
    iprim_th = round(int_p / 11.0, 1)        # Ltrs
    ipnt_th = round(int_p / 7.5, 1)          # Ltrs (2 coats)
    eprim_th = round(ext_p / 10.0, 1)
    epnt_th = round(ext_p / 7.0, 1)
    enml_th = round(enm_p / 8.0, 1)
    redox_th = round(enm_p / 11.0, 1)
    thin_th = round((enml_th + redox_th) * 0.25, 1)
    nails_th = round(shut_a * 0.12, 1)       # kg
    bwire_th = round((stl_k / 1000.0) * 10.0, 1)  # 10 kg per MT of rebar
    cblks_th = round(shut_a * 3.0, 0)        # 3 cover blocks / Sqm

    # (Material, TheoQty, WastagePct, Unit, RateINR)
    mats = [
        ("Cement (OPC 53 / PPC — 50 kg bags)", cement_th, 0.03, "Bags", 395.0),
        ("M-Sand (Concrete & Masonry)", msand_th, 0.05, "Cum", 1750.0),
        ("P-Sand (Fine Screened Plastering Sand)", psand_th, 0.05, "Cum", 1950.0),
        ("Coarse Aggregate 20 mm (RCC Concrete)", agg20_th, 0.05, "Cum", 1450.0),
        ("Coarse Aggregate 40 mm (PCC & Base)", agg40_th, 0.05, "Cum", 1300.0),
        ("Gravel / Filling Earth (Foundations & Basement)", grvl_v, 0.10, "Cum", 550.0),
        ("Reinforcement Steel (TMT Fe-500D / Fe-550D)", stl_k, 0.03, "kg", steel_rate),
        ("Bricks / Solid Masonry Blocks (190×90×90 mm)", bricks_th, 0.05, "Nos", 9.50),
        ("Joinery Timber / Teak & Sal Wood Volume", wood_c, 0.05, "CFT", 2600.0),
        ("White Cement (50 kg bags — base finish)", w_cem_th, 0.03, "Bags", 1050.0),
        ("Wall Putty (40 kg bags — 2 coats internal)", putty_th, 0.05, "Bags", 820.0),
        ("Interior Wall Primer (Water-Based)", iprim_th, 0.05, "Ltrs", 210.0),
        ("Interior Acrylic Emulsion Paint (2 coats)", ipnt_th, 0.05, "Ltrs", 420.0),
        ("Exterior Weather-Shield Primer", eprim_th, 0.05, "Ltrs", 240.0),
        ("Exterior Weatherproof Emulsion Paint (2 coats)", epnt_th, 0.05, "Ltrs", 480.0),
        ("Synthetic Enamel Paint (Grills & Frames)", enml_th, 0.05, "Ltrs", 340.0),
        ("Red Oxide / Zinc Chromate Metal Primer", redox_th, 0.05, "Ltrs", 220.0),
        ("Paint Thinner (Mineral Turpentine Oil)", thin_th, 0.05, "Ltrs", 140.0),
        ("Steel Wire Nails (Shuttering & Formwork)", nails_th, 0.05, "kg", 95.0),
        ("Rebar Binding Wire (GI 18-Gauge — 10 kg/MT)", bwire_th, 0.05, "kg", 110.0),
        ("Concrete Cover Blocks (20 / 25 / 40 / 50 mm)", cblks_th, 0.05, "Nos", 6.00),
        ("Centring & Shuttering Material Contact Area", shut_a, 0.00, "Sqm", 260.0),
    ]

    start_r = 5
    tot_supply = 0.0
    for idx, (name, theo, wast, unit, rate) in enumerate(mats, start=1):
        r = start_r + idx - 1
        qty_w = round(theo * (1.0 + wast), 2)
        tot_supply += round(qty_w * rate, 2)
        ws.row_dimensions[r].height = 22
        ws.cell(r, 1, idx).alignment = ALIGN_C
        ws.cell(r, 2, name).alignment = ALIGN_L
        ws.cell(r, 3, theo).alignment = ALIGN_R
        ws.cell(r, 4, wast).alignment = ALIGN_C
        ws.cell(r, 5, f"=ROUND(C{r}*(1+D{r}), 2)").alignment = ALIGN_R
        ws.cell(r, 6, unit).alignment = ALIGN_C
        ws.cell(r, 7, rate).alignment = ALIGN_R
        ws.cell(r, 8, f"=ROUND(E{r}*G{r}, 2)").alignment = ALIGN_R
        for ci in range(1, 9):
            c = ws.cell(r, ci)
            c.border = BORDER_GRID
            if ci in (3, 4, 7):
                c.fill = FILL_INPUT
                c.font = FONT_INPUT
            else:
                c.fill = FILL_CALC
                c.font = FONT_BOLD if ci in (5, 8) else FONT_DATA
        ws.cell(r, 3).number_format = "#,##0.00"
        ws.cell(r, 4).number_format = "0.0%"
        ws.cell(r, 5).number_format = "#,##0.00"
        ws.cell(r, 7).number_format = "₹#,##0.00"
        ws.cell(r, 8).number_format = "₹#,##0.00"

    end_r = start_r + len(mats) - 1
    tot_r = end_r + 1
    ws.row_dimensions[tot_r].height = 25
    ws.merge_cells(start_row=tot_r, start_column=1, end_row=tot_r, end_column=7)
    ws.cell(tot_r, 1, "TOTAL ESTIMATED MATERIAL PROCUREMENT COST (SUPPLY ONLY)").alignment = ALIGN_R
    ws.cell(tot_r, 8, f"=SUM(H{start_r}:H{end_r})").alignment = ALIGN_R
    ws.cell(tot_r, 8).number_format = "₹#,##0.00"
    for ci in range(1, 9):
        ws.cell(tot_r, ci).fill = FILL_GREEN if ci == 8 else FILL_TOTAL
        ws.cell(tot_r, ci).font = FONT_BOLD
        ws.cell(tot_r, ci).border = BORDER_TOT

    # Key Material Consumption Thumb Checks (per Sq.ft)
    chk_r = tot_r + 2
    ws.row_dimensions[chk_r].height = 23
    ws.merge_cells(start_row=chk_r, start_column=1, end_row=chk_r, end_column=8)
    ws.cell(chk_r, 1, "KEY MATERIAL CONSUMPTION THUMB-RULE CHECKS (PER SQ.FT BUILT-UP AREA)").font = FONT_SEC
    ws.cell(chk_r, 1).fill = FILL_SEC
    for ci in range(1, 9):
        ws.cell(chk_r, ci).fill = FILL_SEC
        ws.cell(chk_r, ci).border = BORDER_HDR

    checks = [
        ("Cement Consumption (Bags / Sq.ft)", f"=ROUND(E5/{builtup_sqft:.2f}, 2)", "Bags / Sq.ft", "Typical: 0.40 - 0.55 Bags/Sq.ft"),
        ("TMT Steel Consumption (kg / Sq.ft)", f"=ROUND(E11/{builtup_sqft:.2f}, 2)", "kg / Sq.ft", "Typical: 3.50 - 4.80 kg/Sq.ft"),
        ("Bricks / Blocks (Nos / Sq.ft)", f"=ROUND(E12/{builtup_sqft:.2f}, 1)", "Nos / Sq.ft", "Typical: 18.0 - 28.0 Nos/Sq.ft"),
    ]
    for off, (lbl, fml, u, norm) in enumerate(checks, start=1):
        rr = chk_r + off
        ws.row_dimensions[rr].height = 21
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=4)
        ws.cell(rr, 1, lbl).alignment = ALIGN_L
        ws.cell(rr, 5, fml).alignment = ALIGN_R
        ws.cell(rr, 5).number_format = "#,##0.00"
        ws.cell(rr, 6, u).alignment = ALIGN_C
        ws.merge_cells(start_row=rr, start_column=7, end_row=rr, end_column=8)
        ws.cell(rr, 7, norm).alignment = ALIGN_L
        for ci in range(1, 9):
            ws.cell(rr, ci).fill = FILL_AMBER if ci == 5 else FILL_CALC
            ws.cell(rr, ci).font = FONT_BOLD if ci == 5 else FONT_DATA
            ws.cell(rr, ci).border = BORDER_GRID

    ws.freeze_panes = "A5"
    _configure_print(ws, landscape=True)
    return {"sheet_name": ws.title, "total_row": tot_r, "amount_col": "H", "tot_supply": round(tot_supply, 2), "mats": mats}


def build_client_boq_sheet(
    wb: Workbook,
    all_trade_rows: list[tuple],
    cont_pct: float,
    ohp_pct: float,
) -> dict[str, Any]:
    """Tab 2: Client-Facing Quotation BOQ (Unit Rates bake in Contingency % + Overheads & Profit %)."""
    ws = wb.create_sheet("2. Client BOQ & Quotation")
    _banner(
        ws,
        7,
        "CLIENT QUOTATION & PRICED BOQ (COMMERCIAL UNIT RATES INCL. CONTINGENCY & OHP)",
        f"Commercial Markup Applied: +{cont_pct*100:.1f}% Contingency & +{ohp_pct*100:.1f}% Overheads/Profit baked transparently into Client Unit Rates.",
    )
    ws.row_dimensions[3].height = 20
    ws.cell(3, 1, "Markup Factor:").font = FONT_BOLD
    ws.cell(3, 2, round(1.0 + cont_pct + ohp_pct, 4)).font = FONT_INPUT
    ws.cell(3, 2).fill = FILL_INPUT
    ws.cell(3, 2).number_format = "0.000"
    ws.cell(3, 3, "Client Unit Rate = Internal Rate × $B$3 (Editable Markup)").font = FONT_DATA

    _write_table_headers(ws, 4, [
        ("Code", 9), ("Trade", 16), ("Item Description (Client BOQ)", 52),
        ("Unit", 8), ("Quantity", 11), ("Client Rate (₹)", 14), ("Client Amount (₹)", 16),
    ])

    start_r = 5
    for idx, (code, trade, desc, unit, qty, int_rate) in enumerate(all_trade_rows):
        r = start_r + idx
        ws.row_dimensions[r].height = 24
        ws.cell(r, 1, code).alignment = ALIGN_C
        ws.cell(r, 2, trade).alignment = ALIGN_C
        ws.cell(r, 3, desc).alignment = ALIGN_L
        ws.cell(r, 4, unit).alignment = ALIGN_C
        ws.cell(r, 5, qty).alignment = ALIGN_R
        ws.cell(r, 6, f"=ROUND({int_rate}*$B$3, 2)").alignment = ALIGN_R
        ws.cell(r, 7, f"=ROUND(E{r}*F{r}, 2)").alignment = ALIGN_R
        for ci in range(1, 8):
            c = ws.cell(r, ci)
            c.border = BORDER_GRID
            c.fill = FILL_CALC
            c.font = FONT_BOLD if ci == 7 else FONT_DATA
        ws.cell(r, 5).number_format = "#,##0.00"
        ws.cell(r, 6).number_format = "₹#,##0.00"
        ws.cell(r, 7).number_format = "₹#,##0.00"

    end_r = start_r + len(all_trade_rows) - 1
    tot_r = end_r + 1
    ws.row_dimensions[tot_r].height = 25
    ws.merge_cells(start_row=tot_r, start_column=1, end_row=tot_r, end_column=6)
    ws.cell(tot_r, 1, "TOTAL CLIENT QUOTATION VALUE (BEFORE GST)").alignment = ALIGN_R
    ws.cell(tot_r, 7, f"=SUM(G{start_r}:G{end_r})").alignment = ALIGN_R
    ws.cell(tot_r, 7).number_format = "₹#,##0.00"
    for ci in range(1, 8):
        ws.cell(tot_r, ci).fill = FILL_GREEN if ci == 7 else FILL_TOTAL
        ws.cell(tot_r, ci).font = FONT_BOLD
        ws.cell(tot_r, ci).border = BORDER_TOT

    ws.freeze_panes = "A5"
    _configure_print(ws, landscape=True)
    return {"sheet_name": ws.title, "total_row": tot_r}


def build_bank_loan_estimate_sheet(
    wb: Workbook,
    grand_total_expr: str | float,
    builtup_sqm: float,
) -> dict[str, Any]:
    """Tab 11: Bank Loan Abstract & Stage-Wise Payment Schedule."""
    ws = wb.create_sheet("11. Bank Loan Estimate")
    _banner(
        ws,
        6,
        "BANK LOAN ESTIMATE ABSTRACT & STAGE-WISE PAYMENT SCHEDULE (L × B × D BASIS)",
        "Formatted for Housing / Commercial Construction Loan Sanction & Stage-Wise Bank Disbursement.",
    )
    _write_table_headers(ws, 4, [
        ("Stage", 8), ("Construction Milestone / Disbursement Stage", 46),
        ("Stage %", 10), ("Stage Amount (₹)", 16), ("Cumulative (₹)", 16), ("Verification Required", 22),
    ])

    stages = [
        ("1", "Advance & Foundation / Plinth Beam Completion (incl. Excavation, PCC, Footings, Backfill)", 0.20, "Engineer Plinth Certificate"),
        ("2", "Superstructure Masonry up to Lintel / Roof Beam Level (Columns, Walls, Door/Window Frames)", 0.20, "Lintel Stage Inspection"),
        ("3", "Roof Slab Casting & Staircase Completion (RCC M25, Formwork, Fe-500D Steel)", 0.25, "Slab Casting Photo + Bill"),
        ("4", "Internal/External Plastering, Flooring Tiles, Granite, Electrical & Plumbing Rough-in", 0.25, "Finishing Stage Audit"),
        ("5", "Painting, Doors/Windows Shutters, CP Fittings, Additional Works & Handover", 0.10, "Completion Certificate"),
    ]
    base_val = grand_total_expr
    for idx, (st_no, desc, pct, ver) in enumerate(stages):
        r = 5 + idx
        ws.row_dimensions[r].height = 26
        ws.cell(r, 1, st_no).alignment = ALIGN_C
        ws.cell(r, 2, desc).alignment = ALIGN_L
        ws.cell(r, 3, pct).alignment = ALIGN_C
        ws.cell(r, 3).number_format = "0.0%"
        base_ref = base_val if isinstance(base_val, str) else str(base_val)
        ws.cell(r, 4, f"=ROUND({base_ref}*C{r}, 2)").alignment = ALIGN_R
        ws.cell(r, 4).number_format = "₹#,##0.00"
        ws.cell(r, 5, f"=SUM($D$5:D{r})").alignment = ALIGN_R
        ws.cell(r, 5).number_format = "₹#,##0.00"
        ws.cell(r, 6, ver).alignment = ALIGN_L
        for ci in range(1, 7):
            c = ws.cell(r, ci)
            c.border = BORDER_GRID
            c.fill = FILL_INPUT if ci == 3 else FILL_CALC
            c.font = FONT_INPUT if ci == 3 else (FONT_BOLD if ci in (4, 5) else FONT_DATA)

    tot_r = 10
    ws.row_dimensions[tot_r].height = 25
    ws.merge_cells(start_row=tot_r, start_column=1, end_row=tot_r, end_column=2)
    ws.cell(tot_r, 1, f"TOTAL SANCTIONED ESTIMATE ({builtup_sqm:.1f} Sq.m / {builtup_sqm*10.7639:.0f} Sq.ft)").alignment = ALIGN_R
    ws.cell(tot_r, 3, "=SUM(C5:C9)").alignment = ALIGN_C
    ws.cell(tot_r, 3).number_format = "0.0%"
    ws.cell(tot_r, 4, "=SUM(D5:D9)").alignment = ALIGN_R
    ws.cell(tot_r, 4).number_format = "₹#,##0.00"
    ws.cell(tot_r, 5, "=E9").alignment = ALIGN_R
    ws.cell(tot_r, 5).number_format = "₹#,##0.00"
    for ci in range(1, 7):
        ws.cell(tot_r, ci).fill = FILL_GREEN if ci == 4 else FILL_TOTAL
        ws.cell(tot_r, ci).font = FONT_BOLD
        ws.cell(tot_r, ci).border = BORDER_TOT

    ws.freeze_panes = "A5"
    _configure_print(ws, landscape=True)
    return {"sheet_name": ws.title}


def build_commercial_summary_sheet(
    wb: Workbook,
    trade_refs: list[tuple[str, str, str, int]],
    mat_ref: dict[str, Any],
    builtup_area_sqm: float,
    params: dict[str, Any],
) -> dict[str, Any]:
    """Tab 1: Commercial Summary & Executive Abstract (linked live to Tabs 3–10)."""
    ws = wb.create_sheet("1. Commercial Summary", 0)
    _banner(
        ws,
        6,
        "MODULAR BUILDING BOQ — EXECUTIVE COMMERCIAL SUMMARY (IS 1200 / CPWD DSR)",
        "Complete 12-Tab System: Separate Building Parts + 22-Item Material Takeoff + Client Quotation + Bank Loan Estimate.",
    )
    _write_table_headers(ws, 4, [
        ("S.No", 7), ("Building Part / Trade Module", 44), ("Source Tab", 26),
        ("Amount (₹)", 16), ("% Share", 11), ("Remarks", 22),
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

    ws.row_dimensions[sub_r].height = 24
    ws.merge_cells(start_row=sub_r, start_column=1, end_row=sub_r, end_column=3)
    ws.cell(sub_r, 1, "DIRECT CONSTRUCTION COST SUB-TOTAL (ALL PARTS)").alignment = ALIGN_R
    ws.cell(sub_r, 4, f"=SUM(D{first_r}:D{last_r})").alignment = ALIGN_R
    ws.cell(sub_r, 4).number_format = "₹#,##0.00"
    ws.cell(sub_r, 5, f"=SUM(E{first_r}:E{last_r})").alignment = ALIGN_R
    ws.cell(sub_r, 5).number_format = "0.0%"
    for ci in range(1, 7):
        ws.cell(sub_r, ci).fill = FILL_TOTAL
        ws.cell(sub_r, ci).font = FONT_BOLD
        ws.cell(sub_r, ci).border = BORDER_TOT

    cont_pct = float(params.get("contingency_pct", 0.05))
    ohp_pct = float(params.get("ohp_pct", 0.10))
    gst_pct = float(params.get("gst_pct", 0.18))

    cont_r, ohp_r, pregst_r, gst_r, grand_r = sub_r + 1, sub_r + 2, sub_r + 3, sub_r + 4, sub_r + 5
    commercial_rows = [
        (cont_r, "A", "Add: Contingencies & Escalation (%)", f"=ROUND(D{sub_r}*E{cont_r}, 2)", cont_pct, "Editable %", True),
        (ohp_r,  "B", "Add: Contractor Overheads & Profit (%)", f"=ROUND(D{sub_r}*E{ohp_r}, 2)", ohp_pct, "Editable %", True),
        (pregst_r, "", "SUB-TOTAL BEFORE GST (CLIENT QUOTATION BASIS)", f"=ROUND(D{sub_r}+D{cont_r}+D{ohp_r}, 2)", "", "Matches Tab 2", False),
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
    ws.cell(grand_r, 6, "Turnkey Estimate").alignment = ALIGN_L
    for ci in range(1, 7):
        ws.cell(grand_r, ci).fill = FILL_GREEN
        ws.cell(grand_r, ci).font = FONT_BOLD
        ws.cell(grand_r, ci).border = BORDER_TOT

    # Materials Supply Memo & Cost per Sq.ft KPI
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

    mat_memo_r = kpi_r + 1
    ws.row_dimensions[mat_memo_r].height = 22
    ws.cell(mat_memo_r, 1, "MAT").alignment = ALIGN_C
    ws.merge_cells(start_row=mat_memo_r, start_column=2, end_row=mat_memo_r, end_column=3)
    ws.cell(mat_memo_r, 2, "Material Procurement Cost Only (22 Supply Items — Tab 3)").alignment = ALIGN_L
    ws.cell(mat_memo_r, 4, f"='{mat_ref['sheet_name']}'!{mat_ref['amount_col']}{mat_ref['total_row']}").alignment = ALIGN_R
    ws.cell(mat_memo_r, 4).number_format = "₹#,##0.00"
    ws.cell(mat_memo_r, 5, f"=IF($D${sub_r}>0, D{mat_memo_r}/$D${sub_r}, 0)").alignment = ALIGN_R
    ws.cell(mat_memo_r, 5).number_format = "0.0%"
    ws.cell(mat_memo_r, 6, "Supply-Only Share").alignment = ALIGN_L
    for ci in range(1, 7):
        ws.cell(mat_memo_r, ci).fill = FILL_SEC
        ws.cell(mat_memo_r, ci).font = FONT_BOLD if ci == 4 else FONT_DATA
        ws.cell(mat_memo_r, ci).border = BORDER_GRID

    _configure_print(ws, landscape=True)
    return {"grand_total_cell": f"'1. Commercial Summary'!D{grand_r}"}


# ---------------------------------------------------------------------------
# Dataset Derivation (100% Aligned with Quick Building Estimator PRO Items)
# ---------------------------------------------------------------------------
def _derive_trade_datasets(
    raw_rows: list[dict[str, Any]],
    params: dict[str, Any],
) -> dict[str, Any]:
    story_ht = float(params.get("story_height_m", 3.30))
    slab_thk = float(params.get("slab_thickness_m", 0.125))
    steel_rate = float(params.get("steel_rate_inr", 82.0))
    conc_rate = float(params.get("conc_rate_inr", 7600.0))

    ftgs = [r for r in raw_rows if r["part"] == "FOOTING"]
    cols = [r for r in raw_rows if r["part"] == "COLUMN"]
    pbeams = [r for r in raw_rows if r["part"] == "PLINTH_BEAM"]
    beams = [r for r in raw_rows if r["part"] == "BEAM"]
    slabs = [r for r in raw_rows if r["part"] == "SLAB"]
    stairs = [r for r in raw_rows if r["part"] == "STAIRCASE"]
    walls = [r for r in raw_rows if r["part"] in ("WALL", "MASONRY")]
    rooms = [r for r in raw_rows if r["part"] == "ROOM"]

    ftg_area = sum(r["nos"] * r["area_m2"] for r in ftgs) or 37.08
    ftg_vol = sum(r["nos"] * r["area_m2"] * r["ht_m"] for r in ftgs) or 20.80
    ftg_shut = sum(r["nos"] * r["perim_m"] * r["ht_m"] for r in ftgs) or 42.84

    col_vol = sum(r["nos"] * r["area_m2"] * (r["ht_m"] or story_ht) for r in cols) or 4.46
    col_shut = sum(r["nos"] * r["perim_m"] * (r["ht_m"] or story_ht) for r in cols) or 49.50

    slab_area = sum(r["nos"] * r["area_m2"] for r in slabs) or 112.00
    builtup_sqft = round(slab_area * 10.7639, 1)
    slab_vol = round(slab_area * slab_thk, 2)
    beam_vol = round(sum(r["nos"] * r["area_m2"] * (r["ht_m"] or 0.45) for r in beams), 2) or round(slab_area * 0.065, 2)
    plinth_beam_vol = round(sum(r["nos"] * r["area_m2"] * (r["ht_m"] or 0.38) for r in pbeams), 2) or round(slab_area * 0.055, 2)
    stair_vol = round(sum(r["nos"] * r["area_m2"] * (r["ht_m"] or 0.175) for r in stairs), 2) or 2.65
    total_rcc_vol = round(ftg_vol + col_vol + plinth_beam_vol + beam_vol + slab_vol + stair_vol, 2)
    total_steel_kg = round(total_rcc_vol * 120.0, 2)
    total_shut_sqm = round(ftg_shut + col_shut + slab_area * 1.35, 2)

    room_perim = sum(r["nos"] * r["perim_m"] for r in rooms) or 84.80
    outer_wall_len = sum(r["nos"] * r["len_m"] for r in walls if r["wid_m"] >= 0.18)
    part_wall_len = sum(r["nos"] * r["len_m"] for r in walls if r["wid_m"] < 0.18)
    outer_perim = round(outer_wall_len, 2) if outer_wall_len > 0 else max(sum(r["perim_m"] for r in slabs), 44.0)
    part_len = round(part_wall_len, 2) if part_wall_len > 0 else round(room_perim * 0.45, 2)

    pcc_cum = round(ftg_area * 1.10 * 0.10, 2)
    gravel_cum = round(slab_area * 0.45, 2)
    open_area_est = 25.04  # MD1 + 3 BD1 + 2 D2 + 4 W1 + 2 W2 + 1 W3 + 2 V1
    mas_cum = round(outer_perim * 0.23 * story_ht + part_len * 0.115 * story_ht - open_area_est * 0.23, 2)
    int_pls_sqm = round(room_perim * story_ht * 0.85, 2)
    ext_pls_sqm = round(outer_perim * story_ht * 0.90, 2)
    ceil_pls_sqm = round(slab_area * 0.92, 2)
    pls_sqm = round(int_pls_sqm + ext_pls_sqm + ceil_pls_sqm, 2)

    # Tuple: (Code, TradeFilter, Description, Unit, Nos, L, B, D, Rate, ScopeBucket)
    fdn_rows = [
        ("FD-01", "Civil Labour", "Earth excavation for foundations & trenches (L×B×2.0m depth)", "Cum", 1, round(ftg_area * 1.15, 2), 1.0, 2.00, 320.0, "Labour & Works"),
        ("FD-02", "Civil Labour", "River / M-sand cushion filling (150 mm) under all footings", "Cum", 1, round(ftg_area * 1.10, 2), 1.0, 0.15, 1650.0, "Labour & Works"),
        ("FD-03", "Civil Labour", "Plain Cement Concrete PCC 1:4:8 (100 mm thick) levelling bed", "Cum", 1, round(ftg_area * 1.10, 2), 1.0, 0.10, 5400.0, "Labour & Works"),
        ("FD-04", "Civil Labour", "RCC M25 in Isolated & Combined Footings (from AutoCAD Polylines)", "Cum", 1, round(ftg_vol, 2), 1.0, 1.0, conc_rate, "Labour & Works"),
        ("FD-05", "Civil Labour", "Centering & shuttering for footing sides (Perimeter × Df)", "Sqm", 1, round(ftg_shut, 2), 1.0, 1.0, 520.0, "Labour & Works"),
        ("FD-06", "Civil Labour", "Backfilling around foundations & basement gravel/earth filling (0.45m)", "Cum", 1, gravel_cum, 1.0, 1.0, 620.0, "Labour & Works"),
        ("FD-07", "Civil Labour", "Pre-construction Anti-Termite soil & plinth chemical treatment (IS 6313)", "Sqm", 1, round(slab_area, 2), 1.0, 1.0, 135.0, "Labour & Works"),
        ("FD-08", "Civil Labour", "Damp Proof Course (DPC 40 mm CM 1:2 + waterproofing) & Plinth Protection", "Sqm", 1, round(outer_perim * 0.60, 2), 1.0, 1.0, 580.0, "Labour & Works"),
    ]

    rcc_rows = [
        ("RC-01", "Civil Labour", "RCC M25 in Columns (Fdn Neck + Superstructure) from AutoCAD Polylines", "Cum", 1, round(col_vol, 2), 1.0, 1.0, conc_rate + 600, "Labour & Works"),
        ("RC-02", "Civil Labour", "RCC M25 in Plinth Beams & Tie Beams (230×380 mm)", "Cum", 1, plinth_beam_vol, 1.0, 1.0, conc_rate + 400, "Labour & Works"),
        ("RC-03", "Civil Labour", "RCC M25 in Roof / Floor Beams (230×450 mm)", "Cum", 1, beam_vol, 1.0, 1.0, conc_rate + 600, "Labour & Works"),
        ("RC-04", "Civil Labour", "RCC M25 in Roof Slab & Staircase Waist Slab + Steps", "Cum", 1, round(slab_vol + stair_vol, 2), 1.0, 1.0, conc_rate + 500, "Labour & Works"),
        ("RC-05", "Civil Labour", "Centering, shuttering & staging for Columns, Beams, Slab & Staircase", "Sqm", 1, round(col_shut + slab_area * 1.35, 2), 1.0, 1.0, 650.0, "Labour & Works"),
        ("RC-06", "Civil Labour", "TMT Fe-500D Reinforcement Steel — supply, cutting, bending & binding (~120 kg/m³)", "kg", 1, total_steel_kg, 1.0, 1.0, steel_rate, "Labour & Works"),
        ("RC-07", "Civil Labour", "Rebar Binding Wire (GI 18-G), Concrete Cover Blocks & Shuttering Wire Nails", "Lot", 1, 1.0, 1.0, 1.0, 14500.0, "Labour & Works"),
    ]

    mas_rows = [
        ("MS-01", "Civil Labour", "230 mm Outer Brick / Solid Block Masonry in CM 1:6 (Perimeter × Story Ht)", "Cum", 1, outer_perim, 0.23, story_ht, 6900.0, "Labour & Works"),
        ("MS-02", "Civil Labour", "115 mm Internal Partition Masonry in CM 1:4", "Cum", 1, part_len, 0.115, story_ht, 7200.0, "Labour & Works"),
        ("MS-03", "Civil Labour", "Less: Deduction for Doors, Windows & Ventilators (Auto-linked from Tab 7)", "Cum", -1, "='7. Doors & Windows'!F12", 0.23, 1.0, 6900.0, "Labour & Works"),
        ("PL-01", "Civil Labour", "12 mm Internal Cement Plastering in CM 1:6 on walls (net of openings)", "Sqm", 1, int_pls_sqm, 1.0, 1.0, 310.0, "Labour & Works"),
        ("PL-02", "Civil Labour", "18 mm External Two-Coat Waterproof Plastering in CM 1:4", "Sqm", 1, ext_pls_sqm, 1.0, 1.0, 390.0, "Labour & Works"),
        ("PL-03", "Civil Labour", "10 mm Ceiling Plastering in CM 1:4 (Soffit of slab)", "Sqm", 1, ceil_pls_sqm, 1.0, 1.0, 330.0, "Labour & Works"),
        ("WC-01", "Civil Labour", "Terrace Weathering Course (75 mm brick jelly lime + cool pressed clay tiles & curing)", "Sqm", 1, round(slab_area, 2), 1.0, 1.0, 920.0, "Labour & Works"),
    ]

    flr_rows = []
    for i, rm in enumerate(rooms[:6], start=1):
        is_bath = "bath" in rm["mark"].lower() or "toilet" in rm["mark"].lower()
        tile_desc = "Anti-Skid Ceramic Bathroom Floor Tiles" if is_bath else "Vitrified 600×600 Floor Tiles"
        flr_rows.append((
            f"FL-{i:02d}", "Tiles & Granite",
            f"Flooring — {rm['mark']} ({rm['len_m']}×{rm['wid_m']} m) {tile_desc} over CM 1:4 bed",
            "Sqm", rm["nos"], rm["area_m2"], 1.0, 1.0, 1320.0 if is_bath else 1480.0, "Labour & Works",
        ))
    flr_rows.extend([
        ("FL-07", "Tiles & Granite", "Parking Tiles / Heavy-Duty Interlocking Pavers (car porch & setback)", "Sqm", 1, 28.00, 1.0, 1.0, 980.0, "Labour & Works"),
        ("FL-08", "Tiles & Granite", "Kitchen Dado Glazed Ceramic Wall Tiles (0.6 m above counter) & Bathroom Wall Tiles (2.1 m ht)", "Sqm", 1, 48.00, 1.0, 1.0, 1380.0, "Labour & Works"),
        ("FL-09", "Tiles & Granite", "100 mm Matching Tile / Granite Skirting along room perimeters", "Rmt", 1, round(room_perim * 0.88, 2), 1.0, 1.0, 240.0, "Labour & Works"),
        ("GR-01", "Tiles & Granite", "Kitchen Counter Black Galaxy Granite Slab (20 mm bullnosed) with vertical support", "Sqm", 1, 4.20, 1.0, 1.0, 4200.0, "Labour & Works"),
        ("GR-02", "Tiles & Granite", "Staircase Granite Treads (20 mm double-nosed) & Risers with anti-skid grooves", "Sqm", 1, 16.50, 1.0, 1.0, 3400.0, "Labour & Works"),
        ("GR-03", "Carpentry / Doors", "Staircase SS 304 / Matte-Black MS Handrail (1.05 m ht) fixed on steps & landings", "Rmt", 1, 11.50, 1.0, 1.0, 2400.0, "Labour & Works"),
        ("PT-01", "Painting Labour", "Internal Painting — 2 coats Wall Putty + 1 coat Primer + 2 coats Acrylic Emulsion", "Sqm", 1, round(int_pls_sqm + ceil_pls_sqm, 2), 1.0, 1.0, 195.0, "Labour & Works"),
        ("PT-02", "Painting Labour", "External Painting — 1 coat Weather-Shield Primer + 2 coats Weatherproof Emulsion", "Sqm", 1, ext_pls_sqm, 1.0, 1.0, 215.0, "Labour & Works"),
        ("PT-03", "Painting Labour", "Synthetic Enamel Painting on MS Grills, Gates & Frames (1 coat Red Oxide + 2 coats Enamel)", "Sqm", 1, 24.00, 1.0, 1.0, 185.0, "Labour & Works"),
    ])

    mep_rows = [
        ("EL-01", "Electrical Labour", "Concealed FRLS PVC Conduit Wiring — Light, Fan & 6A Modular Points", "Pts", 48, 1.0, 1.0, 1.0, 1150.0, "Labour & Works"),
        ("EL-02", "Electrical Labour", "16A Power Points (AC, Geyser, Kitchen) + 8-Way TPN DB & Copper Plate Earthing", "Pts", 12, 1.0, 1.0, 1.0, 2850.0, "Labour & Works"),
        ("PL-01", "Plumbing Labour", "Toilet Plumbing & Sanitaryware (CPVC/SWR Piping, EWC, Wash Basin, Concealed Diverter)", "Set", 2, 1.0, 1.0, 1.0, 44000.0, "Labour & Works"),
        ("PL-02", "Plumbing Labour", "Bathroom Water Taps, Bib Cocks, Angle Valves, Health Faucets & Floor Traps", "Set", 2, 1.0, 1.0, 1.0, 9500.0, "Labour & Works"),
        ("PL-03", "Plumbing Labour", "Kitchen Plumbing, SS 304 Sink (600×450 mm), Swan Neck Tap & RO/Drain Provision", "Job", 1, 1.0, 1.0, 1.0, 18500.0, "Labour & Works"),
        ("LB-01", "Civil Labour", "Civil General Site Supervision, Curing & Bar-Bending Misc Labour Allowance (Sq.ft basis)", "Sqft", 1, builtup_sqft, 1.0, 1.0, 45.0, "Labour & Works"),
    ]

    add_rows = [
        ("AD-01", "Additional Works", "Underground RCC Water Sump (6,000 Ltr) with waterproofing, plaster & CI manhole cover", "Job", 1, 1.0, 1.0, 1.0, 92000.0, "Additional Works"),
        ("AD-02", "Additional Works", "Septic Tank (10-User RCC / Masonry) with baffle walls, vent pipe & Soak Pit", "Job", 1, 1.0, 1.0, 1.0, 78000.0, "Additional Works"),
        ("AD-03", "Additional Works", "Overhead PVC Water Tank (2,000 Ltr, 3-Layer) incl. MS/RCC Stand & Plumbing Connections", "Job", 1, 1.0, 1.0, 1.0, 34000.0, "Additional Works"),
        ("AD-04", "Additional Works", "Compound Wall (230 mm masonry, 1.8 m ht, foundation, plinth beam & 2-side plaster)", "Rmt", 1, 36.0, 1.0, 1.0, 4800.0, "Additional Works"),
        ("AD-05", "Additional Works", "Main MS Compound Gate (3.6 m wide × 1.8 m ht) with pillars, rollers & enamel painting", "Job", 1, 1.0, 1.0, 1.0, 58000.0, "Additional Works"),
        ("AD-06", "Additional Works", "Balcony SS 304 / Toughened Glass Railing (1.05 m high)", "Rmt", 1, 6.50, 1.0, 1.0, 3800.0, "Additional Works"),
        ("AD-07", "Additional Works", "Car Parking Entry Ramp (RCC M20 with anti-skid grooves)", "Sqm", 1, 9.00, 1.0, 1.0, 1650.0, "Additional Works"),
        ("AD-08", "Additional Works", "Rainwater Harvesting Pit (1.2 m dia × 2.5 m deep, rings, gravel/charcoal filter & recharge pipe)", "Job", 1, 1.0, 1.0, 1.0, 26000.0, "Additional Works"),
    ]

    metrics = {
        "rcc_cum": total_rcc_vol,
        "pcc_cum": pcc_cum,
        "mas_cum": mas_cum,
        "pls_sqm": pls_sqm,
        "int_paint_sqm": round(int_pls_sqm + ceil_pls_sqm, 2),
        "ext_paint_sqm": ext_pls_sqm,
        "enamel_sqm": 24.0,
        "steel_kg": total_steel_kg,
        "shuttering_sqm": total_shut_sqm,
        "gravel_cum": gravel_cum,
        "wood_cft": 22.90,  # MD1 (8.5) + 3*BD1 (14.4)
    }

    return {
        "fdn_rows": fdn_rows,
        "rcc_rows": rcc_rows,
        "mas_rows": mas_rows,
        "flr_rows": flr_rows,
        "mep_rows": mep_rows,
        "add_rows": add_rows,
        "metrics": metrics,
        "builtup_area_sqm": slab_area,
        "builtup_sqft": builtup_sqft,
    }


def _eval_rows_for_preview(rows_data: list[tuple], open_area: float = 25.04) -> list[dict[str, Any]]:
    out = []
    for code, trade, desc, unit, nos, lv, bv, dv, rate, scope in rows_data:
        l_num = open_area if isinstance(lv, str) else float(lv)
        qty = round(float(nos) * l_num * float(bv) * float(dv), 2)
        amt = round(qty * float(rate), 2)
        fml = f"{nos} × {l_num:.2f} × {bv} × {dv} = {qty:.2f} {unit}"
        out.append({
            "Code": code,
            "Trade": trade,
            "Scope": scope,
            "Description": desc,
            "Formula": fml,
            "Qty": qty,
            "Unit": unit,
            "Rate (₹)": rate,
            "Amount (₹)": amt,
        })
    return out


def get_modular_boq_preview(
    raw_rows: list[dict[str, Any]] | None = None,
    params: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Return structured KPIs & itemized rows for Streamlit Tab 4 Dashboard & Scope/Trade Filters."""
    p = params or {}
    raw = raw_rows if raw_rows else default_autocad_sample_rows()
    ds = _derive_trade_datasets(raw, p)
    cont = float(p.get("contingency_pct", 0.05))
    ohp = float(p.get("ohp_pct", 0.10))
    gst = float(p.get("gst_pct", 0.18))

    # Build temporary workbook to read exact Doors/Windows & Material totals
    tmp_wb = Workbook()
    dw_ref = build_doors_windows_sheet(tmp_wb)
    mat_ref = build_material_takeoff_sheet(tmp_wb, ds["metrics"], ds["builtup_sqft"], float(p.get("steel_rate_inr", 82.0)))

    items: list[dict[str, Any]] = []
    for grp in (ds["fdn_rows"], ds["rcc_rows"], ds["mas_rows"], ds["flr_rows"], ds["mep_rows"], ds["add_rows"]):
        items.extend(_eval_rows_for_preview(grp, dw_ref["tot_open_area"]))

    # Append Doors & Windows items for filtering
    items.append({
        "Code": "DW-ALL",
        "Trade": "Carpentry / Doors",
        "Scope": "Doors & Windows",
        "Description": f"Doors, Windows & Ventilators Complete Schedule (MD1, BD1, Bath, W1–W3, V1 | {dw_ref['tot_open_area']} m² | {dw_ref['tot_wood_cft']} CFT wood)",
        "Formula": f"Opening Area = {dw_ref['tot_open_area']} m²; Wood = {dw_ref['tot_wood_cft']} CFT",
        "Qty": 15.0,
        "Unit": "Nos",
        "Rate (₹)": round(dw_ref["tot_amt"] / 15.0, 2),
        "Amount (₹)": dw_ref["tot_amt"],
    })

    # Append Supply Material rows for "Materials Only" scope filter
    for idx, (mname, theo, wast, unit, rate) in enumerate(mat_ref["mats"], start=1):
        qw = round(theo * (1.0 + wast), 2)
        items.append({
            "Code": f"MT-{idx:02d}",
            "Trade": "Materials Supply",
            "Scope": "Materials Only",
            "Description": mname,
            "Formula": f"Theoretical {theo} × (1 + {wast*100:.0f}% wastage) = {qw} {unit}",
            "Qty": qw,
            "Unit": unit,
            "Rate (₹)": rate,
            "Amount (₹)": round(qw * rate, 2),
        })

    labour_works_tot = sum(r["Amount (₹)"] for r in items if r["Scope"] == "Labour & Works")
    dw_tot = dw_ref["tot_amt"]
    add_tot = sum(r["Amount (₹)"] for r in items if r["Scope"] == "Additional Works")
    mat_tot = mat_ref["tot_supply"]
    direct_tot = round(labour_works_tot + dw_tot + add_tot, 2)
    pregst = round(direct_tot * (1.0 + cont + ohp), 2)
    grand = round(pregst * (1.0 + gst), 2)

    return {
        "items": items,
        "kpis": {
            "materials_supply": round(mat_tot, 2),
            "labour_works": round(labour_works_tot, 2),
            "doors_windows": round(dw_tot, 2),
            "additional_works": round(add_tot, 2),
            "direct_cost": direct_tot,
            "pregst": pregst,
            "grand_total": grand,
            "rate_per_sqft": round(grand / max(ds["builtup_sqft"], 1.0), 2),
            "builtup_sqft": ds["builtup_sqft"],
        },
    }


def _wb_bytes(wb: Workbook) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def compile_modular_boq_bytes(
    raw_rows: list[dict[str, Any]] | None = None,
    params: dict[str, Any] | None = None,
) -> dict[str, bytes]:
    """Compile Master 12-Tab Modular BOQ + Standalone Individual Workbooks in memory (0 #REF!)."""
    p = params or {}
    raw = raw_rows if raw_rows else default_autocad_sample_rows()
    ds = _derive_trade_datasets(raw, p)
    cont_pct = float(p.get("contingency_pct", 0.05))
    ohp_pct = float(p.get("ohp_pct", 0.10))
    steel_rate = float(p.get("steel_rate_inr", 82.0))
    out: dict[str, bytes] = {}

    # 1. Master 12-Tab Workbook
    wb = Workbook()
    d_ws = wb.active

    mat_ref = build_material_takeoff_sheet(wb, ds["metrics"], ds["builtup_sqft"], steel_rate)
    f_ref = _build_trade_sheet(wb, "4. Foundation & Basement", "MODULE 1: FOUNDATION, EARTHWORK & BASEMENT BOQ", "Excavation, Sand Fill, PCC 1:4:8, Footing RCC, Basement Gravel Fill, Anti-Termite & DPC.", ds["fdn_rows"])
    r_ref = _build_trade_sheet(wb, "5. RCC Structure & Steel", "MODULE 2: RCC STRUCTURE, STAIRCASE & STEEL BOQ", "Columns, Plinth Beams, Floor Beams, Roof Slab, Staircase, Shuttering, TMT Steel, Binding Wire & Cover Blocks.", ds["rcc_rows"])
    m_ref = _build_trade_sheet(wb, "6. Masonry & Plastering", "MODULE 3: MASONRY, PLASTERING & WEATHERING COURSE BOQ", "230/115 mm Masonry with Live Opening Deductions (Tab 7), Plastering & Terrace Weathering Course.", ds["mas_rows"])
    dw_ref = build_doors_windows_sheet(wb)
    fl_ref = _build_trade_sheet(wb, "8. Flooring, Tiles & Paint", "MODULE 5: FLOORING, TILES, GRANITE & PAINTING BOQ", "Room Tiles, Parking Pavers, Dado, Skirting, Kitchen/Stair Granite, Handrail & Interior/Exterior/Enamel Paints.", ds["flr_rows"])
    ep_ref = _build_trade_sheet(wb, "9. Electrical & Plumbing", "MODULE 6: ELECTRICAL, PLUMBING & GENERAL LABOUR BOQ", "Concealed Wiring, DB, Toilet Sanitaryware, Taps/Bib Cocks, Kitchen SS Sink & Civil Labour.", ds["mep_rows"])
    ad_ref = _build_trade_sheet(wb, "10. Additional Works", "MODULE 7: ADDITIONAL EXTERNAL WORKS BOQ", "Underground Sump, Septic Tank, Overhead Tank, Compound Wall, Gate, Balcony Railing, Ramp & RWH Pit.", ds["add_rows"])

    # Flat client BOQ rows
    client_rows = []
    for grp in (ds["fdn_rows"], ds["rcc_rows"], ds["mas_rows"], ds["flr_rows"], ds["mep_rows"], ds["add_rows"]):
        for row in _eval_rows_for_preview(grp, dw_ref["tot_open_area"]):
            client_rows.append((row["Code"], row["Trade"], row["Description"], row["Unit"], row["Qty"], row["Rate (₹)"]))
    client_rows.append(("DW-01", "Carpentry / Doors", "Complete Doors, Windows & Ventilators Joinery Package (MD1, BD1, Bath, W1–W3, V1)", "Lot", 1.0, dw_ref["tot_amt"]))
    build_client_boq_sheet(wb, client_rows, cont_pct, ohp_pct)

    trade_refs = [
        ("Foundation, Earthwork & Basement", f_ref["sheet_name"], "K", f_ref["total_row"]),
        ("RCC Structure, Staircase & Steel", r_ref["sheet_name"], "K", r_ref["total_row"]),
        ("Masonry, Plastering & Weathering Course", m_ref["sheet_name"], "K", m_ref["total_row"]),
        ("Doors, Windows & Joinery Schedule", dw_ref["sheet_name"], dw_ref["amount_col"], dw_ref["total_row"]),
        ("Flooring, Tiles, Granite & Painting", fl_ref["sheet_name"], "K", fl_ref["total_row"]),
        ("Electrical, Plumbing & General Labour", ep_ref["sheet_name"], "K", ep_ref["total_row"]),
        ("Additional Works (Sump, Septic, Wall, Gate, RWH)", ad_ref["sheet_name"], "K", ad_ref["total_row"]),
    ]
    sum_ref = build_commercial_summary_sheet(wb, trade_refs, mat_ref, ds["builtup_area_sqm"], p)
    build_bank_loan_estimate_sheet(wb, sum_ref["grand_total_cell"], ds["builtup_area_sqm"])
    build_autocad_raw_sheet(wb, raw)
    wb.remove(d_ws)

    # Enforce logical tab ordering 1..12
    order = [
        "1. Commercial Summary",
        "2. Client BOQ & Quotation",
        "3. Material Takeoff (Supply)",
        "4. Foundation & Basement",
        "5. RCC Structure & Steel",
        "6. Masonry & Plastering",
        "7. Doors & Windows",
        "8. Flooring, Tiles & Paint",
        "9. Electrical & Plumbing",
        "10. Additional Works",
        "11. Bank Loan Estimate",
        "12. AutoCAD Raw Takeoff",
    ]
    wb._sheets = [wb[n] for n in order]
    wb.active = 0
    out[MASTER_FILENAME] = _wb_bytes(wb)

    # 2. Standalone Exports (zero broken links!)
    # Materials Takeoff Only
    wb_mat = Workbook()
    dw = wb_mat.active
    build_material_takeoff_sheet(wb_mat, ds["metrics"], ds["builtup_sqft"], steel_rate)
    wb_mat.remove(dw)
    out["BOQ_Materials_Takeoff_Only.xlsx"] = _wb_bytes(wb_mat)

    # Client Quotation BOQ Only
    wb_cli = Workbook()
    dw = wb_cli.active
    build_client_boq_sheet(wb_cli, client_rows, cont_pct, ohp_pct)
    wb_cli.remove(dw)
    out["BOQ_Client_Quotation.xlsx"] = _wb_bytes(wb_cli)

    # Bank Loan Estimate Only
    prev = get_modular_boq_preview(raw, p)
    wb_bnk = Workbook()
    dw = wb_bnk.active
    build_bank_loan_estimate_sheet(wb_bnk, prev["kpis"]["grand_total"], ds["builtup_area_sqm"])
    wb_bnk.remove(dw)
    out["BOQ_Bank_Loan_Estimate.xlsx"] = _wb_bytes(wb_bnk)

    # Standalone Trade Workbooks
    standalone_specs = [
        ("BOQ_Foundation_Earthwork.xlsx", "4. Foundation & Basement", "MODULE 1: FOUNDATION & BASEMENT BOQ", ds["fdn_rows"]),
        ("BOQ_RCC_Structure.xlsx", "5. RCC Structure & Steel", "MODULE 2: RCC STRUCTURE & STEEL BOQ", ds["rcc_rows"]),
        ("BOQ_Flooring_Paint.xlsx", "8. Flooring, Tiles & Paint", "MODULE 5: FLOORING, TILES, GRANITE & PAINT BOQ", ds["flr_rows"]),
        ("BOQ_Electrical_Plumbing.xlsx", "9. Electrical & Plumbing", "MODULE 6: ELECTRICAL & PLUMBING BOQ", ds["mep_rows"]),
        ("BOQ_Additional_Works.xlsx", "10. Additional Works", "MODULE 7: ADDITIONAL EXTERNAL WORKS BOQ", ds["add_rows"]),
    ]
    for fname, sname, title, rdata in standalone_specs:
        wbs = Workbook()
        dw = wbs.active
        _build_trade_sheet(wbs, sname, title, "Standalone Trade Export — AutoBOQ 2D Modular System.", rdata)
        wbs.remove(dw)
        out[fname] = _wb_bytes(wbs)

    wb_dw = Workbook()
    dw = wb_dw.active
    build_doors_windows_sheet(wb_dw)
    wb_dw.remove(dw)
    out["BOQ_Doors_Windows.xlsx"] = _wb_bytes(wb_dw)

    # Masonry & Plaster standalone includes Tab 7 so opening deduction never #REF!s
    wb_mp = Workbook()
    dw = wb_mp.active
    _build_trade_sheet(wb_mp, "6. Masonry & Plastering", "MODULE 3: MASONRY, PLASTERING & WEATHERING COURSE BOQ", "Includes Tab 7 for live opening deduction.", ds["mas_rows"])
    build_doors_windows_sheet(wb_mp)
    wb_mp.remove(dw)
    wb_mp.active = 0
    out["BOQ_Masonry_Plaster.xlsx"] = _wb_bytes(wb_mp)

    return out


def generate_modular_boq_files() -> list[str]:
    """Generate and save Master 12-Tab Modular BOQ (.xlsx + .pdf) and all standalone part files."""
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
