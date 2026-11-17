"""
Standalone Excel Bar Bending Schedule (BBS) Generator for:
  1. Summary & Tender BOQ (kg & MT)
  2. Footing BBS (Lx & Ly Dual-Way Mesh)
  3. Column BBS (Vertical Main Bars & Confining Lateral Ties)
  4. Beam BBS (Top, Bottom, Extra Support, Cranked & Stirrups)
  5. Slab BBS (Main, Distribution, Cranked & Top Extra Bars)
  6. Shape Codes & Bar Details (Embedded Figures for Cranked Bar, Closed Stirrup 135 deg Hook, Extra Support Bars)

Features:
  - Dynamic Excel formula in 'Shape Figure' column: changing Shape Code OR Bend/Crank Angle (°)
    automatically updates the Shape Figure AND recalculates bend deductions / crank lengths.
  - Supports Bend / Crank Angles: 0°, 45°, 90°, 135°, 180° per IS 2502 / SP 34.
"""

from __future__ import annotations

import os
import shutil
import subprocess
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

from bbs_diagrams import ensure_bbs_diagrams

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXPORTS_DIR = os.path.join(BASE_DIR, "exports")
OUTPUT_FILENAME = "Bar_Bending_Schedule_Footing_Column_Beam_Slab.xlsx"

DIAS = [8, 10, 12, 16, 20, 25, 32]
ALL_SHAPES = "SC-00,SC-01,SC-02,SC-41,SC-51"
ALL_ANGLES = "0,45,90,135,180"
BLANK_EXTRA_ROWS = 8
LOOKUP_SHEET = "6. Shape Codes & Bar Details"
LOOKUP_END_ROW = 22

NAVY = "0F172A"
SLATE_HDR = "1E293B"
BLUE_SUB = "DBEAFE"
YELLOW_INPUT = "FEF9C3"
CALC_BG = "F8FAFC"
TOTAL_BG = "E2E8F0"
GREEN_TOTAL = "DCFCE7"
NOTE_BG = "EFF6FF"

THIN_SIDE = Side(style="thin", color="CBD5E1")
MED_NAVY = Side(style="medium", color="1E293B")
DOUBLE_SIDE = Side(style="double", color="0F172A")

BORDER_GRID = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
BORDER_HDR = Border(left=THIN_SIDE, right=THIN_SIDE, top=MED_NAVY, bottom=MED_NAVY)
BORDER_TOTAL = Border(left=THIN_SIDE, right=THIN_SIDE, top=MED_NAVY, bottom=DOUBLE_SIDE)

FONT_TITLE = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
FONT_NOTE = Font(name="Calibri", size=10, italic=True, color="1E3A8A")
FONT_HDR = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
FONT_INPUT = Font(name="Calibri", size=10, bold=True, color="0F172A")
FONT_CALC = Font(name="Calibri", size=10, color="1E293B")
FONT_CALC_BOLD = Font(name="Calibri", size=10, bold=True, color="0F172A")
FONT_FIG = Font(name="Consolas", size=11, bold=True, color="B91C1C")
FONT_TOTAL = Font(name="Calibri", size=11, bold=True, color="0F172A")
FONT_SECTION = Font(name="Calibri", size=11, bold=True, color="0F172A")

FILL_TITLE = PatternFill("solid", fgColor=NAVY)
FILL_HDR = PatternFill("solid", fgColor=SLATE_HDR)
FILL_INPUT = PatternFill("solid", fgColor=YELLOW_INPUT)
FILL_CALC = PatternFill("solid", fgColor=CALC_BG)
FILL_TOTAL = PatternFill("solid", fgColor=TOTAL_BG)
FILL_GREEN = PatternFill("solid", fgColor=GREEN_TOTAL)
FILL_NOTE = PatternFill("solid", fgColor=NOTE_BG)
FILL_SUB = PatternFill("solid", fgColor=BLUE_SUB)
FILL_WHITE = PatternFill("solid", fgColor="FFFFFF")

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)


def _setup_sheet_banner(ws, max_col: int, title: str, subtitle: str) -> None:
    ws.views.sheetView[0].showGridLines = True
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    c1 = ws.cell(1, 1, title)
    c1.font = FONT_TITLE
    c1.fill = FILL_TITLE
    c1.alignment = ALIGN_LEFT
    ws.row_dimensions[1].height = 28

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
    c2 = ws.cell(2, 1, subtitle)
    c2.font = FONT_NOTE
    c2.fill = FILL_NOTE
    c2.alignment = ALIGN_LEFT
    ws.row_dimensions[2].height = 20


def _write_headers(ws, row: int, headers: list[tuple[str, float]], freeze_col: str = "F") -> None:
    ws.row_dimensions[row].height = 34
    for col_idx, (text, width) in enumerate(headers, start=1):
        c = ws.cell(row, col_idx, text)
        c.font = FONT_HDR
        c.fill = FILL_HDR
        c.alignment = ALIGN_CENTER
        c.border = BORDER_HDR
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.freeze_panes = f"{freeze_col}{row + 1}"


def _add_list_dv(ws, formula_list: str, col_letter: str, start_r: int, end_r: int) -> None:
    dv = DataValidation(type="list", formula1=f'"{formula_list}"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{col_letter}{start_r}:{col_letter}{end_r}")


def _shape_figure_formula(r: int, sc_col: str = "C", ang_col: str = "E") -> str:
    """Live Excel formula that updates Shape Figure automatically when Shape Code or Angle changes."""
    return (
        f'=IF({sc_col}{r}="SC-00","━━━━━━━━━━",'
        f'IF({sc_col}{r}="SC-01",IF({ang_col}{r}=135,"⟋━━━━━━━━",IF({ang_col}{r}=45,"╲━━━━━━━━","┗━━━━━━━━")),'
        f'IF({sc_col}{r}="SC-02",IF({ang_col}{r}=135,"⟋━━━━━━━⟍","┗━━━━━━━━┛"),'
        f'IF({sc_col}{r}="SC-41",IF({ang_col}{r}=90,"━┃____┃━ (90°)",IF({ang_col}{r}=135,"━╱____╲━ (135°)","━╲____╱━ (45°)")),'
        f'IF({sc_col}{r}="SC-51","▭↺ ("&{ang_col}{r}&"°)","━━━━━━━━━━")))))'
    )


def _angle_ded_expr(ang_cell: str) -> str:
    return f"IF({ang_cell}=45,1,IF({ang_cell}=90,2,IF({ang_cell}=135,3,IF({ang_cell}=180,4,0))))"


def _crank_factor_expr(ang_cell: str) -> str:
    return f"IF({ang_cell}=45,0.42,IF({ang_cell}=90,1,IF({ang_cell}=135,1.3,IF({ang_cell}=30,0.27,IF({ang_cell}=60,0.58,0)))))"


def _embed_thumb(ws, img_path: str, cell_ref: str) -> None:
    if not img_path or not os.path.exists(img_path):
        return
    xl_img = XLImage(img_path)
    xl_img.width = 122
    xl_img.height = 36
    ws.add_image(xl_img, cell_ref)


def _get_thumb(diag_paths: dict[str, str], shape_code: str, angle: int) -> str | None:
    return diag_paths.get(f"{shape_code}_{angle}") or diag_paths.get(shape_code)


def _add_dynamic_shape_pic(
    wb: Workbook,
    ws,
    r: int,
    prefix: str,
    diag_paths: dict[str, str],
    default_sc: str,
    default_ang: int,
) -> None:
    """Places a thumbnail in D{r} and registers a Linked Picture DefinedName for auto-switching."""
    _embed_thumb(ws, _get_thumb(diag_paths, default_sc, default_ang), f"D{r}")
    sname = ws.title
    formula = (
        f"INDEX('{LOOKUP_SHEET}'!$K$5:$K${LOOKUP_END_ROW},"
        f"IFERROR(MATCH('{sname}'!$C${r}&\"_\"&'{sname}'!$E${r},'{LOOKUP_SHEET}'!$J$5:$J${LOOKUP_END_ROW},0),"
        f"MATCH('{sname}'!$C${r},'{LOOKUP_SHEET}'!$J$5:$J${LOOKUP_END_ROW},0)))"
    )
    wb.defined_names.add(DefinedName(f"{prefix}{r}", attr_text=formula))


def _style_row(
    ws,
    r: int,
    max_col: int,
    input_cols: set[int],
    center_cols: set[int],
    left_cols: set[int],
    fmt_map: dict[int, str],
    fig_col: int | None = 4,
) -> None:
    ws.row_dimensions[r].height = 33
    for ci in range(1, max_col + 1):
        cell = ws.cell(r, ci)
        cell.border = BORDER_GRID
        if fig_col and ci == fig_col:
            cell.fill = FILL_WHITE
            cell.alignment = ALIGN_CENTER
        elif ci in input_cols:
            cell.fill = FILL_INPUT
            cell.font = FONT_INPUT
        else:
            cell.fill = FILL_CALC
            cell.font = FONT_CALC
        if ci in center_cols:
            cell.alignment = ALIGN_CENTER
        elif ci in left_cols:
            cell.alignment = ALIGN_LEFT
        else:
            cell.alignment = ALIGN_RIGHT
        if ci in fmt_map:
            cell.number_format = fmt_map[ci]


# =============================================================================
# SHEET 2: FOOTING BBS (Lx & Ly)
# =============================================================================
def build_footing_sheet(wb: Workbook, diag_paths: dict[str, str], params: dict | None = None) -> dict:
    ws = wb.create_sheet("2. Footing BBS (Lx & Ly)")
    max_col = 31  # A..AE
    _setup_sheet_banner(
        ws,
        max_col,
        "FOOTING BAR BENDING SCHEDULE — DYNAMIC SHAPE FIGURE, ANGLE (°), BAR TYPE & Lx/Ly MESH (IS 2502)",
        "YELLOW CELLS = Change Shape Code or Angle (0°/45°/90°/135°/180°) and Shape Figure & Bend Deduction update automatically!",
    )

    headers = [
        ("S.No", 6),              # 1 A
        ("Footing Mark", 11),     # 2 B
        ("Shape Code", 10),       # 3 C
        ("Shape Figure", 19),     # 4 D
        ("Angle (°)", 9),         # 5 E
        ("Bar Type", 12),         # 6 F
        ("Bar Description", 28),  # 7 G
        ("No. of Footings", 10),  # 8 H
        ("Length Lx (mm)", 11),   # 9 I
        ("Breadth Ly (mm)", 11),  # 10 J
        ("Depth D (mm)", 10),     # 11 K
        ("Cover c (mm)", 10),     # 12 L
        ("Bar Span (Lx/Ly)", 10), # 13 M
        ("Dia d (mm)", 9),        # 14 N
        ("Spacing s (mm)", 10),   # 15 O
        ("Clear Run (mm)", 11),   # 16 P
        ("L-Bend Leg (mm)", 11),  # 17 Q
        ("Bend Ded. (mm)", 10),   # 18 R
        ("Cutting Length (m)", 12), # 19 S
        ("Bars / Footing", 10),   # 20 T
        ("Total Bars", 10),       # 21 U
        ("Total Length (m)", 12), # 22 V
        ("Unit Wt (kg/m)", 11),   # 23 W
        ("8 mm (kg)", 11),        # 24 X
        ("10 mm (kg)", 11),       # 25 Y
        ("12 mm (kg)", 11),       # 26 Z
        ("16 mm (kg)", 11),       # 27 AA
        ("20 mm (kg)", 11),       # 28 AB
        ("25 mm (kg)", 11),       # 29 AC
        ("32 mm (kg)", 11),       # 30 AD
        ("Total Wt (kg)", 13),    # 31 AE
    ]
    _write_headers(ws, 4, headers, freeze_col="H")

    sample_rows = [
        ("F1", "SC-02", 90, "BOT-MESH-LX", "Longitudinal Bottom Mesh (along Lx) Y10@125", 6, 1800, 1800, 525, 50, "Lx", 10, 125),
        ("F1", "SC-02", 90, "BOT-MESH-LY", "Transverse Bottom Mesh (along Ly) Y10@125",   6, 1800, 1800, 525, 50, "Ly", 10, 125),
        ("F2", "SC-02", 90, "BOT-MESH-LX", "Longitudinal Bottom Mesh (along Lx) Y10@100", 4, 2100, 2100, 600, 50, "Lx", 10, 100),
        ("F2", "SC-02", 90, "BOT-MESH-LY", "Transverse Bottom Mesh (along Ly) Y10@100",   4, 2100, 2100, 600, 50, "Ly", 10, 100),
        ("F3", "SC-02", 90, "BOT-MESH-LX", "Longitudinal Bottom Mesh (along Lx) Y12@125", 2, 2400, 2400, 680, 50, "Lx", 12, 125),
        ("F3", "SC-02", 90, "BOT-MESH-LY", "Transverse Bottom Mesh (along Ly) Y12@125",   2, 2400, 2400, 680, 50, "Ly", 12, 125),
        ("F3", "SC-02", 90, "TOP-MAT-LX",  "Top Mat Mesh (along Lx) Y10@150",             2, 2400, 2400, 680, 50, "Lx", 10, 150),
        ("F3", "SC-02", 90, "TOP-MAT-LY",  "Top Mat Mesh (along Ly) Y10@150",             2, 2400, 2400, 680, 50, "Ly", 10, 150),
        ("F4", "SC-02", 90, "BOT-MESH-LX", "Longitudinal Bottom Mesh (along Lx) Y12@100", 2, 2700, 2700, 750, 50, "Lx", 12, 100),
        ("F4", "SC-02", 90, "BOT-MESH-LY", "Transverse Bottom Mesh (along Ly) Y12@100",   2, 2700, 2700, 750, 50, "Ly", 12, 100),
        ("F4", "SC-02", 90, "TOP-MAT-LX",  "Top Mat Mesh (along Lx) Y10@100",             2, 2700, 2700, 750, 50, "Lx", 10, 100),
        ("F4", "SC-02", 90, "TOP-MAT-LY",  "Top Mat Mesh (along Ly) Y10@100",             2, 2700, 2700, 750, 50, "Ly", 10, 100),
        ("F5", "SC-02", 90, "BOT-MESH-LX", "Longitudinal Bottom Mesh (along Lx) Y12@100", 2, 3000, 3000, 830, 50, "Lx", 12, 100),
        ("F5", "SC-02", 90, "BOT-MESH-LY", "Transverse Bottom Mesh (along Ly) Y12@100",   2, 3000, 3000, 830, 50, "Ly", 12, 100),
        ("F5", "SC-02", 90, "TOP-MAT-LX",  "Top Mat Mesh (along Lx) Y12@200",             2, 3000, 3000, 830, 50, "Lx", 12, 200),
        ("F5", "SC-02", 90, "TOP-MAT-LY",  "Top Mat Mesh (along Ly) Y12@200",             2, 3000, 3000, 830, 50, "Ly", 12, 200),
        ("F6", "SC-02", 90, "BOT-MESH-LX", "Longitudinal Bottom Mesh (along Lx) Y12@100", 2, 3200, 3200, 830, 50, "Lx", 12, 100),
        ("F6", "SC-02", 90, "BOT-MESH-LY", "Transverse Bottom Mesh (along Ly) Y12@100",   2, 3200, 3200, 830, 50, "Ly", 12, 100),
        ("F6", "SC-02", 90, "TOP-MAT-LX",  "Top Mat Mesh (along Lx) Y12@200",             2, 3200, 3200, 830, 50, "Lx", 12, 200),
        ("F6", "SC-02", 90, "TOP-MAT-LY",  "Top Mat Mesh (along Ly) Y12@200",             2, 3200, 3200, 830, 50, "Ly", 12, 200),
        ("SF", "SC-02", 90, "BOT-MESH-LX", "Stub Footing Bottom Mesh (along Lx) Y10@150", 2, 1500, 1500, 380, 50, "Lx", 10, 150),
        ("SF", "SC-02", 90, "BOT-MESH-LY", "Stub Footing Bottom Mesh (along Ly) Y10@150", 2, 1500, 1500, 380, 50, "Ly", 10, 150),
        ("CF1","SC-02", 90, "BOT-MESH-LX", "Combined Footing Bottom Mesh (Lx) Y12@100",   1, 3600, 3600, 750, 50, "Lx", 12, 100),
        ("CF1","SC-02", 90, "BOT-MESH-LY", "Combined Footing Bottom Mesh (Ly) Y12@100",   1, 3600, 3600, 750, 50, "Ly", 12, 100),
        ("CF1","SC-02", 90, "TOP-MAT-LX",  "Combined Footing Top Mat (Lx) Y12@150",       1, 3600, 3600, 750, 50, "Lx", 12, 150),
        ("CF1","SC-02", 90, "TOP-MAT-LY",  "Combined Footing Top Mat (Ly) Y12@150",       1, 3600, 3600, 750, 50, "Ly", 12, 150),
    ]
    if params:
        f_cov = int(params.get("footing_cover_mm", 50))
        f_ov = params.get("footings") or {}
        updated = []
        seen_marks = set()
        for row in sample_rows:
            mk = row[0]
            seen_marks.add(mk)
            if mk in f_ov:
                ov = f_ov[mk]
                updated.append((
                    mk, row[1], row[2], row[3], row[4],
                    int(ov.get("nos", row[5])),
                    int(ov.get("lx", row[6])),
                    int(ov.get("ly", row[7])),
                    int(ov.get("df", row[8])),
                    f_cov, row[10], row[11], row[12],
                ))
            else:
                updated.append((*row[:9], f_cov, *row[10:]))
        for mk, ov in f_ov.items():
            if mk not in seen_marks and int(ov.get("nos", 0)) > 0:
                nos_v = int(ov["nos"])
                lx_v = int(ov.get("lx", 1800))
                ly_v = int(ov.get("ly", 1500))
                df_v = int(ov.get("df", 450))
                updated.append((mk, "SC-02", 90, "BOT-MESH-LX", "Longitudinal Bottom Mesh (along Lx)", nos_v, lx_v, ly_v, df_v, f_cov, "Lx", 12, 150))
                updated.append((mk, "SC-02", 90, "BOT-MESH-LY", "Transverse Bottom Mesh (along Ly)", nos_v, lx_v, ly_v, df_v, f_cov, "Ly", 12, 150))
        sample_rows = updated

    start_r = 5
    total_rows = len(sample_rows) + BLANK_EXTRA_ROWS
    input_cols = {2, 3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}
    center_cols = {1, 2, 3, 4, 5, 6, 8, 12, 13, 14, 15, 20, 21}
    left_cols = {7}
    fmt_map = {
        5: '0"°"', 8: "#,##0", 9: "#,##0", 10: "#,##0", 11: "#,##0", 12: "#,##0", 14: "0", 15: "#,##0",
        16: "#,##0", 17: "#,##0", 18: "#,##0", 19: "0.000", 20: "#,##0", 21: "#,##0",
        22: "#,##0.00", 23: "0.000",
    }
    for ci in range(24, 32):
        fmt_map[ci] = "#,##0.00"

    for idx in range(total_rows):
        r = start_r + idx
        ws.cell(r, 1, idx + 1)
        if idx < len(sample_rows):
            rv = sample_rows[idx]
            ws.cell(r, 2, rv[0])
            ws.cell(r, 3, rv[1])
            _add_dynamic_shape_pic(wb, ws, r, "Pic_F_", diag_paths, rv[1], rv[2])
            ws.cell(r, 5, rv[2])
            for off, val in enumerate(rv[3:], start=6):
                ws.cell(r, off, val)
        else:
            ws.cell(r, 3, "SC-02")
            _add_dynamic_shape_pic(wb, ws, r, "Pic_F_", diag_paths, "SC-02", 90)
            ws.cell(r, 5, 90)
            ws.cell(r, 6, "BOT-MESH-LX")
            ws.cell(r, 12, 50)
            ws.cell(r, 13, "Lx")

        ws.cell(r, 16, f'=IF(H{r}>0, IF(M{r}="Lx", I{r}-2*L{r}, J{r}-2*L{r}), 0)')
        ws.cell(r, 17, f"=IF(H{r}>0, MAX(0, K{r}-2*L{r}), 0)")
        ws.cell(r, 18, f"=IF(H{r}>0, 2*{_angle_ded_expr(f'E{r}')}*N{r}, 0)")
        ws.cell(r, 19, f"=IF(H{r}>0, ROUND((P{r}+2*Q{r}-R{r})/1000, 3), 0)")
        ws.cell(r, 20, f'=IF(AND(H{r}>0, O{r}>0), IF(M{r}="Lx", ROUNDDOWN((J{r}-2*L{r})/O{r},0)+1, ROUNDDOWN((I{r}-2*L{r})/O{r},0)+1), 0)')
        ws.cell(r, 21, f"=H{r}*T{r}")
        ws.cell(r, 22, f"=ROUND(S{r}*U{r}, 2)")
        ws.cell(r, 23, f"=IF(N{r}>0, ROUND((N{r}*N{r})/162.2, 3), 0)")
        for d_idx, dia in enumerate(DIAS):
            ws.cell(r, 24 + d_idx, f"=IF(N{r}={dia}, ROUND(V{r}*W{r}, 2), 0)")
        ws.cell(r, 31, f"=SUM(X{r}:AD{r})")

        _style_row(ws, r, max_col, input_cols, center_cols, left_cols, fmt_map, fig_col=4)
        ws.cell(r, 19).font = FONT_CALC_BOLD
        ws.cell(r, 31).font = FONT_CALC_BOLD

    end_r = start_r + total_rows - 1
    tot_r = end_r + 1
    ws.row_dimensions[tot_r].height = 24
    ws.merge_cells(start_row=tot_r, start_column=1, end_row=tot_r, end_column=20)
    lbl = ws.cell(tot_r, 1, "TOTAL FOOTING REINFORCEMENT STEEL (NET WEIGHT)")
    lbl.font = FONT_TOTAL
    lbl.alignment = ALIGN_RIGHT

    for ci in (21, 22, 24, 25, 26, 27, 28, 29, 30, 31):
        col_let = get_column_letter(ci)
        c = ws.cell(tot_r, ci, f"=SUM({col_let}{start_r}:{col_let}{end_r})")
        c.font = FONT_TOTAL
        c.alignment = ALIGN_RIGHT if ci != 21 else ALIGN_CENTER
        c.number_format = "#,##0" if ci == 21 else "#,##0.00"
    for ci in range(1, max_col + 1):
        ws.cell(tot_r, ci).fill = FILL_GREEN if ci == max_col else FILL_TOTAL
        ws.cell(tot_r, ci).border = BORDER_TOTAL

    ws.auto_filter.ref = f"A4:{get_column_letter(max_col)}{end_r}"
    _add_list_dv(ws, ALL_SHAPES, "C", start_r, end_r)
    _add_list_dv(ws, ALL_ANGLES, "E", start_r, end_r)
    _add_list_dv(ws, "BOT-MESH-LX,BOT-MESH-LY,TOP-MAT-LX,TOP-MAT-LY", "F", start_r, end_r)
    _add_list_dv(ws, "Lx,Ly", "M", start_r, end_r)
    _add_list_dv(ws, "8,10,12,16,20,25,32", "N", start_r, end_r)

    dia_totals = [0.0] * len(DIAS)
    footing_lookup: dict[str, tuple[int, int]] = {}
    footing_conc_m3 = 0.0
    for rv in sample_rows:
        mk, _, ang, _, _, nos_f, lx, ly, df, cov, sdir, dia, spc = rv
        if mk not in footing_lookup:
            footing_lookup[mk] = (int(nos_f), int(df))
            footing_conc_m3 += nos_f * (lx / 1000.0) * (ly / 1000.0) * (df / 1000.0)
        if nos_f > 0 and spc > 0:
            span = (lx - 2 * cov) if sdir == "Lx" else (ly - 2 * cov)
            leg = max(0, df - 2 * cov)
            ded_k = 2 if ang == 90 else (1 if ang == 45 else (3 if ang == 135 else (4 if ang == 180 else 0)))
            ded = 2 * ded_k * dia
            cut_m = round((span + 2 * leg - ded) / 1000.0, 3)
            bars_per = int(((ly - 2 * cov) if sdir == "Lx" else (lx - 2 * cov)) // spc) + 1
            tot_len = round(cut_m * nos_f * bars_per, 2)
            unit_wt = round((dia * dia) / 162.2, 3)
            wt = round(tot_len * unit_wt, 2)
            if dia in DIAS:
                dia_totals[DIAS.index(dia)] = round(dia_totals[DIAS.index(dia)] + wt, 2)

    return {
        "sheet_name": ws.title,
        "total_row": tot_r,
        "dia_cols": [get_column_letter(24 + i) for i in range(len(DIAS))],
        "total_col": get_column_letter(31),
        "default_concrete_m3": round(footing_conc_m3, 2),
        "dia_totals": dia_totals,
        "footing_lookup": footing_lookup,
    }


def generate_column_shorthand_rows(
    shorthand: str,
    b_mm: int = 300,
    d_mm: int = 450,
    main_bars_ffr: list[tuple[int, int]] | None = None,
    main_bars_top: list[tuple[int, int]] | None = None,
) -> list[tuple]:
    """Expand a shorthand like 'C1: 19 nos on F1, 3 nos on F2' into full Zone A + Zone B openpyxl rows."""
    import re

    col_mark = shorthand.split(":", 1)[0].strip() if ":" in shorthand else "C1"
    groups = re.findall(r"(\d+)\s*(?:nos?)?\s*(?:on|@)\s*([A-Za-z0-9_]+)", shorthand, flags=re.IGNORECASE)
    if not groups:
        groups = [("1", "F1")]
    ffr = main_bars_ffr or [(6, 20), (4, 16)]
    top = main_bars_top or [(10, 16)]

    rows: list[tuple] = []
    for nos_str, fref in groups:
        nos = int(nos_str)
        fref_u = fref.upper()
        base = f"{col_mark} on {fref_u}"
        for cnt, dia in ffr:
            rows.append((f"{base} [Zone A Fdn]", "SC-01", 90, "MAIN", f"Zone A: Starter ({cnt}-Y{dia}) Fdn Pad + Neck + 40d Lap", fref_u, nos, b_mm, d_mm, "=$C$3", "=$R$3", "=$I$3", dia, cnt))
        for cnt, dia in ffr:
            rows.append((f"{base} [Zone B FFR]", "SC-00", 0, "MAIN", f"Zone B: Superstructure ({cnt}-Y{dia}) — Upto FFR", fref_u, nos, b_mm, d_mm, "=$C$3", "=$L$3", 0, dia, cnt))
        for cnt, dia in top:
            rows.append((f"{base} [Zone B Top]", "SC-00", 0, "MAIN", f"Zone B: Superstructure ({cnt}-Y{dia}) — FFR to Top", fref_u, nos, b_mm, d_mm, "=$C$3", "=$L$3", 0, dia, cnt))
        rows.extend(_column_tie_block(base, fref_u, nos, b_mm, d_mm))
    return rows


def _column_tie_block(base: str, fref: str, nos: int, b_mm: int, d_mm: int) -> list[tuple]:
    """IS 13920 tie block: Zone A ring over Df (=DF -> Col N), Zone A neck, Zone B HC/4 + HC/2 zones."""
    return [
        (f"{base} [Zone A Ring]",  "SC-51", 135, "TIE", "Zone A: Footing Depth Outer Rings (Y10 @ 100 c/c over Df)", fref, nos, b_mm, d_mm, "=$C$3", "=DF",     0, 10, 100),
        (f"{base} [Zone A Neck]",  "SC-51", 135, "TIE", "Zone A: Foundation Neck Ties (Y8 @ 150 c/c)",              fref, nos, b_mm, d_mm, "=$C$3", "=$R$3",   0, 8,  150),
        (f"{base} [Zone B Bot]",   "SC-51", 135, "TIE", "Zone B: Confining Ties — Bot HC/4 (Y8 @ 100 c/c)",         fref, nos, b_mm, d_mm, "=$C$3", "=$L$3/4", 0, 8,  100),
        (f"{base} [Zone B Top]",   "SC-51", 135, "TIE", "Zone B: Confining Ties — Top HC/4 (Y8 @ 100 c/c)",         fref, nos, b_mm, d_mm, "=$C$3", "=$L$3/4", 0, 8,  100),
        (f"{base} [Zone B Mid]",   "SC-51", 135, "TIE", "Zone B: Lap Splice Ties — Mid HC/2 (Y8 @ 150 c/c)",        fref, nos, b_mm, d_mm, "=$C$3", "=$L$3/2", 0, 8,  150),
    ]


def _floor_multiplier_expr(mark: str, btype: str, marks_with_ffr: set[str]) -> str:
    """Default Floor Multiplier (Col AG). Zone A/C = 1. Zone B ties & single-tier mains = Typical Floors ($X$3).
    Where a column has a separate 'Upto FFR' tier: FFR tier = 1 floor, 'FFR to Top' tier = $X$3 - 1."""
    if "Zone B" not in mark:
        return "1"
    base = mark.split("[", 1)[0].strip()
    if btype == "MAIN" and "[Zone B FFR]" in mark:
        return "1"
    if btype == "MAIN" and "[Zone B Top]" in mark and base in marks_with_ffr:
        return "=MAX($X$3-1,0)"
    return "=$X$3"


# =============================================================================
# SHEET 3: COLUMN BBS (MAIN VERTICAL BARS & LATERAL TIES)
# =============================================================================
def build_column_sheet(
    wb: Workbook,
    diag_paths: dict[str, str],
    params: dict | None = None,
    footing_lookup: dict[str, tuple[int, int]] | None = None,
) -> dict:
    ws = wb.create_sheet("3. Column BBS (Main & Ties)")
    max_col = 33  # A..AG (AG = Floor Multiplier)
    _setup_sheet_banner(
        ws,
        max_col,
        "COLUMN BAR BENDING SCHEDULE — DYNAMIC VLOOKUP FOOTING DEPTH (F1/F2/F3) & IS 13920 VERTICAL STACK",
        "Auto VLOOKUP pulls Footing Pad Depth Df from Sheet 2 (F1/F2/F3/CF1) | Zone A = Base L-Bend + VLOOKUP(Df) + Neck + 40d Lap | Zone B = HC/4 & HC/2 Ties.",
    )

    # Row 3: Drawing Defaults / Constants (Yellow = Editable once for whole building)
    ws.row_dimensions[3].height = 22
    p = params or {}
    c_cov = int(p.get("col_cover_mm", 40))
    c_ld = int(p.get("lap_factor_d", 40))
    c_bend = int(p.get("base_bend_mm", 400))
    c_hc = int(p.get("clear_height_mm", 3660))
    c_lap_f = float(p.get("lap_multiplier", 1.0))
    c_neck = int(p.get("neck_height_mm", 1800))
    c_para = int(p.get("parapet_height_mm", 1000))
    c_floors = int(p.get("typical_floors", 3))
    defaults = [
        (1, "Cover c (mm):", 3, c_cov),
        (4, "Col Ld Factor (×d):", 6, c_ld),
        (7, "Footing L-Bend (mm):", 9, c_bend),
        (10, "Floor Height HC (mm):", 12, c_hc),
        (13, "Lap Factor (1=40d):", 15, c_lap_f),
        (16, "Zone A Neck Ht (mm):", 18, c_neck),
        (19, "Zone C Parapet Ht (mm):", 21, c_para),
        (22, "Typical Floors (×):", 24, c_floors),
    ]
    for lc, lt, vc, val in defaults:
        ws.merge_cells(start_row=3, start_column=lc, end_row=3, end_column=lc + 1)
        lcell = ws.cell(3, lc, lt)
        lcell.font = FONT_CALC_BOLD
        lcell.fill = FILL_CALC
        lcell.alignment = ALIGN_RIGHT
        lcell.border = BORDER_GRID
        ws.cell(3, lc + 1).border = BORDER_GRID
        vcell = ws.cell(3, vc, val)
        vcell.font = FONT_INPUT
        vcell.fill = FILL_INPUT
        vcell.alignment = ALIGN_CENTER
        vcell.border = BORDER_GRID

    # Row 4: Plain-English Beginner Guide
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=max_col)
    g4 = ws.cell(
        4,
        1,
        "DYNAMIC FOOTING LOOKUP + VERTICAL STACK:  "
        "1) Pick Footing Ref (Col H: F1/F2/F3/CF1) — Col N automatically VLOOKUPs Pad Depth Df from Sheet 2!  "
        "2) Enter No. of Cols (Col I).  "
        "3) Zone A Cut = Base L-Bend + VLOOKUP(Df) + Neck + 40d Lap - Bend Ded.  "
        "4) Zone B Ties = Bot HC/4 (@100), Mid HC/2 (@150), Top HC/4 (@100).  "
        "5) Col AG Floor Multiplier scales identical floors (Typical Floors in X3) — Total Bars U = I × T × AG.",
    )
    g4.font = FONT_NOTE
    g4.fill = FILL_SUB
    g4.alignment = ALIGN_LEFT
    ws.row_dimensions[4].height = 22

    headers = [
        ("S.No", 5),                       # 1 A
        ("Col Mark & Zone", 16),           # 2 B
        ("Shape Code", 10),                # 3 C
        ("Shape Figure", 19),              # 4 D
        ("Angle (°)", 9),                  # 5 E
        ("Bar Type", 10),                  # 6 F
        ("Plain-English Bar Description", 34), # 7 G
        ("Footing Ref", 11),               # 8 H
        ("No. of Cols", 10),               # 9 I
        ("Width b (mm)", 10),              # 10 J
        ("Depth D (mm)", 10),              # 11 K
        ("Cover c (mm)", 9),               # 12 L
        ("Neck / Clear H (mm)", 12),       # 13 M
        ("Auto Pad Df (VLOOKUP)", 12),     # 14 N
        ("Base L-Bend (mm)", 11),          # 15 O
        ("Dia d (mm)", 9),                 # 16 P
        ("Bars/Col or Spacing", 12),       # 17 Q
        ("Auto Ld=40d (mm)", 11),          # 18 R
        ("135° Hook MAX(10d,75)", 12),     # 19 S
        ("Calc Bars / Col", 11),           # 20 T
        ("Total Bars", 10),                # 21 U
        ("Cutting Length (m)", 12),        # 22 V
        ("Total Length (m)", 12),          # 23 W
        ("Unit Wt (kg/m)", 11),            # 24 X
        ("8 mm (kg)", 11),                 # 25 Y
        ("10 mm (kg)", 11),                # 26 Z
        ("12 mm (kg)", 11),                # 27 AA
        ("16 mm (kg)", 11),                # 28 AB
        ("20 mm (kg)", 11),                # 29 AC
        ("25 mm (kg)", 11),                # 30 AD
        ("32 mm (kg)", 11),                # 31 AE
        ("Total Wt (kg)", 13),             # 32 AF
        ("Floor Mult (×)", 11),            # 33 AG
    ]
    _write_headers(ws, 6, headers, freeze_col="I")

    # Exact Column Schedule from User's Structural Drawing (C1..C9 & SC)
    # Tuple: (Mark, Shape, Angle, Type, Desc, FootingRef, Nos, b, D, Cover, NeckOrClearH, BaseBend, Dia, BarsOrSpacing)
    def _std_tie_zones(mark_base: str, footing_ref: str, b_mm: int, d_mm: int):
        return _column_tie_block(mark_base, footing_ref, 1, b_mm, d_mm)

    c1_b = int(p.get("c1_b_mm", 300))
    c1_d = int(p.get("c1_d_mm", 450))

    # Tuple: (Mark, Shape, Angle, Type, Desc, FootingRef, Nos, b, D, Cover, NeckOrClearH, BaseBend, Dia, BarsOrSpacing)
    sample_rows = [
        # Multi-Footing Column Variations: C1 on F1 (6 nos, Df=450) and C1 on F2 (4 nos, Df=500)
        ("C1 on F1 [Zone A Fdn]", "SC-01", 90, "MAIN", "Zone A: Starter (6-Y20) Fdn Pad + Neck + 40d Lap", "F1",   1, c1_b, c1_d, "=$C$3", "=$R$3", "=$I$3", 20, 6),
        ("C1 on F1 [Zone A Fdn]", "SC-01", 90, "MAIN", "Zone A: Starter (4-Y16) Fdn Pad + Neck + 40d Lap", "F1",   1, c1_b, c1_d, "=$C$3", "=$R$3", "=$I$3", 16, 4),
        ("C1 on F1 [Zone B FFR]", "SC-00", 0,  "MAIN", "Zone B: Superstructure (6-Y20) — Upto FFR",         "F1",   1, c1_b, c1_d, "=$C$3", "=$L$3", 0,       20, 6),
        ("C1 on F1 [Zone B FFR]", "SC-00", 0,  "MAIN", "Zone B: Superstructure (4-Y16) — Upto FFR",         "F1",   1, c1_b, c1_d, "=$C$3", "=$L$3", 0,       16, 4),
        ("C1 on F1 [Zone B Top]", "SC-00", 0,  "MAIN", "Zone B: Superstructure (10-Y16) — FFR to Top",      "F1",   1, c1_b, c1_d, "=$C$3", "=$L$3", 0,       16, 10),
        *_std_tie_zones("C1 on F1", "F1", c1_b, c1_d),

        ("C1 on F2 [Zone A Fdn]", "SC-01", 90, "MAIN", "Zone A: Starter (6-Y20) Fdn Pad + Neck + 40d Lap", "F2",   1, c1_b, c1_d, "=$C$3", "=$R$3", "=$I$3", 20, 6),
        ("C1 on F2 [Zone A Fdn]", "SC-01", 90, "MAIN", "Zone A: Starter (4-Y16) Fdn Pad + Neck + 40d Lap", "F2",   1, c1_b, c1_d, "=$C$3", "=$R$3", "=$I$3", 16, 4),
        ("C1 on F2 [Zone B FFR]", "SC-00", 0,  "MAIN", "Zone B: Superstructure (6-Y20) — Upto FFR",         "F2",   1, c1_b, c1_d, "=$C$3", "=$L$3", 0,       20, 6),
        ("C1 on F2 [Zone B FFR]", "SC-00", 0,  "MAIN", "Zone B: Superstructure (4-Y16) — Upto FFR",         "F2",   1, c1_b, c1_d, "=$C$3", "=$L$3", 0,       16, 4),
        ("C1 on F2 [Zone B Top]", "SC-00", 0,  "MAIN", "Zone B: Superstructure (10-Y16) — FFR to Top",      "F2",   1, c1_b, c1_d, "=$C$3", "=$L$3", 0,       16, 10),
        *_std_tie_zones("C1 on F2", "F2", c1_b, c1_d),

        ("C2 [Zone A Fdn]", "SC-01", 90, "MAIN", "Zone A: Starter (6-Y25) Fdn Pad + Neck + 40d Lap", "F2",   1, 300, 450, "=$C$3", "=$R$3", "=$I$3", 25, 6),
        ("C2 [Zone A Fdn]", "SC-01", 90, "MAIN", "Zone A: Starter (4-Y20) Fdn Pad + Neck + 40d Lap", "F2",   1, 300, 450, "=$C$3", "=$R$3", "=$I$3", 20, 4),
        ("C2 [Zone B FFR]", "SC-00", 0,  "MAIN", "Zone B: Superstructure (6-Y25 + 4-Y20) — Upto FFR", "F2",   1, 300, 450, "=$C$3", "=$L$3", 0,       25, 6),
        ("C2 [Zone B Top]", "SC-00", 0,  "MAIN", "Zone B: Superstructure (10-Y20) — FFR to Top",      "F2",   1, 300, 450, "=$C$3", "=$L$3", 0,       20, 10),
        *_std_tie_zones("C2", "F2", 300, 450),

        ("C3 [Zone A Fdn]", "SC-01", 90, "MAIN", "Zone A: Starter (4-Y25) Fdn + Neck + 40d Lap",     "F2",   1, 300, 600, "=$C$3", "=$R$3", "=$I$3", 25, 4),
        ("C3 [Zone A Fdn]", "SC-01", 90, "MAIN", "Zone A: Starter (8-Y20) Fdn + Neck + 40d Lap",     "F2",   1, 300, 600, "=$C$3", "=$R$3", "=$I$3", 20, 8),
        ("C3 [Zone B Top]", "SC-00", 0,  "MAIN", "Zone B: Superstructure (12-Y20) — FFR to Top",     "F2",   1, 300, 600, "=$C$3", "=$L$3", 0,       20, 12),
        *_std_tie_zones("C3", "F2", 300, 600),

        ("C4 [Zone A Fdn]", "SC-01", 90, "MAIN", "Zone A: Starter (8-Y25) Fdn + Neck + 40d Lap",     "F2",   1, 300, 600, "=$C$3", "=$R$3", "=$I$3", 25, 8),
        ("C4 [Zone A Fdn]", "SC-01", 90, "MAIN", "Zone A: Starter (4-Y20) Fdn + Neck + 40d Lap",     "F2",   1, 300, 600, "=$C$3", "=$R$3", "=$I$3", 20, 4),
        ("C4 [Zone B Top]", "SC-00", 0,  "MAIN", "Zone B: Superstructure (4-Y25) — FFR to Top",      "F2",   1, 300, 600, "=$C$3", "=$L$3", 0,       25, 4),
        ("C4 [Zone B Top]", "SC-00", 0,  "MAIN", "Zone B: Superstructure (8-Y20) — FFR to Top",      "F2",   1, 300, 600, "=$C$3", "=$L$3", 0,       20, 8),
        *_std_tie_zones("C4", "F2", 300, 600),

        ("C5 [Zone A Fdn]", "SC-01", 90, "MAIN", "Zone A: Starter (10-Y20) Fdn + Neck + 40d Lap",    "F1",   1, 230, 450, "=$C$3", "=$R$3", "=$I$3", 20, 10),
        ("C5 [Zone B Top]", "SC-00", 0,  "MAIN", "Zone B: Superstructure (6-Y20) — FFR to Top",      "F1",   1, 230, 450, "=$C$3", "=$L$3", 0,       20, 6),
        ("C5 [Zone B Top]", "SC-00", 0,  "MAIN", "Zone B: Superstructure (4-Y16) — FFR to Top",      "F1",   1, 230, 450, "=$C$3", "=$L$3", 0,       16, 4),
        *_std_tie_zones("C5", "F1", 230, 450),

        ("C6 [Zone A Fdn]", "SC-01", 90, "MAIN", "Zone A: Starter (16-Y20) Fdn + Neck + 40d Lap",    "F3",   1, 600, 600, "=$C$3", "=$R$3", "=$I$3", 20, 16),
        ("C6 [Zone B Flr]", "SC-00", 0,  "MAIN", "Zone B: Superstructure (16-Y20) — Upto Top",       "F3",   1, 600, 600, "=$C$3", "=$L$3", 0,       20, 16),
        *_std_tie_zones("C6", "F3", 600, 600),

        ("C7 [Zone A Fdn]", "SC-01", 90, "MAIN", "Zone A: Starter (16-Y25) Fdn + Neck + 40d Lap",    "CF1",  1, 600, 600, "=$C$3", "=$R$3", "=$I$3", 25, 16),
        ("C7 [Zone B Flr]", "SC-00", 0,  "MAIN", "Zone B: Superstructure (16-Y25) — Upto Top",       "CF1",  1, 600, 600, "=$C$3", "=$L$3", 0,       25, 16),
        *_std_tie_zones("C7", "CF1", 600, 600),

        ("C8 [Zone A Fdn]", "SC-01", 90, "MAIN", "Zone A: Starter (28-Y32) Fdn + Neck + 40d Lap",    "CF1",  1, 900, 900, "=$C$3", "=$R$3", "=$I$3", 32, 28),
        ("C8 [Zone B Flr]", "SC-00", 0,  "MAIN", "Zone B: Superstructure (28-Y32) — Upto Top",       "CF1",  1, 900, 900, "=$C$3", "=$L$3", 0,       32, 28),
        *_std_tie_zones("C8", "CF1", 900, 900),

        ("C9 [Zone A Fdn]", "SC-01", 90, "MAIN", "Zone A: Starter (10-Y16) Fdn + Neck + 40d Lap",    "F1",   1, 230, 450, "=$C$3", "=$R$3", "=$I$3", 16, 10),
        ("C9 [Zone B Top]", "SC-00", 0,  "MAIN", "Zone B: Superstructure (6-Y16) — FFR to Top",      "F1",   1, 230, 450, "=$C$3", "=$L$3", 0,       16, 6),
        ("C9 [Zone B Top]", "SC-00", 0,  "MAIN", "Zone B: Superstructure (4-Y12) — FFR to Top",      "F1",   1, 230, 450, "=$C$3", "=$L$3", 0,       12, 4),
        *_std_tie_zones("C9", "F1", 230, 450),

        ("SC [Zone A Tie]", "SC-01", 90, "MAIN", "Zone A: Stub Column (8-Y12) — Upto Tie Beam",      "SF",   1, 300, 300, "=$C$3", "=$R$3", "=$I$3", 12, 8),
        ("SC [Zone A Ring]","SC-51", 135,"TIE",  "Zone A: Stub Footing Depth Outer Rings (Y10 @ 100 c/c over Df)", "SF", 1, 300, 300, "=$C$3", "=DF", 0, 10, 100),
        ("SC [Zone A Neck]","SC-51", 135,"TIE",  "Zone A: Stub Neck Ties (Y8 @ 150 c/c)",                         "SF", 1, 300, 300, "=$C$3", "=$R$3", 0, 8, 150),

        # Zone C: Terrace Termination & Parapet Level (No mid-height lap splice)
        ("Zone C [Roof]",   "SC-01", 90, "MAIN", "Zone C: Top Floor Roof Beam Anchorage (40d Bend)", "NONE", 1, 300, 450, "=$C$3", 0,       "=$I$3", 16, 10),
        ("Zone C [Parapet]","SC-01", 90, "MAIN", "Zone C: Parapet Vertical Bars (Y8 @ 150 c/c /m)", "NONE", 1, 150, 1000, 25,     "=$U$3", 300,     8,  7),
        ("Zone C [Parapet]","SC-00", 0,  "MAIN", "Zone C: Parapet Horizontal Bars (Y8 @ 150 c/c /m)","NONE", 1, 150, 1000, 25,     1000,    0,       8,  7),
    ]

    start_r = 7
    total_rows = len(sample_rows) + BLANK_EXTRA_ROWS
    marks_with_ffr = {rv[0].split("[", 1)[0].strip() for rv in sample_rows if "[Zone B FFR]" in rv[0]}
    input_cols = {2, 3, 5, 6, 7, 8, 10, 11, 12, 13, 15, 16, 17, 33}
    center_cols = {1, 2, 3, 4, 5, 6, 8, 9, 12, 16, 17, 20, 21, 33}
    left_cols = {7}
    fmt_map = {
        5: '0"°"', 9: "#,##0", 10: "#,##0", 11: "#,##0", 12: "#,##0", 13: "#,##0", 14: "#,##0", 15: "#,##0",
        16: "0", 17: "#,##0", 18: "#,##0", 19: "#,##0", 20: "#,##0", 21: "#,##0",
        22: "0.000", 23: "#,##0.00", 24: "0.000", 33: '0"×"',
    }
    for ci in range(25, 33):
        fmt_map[ci] = "#,##0.00"

    for idx in range(total_rows):
        r = start_r + idx
        ws.cell(r, 1, idx + 1)
        if idx < len(sample_rows):
            rv = sample_rows[idx]
            ws.cell(r, 2, rv[0])
            ws.cell(r, 3, rv[1])
            _add_dynamic_shape_pic(wb, ws, r, "Pic_C_", diag_paths, rv[1], rv[2])
            ws.cell(r, 5, rv[2])
            ws.cell(r, 6, rv[3])
            ws.cell(r, 7, rv[4])
            ws.cell(r, 8, rv[5])
            ws.cell(r, 10, rv[7])
            ws.cell(r, 11, rv[8])
            ws.cell(r, 12, rv[9])
            ws.cell(r, 13, f"=N{r}" if rv[10] == "=DF" else rv[10])
            ws.cell(r, 15, rv[11])
            ws.cell(r, 16, rv[12])
            ws.cell(r, 17, rv[13])
            fm = _floor_multiplier_expr(rv[0], rv[3], marks_with_ffr)
            ws.cell(r, 33, int(fm) if fm.isdigit() else fm)
        else:
            ws.cell(r, 3, "SC-01")
            _add_dynamic_shape_pic(wb, ws, r, "Pic_C_", diag_paths, "SC-01", 90)
            ws.cell(r, 5, 90)
            ws.cell(r, 6, "MAIN")
            ws.cell(r, 8, "F1")
            ws.cell(r, 12, "=$C$3")
            ws.cell(r, 33, 1)

        ws.cell(r, 9, f'=IF(OR(H{r}="",H{r}="NONE"),1,IFERROR(VLOOKUP(H{r},\'2. Footing BBS (Lx & Ly)\'!$B$5:$H$60,7,FALSE),1))')
        ws.cell(r, 14, f'=IF(OR(H{r}="",H{r}="NONE",NOT(ISNUMBER(SEARCH("Zone A",B{r})))),0,IFERROR(VLOOKUP(H{r},\'2. Footing BBS (Lx & Ly)\'!$B$5:$K$60,10,FALSE),0))')
        ws.cell(r, 18, f'=IF(F{r}="MAIN", $F$3*P{r}, 0)')
        ws.cell(r, 19, f'=IF(F{r}="TIE", MAX(10*P{r}, 75), 0)')
        ws.cell(r, 20, f'=IF(I{r}>0, IF(F{r}="MAIN", Q{r}, IF(Q{r}>0, ROUNDDOWN(M{r}/Q{r},0)+1, 0)), 0)')
        ws.cell(r, 21, f"=I{r}*T{r}*IF(ISNUMBER(AG{r}),AG{r},1)")
        ws.cell(
            r,
            22,
            f'=IF(I{r}>0, IF(F{r}="MAIN", ROUND((M{r}+N{r}+O{r}+IF(ISNUMBER(SEARCH("Zone C",B{r})),0,R{r}*$O$3)-IF(O{r}>0,{_angle_ded_expr(f"E{r}")}*P{r},0))/1000, 3), '
            f"ROUND((2*((J{r}-2*L{r})+(K{r}-2*L{r}))+2*S{r}-(6+2*{_angle_ded_expr(f'E{r}')})*P{r})/1000, 3)), 0)",
        )
        ws.cell(r, 23, f"=ROUND(U{r}*V{r}, 2)")
        ws.cell(r, 24, f"=IF(P{r}>0, ROUND((P{r}*P{r})/162.2, 3), 0)")
        for d_idx, dia in enumerate(DIAS):
            ws.cell(r, 25 + d_idx, f"=IF(P{r}={dia}, ROUND(W{r}*X{r}, 2), 0)")
        ws.cell(r, 32, f"=SUM(Y{r}:AE{r})")

        _style_row(ws, r, max_col, input_cols, center_cols, left_cols, fmt_map, fig_col=4)
        ws.cell(r, 14).font = FONT_CALC_BOLD
        ws.cell(r, 22).font = FONT_CALC_BOLD
        ws.cell(r, 32).font = FONT_CALC_BOLD

    end_r = start_r + total_rows - 1
    tot_r = end_r + 1
    ws.row_dimensions[tot_r].height = 24
    ws.merge_cells(start_row=tot_r, start_column=1, end_row=tot_r, end_column=20)
    lbl = ws.cell(tot_r, 1, "TOTAL COLUMN REINFORCEMENT STEEL (NET WEIGHT)")
    lbl.font = FONT_TOTAL
    lbl.alignment = ALIGN_RIGHT

    for ci in (21, 23, 25, 26, 27, 28, 29, 30, 31, 32):
        col_let = get_column_letter(ci)
        c = ws.cell(tot_r, ci, f"=SUM({col_let}{start_r}:{col_let}{end_r})")
        c.font = FONT_TOTAL
        c.alignment = ALIGN_RIGHT if ci != 21 else ALIGN_CENTER
        c.number_format = "#,##0" if ci == 21 else "#,##0.00"
    for ci in range(1, max_col + 1):
        ws.cell(tot_r, ci).fill = FILL_GREEN if ci == 32 else FILL_TOTAL
        ws.cell(tot_r, ci).border = BORDER_TOTAL

    ws.auto_filter.ref = f"A6:{get_column_letter(max_col)}{end_r}"
    _add_list_dv(ws, ALL_SHAPES, "C", start_r, end_r)
    _add_list_dv(ws, ALL_ANGLES, "E", start_r, end_r)
    _add_list_dv(ws, "MAIN,TIE", "F", start_r, end_r)
    _add_list_dv(ws, "NONE,F1,F2,F3,F4,F5,F6,SF,CF1", "H", start_r, end_r)
    _add_list_dv(ws, "8,10,12,16,20,25,32", "P", start_r, end_r)

    flookup = footing_lookup or {"F1": (6, 525), "F2": (4, 600), "F3": (2, 680), "F4": (2, 750), "F5": (2, 830), "F6": (2, 830), "SF": (2, 380), "CF1": (1, 750)}
    col_dia_totals = [0.0] * len(DIAS)
    for rv in sample_rows:
        mk, _, ang, btype, _, fref, _, b_mm, d_mm, cov_v, zh_v, bend_v, dia, q_val = rv
        nos_cols = flookup.get(fref, (1, 0))[0] if fref and fref != "NONE" else 1
        df_val = flookup.get(fref, (1, 0))[1] if (fref and fref != "NONE" and "Zone A" in mk) else 0
        cov = c_cov if cov_v == "=$C$3" else int(cov_v)
        bend = c_bend if bend_v == "=$I$3" else int(bend_v)
        if zh_v == "=DF":
            zh = float(df_val)
        elif zh_v == "=$R$3":
            zh = float(c_neck)
        elif zh_v == "=$L$3":
            zh = float(c_hc)
        elif zh_v == "=$L$3/4":
            zh = c_hc / 4.0
        elif zh_v == "=$L$3/2":
            zh = c_hc / 2.0
        elif zh_v == "=$U$3":
            zh = float(c_para)
        else:
            zh = float(zh_v)
        ded_k = 2 if ang == 90 else (1 if ang == 45 else (3 if ang == 135 else (4 if ang == 180 else 0)))
        if btype == "MAIN":
            lap_len = c_ld * dia
            bars_per = int(q_val)
            lap_add = 0.0 if "Zone C" in mk else lap_len * c_lap_f
            bend_ded = (ded_k * dia) if bend > 0 else 0.0
            cut_m = round((zh + df_val + bend + lap_add - bend_ded) / 1000.0, 3)
        else:
            bars_per = int(zh // q_val) + 1 if q_val > 0 else 0
            hook = max(10 * dia, 75)
            cut_m = round((2 * ((b_mm - 2 * cov) + (d_mm - 2 * cov)) + 2 * hook - (6 + 2 * ded_k) * dia) / 1000.0, 3)
        fm_expr = _floor_multiplier_expr(mk, btype, marks_with_ffr)
        if fm_expr == "=$X$3":
            fm_mult = c_floors
        elif fm_expr == "=MAX($X$3-1,0)":
            fm_mult = max(c_floors - 1, 0)
        else:
            fm_mult = int(fm_expr)
        tot_len = round(nos_cols * bars_per * fm_mult * cut_m, 2)
        unit_wt = round((dia * dia) / 162.2, 3)
        wt = round(tot_len * unit_wt, 2)
        if dia in DIAS:
            col_dia_totals[DIAS.index(dia)] = round(col_dia_totals[DIAS.index(dia)] + wt, 2)

    return {
        "sheet_name": ws.title,
        "total_row": tot_r,
        "dia_cols": [get_column_letter(25 + i) for i in range(len(DIAS))],
        "total_col": get_column_letter(32),
        "default_concrete_m3": round(31.50 * (max(c_floors, 1) / 3.0), 2),
        "dia_totals": col_dia_totals,
    }


# =============================================================================
# SHEET 4: BEAM BBS (TOP, BOTTOM, EXTRA SUPPORT, CRANKED & STIRRUPS)
# =============================================================================
def build_beam_sheet(wb: Workbook, diag_paths: dict[str, str]) -> dict:
    ws = wb.create_sheet("4. Beam BBS (Top, Bot, Stirrup)")
    max_col = 30  # A..AD
    _setup_sheet_banner(
        ws,
        max_col,
        "BEAM BAR BENDING SCHEDULE — DYNAMIC SHAPE FIGURE, CRANK ANGLE (°), EXTRA SUPPORT & STIRRUP HOOKS (IS 2502)",
        "YELLOW CELLS = Change Shape Code or Crank/Hook Angle (0°/45°/90°/135°/180°) and Shape Figure, Crank Factor & Bend Deduction update automatically!",
    )

    headers = [
        ("S.No", 6),                     # 1 A
        ("Beam Mark", 10),               # 2 B
        ("Shape Code", 10),              # 3 C
        ("Shape Figure", 19),            # 4 D
        ("Angle (°)", 9),                # 5 E
        ("Bar Type", 15),                # 6 F
        ("Bar Description", 31),         # 7 G
        ("No. of Beams", 10),            # 8 H
        ("Clear Span / Zone L (mm)", 12),# 9 I
        ("Width b (mm)", 10),            # 10 J
        ("Depth D (mm)", 10),            # 11 K
        ("Cover c (mm)", 9),             # 12 L
        ("Left Ld (mm)", 10),            # 13 M
        ("Right Ld (mm)", 10),           # 14 N
        ("No. of Cranks", 9),            # 15 O
        ("Dia d (mm)", 9),               # 16 P
        ("Bars/Beam or Spacing s", 12),  # 17 Q
        ("Calc Bars / Beam", 11),        # 18 R
        ("Total Bars", 10),              # 19 S
        ("Cutting Length (m)", 12),      # 20 T
        ("Total Length (m)", 12),        # 21 U
        ("Unit Wt (kg/m)", 11),          # 22 V
        ("8 mm (kg)", 11),               # 23 W
        ("10 mm (kg)", 11),              # 24 X
        ("12 mm (kg)", 11),              # 25 Y
        ("16 mm (kg)", 11),              # 26 Z
        ("20 mm (kg)", 11),              # 27 AA
        ("25 mm (kg)", 11),              # 28 AB
        ("32 mm (kg)", 11),              # 29 AC
        ("Total Wt (kg)", 13),           # 30 AD
    ]
    _write_headers(ws, 4, headers, freeze_col="H")

    sample_rows = [
        ("PB1", "SC-02", 90, "LONG", "Bottom Main Straight Bars (2×L-Hook)", 4, 4200, 230, 380, 25, 350, 350, 0, 12, 3),
        ("PB1", "SC-02", 90, "LONG", "Top Through Hanger Bars (2×L-Hook)", 4, 4200, 230, 380, 25, 350, 350, 0, 12, 2),
        ("PB1", "SC-51", 135, "STIRRUP", "Closed Stirrups (135° Seismic Hooks)", 4, 4200, 230, 380, 25, 0, 0, 0, 8, 150),
        ("B1", "SC-02", 90, "LONG", "Bottom Main Straight Bars (Through)", 4, 5200, 230, 450, 25, 450, 450, 0, 16, 3),
        ("B1", "SC-41", 45, "CRANK", "Cranked / Bent-Up Bars (Double 45° Crank)", 4, 5200, 230, 450, 25, 450, 450, 2, 16, 2),
        ("B1", "SC-02", 90, "LONG", "Top Anchor / Hanger Bars (Through)", 4, 5200, 230, 450, 25, 450, 450, 0, 12, 2),
        ("B1", "SC-01", 90, "EXTRA_SUPPORT", "Top Extra Support Bars (0.3L + End Ld)", 4, 1560, 230, 450, 25, 450, 0, 0, 16, 4),
        ("B1", "SC-00", 0, "EXTRA_SUPPORT", "Bottom Extra Mid-Span Curtailment (0.7L)", 4, 3640, 230, 450, 25, 0, 0, 0, 16, 2),
        ("B1", "SC-51", 135, "STIRRUP", "Support Zone Closed Stirrups (135° Hooks)", 4, 2600, 230, 450, 25, 0, 0, 0, 8, 100),
        ("B1", "SC-51", 135, "STIRRUP", "Mid-Span Closed Stirrups (135° Hooks)", 4, 2600, 230, 450, 25, 0, 0, 0, 8, 175),
        ("B2", "SC-02", 90, "LONG", "Bottom Main Bars (2×L-Hook)", 3, 3600, 230, 380, 25, 350, 350, 0, 12, 3),
        ("B2", "SC-02", 90, "LONG", "Top Main Bars (2×L-Hook)", 3, 3600, 230, 380, 25, 350, 350, 0, 12, 2),
        ("B2", "SC-51", 135, "STIRRUP", "Uniform Closed Stirrups (135° Hooks)", 3, 3600, 230, 380, 25, 0, 0, 0, 8, 150),
    ]

    start_r = 5
    total_rows = len(sample_rows) + BLANK_EXTRA_ROWS
    input_cols = {2, 3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17}
    center_cols = {1, 2, 3, 4, 5, 6, 8, 12, 15, 16, 17, 18, 19}
    left_cols = {7}
    fmt_map = {
        5: '0"°"', 8: "#,##0", 9: "#,##0", 10: "#,##0", 11: "#,##0", 12: "#,##0", 13: "#,##0", 14: "#,##0",
        15: "0", 16: "0", 17: "#,##0", 18: "#,##0", 19: "#,##0", 20: "0.000", 21: "#,##0.00", 22: "0.000",
    }
    for ci in range(23, 31):
        fmt_map[ci] = "#,##0.00"

    for idx in range(total_rows):
        r = start_r + idx
        ws.cell(r, 1, idx + 1)
        if idx < len(sample_rows):
            rv = sample_rows[idx]
            ws.cell(r, 2, rv[0])
            ws.cell(r, 3, rv[1])
            _add_dynamic_shape_pic(wb, ws, r, "Pic_B_", diag_paths, rv[1], rv[2])
            ws.cell(r, 5, rv[2])
            for off, val in enumerate(rv[3:], start=6):
                ws.cell(r, off, val)
        else:
            ws.cell(r, 3, "SC-02")
            _add_dynamic_shape_pic(wb, ws, r, "Pic_B_", diag_paths, "SC-02", 90)
            ws.cell(r, 5, 90)
            ws.cell(r, 6, "LONG")
            ws.cell(r, 12, 25)
            ws.cell(r, 15, 0)

        ws.cell(r, 18, f'=IF(H{r}>0, IF(F{r}="STIRRUP", IF(Q{r}>0, ROUNDDOWN(I{r}/Q{r},0)+1, 0), Q{r}), 0)')
        ws.cell(r, 19, f"=H{r}*R{r}")
        ws.cell(
            r,
            20,
            f'=IF(H{r}>0, IF(F{r}="STIRRUP", '
            f"ROUND((2*((J{r}-2*L{r})+(K{r}-2*L{r}))+20*P{r}-(6+2*{_angle_ded_expr(f'E{r}')})*P{r})/1000, 3), "
            f"ROUND((I{r}+M{r}+N{r}+O{r}*{_crank_factor_expr(f'E{r}')}*(K{r}-2*L{r})-(IF(M{r}>0,2*P{r},0)+IF(N{r}>0,2*P{r},0))-O{r}*{_angle_ded_expr(f'E{r}')}*P{r})/1000, 3)), 0)",
        )
        ws.cell(r, 21, f"=ROUND(S{r}*T{r}, 2)")
        ws.cell(r, 22, f"=IF(P{r}>0, ROUND((P{r}*P{r})/162.2, 3), 0)")
        for d_idx, dia in enumerate(DIAS):
            ws.cell(r, 23 + d_idx, f"=IF(P{r}={dia}, ROUND(U{r}*V{r}, 2), 0)")
        ws.cell(r, 30, f"=SUM(W{r}:AC{r})")

        _style_row(ws, r, max_col, input_cols, center_cols, left_cols, fmt_map, fig_col=4)
        ws.cell(r, 20).font = FONT_CALC_BOLD
        ws.cell(r, 30).font = FONT_CALC_BOLD

    end_r = start_r + total_rows - 1
    tot_r = end_r + 1
    ws.row_dimensions[tot_r].height = 24
    ws.merge_cells(start_row=tot_r, start_column=1, end_row=tot_r, end_column=18)
    lbl = ws.cell(tot_r, 1, "TOTAL BEAM REINFORCEMENT STEEL (NET WEIGHT)")
    lbl.font = FONT_TOTAL
    lbl.alignment = ALIGN_RIGHT

    for ci in (19, 21, 23, 24, 25, 26, 27, 28, 29, 30):
        col_let = get_column_letter(ci)
        c = ws.cell(tot_r, ci, f"=SUM({col_let}{start_r}:{col_let}{end_r})")
        c.font = FONT_TOTAL
        c.alignment = ALIGN_RIGHT if ci != 19 else ALIGN_CENTER
        c.number_format = "#,##0" if ci == 19 else "#,##0.00"
    for ci in range(1, max_col + 1):
        ws.cell(tot_r, ci).fill = FILL_GREEN if ci == max_col else FILL_TOTAL
        ws.cell(tot_r, ci).border = BORDER_TOTAL

    ws.auto_filter.ref = f"A4:{get_column_letter(max_col)}{end_r}"
    _add_list_dv(ws, ALL_SHAPES, "C", start_r, end_r)
    _add_list_dv(ws, ALL_ANGLES, "E", start_r, end_r)
    _add_list_dv(ws, "LONG,CRANK,EXTRA_SUPPORT,STIRRUP", "F", start_r, end_r)
    _add_list_dv(ws, "8,10,12,16,20,25,32", "P", start_r, end_r)

    return {
        "sheet_name": ws.title,
        "total_row": tot_r,
        "dia_cols": [get_column_letter(23 + i) for i in range(len(DIAS))],
        "total_col": get_column_letter(30),
        "default_concrete_m3": round((4 * 4.2 * 0.23 * 0.38) + (4 * 5.2 * 0.23 * 0.45) + (3 * 3.6 * 0.23 * 0.38), 2),
        "dia_totals": [163.01, 0.0, 185.8, 290.22, 0.0, 0.0, 0.0],
    }


# =============================================================================
# SHEET 5: SLAB BBS (MAIN, DISTRIBUTION, CRANK & TOP EXTRA SUPPORT)
# =============================================================================
def build_slab_sheet(wb: Workbook, diag_paths: dict[str, str]) -> dict:
    ws = wb.create_sheet("5. Slab BBS (Main, Dist, Crank)")
    max_col = 31  # A..AE
    _setup_sheet_banner(
        ws,
        max_col,
        "SLAB BAR BENDING SCHEDULE — DYNAMIC SHAPE FIGURE, CRANK ANGLE (°), CRANKED & EXTRA SUPPORT BARS (IS 2502)",
        "YELLOW CELLS = Change Shape Code or Crank Angle (0°/45°/90°/135°/180°) and Shape Figure, Crank Extra & Bend Deduction update automatically!",
    )

    headers = [
        ("S.No", 6),                    # 1 A
        ("Slab Mark", 10),              # 2 B
        ("Shape Code", 10),             # 3 C
        ("Shape Figure", 19),           # 4 D
        ("Angle (°)", 9),               # 5 E
        ("Bar Type", 15),               # 6 F
        ("Bar Description", 31),        # 7 G
        ("Span Dir (Lx/Ly)", 10),       # 8 H
        ("No. of Panels", 10),          # 9 I
        ("Short Span Lx (mm)", 11),     # 10 J
        ("Long Span Ly (mm)", 11),      # 11 K
        ("Thickness D (mm)", 10),       # 12 L
        ("Cover c (mm)", 9),            # 13 M
        ("Bearing / Ld (mm)", 10),      # 14 N
        ("No. of Cranks", 9),           # 15 O
        ("Dia d (mm)", 9),              # 16 P
        ("Spacing s (mm)", 10),         # 17 Q
        ("Bars / Panel", 10),           # 18 R
        ("Total Bars", 10),             # 19 S
        ("Crank Ht Hc (mm)", 11),       # 20 T
        ("Cutting Length (m)", 12),     # 21 U
        ("Total Length (m)", 12),       # 22 V
        ("Unit Wt (kg/m)", 11),         # 23 W
        ("8 mm (kg)", 11),              # 24 X
        ("10 mm (kg)", 11),             # 25 Y
        ("12 mm (kg)", 11),             # 26 Z
        ("16 mm (kg)", 11),             # 27 AA
        ("20 mm (kg)", 11),             # 28 AB
        ("25 mm (kg)", 11),             # 29 AC
        ("32 mm (kg)", 11),             # 30 AD
        ("Total Wt (kg)", 13),          # 31 AE
    ]
    _write_headers(ws, 4, headers, freeze_col="H")

    sample_rows = [
        ("S1", "SC-41", 45, "CRANK", "Two-Way Short Span Main Cranked (45°)", "Lx", 2, 3600, 4800, 125, 20, 150, 2, 10, 150),
        ("S1", "SC-41", 45, "CRANK", "Two-Way Long Span Main Cranked (45°)", "Ly", 2, 3600, 4800, 125, 20, 150, 2, 8, 150),
        ("S1", "SC-00", 0, "EXTRA_SUPPORT", "Top Extra Support Bars over Crank (L/4)", "Lx", 2, 1100, 4800, 125, 20, 150, 0, 8, 200),
        ("S2", "SC-41", 45, "CRANK", "One-Way Short Span Main Cranked (45°)", "Lx", 2, 2400, 5000, 125, 20, 150, 1, 10, 150),
        ("S2", "SC-00", 0, "DIST", "One-Way Long Span Distribution Bars", "Ly", 2, 2400, 5000, 125, 20, 150, 0, 8, 200),
        ("S3", "SC-01", 90, "EXTRA_SUPPORT", "Cantilever Top Main + Backspan Anchor", "Lx", 2, 1500, 4200, 150, 20, 600, 0, 10, 125),
        ("S3", "SC-00", 0, "DIST", "Cantilever Distribution Bars", "Ly", 2, 1500, 4200, 150, 20, 150, 0, 8, 175),
    ]

    start_r = 5
    total_rows = len(sample_rows) + BLANK_EXTRA_ROWS
    input_cols = {2, 3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17}
    center_cols = {1, 2, 3, 4, 5, 6, 8, 9, 13, 15, 16, 17, 18, 19}
    left_cols = {7}
    fmt_map = {
        5: '0"°"', 9: "#,##0", 10: "#,##0", 11: "#,##0", 12: "#,##0", 13: "#,##0", 14: "#,##0", 15: "0",
        16: "0", 17: "#,##0", 18: "#,##0", 19: "#,##0", 20: "#,##0", 21: "0.000", 22: "#,##0.00", 23: "0.000",
    }
    for ci in range(24, 32):
        fmt_map[ci] = "#,##0.00"

    for idx in range(total_rows):
        r = start_r + idx
        ws.cell(r, 1, idx + 1)
        if idx < len(sample_rows):
            rv = sample_rows[idx]
            ws.cell(r, 2, rv[0])
            ws.cell(r, 3, rv[1])
            _add_dynamic_shape_pic(wb, ws, r, "Pic_S_", diag_paths, rv[1], rv[2])
            ws.cell(r, 5, rv[2])
            for off, val in enumerate(rv[3:], start=6):
                ws.cell(r, off, val)
        else:
            ws.cell(r, 3, "SC-41")
            _add_dynamic_shape_pic(wb, ws, r, "Pic_S_", diag_paths, "SC-41", 45)
            ws.cell(r, 5, 45)
            ws.cell(r, 6, "CRANK")
            ws.cell(r, 8, "Lx")
            ws.cell(r, 13, 20)
            ws.cell(r, 14, 150)
            ws.cell(r, 15, 1)

        ws.cell(r, 18, f'=IF(AND(I{r}>0, Q{r}>0), IF(H{r}="Lx", ROUNDDOWN(K{r}/Q{r},0)+1, ROUNDDOWN(J{r}/Q{r},0)+1), 0)')
        ws.cell(r, 19, f"=I{r}*R{r}")
        ws.cell(r, 20, f"=IF(AND(I{r}>0, O{r}>0), MAX(0, L{r}-2*M{r}-P{r}), 0)")
        ws.cell(
            r,
            21,
            f'=IF(I{r}>0, ROUND((IF(H{r}="Lx",J{r},K{r})+2*N{r}+O{r}*{_crank_factor_expr(f"E{r}")}*T{r}-O{r}*{_angle_ded_expr(f"E{r}")}*P{r})/1000, 3), 0)',
        )
        ws.cell(r, 22, f"=ROUND(S{r}*U{r}, 2)")
        ws.cell(r, 23, f"=IF(P{r}>0, ROUND((P{r}*P{r})/162.2, 3), 0)")
        for d_idx, dia in enumerate(DIAS):
            ws.cell(r, 24 + d_idx, f"=IF(P{r}={dia}, ROUND(V{r}*W{r}, 2), 0)")
        ws.cell(r, 31, f"=SUM(X{r}:AD{r})")

        _style_row(ws, r, max_col, input_cols, center_cols, left_cols, fmt_map, fig_col=4)
        ws.cell(r, 21).font = FONT_CALC_BOLD
        ws.cell(r, 31).font = FONT_CALC_BOLD

    end_r = start_r + total_rows - 1
    tot_r = end_r + 1
    ws.row_dimensions[tot_r].height = 24
    ws.merge_cells(start_row=tot_r, start_column=1, end_row=tot_r, end_column=18)
    lbl = ws.cell(tot_r, 1, "TOTAL SLAB REINFORCEMENT STEEL (NET WEIGHT)")
    lbl.font = FONT_TOTAL
    lbl.alignment = ALIGN_RIGHT

    for ci in (19, 22, 24, 25, 26, 27, 28, 29, 30, 31):
        col_let = get_column_letter(ci)
        c = ws.cell(tot_r, ci, f"=SUM({col_let}{start_r}:{col_let}{end_r})")
        c.font = FONT_TOTAL
        c.alignment = ALIGN_RIGHT if ci != 19 else ALIGN_CENTER
        c.number_format = "#,##0" if ci == 19 else "#,##0.00"
    for ci in range(1, max_col + 1):
        ws.cell(tot_r, ci).fill = FILL_GREEN if ci == max_col else FILL_TOTAL
        ws.cell(tot_r, ci).border = BORDER_TOTAL

    ws.auto_filter.ref = f"A4:{get_column_letter(max_col)}{end_r}"
    _add_list_dv(ws, ALL_SHAPES, "C", start_r, end_r)
    _add_list_dv(ws, ALL_ANGLES, "E", start_r, end_r)
    _add_list_dv(ws, "CRANK,DIST,EXTRA_SUPPORT,MAIN", "F", start_r, end_r)
    _add_list_dv(ws, "Lx,Ly", "H", start_r, end_r)
    _add_list_dv(ws, "8,10,12,16,20,25,32", "P", start_r, end_r)

    return {
        "sheet_name": ws.title,
        "total_row": tot_r,
        "dia_cols": [get_column_letter(24 + i) for i in range(len(DIAS))],
        "total_col": get_column_letter(31),
        "default_concrete_m3": round((2 * 3.6 * 4.8 * 0.125) + (2 * 2.4 * 5.0 * 0.125) + (2 * 1.5 * 4.2 * 0.150), 2),
        "dia_totals": [215.77, 388.06, 0.0, 0.0, 0.0, 0.0, 0.0],
    }


# =============================================================================
# SHEET 6: SHAPE CODES, BAR TYPES & ENGINEERING DETAIL FIGURES
# =============================================================================
def build_shape_codes_sheet(wb: Workbook, diag_paths: dict[str, str]) -> None:
    ws = wb.create_sheet("6. Shape Codes & Bar Details")
    max_col = 8  # A..H
    _setup_sheet_banner(
        ws,
        max_col,
        "IS 2502 / SP 34 SHAPE CODES, ANGLES (45° / 90° / 135°), BAR TYPES & DETAILING FIGURES",
        "Visual guide for site engineers & estimators: (1) Cranked Bar Arrangement, (2) Closed Stirrup 135° Hook Detail, (3) Extra Support Bars.",
    )

    headers = [
        ("S.No", 7),
        ("Shape Code", 13),
        ("Shape Figure", 18),
        ("Bar Type", 16),
        ("Bar Description & Structural Usage", 36),
        ("Cutting Length Formula (IS 2502)", 34),
        ("Angle & Bend Deduction", 24),
        ("Used In Members", 20),
    ]
    _write_headers(ws, 4, headers, freeze_col="A")

    shape_rows = [
        (1, "SC-00", "DIST / STRAIGHT", "Straight Bar (0° Bend) — Slab distribution, mid-span extra curtailment", "L + 2 × Bearing", "0°: No bend deduction (0)", "Slab, Beam, Column"),
        (2, "SC-01", "MAIN / EXTRA_SUPPORT", "Single L-Bent Bar (90° / 45° / 135°) — Column starter, beam support extra", "A + B - Bend Ded.", "45°=1d | 90°=2d | 135°=3d", "Column, Beam, Slab"),
        (3, "SC-02", "MESH / LONG", "Double 90° U-Bar — Footing Lx/Ly mesh & beam main anchored bars", "A + B + C - (2 × 2d)", "90°: Deduct 2×2d = 4d", "Footing, Beam"),
        (4, "SC-41", "CRANK", "Cranked / Bent-Up Bar (45° / 90° / 135°) — Alternate slab & beam bars", "L + 2×Ld + n×(k×Hc) - n×Ded", "45°: k=0.42 | 90°: k=1.00", "Slab, Beam"),
        (5, "SC-51", "STIRRUP / TIE", "Closed Stirrup / Lateral Tie with 135° Seismic Hooks", "2(a + b') + 2×(10d) - 12d", "+20d Hooks; 135° Ded=12d", "Column, Beam"),
    ]

    for idx, (sn, scode, btype, bdesc, form, allow, mem) in enumerate(shape_rows):
        r = 5 + idx
        ws.row_dimensions[r].height = 35
        ws.cell(r, 1, sn).alignment = ALIGN_CENTER
        ws.cell(r, 2, scode).alignment = ALIGN_CENTER
        ws.cell(r, 2).font = FONT_CALC_BOLD
        _embed_thumb(ws, diag_paths.get(scode), f"C{r}")
        ws.cell(r, 4, btype).alignment = ALIGN_CENTER
        ws.cell(r, 4).font = FONT_CALC_BOLD
        ws.cell(r, 5, bdesc).alignment = ALIGN_LEFT
        ws.cell(r, 6, form).alignment = ALIGN_LEFT
        ws.cell(r, 7, allow).alignment = ALIGN_LEFT
        ws.cell(r, 8, mem).alignment = ALIGN_CENTER
        for ci in range(1, max_col + 1):
            ws.cell(r, ci).fill = FILL_WHITE if ci == 3 else FILL_CALC
            ws.cell(r, ci).border = BORDER_GRID

    # Dynamic Picture Lookup Gallery (Columns J & K, rows 5..22) used by Linked Pictures on Sheets 2..5
    ws.cell(4, 10, "Lookup Key").font = FONT_HDR
    ws.cell(4, 10).fill = FILL_HDR
    ws.cell(4, 10).alignment = ALIGN_CENTER
    ws.cell(4, 10).border = BORDER_HDR
    ws.column_dimensions["J"].width = 13

    ws.cell(4, 11, "Lookup Shape").font = FONT_HDR
    ws.cell(4, 11).fill = FILL_HDR
    ws.cell(4, 11).alignment = ALIGN_CENTER
    ws.cell(4, 11).border = BORDER_HDR
    ws.column_dimensions["K"].width = 19

    lookup_keys = [
        ("SC-00_0", "SC-00_0"),
        ("SC-01_45", "SC-01_45"),
        ("SC-01_90", "SC-01_90"),
        ("SC-01_135", "SC-01_135"),
        ("SC-01_180", "SC-01_180"),
        ("SC-02_45", "SC-02_45"),
        ("SC-02_90", "SC-02_90"),
        ("SC-02_135", "SC-02_135"),
        ("SC-41_45", "SC-41_45"),
        ("SC-41_90", "SC-41_90"),
        ("SC-41_135", "SC-41_135"),
        ("SC-51_90", "SC-51_90"),
        ("SC-51_135", "SC-51_135"),
        ("SC-00", "SC-00_0"),
        ("SC-01", "SC-01_90"),
        ("SC-02", "SC-02_90"),
        ("SC-41", "SC-41_45"),
        ("SC-51", "SC-51_135"),
    ]
    for off, (lkey, dkey) in enumerate(lookup_keys):
        lr = 5 + off
        ws.row_dimensions[lr].height = 33
        kc = ws.cell(lr, 10, lkey)
        kc.font = FONT_CALC_BOLD
        kc.alignment = ALIGN_CENTER
        kc.fill = FILL_CALC
        kc.border = BORDER_GRID
        vc = ws.cell(lr, 11, "")
        vc.fill = FILL_WHITE
        vc.border = BORDER_GRID
        _embed_thumb(ws, diag_paths.get(dkey), f"K{lr}")

    details = [
        ("DETAIL_CRANK", 24, "1. CRANKED BAR ARRANGEMENT (SLAB & BEAM BENT-UP DETAIL — 45° / 90° / 135°)"),
        ("DETAIL_STIRRUP", 40, "2. CLOSED STIRRUP / TIE 135° SEISMIC HOOK DETAIL"),
        ("DETAIL_EXTRA", 56, "3. EXTRA SUPPORT BARS ARRANGEMENT (TOP EXTRA AT SUPPORT & BOTTOM CURTAILMENT)"),
    ]
    for key, anchor_r, caption in details:
        ws.merge_cells(start_row=anchor_r, start_column=1, end_row=anchor_r, end_column=max_col)
        cap = ws.cell(anchor_r, 1, caption)
        cap.font = FONT_SECTION
        cap.fill = FILL_SUB
        cap.alignment = ALIGN_LEFT
        ws.row_dimensions[anchor_r].height = 22
        for ci in range(1, max_col + 1):
            ws.cell(anchor_r, ci).border = BORDER_HDR

        img_path = diag_paths.get(key)
        if img_path and os.path.exists(img_path):
            xl_img = XLImage(img_path)
            xl_img.width = 880
            xl_img.height = 278
            ws.add_image(xl_img, f"A{anchor_r + 1}")


# =============================================================================
# SHEET 1: SUMMARY & TENDER BOQ ABSTRACT (LIVE LINKS TO SHEETS 2..5)
# =============================================================================
def build_summary_sheet(
    wb: Workbook,
    refs: list[dict],
    static_summary_values: list[list[float]] | None = None,
    params: dict | None = None,
) -> None:
    ws = wb.create_sheet("1. Summary & Tender BOQ (kg)", 0)
    max_col = 11  # A..K
    _setup_sheet_banner(
        ws,
        max_col,
        "REINFORCEMENT STEEL BAR BENDING SCHEDULE — MASTER TENDER SUMMARY (IS 2502 / CPWD DSR)",
        "All weights below automatically update from Sheets 2–5. Changing Shape Code or Angle (°) updates Figures & Weights automatically.",
    )

    meta_fields = [
        (4, 1, "Project Name:", 2, "Proposed Residential / Commercial Building (Fe-500D TMT)"),
        (4, 5, "Client / Authority:", 6, "Principal Project Developer"),
        (5, 1, "Structural Drawing Ref:", 2, "STR-GA-001 (Footing, Column, Beam & Slab)"),
        (5, 5, "Steel Grade & Code:", 6, "Fe-500D TMT  |  IS 1786 / IS 2502 / SP 34"),
    ]
    for r, lc, lt, vc, vt in meta_fields:
        ws.row_dimensions[r].height = 20
        lcell = ws.cell(r, lc, lt)
        lcell.font = FONT_CALC_BOLD
        lcell.fill = FILL_CALC
        lcell.border = BORDER_GRID
        ws.merge_cells(start_row=r, start_column=vc, end_row=r, end_column=vc + 2)
        vcell = ws.cell(r, vc, vt)
        vcell.font = FONT_INPUT
        vcell.fill = FILL_INPUT
        for c_i in range(vc, vc + 3):
            ws.cell(r, c_i).border = BORDER_GRID

    headers = [
        ("S.No", 7),
        ("Structural Member Group", 32),
        ("8 mm (kg)", 13),
        ("10 mm (kg)", 13),
        ("12 mm (kg)", 13),
        ("16 mm (kg)", 13),
        ("20 mm (kg)", 13),
        ("25 mm (kg)", 13),
        ("32 mm (kg)", 13),
        ("Total Weight (kg)", 16),
        ("% Share", 12),
    ]
    hdr_r = 7
    _write_headers(ws, hdr_r, headers, freeze_col="A")

    labels = [
        "1. Isolated & Combined Footings (Lx & Ly)",
        "2. RCC Columns (Vertical Main & Ties)",
        "3. Plinth & Floor Beams (Top, Bot, Crank, Extra & Stirrups)",
        "4. RCC Slabs (Main, Dist, Crank & Extra)",
    ]

    first_data_r = 8
    last_data_r = first_data_r + len(refs) - 1
    net_tot_r = last_data_r + 1

    for idx, (label, ref) in enumerate(zip(labels, refs)):
        r = first_data_r + idx
        ws.row_dimensions[r].height = 22
        sname = ref["sheet_name"]
        trow = ref["total_row"]

        c1 = ws.cell(r, 1, idx + 1)
        c1.alignment = ALIGN_CENTER
        c2 = ws.cell(r, 2, label)
        c2.font = FONT_CALC_BOLD
        c2.alignment = ALIGN_LEFT

        for d_idx, col_let in enumerate(ref["dia_cols"]):
            val = static_summary_values[idx][d_idx] if static_summary_values else f"='{sname}'!{col_let}{trow}"
            dc = ws.cell(r, 3 + d_idx, val)
            dc.alignment = ALIGN_RIGHT
            dc.number_format = "#,##0.00"

        tot_val = f"=SUM(C{r}:I{r})" if static_summary_values else f"='{sname}'!{ref['total_col']}{trow}"
        tc = ws.cell(r, 10, tot_val)
        tc.font = FONT_CALC_BOLD
        tc.alignment = ALIGN_RIGHT
        tc.number_format = "#,##0.00"

        pc = ws.cell(r, 11, f"=IF($J${net_tot_r}>0, J{r}/$J${net_tot_r}, 0)")
        pc.alignment = ALIGN_RIGHT
        pc.number_format = "0.0%"

        for ci in range(1, max_col + 1):
            ws.cell(r, ci).fill = FILL_CALC if ci != 10 else FILL_SUB
            ws.cell(r, ci).border = BORDER_GRID

    ws.row_dimensions[net_tot_r].height = 24
    ws.merge_cells(start_row=net_tot_r, start_column=1, end_row=net_tot_r, end_column=2)
    c_net = ws.cell(net_tot_r, 1, "NET REINFORCEMENT WEIGHT BY DIAMETER (kg)")
    c_net.font = FONT_TOTAL
    c_net.alignment = ALIGN_RIGHT
    for ci in range(3, 11):
        col_let = get_column_letter(ci)
        c = ws.cell(net_tot_r, ci, f"=SUM({col_let}{first_data_r}:{col_let}{last_data_r})")
        c.font = FONT_TOTAL
        c.alignment = ALIGN_RIGHT
        c.number_format = "#,##0.00"
    p_tot = ws.cell(net_tot_r, 11, f"=SUM(K{first_data_r}:K{last_data_r})")
    p_tot.font = FONT_TOTAL
    p_tot.alignment = ALIGN_RIGHT
    p_tot.number_format = "0.0%"
    for ci in range(1, max_col + 1):
        ws.cell(net_tot_r, ci).fill = FILL_TOTAL
        ws.cell(net_tot_r, ci).border = BORDER_TOTAL

    wast_r = net_tot_r + 1
    ws.row_dimensions[wast_r].height = 22
    ws.cell(wast_r, 1, "").border = BORDER_GRID
    ws.cell(wast_r, 2, "Add: Rolling Margin & Wastage (%)").font = FONT_CALC_BOLD
    ws.cell(wast_r, 2).border = BORDER_GRID
    wpct = ws.cell(wast_r, 3, float((params or {}).get("wastage_pct", 0.03)))
    wpct.fill = FILL_INPUT
    wpct.font = FONT_INPUT
    wpct.alignment = ALIGN_CENTER
    wpct.number_format = "0.0%"
    wpct.border = BORDER_GRID
    for ci in range(4, 10):
        ws.cell(wast_r, ci, "").border = BORDER_GRID
    wkg = ws.cell(wast_r, 10, f"=ROUND(J{net_tot_r}*C{wast_r}, 2)")
    wkg.font = FONT_CALC
    wkg.alignment = ALIGN_RIGHT
    wkg.number_format = "#,##0.00"
    wkg.border = BORDER_GRID
    ws.cell(wast_r, 11, "kg").alignment = ALIGN_CENTER
    ws.cell(wast_r, 11).border = BORDER_GRID

    gross_r = wast_r + 1
    ws.row_dimensions[gross_r].height = 24
    ws.merge_cells(start_row=gross_r, start_column=1, end_row=gross_r, end_column=9)
    ws.cell(gross_r, 1, "GROSS TENDER STEEL QUANTITY (INCLUDING WASTAGE) — KG").font = FONT_TOTAL
    ws.cell(gross_r, 1).alignment = ALIGN_RIGHT
    g_kg = ws.cell(gross_r, 10, f"=ROUND(J{net_tot_r}+J{wast_r}, 2)")
    g_kg.font = FONT_TOTAL
    g_kg.alignment = ALIGN_RIGHT
    g_kg.number_format = "#,##0.00"
    ws.cell(gross_r, 11, "kg").alignment = ALIGN_CENTER
    ws.cell(gross_r, 11).font = FONT_TOTAL
    for ci in range(1, max_col + 1):
        ws.cell(gross_r, ci).fill = FILL_GREEN
        ws.cell(gross_r, ci).border = BORDER_TOTAL

    mt_r = gross_r + 1
    ws.row_dimensions[mt_r].height = 22
    ws.merge_cells(start_row=mt_r, start_column=1, end_row=mt_r, end_column=9)
    ws.cell(mt_r, 1, "GROSS TENDER STEEL IN METRIC TONNES (MT = kg / 1000)").font = FONT_CALC_BOLD
    ws.cell(mt_r, 1).alignment = ALIGN_RIGHT
    g_mt = ws.cell(mt_r, 10, f"=ROUND(J{gross_r}/1000, 3)")
    g_mt.font = FONT_CALC_BOLD
    g_mt.alignment = ALIGN_RIGHT
    g_mt.number_format = "#,##0.000"
    ws.cell(mt_r, 11, "MT").alignment = ALIGN_CENTER
    for ci in range(1, max_col + 1):
        ws.cell(mt_r, ci).fill = FILL_CALC
        ws.cell(mt_r, ci).border = BORDER_GRID

    rate_r = mt_r + 1
    ws.row_dimensions[rate_r].height = 24
    ws.merge_cells(start_row=rate_r, start_column=1, end_row=rate_r, end_column=7)
    ws.cell(rate_r, 1, "Supply, Cutting, Bending & Binding Fe-500D TMT Steel — Tender Rate (Rs/kg):").font = FONT_CALC_BOLD
    ws.cell(rate_r, 1).alignment = ALIGN_RIGHT
    rcell = ws.cell(rate_r, 8, float((params or {}).get("tender_rate_inr", 82.00)))
    rcell.fill = FILL_INPUT
    rcell.font = FONT_INPUT
    rcell.alignment = ALIGN_RIGHT
    rcell.number_format = "₹#,##0.00"
    ws.cell(rate_r, 9, "BOQ Amt:").font = FONT_TOTAL
    ws.cell(rate_r, 9).alignment = ALIGN_RIGHT
    acell = ws.cell(rate_r, 10, f"=ROUND(J{gross_r}*H{rate_r}, 2)")
    acell.font = FONT_TOTAL
    acell.alignment = ALIGN_RIGHT
    acell.number_format = "₹#,##0.00"
    ws.cell(rate_r, 11, "INR").alignment = ALIGN_CENTER
    ws.cell(rate_r, 11).font = FONT_TOTAL
    for ci in range(1, max_col + 1):
        if ci != 8:
            ws.cell(rate_r, ci).fill = FILL_SUB
        ws.cell(rate_r, ci).border = BORDER_TOTAL

    chk_hdr_r = rate_r + 2
    ws.row_dimensions[chk_hdr_r].height = 24
    ws.merge_cells(start_row=chk_hdr_r, start_column=1, end_row=chk_hdr_r, end_column=7)
    sh = ws.cell(chk_hdr_r, 1, "CONCRETE VOLUME vs STEEL CONSUMPTION THUMB-RULE CHECK (kg / Cu.m)")
    sh.font = FONT_SECTION
    sh.fill = FILL_SUB
    sh.alignment = ALIGN_LEFT
    for ci in range(1, 8):
        ws.cell(chk_hdr_r, ci).border = BORDER_HDR

    chk_col_r = chk_hdr_r + 1
    chk_cols = [
        (1, "S.No"),
        (2, "Member Group"),
        (3, "Net Steel (kg)"),
        (4, "Concrete Vol (Cu.m)"),
        (5, "Steel Ratio (kg/Cu.m)"),
        (6, "IS / CPWD Typical Range"),
        (7, "Status Check"),
    ]
    ws.row_dimensions[chk_col_r].height = 24
    for ci, txt in chk_cols:
        c = ws.cell(chk_col_r, ci, txt)
        c.font = FONT_HDR
        c.fill = FILL_HDR
        c.alignment = ALIGN_CENTER
        c.border = BORDER_HDR

    norms = [
        ("Footing", 50, 95, "50 - 95 kg/Cu.m"),
        ("Column", 150, 320, "150 - 320 kg/Cu.m"),
        ("Beam", 110, 200, "110 - 200 kg/Cu.m"),
        ("Slab", 70, 120, "70 - 120 kg/Cu.m"),
    ]
    chk_first = chk_col_r + 1
    for i, ((mname, lo, hi, norm_txt), ref) in enumerate(zip(norms, refs)):
        r = chk_first + i
        src_r = first_data_r + i
        ws.row_dimensions[r].height = 20
        ws.cell(r, 1, i + 1).alignment = ALIGN_CENTER
        ws.cell(r, 2, mname).alignment = ALIGN_LEFT
        ws.cell(r, 2).font = FONT_CALC_BOLD
        ws.cell(r, 3, f"=J{src_r}").number_format = "#,##0.00"
        vc = ws.cell(r, 4, ref["default_concrete_m3"])
        vc.fill = FILL_INPUT
        vc.font = FONT_INPUT
        vc.number_format = "#,##0.00"
        ws.cell(r, 5, f"=IF(D{r}>0, ROUND(C{r}/D{r}, 1), 0)").number_format = "#,##0.0"
        ws.cell(r, 6, norm_txt).alignment = ALIGN_CENTER
        ws.cell(r, 7, f'=IF(D{r}<=0,"Enter Vol",IF(AND(E{r}>={lo},E{r}<={hi}),"OK (Typical)","Check Design"))').alignment = ALIGN_CENTER
        for ci in range(1, 8):
            if ci != 4:
                ws.cell(r, ci).fill = FILL_CALC
                if ci != 2:
                    ws.cell(r, ci).font = FONT_CALC
            ws.cell(r, ci).border = BORDER_GRID

    ref_notes = [
        ("IS 2502 / SP 34 FORMULA REFERENCE", True),
        ("Unit Weight (kg/m) = d² / 162.2", False),
        ("45° Bend = 1d  |  90° = 2d  |  135° = 3d", False),
        ("180° Hook = 4d  |  135° Stirrup Hook = 10d", False),
        ("45° Crank = 0.42×Hc  |  90° Step = 1.00×Hc", False),
        ("135° Crank = 1.30×Hc", False),
    ]
    for off, (txt, is_h) in enumerate(ref_notes):
        rr = chk_hdr_r + off
        ws.merge_cells(start_row=rr, start_column=9, end_row=rr, end_column=11)
        rc = ws.cell(rr, 9, txt)
        rc.font = FONT_HDR if is_h else FONT_NOTE
        rc.fill = FILL_HDR if is_h else FILL_NOTE
        rc.alignment = ALIGN_CENTER if is_h else ALIGN_LEFT
        for ci in range(9, 12):
            ws.cell(rr, ci).border = BORDER_GRID


def _bind_pictures_and_export_pdf(xlsx_path: str, pdf_path: str | None = None) -> None:
    ps1 = os.path.join(BASE_DIR, "link_excel_pictures.ps1")
    if os.path.exists(ps1):
        try:
            subprocess.run(
                ["powershell", "-ExecutionPolicy", "Bypass", "-File", ps1, "-WorkbookPath", xlsx_path],
                check=True,
                capture_output=True,
                text=True,
            )
        except Exception as exc:
            print(f"Warning: Could not run Excel Linked Picture binder for {xlsx_path}:", exc)

    if pdf_path:
        pdf_ps1 = os.path.join(BASE_DIR, "export_xlsx_to_pdf.ps1")
        if os.path.exists(pdf_ps1):
            try:
                subprocess.run(
                    ["powershell", "-ExecutionPolicy", "Bypass", "-File", pdf_ps1, "-XlsxPath", xlsx_path, "-PdfPath", pdf_path],
                    check=True,
                    capture_output=True,
                    text=True,
                )
            except Exception as exc:
                print(f"Warning: Could not auto-export PDF for {xlsx_path}:", exc)


def generate_individual_trade_workbooks(
    diag_paths: dict[str, str],
    refs: list[dict],
    master_xlsx_path: str,
) -> list[str]:
    """Generate standalone trade-specific workbooks in exports/ with zero broken links."""
    from openpyxl import load_workbook

    generated: list[str] = []

    # 1. Footing_BBS_Only.xlsx
    wb_f = Workbook()
    d_ws = wb_f.active
    build_footing_sheet(wb_f, diag_paths)
    build_shape_codes_sheet(wb_f, diag_paths)
    wb_f.remove(d_ws)
    f_xlsx = os.path.join(EXPORTS_DIR, "Footing_BBS_Only.xlsx")
    wb_f.save(f_xlsx)
    _bind_pictures_and_export_pdf(f_xlsx)
    generated.append(f_xlsx)

    # 2. Column_BBS_Only.xlsx (includes Sheet 2 lookup table + Sheet 6 diagrams so VLOOKUP & Pictures never #REF!)
    wb_c = Workbook()
    d_ws = wb_c.active
    ws_f_ref = build_footing_sheet(wb_c, diag_paths)
    ws_c_ref = build_column_sheet(wb_c, diag_paths)
    build_shape_codes_sheet(wb_c, diag_paths)
    wb_c.remove(d_ws)
    col_ws = wb_c[ws_c_ref["sheet_name"]]
    ftg_ws = wb_c[ws_f_ref["sheet_name"]]
    shp_ws = wb_c["6. Shape Codes & Bar Details"]
    wb_c._sheets = [col_ws, ftg_ws, shp_ws]
    wb_c.active = 0
    c_xlsx = os.path.join(EXPORTS_DIR, "Column_BBS_Only.xlsx")
    wb_c.save(c_xlsx)
    _bind_pictures_and_export_pdf(c_xlsx)
    generated.append(c_xlsx)

    # 3. Beam_BBS_Only.xlsx
    wb_b = Workbook()
    d_ws = wb_b.active
    build_beam_sheet(wb_b, diag_paths)
    build_shape_codes_sheet(wb_b, diag_paths)
    wb_b.remove(d_ws)
    b_xlsx = os.path.join(EXPORTS_DIR, "Beam_BBS_Only.xlsx")
    wb_b.save(b_xlsx)
    _bind_pictures_and_export_pdf(b_xlsx)
    generated.append(b_xlsx)

    # 4. Slab_BBS_Only.xlsx
    wb_s = Workbook()
    d_ws = wb_s.active
    build_slab_sheet(wb_s, diag_paths)
    build_shape_codes_sheet(wb_s, diag_paths)
    wb_s.remove(d_ws)
    s_xlsx = os.path.join(EXPORTS_DIR, "Slab_BBS_Only.xlsx")
    wb_s.save(s_xlsx)
    _bind_pictures_and_export_pdf(s_xlsx)
    generated.append(s_xlsx)

    # 5. Summary_Tender_BOQ.xlsx (self-contained with evaluated master sheet totals in C8:I11)
    wb_master_val = load_workbook(master_xlsx_path, data_only=True)
    ws_master_sum = wb_master_val["1. Summary & Tender BOQ (kg)"]
    static_vals: list[list[float]] = []
    for r in range(8, 12):
        row_vals = [float(ws_master_sum.cell(r, c).value or 0.0) for c in range(3, 10)]
        static_vals.append(row_vals)
    wb_master_val.close()

    wb_sum = Workbook()
    d_ws = wb_sum.active
    build_summary_sheet(wb_sum, refs, static_summary_values=static_vals)
    wb_sum.remove(d_ws)
    sum_xlsx = os.path.join(EXPORTS_DIR, "Summary_Tender_BOQ.xlsx")
    wb_sum.save(sum_xlsx)
    generated.append(sum_xlsx)

    return generated


def _wb_to_bytes(wb: Workbook) -> bytes:
    import io

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def compile_bbs_workbooks_bytes(params: dict | None = None) -> dict[str, bytes]:
    """Dynamically compile all 6 BBS workbooks in memory from live UI parameters (zero #REF! errors)."""
    diag_paths = ensure_bbs_diagrams()
    out: dict[str, bytes] = {}

    # 1. Master Combined Workbook
    wb_m = Workbook()
    d_ws = wb_m.active
    f_ref = build_footing_sheet(wb_m, diag_paths, params=params)
    c_ref = build_column_sheet(wb_m, diag_paths, params=params, footing_lookup=f_ref["footing_lookup"])
    b_ref = build_beam_sheet(wb_m, diag_paths)
    s_ref = build_slab_sheet(wb_m, diag_paths)
    build_shape_codes_sheet(wb_m, diag_paths)
    refs = [f_ref, c_ref, b_ref, s_ref]
    build_summary_sheet(wb_m, refs, params=params)
    wb_m.remove(d_ws)
    out[OUTPUT_FILENAME] = _wb_to_bytes(wb_m)

    # 2. Footing_BBS_Only.xlsx
    wb_f = Workbook()
    d_ws = wb_f.active
    build_footing_sheet(wb_f, diag_paths, params=params)
    build_shape_codes_sheet(wb_f, diag_paths)
    wb_f.remove(d_ws)
    out["Footing_BBS_Only.xlsx"] = _wb_to_bytes(wb_f)

    # 3. Column_BBS_Only.xlsx (Sheet 2 lookup + Sheet 3 active + Sheet 6 diagrams -> 0 #REF!)
    wb_c = Workbook()
    d_ws = wb_c.active
    ws_f_ref = build_footing_sheet(wb_c, diag_paths, params=params)
    ws_c_ref = build_column_sheet(wb_c, diag_paths, params=params, footing_lookup=ws_f_ref["footing_lookup"])
    build_shape_codes_sheet(wb_c, diag_paths)
    wb_c.remove(d_ws)
    wb_c._sheets = [wb_c[ws_c_ref["sheet_name"]], wb_c[ws_f_ref["sheet_name"]], wb_c["6. Shape Codes & Bar Details"]]
    wb_c.active = 0
    out["Column_BBS_Only.xlsx"] = _wb_to_bytes(wb_c)

    # 4. Beam_BBS_Only.xlsx
    wb_b = Workbook()
    d_ws = wb_b.active
    build_beam_sheet(wb_b, diag_paths)
    build_shape_codes_sheet(wb_b, diag_paths)
    wb_b.remove(d_ws)
    out["Beam_BBS_Only.xlsx"] = _wb_to_bytes(wb_b)

    # 5. Slab_BBS_Only.xlsx
    wb_s = Workbook()
    d_ws = wb_s.active
    build_slab_sheet(wb_s, diag_paths)
    build_shape_codes_sheet(wb_s, diag_paths)
    wb_s.remove(d_ws)
    out["Slab_BBS_Only.xlsx"] = _wb_to_bytes(wb_s)

    # 6. Summary_Tender_BOQ.xlsx (self-contained numeric trade totals, live SUM/wastage/INR/kg-per-m3 formulas)
    static_vals = [r["dia_totals"] for r in refs]
    wb_sum = Workbook()
    d_ws = wb_sum.active
    build_summary_sheet(wb_sum, refs, static_summary_values=static_vals, params=params)
    wb_sum.remove(d_ws)
    out["Summary_Tender_BOQ.xlsx"] = _wb_to_bytes(wb_sum)

    return out


def generate_standalone_bbs_workbook() -> list[str]:
    os.makedirs(EXPORTS_DIR, exist_ok=True)
    diag_paths = ensure_bbs_diagrams()

    wb = Workbook()
    default_ws = wb.active

    f_ref = build_footing_sheet(wb, diag_paths)
    c_ref = build_column_sheet(wb, diag_paths)
    b_ref = build_beam_sheet(wb, diag_paths)
    s_ref = build_slab_sheet(wb, diag_paths)
    build_shape_codes_sheet(wb, diag_paths)
    refs = [f_ref, c_ref, b_ref, s_ref]
    build_summary_sheet(wb, refs)

    wb.remove(default_ws)

    root_out = os.path.join(BASE_DIR, OUTPUT_FILENAME)
    export_out = os.path.join(EXPORTS_DIR, OUTPUT_FILENAME)
    pdf_filename = os.path.splitext(OUTPUT_FILENAME)[0] + ".pdf"
    root_pdf = os.path.join(BASE_DIR, pdf_filename)
    export_pdf = os.path.join(EXPORTS_DIR, pdf_filename)
    wb.save(root_out)

    _bind_pictures_and_export_pdf(root_out, root_pdf)
    shutil.copy2(root_out, export_out)
    if os.path.exists(root_pdf):
        shutil.copy2(root_pdf, export_pdf)

    trade_files = generate_individual_trade_workbooks(diag_paths, refs, root_out)
    return [root_out, export_out, root_pdf, export_pdf, *trade_files]


# Explicit public aliases for cross-module imports
generate_bbs_master = generate_standalone_bbs_workbook
export_all_bbs_workbooks = generate_standalone_bbs_workbook


if __name__ == "__main__":
    paths = generate_standalone_bbs_workbook()
    for p in paths:
        print("Generated:", p)
