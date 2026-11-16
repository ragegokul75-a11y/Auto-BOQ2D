"""
Multi-Sheet Excel BOQ, Measurement Sheet (Takeoff), and General Cost Abstract Generator.
Generates an industry-grade Indian Standard (IS 1200 / CPWD DSR) Excel Workbook (.xlsx)
with live cross-sheet formulas (`=SUM()`, `'3. Measurement Sheet (Takeoff)'!H...`, GST, Contingency).
"""

import io
from datetime import datetime
from typing import Dict, List, Any
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from bbs_engine import generate_bbs_from_takeoff, summarize_bbs


def generate_boq_excel(payload: Dict[str, Any]) -> bytes:
    """
    Creates a 4-sheet Excel workbook with live formulas:
      Sheet 1: "1. Abstract of Cost"
      Sheet 2: "2. Bill of Quantities (BOQ)"
      Sheet 3: "3. Measurement Sheet (Takeoff)"
      Sheet 4: "4. Bar Bending Schedule (BBS)"
    """
    project_name = payload.get("project_name", "Proposed Landscape, Residential & Commercial Development")
    client_name = payload.get("client_name", "M/s Principal Employer / Project Developer")
    consultant_name = payload.get("consultant_name", "Automated 2D CAD QS & BOQ Studio")
    drawing_no = payload.get("drawing_no", "CAD-2D-GA-001")
    project_type = payload.get("project_type", "Landscape, Residential & Commercial")
    contingency_pct = float(payload.get("contingency_pct", 3.0)) / 100.0
    gst_pct = float(payload.get("gst_pct", 18.0)) / 100.0
    date_str = datetime.now().strftime("%d-%b-%Y")

    groups: List[Dict[str, Any]] = [
        g for g in payload.get("takeoff_groups", []) if g.get("included", True) and len(g.get("measurements", [])) > 0
    ]

    wb = openpyxl.Workbook()
    # Create the 4 sheets in order
    ws_abs = wb.active
    ws_abs.title = "1. Abstract of Cost"
    ws_boq = wb.create_sheet(title="2. Bill of Quantities (BOQ)")
    ws_meas = wb.create_sheet(title="3. Measurement Sheet (Takeoff)")
    ws_bbs = wb.create_sheet(title="4. Bar Bending Schedule (BBS)")

    # Styling palette
    navy_header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    slate_sub_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    cat_banner_fill = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
    sub_total_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    grand_total_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    font_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    font_subtitle = Font(name="Calibri", size=10, italic=True, color="E2E8F0")
    font_tbl_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_cat = Font(name="Calibri", size=10, bold=True, color="0369A1")
    font_bold = Font(name="Calibri", size=10, bold=True, color="0F172A")
    font_regular = Font(name="Calibri", size=10, bold=False, color="1E293B")
    font_muted = Font(name="Calibri", size=9, italic=True, color="475569")
    font_grand = Font(name="Calibri", size=11, bold=True, color="166534")

    thin_side = Side(border_style="thin", color="CBD5E1")
    double_bottom = Side(border_style="double", color="0F172A")
    top_thin = Side(border_style="thin", color="0F172A")

    grid_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    total_border = Border(left=thin_side, right=thin_side, top=top_thin, bottom=double_bottom)

    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")

    inr_format = '₹ #,##0.00'
    qty_format = '#,##0.00'

    def write_project_banner(ws, max_col: int, sheet_subtitle: str):
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
        ws.cell(1, 1, f"{project_name.upper()} — {sheet_subtitle.upper()}").font = font_title
        ws.cell(1, 1).fill = navy_header_fill
        ws.cell(1, 1).alignment = align_center
        ws.row_dimensions[1].height = 28

        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
        meta_text = (
            f"Client: {client_name}   |   Project Type: {project_type}   |   "
            f"Drawing Ref: {drawing_no}   |   Standard: IS 1200 / CPWD DSR   |   Date: {date_str}"
        )
        ws.cell(2, 1, meta_text).font = font_subtitle
        ws.cell(2, 1).fill = slate_sub_fill
        ws.cell(2, 1).alignment = align_center
        ws.row_dimensions[2].height = 20

    # =========================================================================
    # STEP 1: POPULATE SHEET 3 — "3. Measurement Sheet (Takeoff)" FIRST
    # so we know the exact cell coordinates of each item's Total Quantity!
    # =========================================================================
    write_project_banner(ws_meas, 9, "Detailed 2D CAD Measurement Sheet (IS 1200 Takeoff)")

    meas_headers = [
        ("Item No.", 10),
        ("DSR Code", 13),
        ("CAD Layer / Element Description & Location", 44),
        ("Nos", 10),
        ("Length (L) m", 14),
        ("Breadth/Width (B) m", 18),
        ("Depth/Height (D) m", 18),
        ("Quantity / Area", 16),
        ("Unit & Remarks", 16),
    ]
    ws_meas.row_dimensions[4].height = 24
    for col_idx, (h_text, width) in enumerate(meas_headers, start=1):
        c = ws_meas.cell(4, col_idx, h_text)
        c.font = font_tbl_header
        c.fill = navy_header_fill
        c.alignment = align_center
        c.border = grid_border
        ws_meas.column_dimensions[get_column_letter(col_idx)].width = width

    # Group items by category to keep consistent ordering across all 3 sheets
    categories_order: List[str] = []
    groups_by_cat: Dict[str, List[Dict[str, Any]]] = {}
    for g in groups:
        cat = g.get("category", "General Works")
        if cat not in groups_by_cat:
            categories_order.append(cat)
            groups_by_cat[cat] = []
        groups_by_cat[cat].append(g)

    meas_row = 5
    # Map group_id -> row number in Sheet 3 where that group's Total Quantity resides
    group_meas_total_row: Dict[str, int] = {}

    item_serial = 0
    for cat in categories_order:
        # Category header row
        ws_meas.merge_cells(start_row=meas_row, start_column=1, end_row=meas_row, end_column=9)
        cat_cell = ws_meas.cell(meas_row, 1, f"SECTION: {cat.upper()}")
        cat_cell.font = font_cat
        cat_cell.fill = cat_banner_fill
        cat_cell.alignment = align_left
        ws_meas.row_dimensions[meas_row].height = 22
        for ci in range(1, 10):
            ws_meas.cell(meas_row, ci).border = grid_border
        meas_row += 1

        for grp in groups_by_cat[cat]:
            item_serial += 1
            gid = grp["group_id"]
            calc_mode = grp.get("calc_mode", "AREA")
            unit = grp.get("unit", "Sq.m")

            # Item Sub-header in Measurement sheet
            ws_meas.cell(meas_row, 1, f"{item_serial}.0").font = font_bold
            ws_meas.cell(meas_row, 1).alignment = align_center
            ws_meas.cell(meas_row, 2, grp.get("dsr_code", "")).font = font_bold
            ws_meas.cell(meas_row, 2).alignment = align_center
            ws_meas.merge_cells(start_row=meas_row, start_column=3, end_row=meas_row, end_column=8)
            ws_meas.cell(meas_row, 3, f"{grp['display_name']} — {grp['description'][:110]}...").font = font_bold
            ws_meas.cell(meas_row, 3).alignment = align_left
            ws_meas.cell(meas_row, 9, unit).font = font_bold
            ws_meas.cell(meas_row, 9).alignment = align_center
            for ci in range(1, 10):
                ws_meas.cell(meas_row, ci).fill = zebra_fill
                ws_meas.cell(meas_row, ci).border = grid_border
            ws_meas.row_dimensions[meas_row].height = 20
            meas_row += 1

            start_sub_row = meas_row
            measurements = grp.get("measurements", [])
            for sub_idx, m in enumerate(measurements, start=1):
                c1 = ws_meas.cell(meas_row, 1, f"{item_serial}.{sub_idx}")
                c1.font = font_muted
                c1.alignment = align_center

                c2 = ws_meas.cell(meas_row, 2, grp.get("layer", ""))
                c2.font = font_muted
                c2.alignment = align_center

                c3 = ws_meas.cell(meas_row, 3, m.get("label", f"Element #{sub_idx}"))
                c3.font = font_regular
                c3.alignment = align_left

                c4 = ws_meas.cell(meas_row, 4, float(m.get("nos", 1)))
                c4.font = font_regular
                c4.alignment = align_center
                c4.number_format = '#,##0'

                # Populate L, B, D depending on calc_mode
                l_val = float(m.get("length", 1.0))
                b_val = float(m.get("breadth", 1.0))
                d_val = float(m.get("depth", 1.0))

                c5 = ws_meas.cell(meas_row, 5, l_val)
                c6 = ws_meas.cell(meas_row, 6, b_val)
                c7 = ws_meas.cell(meas_row, 7, d_val)
                for cn in (c5, c6, c7):
                    cn.font = font_regular
                    cn.alignment = align_right
                    cn.number_format = qty_format

                # Live Excel Formula for Quantity = D{row} * E{row} * F{row} * G{row}
                c8 = ws_meas.cell(meas_row, 8, f"=ROUND(D{meas_row}*E{meas_row}*F{meas_row}*G{meas_row}, 3)")
                c8.font = font_regular
                c8.alignment = align_right
                c8.number_format = qty_format

                remark = "CAD Polygon Shoelace" if m.get("raw_area", 0) > 0 else f"Mode: {calc_mode}"
                c9 = ws_meas.cell(meas_row, 9, remark)
                c9.font = font_muted
                c9.alignment = align_center

                for ci in range(1, 10):
                    ws_meas.cell(meas_row, ci).border = grid_border
                ws_meas.row_dimensions[meas_row].height = 18
                meas_row += 1

            end_sub_row = meas_row - 1

            # Subtotal Row for this BOQ Item
            ws_meas.merge_cells(start_row=meas_row, start_column=1, end_row=meas_row, end_column=7)
            tot_lbl = ws_meas.cell(meas_row, 1, f"Total Measured Quantity for Item {item_serial}.0 ({grp['display_name']})")
            tot_lbl.font = font_bold
            tot_lbl.alignment = align_right

            tot_val = ws_meas.cell(meas_row, 8, f"=ROUND(SUM(H{start_sub_row}:H{end_sub_row}), 2)")
            tot_val.font = font_bold
            tot_val.alignment = align_right
            tot_val.number_format = qty_format

            tot_u = ws_meas.cell(meas_row, 9, unit)
            tot_u.font = font_bold
            tot_u.alignment = align_center

            for ci in range(1, 10):
                ws_meas.cell(meas_row, ci).fill = sub_total_fill
                ws_meas.cell(meas_row, ci).border = total_border
            ws_meas.row_dimensions[meas_row].height = 21

            group_meas_total_row[gid] = meas_row
            meas_row += 2  # Blank spacing row after each item

    # =========================================================================
    # STEP 2: POPULATE SHEET 2 — "2. Bill of Quantities (BOQ)"
    # Referencing Sheet 3 Totals via live Excel formulas!
    # =========================================================================
    write_project_banner(ws_boq, 9, "Detailed Bill of Quantities (IS 1200 / CPWD DSR Schedule)")

    boq_headers = [
        ("S.No.", 8),
        ("DSR Code", 13),
        ("IS / DSR Ref", 20),
        ("CAD Layer / Source", 22),
        ("Item Description & Technical Specification", 54),
        ("Quantity", 15),
        ("Unit", 10),
        ("Rate (INR)", 16),
        ("Amount (INR)", 18),
    ]
    ws_boq.row_dimensions[4].height = 25
    for col_idx, (h_text, width) in enumerate(boq_headers, start=1):
        c = ws_boq.cell(4, col_idx, h_text)
        c.font = font_tbl_header
        c.fill = navy_header_fill
        c.alignment = align_center
        c.border = grid_border
        ws_boq.column_dimensions[get_column_letter(col_idx)].width = width

    boq_row = 5
    cat_boq_subtotal_rows: Dict[str, int] = {}
    item_serial = 0

    for cat_idx, cat in enumerate(categories_order, start=1):
        ws_boq.merge_cells(start_row=boq_row, start_column=1, end_row=boq_row, end_column=9)
        cat_cell = ws_boq.cell(boq_row, 1, f"PART {chr(64 + cat_idx)} — {cat.upper()}")
        cat_cell.font = font_cat
        cat_cell.fill = cat_banner_fill
        cat_cell.alignment = align_left
        for ci in range(1, 10):
            ws_boq.cell(boq_row, ci).border = grid_border
        ws_boq.row_dimensions[boq_row].height = 23
        boq_row += 1

        cat_first_row = boq_row
        for grp in groups_by_cat[cat]:
            item_serial += 1
            gid = grp["group_id"]
            meas_tot_r = group_meas_total_row[gid]

            c1 = ws_boq.cell(boq_row, 1, item_serial)
            c1.font = font_bold
            c1.alignment = align_center

            c2 = ws_boq.cell(boq_row, 2, grp.get("dsr_code", ""))
            c2.font = font_bold
            c2.alignment = align_center

            c3 = ws_boq.cell(boq_row, 3, grp.get("is_ref", "IS 1200"))
            c3.font = font_muted
            c3.alignment = align_center

            c4 = ws_boq.cell(boq_row, 4, grp.get("display_name", ""))
            c4.font = font_regular
            c4.alignment = align_left

            c5 = ws_boq.cell(boq_row, 5, grp.get("description", ""))
            c5.font = font_regular
            c5.alignment = align_left

            # Live formula pulling Quantity directly from Sheet 3!
            c6 = ws_boq.cell(boq_row, 6, f"='3. Measurement Sheet (Takeoff)'!H{meas_tot_r}")
            c6.font = font_bold
            c6.alignment = align_right
            c6.number_format = qty_format

            c7 = ws_boq.cell(boq_row, 7, grp.get("unit", "Sq.m"))
            c7.font = font_bold
            c7.alignment = align_center

            c8 = ws_boq.cell(boq_row, 8, float(grp.get("rate", 0.0)))
            c8.font = font_regular
            c8.alignment = align_right
            c8.number_format = inr_format

            # Live formula: Amount = Quantity * Rate
            c9 = ws_boq.cell(boq_row, 9, f"=ROUND(F{boq_row}*H{boq_row}, 2)")
            c9.font = font_bold
            c9.alignment = align_right
            c9.number_format = inr_format

            for ci in range(1, 10):
                ws_boq.cell(boq_row, ci).border = grid_border
                if item_serial % 2 == 0:
                    ws_boq.cell(boq_row, ci).fill = zebra_fill

            ws_boq.row_dimensions[boq_row].height = 42
            boq_row += 1

        cat_last_row = boq_row - 1
        # Category Sub-Total Row in BOQ
        ws_boq.merge_cells(start_row=boq_row, start_column=1, end_row=boq_row, end_column=8)
        sub_lbl = ws_boq.cell(boq_row, 1, f"Sub-Total for Part {chr(64 + cat_idx)} — {cat}")
        sub_lbl.font = font_bold
        sub_lbl.alignment = align_right

        sub_amt = ws_boq.cell(boq_row, 9, f"=SUM(I{cat_first_row}:I{cat_last_row})")
        sub_amt.font = font_bold
        sub_amt.alignment = align_right
        sub_amt.number_format = inr_format

        for ci in range(1, 10):
            ws_boq.cell(boq_row, ci).fill = sub_total_fill
            ws_boq.cell(boq_row, ci).border = total_border
        ws_boq.row_dimensions[boq_row].height = 22

        cat_boq_subtotal_rows[cat] = boq_row
        boq_row += 2

    # Grand Total at bottom of BOQ Sheet
    subtotal_refs = "+".join([f"I{r}" for r in cat_boq_subtotal_rows.values()]) or "0"
    ws_boq.merge_cells(start_row=boq_row, start_column=1, end_row=boq_row, end_column=8)
    gt_lbl = ws_boq.cell(boq_row, 1, "TOTAL BASE ESTIMATE COST (CARRIED TO GENERAL ABSTRACT)")
    gt_lbl.font = font_grand
    gt_lbl.alignment = align_right

    gt_val = ws_boq.cell(boq_row, 9, f"={subtotal_refs}")
    gt_val.font = font_grand
    gt_val.alignment = align_right
    gt_val.number_format = inr_format

    for ci in range(1, 10):
        ws_boq.cell(boq_row, ci).fill = grand_total_fill
        ws_boq.cell(boq_row, ci).border = total_border
    ws_boq.row_dimensions[boq_row].height = 26

    # =========================================================================
    # STEP 3: POPULATE SHEET 1 — "1. Abstract of Cost"
    # Executive Summary linking directly to Sheet 2 Category Sub-totals
    # =========================================================================
    write_project_banner(ws_abs, 6, "General Abstract of Cost (Executive Summary)")

    abs_headers = [
        ("Part / S.No.", 14),
        ("Work Trade / Category Head (IS 1200 / CPWD DSR)", 46),
        ("No. of BOQ Items", 18),
        ("Reference Sheet", 24),
        ("Estimated Amount (INR)", 24),
        ("% Share of Base Cost", 20),
    ]
    ws_abs.row_dimensions[4].height = 25
    for col_idx, (h_text, width) in enumerate(abs_headers, start=1):
        c = ws_abs.cell(4, col_idx, h_text)
        c.font = font_tbl_header
        c.fill = navy_header_fill
        c.alignment = align_center
        c.border = grid_border
        ws_abs.column_dimensions[get_column_letter(col_idx)].width = width

    abs_start_row = 5
    abs_row = abs_start_row
    total_cat_count = len(categories_order)
    base_total_row = abs_start_row + total_cat_count

    for cat_idx, cat in enumerate(categories_order, start=1):
        boq_sub_r = cat_boq_subtotal_rows[cat]
        num_items = len(groups_by_cat[cat])

        c1 = ws_abs.cell(abs_row, 1, f"Part {chr(64 + cat_idx)}")
        c1.font = font_bold
        c1.alignment = align_center

        c2 = ws_abs.cell(abs_row, 2, cat)
        c2.font = font_bold
        c2.alignment = align_left

        c3 = ws_abs.cell(abs_row, 3, num_items)
        c3.font = font_regular
        c3.alignment = align_center

        c4 = ws_abs.cell(abs_row, 4, f"BOQ Row #{boq_sub_r}")
        c4.font = font_muted
        c4.alignment = align_center

        c5 = ws_abs.cell(abs_row, 5, f"='2. Bill of Quantities (BOQ)'!I{boq_sub_r}")
        c5.font = font_bold
        c5.alignment = align_right
        c5.number_format = inr_format

        c6 = ws_abs.cell(abs_row, 6, f"=IF(E{base_total_row}>0, E{abs_row}/E{base_total_row}, 0)")
        c6.font = font_regular
        c6.alignment = align_right
        c6.number_format = "0.0%"

        for ci in range(1, 7):
            ws_abs.cell(abs_row, ci).border = grid_border
            if cat_idx % 2 == 0:
                ws_abs.cell(abs_row, ci).fill = zebra_fill
        ws_abs.row_dimensions[abs_row].height = 24
        abs_row += 1

    abs_end_cat_row = abs_row - 1

    # (A) Sub-Total Base Construction & Landscape Cost
    ws_abs.merge_cells(start_row=abs_row, start_column=1, end_row=abs_row, end_column=4)
    ws_abs.cell(abs_row, 1, "(A) SUB-TOTAL BASE WORK COST (EXCLUDING TAXES & CONTINGENCY)").font = font_bold
    ws_abs.cell(abs_row, 1).alignment = align_right
    c_base = ws_abs.cell(abs_row, 5, f"=SUM(E{abs_start_row}:E{abs_end_cat_row})" if total_cat_count > 0 else 0)
    c_base.font = font_bold
    c_base.alignment = align_right
    c_base.number_format = inr_format
    ws_abs.cell(abs_row, 6, 1.0).number_format = "0.0%"
    ws_abs.cell(abs_row, 6).font = font_bold
    ws_abs.cell(abs_row, 6).alignment = align_right
    for ci in range(1, 7):
        ws_abs.cell(abs_row, ci).fill = sub_total_fill
        ws_abs.cell(abs_row, ci).border = total_border
    ws_abs.row_dimensions[abs_row].height = 24
    row_a = abs_row
    abs_row += 1

    # (B) Physical & Unforeseen Contingencies
    ws_abs.merge_cells(start_row=abs_row, start_column=1, end_row=abs_row, end_column=4)
    ws_abs.cell(abs_row, 1, f"(B) Add: Physical & Site Contingencies / Sundries @ {contingency_pct*100:.1f}% on (A)").font = font_regular
    ws_abs.cell(abs_row, 1).alignment = align_right
    c_cont = ws_abs.cell(abs_row, 5, f"=ROUND(E{row_a}*{contingency_pct}, 2)")
    c_cont.font = font_regular
    c_cont.alignment = align_right
    c_cont.number_format = inr_format
    ws_abs.cell(abs_row, 6, contingency_pct).number_format = "0.0%"
    ws_abs.cell(abs_row, 6).alignment = align_right
    for ci in range(1, 7):
        ws_abs.cell(abs_row, ci).border = grid_border
    ws_abs.row_dimensions[abs_row].height = 22
    row_b = abs_row
    abs_row += 1

    # (C) Taxable Project Cost = A + B
    ws_abs.merge_cells(start_row=abs_row, start_column=1, end_row=abs_row, end_column=4)
    ws_abs.cell(abs_row, 1, "(C) TAXABLE PROJECT COST [(A) + (B)]").font = font_bold
    ws_abs.cell(abs_row, 1).alignment = align_right
    c_taxable = ws_abs.cell(abs_row, 5, f"=E{row_a}+E{row_b}")
    c_taxable.font = font_bold
    c_taxable.alignment = align_right
    c_taxable.number_format = inr_format
    for ci in range(1, 7):
        ws_abs.cell(abs_row, ci).fill = sub_total_fill
        ws_abs.cell(abs_row, ci).border = grid_border
    ws_abs.row_dimensions[abs_row].height = 22
    row_c = abs_row
    abs_row += 1

    # (D) Goods & Services Tax (GST)
    ws_abs.merge_cells(start_row=abs_row, start_column=1, end_row=abs_row, end_column=4)
    ws_abs.cell(abs_row, 1, f"(D) Add: Goods & Services Tax (CGST + SGST) @ {gst_pct*100:.1f}% on (C)").font = font_regular
    ws_abs.cell(abs_row, 1).alignment = align_right
    c_gst = ws_abs.cell(abs_row, 5, f"=ROUND(E{row_c}*{gst_pct}, 2)")
    c_gst.font = font_regular
    c_gst.alignment = align_right
    c_gst.number_format = inr_format
    ws_abs.cell(abs_row, 6, gst_pct).number_format = "0.0%"
    ws_abs.cell(abs_row, 6).alignment = align_right
    for ci in range(1, 7):
        ws_abs.cell(abs_row, ci).border = grid_border
    ws_abs.row_dimensions[abs_row].height = 22
    row_d = abs_row
    abs_row += 1

    # (E) GRAND TOTAL PROJECT COST = C + D
    ws_abs.merge_cells(start_row=abs_row, start_column=1, end_row=abs_row, end_column=4)
    ws_abs.cell(abs_row, 1, "(E) GRAND TOTAL ESTIMATED PROJECT COST [(C) + (D)] (INCL. GST)").font = font_grand
    ws_abs.cell(abs_row, 1).alignment = align_right
    c_grand = ws_abs.cell(abs_row, 5, f"=ROUND(E{row_c}+E{row_d}, 0)")
    c_grand.font = font_grand
    c_grand.alignment = align_right
    c_grand.number_format = inr_format
    for ci in range(1, 7):
        ws_abs.cell(abs_row, ci).fill = grand_total_fill
        ws_abs.cell(abs_row, ci).border = total_border
    ws_abs.row_dimensions[abs_row].height = 28
    abs_row += 3

    # Sign-off block
    ws_abs.cell(abs_row, 2, "Prepared By (QS / Billing Engineer):").font = font_bold
    ws_abs.cell(abs_row, 4, "Verified By (Project Manager):").font = font_bold
    ws_abs.cell(abs_row, 5, "Approved By (Client / Principal):").font = font_bold
    abs_row += 1
    ws_abs.cell(abs_row, 2, consultant_name).font = font_muted
    ws_abs.cell(abs_row, 4, "Checked per IS 1200 & 2D CAD Takeoff").font = font_muted
    ws_abs.cell(abs_row, 5, client_name).font = font_muted

    # =========================================================================
    # STEP 4: POPULATE SHEET 4 — "4. Bar Bending Schedule (BBS)"
    # IS 2502 / SP 34 Schedule with Live Cutting Length & Dia-wise Weight Formulas
    # =========================================================================
    bbs_rows_input = payload.get("bbs_rows")
    if bbs_rows_input is None:
        bbs_bundle = generate_bbs_from_takeoff(groups, project_type=project_type)
        bbs_rows_input = bbs_bundle["bbs_rows"]
    else:
        bbs_bundle = summarize_bbs(bbs_rows_input)

    write_project_banner(ws_bbs, 24, "Bar Bending Schedule (BBS — IS 2502 / SP 34 / Fe-500D TMT Reinforcement)")

    bbs_headers = [
        ("S.No", 7),
        ("Bar Mark", 11),
        ("Structural Member & Bar Description", 40),
        ("Shape Code", 12),
        ("Dia (mm)", 10),
        ("Spacing (mm)", 12),
        ("No. of Members", 14),
        ("Bars / Member", 14),
        ("Total Bars", 12),
        ("A (m)", 9),
        ("B (m)", 9),
        ("C (m)", 9),
        ("D (m)", 9),
        ("Bend Ded. (m)", 13),
        ("Cutting Length (m)", 16),
        ("Total Length (m)", 15),
        ("Unit Wt (kg/m)", 14),
        ("8mm (kg)", 12),
        ("10mm (kg)", 12),
        ("12mm (kg)", 12),
        ("16mm (kg)", 12),
        ("20mm (kg)", 12),
        ("25mm+ (kg)", 12),
        ("Total Wt (kg)", 15),
    ]
    ws_bbs.row_dimensions[4].height = 26
    for col_idx, (h_text, width) in enumerate(bbs_headers, start=1):
        c = ws_bbs.cell(4, col_idx, h_text)
        c.font = font_tbl_header
        c.fill = navy_header_fill
        c.alignment = align_center
        c.border = grid_border
        ws_bbs.column_dimensions[get_column_letter(col_idx)].width = width

    bbs_r = 5
    bbs_data_rows: List[int] = []
    current_member_grp = None

    for idx, brow in enumerate(bbs_rows_input, start=1):
        mgrp = brow.get("member_group", "Structural RCC Member")
        if mgrp != current_member_grp:
            current_member_grp = mgrp
            ws_bbs.merge_cells(start_row=bbs_r, start_column=1, end_row=bbs_r, end_column=24)
            mc = ws_bbs.cell(bbs_r, 1, f"MEMBER: {mgrp.upper()}")
            mc.font = font_cat
            mc.fill = cat_banner_fill
            mc.alignment = align_left
            for ci in range(1, 25):
                ws_bbs.cell(bbs_r, ci).border = grid_border
            ws_bbs.row_dimensions[bbs_r].height = 22
            bbs_r += 1

        bbs_data_rows.append(bbs_r)
        scode = brow.get("shape_code", "SC-00")

        ws_bbs.cell(bbs_r, 1, idx).alignment = align_center
        ws_bbs.cell(bbs_r, 2, brow.get("bar_mark", f"B-{idx}")).alignment = align_center
        ws_bbs.cell(bbs_r, 2).font = font_bold
        ws_bbs.cell(bbs_r, 3, brow.get("description", "")).alignment = align_left
        ws_bbs.cell(bbs_r, 4, scode).alignment = align_center
        ws_bbs.cell(bbs_r, 5, int(brow.get("dia_mm", 12))).alignment = align_center
        ws_bbs.cell(bbs_r, 5).font = font_bold
        ws_bbs.cell(bbs_r, 6, int(brow.get("spacing_mm", 150))).alignment = align_center
        ws_bbs.cell(bbs_r, 7, int(brow.get("num_members", 1))).alignment = align_center
        ws_bbs.cell(bbs_r, 8, int(brow.get("bars_per_member", 1))).alignment = align_center

        # Col I (9): Total Bars = G * H
        c_tbars = ws_bbs.cell(bbs_r, 9, f"=G{bbs_r}*H{bbs_r}")
        c_tbars.font = font_bold
        c_tbars.alignment = align_center

        # Cols J, K, L, M (10..13): Segments A, B, C, D (m)
        for offset, key in enumerate(("a_m", "b_m", "c_m", "d_m"), start=10):
            sc = ws_bbs.cell(bbs_r, offset, float(brow.get(key, 0.0)))
            sc.alignment = align_right
            sc.number_format = "0.000"

        # Col N (14): Bend Deduction (m)
        c_bded = ws_bbs.cell(bbs_r, 14, float(brow.get("bend_deduction_m", 0.0)))
        c_bded.alignment = align_right
        c_bded.number_format = "0.000"

        # Col O (15): Cutting Length (m) formula per Shape Code
        if scode == "SC-51":
            cut_formula = f"=ROUND(2*(J{bbs_r}+K{bbs_r})+L{bbs_r}+M{bbs_r}-N{bbs_r}, 3)"
        else:
            cut_formula = f"=ROUND(J{bbs_r}+K{bbs_r}+L{bbs_r}+M{bbs_r}-N{bbs_r}, 3)"
        c_cut = ws_bbs.cell(bbs_r, 15, cut_formula)
        c_cut.font = font_bold
        c_cut.alignment = align_right
        c_cut.number_format = "0.000"

        # Col P (16): Total Length (m) = I * O
        c_tlen = ws_bbs.cell(bbs_r, 16, f"=ROUND(I{bbs_r}*O{bbs_r}, 2)")
        c_tlen.font = font_bold
        c_tlen.alignment = align_right
        c_tlen.number_format = qty_format

        # Col Q (17): Unit Weight (kg/m) = ROUND(E^2 / 162.2, 3)
        c_uwt = ws_bbs.cell(bbs_r, 17, f"=ROUND((E{bbs_r}*E{bbs_r})/162.2, 3)")
        c_uwt.alignment = align_right
        c_uwt.number_format = "0.000"

        # Cols R..W (18..23): Diameter-wise Weight (kg)
        dia_cols = [(18, "=8"), (19, "=10"), (20, "=12"), (21, "=16"), (22, "=20"), (23, ">=25")]
        for col_i, cond in dia_cols:
            dc = ws_bbs.cell(bbs_r, col_i, f"=IF(E{bbs_r}{cond}, ROUND(P{bbs_r}*Q{bbs_r}, 2), 0)")
            dc.alignment = align_right
            dc.number_format = qty_format

        # Col X (24): Row Total Weight (kg) = SUM(R..W)
        c_rwt = ws_bbs.cell(bbs_r, 24, f"=SUM(R{bbs_r}:W{bbs_r})")
        c_rwt.font = font_bold
        c_rwt.alignment = align_right
        c_rwt.number_format = qty_format

        for ci in range(1, 25):
            ws_bbs.cell(bbs_r, ci).border = grid_border
            if idx % 2 == 0 and ws_bbs.cell(bbs_r, ci).fill.fill_type is None:
                ws_bbs.cell(bbs_r, ci).fill = zebra_fill
        ws_bbs.row_dimensions[bbs_r].height = 20
        bbs_r += 1

    # Summary Footer Rows in BBS Sheet
    first_bbs_r = bbs_data_rows[0] if bbs_data_rows else 5
    last_bbs_r = bbs_r - 1

    # 1. Net Steel Weight (kg)
    ws_bbs.merge_cells(start_row=bbs_r, start_column=1, end_row=bbs_r, end_column=17)
    ws_bbs.cell(bbs_r, 1, "NET REINFORCEMENT WEIGHT BY DIAMETER (KG)").font = font_bold
    ws_bbs.cell(bbs_r, 1).alignment = align_right
    for col_i in range(18, 25):
        col_let = get_column_letter(col_i)
        sc = ws_bbs.cell(bbs_r, col_i, f"=SUM({col_let}{first_bbs_r}:{col_let}{last_bbs_r})")
        sc.font = font_bold
        sc.alignment = align_right
        sc.number_format = qty_format
    for ci in range(1, 25):
        ws_bbs.cell(bbs_r, ci).fill = sub_total_fill
        ws_bbs.cell(bbs_r, ci).border = total_border
    ws_bbs.row_dimensions[bbs_r].height = 22
    net_kg_row = bbs_r
    bbs_r += 1

    # 2. Rolling Margin & Cutting Wastage @ 3% (kg)
    ws_bbs.merge_cells(start_row=bbs_r, start_column=1, end_row=bbs_r, end_column=17)
    ws_bbs.cell(bbs_r, 1, "Add: Rolling Margin, Lap & Cutting Wastage @ 3.0% (KG)").font = font_regular
    ws_bbs.cell(bbs_r, 1).alignment = align_right
    for col_i in range(18, 25):
        col_let = get_column_letter(col_i)
        sc = ws_bbs.cell(bbs_r, col_i, f"=ROUND({col_let}{net_kg_row}*0.03, 2)")
        sc.font = font_regular
        sc.alignment = align_right
        sc.number_format = qty_format
    for ci in range(1, 25):
        ws_bbs.cell(bbs_r, ci).border = grid_border
    ws_bbs.row_dimensions[bbs_r].height = 20
    wastage_row = bbs_r
    bbs_r += 1

    # 3. Gross Steel Weight in KG
    ws_bbs.merge_cells(start_row=bbs_r, start_column=1, end_row=bbs_r, end_column=17)
    ws_bbs.cell(bbs_r, 1, "GROSS TMT FE-500D STEEL REQUIREMENT (IN KILOGRAMS - KG)").font = font_bold
    ws_bbs.cell(bbs_r, 1).alignment = align_right
    for col_i in range(18, 25):
        col_let = get_column_letter(col_i)
        sc = ws_bbs.cell(bbs_r, col_i, f"={col_let}{net_kg_row}+{col_let}{wastage_row}")
        sc.font = font_bold
        sc.alignment = align_right
        sc.number_format = qty_format
    for ci in range(1, 25):
        ws_bbs.cell(bbs_r, ci).fill = sub_total_fill
        ws_bbs.cell(bbs_r, ci).border = total_border
    ws_bbs.row_dimensions[bbs_r].height = 22
    gross_kg_row = bbs_r
    bbs_r += 1

    # 4. Gross Steel Weight in Metric Tonnes (MT)
    ws_bbs.merge_cells(start_row=bbs_r, start_column=1, end_row=bbs_r, end_column=17)
    ws_bbs.cell(bbs_r, 1, "GROSS TMT FE-500D STEEL REQUIREMENT (IN METRIC TONNES - MT)").font = font_grand
    ws_bbs.cell(bbs_r, 1).alignment = align_right
    for col_i in range(18, 25):
        col_let = get_column_letter(col_i)
        sc = ws_bbs.cell(bbs_r, col_i, f"=ROUND({col_let}{gross_kg_row}/1000, 3)")
        sc.font = font_grand
        sc.alignment = align_right
        sc.number_format = "0.000"
    for ci in range(1, 25):
        ws_bbs.cell(bbs_r, ci).fill = grand_total_fill
        ws_bbs.cell(bbs_r, ci).border = total_border
    ws_bbs.row_dimensions[bbs_r].height = 25

    # Freeze header panes on all 4 sheets
    ws_abs.freeze_panes = "A5"
    ws_boq.freeze_panes = "A5"
    ws_meas.freeze_panes = "A5"
    ws_bbs.freeze_panes = "A5"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
