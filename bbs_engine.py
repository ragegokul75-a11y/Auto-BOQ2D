"""
IS 2502 / SP 34 / IS 456 Automatic Bar Bending Schedule (BBS) Engine.

Automatically generates complete Bar Bending Schedules from 2D CAD structural & landscape
elements (RCC Columns, Footings, Plinth/Floor Beams, Slabs, Retaining Walls, Water Body Walls)
OR from custom structural member schedules.

Formulas implemented per Indian Standard (IS 2502 / IS 456:2000 / SP 34):
- Unit Weight of TMT Bar (kg/m) = d^2 / 162.2  (where d = nominal diameter in mm)
- Bend Deductions (IS 2502):
    * 45° bend = 1 * d
    * 90° bend = 2 * d
    * 135° seismic hook bend = 3 * d
    * 180° hook bend = 4 * d
- Standard Anchorage / Hook Length = max(9 * d, 75mm) or 10 * d (for 135° seismic stirrup hook)
- Development / Lap Length (Ld) = 50 * d (Fe-500D TMT in M-25/M-30 concrete)
"""

import math
from typing import Dict, List, Any

STANDARD_DIAS = [8, 10, 12, 16, 20, 25, 32]

SHAPE_CODES = {
    "SC-00": {
        "code": "SC-00",
        "name": "Straight Bar",
        "formula_desc": "A",
        "bends_90": 0,
        "bends_45": 0,
        "bends_135": 0,
        "svg_path": "M 12 22 L 108 22",
    },
    "SC-01": {
        "code": "SC-01",
        "name": "L-Bent Bar (Single 90° Hook/Foot)",
        "formula_desc": "A + B - (1 × 2d)",
        "bends_90": 1,
        "bends_45": 0,
        "bends_135": 0,
        "svg_path": "M 16 8 L 16 28 L 104 28",
    },
    "SC-02": {
        "code": "SC-02",
        "name": "U-Bar / Both Ends 90° L-Hooked",
        "formula_desc": "A + B + C - (2 × 2d)",
        "bends_90": 2,
        "bends_45": 0,
        "bends_135": 0,
        "svg_path": "M 16 10 L 16 28 L 104 28 L 104 10",
    },
    "SC-51": {
        "code": "SC-51",
        "name": "Rectangular Closed Stirrup / Ring (135° Seismic Hooks)",
        "formula_desc": "2(A + B) + 2×Hook(10d) - (3×2d + 2×3d)",
        "bends_90": 3,
        "bends_45": 0,
        "bends_135": 2,
        "svg_path": "M 30 8 L 90 8 L 90 32 L 30 32 L 30 8 M 30 8 L 38 16 M 30 8 L 40 12",
    },
    "SC-41": {
        "code": "SC-41",
        "name": "Cranked / Bent-Up Slab Bar (Double 45° Crank)",
        "formula_desc": "A + 2(0.42×D) + 2×Hook - (2 × 1d)",
        "bends_90": 0,
        "bends_45": 2,
        "bends_135": 0,
        "svg_path": "M 8 12 L 28 12 L 40 28 L 80 28 L 92 12 L 112 12",
    },
    "SC-99": {
        "code": "SC-99",
        "name": "Chair Bar (Slab Top Mesh Support)",
        "formula_desc": "A + 2×B + 2×C - (4 × 2d)",
        "bends_90": 4,
        "bends_45": 0,
        "bends_135": 0,
        "svg_path": "M 12 30 L 36 30 L 36 12 L 84 12 L 84 30 L 108 30",
    },
}


def unit_weight_kg_per_m(dia_mm: int) -> float:
    """IS 1786 / IS 2502 unit weight: d^2 / 162.2 kg/m."""
    return round((dia_mm * dia_mm) / 162.2, 3)


def compute_bbs_row(row: Dict[str, Any]) -> Dict[str, Any]:
    """
    Computes exact cutting length (m), bend deduction (m), total bars, total length (m),
    unit weight (kg/m), total weight (kg), and diameter-wise breakdown for a BBS entry.
    All segment inputs (a_m, b_m, c_m, d_m, e_m) are in Meters (m).
    """
    dia = int(row.get("dia_mm", 12))
    shape_code = row.get("shape_code", "SC-00")
    shape_info = SHAPE_CODES.get(shape_code, SHAPE_CODES["SC-00"])

    num_members = int(row.get("num_members", 1))
    bars_per_member = int(row.get("bars_per_member", 1))
    total_bars = num_members * bars_per_member

    a_m = round(float(row.get("a_m", 0.0)), 3)
    b_m = round(float(row.get("b_m", 0.0)), 3)
    c_m = round(float(row.get("c_m", 0.0)), 3)
    d_m = round(float(row.get("d_m", 0.0)), 3)
    e_m = round(float(row.get("e_m", 0.0)), 3)

    d_meters = dia / 1000.0  # Convert mm to meters

    # Bend deduction per IS 2502: 90° = 2d, 45° = 1d, 135° = 3d
    b90 = int(row.get("bends_90", shape_info["bends_90"]))
    b45 = int(row.get("bends_45", shape_info["bends_45"]))
    b135 = int(row.get("bends_135", shape_info["bends_135"]))
    bend_deduction_m = round((b90 * 2 + b45 * 1 + b135 * 3) * d_meters, 3)

    # Gross length before bend deduction
    if shape_code == "SC-51":
        # Rectangular stirrup: 2*(A + B) + 2*Hooks (C + D)
        gross_len_m = 2.0 * (a_m + b_m) + c_m + d_m
    else:
        gross_len_m = a_m + b_m + c_m + d_m + e_m

    cutting_length_m = round(max(0.10, gross_len_m - bend_deduction_m), 3)
    total_length_m = round(total_bars * cutting_length_m, 2)
    u_wt = unit_weight_kg_per_m(dia)
    total_weight_kg = round(total_length_m * u_wt, 2)

    row.update({
        "dia_mm": dia,
        "shape_code": shape_code,
        "shape_name": shape_info["name"],
        "svg_path": shape_info["svg_path"],
        "num_members": num_members,
        "bars_per_member": bars_per_member,
        "total_bars": total_bars,
        "a_m": a_m,
        "b_m": b_m,
        "c_m": c_m,
        "d_m": d_m,
        "e_m": e_m,
        "bends_90": b90,
        "bends_45": b45,
        "bends_135": b135,
        "bend_deduction_m": bend_deduction_m,
        "cutting_length_m": cutting_length_m,
        "total_length_m": total_length_m,
        "unit_weight": u_wt,
        "total_weight_kg": total_weight_kg,
    })
    return row


def generate_bbs_from_takeoff(takeoff_groups: List[Dict[str, Any]], project_type: str = "Residential") -> Dict[str, Any]:
    """
    Automatically generates a detailed Bar Bending Schedule (BBS) per IS 2502 / SP 34
    by inspecting the 2D CAD takeoff groups (Columns, RCC Slabs, Walls/Beams, Water Bodies, Paver Kerb Haunching)
    and creating the corresponding Footing, Column, Plinth/Roof Beam, Slab, and Retaining Wall reinforcement bars.
    """
    bbs_rows: List[Dict[str, Any]] = []
    bar_id_seq = 0

    def next_id() -> str:
        nonlocal bar_id_seq
        bar_id_seq += 1
        return f"BBS-{bar_id_seq:03d}"

    # Inspect what structural / civil / landscape RCC elements exist in the 2D CAD drawing
    col_count = 0
    wall_runs: List[Dict[str, Any]] = []
    slab_zones: List[Dict[str, Any]] = []
    water_zones: List[Dict[str, Any]] = []

    for grp in takeoff_groups:
        if not grp.get("included", True):
            continue
        code = grp.get("dsr_code", "")
        lname = grp.get("layer", "").upper()
        disp = grp.get("display_name", "").upper()

        if code == "DSR-C04" or "COL" in lname or "COL" in disp:
            col_count += int(round(grp.get("total_qty", len(grp.get("measurements", [])))))
        elif code in ("DSR-C05", "DSR-C06") or "WALL" in lname:
            wall_runs.extend(grp.get("measurements", []))
        elif code == "DSR-C03" or "SLAB" in lname or "RCC" in lname:
            slab_zones.extend(grp.get("measurements", []))
        elif code == "DSR-L14" or "WATER" in lname or "POND" in disp:
            water_zones.extend(grp.get("measurements", []))
        elif code == "DSR-F01" and len(slab_zones) == 0:
            # In a 2D floor plan without a separate RCC slab layer, each room floor polygon is supported by an RCC slab panel!
            slab_zones.extend(grp.get("measurements", []))

    # Fallback defaults if a pure landscape drawing has no building columns/walls
    if col_count == 0 and len(wall_runs) == 0 and len(slab_zones) == 0 and len(water_zones) == 0:
        col_count = 8
        slab_zones = [{"label": "Gazebo / Pavilion RCC Deck Slab", "length": 8.0, "breadth": 5.0}]

    # -------------------------------------------------------------------------
    # 1. ISOLATED RCC FOOTINGS & COLUMNS (if columns exist or derived from grid)
    # -------------------------------------------------------------------------
    if col_count == 0 and len(wall_runs) > 0:
        col_count = max(6, len(wall_runs) + 2)

    if col_count > 0:
        # Footing Bottom Mesh (1.8m x 1.8m isolated footing, 50mm cover, 12mm @ 150 c/c both ways)
        foot_span = 1.80 - 2 * 0.05  # 1.70m
        foot_leg = 0.30              # 300mm L-upstand at both ends
        bars_one_way = int(math.ceil(foot_span / 0.15)) + 1  # 13 bars
        bbs_rows.append(compute_bbs_row({
            "id": next_id(),
            "member_group": f"1. Isolated RCC Footings F1 (1800×1800×450 mm) — {col_count} Nos",
            "bar_mark": "F1-X",
            "description": "Footing Bottom Mesh Along X-Axis (12mm Ø @ 150mm c/c, Cover 50mm)",
            "dia_mm": 12,
            "spacing_mm": 150,
            "shape_code": "SC-02",
            "num_members": col_count,
            "bars_per_member": bars_one_way,
            "a_m": foot_leg,
            "b_m": foot_span,
            "c_m": foot_leg,
            "d_m": 0.0,
            "e_m": 0.0,
        }))
        bbs_rows.append(compute_bbs_row({
            "id": next_id(),
            "member_group": f"1. Isolated RCC Footings F1 (1800×1800×450 mm) — {col_count} Nos",
            "bar_mark": "F1-Y",
            "description": "Footing Bottom Mesh Along Y-Axis (12mm Ø @ 150mm c/c, Cover 50mm)",
            "dia_mm": 12,
            "spacing_mm": 150,
            "shape_code": "SC-02",
            "num_members": col_count,
            "bars_per_member": bars_one_way,
            "a_m": foot_leg,
            "b_m": foot_span,
            "c_m": foot_leg,
            "d_m": 0.0,
            "e_m": 0.0,
        }))

        # RCC Columns C1 (300x450mm, Footing depth 1.5m + Floor height 3.2m + Lap 50d)
        # 4 Nos 20mm Corner Main Bars + 4 Nos 16mm Intermediate Bars
        col_vert_len = 1.40 + 3.20 + (50 * 0.020)  # 5.60m
        bbs_rows.append(compute_bbs_row({
            "id": next_id(),
            "member_group": f"2. RCC Structural Columns C1 (300×450 mm) — {col_count} Nos",
            "bar_mark": "C1-M1",
            "description": "Column Corner Longitudinal Main Bars (4 Nos - 20mm Ø with 450mm Footing L-Bend + 50d Lap)",
            "dia_mm": 20,
            "spacing_mm": 0,
            "shape_code": "SC-01",
            "num_members": col_count,
            "bars_per_member": 4,
            "a_m": 0.45,
            "b_m": round(col_vert_len, 2),
            "c_m": 0.0,
            "d_m": 0.0,
            "e_m": 0.0,
        }))
        bbs_rows.append(compute_bbs_row({
            "id": next_id(),
            "member_group": f"2. RCC Structural Columns C1 (300×450 mm) — {col_count} Nos",
            "bar_mark": "C1-M2",
            "description": "Column Mid-Face Longitudinal Bars (4 Nos - 16mm Ø with 450mm Footing L-Bend + 50d Lap)",
            "dia_mm": 16,
            "spacing_mm": 0,
            "shape_code": "SC-01",
            "num_members": col_count,
            "bars_per_member": 4,
            "a_m": 0.45,
            "b_m": round(1.40 + 3.20 + (50 * 0.016), 2),
            "c_m": 0.0,
            "d_m": 0.0,
            "e_m": 0.0,
        }))

        # Column Lateral Ties / Stirrups (8mm @ 150mm c/c, core 220x370mm, 135° seismic hooks 10d = 80mm)
        num_ties = int(math.ceil(4.50 / 0.15)) + 1  # 31 stirrups per column
        bbs_rows.append(compute_bbs_row({
            "id": next_id(),
            "member_group": f"2. RCC Structural Columns C1 (300×450 mm) — {col_count} Nos",
            "bar_mark": "C1-S1",
            "description": "Column Master Rectangular Lateral Ties (8mm Ø @ 150mm c/c, 40mm Cover, 135° Seismic Hooks)",
            "dia_mm": 8,
            "spacing_mm": 150,
            "shape_code": "SC-51",
            "num_members": col_count,
            "bars_per_member": num_ties,
            "a_m": 0.22,
            "b_m": 0.37,
            "c_m": 0.08,
            "d_m": 0.08,
            "e_m": 0.0,
        }))

    # -------------------------------------------------------------------------
    # 2. PLINTH BEAMS & FLOOR ROOF BEAMS (Derived directly from 2D Wall Runs)
    # -------------------------------------------------------------------------
    if len(wall_runs) > 0:
        total_wall_len = sum(float(w.get("length", 4.0)) for w in wall_runs)
        # Group into typical 4.5m beam spans or summarize across all extracted wall spans
        avg_beam_span = 4.50
        num_beams = max(4, int(round(total_wall_len / avg_beam_span)))

        # Beam Bottom Main Bars (3 Nos 16mm straight/L-anchored)
        bbs_rows.append(compute_bbs_row({
            "id": next_id(),
            "member_group": f"3. RCC Plinth & Roof Floor Beams PB/FB (230×450 mm) — {num_beams} Spans ({total_wall_len:.1f}m Total Run)",
            "bar_mark": "BM-B1",
            "description": "Beam Bottom Tension Main Bars (3 Nos - 16mm Ø Through Bars + 300mm End Development)",
            "dia_mm": 16,
            "spacing_mm": 0,
            "shape_code": "SC-02",
            "num_members": num_beams,
            "bars_per_member": 3,
            "a_m": 0.30,
            "b_m": avg_beam_span,
            "c_m": 0.30,
            "d_m": 0.0,
            "e_m": 0.0,
        }))
        # Beam Top Anchor Bars (2 Nos 12mm) + Extra Top Over Supports (2 Nos 16mm)
        bbs_rows.append(compute_bbs_row({
            "id": next_id(),
            "member_group": f"3. RCC Plinth & Roof Floor Beams PB/FB (230×450 mm) — {num_beams} Spans ({total_wall_len:.1f}m Total Run)",
            "bar_mark": "BM-T1",
            "description": "Beam Top Hanger / Anchor Bars (2 Nos - 12mm Ø Continuous + 250mm L-Anchors)",
            "dia_mm": 12,
            "spacing_mm": 0,
            "shape_code": "SC-02",
            "num_members": num_beams,
            "bars_per_member": 2,
            "a_m": 0.25,
            "b_m": avg_beam_span,
            "c_m": 0.25,
            "d_m": 0.0,
            "e_m": 0.0,
        }))
        bbs_rows.append(compute_bbs_row({
            "id": next_id(),
            "member_group": f"3. RCC Plinth & Roof Floor Beams PB/FB (230×450 mm) — {num_beams} Spans ({total_wall_len:.1f}m Total Run)",
            "bar_mark": "BM-T2",
            "description": "Beam Top Support Extra Negative Bars up to L/3 (2 Nos - 16mm Ø @ Each Support)",
            "dia_mm": 16,
            "spacing_mm": 0,
            "shape_code": "SC-01",
            "num_members": num_beams,
            "bars_per_member": 4,
            "a_m": 0.30,
            "b_m": round(avg_beam_span / 3.0, 2),
            "c_m": 0.0,
            "d_m": 0.0,
            "e_m": 0.0,
        }))
        # Beam Shear Stirrups (8mm 2-legged stirrups @ 150mm c/c, core 170x390mm)
        beam_stirrups = int(math.ceil(avg_beam_span / 0.15)) + 1
        bbs_rows.append(compute_bbs_row({
            "id": next_id(),
            "member_group": f"3. RCC Plinth & Roof Floor Beams PB/FB (230×450 mm) — {num_beams} Spans ({total_wall_len:.1f}m Total Run)",
            "bar_mark": "BM-S1",
            "description": "Beam 2-Legged Shear Stirrups (8mm Ø @ 150mm c/c, 30mm Cover, 135° Hooks)",
            "dia_mm": 8,
            "spacing_mm": 150,
            "shape_code": "SC-51",
            "num_members": num_beams,
            "bars_per_member": beam_stirrups,
            "a_m": 0.17,
            "b_m": 0.39,
            "c_m": 0.08,
            "d_m": 0.08,
            "e_m": 0.0,
        }))

    # -------------------------------------------------------------------------
    # 3. TWO-WAY RCC FLOOR / ROOF SLAB PANELS (Derived from 2D Room / Slab Polygons)
    # -------------------------------------------------------------------------
    for idx, sz in enumerate(slab_zones[:6], start=1):
        lx = max(2.0, float(sz.get("length", 5.0)))
        ly = max(2.0, float(sz.get("breadth", 4.0)))
        short_span = round(min(lx, ly), 2)
        long_span = round(max(lx, ly), 2)
        slab_lbl = sz.get("label", f"Slab Panel S{idx}")

        # Main Cranked / Bent-up Bars along Shorter Span (10mm @ 150mm c/c)
        # Crank height D = 150mm - 2*20mm(cover) - 10mm = 100mm -> extra length = 2 * 0.42 * 0.10 = 0.084m
        num_short_bars = int(math.ceil(long_span / 0.15)) + 1
        bbs_rows.append(compute_bbs_row({
            "id": next_id(),
            "member_group": f"4.{idx} RCC Two-Way Slab Panel — {slab_lbl} ({short_span:.2f}m × {long_span:.2f}m, 150mm Thk)",
            "bar_mark": f"S{idx}-M1",
            "description": f"Main Cranked Bars Along Short Span ({short_span}m) — 10mm Ø @ 150mm c/c (Alternate Bent-Up)",
            "dia_mm": 10,
            "spacing_mm": 150,
            "shape_code": "SC-41",
            "num_members": 1,
            "bars_per_member": num_short_bars,
            "a_m": round(short_span - 0.04, 2),
            "b_m": 0.084,  # 2 * 0.42 * D
            "c_m": 0.10,   # Left anchorage
            "d_m": 0.10,   # Right anchorage
            "e_m": 0.0,
        }))

        # Distribution Straight Bars along Longer Span (8mm @ 150mm c/c)
        num_long_bars = int(math.ceil(short_span / 0.15)) + 1
        bbs_rows.append(compute_bbs_row({
            "id": next_id(),
            "member_group": f"4.{idx} RCC Two-Way Slab Panel — {slab_lbl} ({short_span:.2f}m × {long_span:.2f}m, 150mm Thk)",
            "bar_mark": f"S{idx}-D1",
            "description": f"Distribution Straight Bars Along Long Span ({long_span}m) — 8mm Ø @ 150mm c/c",
            "dia_mm": 8,
            "spacing_mm": 150,
            "shape_code": "SC-00",
            "num_members": 1,
            "bars_per_member": num_long_bars,
            "a_m": round(long_span - 0.04 + 0.20, 2),
            "b_m": 0.0,
            "c_m": 0.0,
            "d_m": 0.0,
            "e_m": 0.0,
        }))

        # Chair Bars (12mm Ø @ 1 per Sq.m for top mesh support)
        chair_count = max(4, int(round((short_span * long_span) / 1.0)))
        bbs_rows.append(compute_bbs_row({
            "id": next_id(),
            "member_group": f"4.{idx} RCC Two-Way Slab Panel — {slab_lbl} ({short_span:.2f}m × {long_span:.2f}m, 150mm Thk)",
            "bar_mark": f"S{idx}-CH",
            "description": "Slab Chair Bars for Top Mesh Support (12mm Ø @ 1.0m × 1.0m Grid)",
            "dia_mm": 12,
            "spacing_mm": 1000,
            "shape_code": "SC-99",
            "num_members": 1,
            "bars_per_member": chair_count,
            "a_m": 0.45,   # Head width
            "b_m": 0.20,   # 2 x 0.10m vertical height
            "c_m": 0.60,   # 2 x 0.30m bottom legs
            "d_m": 0.0,
            "e_m": 0.0,
        }))

    # -------------------------------------------------------------------------
    # 4. LANDSCAPE WATER BODY / RETAINING WALL RCC RAFT & STEM (If Water Body exists)
    # -------------------------------------------------------------------------
    for idx, wz in enumerate(water_zones, start=1):
        lx = max(3.0, float(wz.get("length", 7.0)))
        ly = max(2.5, float(wz.get("breadth", 5.0)))
        perim = 2.0 * (lx + ly)
        w_lbl = wz.get("label", f"Reflective Water Body #{idx}")

        # Base Raft Double Mesh (12mm @ 150mm c/c Top & Bottom)
        bars_x = (int(math.ceil(ly / 0.15)) + 1) * 2  # Top + Bottom layers
        bars_y = (int(math.ceil(lx / 0.15)) + 1) * 2
        bbs_rows.append(compute_bbs_row({
            "id": next_id(),
            "member_group": f"5.{idx} Landscape Water Body / Retaining Pool RCC Structure — {w_lbl} ({lx:.1f}m × {ly:.1f}m)",
            "bar_mark": f"WB{idx}-RX",
            "description": "Water Body RCC Base Raft Double Mesh (Top & Bottom) Along X — 12mm Ø @ 150mm c/c",
            "dia_mm": 12,
            "spacing_mm": 150,
            "shape_code": "SC-02",
            "num_members": 1,
            "bars_per_member": bars_x,
            "a_m": 0.25,
            "b_m": round(lx - 0.10, 2),
            "c_m": 0.25,
            "d_m": 0.0,
            "e_m": 0.0,
        }))
        bbs_rows.append(compute_bbs_row({
            "id": next_id(),
            "member_group": f"5.{idx} Landscape Water Body / Retaining Pool RCC Structure — {w_lbl} ({lx:.1f}m × {ly:.1f}m)",
            "bar_mark": f"WB{idx}-RY",
            "description": "Water Body RCC Base Raft Double Mesh (Top & Bottom) Along Y — 12mm Ø @ 150mm c/c",
            "dia_mm": 12,
            "spacing_mm": 150,
            "shape_code": "SC-02",
            "num_members": 1,
            "bars_per_member": bars_y,
            "a_m": 0.25,
            "b_m": round(ly - 0.10, 2),
            "c_m": 0.25,
            "d_m": 0.0,
            "e_m": 0.0,
        }))

        # Perimeter Retaining Wall L-Starter & Stem Vertical Bars (10mm @ 150mm c/c inner + outer faces)
        vert_bars = (int(math.ceil(perim / 0.15)) + 1) * 2
        bbs_rows.append(compute_bbs_row({
            "id": next_id(),
            "member_group": f"5.{idx} Landscape Water Body / Retaining Pool RCC Structure — {w_lbl} ({lx:.1f}m × {ly:.1f}m)",
            "bar_mark": f"WB{idx}-WV",
            "description": "Water Body Perimeter Retaining Wall Vertical L-Bars (Inner & Outer Faces, 1.2m High) — 10mm Ø @ 150mm c/c",
            "dia_mm": 10,
            "spacing_mm": 150,
            "shape_code": "SC-01",
            "num_members": 1,
            "bars_per_member": vert_bars,
            "a_m": 0.45,
            "b_m": 1.15,
            "c_m": 0.0,
            "d_m": 0.0,
            "e_m": 0.0,
        }))

    return summarize_bbs(bbs_rows)


def summarize_bbs(bbs_rows: List[Dict[str, Any]], wastage_pct: float = 3.0, steel_rate_per_kg: float = 88.0) -> Dict[str, Any]:
    """
    Recomputes every BBS row and builds the Diameter-wise Weight Abstract (8mm to 32mm)
    in Kilograms (kg) and Metric Tonnes (MT), including rolling margin/wastage and cost.
    """
    dia_weights_kg: Dict[int, float] = {d: 0.0 for d in STANDARD_DIAS}
    dia_lengths_m: Dict[int, float] = {d: 0.0 for d in STANDARD_DIAS}

    for r in bbs_rows:
        compute_bbs_row(r)
        d = int(r["dia_mm"])
        if d not in dia_weights_kg:
            dia_weights_kg[d] = 0.0
            dia_lengths_m[d] = 0.0
        dia_weights_kg[d] = round(dia_weights_kg[d] + r["total_weight_kg"], 2)
        dia_lengths_m[d] = round(dia_lengths_m[d] + r["total_length_m"], 2)

    net_weight_kg = round(sum(dia_weights_kg.values()), 2)
    wastage_kg = round(net_weight_kg * (wastage_pct / 100.0), 2)
    gross_weight_kg = round(net_weight_kg + wastage_kg, 2)
    gross_weight_mt = round(gross_weight_kg / 1000.0, 3)
    estimated_steel_cost = round(gross_weight_kg * steel_rate_per_kg, 2)

    return {
        "bbs_rows": bbs_rows,
        "shape_codes": list(SHAPE_CODES.values()),
        "summary": {
            "dia_weights_kg": {str(k): v for k, v in dia_weights_kg.items()},
            "dia_lengths_m": {str(k): v for k, v in dia_lengths_m.items()},
            "net_weight_kg": net_weight_kg,
            "wastage_pct": wastage_pct,
            "wastage_kg": wastage_kg,
            "gross_weight_kg": gross_weight_kg,
            "gross_weight_mt": gross_weight_mt,
            "steel_rate_per_kg": steel_rate_per_kg,
            "estimated_steel_cost": estimated_steel_cost,
        },
    }
