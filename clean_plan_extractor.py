"""
Precision 2D Drawing Renderer & Clean Geometry Extraction Engine (.PDF, .DXF, .DWG).
Eliminates messy auto-guesswork by:
1. Rendering a high-resolution crisp image of PDF pages (or crisp vector geometry for CAD)
   along with exact vertex snap points, closed polygons, line segments, and text callouts.
2. Providing exact 1-Click Room/Zone Fill, 2-Click Scale Calibration, Wall Running-Length Tracer,
   Count Stamp, and Text-Dimension (`4.50 x 3.60`) Click-to-Add.
"""

import base64
import math
import os
import re
import tempfile
from typing import Dict, List, Any, Tuple, Optional

import fitz  # PyMuPDF
import ezdxf

from dxf_takeoff_engine import (
    polygon_area_and_perimeter,
    polyline_length,
    bounding_dims,
    point_in_polygon,
    parse_inline_dimension_text,
    match_text_callout,
)
from pdf_dwg_converter import find_dwg_converter_exe


def extract_clean_dxf_plan(file_path: str, unit_scale: float = 1.0) -> Dict[str, Any]:
    """
    Extracts clean, exact 2D CAD lines, closed polygons, circles/blocks, snap vertices,
    and text labels WITHOUT lumping random lines into wrong BOQ items.
    """
    doc = ezdxf.readfile(file_path)
    msp = doc.modelspace()

    insunits = doc.header.get("$INSUNITS", 0)
    if unit_scale <= 0:
        if insunits == 4:
            unit_scale = 0.001
        elif insunits == 5:
            unit_scale = 0.01
        elif insunits == 1:
            unit_scale = 0.0254
        elif insunits == 2:
            unit_scale = 0.3048
        else:
            unit_scale = 1.0

    lines_and_polys: List[Dict[str, Any]] = []
    closed_regions: List[Dict[str, Any]] = []
    blocks_and_circles: List[Dict[str, Any]] = []
    texts: List[Dict[str, Any]] = []
    snap_points: List[Tuple[float, float]] = []

    min_x, min_y = float("inf"), float("inf")
    max_x, max_y = float("-inf"), float("-inf")

    def add_pt(x: float, y: float):
        nonlocal min_x, min_y, max_x, max_y
        if x < min_x:
            min_x = x
        if y < min_y:
            min_y = y
        if x > max_x:
            max_x = x
        if y > max_y:
            max_y = y
        if len(snap_points) < 4000:
            snap_points.append((round(x, 3), round(y, 3)))

    seq = 0
    for ent in msp:
        etype = ent.dxftype()
        layer = ent.dxf.layer or "0"
        seq += 1
        eid = f"CAD-{seq:04d}"

        if etype in ("TEXT", "MTEXT"):
            try:
                txt = (ent.plain_text() if etype == "MTEXT" else ent.dxf.text).strip()
                if not txt:
                    continue
                ins = ent.dxf.insert
                x, y = round(float(ins.x) * unit_scale, 3), round(float(ins.y) * unit_scale, 3)
                add_pt(x, y)
                dims = parse_inline_dimension_text(txt)
                callout = match_text_callout(txt)
                texts.append({
                    "id": eid,
                    "text": txt,
                    "x": x,
                    "y": y,
                    "layer": layer,
                    "dims": list(dims) if dims else None,
                    "suggested_dsr": callout[0] if callout else None,
                    "suggested_label": callout[1] if callout else None,
                })
            except Exception:
                pass

        elif etype in ("LWPOLYLINE", "POLYLINE"):
            pts: List[Tuple[float, float]] = []
            is_closed = False
            if etype == "LWPOLYLINE":
                is_closed = bool(ent.closed)
                for p in ent.get_points("xy"):
                    px, py = round(float(p[0]) * unit_scale, 3), round(float(p[1]) * unit_scale, 3)
                    pts.append((px, py))
                    add_pt(px, py)
            else:
                is_closed = bool(ent.is_closed)
                for v in ent.vertices:
                    loc = v.dxf.location
                    px, py = round(float(loc.x) * unit_scale, 3), round(float(loc.y) * unit_scale, 3)
                    pts.append((px, py))
                    add_pt(px, py)

            if len(pts) < 2:
                continue
            if len(pts) >= 3 and math.hypot(pts[0][0] - pts[-1][0], pts[0][1] - pts[-1][1]) < 1e-3:
                is_closed = True
                pts = pts[:-1]

            if is_closed and len(pts) >= 3:
                area, perim = polygon_area_and_perimeter(pts)
                b_len, b_wid, centroid = bounding_dims(pts)
                closed_regions.append({
                    "id": eid,
                    "layer": layer,
                    "points": pts,
                    "area": round(area, 3),
                    "perimeter": round(perim, 3),
                    "length": round(b_len, 3),
                    "breadth": round(area / b_len, 3) if b_len > 0 else 1.0,
                    "centroid": list(centroid),
                })
            else:
                length = polyline_length(pts, closed=False)
                _, _, centroid = bounding_dims(pts)
                lines_and_polys.append({
                    "id": eid,
                    "layer": layer,
                    "points": pts,
                    "length": round(length, 3),
                    "centroid": list(centroid),
                })

        elif etype == "HATCH":
            pts = []
            try:
                for path in ent.paths:
                    if hasattr(path, "vertices") and len(path.vertices) >= 3:
                        for v in path.vertices:
                            px, py = round(float(v[0]) * unit_scale, 3), round(float(v[1]) * unit_scale, 3)
                            pts.append((px, py))
                            add_pt(px, py)
                        break
            except Exception:
                pass
            if len(pts) >= 3:
                area, perim = polygon_area_and_perimeter(pts)
                b_len, _, centroid = bounding_dims(pts)
                closed_regions.append({
                    "id": eid,
                    "layer": layer,
                    "points": pts,
                    "area": round(area, 3),
                    "perimeter": round(perim, 3),
                    "length": round(b_len, 3),
                    "breadth": round(area / b_len, 3) if b_len > 0 else 1.0,
                    "centroid": list(centroid),
                })

        elif etype == "LINE":
            s, e = ent.dxf.start, ent.dxf.end
            p1 = (round(float(s.x) * unit_scale, 3), round(float(s.y) * unit_scale, 3))
            p2 = (round(float(e.x) * unit_scale, 3), round(float(e.y) * unit_scale, 3))
            add_pt(*p1)
            add_pt(*p2)
            length = math.hypot(p2[0] - p1[0], p2[1] - p1[1])
            if length > 1e-3:
                lines_and_polys.append({
                    "id": eid,
                    "layer": layer,
                    "points": [p1, p2],
                    "length": round(length, 3),
                    "centroid": [round((p1[0] + p2[0]) / 2, 3), round((p1[1] + p2[1]) / 2, 3)],
                })

        elif etype == "INSERT":
            bname = ent.dxf.name
            if not bname.startswith("*"):
                ins = ent.dxf.insert
                x, y = round(float(ins.x) * unit_scale, 3), round(float(ins.y) * unit_scale, 3)
                add_pt(x, y)
                blocks_and_circles.append({
                    "id": eid,
                    "layer": layer,
                    "name": bname,
                    "x": x,
                    "y": y,
                    "radius": 0.45,
                })

        elif etype == "CIRCLE":
            c = ent.dxf.center
            r = round(float(ent.dxf.radius) * unit_scale, 3)
            cx, cy = round(float(c.x) * unit_scale, 3), round(float(c.y) * unit_scale, 3)
            add_pt(cx, cy)
            blocks_and_circles.append({
                "id": eid,
                "layer": layer,
                "name": f"Circle Ø{2*r:.2f}m",
                "x": cx,
                "y": cy,
                "radius": r,
            })

    # Attach room text labels to closed regions if a text sits inside
    for reg in closed_regions:
        for t in texts:
            if point_in_polygon(t["x"], t["y"], reg["points"]):
                reg["label"] = t["text"]
                reg["suggested_dsr"] = t["suggested_dsr"]
                if t["dims"]:
                    reg["length"] = t["dims"][0]
                    reg["breadth"] = t["dims"][1]
                    reg["area"] = round(t["dims"][0] * t["dims"][1], 3)
                break

    if min_x == float("inf"):
        min_x, min_y, max_x, max_y = 0.0, 0.0, 40.0, 30.0

    return {
        "drawing_format": "DXF",
        "background_image_b64": None,
        "bounds": {
            "min_x": round(min_x, 2),
            "min_y": round(min_y, 2),
            "max_x": round(max_x, 2),
            "max_y": round(max_y, 2),
            "width": round(max(1.0, max_x - min_x), 2),
            "height": round(max(1.0, max_y - min_y), 2),
        },
        "unit_scale": unit_scale,
        "closed_regions": closed_regions,
        "lines_and_polys": lines_and_polys,
        "blocks_and_circles": blocks_and_circles,
        "texts": texts,
        "snap_points": snap_points,
    }


def extract_clean_pdf_plan(pdf_path: str, scale_m_per_pt: float = 0.0352778) -> Dict[str, Any]:
    """
    Renders the actual PDF page as a crisp high-DPI PNG image (so the user sees their exact
    true PDF drawing on screen without any distortion!) AND extracts all vector polygons, lines,
    snap points, and text labels on top for 1-click measurement.
    """
    doc = fitz.open(pdf_path)
    if len(doc) == 0:
        raise ValueError("PDF has no pages.")

    page = doc[0]
    rect = page.rect
    pw, ph = float(rect.width), float(rect.height)

    # Render crisp 2x DPI PNG image of the PDF page
    pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False)
    png_bytes = pix.tobytes("png")
    b64_img = "data:image/png;base64," + base64.b64encode(png_bytes).decode("ascii")

    eff_scale = scale_m_per_pt if scale_m_per_pt > 0 else 0.0352778
    world_w = round(pw * eff_scale, 2)
    world_h = round(ph * eff_scale, 2)

    closed_regions: List[Dict[str, Any]] = []
    lines_and_polys: List[Dict[str, Any]] = []
    texts: List[Dict[str, Any]] = []
    snap_points: List[Tuple[float, float]] = []

    # 1. Extract text callouts
    t_seq = 0
    for b in page.get_text("blocks"):
        if len(b) < 5:
            continue
        txt = str(b[4]).strip().replace("\n", " ")
        if not txt:
            continue
        cx = round(((float(b[0]) + float(b[2])) / 2.0) * eff_scale, 3)
        cy = round((ph - ((float(b[1]) + float(b[3])) / 2.0)) * eff_scale, 3)
        t_seq += 1
        dims = parse_inline_dimension_text(txt)
        callout = match_text_callout(txt)
        texts.append({
            "id": f"PDF-T-{t_seq:04d}",
            "text": txt[:60],
            "x": cx,
            "y": cy,
            "layer": "PDF",
            "dims": list(dims) if dims else None,
            "suggested_dsr": callout[0] if callout else None,
            "suggested_label": callout[1] if callout else None,
        })

    # 2. Extract vector regions & lines for snapping and 1-click room selection
    seq = 0
    for path in page.get_drawings():
        items = path.get("items", [])
        if not items:
            continue
        pts: List[Tuple[float, float]] = []
        is_closed = bool(path.get("closePath", False)) or (path.get("fill") is not None)

        for item in items:
            op = item[0]
            if op == "re":
                r_box = item[1]
                if abs(r_box.width - pw) < 12 and abs(r_box.height - ph) < 12:
                    continue
                x0 = round(float(r_box.x0) * eff_scale, 3)
                y0 = round((ph - float(r_box.y1)) * eff_scale, 3)
                x1 = round(float(r_box.x1) * eff_scale, 3)
                y1 = round((ph - float(r_box.y0)) * eff_scale, 3)
                pts = [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]
                is_closed = True
                break
            elif op == "l":
                p1, p2 = item[1], item[2]
                pt1 = (round(float(p1.x) * eff_scale, 3), round((ph - float(p1.y)) * eff_scale, 3))
                pt2 = (round(float(p2.x) * eff_scale, 3), round((ph - float(p2.y)) * eff_scale, 3))
                if not pts:
                    pts.append(pt1)
                pts.append(pt2)

        for px, py in pts:
            if len(snap_points) < 4000:
                snap_points.append((px, py))

        if len(pts) < 2:
            continue
        seq += 1
        eid = f"PDF-G-{seq:04d}"

        if is_closed and len(pts) >= 3:
            area, perim = polygon_area_and_perimeter(pts)
            if area < 0.05:
                continue
            b_len, _, centroid = bounding_dims(pts)
            closed_regions.append({
                "id": eid,
                "layer": "PDF",
                "points": pts,
                "area": round(area, 3),
                "perimeter": round(perim, 3),
                "length": round(b_len, 3),
                "breadth": round(area / b_len, 3) if b_len > 0 else 1.0,
                "centroid": list(centroid),
            })
        else:
            length = polyline_length(pts, closed=False)
            if length > 0.1:
                _, _, centroid = bounding_dims(pts)
                lines_and_polys.append({
                    "id": eid,
                    "layer": "PDF",
                    "points": pts,
                    "length": round(length, 3),
                    "centroid": list(centroid),
                })

    for reg in closed_regions:
        for t in texts:
            if point_in_polygon(t["x"], t["y"], reg["points"]):
                reg["label"] = t["text"]
                reg["suggested_dsr"] = t["suggested_dsr"]
                if t["dims"]:
                    reg["length"] = t["dims"][0]
                    reg["breadth"] = t["dims"][1]
                    reg["area"] = round(t["dims"][0] * t["dims"][1], 3)
                break

    doc.close()
    return {
        "drawing_format": "PDF",
        "background_image_b64": b64_img,
        "bounds": {
            "min_x": 0.0,
            "min_y": 0.0,
            "max_x": world_w,
            "max_y": world_h,
            "width": world_w,
            "height": world_h,
        },
        "unit_scale": eff_scale,
        "closed_regions": closed_regions,
        "lines_and_polys": lines_and_polys,
        "blocks_and_circles": [],
        "texts": texts,
        "snap_points": snap_points,
    }
