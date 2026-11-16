"""
Command-Line / Batch Folder Processor for 2D AutoCAD (.DXF) Takeoff, BOQ & Abstract Generation.
Usage:
  python batch_boq_cli.py                       # Generates Excel BOQ workbooks for all .dxf files in samples/ into exports/
  python batch_boq_cli.py path/to/drawing.dxf   # Generates Excel BOQ workbook for a specific .dxf drawing
"""

import os
import sys
from sample_dxf_generator import ensure_all_samples
from dxf_takeoff_engine import parse_dxf_takeoff
from pdf_dwg_converter import parse_pdf_takeoff, parse_dwg_takeoff
from excel_boq_exporter import generate_boq_excel

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXPORTS_DIR = os.path.join(BASE_DIR, "exports")


def process_single_drawing(file_path: str, project_type: str = "Landscape", unit_scale: float = 1.0) -> str:
    os.makedirs(EXPORTS_DIR, exist_ok=True)
    basename = os.path.splitext(os.path.basename(file_path))[0]
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".pdf":
        parsed = parse_pdf_takeoff(file_path, unit_scale=unit_scale, project_type=project_type)
    elif ext == ".dwg":
        parsed = parse_dwg_takeoff(file_path, original_filename=os.path.basename(file_path), unit_scale=unit_scale, project_type=project_type)
    else:
        parsed = parse_dxf_takeoff(file_path, unit_scale=unit_scale, project_type=project_type)

    payload = {
        "project_name": basename.replace("_", " ").title(),
        "client_name": "M/s Principal Project Developer",
        "drawing_no": f"CAD-2D-{basename[:15].upper()}",
        "project_type": project_type,
        "contingency_pct": 3.0,
        "gst_pct": 18.0,
        "takeoff_groups": parsed["takeoff_groups"],
    }

    xlsx_bytes = generate_boq_excel(payload)
    out_path = os.path.join(EXPORTS_DIR, f"{basename}_BOQ_Takeoff_Abstract.xlsx")
    with open(out_path, "wb") as f:
        f.write(xlsx_bytes)
    return out_path


if __name__ == "__main__":
    samples = ensure_all_samples()
    if len(sys.argv) > 1:
        target = sys.argv[1]
        out = process_single_drawing(target, project_type="All")
        print(f"[OK] Generated BOQ Excel Workbook: {out}")
    else:
        for key, path in samples.items():
            ptype = "Landscape" if "pdf" in key or "landscape" in key or "mixed" in key else key.capitalize()
            out = process_single_drawing(path, project_type=ptype)
            print(f"[OK] {key.upper()} ({os.path.basename(path)}) -> Excel BOQ Workbook: {out}")
