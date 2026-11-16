"""
FastAPI Backend Server for Automated 2D CAD Takeoff, BOQ & Abstract Preparation
(Landscape, Residential & Commercial — IS 1200 / CPWD DSR Standard).
"""

import os
import tempfile
from typing import Dict, Any
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, Response, FileResponse
from fastapi.staticfiles import StaticFiles

from dsr_library import DSR_ITEMS, ITEM_BY_CODE
from dxf_takeoff_engine import parse_dxf_takeoff, recalculate_groups
from excel_boq_exporter import generate_boq_excel
from sample_dxf_generator import ensure_all_samples
from bbs_engine import generate_bbs_from_takeoff, summarize_bbs, SHAPE_CODES
from pdf_dwg_converter import parse_pdf_takeoff, parse_dwg_takeoff
from clean_plan_extractor import extract_clean_dxf_plan, extract_clean_pdf_plan

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
os.makedirs(STATIC_DIR, exist_ok=True)

SAMPLE_PATHS = ensure_all_samples()

app = FastAPI(
    title="AutoBOQ 2D Studio — Landscape, Residential & Commercial CAD Takeoff & BOQ Engine",
    version="2.0.0",
)

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.get("/", response_class=HTMLResponse)
async def serve_index():
    index_path = os.path.join(STATIC_DIR, "index.html")
    if not os.path.exists(index_path):
        return HTMLResponse("<h3>Frontend initializing...</h3>")
    with open(index_path, "r", encoding="utf-8") as f:
        return HTMLResponse(f.read())


@app.get("/api/dsr-library")
async def get_dsr_library():
    """Returns the complete IS 1200 / CPWD DSR Item Master for Landscape, Residential & Commercial works."""
    return {"items": DSR_ITEMS}


@app.get("/api/sample/{sample_key}")
async def load_sample_dxf(sample_key: str, unit_scale: float = 1.0, grouping_mode: str = "SMART"):
    """Parses one of the built-in 2D CAD / PDF sample drawings (mixed, landscape, residential, commercial, pdf_sample)."""
    key = sample_key.lower()
    if key not in SAMPLE_PATHS:
        raise HTTPException(status_code=404, detail=f"Unknown sample key: {sample_key}")

    proj_type_map = {
        "mixed": "Landscape",
        "landscape": "Landscape",
        "residential": "Residential",
        "commercial": "Commercial",
        "pdf_sample": "Landscape",
    }
    project_meta = {
        "mixed": {
            "project_name": "Shared-Layer Villa & Landscape Plan (All Items on '0', 'WALL' & 'WORK_LAYER')",
            "client_name": "M/s Mixed-Layer Architectural & Landscape Practice",
            "drawing_no": "CAD-MIXED-001 (No Separate Layers)",
            "project_type": "Landscape & Residential",
        },
        "landscape": {
            "project_name": "Greenwoods Integrated Township — Phase-1 Landscape Masterplan",
            "client_name": "M/s Prestige Estates & Infrastructure Ltd.",
            "drawing_no": "CAD-LAND-MP-101 (Rev-A)",
            "project_type": "Landscape",
        },
        "residential": {
            "project_name": "Palm Meadows Luxury 3BHK Residential Villa — Ground Floor",
            "client_name": "Mr. & Mrs. Rajeshwar Rao (Private Residence)",
            "drawing_no": "CAD-RES-GA-204 (Rev-02)",
            "project_type": "Residential",
        },
        "commercial": {
            "project_name": "TechnoPark Grade-A Corporate Office Floorplate (Level-04)",
            "client_name": "M/s Apex Global Technologies Pvt. Ltd.",
            "drawing_no": "CAD-COM-ID-401 (GFC)",
            "project_type": "Commercial",
        },
        "pdf_sample": {
            "project_name": "2D Vector PDF Architectural & Landscape Masterplan (Direct PDF Takeoff)",
            "client_name": "M/s PDF Vector Takeoff Demonstration",
            "drawing_no": "PDF-GA-501 (1:100 Vector PDF)",
            "project_type": "Landscape & Residential",
        },
    }

    target_path = SAMPLE_PATHS[key]
    if target_path.lower().endswith(".pdf"):
        parsed = parse_pdf_takeoff(
            target_path,
            unit_scale=unit_scale,
            project_type=proj_type_map[key],
            grouping_mode=grouping_mode,
        )
        parsed["clean_plan"] = extract_clean_pdf_plan(target_path)
    else:
        parsed = parse_dxf_takeoff(
            target_path,
            unit_scale=unit_scale,
            project_type=proj_type_map[key],
            grouping_mode=grouping_mode,
        )
        parsed["clean_plan"] = extract_clean_dxf_plan(target_path, unit_scale=unit_scale)
    parsed["meta"] = project_meta[key]
    parsed["filename"] = os.path.basename(target_path)
    parsed["bbs"] = generate_bbs_from_takeoff(parsed["takeoff_groups"], project_type=proj_type_map[key])
    return parsed


@app.get("/api/download-sample-dxf/{sample_key}")
async def download_sample_dxf_file(sample_key: str):
    key = sample_key.lower()
    if key not in SAMPLE_PATHS:
        raise HTTPException(status_code=404, detail="Sample not found")
    path = SAMPLE_PATHS[key]
    mtype = "application/pdf" if path.lower().endswith(".pdf") else "application/dxf"
    return FileResponse(path, filename=os.path.basename(path), media_type=mtype)


@app.post("/api/upload-dxf")
async def upload_and_parse_drawing(
    file: UploadFile = File(...),
    unit_scale: float = Form(0.0),
    project_type: str = Form("Landscape"),
    grouping_mode: str = Form("SMART"),
):
    """
    Uploads any user 2D AutoCAD (.DXF / .DWG) or 2D Vector (.PDF) drawing file
    and performs instant automated takeoff, DSR mapping, and Bar Bending Schedule generation.
    """
    fname = file.filename or "uploaded_drawing.dxf"
    ext = os.path.splitext(fname)[1].lower()
    if ext not in (".dxf", ".dwg", ".pdf"):
        raise HTTPException(
            status_code=400,
            detail="Unsupported file format. Please upload a .DXF, .DWG, or .PDF 2D drawing file.",
        )

    content = await file.read()
    with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
        tmp.write(content)
        tmp_path = tmp.name

    try:
        if ext == ".pdf":
            parsed = parse_pdf_takeoff(
                tmp_path,
                unit_scale=unit_scale,
                project_type=project_type,
                grouping_mode=grouping_mode,
            )
            parsed["clean_plan"] = extract_clean_pdf_plan(tmp_path)
        elif ext == ".dwg":
            parsed = parse_dwg_takeoff(
                tmp_path,
                original_filename=fname,
                unit_scale=unit_scale,
                project_type=project_type,
                grouping_mode=grouping_mode,
            )
            parsed["clean_plan"] = None
        else:
            parsed = parse_dxf_takeoff(
                tmp_path,
                unit_scale=unit_scale,
                project_type=project_type,
                grouping_mode=grouping_mode,
            )
            parsed["clean_plan"] = extract_clean_dxf_plan(tmp_path, unit_scale=unit_scale)

        parsed["filename"] = fname
        parsed["meta"] = {
            "project_name": os.path.splitext(fname)[0].replace("_", " ").title(),
            "client_name": "M/s Principal Project Developer",
            "drawing_no": f"2D-{ext.lstrip('.').upper()}-{os.path.splitext(fname)[0][:12].upper()}",
            "project_type": project_type,
        }
        parsed["bbs"] = generate_bbs_from_takeoff(parsed["takeoff_groups"], project_type=project_type)
        return parsed
    except Exception as exc:
        raise HTTPException(status_code=422, detail=f"Failed to parse {ext.upper()} drawing: {exc}") from exc
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


@app.post("/api/recalculate")
async def recalculate_boq_payload(payload: Dict[str, Any]):
    """Recalculates takeoff quantities when the user changes DSR mapping, depth/height, or rates."""
    groups = payload.get("takeoff_groups", [])
    recalculate_groups(groups)
    return {"takeoff_groups": groups}


@app.post("/api/generate-bbs")
async def api_generate_bbs(payload: Dict[str, Any]):
    """Regenerates the IS 2502 Bar Bending Schedule from the current 2D CAD takeoff groups."""
    groups = payload.get("takeoff_groups", [])
    ptype = payload.get("project_type", "Residential")
    return generate_bbs_from_takeoff(groups, project_type=ptype)


@app.post("/api/export-excel")
async def export_boq_workbook(payload: Dict[str, Any]):
    """Generates and downloads the 3-sheet IS 1200 / CPWD DSR Excel Workbook (.xlsx) with live formulas."""
    groups = payload.get("takeoff_groups", [])
    recalculate_groups(groups)
    xlsx_bytes = generate_boq_excel(payload)
    proj_slug = payload.get("project_name", "AutoBOQ_Estimate").replace(" ", "_")[:35]
    filename = f"{proj_slug}_BOQ_Takeoff_Abstract.xlsx"
    return Response(
        content=xlsx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
