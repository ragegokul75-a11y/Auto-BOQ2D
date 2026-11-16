"""
AutoBOQ 2D - On-screen Takeoff -> Measurement Sheet -> BOQ -> Abstract (IS 1200 / CPWD DSR).

The drawing is rendered faithfully (no auto-guessing). The user measures directly on it:
length / polyline, area polygon, rectangle, count, and picks text callouts.
"""

import os
import shutil
import subprocess
import tempfile
from typing import Any, Dict

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, Response
from fastapi.staticfiles import StaticFiles

from drawing_renderer import render_any
from dsr_library import DSR_ITEMS
from excel_boq_exporter import generate_boq_excel
from create_bbs_excel import OUTPUT_FILENAME as BBS_TEMPLATE_FILENAME, generate_standalone_bbs_workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
SAMPLES_DIR = os.path.join(BASE_DIR, "samples")

SAMPLES = {
    "office_mm": "test_office_style_mm.dxf",
    "villa": "sample_residential_villa.dxf",
    "mixed": "sample_mixed_unlayered_firm_style.dxf",
    "landscape": "sample_landscape_masterplan.dxf",
    "commercial": "sample_commercial_office.dxf",
    "pdf": "sample_2d_drawing_plan.pdf",
}

app = FastAPI(title="AutoBOQ 2D - On-screen Takeoff & BOQ", version="3.0.0")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.get("/", response_class=HTMLResponse)
async def index():
    with open(os.path.join(STATIC_DIR, "index.html"), "r", encoding="utf-8") as f:
        return HTMLResponse(f.read(), headers={"Cache-Control": "no-store"})


@app.get("/api/dsr-library")
async def dsr_library():
    return {"items": DSR_ITEMS}


@app.get("/api/samples")
async def list_samples():
    return {"samples": [{"key": k, "file": v} for k, v in SAMPLES.items()
                        if os.path.exists(os.path.join(SAMPLES_DIR, v))]}


@app.get("/api/sample/{key}")
async def open_sample(key: str, page: int = 0):
    fn = SAMPLES.get(key)
    if not fn:
        raise HTTPException(404, "Unknown sample")
    path = os.path.join(SAMPLES_DIR, fn)
    try:
        data = render_any(path, os.path.splitext(fn)[1], page)
    except ValueError as exc:
        raise HTTPException(422, str(exc)) from exc
    data["filename"] = fn
    return data


@app.post("/api/open")
async def open_drawing(file: UploadFile = File(...), page: int = Form(0)):
    fname = file.filename or "drawing"
    ext = os.path.splitext(fname)[1].lower()
    if ext not in (".dxf", ".dwg", ".pdf"):
        raise HTTPException(400, "Please upload a .DXF, .DWG or .PDF drawing.")
    content = await file.read()
    if not content:
        raise HTTPException(400, "The uploaded file is empty.")
    fd, tmp_path = tempfile.mkstemp(suffix=ext)
    try:
        with os.fdopen(fd, "wb") as fh:
            fh.write(content)
        data = render_any(tmp_path, ext, page)
    except ValueError as exc:
        raise HTTPException(422, str(exc)) from exc
    except Exception as exc:  # unexpected - report clearly instead of a broken screen
        raise HTTPException(500, f"Could not open '{fname}': {type(exc).__name__}: {exc}") from exc
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
    data["filename"] = fname
    return data


@app.post("/api/export-excel")
async def export_excel(payload: Dict[str, Any]):
    xlsx = generate_boq_excel(payload)
    slug = "".join(ch if ch.isalnum() else "_" for ch in payload.get("project_name", "BOQ"))[:40] or "BOQ"
    return Response(
        content=xlsx,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{slug}_BOQ.xlsx"'},
    )


@app.get("/api/download-bbs-template")
async def download_bbs_template():
    bbs_path = os.path.join(BASE_DIR, BBS_TEMPLATE_FILENAME)
    if not os.path.exists(bbs_path):
        generate_standalone_bbs_workbook()
    return FileResponse(
        bbs_path,
        filename=BBS_TEMPLATE_FILENAME,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/api/download-bbs-pdf")
async def download_bbs_pdf():
    pdf_filename = os.path.splitext(BBS_TEMPLATE_FILENAME)[0] + ".pdf"
    pdf_path = os.path.join(BASE_DIR, pdf_filename)
    if not os.path.exists(pdf_path):
        generate_standalone_bbs_workbook()
    if not os.path.exists(pdf_path):
        raise HTTPException(500, "Could not generate BBS PDF.")
    return FileResponse(pdf_path, filename=pdf_filename, media_type="application/pdf")


@app.post("/api/export-pdf")
async def export_pdf(payload: Dict[str, Any]):
    xlsx_bytes = generate_boq_excel(payload)
    slug = "".join(ch if ch.isalnum() else "_" for ch in payload.get("project_name", "BOQ"))[:40] or "BOQ"
    fd_x, tmp_xlsx = tempfile.mkstemp(suffix=".xlsx")
    fd_p, tmp_pdf = tempfile.mkstemp(suffix=".pdf")
    os.close(fd_p)
    try:
        with os.fdopen(fd_x, "wb") as fh:
            fh.write(xlsx_bytes)
        ps1 = os.path.join(BASE_DIR, "export_xlsx_to_pdf.ps1")
        if os.name == "nt" and os.path.exists(ps1):
            subprocess.run(
                ["powershell", "-ExecutionPolicy", "Bypass", "-File", ps1, "-XlsxPath", tmp_xlsx, "-PdfPath", tmp_pdf],
                check=True,
                capture_output=True,
                text=True,
            )
        else:
            soffice = shutil.which("soffice") or shutil.which("libreoffice")
            if not soffice:
                raise RuntimeError("Neither Microsoft Excel (Windows) nor LibreOffice (soffice) is installed for PDF export.")
            out_dir = os.path.dirname(tmp_pdf)
            subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, tmp_xlsx],
                check=True,
                capture_output=True,
                text=True,
            )
            generated_pdf = os.path.splitext(tmp_xlsx)[0] + ".pdf"
            if generated_pdf != tmp_pdf and os.path.exists(generated_pdf):
                shutil.move(generated_pdf, tmp_pdf)
        with open(tmp_pdf, "rb") as pf:
            pdf_bytes = pf.read()
        return Response(
            content=pdf_bytes,
            media_type="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="{slug}_BOQ.pdf"'},
        )
    except Exception as exc:
        raise HTTPException(500, f"PDF export failed: {exc}") from exc
    finally:
        for p in (tmp_xlsx, tmp_pdf):
            try:
                if os.path.exists(p):
                    os.remove(p)
            except OSError:
                pass
