"""
AutoBOQ 2D — Streamlit Cloud Entrypoint (https://share.streamlit.io/)
Supports:
  1. 2D Drawing Viewer & Takeoff (.DXF / .PDF / .DWG) with Text-Dimension extraction
  2. IS 1200 / CPWD DSR Bill of Quantities (BOQ) & Cost Abstract
  3. IS 2502 Bar Bending Schedule (Footing Bottom + Top Mat Lx/Ly, Column, Beam, Slab)
     with Dynamic Shape Figures, Bend/Crank Angles (0°-180°), and 1-Click Excel & PDF Export
"""

from __future__ import annotations

import base64
import io
import math
import os
import re
import shutil
import subprocess
import tempfile
from typing import Any, Dict, List

import pymupdf
import streamlit as st
from PIL import Image, ImageDraw

from bbs_diagrams import ensure_bbs_diagrams
import importlib
import create_bbs_excel as _bbs_mod
if not hasattr(_bbs_mod, "compile_bbs_workbooks_bytes"):
    _bbs_mod = importlib.reload(_bbs_mod)
from create_bbs_excel import (
    OUTPUT_FILENAME as BBS_XLSX_NAME,
    compile_bbs_workbooks_bytes,
    generate_standalone_bbs_workbook,
)
from drawing_renderer import render_any
from dsr_library import DSR_ITEMS, ITEM_BY_CODE
from excel_boq_exporter import generate_boq_excel
from modular_boq_generator import (
    MASTER_FILENAME as MODULAR_BOQ_NAME,
    compile_modular_boq_bytes,
    default_autocad_sample_rows,
    get_modular_boq_preview,
    parse_autocad_takeoff,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SAMPLES_DIR = os.path.join(BASE_DIR, "samples")

st.set_page_config(
    page_title="AutoBOQ 2D — Takeoff, BOQ & BBS",
    page_icon="📐",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _xlsx_bytes_to_pdf_bytes(xlsx_bytes: bytes, payload: Dict[str, Any] | None = None) -> bytes:
    """Converts an .xlsx workbook to .pdf using Excel COM (Windows), LibreOffice, or pure-Python PyMuPDF."""
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
            with open(tmp_pdf, "rb") as pf:
                return pf.read()
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if soffice:
            out_dir = os.path.dirname(tmp_pdf)
            subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, tmp_xlsx],
                check=True,
                capture_output=True,
                text=True,
            )
            gen_pdf = os.path.splitext(tmp_xlsx)[0] + ".pdf"
            if gen_pdf != tmp_pdf and os.path.exists(gen_pdf):
                shutil.move(gen_pdf, tmp_pdf)
            with open(tmp_pdf, "rb") as pf:
                return pf.read()
        # Pure-Python fallback using PyMuPDF (works on Streamlit Cloud with zero system packages)
        return _payload_to_pdf_pymupdf(payload or {})
    finally:
        for p in (tmp_xlsx, tmp_pdf):
            try:
                if os.path.exists(p):
                    os.remove(p)
            except OSError:
                pass


def _payload_to_pdf_pymupdf(payload: Dict[str, Any]) -> bytes:
    """Renders a clean A4 landscape PDF summary of the BOQ, Abstract & BBS using PyMuPDF."""
    doc = pymupdf.open()
    page = doc.new_page(width=842, height=595)  # A4 landscape
    y = 40
    proj = payload.get("project_name", "AutoBOQ Estimate")
    client = payload.get("client_name", "")
    dwg = payload.get("drawing_no", "")

    def line(txt: str, size: float = 10.0, bold: bool = False):
        nonlocal page, y
        if y > 555:
            page = doc.new_page(width=842, height=595)
            y = 40
        fontname = "helv" if not bold else "hebo"
        page.insert_text((36, y), txt, fontsize=size, fontname=fontname, color=(0.06, 0.09, 0.16))
        y += size + 6

    line(f"BILL OF QUANTITIES & ABSTRACT — {proj}", size=14, bold=True)
    line(f"Client: {client}   |   Drawing No: {dwg}", size=10, bold=False)
    y += 6
    line("S.No  Code        Description                                              Qty       Unit     Rate (INR)    Amount (INR)", size=9.5, bold=True)
    sub = 0.0
    for idx, grp in enumerate(payload.get("takeoff_groups", []), start=1):
        qty = sum(m["nos"] * m["length"] * m["breadth"] * m["depth"] for m in grp.get("measurements", []))
        amt = qty * grp.get("rate", 0.0)
        sub += amt
        desc = (grp.get("description") or "")[:48]
        line(f"{idx:<5} {grp.get('dsr_code',''):<11} {desc:<48} {qty:>8.2f}  {grp.get('unit',''):<7} {grp.get('rate',0):>10.2f}  {amt:>12.2f}", size=8.5)
    cont_pct = float(payload.get("contingency_pct", 3.0))
    gst_pct = float(payload.get("gst_pct", 18.0))
    cont = sub * cont_pct / 100.0
    gst = (sub + cont) * gst_pct / 100.0
    y += 6
    line(f"Sub-Total: INR {sub:,.2f}   |   Contingency ({cont_pct}%): INR {cont:,.2f}   |   GST ({gst_pct}%): INR {gst:,.2f}", size=10, bold=True)
    line(f"GRAND TOTAL: INR {sub + cont + gst:,.2f}", size=12, bold=True)

    bbs_rows = payload.get("bbs_rows") or []
    if bbs_rows:
        y += 12
        line("BAR BENDING SCHEDULE (REINFORCEMENT STEEL)", size=12, bold=True)
        tot_kg = 0.0
        for idx, b in enumerate(bbs_rows, start=1):
            dia = int(b.get("dia_mm", 12))
            tl = int(b.get("num_members", 1)) * int(b.get("bars_per_member", 1)) * float(b.get("a_m", 0))
            kg = tl * ((dia * dia) / 162.2)
            tot_kg += kg
            line(f"{idx:<4} {b.get('member_group','')[:18]:<19} {b.get('shape_code','SC-00'):<7} Ø{dia:<3}mm  Cut:{b.get('a_m',0):>6.3f}m  Total:{tl:>7.2f}m  Wt:{kg:>8.2f} kg", size=8.5)
        line(f"TOTAL STEEL WEIGHT: {tot_kg:,.2f} kg ({tot_kg/1000:.3f} MT)", size=11, bold=True)

    return doc.tobytes()


def _render_dxf_preview_image(dwg: Dict[str, Any]) -> Image.Image:
    """Renders DXF strokes and fills onto a crisp white PNG preview."""
    w, h = 1100, 650
    img = Image.new("RGB", (w, h), "#FFFFFF")
    draw = ImageDraw.Draw(img)
    b = dwg["bounds"]
    bw = max(b["maxx"] - b["minx"], 1e-6)
    bh = max(b["maxy"] - b["miny"], 1e-6)
    pad = 28
    sc = min((w - 2 * pad) / bw, (h - 2 * pad) / bh)
    ox = (w - bw * sc) / 2.0 - b["minx"] * sc
    oy = (h - bh * sc) / 2.0 + b["maxy"] * sc

    def to_px(x: float, y: float):
        return (x * sc + ox, oy - y * sc)

    for fill_grp in dwg.get("fills", []):
        for poly in fill_grp["p"]:
            if len(poly) >= 6:
                pts = [to_px(poly[i], poly[i + 1]) for i in range(0, len(poly) - 1, 2)]
                draw.polygon(pts, fill="#E2E8F0")

    for stroke_grp in dwg.get("strokes", []):
        col = stroke_grp.get("c", "#1E293B")
        if col.lower() in ("#ffffff", "#fefefe"):
            col = "#1E293B"
        for pl in stroke_grp["p"]:
            if len(pl) >= 4:
                pts = [to_px(pl[i], pl[i + 1]) for i in range(0, len(pl) - 1, 2)]
                draw.line(pts, fill=col, width=2)
    return img


def _parse_dims_from_text(s: str):
    s = s.replace("×", "x")
    m = re.search(r"(\d+)\s*'\s*-?\s*(\d+(?:\.\d+)?)?\s*\"?\s*[xX*]\s*(\d+)\s*'\s*-?\s*(\d+(?:\.\d+)?)?\s*\"?", s)
    if m:
        a = (int(m.group(1)) * 12 + float(m.group(2) or 0)) * 0.0254
        b = (int(m.group(3)) * 12 + float(m.group(4) or 0)) * 0.0254
        return round(a, 3), round(b, 3)
    m = re.search(r"(\d+(?:\.\d+)?)\s*(mm|m)?\s*[xX*]\s*(\d+(?:\.\d+)?)\s*(mm|m)?", s)
    if m:
        a, b = float(m.group(1)), float(m.group(3))
        unit = (m.group(2) or m.group(4) or "").lower()
        if unit == "mm" or (unit == "" and a >= 100 and b >= 100):
            a /= 1000.0
            b /= 1000.0
        return round(a, 3), round(b, 3)
    return None


def _ensure_bbs_files():
    xlsx_path = os.path.join(BASE_DIR, BBS_XLSX_NAME)
    pdf_path = os.path.join(BASE_DIR, os.path.splitext(BBS_XLSX_NAME)[0] + ".pdf")
    if not os.path.exists(xlsx_path):
        generate_standalone_bbs_workbook()
    if not os.path.exists(pdf_path) and os.path.exists(xlsx_path):
        try:
            with open(xlsx_path, "rb") as xf:
                pdf_b = _xlsx_bytes_to_pdf_bytes(xf.read())
            with open(pdf_path, "wb") as pf:
                pf.write(pdf_b)
        except Exception:
            pass
    return xlsx_path, pdf_path


def _safe_read_export_bytes(filename: str) -> bytes:
    """Read an exported workbook from exports/ or root, auto-generating if missing."""
    candidates = [
        os.path.join(BASE_DIR, "exports", filename),
        os.path.join(BASE_DIR, filename),
    ]
    for path in candidates:
        try:
            with open(path, "rb") as fh:
                return fh.read()
        except (FileNotFoundError, OSError):
            continue
    generate_standalone_bbs_workbook()
    for path in candidates:
        try:
            with open(path, "rb") as fh:
                return fh.read()
        except (FileNotFoundError, OSError):
            continue
    raise FileNotFoundError(f"Could not generate or read {filename}")


@st.cache_data(show_spinner=False)
def _get_compiled_bbs_bytes(params_json: str) -> Dict[str, bytes]:
    import json

    try:
        return compile_bbs_workbooks_bytes(json.loads(params_json))
    except Exception:
        trade_names = [
            BBS_XLSX_NAME,
            "Footing_BBS_Only.xlsx",
            "Column_BBS_Only.xlsx",
            "Beam_BBS_Only.xlsx",
            "Slab_BBS_Only.xlsx",
            "Summary_Tender_BOQ.xlsx",
        ]
        return {name: _safe_read_export_bytes(name) for name in trade_names}


# ---------------------------------------------------------------------------
# Session State
# ---------------------------------------------------------------------------
if "boq_rows" not in st.session_state:
    st.session_state.boq_rows = []
if "steel_rows" not in st.session_state:
    st.session_state.steel_rows = []
if "dwg_data" not in st.session_state:
    st.session_state.dwg_data = None
if "diag_paths" not in st.session_state:
    st.session_state.diag_paths = ensure_bbs_diagrams()

# ---------------------------------------------------------------------------
# Sidebar: Live Structural Parameters & Download Center
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ BBS Structural Inputs")
    sh1, sh2 = st.columns(2)
    sb_hc = sh1.number_input("Floor Height HC (mm)", value=3660, min_value=2400, max_value=6000, step=60)
    sb_floors = sh2.number_input("Typical Floors (×)", value=3, min_value=1, max_value=20, step=1)
    sb_neck = st.number_input("Zone A Neck Height (mm)", value=1800, min_value=600, max_value=4000, step=100)
    sc1, sc2 = st.columns(2)
    sb_c1_b = sc1.number_input("C1 Width b (mm)", value=300, min_value=200, max_value=1200, step=25)
    sb_c1_d = sc2.number_input("C1 Depth D (mm)", value=450, min_value=200, max_value=1500, step=25)
    sv1, sv2 = st.columns(2)
    sb_col_cov = sv1.number_input("Col Cover (mm)", value=40, min_value=25, max_value=75, step=5)
    sb_ftg_cov = sv2.number_input("Footing Cover (mm)", value=50, min_value=40, max_value=75, step=5)

    with st.expander("Footing Marks (F1–F6) Overrides", expanded=False):
        st.caption("Edit Nos, Dimensions & Pad Depth Df (mm) — Column VLOOKUP updates automatically.")
        f_cfg: Dict[str, Dict[str, int]] = {}
        f_defaults = {
            "F1": (6, 1800, 1800, 525),
            "F2": (4, 2100, 2100, 600),
            "F3": (2, 2400, 2400, 680),
            "F4": (2, 2700, 2700, 750),
            "F5": (2, 3000, 3000, 830),
            "F6": (2, 3200, 3200, 830),
        }
        for mk, (d_nos, d_lx, d_ly, d_df) in f_defaults.items():
            st.markdown(f"**{mk}**")
            fc_a, fc_b, fc_c, fc_d = st.columns(4)
            nos_v = fc_a.number_input("Nos", value=d_nos, min_value=0, step=1, key=f"sb_{mk}_nos")
            lx_v = fc_b.number_input("Lx", value=d_lx, min_value=600, step=100, key=f"sb_{mk}_lx")
            ly_v = fc_c.number_input("Ly", value=d_ly, min_value=600, step=100, key=f"sb_{mk}_ly")
            df_v = fc_d.number_input("Df", value=d_df, min_value=200, step=25, key=f"sb_{mk}_df")
            if nos_v > 0 or mk in ("F1", "F2", "F3"):
                f_cfg[mk] = {"nos": int(nos_v), "lx": int(lx_v), "ly": int(ly_v), "df": int(df_v)}

    sr1, sr2 = st.columns(2)
    sb_rate = sr1.number_input("Steel Rate (₹/kg)", value=82.0, min_value=10.0, step=1.0, key="sb_rate")
    sb_wast = sr2.number_input("Wastage (%)", value=3.0, min_value=0.0, max_value=15.0, step=0.5, key="sb_wast")

    import json as _json

    bbs_params = {
        "clear_height_mm": int(sb_hc),
        "typical_floors": int(sb_floors),
        "neck_height_mm": int(sb_neck),
        "c1_b_mm": int(sb_c1_b),
        "c1_d_mm": int(sb_c1_d),
        "col_cover_mm": int(sb_col_cov),
        "footing_cover_mm": int(sb_ftg_cov),
        "footings": f_cfg,
        "tender_rate_inr": float(sb_rate),
        "wastage_pct": float(sb_wast) / 100.0,
    }
    compiled_bbs = _get_compiled_bbs_bytes(_json.dumps(bbs_params, sort_keys=True))

    st.divider()
    st.header("📥 Download Center")
    st.caption("Instant single-trade or complete master Excel downloads (0 formula errors).")
    _xlsx_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    st.download_button(
        "⬇️ Download Footing BBS",
        data=compiled_bbs["Footing_BBS_Only.xlsx"],
        file_name="Footing_BBS_Only.xlsx",
        mime=_xlsx_mime,
        use_container_width=True,
        key="sb_dl_ftg",
    )
    st.download_button(
        "⬇️ Download Column BBS",
        data=compiled_bbs["Column_BBS_Only.xlsx"],
        file_name="Column_BBS_Only.xlsx",
        mime=_xlsx_mime,
        use_container_width=True,
        key="sb_dl_col",
    )
    st.download_button(
        "⬇️ Download Beam BBS",
        data=compiled_bbs["Beam_BBS_Only.xlsx"],
        file_name="Beam_BBS_Only.xlsx",
        mime=_xlsx_mime,
        use_container_width=True,
        key="sb_dl_bm",
    )
    st.download_button(
        "⬇️ Download Slab BBS",
        data=compiled_bbs["Slab_BBS_Only.xlsx"],
        file_name="Slab_BBS_Only.xlsx",
        mime=_xlsx_mime,
        use_container_width=True,
        key="sb_dl_slb",
    )
    st.download_button(
        "⬇️ Download Summary BOQ",
        data=compiled_bbs["Summary_Tender_BOQ.xlsx"],
        file_name="Summary_Tender_BOQ.xlsx",
        mime=_xlsx_mime,
        use_container_width=True,
        key="sb_dl_sum",
    )
    st.download_button(
        "📦 Download Complete Master BBS Workbook (All Sheets)",
        data=compiled_bbs[BBS_XLSX_NAME],
        file_name=BBS_XLSX_NAME,
        mime=_xlsx_mime,
        type="primary",
        use_container_width=True,
        key="sb_dl_master",
    )

# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
st.title("📐 AutoBOQ 2D — Takeoff, BOQ & Bar Bending Schedule")
st.caption("IS 1200 / CPWD DSR Quantity Takeoff, Abstract of Cost, and IS 2502 Bar Bending Schedule (Footing, Column, Beam, Slab)")

tab1, tab2, tab3, tab4 = st.tabs([
    "1. Drawing & Takeoff",
    "2. BOQ & Abstract (Excel / PDF)",
    "3. Bar Bending Schedule (Steel kg)",
    "4. Smart AutoCAD → Modular BOQ (12 Tabs)",
])

# ===========================================================================
# TAB 1: DRAWING & TAKEOFF
# ===========================================================================
with tab1:
    c_up, c_smp = st.columns([2, 1])
    with c_up:
        up = st.file_uploader("Upload 2D Drawing (.DXF, .PDF, .DWG)", type=["dxf", "pdf", "dwg"])
    with c_smp:
        sample_files = {
            "— Choose a sample —": None,
            "Office-style plan (mm, layer 0)": "test_office_style_mm.dxf",
            "Residential villa (m)": "sample_residential_villa.dxf",
            "Mixed / unlayered plan (m)": "sample_mixed_unlayered_firm_style.dxf",
            "Landscape masterplan (m)": "sample_landscape_masterplan.dxf",
            "Commercial office (m)": "sample_commercial_office.dxf",
            "PDF 2D plan": "sample_2d_drawing_plan.pdf",
        }
        smp_choice = st.selectbox("Or load a built-in sample", list(sample_files.keys()))

    if up is not None:
        ext = os.path.splitext(up.name)[1].lower()
        fd, tmp_p = tempfile.mkstemp(suffix=ext)
        try:
            with os.fdopen(fd, "wb") as fh:
                fh.write(up.getvalue())
            st.session_state.dwg_data = render_any(tmp_p, ext, 0)
            st.session_state.dwg_data["filename"] = up.name
        except Exception as exc:
            st.error(f"Could not open {up.name}: {exc}")
        finally:
            if os.path.exists(tmp_p):
                os.remove(tmp_p)
    elif sample_files.get(smp_choice):
        sp = os.path.join(SAMPLES_DIR, sample_files[smp_choice])
        if os.path.exists(sp):
            try:
                st.session_state.dwg_data = render_any(sp, os.path.splitext(sp)[1], 0)
                st.session_state.dwg_data["filename"] = sample_files[smp_choice]
            except Exception as exc:
                st.error(f"Error loading sample: {exc}")

    col_plan, col_meas = st.columns([1.4, 1])
    with col_plan:
        dwg = st.session_state.dwg_data
        if dwg:
            st.subheader(f"Drawing Preview: {dwg.get('filename', '')} ({dwg['format']})")
            if dwg.get("image"):
                b64 = dwg["image"]["src"].split(",", 1)[1]
                st.image(base64.b64decode(b64), width=680)
            else:
                st.image(_render_dxf_preview_image(dwg), width=680)
            st.caption(
                f"Units: {dwg.get('unit_name')} | Polylines: {dwg['stats']['polylines']} | Texts: {dwg['stats']['texts']}"
            )
        else:
            st.info("Upload a `.DXF` or `.PDF` drawing above (or choose a sample) to preview it here.")

    with col_meas:
        st.subheader("Add Measurement to BOQ")
        dsr_labels = [f"{d['code']} · {d['description'][:55]} ({d['unit']})" for d in DSR_ITEMS]
        dsr_sel = st.selectbox("Select IS 1200 / CPWD DSR Item", dsr_labels)
        dsr_obj = DSR_ITEMS[dsr_labels.index(dsr_sel)]

        # Quick pick from detected text dimensions on the drawing
        detected_dims = []
        if dwg:
            for t in dwg.get("texts", []):
                dm = _parse_dims_from_text(t["t"])
                if dm:
                    detected_dims.append((t["t"], dm[0], dm[1]))

        if detected_dims:
            pick_lbl = st.selectbox(
                "Detected Text Dimensions (optional)",
                ["— Manual entry —"] + [f"{lbl}  ({a}m × {b}m)" for lbl, a, b in detected_dims],
            )
            if pick_lbl != "— Manual entry —":
                idx_d = [f"{lbl}  ({a}m × {b}m)" for lbl, a, b in detected_dims].index(pick_lbl)
                def_lbl, def_l, def_b = detected_dims[idx_d]
            else:
                def_lbl, def_l, def_b = "Measurement", 4.5, 3.6
        else:
            def_lbl, def_l, def_b = "Measurement", 4.5, 3.6

        m_lbl = st.text_input("Location / Room Label", value=def_lbl)
        c1, c2, c3, c4 = st.columns(4)
        m_nos = c1.number_input("Nos", value=1, step=1)
        m_l = c2.number_input("L (m)", value=float(def_l), step=0.1)
        m_b = c3.number_input("B (m)", value=float(def_b), step=0.1)
        m_d = c4.number_input("D / H (m)", value=float(dsr_obj.get("default_depth", 1.0) or 1.0), step=0.05)

        if st.button("＋ Add to BOQ", type="primary", use_container_width=True):
            st.session_state.boq_rows.append({
                "dsr_code": dsr_obj["code"],
                "category": dsr_obj["category"],
                "description": dsr_obj["description"],
                "unit": dsr_obj["unit"],
                "rate": float(dsr_obj["rate"]),
                "calc_mode": dsr_obj["calc_mode"],
                "is_ref": dsr_obj["is_ref"],
                "label": m_lbl,
                "nos": float(m_nos),
                "length": float(m_l),
                "breadth": float(m_b),
                "depth": float(m_d),
            })
            st.success(f"Added {m_lbl} to {dsr_obj['code']}")


# ===========================================================================
# TAB 2: BOQ & ABSTRACT (EXCEL / PDF EXPORT)
# ===========================================================================
with tab2:
    c_m1, c_m2, c_m3, c_m4, c_m5 = st.columns([2, 1.5, 1.2, 0.8, 0.8])
    proj_name = c_m1.text_input("Project Name", "Proposed Building & Landscape Works")
    client_name = c_m2.text_input("Client", "Principal Project Developer")
    dwg_no = c_m3.text_input("Drawing No.", "GA-001")
    cont_pct = c_m4.number_input("Contingency %", value=3.0, step=0.5)
    gst_pct = c_m5.number_input("GST %", value=18.0, step=1.0)

    if st.session_state.boq_rows:
        display_rows = []
        sub_tot = 0.0
        for idx, r in enumerate(st.session_state.boq_rows, start=1):
            qty = round(r["nos"] * r["length"] * r["breadth"] * r["depth"], 3)
            amt = round(qty * r["rate"], 2)
            sub_tot += amt
            display_rows.append({
                "S.No": idx,
                "Code": r["dsr_code"],
                "Location": r["label"],
                "Nos": r["nos"],
                "L": r["length"],
                "B": r["breadth"],
                "D": r["depth"],
                "Qty": qty,
                "Unit": r["unit"],
                "Rate (₹)": r["rate"],
                "Amount (₹)": amt,
            })
        st.dataframe(display_rows, use_container_width=True, hide_index=True)

        cont_amt = sub_tot * cont_pct / 100.0
        gst_amt = (sub_tot + cont_amt) * gst_pct / 100.0
        grand_tot = sub_tot + cont_amt + gst_amt
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Sub-Total", f"₹ {sub_tot:,.2f}")
        m2.metric(f"Contingency ({cont_pct}%)", f"₹ {cont_amt:,.2f}")
        m3.metric(f"GST ({gst_pct}%)", f"₹ {gst_amt:,.2f}")
        m4.metric("Grand Total", f"₹ {grand_tot:,.2f}")

        if st.button("Clear BOQ Rows"):
            st.session_state.boq_rows = []
            st.rerun()
    else:
        st.info("No BOQ items yet. Add measurements in Tab 1 or push Steel (kg) from Tab 3.")

    # Build payload for Excel / PDF export
    grouped: Dict[str, Dict[str, Any]] = {}
    for r in st.session_state.boq_rows:
        c = r["dsr_code"]
        if c not in grouped:
            grouped[c] = {
                "group_id": c,
                "dsr_code": c,
                "category": r["category"],
                "display_name": r["description"][:55],
                "description": r["description"],
                "unit": r["unit"],
                "rate": r["rate"],
                "calc_mode": r["calc_mode"],
                "is_ref": r["is_ref"],
                "layer": "TAKEOFF",
                "included": True,
                "measurements": [],
            }
        grouped[c]["measurements"].append({
            "label": r["label"],
            "nos": r["nos"],
            "length": r["length"],
            "breadth": r["breadth"],
            "depth": r["depth"],
            "raw_area": 0,
        })

    bbs_export_rows = []
    for i, sr in enumerate(st.session_state.steel_rows, start=1):
        bbs_export_rows.append({
            "member_group": sr["member"],
            "bar_mark": f"B{i}",
            "description": sr["desc"],
            "shape_code": sr["shape"],
            "dia_mm": sr["dia"],
            "spacing_mm": sr.get("spacing", 0),
            "num_members": sr["nos"],
            "bars_per_member": sr["bars"],
            "a_m": sr["cut"],
            "b_m": 0.0,
            "c_m": 0.0,
            "d_m": 0.0,
            "bend_deduction_m": 0.0,
        })

    export_payload = {
        "project_name": proj_name,
        "client_name": client_name,
        "drawing_no": dwg_no,
        "contingency_pct": cont_pct,
        "gst_pct": gst_pct,
        "project_type": "Landscape, Residential & Commercial",
        "takeoff_groups": list(grouped.values()),
        "bbs_rows": bbs_export_rows,
    }

    if st.session_state.boq_rows or st.session_state.steel_rows:
        xlsx_data = generate_boq_excel(export_payload)
        col_ex, col_pd = st.columns(2)
        col_ex.download_button(
            "⬇ Download Project BOQ Excel (.xlsx)",
            data=xlsx_data,
            file_name="AutoBOQ_Estimate.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        with col_pd:
            if st.button("⚙ Prepare Project BOQ PDF (.pdf)", use_container_width=True):
                try:
                    pdf_data = _xlsx_bytes_to_pdf_bytes(xlsx_data, export_payload)
                    st.download_button(
                        "⬇ Save Project BOQ PDF",
                        data=pdf_data,
                        file_name="AutoBOQ_Estimate.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )
                except Exception as exc:
                    st.error(f"PDF export error: {exc}")

# ===========================================================================
# TAB 3: BAR BENDING SCHEDULE (FOOTING, COLUMN, BEAM, SLAB)
# ===========================================================================
with tab3:
    diags = st.session_state.diag_paths
    _, pdf_tpl = _ensure_bbs_files()

    st.subheader("📥 BBS Download Center (Single-Trade & Complete Master Workbook)")
    st.caption(
        f"Live-compiled with **HC = {sb_hc} mm**, **Floors = {sb_floors}×**, **Neck = {sb_neck} mm**, "
        f"**C1 = {sb_c1_b}×{sb_c1_d} mm**, **Rate = ₹{sb_rate:.0f}/kg**, **Wastage = {sb_wast:.1f}%**."
    )
    tc1, tc2, tc3, tc4, tc5 = st.columns(5)
    tc1.download_button(
        "⬇️ Download Footing BBS",
        data=compiled_bbs["Footing_BBS_Only.xlsx"],
        file_name="Footing_BBS_Only.xlsx",
        mime=_xlsx_mime,
        use_container_width=True,
        key="tab_dl_ftg",
    )
    tc2.download_button(
        "⬇️ Download Column BBS",
        data=compiled_bbs["Column_BBS_Only.xlsx"],
        file_name="Column_BBS_Only.xlsx",
        mime=_xlsx_mime,
        use_container_width=True,
        key="tab_dl_col",
    )
    tc3.download_button(
        "⬇️ Download Beam BBS",
        data=compiled_bbs["Beam_BBS_Only.xlsx"],
        file_name="Beam_BBS_Only.xlsx",
        mime=_xlsx_mime,
        use_container_width=True,
        key="tab_dl_bm",
    )
    tc4.download_button(
        "⬇️ Download Slab BBS",
        data=compiled_bbs["Slab_BBS_Only.xlsx"],
        file_name="Slab_BBS_Only.xlsx",
        mime=_xlsx_mime,
        use_container_width=True,
        key="tab_dl_slb",
    )
    tc5.download_button(
        "⬇️ Download Summary BOQ",
        data=compiled_bbs["Summary_Tender_BOQ.xlsx"],
        file_name="Summary_Tender_BOQ.xlsx",
        mime=_xlsx_mime,
        use_container_width=True,
        key="tab_dl_sum",
    )

    d1, d2 = st.columns(2)
    d1.download_button(
        "📦 Download Complete Master BBS Workbook (All Sheets)",
        data=compiled_bbs[BBS_XLSX_NAME],
        file_name=BBS_XLSX_NAME,
        mime=_xlsx_mime,
        type="primary",
        use_container_width=True,
        key="tab_dl_master",
    )
    if os.path.exists(pdf_tpl):
        with open(pdf_tpl, "rb") as pf:
            d2.download_button(
                "⬇ Download 7-Page BBS PDF (Footing, Column, Beam, Slab)",
                data=pf.read(),
                file_name=os.path.splitext(BBS_XLSX_NAME)[0] + ".pdf",
                mime="application/pdf",
                use_container_width=True,
                key="tab_dl_pdf",
            )

    st.divider()
    st.subheader("1. Isolated Footing Steel (Bottom Mesh & Top Mat Lx / Ly)")
    fc1, fc2, fc3, fc4, fc5, fc6 = st.columns(6)
    ft_mark = fc1.text_input("Footing Mark", "F1")
    ft_nos = fc2.number_input("No. of Footings", value=4, min_value=1)
    ft_L = fc3.number_input("Lx (mm)", value=1800, step=50)
    ft_B = fc4.number_input("Ly (mm)", value=1500, step=50)
    ft_D = fc5.number_input("Depth D (mm)", value=450, step=25)
    ft_c = fc6.number_input("Cover (mm)", value=50, step=5)

    fb1, fb2, fb3, fb4 = st.columns(4)
    ft_dx = fb1.selectbox("Bottom Lx Dia (mm)", [8, 10, 12, 16, 20, 25], index=2)
    ft_sx = fb2.number_input("Bottom Lx Spacing (mm)", value=150, step=25)
    ft_dy = fb3.selectbox("Bottom Ly Dia (mm)", [8, 10, 12, 16, 20, 25], index=2)
    ft_sy = fb4.number_input("Bottom Ly Spacing (mm)", value=150, step=25)

    has_top_mat = st.checkbox("Include Top Mat (Lx & Ly)")
    if has_top_mat:
        tb1, tb2, tb3, tb4 = st.columns(4)
        ft_tdx = tb1.selectbox("Top Mat Lx Dia (mm)", [8, 10, 12, 16, 20], index=1)
        ft_tsx = tb2.number_input("Top Mat Lx Spacing (mm)", value=150, step=25)
        ft_tdy = tb3.selectbox("Top Mat Ly Dia (mm)", [8, 10, 12, 16, 20], index=1)
        ft_tsy = tb4.number_input("Top Mat Ly Spacing (mm)", value=150, step=25)

    if st.button("＋ Add Footing Steel to List", type="primary"):
        leg = max(0, ft_D - 2 * ft_c)
        nx = math.floor((ft_B - 2 * ft_c) / ft_sx) + 1
        cut_x = round((ft_L - 2 * ft_c + 2 * leg - 4 * ft_dx) / 1000.0, 3)
        ny = math.floor((ft_L - 2 * ft_c) / ft_sy) + 1
        cut_y = round((ft_B - 2 * ft_c + 2 * leg - 4 * ft_dy) / 1000.0, 3)
        st.session_state.steel_rows.append({
            "member": f"Footing {ft_mark}", "shape": "SC-02", "angle": 90, "bar_type": "BOT-MESH-LX",
            "desc": f"Bottom Mesh Lx Ø{ft_dx}@{ft_sx}", "nos": int(ft_nos), "bars": nx, "dia": ft_dx, "spacing": ft_sx, "cut": cut_x,
        })
        st.session_state.steel_rows.append({
            "member": f"Footing {ft_mark}", "shape": "SC-02", "angle": 90, "bar_type": "BOT-MESH-LY",
            "desc": f"Bottom Mesh Ly Ø{ft_dy}@{ft_sy}", "nos": int(ft_nos), "bars": ny, "dia": ft_dy, "spacing": ft_sy, "cut": cut_y,
        })
        if has_top_mat:
            tnx = math.floor((ft_B - 2 * ft_c) / ft_tsx) + 1
            tcut_x = round((ft_L - 2 * ft_c + 2 * leg - 4 * ft_tdx) / 1000.0, 3)
            tny = math.floor((ft_L - 2 * ft_c) / ft_tsy) + 1
            tcut_y = round((ft_B - 2 * ft_c + 2 * leg - 4 * ft_tdy) / 1000.0, 3)
            st.session_state.steel_rows.append({
                "member": f"Footing {ft_mark}", "shape": "SC-02", "angle": 90, "bar_type": "TOP-MAT-LX",
                "desc": f"Top Mat Lx Ø{ft_tdx}@{ft_tsx}", "nos": int(ft_nos), "bars": tnx, "dia": ft_tdx, "spacing": ft_tsx, "cut": tcut_x,
            })
            st.session_state.steel_rows.append({
                "member": f"Footing {ft_mark}", "shape": "SC-02", "angle": 90, "bar_type": "TOP-MAT-LY",
                "desc": f"Top Mat Ly Ø{ft_tdy}@{ft_tsy}", "nos": int(ft_nos), "bars": tny, "dia": ft_tdy, "spacing": ft_tsy, "cut": tcut_y,
            })
        st.success(f"Added Footing {ft_mark} to Steel List")

    st.divider()
    st.subheader("2. Column / Beam / Slab Bars (Shape Code, Angle & Live Figure)")
    o1, o2, o3, o4 = st.columns([1.1, 1, 0.8, 1.1])
    ob_mem = o1.text_input("Member", "Beam B1")
    ob_shape = o2.selectbox("Shape Code", ["SC-00", "SC-01", "SC-02", "SC-41", "SC-51"], index=3)
    ob_ang = o3.selectbox("Angle (°)", [0, 45, 90, 135, 180], index=1)
    ob_type = o4.selectbox("Bar Type", ["CRANK", "STIRRUP", "EXTRA_SUPPORT", "BOT-MESH-LX", "BOT-MESH-LY", "TOP-MAT-LX", "TOP-MAT-LY", "MAIN", "DIST"])

    o5, o6, o7, o8, o9 = st.columns([2, 0.7, 0.7, 0.8, 0.9])
    ob_desc = o5.text_input("Bar Description", "Bottom Cranked Bar (45° Bent-up)")
    ob_nos = o6.number_input("Members", value=4, min_value=1)
    ob_bars = o7.number_input("Bars/mem", value=2, min_value=1)
    ob_dia = o8.selectbox("Dia (mm)", [8, 10, 12, 16, 20, 25, 32], index=3)
    ob_cut = o9.number_input("Cut Len (m)", value=5.98, step=0.05)

    thumb_key = f"{ob_shape}_{ob_ang}"
    thumb_path = diags.get(thumb_key) or diags.get(ob_shape)
    p_col1, p_col2 = st.columns([1, 3])
    if thumb_path and os.path.exists(thumb_path):
        p_col1.image(thumb_path, width=180)
    p_col2.write(f"**Live Shape Preview:** `{ob_shape}` at `{ob_ang}°` | Type: `{ob_type}`")
    if p_col2.button("＋ Add Bar to Steel List"):
        st.session_state.steel_rows.append({
            "member": ob_mem, "shape": ob_shape, "angle": ob_ang, "bar_type": ob_type,
            "desc": ob_desc, "nos": int(ob_nos), "bars": int(ob_bars), "dia": ob_dia, "spacing": 0, "cut": float(ob_cut),
        })

    if st.session_state.steel_rows:
        st.subheader("Steel Schedule & Weight Summary (kg)")
        tbl = []
        tot_kg = 0.0
        for r in st.session_state.steel_rows:
            tl = r["nos"] * r["bars"] * r["cut"]
            uw = round((r["dia"] ** 2) / 162.2, 3)
            kg = round(tl * uw, 2)
            tot_kg += kg
            tbl.append({
                "Member": r["member"], "Shape": r["shape"], "Angle": f"{r['angle']}°",
                "Bar Type": r["bar_type"], "Description": r["desc"],
                "Nos": r["nos"], "Bars/Mem": r["bars"], "Dia (mm)": r["dia"],
                "Cut (m)": r["cut"], "Total (m)": round(tl, 2), "Weight (kg)": kg,
            })
        st.dataframe(tbl, use_container_width=True, hide_index=True)
        steel_rate = st.number_input("Steel Tender Rate (₹/kg)", value=82.0, step=1.0)
        st.metric("Total Reinforcement Steel", f"{tot_kg:,.2f} kg  ({tot_kg/1000:.3f} MT)", f"₹ {tot_kg * steel_rate:,.2f}")
        c_push, c_clr = st.columns(2)
        if c_push.button("→ Push Total Steel (kg) into BOQ", type="primary", use_container_width=True):
            st.session_state.boq_rows = [x for x in st.session_state.boq_rows if x["dsr_code"] != "STEEL-FE500D"]
            st.session_state.boq_rows.append({
                "dsr_code": "STEEL-FE500D",
                "category": "Civil - RCC Reinforcement",
                "description": "Providing and placing TMT Fe-500D reinforcement steel complete as per drawing (IS 1786 / IS 2502).",
                "unit": "kg",
                "rate": float(steel_rate),
                "calc_mode": "LENGTH",
                "is_ref": "IS 1200 (Part 2) / IS 2502",
                "label": "Total Reinforcement Steel (BBS)",
                "nos": 1.0,
                "length": round(tot_kg, 2),
                "breadth": 1.0,
                "depth": 1.0,
            })
            st.success(f"Pushed {tot_kg:,.2f} kg to BOQ Tab!")
        if c_clr.button("Clear Steel List", use_container_width=True):
            st.session_state.steel_rows = []
            st.rerun()

    st.divider()
    st.subheader("3. Detailing Reference Figures")
    g1, g2, g3 = st.columns(3)
    if os.path.exists(diags["DETAIL_CRANK"]):
        g1.image(diags["DETAIL_CRANK"], caption="1. Cranked Bar Arrangement (SC-41)", width=360)
    if os.path.exists(diags["DETAIL_STIRRUP"]):
        g2.image(diags["DETAIL_STIRRUP"], caption="2. Closed Stirrup 135° Hook Detail (SC-51)", width=360)
    if os.path.exists(diags["DETAIL_EXTRA"]):
        g3.image(diags["DETAIL_EXTRA"], caption="3. Extra Support Bars Arrangement", width=360)

# ===========================================================================
# TAB 4: SMART AUTOCAD -> COMPLETE 12-TAB MODULAR BOQ (EVERY PART SEPARATE)
# ===========================================================================
with tab4:
    st.subheader("🏗️ Smart AutoCAD → Modular 12-Tab BOQ (Every Building Part Separate)")
    st.caption(
        "Upload a `.DXF` drawing OR the `AutoBOQ_Takeoff.csv` exported from AutoCAD using our free "
        "`AutoBOQ_SmartTools.lsp` plugin — just type **`AD`** in AutoCAD to open the 10-part pop-up menu "
        "(Footing, Column, Plinth Beam, Beam, Slab, Staircase, Water Tank, Wall, Room) and draw native BBS & BOQ tables!"
    )

    c_lsp1, c_lsp2 = st.columns([2, 1])
    with c_lsp1:
        cad_up = st.file_uploader(
            "Upload AutoCAD Drawing (.DXF) or LISP Takeoff (.CSV)",
            type=["dxf", "csv"],
            key="mod_cad_upload",
        )
    with c_lsp2:
        lsp_path = os.path.join(BASE_DIR, "autocad", "AutoBOQ_SmartTools.lsp")
        if os.path.exists(lsp_path):
            with open(lsp_path, "rb") as lf:
                st.download_button(
                    "⬇️ Download AutoCAD LISP Plugin v2 (.LSP)",
                    data=lf.read(),
                    file_name="AutoBOQ_SmartTools.lsp",
                    mime="text/plain",
                    use_container_width=True,
                    key="dl_autocad_lsp",
                )
        st.caption("Load via `APPLOAD` → type **`AD`** (10 parts + native AutoCAD BBS & BOQ Table + CSV export).")

    if cad_up is not None:
        try:
            parsed_rows = parse_autocad_takeoff(cad_up.getvalue(), cad_up.name)
            if parsed_rows:
                st.success(f"Extracted {len(parsed_rows)} closed elements from `{cad_up.name}`!")
                mod_raw_rows = parsed_rows
            else:
                st.warning("No closed polylines detected in uploaded file; showing default sample takeoff.")
                mod_raw_rows = default_autocad_sample_rows()
        except Exception as exc:
            st.error(f"Could not parse `{cad_up.name}`: {exc}")
            mod_raw_rows = default_autocad_sample_rows()
    else:
        mod_raw_rows = default_autocad_sample_rows()

    with st.expander("📐 AutoCAD Raw Polyline Takeoff & Room Schedule", expanded=False):
        st.dataframe(mod_raw_rows, use_container_width=True, hide_index=True)

    st.markdown("#### Project & Commercial Parameters (Live Recompile)")
    mp1, mp2, mp3, mp4, mp5, mp6, mp7 = st.columns(7)
    m_story_ht = mp1.number_input("Story Ht (m)", value=3.30, min_value=2.4, max_value=6.0, step=0.1, key="mod_ht")
    m_slab_thk = mp2.number_input("Slab Thk (m)", value=0.125, min_value=0.10, max_value=0.30, step=0.005, format="%.3f", key="mod_sthk")
    m_conc_rt = mp3.number_input("M25 (₹/m³)", value=7600.0, step=100.0, key="mod_conc")
    m_stl_rt = mp4.number_input("Steel (₹/kg)", value=82.0, step=1.0, key="mod_stl")
    m_cont = mp5.number_input("Contingency %", value=5.0, step=0.5, key="mod_cont")
    m_ohp = mp6.number_input("OHP %", value=10.0, step=0.5, key="mod_ohp")
    m_gst = mp7.number_input("GST %", value=18.0, min_value=0.0, max_value=28.0, step=1.0, key="mod_gst")

    mod_params = {
        "story_height_m": float(m_story_ht),
        "slab_thickness_m": float(m_slab_thk),
        "conc_rate_inr": float(m_conc_rt),
        "steel_rate_inr": float(m_stl_rt),
        "contingency_pct": float(m_cont) / 100.0,
        "ohp_pct": float(m_ohp) / 100.0,
        "gst_pct": float(m_gst) / 100.0,
    }
    mod_prev = get_modular_boq_preview(mod_raw_rows, mod_params)
    mod_files = compile_modular_boq_bytes(mod_raw_rows, mod_params)
    k = mod_prev["kpis"]

    # Executive Scope Breakdown Cards (Matching Quick Building Estimator PRO)
    st.markdown("#### 📊 Commercial & Scope Breakdown")
    sk1, sk2, sk3, sk4, sk5 = st.columns(5)
    sk1.metric("Materials (Supply Only)", f"₹{k['materials_supply']:,.0f}", "22 Supply Items")
    sk2.metric("Labour & Works", f"₹{k['labour_works']:,.0f}", "All Civil/Finish/MEP")
    sk3.metric("Doors & Windows", f"₹{k['doors_windows']:,.0f}", "Joinery Schedule")
    sk4.metric("Additional Works", f"₹{k['additional_works']:,.0f}", "Sump/Septic/Wall/RWH")
    sk5.metric("Grand Total (incl. GST)", f"₹{k['grand_total']:,.0f}", f"₹{k['rate_per_sqft']:,.0f} / Sq.ft")

    # Scope & Trade Filterable Itemized BOQ + Show Formulas Checkbox
    st.markdown("#### 🔍 Itemized BOQ Explorer (Filter by Scope / Trade)")
    fc1, fc2, fc3 = st.columns([1.4, 1.4, 1])
    scope_sel = fc1.selectbox(
        "Primary Scope Filter",
        ["All Items", "Materials Only", "Labour & Works", "Doors & Windows", "Additional Works"],
        key="mod_scope_flt",
    )
    all_trades = sorted({r["Trade"] for r in mod_prev["items"]})
    trade_sel = fc2.selectbox("Filter by Trade", ["All Trades"] + all_trades, key="mod_trade_flt")
    show_fml = fc3.checkbox("Show First-Principles Formulas", value=True, key="mod_show_fml")

    filt_items = mod_prev["items"]
    if scope_sel != "All Items":
        filt_items = [r for r in filt_items if r["Scope"] == scope_sel]
    if trade_sel != "All Trades":
        filt_items = [r for r in filt_items if r["Trade"] == trade_sel]

    disp_rows = []
    for r in filt_items:
        row_obj = {
            "Code": r["Code"],
            "Trade": r["Trade"],
            "Description": r["Description"],
            "Qty": r["Qty"],
            "Unit": r["Unit"],
            "Rate (₹)": r["Rate (₹)"],
            "Amount (₹)": r["Amount (₹)"],
        }
        if show_fml:
            row_obj["First-Principles Formula"] = r["Formula"]
        disp_rows.append(row_obj)
    st.dataframe(disp_rows, use_container_width=True, hide_index=True)

    st.divider()
    st.markdown("#### 📥 Modular BOQ Download Center (Complete 12-Tab Master, Reports & Individual Trade Files)")
    st.download_button(
        "📦 Download Complete 12-Tab Master BOQ Workbook (All Parts + 22 Materials + Client BOQ + Bank Loan)",
        data=mod_files[MODULAR_BOQ_NAME],
        file_name=MODULAR_BOQ_NAME,
        mime=_xlsx_mime,
        type="primary",
        use_container_width=True,
        key="mod_dl_master",
    )
    r1, r2, r3 = st.columns(3)
    r1.download_button("🧱 Materials Takeoff Only (22 Items)", data=mod_files["BOQ_Materials_Takeoff_Only.xlsx"], file_name="BOQ_Materials_Takeoff_Only.xlsx", mime=_xlsx_mime, use_container_width=True, key="mod_dl_mat")
    r2.download_button("💼 Client Quotation BOQ (incl. OHP)", data=mod_files["BOQ_Client_Quotation.xlsx"], file_name="BOQ_Client_Quotation.xlsx", mime=_xlsx_mime, use_container_width=True, key="mod_dl_cli")
    r3.download_button("🏦 Bank Loan Estimate (Stages)", data=mod_files["BOQ_Bank_Loan_Estimate.xlsx"], file_name="BOQ_Bank_Loan_Estimate.xlsx", mime=_xlsx_mime, use_container_width=True, key="mod_dl_bnk")

    mc1, mc2, mc3, mc4 = st.columns(4)
    mc1.download_button("⬇️ Foundation & Basement", data=mod_files["BOQ_Foundation_Earthwork.xlsx"], file_name="BOQ_Foundation_Earthwork.xlsx", mime=_xlsx_mime, use_container_width=True, key="mod_dl_fdn")
    mc2.download_button("⬇️ RCC Structure & Steel", data=mod_files["BOQ_RCC_Structure.xlsx"], file_name="BOQ_RCC_Structure.xlsx", mime=_xlsx_mime, use_container_width=True, key="mod_dl_rcc")
    mc3.download_button("⬇️ Masonry & Plastering", data=mod_files["BOQ_Masonry_Plaster.xlsx"], file_name="BOQ_Masonry_Plaster.xlsx", mime=_xlsx_mime, use_container_width=True, key="mod_dl_mas")
    mc4.download_button("⬇️ Doors & Windows", data=mod_files["BOQ_Doors_Windows.xlsx"], file_name="BOQ_Doors_Windows.xlsx", mime=_xlsx_mime, use_container_width=True, key="mod_dl_dw")

    mc5, mc6, mc7 = st.columns(3)
    mc5.download_button("⬇️ Flooring, Granite & Paint", data=mod_files["BOQ_Flooring_Paint.xlsx"], file_name="BOQ_Flooring_Paint.xlsx", mime=_xlsx_mime, use_container_width=True, key="mod_dl_flr")
    mc6.download_button("⬇️ Electrical & Plumbing", data=mod_files["BOQ_Electrical_Plumbing.xlsx"], file_name="BOQ_Electrical_Plumbing.xlsx", mime=_xlsx_mime, use_container_width=True, key="mod_dl_mep")
    mc7.download_button("⬇️ Additional Works (Sump/Wall)", data=mod_files["BOQ_Additional_Works.xlsx"], file_name="BOQ_Additional_Works.xlsx", mime=_xlsx_mime, use_container_width=True, key="mod_dl_add")
