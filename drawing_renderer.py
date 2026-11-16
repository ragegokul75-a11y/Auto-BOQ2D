"""
Faithful 2D drawing renderer for DXF / DWG / PDF.

Uses ezdxf's official drawing frontend (the same engine used to plot DXF to PDF/PNG),
so blocks, nested blocks, arcs, circles, ellipses, splines, hatches, dimensions,
leaders, layer on/off/frozen and linetypes are all rendered the way CAD shows them.

Output is plain JSON geometry in drawing units that the browser draws on a canvas:
    strokes  : polylines grouped by colour
    fills    : filled polygons (solid hatches / solids) grouped by colour
    texts    : text strings with position, height, rotation (for display + click-pick)
    snaps    : endpoint / vertex snap points
No automatic "guessing" of BOQ items is done here - the user measures on the true drawing.
"""

from __future__ import annotations

import base64
import math
import os
import shutil
from typing import Any, Dict, List, Tuple

import ezdxf
from ezdxf import recover
from ezdxf.addons.drawing import Frontend, RenderContext
from ezdxf.addons.drawing import recorder
from ezdxf.addons.drawing.config import Configuration, TextPolicy, BackgroundPolicy
from ezdxf.disassemble import recursive_decompose

try:
    import pymupdf as fitz  # PyMuPDF >= 1.24
except ImportError:  # pragma: no cover
    import fitz  # type: ignore

INSUNITS_TO_M = {1: 0.0254, 2: 0.3048, 4: 0.001, 5: 0.01, 6: 1.0, 14: 0.1}
INSUNITS_NAME = {1: "inch", 2: "feet", 4: "mm", 5: "cm", 6: "m", 14: "dm"}

MAX_SNAPS = 80000


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def _r(v: float) -> float:
    return round(float(v), 4)


def _hex(color: str) -> str:
    c = (color or "#000000")[:7]
    return c if c.startswith("#") and len(c) == 7 else "#000000"


class _Collector:
    def __init__(self) -> None:
        self.strokes: Dict[str, List[List[float]]] = {}
        self.fills: Dict[str, List[List[float]]] = {}
        self.snap_set: set = set()
        self.snaps: List[float] = []
        self.minx = self.miny = math.inf
        self.maxx = self.maxy = -math.inf

    def _ext(self, x: float, y: float) -> None:
        if x < self.minx:
            self.minx = x
        if x > self.maxx:
            self.maxx = x
        if y < self.miny:
            self.miny = y
        if y > self.maxy:
            self.maxy = y

    def snap(self, x: float, y: float) -> None:
        if len(self.snaps) >= MAX_SNAPS * 2:
            return
        key = (round(x, 3), round(y, 3))
        if key in self.snap_set:
            return
        self.snap_set.add(key)
        self.snaps.extend((_r(x), _r(y)))

    def add_stroke(self, color: str, pts: List[Tuple[float, float]], snap_all: bool) -> None:
        if len(pts) < 2:
            return
        flat: List[float] = []
        for x, y in pts:
            if not (math.isfinite(x) and math.isfinite(y)):
                return
            flat.append(_r(x))
            flat.append(_r(y))
            self._ext(x, y)
        self.strokes.setdefault(color, []).append(flat)
        if snap_all:
            for x, y in pts:
                self.snap(x, y)
        else:
            self.snap(*pts[0])
            self.snap(*pts[-1])

    def add_fill(self, color: str, pts: List[Tuple[float, float]]) -> None:
        if len(pts) < 3:
            return
        flat: List[float] = []
        for x, y in pts:
            if not (math.isfinite(x) and math.isfinite(y)):
                return
            flat.append(_r(x))
            flat.append(_r(y))
            self._ext(x, y)
        self.fills.setdefault(color, []).append(flat)


def _read_dxf(path: str):
    """Reads DXF with ezdxf's recover module so slightly damaged files still open."""
    try:
        doc, auditor = recover.readfile(path)
    except IOError as exc:
        raise ValueError(f"File is not a readable DXF: {exc}") from exc
    except ezdxf.DXFStructureError as exc:
        raise ValueError(f"DXF file structure is damaged beyond repair: {exc}") from exc
    return doc


def _choose_layout(doc):
    msp = doc.modelspace()
    if len(msp) > 0:
        return msp, "Model"
    for name in doc.layout_names_in_taborder():
        if name.lower() == "model":
            continue
        lay = doc.layouts.get(name)
        if len(lay) > 0:
            return lay, name
    return msp, "Model"


def _collect_texts(layout, flat_dist: float) -> List[Dict[str, Any]]:
    texts: List[Dict[str, Any]] = []
    try:
        entities = recursive_decompose(layout)
    except Exception:
        entities = iter(layout)
    for e in entities:
        try:
            et = e.dxftype()
            if et in ("TEXT", "ATTRIB"):
                s = e.plain_text() if hasattr(e, "plain_text") else e.dxf.text
                h = float(e.dxf.get("height", 2.5))
                rot = float(e.dxf.get("rotation", 0.0))
                halign = e.dxf.get("halign", 0)
                valign = e.dxf.get("valign", 0)
                p = e.dxf.align_point if (halign or valign) and e.dxf.hasattr("align_point") else e.dxf.insert
                anchor = "l"
                if halign in (1, 4):
                    anchor = "c"
                elif halign == 2:
                    anchor = "r"
                x, y = float(p.x), float(p.y)
                base = "b" if valign in (0, 1) else ("m" if valign == 2 else "t")
            elif et == "MTEXT":
                s = e.plain_text()
                h = float(e.dxf.get("char_height", 2.5))
                rot = float(e.get_rotation())
                p = e.dxf.insert
                x, y = float(p.x), float(p.y)
                ap = int(e.dxf.get("attachment_point", 1))
                anchor = {0: "l", 1: "c", 2: "r"}[(ap - 1) % 3]
                base = {0: "t", 1: "m", 2: "b"}[(ap - 1) // 3]
            else:
                continue
            s = (s or "").strip()
            if not s or not math.isfinite(x) or not math.isfinite(y):
                continue
            for i, line in enumerate(s.split("\n")[:12]):
                line = line.strip()
                if not line:
                    continue
                ly = y - i * h * 1.6 if base != "b" else y + (len(s.split("\n")) - 1 - i) * h * 1.6
                texts.append({
                    "t": line[:120], "x": _r(x), "y": _r(ly), "h": _r(h),
                    "r": round(rot, 2), "a": anchor, "b": base,
                    "layer": e.dxf.get("layer", "0"),
                })
        except Exception:
            continue
        if len(texts) > 20000:
            break
    return texts


# --------------------------------------------------------------------------- #
# DXF
# --------------------------------------------------------------------------- #
def render_dxf(path: str) -> Dict[str, Any]:
    doc = _read_dxf(path)
    layout, layout_name = _choose_layout(doc)

    # Rough extents first to pick a sensible curve flattening tolerance
    ctx = RenderContext(doc)
    rec = recorder.Recorder()
    cfg = Configuration(text_policy=TextPolicy.IGNORE, background_policy=BackgroundPolicy.WHITE)
    Frontend(ctx, rec, config=cfg).draw_layout(layout, finalize=True)
    player = rec.player()

    bbox = player.bbox()
    if bbox.has_data:
        span = max(bbox.size.x, bbox.size.y, 1e-6)
    else:
        span = 1000.0
    flat_dist = span / 6000.0

    col = _Collector()
    for record, props in player.recordings():
        color = _hex(props.color)
        if isinstance(record, recorder.PathRecord):
            try:
                pts = [(v.x, v.y) for v in record.path.flattening(flat_dist)]
            except Exception:
                continue
            col.add_stroke(color, pts, snap_all=len(pts) <= 6)
        elif isinstance(record, recorder.SolidLinesRecord):
            v = list(record.lines.vertices())
            for i in range(0, len(v) - 1, 2):
                col.add_stroke(color, [(v[i].x, v[i].y), (v[i + 1].x, v[i + 1].y)], snap_all=True)
        elif isinstance(record, recorder.PointsRecord):
            v = [(p.x, p.y) for p in record.points.vertices()]
            if len(v) == 2:
                col.add_stroke(color, v, snap_all=True)
            elif len(v) > 2:
                col.add_fill(color, v)
        elif isinstance(record, recorder.FilledPathsRecord):
            for path_ in record.paths:
                try:
                    pts = [(v.x, v.y) for v in path_.flattening(flat_dist)]
                except Exception:
                    continue
                col.add_fill(color, pts)

    texts = _collect_texts(layout, flat_dist)
    for t in texts:
        col._ext(t["x"], t["y"])

    if col.minx == math.inf:
        raise ValueError("The drawing opened but contains no visible geometry in Model space or any Layout.")

    insunits = int(doc.header.get("$INSUNITS", 0) or 0)
    if insunits in INSUNITS_TO_M:
        unit_to_m = INSUNITS_TO_M[insunits]
        unit_name = INSUNITS_NAME[insunits]
        unit_source = "header ($INSUNITS)"
    else:
        span_w = max(col.maxx - col.minx, col.maxy - col.miny)
        if span_w > 2500:
            unit_to_m, unit_name = 0.001, "mm"
        else:
            unit_to_m, unit_name = 1.0, "m"
        unit_source = "guessed from drawing size - please confirm"

    layers = sorted({p[2] for p in player.properties.values()} if hasattr(player, "properties") else [])

    return {
        "format": "DXF",
        "layout": layout_name,
        "unit_to_m": unit_to_m,
        "unit_name": unit_name,
        "unit_source": unit_source,
        "bounds": {"minx": _r(col.minx), "miny": _r(col.miny), "maxx": _r(col.maxx), "maxy": _r(col.maxy)},
        "strokes": [{"c": c, "p": p} for c, p in col.strokes.items()],
        "fills": [{"c": c, "p": p} for c, p in col.fills.items()],
        "texts": texts,
        "snaps": col.snaps,
        "layers": layers,
        "image": None,
        "stats": {
            "polylines": sum(len(p) for p in col.strokes.values()),
            "fills": sum(len(p) for p in col.fills.values()),
            "texts": len(texts),
            "snaps": len(col.snaps) // 2,
        },
    }


# --------------------------------------------------------------------------- #
# DWG (via free ODA File Converter, if installed)
# --------------------------------------------------------------------------- #
def _find_oda() -> str | None:
    found = shutil.which("ODAFileConverter")
    if found:
        return found
    for root in (r"C:\Program Files\ODA", r"C:\Program Files (x86)\ODA"):
        if os.path.isdir(root):
            for dp, _, fns in os.walk(root):
                for fn in fns:
                    if fn.lower() == "odafileconverter.exe":
                        return os.path.join(dp, fn)
    return None


def render_dwg(path: str) -> Dict[str, Any]:
    exe = _find_oda()
    if not exe:
        raise ValueError(
            "DWG files need the free 'ODA File Converter' to be installed on this PC "
            "(https://www.opendesign.com/guestfiles/oda_file_converter). "
            "Quick alternative: in AutoCAD type SAVEAS and choose 'AutoCAD DXF (*.dxf)', then upload the DXF."
        )
    from ezdxf.addons import odafc
    ezdxf.options.set("odafc-addon", "win_exec_path", exe)
    out_dir = os.path.dirname(path)
    dxf_path = os.path.splitext(path)[0] + "_conv.dxf"
    try:
        odafc.convert(path, dxf_path, version="R2018", replace=True)
    except Exception as exc:
        raise ValueError(f"ODA File Converter could not convert this DWG: {exc}") from exc
    try:
        data = render_dxf(dxf_path)
        data["format"] = "DWG"
        return data
    finally:
        if os.path.exists(dxf_path):
            os.remove(dxf_path)


# --------------------------------------------------------------------------- #
# PDF
# --------------------------------------------------------------------------- #
def render_pdf(path: str, page_no: int = 0) -> Dict[str, Any]:
    try:
        doc = fitz.open(path)
    except Exception as exc:
        raise ValueError(f"File is not a readable PDF: {exc}") from exc
    if doc.page_count == 0:
        raise ValueError("PDF has no pages.")
    page_no = max(0, min(page_no, doc.page_count - 1))
    page = doc[page_no]
    pw, ph = float(page.rect.width), float(page.rect.height)
    x0, y0 = float(page.rect.x0), float(page.rect.y0)

    # Render the true page image (longest side ~ 4000 px, capped for memory)
    zoom = min(4000.0 / max(pw, ph), 6.0)
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    img_b64 = "data:image/png;base64," + base64.b64encode(pix.tobytes("png")).decode("ascii")

    def W(px: float, py: float) -> Tuple[float, float]:
        # PDF y runs downward -> convert to y-up world coordinates in PDF points
        return px - x0, ph - (py - y0)

    col = _Collector()
    try:
        drawings = page.get_drawings()
    except Exception:
        drawings = []
    for d in drawings:
        for it in d.get("items", []):
            op = it[0]
            try:
                if op == "l":
                    col.snap(*W(it[1].x, it[1].y))
                    col.snap(*W(it[2].x, it[2].y))
                elif op == "re":
                    r = it[1]
                    if abs(r.width - pw) < 5 and abs(r.height - ph) < 5:
                        continue
                    for px, py in ((r.x0, r.y0), (r.x1, r.y0), (r.x1, r.y1), (r.x0, r.y1)):
                        col.snap(*W(px, py))
                elif op == "qu":
                    q = it[1]
                    for p in (q.ul, q.ur, q.lr, q.ll):
                        col.snap(*W(p.x, p.y))
                elif op == "c":
                    col.snap(*W(it[1].x, it[1].y))
                    col.snap(*W(it[4].x, it[4].y))
            except Exception:
                continue

    texts: List[Dict[str, Any]] = []
    try:
        tdict = page.get_text("dict")
        for blk in tdict.get("blocks", []):
            for ln in blk.get("lines", []):
                s = "".join(sp.get("text", "") for sp in ln.get("spans", [])).strip()
                if not s:
                    continue
                bx0, by0, bx1, by1 = ln["bbox"]
                dx, dy = ln.get("dir", (1, 0))
                rot = -math.degrees(math.atan2(dy, dx))
                wx, wy = W(bx0, by1)
                texts.append({"t": s[:120], "x": _r(wx), "y": _r(wy), "h": _r(max(by1 - by0, 1)),
                              "r": round(rot, 2), "a": "l", "b": "b", "layer": "PDF"})
    except Exception:
        pass

    page_count = doc.page_count
    doc.close()
    return {
        "format": "PDF",
        "layout": f"Page {page_no + 1} of {page_count}",
        "page": page_no,
        "page_count": page_count,
        # PDF has no real-world units: 1 pt = 0.3528 mm on paper. Default assumes 1:100.
        "unit_to_m": 0.0254 / 72.0 * 100.0,
        "unit_name": "pt (1:100 assumed)",
        "unit_source": "PDF has no real scale - choose the drawing scale or use Calibrate",
        "bounds": {"minx": 0.0, "miny": 0.0, "maxx": _r(pw), "maxy": _r(ph)},
        "strokes": [],
        "fills": [],
        "texts": texts,
        "snaps": col.snaps,
        "layers": [],
        "image": {"src": img_b64, "minx": 0.0, "miny": 0.0, "maxx": _r(pw), "maxy": _r(ph)},
        "stats": {"polylines": 0, "fills": 0, "texts": len(texts), "snaps": len(col.snaps) // 2},
    }


def render_any(path: str, ext: str, page_no: int = 0) -> Dict[str, Any]:
    ext = ext.lower()
    if ext == ".dxf":
        return render_dxf(path)
    if ext == ".dwg":
        return render_dwg(path)
    if ext == ".pdf":
        return render_pdf(path, page_no)
    raise ValueError("Unsupported file type. Upload a .DXF, .DWG or .PDF drawing.")
