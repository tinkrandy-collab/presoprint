#!/usr/bin/env python3
"""
PDF Print Production – Web UI.

Run:
    python app.py

Then open http://localhost:5001 in your browser.
"""

import os
import re
import io
import uuid
import base64
import json
import tempfile
import time
import threading
import shutil
import subprocess
import urllib.request
import urllib.error
from pathlib import Path
from flask import Flask, request, jsonify, send_file, send_from_directory

import pikepdf
from pikepdf import Pdf, Page, Name, Array, Dictionary, String
import fitz  # pymupdf – for thumbnail rendering
import numpy as np
from PIL import Image, ImageFilter

# ── Constants ─────────────────────────────────────────────────────────────────
PTS_PER_INCH = 72
DEFAULT_TRIM_WIDTH_IN = 10.5   # fixed width
DEFAULT_BLEED_IN = 0.125

# Safe margins (inches)
DEFAULT_MARGIN_LEFT = 0.25
DEFAULT_MARGIN_RIGHT = 0.25
DEFAULT_MARGIN_TOP = 0.5
DEFAULT_MARGIN_BOTTOM = 0.5

MARK_OFFSET = 3
MARK_LENGTH = 18
WHITE_WEIGHT = 1.0
BLACK_WEIGHT = 0.75

THUMBNAIL_DPI = 96
HOME_LETTER_WIDTH_IN = 11.0
HOME_LETTER_HEIGHT_IN = 8.5
HOME_A4_WIDTH_IN = 11.6929
HOME_A4_HEIGHT_IN = 8.2677
HOME_MARGIN_TOP_BOTTOM_IN = 0.5
HOME_MARGIN_LEFT_RIGHT_IN = 0.25
HOME_CORNER_RADIUS_IN = 0.18
OUTPAINT_API_BASE = "https://api.replicate.com/v1"
OUTPAINT_DEFAULT_MODEL = "fermatresearch/sdxl-outpainting-lora"
OUTPAINT_DEFAULT_PROMPT = (
    "Outpaint only the masked border area to extend the existing background naturally. "
    "Match nearby colors, texture, and lighting. Keep continuation subtle and seamless."
)
OUTPAINT_FERMAT_PROMPT = (
    "Extend the existing background into the outpaint area only. "
    "Preserve gradients, photo texture, and curved shapes from adjacent edges. "
    "Do not add new objects or geometric forms."
)
OUTPAINT_DEFAULT_NEGATIVE_PROMPT = (
    "text, words, letters, numbers, logo, watermark, frame, border, shapes, "
    "objects, people, faces, symbols, artifacts, distortions, sharp seams, "
    "rectangles, boxes, repeated elements, duplicated objects, glow, halo"
)
OUTPAINT_MAX_SIDE_PX = 1024
ADOBE_FIREFLY_BASE = "https://firefly-api.adobe.io"

app = Flask(__name__, static_folder="static")

UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "pdf_print_uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

JOB_PROGRESS = {}
JOB_PROGRESS_LOCK = threading.Lock()
REPLICATE_CREATE_LOCK = threading.Lock()
REPLICATE_LAST_CREATE_TS = 0.0


class HttpApiError(RuntimeError):
    def __init__(self, status, url, detail):
        self.status = int(status)
        self.url = url
        self.detail = detail
        super().__init__(f"HTTP {self.status} from {self.url}: {self.detail}")


def _set_progress(job_id, **fields):
    with JOB_PROGRESS_LOCK:
        state = JOB_PROGRESS.get(job_id, {})
        state.update(fields)
        JOB_PROGRESS[job_id] = state
        return dict(state)


def _phase_profile(phase):
    profiles = {
        "queued": (0, 2),
        "ai_bleed": (2, 55),
        "processing": (57, 13),
        "verifying": (70, 20),
        "thumbnails": (90, 9),
        "finalizing": (99, 1),
        "complete": (100, 0),
        "error": (0, 0),
    }
    return profiles.get(phase, (0, 0))


def _compute_progress(state):
    phase = state.get("phase", "queued")
    current = int(state.get("current", 0) or 0)
    total = int(state.get("total", 0) or 0)
    phase_started_at = float(state.get("phase_started_at", time.time()))
    now = time.time()

    base, span = _phase_profile(phase)
    frac = min(1.0, max(0.0, (current / total))) if total > 0 else 0.0
    percent = 100.0 if phase == "complete" else (base + span * frac)

    eta_seconds = None
    if phase not in ("complete", "error") and total > 0 and current > 0 and current < total:
        elapsed_phase = max(0.1, now - phase_started_at)
        eta_seconds = int(round((elapsed_phase / current) * (total - current)))

    return int(round(percent)), eta_seconds


# ── Helpers ───────────────────────────────────────────────────────────────────

def get_page_dimensions(page):
    mbox = page.mediabox
    x0, y0, x1, y1 = [float(v) for v in mbox]
    return x0, y0, x1 - x0, y1 - y0


def _fitz_page_content_min_dpi(page, min_area_in2=0.04):
    """Estimate minimum effective DPI of raster images on a source fitz page."""
    dpis = []
    seen = set()
    try:
        images = page.get_images(full=True)
    except Exception:
        images = []
    for info in images:
        if len(info) < 4:
            continue
        xref = int(info[0])
        px_w = float(info[2] or 0)
        px_h = float(info[3] or 0)
        if xref <= 0 or px_w <= 0 or px_h <= 0:
            continue
        try:
            rects = page.get_image_rects(xref)
        except Exception:
            rects = []
        for rect in rects:
            w_pt = float(rect.width)
            h_pt = float(rect.height)
            if w_pt <= 0 or h_pt <= 0:
                continue
            w_in = w_pt / 72.0
            h_in = h_pt / 72.0
            if (w_in * h_in) < min_area_in2:
                continue
            dpi_x = px_w / max(1e-6, w_in)
            dpi_y = px_h / max(1e-6, h_in)
            eff_dpi = float(min(dpi_x, dpi_y))
            sig = (xref, round(w_in, 4), round(h_in, 4))
            if sig in seen:
                continue
            seen.add(sig)
            dpis.append(eff_dpi)
    if not dpis:
        return None
    return float(min(dpis))


def _upscale_source_image_xref(doc, xref, scale=2):
    """Upscale a single image object in a fitz document and replace by xref."""
    try:
        pix = fitz.Pixmap(doc, int(xref))
        if pix.alpha or pix.n > 4:
            pix = fitz.Pixmap(fitz.csRGB, pix)
        src = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
    except Exception:
        return False
    up = _run_realesrgan_ncnn(src, max(2, int(scale)))
    if up is None:
        up = src.resize((src.width * 2, src.height * 2), Image.Resampling.LANCZOS)
    out = io.BytesIO()
    up.save(out, format="PNG")
    try:
        # Replaces this image object everywhere it is referenced.
        doc[0].replace_image(int(xref), stream=out.getvalue())
    except Exception:
        # Fallback: try using any page handle from doc.
        try:
            for p in doc:
                p.replace_image(int(xref), stream=out.getvalue())
                break
        except Exception:
            return False
    return True


def detect_background_color(page):
    """Best-effort detection of a page's background fill colour."""
    try:
        obj = page.obj
        if Name.Contents not in obj:
            return (1.0, 1.0, 1.0)
        contents = obj[Name.Contents]
        if isinstance(contents, pikepdf.Array):
            raw = b""
            for ref in contents:
                raw += ref.read_bytes() + b"\n"
        else:
            raw = contents.read_bytes()
        text = raw.decode("latin-1", errors="replace")
        early = text[:3000]

        m = re.search(r"([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+rg", early)
        if m:
            return (float(m.group(1)), float(m.group(2)), float(m.group(3)))
        m = re.search(r"([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+scn?", early)
        if m:
            return (float(m.group(1)), float(m.group(2)), float(m.group(3)))
        m = re.search(r"([\d.]+)\s+g\b", early)
        if m:
            v = float(m.group(1))
            return (v, v, v)
    except Exception:
        pass
    return (1.0, 1.0, 1.0)


def _fitz_page_to_pil(page, target_w_px, target_h_px):
    """Render a fitz page to RGB PIL image at target size."""
    rect = page.rect
    src_w_pt = max(1.0, float(rect.width))
    src_h_pt = max(1.0, float(rect.height))
    zoom_x = max(0.1, target_w_px / src_w_pt)
    zoom_y = max(0.1, target_h_px / src_h_pt)
    matrix = fitz.Matrix(zoom_x, zoom_y)
    pix = page.get_pixmap(matrix=matrix, alpha=False)
    return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


def _round_to_mult(value, mult=64):
    return max(mult, int(round(value / mult) * mult))


def _image_to_data_url(img, fmt="PNG"):
    out = io.BytesIO()
    img.save(out, format=fmt)
    encoded = base64.b64encode(out.getvalue()).decode("ascii")
    mime = "image/png" if fmt.upper() == "PNG" else "image/jpeg"
    return f"data:{mime};base64,{encoded}"


def _http_json(method, url, *, headers=None, payload=None, timeout=120):
    body = None
    req_headers = {
        "User-Agent": "PrintPreso/0.3 (+local)",
        "Accept": "application/json",
    }
    req_headers.update(dict(headers or {}))
    if payload is not None:
        body = json.dumps(payload).encode("utf-8")
        req_headers["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=body, headers=req_headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", errors="ignore")[:800]
        raise HttpApiError(e.code, url, detail) from e


def _http_bytes(url, *, timeout=120):
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "PrintPreso/0.3 (+local)", "Accept": "*/*"},
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


def _http_json_raw(method, url, *, headers=None, payload=None, timeout=120):
    body = None
    req_headers = {"User-Agent": "PrintPreso/0.3 (+local)", "Accept": "application/json"}
    req_headers.update(dict(headers or {}))
    if payload is not None:
        body = payload if isinstance(payload, (bytes, bytearray)) else json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=body, headers=req_headers, method=method)
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        raw = resp.read()
        ctype = (resp.headers.get("Content-Type") or "").lower()
        if "application/json" in ctype:
            return json.loads(raw.decode("utf-8"))
        return raw


def _bleed_quality_ok(out_img, x0, y0, x1, y1):
    """Heuristic guard: reject outpainted bleed that diverges too much from border context."""
    arr = np.array(out_img).astype(np.float32)
    h, w = arr.shape[:2]
    band = max(6, min(24, int(round(min(w, h) * 0.02))))

    def mean_color(region):
        if region.size == 0:
            return None
        return np.mean(region.reshape(-1, 3), axis=0)

    def edge_score(region):
        if region.size == 0:
            return 0.0
        gx = np.abs(np.diff(region, axis=1)).mean()
        gy = np.abs(np.diff(region, axis=0)).mean()
        return float(gx + gy)

    checks = []
    # left
    if x0 > band:
        bleed = arr[y0:y1, :x0, :]
        inner = arr[y0:y1, x0:min(x0 + band, w), :]
        checks.append((bleed, inner))
    # right
    if x1 < w - band:
        bleed = arr[y0:y1, x1:w, :]
        inner = arr[y0:y1, max(0, x1 - band):x1, :]
        checks.append((bleed, inner))
    # top
    if y0 > band:
        bleed = arr[:y0, x0:x1, :]
        inner = arr[y0:min(y0 + band, h), x0:x1, :]
        checks.append((bleed, inner))
    # bottom
    if y1 < h - band:
        bleed = arr[y1:h, x0:x1, :]
        inner = arr[max(0, y1 - band):y1, x0:x1, :]
        checks.append((bleed, inner))

    if not checks:
        return True

    color_delta_max = float(os.getenv("BLEED_GUARD_COLOR_DELTA_MAX", "50"))
    edge_ratio_max = float(os.getenv("BLEED_GUARD_EDGE_RATIO_MAX", "2.6"))
    fail_votes = 0
    total = 0
    for bleed, inner in checks:
        c_bleed = mean_color(bleed)
        c_inner = mean_color(inner)
        if c_bleed is None or c_inner is None:
            continue
        total += 1
        color_delta = float(np.linalg.norm(c_bleed - c_inner))
        e_bleed = edge_score(bleed)
        e_inner = max(1e-3, edge_score(inner))
        edge_ratio = e_bleed / e_inner
        if color_delta > color_delta_max or edge_ratio > edge_ratio_max:
            fail_votes += 1
    if total == 0:
        return True
    # If most bleed sides look inconsistent, fail and fallback.
    return fail_votes <= max(0, total - 2)


def _score_bleed_candidate(img, seed_img, x0, y0, x1, y1):
    """Lower score is better: closer to seed near trim and lower edge noise in bleed."""
    arr = np.array(img).astype(np.float32)
    seed = np.array(seed_img).astype(np.float32)
    h, w = arr.shape[:2]
    bw = max(2, min(14, int(round(min(w, h) * 0.015))))
    score = 0.0

    def edge_energy(region):
        if region.size == 0:
            return 0.0
        gx = np.abs(np.diff(region, axis=1)).mean() if region.shape[1] > 1 else 0.0
        gy = np.abs(np.diff(region, axis=0)).mean() if region.shape[0] > 1 else 0.0
        return float(gx + gy)

    # Seam consistency at trim edges: output should stay close to seed right outside trim.
    if x0 > 0:
        a = arr[y0:y1, max(0, x0 - bw):x0, :]
        s = seed[y0:y1, max(0, x0 - bw):x0, :]
        score += float(np.mean(np.abs(a - s)))
    if x1 < w:
        a = arr[y0:y1, x1:min(w, x1 + bw), :]
        s = seed[y0:y1, x1:min(w, x1 + bw), :]
        score += float(np.mean(np.abs(a - s)))
    if y0 > 0:
        a = arr[max(0, y0 - bw):y0, x0:x1, :]
        s = seed[max(0, y0 - bw):y0, x0:x1, :]
        score += float(np.mean(np.abs(a - s)))
    if y1 < h:
        a = arr[y1:min(h, y1 + bw), x0:x1, :]
        s = seed[y1:min(h, y1 + bw), x0:x1, :]
        score += float(np.mean(np.abs(a - s)))

    # Penalize high-frequency artifacts in bleed regions.
    if x0 > 0:
        score += 0.7 * edge_energy(arr[y0:y1, :x0, :])
    if x1 < w:
        score += 0.7 * edge_energy(arr[y0:y1, x1:w, :])
    if y0 > 0:
        score += 0.7 * edge_energy(arr[:y0, x0:x1, :])
    if y1 < h:
        score += 0.7 * edge_energy(arr[y1:h, x0:x1, :])
    return score


def _suppress_bleed_artifacts(img, x0, y0, x1, y1):
    """Reduce hallucinated text/lines in bleed bands using adaptive edge replacement."""
    arr = np.array(img).astype(np.float32)
    h, w = arr.shape[:2]

    def energy(region):
        if region.size == 0:
            return 0.0
        gx = np.abs(np.diff(region, axis=1)).mean() if region.shape[1] > 1 else 0.0
        gy = np.abs(np.diff(region, axis=0)).mean() if region.shape[0] > 1 else 0.0
        return float(gx + gy)

    def side_detect_and_fix(side):
        nonlocal arr
        band = max(6, min(24, int(round(min(w, h) * 0.02))))
        ratio_thr = float(os.getenv("BLEED_ARTIFACT_RATIO", "1.35"))
        abs_thr = float(os.getenv("BLEED_ARTIFACT_ABS", "10.5"))

        if side == "left" and x0 > 0:
            bw = x0
            bleed = arr[y0:y1, :bw, :]
            inner = arr[y0:y1, x0:min(w, x0 + band), :]
            if inner.size == 0:
                return
            if energy(bleed) > max(abs_thr, energy(inner) * ratio_thr):
                anchor = arr[y0:y1, x0:x0 + 1, :]
                fill = np.repeat(anchor, bw, axis=1)
                arr[y0:y1, :bw, :] = 0.75 * fill + 0.25 * bleed
        elif side == "right" and x1 < w:
            bw = w - x1
            bleed = arr[y0:y1, x1:w, :]
            inner = arr[y0:y1, max(0, x1 - band):x1, :]
            if inner.size == 0:
                return
            if energy(bleed) > max(abs_thr, energy(inner) * ratio_thr):
                anchor = arr[y0:y1, x1 - 1:x1, :]
                fill = np.repeat(anchor, bw, axis=1)
                arr[y0:y1, x1:w, :] = 0.75 * fill + 0.25 * bleed
        elif side == "top" and y0 > 0:
            bh = y0
            bleed = arr[:bh, x0:x1, :]
            inner = arr[y0:min(h, y0 + band), x0:x1, :]
            if inner.size == 0:
                return
            if energy(bleed) > max(abs_thr, energy(inner) * ratio_thr):
                anchor = arr[y0:y0 + 1, x0:x1, :]
                fill = np.repeat(anchor, bh, axis=0)
                arr[:bh, x0:x1, :] = 0.75 * fill + 0.25 * bleed
        elif side == "bottom" and y1 < h:
            bh = h - y1
            bleed = arr[y1:h, x0:x1, :]
            inner = arr[max(0, y1 - band):y1, x0:x1, :]
            if inner.size == 0:
                return
            if energy(bleed) > max(abs_thr, energy(inner) * ratio_thr):
                anchor = arr[y1 - 1:y1, x0:x1, :]
                fill = np.repeat(anchor, bh, axis=0)
                arr[y1:h, x0:x1, :] = 0.75 * fill + 0.25 * bleed

    for side in ("left", "right", "top", "bottom"):
        side_detect_and_fix(side)

    clean = Image.fromarray(np.clip(arr, 0, 255).astype(np.uint8), mode="RGB")
    # Light blur only in bleed area to hide residual line artifacts.
    mask = Image.new("L", (w, h), 255)
    interior = Image.new("L", (max(1, x1 - x0), max(1, y1 - y0)), 0)
    mask.paste(interior, (x0, y0))
    softened = clean.filter(ImageFilter.GaussianBlur(radius=0.9))
    return Image.composite(softened, clean, mask)


def _reinforce_boundary_bleed(img, x0, y0, x1, y1):
    """Force immediate bleed-side boundary strips to match interior edge colors."""
    arr = np.array(img).astype(np.float32)
    h, w = arr.shape[:2]
    bw = max(1, int(os.getenv("BLEED_BOUNDARY_BAND_PX", "3")))
    blend = float(os.getenv("BLEED_BOUNDARY_BLEND", "0.9"))

    # left bleed strip
    if x0 > 0:
        xs = max(0, x0 - bw)
        edge_col = arr[y0:y1, x0:x0 + 1, :]
        fill = np.repeat(edge_col, x0 - xs, axis=1)
        arr[y0:y1, xs:x0, :] = blend * fill + (1.0 - blend) * arr[y0:y1, xs:x0, :]

    # right bleed strip
    if x1 < w:
        xe = min(w, x1 + bw)
        edge_col = arr[y0:y1, x1 - 1:x1, :]
        fill = np.repeat(edge_col, xe - x1, axis=1)
        arr[y0:y1, x1:xe, :] = blend * fill + (1.0 - blend) * arr[y0:y1, x1:xe, :]

    # top bleed strip
    if y0 > 0:
        ys = max(0, y0 - bw)
        edge_row = arr[y0:y0 + 1, x0:x1, :]
        fill = np.repeat(edge_row, y0 - ys, axis=0)
        arr[ys:y0, x0:x1, :] = blend * fill + (1.0 - blend) * arr[ys:y0, x0:x1, :]

    # bottom bleed strip
    if y1 < h:
        ye = min(h, y1 + bw)
        edge_row = arr[y1 - 1:y1, x0:x1, :]
        fill = np.repeat(edge_row, ye - y1, axis=0)
        arr[y1:ye, x0:x1, :] = blend * fill + (1.0 - blend) * arr[y1:ye, x0:x1, :]

    return Image.fromarray(np.clip(arr, 0, 255).astype(np.uint8), mode="RGB")


def _fix_dark_boundary_lines(img, x0, y0, x1, y1):
    """Lift dark seam pixels right outside artwork bounds to match interior luminance."""
    arr = np.array(img).astype(np.float32)
    h, w = arr.shape[:2]
    bw = max(1, int(os.getenv("BLEED_DARKLINE_BAND_PX", "4")))
    dark_thr = float(os.getenv("BLEED_DARKLINE_DELTA", "8.0"))
    blend = float(os.getenv("BLEED_DARKLINE_BLEND", "0.82"))

    def luma(a):
        return 0.2126 * a[..., 0] + 0.7152 * a[..., 1] + 0.0722 * a[..., 2]

    # left
    if x0 > 0:
        xs = max(0, x0 - bw)
        seam = arr[y0:y1, xs:x0, :]
        inner = arr[y0:y1, x0:min(w, x0 + bw), :]
        if seam.size and inner.size:
            d = float(np.mean(luma(inner)) - np.mean(luma(seam)))
            if d > dark_thr:
                anchor = arr[y0:y1, x0:x0 + 1, :]
                fill = np.repeat(anchor, x0 - xs, axis=1)
                arr[y0:y1, xs:x0, :] = blend * fill + (1 - blend) * seam

    # right
    if x1 < w:
        xe = min(w, x1 + bw)
        seam = arr[y0:y1, x1:xe, :]
        inner = arr[y0:y1, max(0, x1 - bw):x1, :]
        if seam.size and inner.size:
            d = float(np.mean(luma(inner)) - np.mean(luma(seam)))
            if d > dark_thr:
                anchor = arr[y0:y1, x1 - 1:x1, :]
                fill = np.repeat(anchor, xe - x1, axis=1)
                arr[y0:y1, x1:xe, :] = blend * fill + (1 - blend) * seam

    # top
    if y0 > 0:
        ys = max(0, y0 - bw)
        seam = arr[ys:y0, x0:x1, :]
        inner = arr[y0:min(h, y0 + bw), x0:x1, :]
        if seam.size and inner.size:
            d = float(np.mean(luma(inner)) - np.mean(luma(seam)))
            if d > dark_thr:
                anchor = arr[y0:y0 + 1, x0:x1, :]
                fill = np.repeat(anchor, y0 - ys, axis=0)
                arr[ys:y0, x0:x1, :] = blend * fill + (1 - blend) * seam

    # bottom
    if y1 < h:
        ye = min(h, y1 + bw)
        seam = arr[y1:ye, x0:x1, :]
        inner = arr[max(0, y1 - bw):y1, x0:x1, :]
        if seam.size and inner.size:
            d = float(np.mean(luma(inner)) - np.mean(luma(seam)))
            if d > dark_thr:
                anchor = arr[y1 - 1:y1, x0:x1, :]
                fill = np.repeat(anchor, ye - y1, axis=0)
                arr[y1:ye, x0:x1, :] = blend * fill + (1 - blend) * seam

    return Image.fromarray(np.clip(arr, 0, 255).astype(np.uint8), mode="RGB")


def _smooth_boundary_transition(img, x0, y0, x1, y1):
    """Blend a tiny seam strip around artwork boundary to hide hard transitions."""
    arr = np.array(img).astype(np.float32)
    h, w = arr.shape[:2]
    bw = max(1, int(os.getenv("BLEED_TRANSITION_BAND_PX", "3")))

    # Left seam (outside only)
    if x0 > 0:
        xs = max(0, x0 - bw)
        for i, x in enumerate(range(xs, x0)):
            t = (i + 1) / max(1.0, (x0 - xs + 1))
            edge = arr[y0:y1, x0:x0 + 1, :]
            arr[y0:y1, x:x + 1, :] = (1 - t) * arr[y0:y1, x:x + 1, :] + t * edge

    # Right seam
    if x1 < w:
        xe = min(w, x1 + bw)
        for i, x in enumerate(range(x1, xe)):
            t = (i + 1) / max(1.0, (xe - x1 + 1))
            edge = arr[y0:y1, x1 - 1:x1, :]
            arr[y0:y1, x:x + 1, :] = (1 - t) * arr[y0:y1, x:x + 1, :] + t * edge

    # Top seam
    if y0 > 0:
        ys = max(0, y0 - bw)
        for i, y in enumerate(range(ys, y0)):
            t = (i + 1) / max(1.0, (y0 - ys + 1))
            edge = arr[y0:y0 + 1, x0:x1, :]
            arr[y:y + 1, x0:x1, :] = (1 - t) * arr[y:y + 1, x0:x1, :] + t * edge

    # Bottom seam
    if y1 < h:
        ye = min(h, y1 + bw)
        for i, y in enumerate(range(y1, ye)):
            t = (i + 1) / max(1.0, (ye - y1 + 1))
            edge = arr[y1 - 1:y1, x0:x1, :]
            arr[y:y + 1, x0:x1, :] = (1 - t) * arr[y:y + 1, x0:x1, :] + t * edge

    return Image.fromarray(np.clip(arr, 0, 255).astype(np.uint8), mode="RGB")


def _denoise_bleed_lines(img, x0, y0, x1, y1):
    """Target thin line artifacts in bleed-only regions while preserving center artwork."""
    w, h = img.size
    # Bleed mask: white outside artwork, black on artwork.
    mask = Image.new("L", (w, h), 255)
    inner = Image.new("L", (max(1, x1 - x0), max(1, y1 - y0)), 0)
    mask.paste(inner, (x0, y0))
    med = img.filter(ImageFilter.MedianFilter(size=3))
    soft = med.filter(ImageFilter.GaussianBlur(radius=0.6))
    den = Image.composite(soft, img, mask)
    return den


def _build_extrapolated_canvas(rgb, pad_left, pad_right, pad_top, pad_bottom):
    """Extend image using local gradient extrapolation (curve-friendly, non-mirrored)."""
    h, w = rgb.shape[:2]
    hh = h + pad_top + pad_bottom
    ww = w + pad_left + pad_right
    out = np.zeros((hh, ww, 3), dtype=np.float32)
    x0, y0 = pad_left, pad_top
    x1, y1 = x0 + w, y0 + h
    core = rgb.astype(np.float32)
    out[y0:y1, x0:x1, :] = core

    # Use averaged boundary bands (not single px) to reduce line artifacts.
    k = max(2, min(10, int(os.getenv("EXTRAP_BAND_PX", "6"))))
    l1 = np.mean(core[:, :k, :], axis=1)
    l2 = np.mean(core[:, k:min(w, 2 * k), :], axis=1) if w > k else l1
    r1 = np.mean(core[:, max(0, w - k):w, :], axis=1)
    r2 = np.mean(core[:, max(0, w - 2 * k):max(1, w - k), :], axis=1) if w > k else r1
    t1 = np.mean(core[:k, :, :], axis=0)
    t2 = np.mean(core[k:min(h, 2 * k), :, :], axis=0) if h > k else t1
    b1 = np.mean(core[max(0, h - k):h, :, :], axis=0)
    b2 = np.mean(core[max(0, h - 2 * k):max(1, h - k), :, :], axis=0) if h > k else b1

    if pad_left > 0:
        for i in range(pad_left):
            t = (pad_left - i) / max(1.0, float(pad_left))
            col = l1 + (l1 - l2) * (0.35 * t)
            out[y0:y1, i:i + 1, :] = col[:, None, :]
    if pad_right > 0:
        for i in range(pad_right):
            t = (pad_right - i) / max(1.0, float(pad_right))
            col = r1 + (r1 - r2) * (0.35 * t)
            out[y0:y1, x1 + i:x1 + i + 1, :] = col[:, None, :]
    if pad_top > 0:
        for i in range(pad_top):
            t = (pad_top - i) / max(1.0, float(pad_top))
            row = t1 + (t1 - t2) * (0.35 * t)
            out[i:i + 1, x0:x1, :] = row[None, :, :]
    if pad_bottom > 0:
        for i in range(pad_bottom):
            t = (pad_bottom - i) / max(1.0, float(pad_bottom))
            row = b1 + (b1 - b2) * (0.35 * t)
            out[y1 + i:y1 + i + 1, x0:x1, :] = row[None, :, :]

    # Corners: blend adjacent extrapolated sides.
    if pad_left > 0 and pad_top > 0:
        tl_row = out[y0:y0 + 1, :x0, :]
        tl_col = out[:y0, x0:x0 + 1, :]
        out[:y0, :x0, :] = 0.5 * (tl_row.repeat(y0, axis=0) + tl_col.repeat(x0, axis=1))
    if pad_right > 0 and pad_top > 0:
        tr_row = out[y0:y0 + 1, x1:, :]
        tr_col = out[:y0, x1 - 1:x1, :]
        out[:y0, x1:, :] = 0.5 * (tr_row.repeat(y0, axis=0) + tr_col.repeat(pad_right, axis=1))
    if pad_left > 0 and pad_bottom > 0:
        bl_row = out[y1 - 1:y1, :x0, :]
        bl_col = out[y1:, x0:x0 + 1, :]
        out[y1:, :x0, :] = 0.5 * (bl_row.repeat(pad_bottom, axis=0) + bl_col.repeat(x0, axis=1))
    if pad_right > 0 and pad_bottom > 0:
        br_row = out[y1 - 1:y1, x1:, :]
        br_col = out[y1:, x1 - 1:x1, :]
        out[y1:, x1:, :] = 0.5 * (br_row.repeat(pad_bottom, axis=0) + br_col.repeat(pad_right, axis=1))

    return np.clip(out, 0, 255).astype(np.uint8)


def _sidewise_ai_guard(ai_img, seed_img, x0, y0, x1, y1):
    """Replace AI bleed sides that diverge too much from deterministic seed."""
    ai = np.array(ai_img).astype(np.float32)
    seed = np.array(seed_img).astype(np.float32)
    h, w = ai.shape[:2]

    def side_ranges():
        if x0 > 0:
            yield ("left", (slice(y0, y1), slice(0, x0)))
        if x1 < w:
            yield ("right", (slice(y0, y1), slice(x1, w)))
        if y0 > 0:
            yield ("top", (slice(0, y0), slice(x0, x1)))
        if y1 < h:
            yield ("bottom", (slice(y1, h), slice(x0, x1)))

    color_thr = float(os.getenv("AI_SIDE_GUARD_COLOR_DELTA", "28"))
    edge_ratio_thr = float(os.getenv("AI_SIDE_GUARD_EDGE_RATIO", "1.5"))

    def edge_energy(region):
        if region.size == 0:
            return 0.0
        gx = np.abs(np.diff(region, axis=1)).mean() if region.shape[1] > 1 else 0.0
        gy = np.abs(np.diff(region, axis=0)).mean() if region.shape[0] > 1 else 0.0
        return float(gx + gy)

    for _, (ys, xs) in side_ranges():
        a = ai[ys, xs, :]
        s = seed[ys, xs, :]
        if a.size == 0 or s.size == 0:
            continue
        c_delta = float(np.linalg.norm(np.mean(a, axis=(0, 1)) - np.mean(s, axis=(0, 1))))
        e_ratio = edge_energy(a) / max(1e-3, edge_energy(s))
        if c_delta > color_thr or e_ratio > edge_ratio_thr:
            ai[ys, xs, :] = s

    return Image.fromarray(np.clip(ai, 0, 255).astype(np.uint8), mode="RGB")


def _blend_ai_near_trim(ai_img, seed_img, x0, y0, x1, y1):
    """Use seed close to trim edge; keep AI mostly in outer bleed to prevent seam/glow artifacts."""
    ai = np.array(ai_img).astype(np.float32)
    seed = np.array(seed_img).astype(np.float32)
    h, w = ai.shape[:2]
    yy, xx = np.mgrid[0:h, 0:w]
    alpha = np.zeros((h, w), dtype=np.float32)  # AI weight
    fade_px = max(6, int(os.getenv("AI_TRIM_FADE_PX", "22")))
    keep_seed_px = max(2, int(os.getenv("AI_TRIM_KEEP_SEED_PX", "5")))

    if x0 > 0:
        d = (x0 - xx).astype(np.float32)
        left = xx < x0
        a = np.clip((d - keep_seed_px) / max(1.0, float(fade_px)), 0.0, 1.0)
        alpha[left] = np.maximum(alpha[left], a[left])
    if x1 < w:
        d = (xx - x1 + 1).astype(np.float32)
        right = xx >= x1
        a = np.clip((d - keep_seed_px) / max(1.0, float(fade_px)), 0.0, 1.0)
        alpha[right] = np.maximum(alpha[right], a[right])
    if y0 > 0:
        d = (y0 - yy).astype(np.float32)
        top = yy < y0
        a = np.clip((d - keep_seed_px) / max(1.0, float(fade_px)), 0.0, 1.0)
        alpha[top] = np.maximum(alpha[top], a[top])
    if y1 < h:
        d = (yy - y1 + 1).astype(np.float32)
        bottom = yy >= y1
        a = np.clip((d - keep_seed_px) / max(1.0, float(fade_px)), 0.0, 1.0)
        alpha[bottom] = np.maximum(alpha[bottom], a[bottom])

    alpha = alpha[..., None]
    out = seed * (1.0 - alpha) + ai * alpha
    # Keep interior untouched; original slide is pasted later regardless.
    return Image.fromarray(np.clip(out, 0, 255).astype(np.uint8), mode="RGB")


def _suppress_structural_hallucinations(ai_img, seed_img, x0, y0, x1, y1):
    """Blend back to seed where AI introduces high-contrast structures in bleed regions."""
    ai = np.array(ai_img).astype(np.float32)
    seed = np.array(seed_img).astype(np.float32)
    h, w = ai.shape[:2]
    bleed_mask = np.ones((h, w), dtype=np.float32)
    bleed_mask[y0:y1, x0:x1] = 0.0

    delta = np.linalg.norm(ai - seed, axis=2)
    gray = 0.2126 * ai[:, :, 0] + 0.7152 * ai[:, :, 1] + 0.0722 * ai[:, :, 2]
    gx = np.zeros_like(gray)
    gy = np.zeros_like(gray)
    gx[:, 1:] = np.abs(gray[:, 1:] - gray[:, :-1])
    gy[1:, :] = np.abs(gray[1:, :] - gray[:-1, :])
    grad = gx + gy

    delta_thr = float(os.getenv("AI_HALLUC_DELTA_THR", "16.0"))
    grad_thr = float(os.getenv("AI_HALLUC_GRAD_THR", "9.0"))
    bad = ((delta > delta_thr) & (grad > grad_thr) & (bleed_mask > 0.5)).astype(np.uint8) * 255
    if np.count_nonzero(bad) == 0:
        return ai_img

    mask = Image.fromarray(bad, mode="L").filter(ImageFilter.GaussianBlur(radius=1.2))
    m = (np.array(mask).astype(np.float32) / 255.0) * float(os.getenv("AI_HALLUC_SEED_BLEND", "0.8"))
    m = np.clip(m, 0.0, 1.0)[..., None]
    out = ai * (1.0 - m) + seed * m
    return Image.fromarray(np.clip(out, 0, 255).astype(np.uint8), mode="RGB")


def _run_realesrgan_ncnn(image, scale):
    executable = shutil.which("realesrgan-ncnn-vulkan")
    if not executable:
        direct_path = Path("/opt/realesrgan/realesrgan-ncnn-vulkan")
        if direct_path.exists():
            executable = str(direct_path)
    if not executable:
        return None

    input_path = Path("/tmp") / f"pp_in_{os.getpid()}_{id(image)}.png"
    output_path = Path("/tmp") / f"pp_out_{os.getpid()}_{id(image)}.png"
    model_name = "realesrgan-x4plus"
    model_dir = os.getenv("REALESRGAN_MODEL_DIR", "/opt/realesrgan/models")
    scale_arg = "4"
    try:
        image.convert("RGB").save(input_path)
        cmd = [
            executable, "-i", str(input_path), "-o", str(output_path),
            "-n", model_name, "-m", model_dir, "-s", scale_arg,
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        if result.returncode != 0 or not output_path.exists():
            return None
        out = Image.open(output_path).convert("RGB")
        return out.resize((image.width * scale, image.height * scale), Image.Resampling.LANCZOS)
    except Exception:
        return None
    finally:
        input_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def _is_low_detail_image(image):
    """Detect flat/gradient-heavy slides where ESRGAN may introduce banding/lines."""
    arr = np.array(image.convert("RGB")).astype(np.float32)
    if arr.size == 0:
        return True
    # Luma variation and edge energy.
    y = 0.2126 * arr[..., 0] + 0.7152 * arr[..., 1] + 0.0722 * arr[..., 2]
    std = float(np.std(y))
    gx = np.abs(np.diff(y, axis=1)).mean() if y.shape[1] > 1 else 0.0
    gy = np.abs(np.diff(y, axis=0)).mean() if y.shape[0] > 1 else 0.0
    edge = float(gx + gy)
    # Low-detail gradients typically have low local edge energy.
    return std < float(os.getenv("LOW_DETAIL_STD_MAX", "28")) and edge < float(os.getenv("LOW_DETAIL_EDGE_MAX", "7.5"))


def _enhance_with_print_rescue_upscaler(image):
    if os.getenv("PRINT_RESCUE_FORCE_NO_UPSCALE", "false").strip().lower() == "true":
        return image
    if _is_low_detail_image(image):
        # Avoid hallucinated lines on smooth backgrounds.
        return image
    scale = int(os.getenv("PRINT_RESCUE_UPSCALE_FACTOR", "2"))
    if scale <= 1:
        return image
    upscaled = _run_realesrgan_ncnn(image, scale)
    if upscaled is None:
        upscaled = image.resize((image.width * scale, image.height * scale), Image.Resampling.LANCZOS)
    # Return to original geometry while preserving center replacement later.
    return upscaled.resize((image.width, image.height), Image.Resampling.LANCZOS)


def _sanitize_edge_halo(rgb):
    if rgb.ndim != 3 or rgb.shape[0] < 4 or rgb.shape[1] < 4:
        return rgb
    fixed = rgb.copy()
    band = 2
    fixed[:band, :, :] = fixed[band:band + 1, :, :]
    fixed[-band:, :, :] = fixed[-band - 1:-band, :, :]
    fixed[:, :band, :] = fixed[:, band:band + 1, :]
    fixed[:, -band:, :] = fixed[:, -band - 1:-band, :]
    return fixed


def _remove_white_frame(rgb):
    if rgb.ndim != 3 or rgb.shape[0] < 32 or rgb.shape[1] < 32:
        return rgb

    h, w = rgb.shape[:2]
    max_scan = max(2, min(36, int(min(h, w) * 0.08)))

    def row_is_white(i):
        row = rgb[i, :, :]
        bright = np.mean(row, axis=1) > 244
        return float(np.mean(bright)) > 0.985 and float(np.std(row)) < 10.0

    def col_is_white(i):
        col = rgb[:, i, :]
        bright = np.mean(col, axis=1) > 244
        return float(np.mean(bright)) > 0.985 and float(np.std(col)) < 10.0

    top = 0
    while top < max_scan and row_is_white(top):
        top += 1
    bottom = 0
    while bottom < max_scan and row_is_white(h - 1 - bottom):
        bottom += 1
    left = 0
    while left < max_scan and col_is_white(left):
        left += 1
    right = 0
    while right < max_scan and col_is_white(w - 1 - right):
        right += 1

    if max(top, bottom, left, right) == 0:
        return rgb

    y0 = min(top, h - 4)
    y1 = max(y0 + 4, h - bottom)
    x0 = min(left, w - 4)
    x1 = max(x0 + 4, w - right)
    if y1 - y0 < int(h * 0.85) or x1 - x0 < int(w * 0.85):
        return rgb

    cropped = rgb[y0:y1, x0:x1, :]
    if cropped.shape[0] == h and cropped.shape[1] == w:
        return rgb
    resized = Image.fromarray(cropped, mode="RGB").resize((w, h), Image.Resampling.BILINEAR)
    return np.array(resized)


def _neutralize_edge_lines(rgb):
    if rgb.ndim != 3 or rgb.shape[0] < 12 or rgb.shape[1] < 12:
        return rgb

    out = rgb.copy()
    h, w = out.shape[:2]

    def row_luma(idx):
        return float(np.mean(out[idx, :, :]))

    def col_luma(idx):
        return float(np.mean(out[:, idx, :]))

    for _ in range(2):
        for i in range(0, min(4, h - 3)):
            ref = min(i + 2, h - 1)
            if row_luma(i) > 220 and (row_luma(i) - row_luma(ref)) > 8:
                out[i, :, :] = out[ref, :, :]
        for i in range(0, min(4, w - 3)):
            ref = min(i + 2, w - 1)
            if col_luma(i) > 220 and (col_luma(i) - col_luma(ref)) > 8:
                out[:, i, :] = out[:, ref, :]
        for i in range(0, min(4, h - 3)):
            idx = h - 1 - i
            ref = max(idx - 2, 0)
            if row_luma(idx) > 220 and (row_luma(idx) - row_luma(ref)) > 8:
                out[idx, :, :] = out[ref, :, :]
        for i in range(0, min(4, w - 3)):
            idx = w - 1 - i
            ref = max(idx - 2, 0)
            if col_luma(idx) > 220 and (col_luma(idx) - col_luma(ref)) > 8:
                out[:, idx, :] = out[:, ref, :]
    return out


def _print_rescue_render_dpi():
    """Configurable render DPI for Print Rescue background generation."""
    raw = os.getenv("PRINT_RESCUE_RENDER_DPI", "240")
    try:
        dpi = int(float(raw))
    except Exception:
        dpi = 240
    return max(96, min(600, dpi))


def _print_rescue_quality_to_dpi(quality):
    q = (quality or "balanced").strip().lower()
    mapping = {"balanced": 240, "high": 300, "ultra": 360}
    return mapping.get(q, mapping["balanced"])


def _build_reflect_seed(source_page, media_w_pt, media_h_pt, x_off_pt, y_off_pt, scaled_w_pt, scaled_h_pt,
                        max_side_px=None, mult=1, render_dpi=None):
    render_dpi = int(render_dpi or _print_rescue_render_dpi())
    px_per_pt = render_dpi / 72.0
    media_w_px = max(8, int(round(media_w_pt * px_per_pt)))
    media_h_px = max(8, int(round(media_h_pt * px_per_pt)))
    placed_w_px = max(8, int(round(scaled_w_pt * px_per_pt)))
    placed_h_px = max(8, int(round(scaled_h_pt * px_per_pt)))
    x0 = int(round(x_off_pt * px_per_pt))
    y0 = int(round(y_off_pt * px_per_pt))
    x0 = max(0, min(media_w_px - 1, x0))
    y0 = max(0, min(media_h_px - 1, y0))
    x1 = max(x0 + 1, min(media_w_px, x0 + placed_w_px))
    y1 = max(y0 + 1, min(media_h_px, y0 + placed_h_px))

    if max_side_px is not None:
        scale_down = min(1.0, max_side_px / max(media_w_px, media_h_px))
        if scale_down < 1.0:
            media_w_px = _round_to_mult(media_w_px * scale_down, max(1, mult))
            media_h_px = _round_to_mult(media_h_px * scale_down, max(1, mult))
            x0 = int(round(x0 * scale_down))
            y0 = int(round(y0 * scale_down))
            x1 = max(x0 + 1, int(round(x1 * scale_down)))
            y1 = max(y0 + 1, int(round(y1 * scale_down)))
        if mult > 1:
            media_w_px = max(mult, min(max_side_px, _round_to_mult(media_w_px, mult)))
            media_h_px = max(mult, min(max_side_px, _round_to_mult(media_h_px, mult)))
        x0 = max(0, min(media_w_px - 2, x0))
        y0 = max(0, min(media_h_px - 2, y0))
        x1 = max(x0 + 1, min(media_w_px, x1))
        y1 = max(y0 + 1, min(media_h_px, y1))

    placed_raw = _fitz_page_to_pil(source_page, max(8, x1 - x0), max(8, y1 - y0)).convert("RGB")
    rgb = np.array(placed_raw)
    trim_white = (os.getenv("OUTPAINT_TRIM_WHITE_FRAME", "false").strip().lower() == "true")
    if trim_white:
        rgb = _remove_white_frame(rgb)
    rgb = _neutralize_edge_lines(rgb)
    rgb = _sanitize_edge_halo(rgb)
    placed = Image.fromarray(np.clip(rgb, 0, 255).astype(np.uint8), mode="RGB")

    pad_top = y0
    pad_bottom = max(0, media_h_px - y1)
    pad_left = x0
    pad_right = max(0, media_w_px - x1)
    extended = _build_extrapolated_canvas(rgb, pad_left, pad_right, pad_top, pad_bottom)
    canvas = Image.fromarray(extended.astype(np.uint8), mode="RGB")
    canvas.paste(placed, (x0, y0))

    mask = Image.new("L", (media_w_px, media_h_px), 255)
    interior = Image.new("L", (max(1, x1 - x0), max(1, y1 - y0)), 0)
    mask.paste(interior, (x0, y0))
    return canvas, mask, placed, x0, y0, media_w_px, media_h_px


def _build_fallback_bleed_background_jpeg(source_page, media_w_pt, media_h_pt,
                                          x_off_pt, y_off_pt, scaled_w_pt, scaled_h_pt,
                                          provider="fallback", render_dpi=None):
    """Deterministic reflective extension (Print Rescue style)."""
    canvas, _, placed, x0, y0, media_w_px, media_h_px = _build_reflect_seed(
        source_page, media_w_pt, media_h_pt, x_off_pt, y_off_pt, scaled_w_pt, scaled_h_pt,
        render_dpi=render_dpi,
    )
    use_upscaler = (
        provider == "print_rescue_local"
        and os.getenv("PRINT_RESCUE_USE_UPSCALER", "true").strip().lower() == "true"
    )
    if use_upscaler:
        canvas = _enhance_with_print_rescue_upscaler(canvas)
    canvas = _suppress_bleed_artifacts(canvas, x0, y0, x0 + placed.width, y0 + placed.height)
    canvas = _reinforce_boundary_bleed(canvas, x0, y0, x0 + placed.width, y0 + placed.height)
    canvas = _fix_dark_boundary_lines(canvas, x0, y0, x0 + placed.width, y0 + placed.height)
    canvas = _smooth_boundary_transition(canvas, x0, y0, x0 + placed.width, y0 + placed.height)
    canvas = _denoise_bleed_lines(canvas, x0, y0, x0 + placed.width, y0 + placed.height)
    canvas.paste(placed, (x0, y0))
    out = io.BytesIO()
    canvas.save(out, format="JPEG", quality=90, optimize=True, progressive=True)
    return out.getvalue(), media_w_px, media_h_px, provider
    render_dpi = 144
    px_per_pt = render_dpi / 72.0
    media_w_px = max(8, int(round(media_w_pt * px_per_pt)))
    media_h_px = max(8, int(round(media_h_pt * px_per_pt)))
    placed_w_px = max(8, int(round(scaled_w_pt * px_per_pt)))
    placed_h_px = max(8, int(round(scaled_h_pt * px_per_pt)))
    x0 = int(round(x_off_pt * px_per_pt))
    y0 = int(round(y_off_pt * px_per_pt))
    x0 = max(0, min(media_w_px - 1, x0))
    y0 = max(0, min(media_h_px - 1, y0))
    x1 = max(x0 + 1, min(media_w_px, x0 + placed_w_px))
    y1 = max(y0 + 1, min(media_h_px, y0 + placed_h_px))

    placed = _fitz_page_to_pil(source_page, x1 - x0, y1 - y0).convert("RGB")
    canvas = Image.new("RGB", (media_w_px, media_h_px), (255, 255, 255))
    canvas.paste(placed, (x0, y0))

    # Use a full-canvas extension derived from the source artwork.
    # This avoids white halos and soft glow around the placed art.
    canvas = ImageOps.fit(
        placed,
        (media_w_px, media_h_px),
        method=Image.Resampling.LANCZOS,
        centering=(0.5, 0.5),
    )
    canvas.paste(placed, (x0, y0))

    out = io.BytesIO()
    canvas.save(out, format="JPEG", quality=88, optimize=True, progressive=True)
    return out.getvalue(), media_w_px, media_h_px, "fallback"


def _run_replicate_outpaint(image_data_url, mask_data_url, prompt, width_px, height_px,
                            pad_left=0, pad_right=0, pad_top=0, pad_bottom=0):
    token = (os.getenv("REPLICATE_API_TOKEN") or "").strip()
    if not token:
        raise RuntimeError("REPLICATE_API_TOKEN is not set")

    model_ref = (os.getenv("REPLICATE_MODEL") or OUTPAINT_DEFAULT_MODEL).strip()
    headers = {"Authorization": f"Token {token}"}
    neg = (os.getenv("OUTPAINT_NEGATIVE_PROMPT") or OUTPAINT_DEFAULT_NEGATIVE_PROMPT)
    if model_ref.startswith("emaph/outpaint-controlnet-union"):
        input_payload = {
            "image": image_data_url,
            "left": int(max(0, pad_left)),
            "right": int(max(0, pad_right)),
            "top": int(max(0, pad_top)),
            "bottom": int(max(0, pad_bottom)),
            "prompt": os.getenv("OUTPAINT_PROMPT_EMAPH", prompt),
            "negative_prompt": neg,
            "steps": int(os.getenv("OUTPAINT_STEPS", "20")),
            "cfg": float(os.getenv("OUTPAINT_CFG", "2.0")),
            "output_format": "png",
            "output_quality": int(os.getenv("OUTPAINT_OUTPUT_QUALITY", "95")),
        }
    elif model_ref.startswith("fermatresearch/sdxl-outpainting-lora"):
        input_payload = {
            "image": image_data_url,
            "outpaint_left": int(max(0, min(512, pad_left))),
            "outpaint_right": int(max(0, min(512, pad_right))),
            "outpaint_up": int(max(0, min(512, pad_top))),
            "outpaint_down": int(max(0, min(512, pad_bottom))),
            "prompt": (os.getenv("OUTPAINT_PROMPT_FERMAT") or OUTPAINT_FERMAT_PROMPT),
            "negative_prompt": neg,
            "num_outputs": 1,
            "guidance_scale": float(os.getenv("OUTPAINT_GUIDANCE_SCALE", "2.2")),
            "condition_scale": float(os.getenv("OUTPAINT_CONDITION_SCALE", "0.35")),
            "lora_scale": float(os.getenv("OUTPAINT_LORA_SCALE", "0.45")),
            "scheduler": os.getenv("OUTPAINT_SCHEDULER", "K_EULER"),
            "apply_watermark": False,
        }
    else:
        # Generic inpainting path.
        input_payload = {
            "image": image_data_url,
            "mask": mask_data_url,
            "prompt": prompt,
            "negative_prompt": neg,
            "width": int(width_px),
            "height": int(height_px),
            "num_outputs": 1,
            "num_inference_steps": int(os.getenv("OUTPAINT_STEPS", "28")),
            "guidance_scale": float(os.getenv("OUTPAINT_GUIDANCE_SCALE", "3.5")),
            "scheduler": os.getenv("OUTPAINT_SCHEDULER", "DPMSolverMultistep"),
            "strength": float(os.getenv("OUTPAINT_STRENGTH", "0.2")),
        }

    def create_prediction_with_version(version_id):
        global REPLICATE_LAST_CREATE_TS
        min_gap = float(os.getenv("REPLICATE_CREATE_MIN_GAP_SECONDS", "10"))
        max_attempts = int(os.getenv("REPLICATE_CREATE_MAX_ATTEMPTS", "8"))
        create_url = f"{OUTPAINT_API_BASE}/predictions"
        payload = {"version": version_id, "input": input_payload}
        for attempt in range(max_attempts):
            with REPLICATE_CREATE_LOCK:
                now = time.time()
                wait = max(0.0, min_gap - (now - REPLICATE_LAST_CREATE_TS))
                if wait > 0:
                    time.sleep(wait)
                try:
                    pred = _http_json(
                        "POST",
                        create_url,
                        headers=headers,
                        payload=payload,
                        timeout=90,
                    )
                    REPLICATE_LAST_CREATE_TS = time.time()
                    return pred
                except HttpApiError as e:
                    if e.status != 429:
                        raise
                    retry_after = 10
                    try:
                        parsed = json.loads(e.detail)
                        retry_after = int(parsed.get("retry_after") or retry_after)
                    except Exception:
                        pass
                    REPLICATE_LAST_CREATE_TS = time.time() + retry_after
            # Sleep outside lock; add slight jitter so concurrent jobs de-sync.
            time.sleep(retry_after + 0.4 + min(1.5, attempt * 0.2))
        raise RuntimeError("Replicate rate-limited too many times; please retry shortly.")

    def latest_model_version(model_slug):
        meta = _http_json(
            "GET",
            f"{OUTPAINT_API_BASE}/models/{model_slug}",
            headers=headers,
            timeout=90,
        )
        latest = (meta.get("latest_version") or {}).get("id")
        if not latest:
            raise RuntimeError(f"Replicate model has no latest_version: {model_slug}")
        return latest

    if ":" in model_ref:
        version = model_ref.split(":", 1)[1].strip()
    else:
        version = latest_model_version(model_ref)
    pred = create_prediction_with_version(version)
    poll_url = pred.get("urls", {}).get("get")
    if not poll_url:
        raise RuntimeError("Replicate did not return a prediction polling URL")

    start = time.time()
    timeout_s = int(os.getenv("OUTPAINT_TIMEOUT_SECONDS", "360"))
    while True:
        status = (pred.get("status") or "").lower()
        if status == "succeeded":
            break
        if status in {"failed", "canceled"}:
            err = pred.get("error") or "prediction failed"
            raise RuntimeError(f"Replicate outpaint failed: {err}")
        if time.time() - start > timeout_s:
            raise RuntimeError("Replicate outpaint timed out")
        time.sleep(1.1)
        pred = _http_json("GET", poll_url, headers=headers, timeout=90)

    output = pred.get("output")
    out_url = None
    if isinstance(output, str):
        out_url = output
    elif isinstance(output, list) and output:
        out_url = output[0]
    elif isinstance(output, dict):
        out_url = output.get("image") or output.get("url")
    if not out_url:
        raise RuntimeError("Replicate returned no output image URL")
    return _http_bytes(out_url, timeout=180)


def _run_adobe_firefly_expand(image_png_bytes, prompt, pad_left=0, pad_right=0, pad_top=0, pad_bottom=0):
    api_key = (os.getenv("ADOBE_FIREFLY_API_KEY") or "").strip()
    access_token = (os.getenv("ADOBE_FIREFLY_ACCESS_TOKEN") or "").strip()
    if not api_key or not access_token:
        raise RuntimeError("ADOBE_FIREFLY_API_KEY / ADOBE_FIREFLY_ACCESS_TOKEN not configured")

    headers = {
        "x-api-key": api_key,
        "Authorization": f"Bearer {access_token}",
    }

    # 1) Upload source image
    upload_payload = {
        "image": {
            "contentType": "image/png",
            "data": base64.b64encode(image_png_bytes).decode("ascii"),
        }
    }
    up = _http_json_raw(
        "POST",
        f"{ADOBE_FIREFLY_BASE}/v2/storage/image",
        headers=headers,
        payload=upload_payload,
        timeout=90,
    )
    source_id = up.get("storageId") or up.get("id") or up.get("imageId")
    if not source_id:
        raise RuntimeError("Adobe Firefly upload did not return storage id")

    # 2) Start async expand job
    expand_payload = {
        "input": {
            "image": {"storageId": source_id},
            "prompt": prompt,
            "negativePrompt": (os.getenv("OUTPAINT_NEGATIVE_PROMPT") or OUTPAINT_DEFAULT_NEGATIVE_PROMPT),
            "padding": {
                "left": int(max(0, pad_left)),
                "right": int(max(0, pad_right)),
                "top": int(max(0, pad_top)),
                "bottom": int(max(0, pad_bottom)),
            },
        }
    }
    job = _http_json_raw(
        "POST",
        f"{ADOBE_FIREFLY_BASE}/v3/images/expand-async",
        headers=headers,
        payload=expand_payload,
        timeout=90,
    )
    status_url = (
        job.get("statusUrl")
        or ((job.get("_links") or {}).get("self") or {}).get("href")
        or job.get("href")
    )
    if not status_url:
        raise RuntimeError("Adobe Firefly did not return status URL")

    timeout_s = int(os.getenv("OUTPAINT_TIMEOUT_SECONDS", "360"))
    start = time.time()
    while True:
        if time.time() - start > timeout_s:
            raise RuntimeError("Adobe Firefly expand timed out")
        state = _http_json_raw("GET", status_url, headers=headers, timeout=90)
        status = (state.get("status") or state.get("state") or "").lower()
        if status in {"succeeded", "success", "done", "completed"}:
            break
        if status in {"failed", "error", "canceled"}:
            raise RuntimeError(f"Adobe Firefly expand failed: {state}")
        time.sleep(1.2)

    out_url = None
    outputs = state.get("outputs") or state.get("result") or []
    if isinstance(outputs, list) and outputs:
        first = outputs[0]
        if isinstance(first, dict):
            out_url = first.get("url") or first.get("href") or ((first.get("image") or {}).get("url"))
    if not out_url:
        out_url = state.get("url") or ((state.get("image") or {}).get("url"))
    if not out_url:
        raise RuntimeError("Adobe Firefly expand returned no output URL")
    return _http_bytes(out_url, timeout=180)


def _build_ai_bleed_background_jpeg(source_page, media_w_pt, media_h_pt,
                                    x_off_pt, y_off_pt, scaled_w_pt, scaled_h_pt,
                                    ai_provider="replicate"):
    """Primary AI bleed path: model outpainting with local fallback."""
    # Seed with reflective extension so the model refines an already-correct bleed,
    # instead of inventing a whole new background from blank pixels.
    canvas, mask, placed, x0, y0, media_w_px, media_h_px = _build_reflect_seed(
        source_page,
        media_w_pt,
        media_h_pt,
        x_off_pt,
        y_off_pt,
        scaled_w_pt,
        scaled_h_pt,
        max_side_px=OUTPAINT_MAX_SIDE_PX,
        mult=64,
    )

    prompt = (os.getenv("OUTPAINT_PROMPT") or OUTPAINT_DEFAULT_PROMPT).strip()
    try:
        pad_left = x0
        pad_right = max(0, media_w_px - (x0 + placed.width))
        pad_top = y0
        pad_bottom = max(0, media_h_px - (y0 + placed.height))
        candidate_count = max(1, min(3, int(os.getenv("OUTPAINT_CANDIDATES", "2"))))
        best_img = None
        best_score = float("inf")
        provider_used = "replicate"
        for _ in range(candidate_count):
            out_bytes = _run_replicate_outpaint(
                _image_to_data_url(canvas, "PNG"),
                _image_to_data_url(mask, "PNG"),
                prompt,
                media_w_px,
                media_h_px,
                pad_left=pad_left,
                pad_right=pad_right,
                pad_top=pad_top,
                pad_bottom=pad_bottom,
            )
            cand = Image.open(io.BytesIO(out_bytes)).convert("RGB")
            if cand.size != (media_w_px, media_h_px):
                cand = cand.resize((media_w_px, media_h_px), Image.Resampling.LANCZOS)
            score = _score_bleed_candidate(cand, canvas, x0, y0, x0 + placed.width, y0 + placed.height)
            if score < best_score:
                best_score = score
                best_img = cand

        out_img = best_img if best_img is not None else canvas.copy()
        # Guardrail chain: keep AI only where it improves bleed and avoid seams.
        out_img = _blend_ai_near_trim(out_img, canvas, x0, y0, x0 + placed.width, y0 + placed.height)
        out_img = _sidewise_ai_guard(out_img, canvas, x0, y0, x0 + placed.width, y0 + placed.height)
        out_img = _suppress_structural_hallucinations(out_img, canvas, x0, y0, x0 + placed.width, y0 + placed.height)
        out_img = _suppress_bleed_artifacts(out_img, x0, y0, x0 + placed.width, y0 + placed.height)
        out_img = _reinforce_boundary_bleed(out_img, x0, y0, x0 + placed.width, y0 + placed.height)
        out_img = _fix_dark_boundary_lines(out_img, x0, y0, x0 + placed.width, y0 + placed.height)
        out_img = _smooth_boundary_transition(out_img, x0, y0, x0 + placed.width, y0 + placed.height)
        out_img = _denoise_bleed_lines(out_img, x0, y0, x0 + placed.width, y0 + placed.height)
        if not _bleed_quality_ok(out_img, x0, y0, x0 + placed.width, y0 + placed.height):
            return _build_fallback_bleed_background_jpeg(
                source_page, media_w_pt, media_h_pt, x_off_pt, y_off_pt, scaled_w_pt, scaled_h_pt,
                provider="quality_guard_fallback",
            )
        # Keep original slide pixels exact inside trim to prevent text drift.
        out_img.paste(placed, (x0, y0))
        out = io.BytesIO()
        out_img.save(out, format="JPEG", quality=90, optimize=True, progressive=True)
        return out.getvalue(), media_w_px, media_h_px, provider_used
    except Exception:
        allow_fallback = (os.getenv("OUTPAINT_ALLOW_FALLBACK", "false").strip().lower() == "true")
        if allow_fallback:
            return _build_fallback_bleed_background_jpeg(
                source_page, media_w_pt, media_h_pt, x_off_pt, y_off_pt, scaled_w_pt, scaled_h_pt, provider="fallback"
            )
        raise


def build_trim_marks_stream(trim_x, trim_y, trim_w, trim_h):
    lines = ["q"]
    corners = [
        (trim_x, trim_y, -1, -1),
        (trim_x + trim_w, trim_y, +1, -1),
        (trim_x + trim_w, trim_y + trim_h, +1, +1),
        (trim_x, trim_y + trim_h, -1, +1),
    ]
    # White halo
    lines.append(f"{WHITE_WEIGHT} w")
    lines.append("1 1 1 RG")
    lines.append("2 J")
    for cx, cy, hdir, vdir in corners:
        hx_s = cx + MARK_OFFSET * hdir
        hx_e = cx + (MARK_OFFSET + MARK_LENGTH) * hdir
        lines.append(f"{hx_s:.4f} {cy:.4f} m {hx_e:.4f} {cy:.4f} l S")
        vy_s = cy + MARK_OFFSET * vdir
        vy_e = cy + (MARK_OFFSET + MARK_LENGTH) * vdir
        lines.append(f"{cx:.4f} {vy_s:.4f} m {cx:.4f} {vy_e:.4f} l S")
    # Black on top
    lines.append(f"{BLACK_WEIGHT} w")
    lines.append("0 0 0 RG")
    lines.append("2 J")
    for cx, cy, hdir, vdir in corners:
        hx_s = cx + MARK_OFFSET * hdir
        hx_e = cx + (MARK_OFFSET + MARK_LENGTH) * hdir
        lines.append(f"{hx_s:.4f} {cy:.4f} m {hx_e:.4f} {cy:.4f} l S")
        vy_s = cy + MARK_OFFSET * vdir
        vy_e = cy + (MARK_OFFSET + MARK_LENGTH) * vdir
        lines.append(f"{cx:.4f} {vy_s:.4f} m {cx:.4f} {vy_e:.4f} l S")
    lines.append("Q")
    return "\n".join(lines)


def build_blue_gradient_resources(pdf, media_w, media_h):
    c0 = (0.04, 0.06, 0.14)
    c1 = (0.11, 0.20, 0.38)
    fn = Dictionary()
    fn[Name("/FunctionType")] = 2
    fn[Name("/Domain")] = Array([0, 1])
    fn[Name("/C0")] = Array([c0[0], c0[1], c0[2]])
    fn[Name("/C1")] = Array([c1[0], c1[1], c1[2]])
    fn[Name("/N")] = 1

    shading = Dictionary()
    shading[Name("/ShadingType")] = 2
    shading[Name("/ColorSpace")] = Name.DeviceRGB
    shading[Name("/Coords")] = Array([0, 0, 0, media_h])
    shading[Name("/Function")] = pdf.make_indirect(fn)
    shading[Name("/Extend")] = Array([True, True])

    shading_name = Name("/BlueGrad")
    return pdf.make_indirect(shading), shading_name


def build_rounded_rect_path(x, y, w, h, r):
    r = max(0.0, min(r, w / 2.0, h / 2.0))
    if r == 0:
        return f"{x:.4f} {y:.4f} {w:.4f} {h:.4f} re"
    k = 0.5522847498 * r
    x0, y0 = x, y
    x1, y1 = x + w, y + h
    return "\n".join([
        f"{x0 + r:.4f} {y0:.4f} m",
        f"{x1 - r:.4f} {y0:.4f} l",
        f"{x1 - r + k:.4f} {y0:.4f} {x1:.4f} {y0 + r - k:.4f} {x1:.4f} {y0 + r:.4f} c",
        f"{x1:.4f} {y1 - r:.4f} l",
        f"{x1:.4f} {y1 - r + k:.4f} {x1 - r + k:.4f} {y1:.4f} {x1 - r:.4f} {y1:.4f} c",
        f"{x0 + r:.4f} {y1:.4f} l",
        f"{x0 + r - k:.4f} {y1:.4f} {x0:.4f} {y1 - r + k:.4f} {x0:.4f} {y1 - r:.4f} c",
        f"{x0:.4f} {y0 + r:.4f} l",
        f"{x0:.4f} {y0 + r - k:.4f} {x0 + r - k:.4f} {y0:.4f} {x0 + r:.4f} {y0:.4f} c",
        "h",
    ])


def get_home_layout(output_mode):
    if output_mode == "a4_rounded":
        return {
            "paper_w_in": HOME_A4_WIDTH_IN,
            "paper_h_in": HOME_A4_HEIGHT_IN,
            "margin_lr_in": HOME_MARGIN_LEFT_RIGHT_IN,
            "margin_tb_in": HOME_MARGIN_TOP_BOTTOM_IN,
            "layout_mode": "a4_rounded",
        }
    return {
        "paper_w_in": HOME_LETTER_WIDTH_IN,
        "paper_h_in": HOME_LETTER_HEIGHT_IN,
        "margin_lr_in": HOME_MARGIN_LEFT_RIGHT_IN,
        "margin_tb_in": HOME_MARGIN_TOP_BOTTOM_IN,
        "layout_mode": "us_letter_rounded",
    }


def compute_trim_height(orig_w, orig_h, trim_width_in, margins):
    """Derive trim height from original aspect ratio and trim width."""
    trim_w_pt = trim_width_in * PTS_PER_INCH
    ml = margins["left"] * PTS_PER_INCH
    mr = margins["right"] * PTS_PER_INCH
    mt = margins["top"] * PTS_PER_INCH
    mb = margins["bottom"] * PTS_PER_INCH

    safe_w = trim_w_pt - ml - mr
    # scale factor to fit width
    scale = safe_w / orig_w
    scaled_h = orig_h * scale
    # trim height = scaled content + top/bottom margins
    trim_h_pt = scaled_h + mt + mb
    return trim_h_pt / PTS_PER_INCH


def process_page(pdf, page, page_index, trim_width_in, trim_height_in,
                 bleed_in, margins, bg_color=None, bg_style="auto",
                 flip=False, bleed_mode="classic",
                 source_fitz_doc=None, source_page_index=None,
                 ai_provider="print_rescue_local",
                 print_rescue_quality="balanced",
                 lowres_upscale=False,
                 lowres_threshold_dpi=240,
                 force_upscale=False):
    """Re-compose a page with safe margins, bleed, crop marks.

    If flip=True, the entire page content (background, slide, crop marks)
    is rotated 180 degrees.  The media-box size stays the same so that
    front/back crop marks land in identical positions for double-sided
    printing.
    """
    bleed = bleed_in * PTS_PER_INCH
    trim_w = trim_width_in * PTS_PER_INCH
    trim_h = trim_height_in * PTS_PER_INCH
    media_w = trim_w + 2 * bleed
    media_h = trim_h + 2 * bleed

    orig_x0, orig_y0, orig_w, orig_h = get_page_dimensions(page)
    rotation = int(page.obj.get(Name.Rotate, 0)) % 360
    if rotation in (90, 270):
        orig_w, orig_h = orig_h, orig_w

    ml = margins["left"] * PTS_PER_INCH
    mr = margins["right"] * PTS_PER_INCH
    mt = margins["top"] * PTS_PER_INCH
    mb = margins["bottom"] * PTS_PER_INCH

    safe_w = trim_w - ml - mr
    safe_h = trim_h - mt - mb
    if safe_w <= 0 or safe_h <= 0:
        raise ValueError("Margins are too large for the trim area.")

    scale_x = safe_w / orig_w
    scale_y = safe_h / orig_h
    scale = min(scale_x, scale_y)
    scaled_w = orig_w * scale
    scaled_h = orig_h * scale

    safe_x0 = bleed + ml
    safe_y0 = bleed + mb
    x_off = safe_x0 + (safe_w - scaled_w) / 2.0
    y_off = safe_y0 + (safe_h - scaled_h) / 2.0

    if bg_color is None:
        bg_color = detect_background_color(page)

    src_page_obj = None
    if source_fitz_doc is not None and source_page_index is not None and source_page_index >= 0:
        try:
            src_page_obj = source_fitz_doc.load_page(int(source_page_index))
        except Exception:
            src_page_obj = None

    # ── Wrap original page as Form XObject ───────────────────────────
    orig_page_obj = page.obj
    if Name.Contents in orig_page_obj:
        contents = orig_page_obj[Name.Contents]
        if isinstance(contents, pikepdf.Array):
            content_data = b""
            for stream_ref in contents:
                content_data += stream_ref.read_bytes() + b"\n"
        else:
            content_data = contents.read_bytes()
    else:
        content_data = b""

    orig_resources = orig_page_obj.get(Name.Resources, Dictionary())

    form_xobj = pikepdf.Stream(pdf, content_data)
    form_xobj[Name.Type] = Name.XObject
    form_xobj[Name.Subtype] = Name.Form
    form_xobj[Name.BBox] = Array([orig_x0, orig_y0,
                                   orig_x0 + orig_w, orig_y0 + orig_h])
    form_xobj[Name.Resources] = orig_resources
    form_xobj[Name.Matrix] = Array([1, 0, 0, 1, -orig_x0, -orig_y0])

    form_name = Name("/OrigPage")
    new_resources = Dictionary()
    xobj_dict = Dictionary()
    xobj_dict[form_name] = form_xobj
    new_resources[Name.XObject] = xobj_dict
    new_resources[Name.ProcSet] = Array([
        Name.PDF, Name.Text, Name.ImageB, Name.ImageC, Name.ImageI,
    ])

    # ── Build content stream ─────────────────────────────────────────
    cl = []

    # If flip, wrap entire drawing in a 180° rotation around centre
    if flip:
        cl.append("q")
        cl.append(f"-1 0 0 -1 {media_w:.4f} {media_h:.4f} cm")

    # 1. Background
    bleed_provider = None
    use_rescue_bg = False
    if (bleed_mode in {"ai_extend", "print_rescue"} and bleed > 0 and src_page_obj is not None):
        rescue_dpi = _print_rescue_quality_to_dpi(print_rescue_quality)
        if bleed_mode == "ai_extend":
            bg_bytes, bg_w_px, bg_h_px, bleed_provider = _build_ai_bleed_background_jpeg(
                src_page_obj, media_w, media_h, x_off, y_off, scaled_w, scaled_h,
                ai_provider=ai_provider,
            )
            use_rescue_bg = True
        else:
            bg_bytes, bg_w_px, bg_h_px, bleed_provider = _build_fallback_bleed_background_jpeg(
                src_page_obj, media_w, media_h, x_off, y_off, scaled_w, scaled_h,
                provider="print_rescue_local",
                render_dpi=rescue_dpi,
            )
            allow_page_fallback = (os.getenv("PRINT_RESCUE_ENABLE_PAGE_FALLBACK", "true").strip().lower() == "true")
            if allow_page_fallback:
                try:
                    bg_img = Image.open(io.BytesIO(bg_bytes)).convert("RGB")
                    ok = _bleed_quality_ok(
                        bg_img,
                        int(round(x_off)),
                        int(round(y_off)),
                        int(round(x_off + scaled_w)),
                        int(round(y_off + scaled_h)),
                    )
                    use_rescue_bg = bool(ok)
                    if not use_rescue_bg:
                        bleed_provider = "print_rescue_sampled_fallback"
                except Exception:
                    use_rescue_bg = False
                    bleed_provider = "print_rescue_sampled_fallback"
            else:
                use_rescue_bg = True

        if use_rescue_bg:
            bg_stream = pikepdf.Stream(pdf, bg_bytes)
            bg_name = Name("/AIBleedBg")
            bg_stream[Name.Type] = Name.XObject
            bg_stream[Name.Subtype] = Name.Image
            bg_stream[Name.Width] = int(bg_w_px)
            bg_stream[Name.Height] = int(bg_h_px)
            bg_stream[Name.ColorSpace] = Name.DeviceRGB
            bg_stream[Name.BitsPerComponent] = 8
            bg_stream[Name.Filter] = Name.DCTDecode
            xobj_dict[bg_name] = bg_stream
            cl.append("q")
            cl.append(f"{media_w:.4f} 0 0 {media_h:.4f} 0 0 cm")
            cl.append(f"{bg_name} Do")
            cl.append("Q")
            r, g, b = bg_color

    if not use_rescue_bg and bg_style == "blue_gradient":
        shading_obj, shading_name = build_blue_gradient_resources(
            pdf, media_w, media_h)
        shading_dict = Dictionary()
        shading_dict[shading_name] = shading_obj
        new_resources[Name("/Shading")] = shading_dict
        cl.append("q")
        cl.append(f"0 0 {media_w:.4f} {media_h:.4f} re W n")
        cl.append(f"{shading_name} sh")
        cl.append("Q")
        r, g, b = (0.04, 0.06, 0.14)
    elif not use_rescue_bg:
        r, g, b = bg_color
        cl.append("q")
        cl.append(f"{r:.6f} {g:.6f} {b:.6f} rg")
        cl.append(f"0 0 {media_w:.4f} {media_h:.4f} re f")
        cl.append("Q")

    # 2. Place original slide (vector default; optional raster upscale on low-DPI pages)
    content_upscaled = False
    source_content_min_dpi = None
    if (lowres_upscale or force_upscale) and src_page_obj is not None:
        source_content_min_dpi = _fitz_page_content_min_dpi(src_page_obj)
    should_upscale = bool(force_upscale)
    if (not should_upscale and lowres_upscale and src_page_obj is not None
            and source_content_min_dpi is not None and source_content_min_dpi < float(lowres_threshold_dpi)):
        should_upscale = True
    if should_upscale and src_page_obj is not None:
        base_dpi = float(source_content_min_dpi) if source_content_min_dpi is not None else float(lowres_threshold_dpi)
        target_dpi = int(max(float(lowres_threshold_dpi), min(480.0, base_dpi * 2.0)))
        target_w_px = max(8, int(round((scaled_w / 72.0) * target_dpi)))
        target_h_px = max(8, int(round((scaled_h / 72.0) * target_dpi)))
        raster_slide = _fitz_page_to_pil(src_page_obj, target_w_px, target_h_px).convert("RGB")
        if os.getenv("LOWRES_UPSCALE_USE_MODEL", "true").strip().lower() == "true":
            raster_slide = _enhance_with_print_rescue_upscaler(raster_slide)
        slide_out = io.BytesIO()
        raster_slide.save(slide_out, format="JPEG", quality=92, optimize=True, progressive=True)
        slide_stream = pikepdf.Stream(pdf, slide_out.getvalue())
        slide_name = Name("/UpscaledPage")
        slide_stream[Name.Type] = Name.XObject
        slide_stream[Name.Subtype] = Name.Image
        slide_stream[Name.Width] = int(raster_slide.width)
        slide_stream[Name.Height] = int(raster_slide.height)
        slide_stream[Name.ColorSpace] = Name.DeviceRGB
        slide_stream[Name.BitsPerComponent] = 8
        slide_stream[Name.Filter] = Name.DCTDecode
        xobj_dict[slide_name] = slide_stream
        cl.append("q")
        cl.append(f"{scaled_w:.4f} 0 0 {scaled_h:.4f} {x_off:.4f} {y_off:.4f} cm")
        cl.append(f"{slide_name} Do")
        cl.append("Q")
        content_upscaled = True
    else:
        cl.append("q")
        cl.append(f"{scale:.6f} 0 0 {scale:.6f} {x_off:.4f} {y_off:.4f} cm")
        cl.append(f"{form_name} Do")
        cl.append("Q")

    # 3. Trim marks
    cl.append(build_trim_marks_stream(bleed, bleed, trim_w, trim_h))

    if flip:
        cl.append("Q")

    new_content = "\n".join(cl).encode("latin-1")
    new_stream = pikepdf.Stream(pdf, new_content)

    # ── Set page boxes ───────────────────────────────────────────────
    orig_page_obj[Name.Contents] = new_stream
    orig_page_obj[Name.Resources] = new_resources
    orig_page_obj[Name.MediaBox] = Array([0, 0, media_w, media_h])
    orig_page_obj[Name.TrimBox] = Array([bleed, bleed,
                                          bleed + trim_w, bleed + trim_h])
    orig_page_obj[Name.BleedBox] = Array([0, 0, media_w, media_h])

    if Name.CropBox in orig_page_obj:
        del orig_page_obj[Name.CropBox]
    if Name.Rotate in orig_page_obj:
        del orig_page_obj[Name.Rotate]

    return {
        "page": page_index + 1,
        "original_size": f"{orig_w/72:.3f}\" x {orig_h/72:.3f}\"",
        "trim_size": f"{trim_w/72:.4f}\" x {trim_h/72:.4f}\"",
        "media_size": f"{media_w/72:.4f}\" x {media_h/72:.4f}\"",
        "safe_area": f"{safe_w/72:.4f}\" x {safe_h/72:.4f}\"",
        "scale_factor": round(scale, 6),
        "flipped": flip,
        "background_color": "blue gradient" if bg_style == "blue_gradient"
                           else ("ai outpaint (replicate)" if bleed_provider == "replicate"
                                 else "ai outpaint (adobe firefly)" if bleed_provider == "adobe_firefly"
                                 else "print rescue local extend" if bleed_provider == "print_rescue_local"
                                 else "print rescue sampled fallback" if bleed_provider == "print_rescue_sampled_fallback"
                                 else "ai quality guard fallback" if bleed_provider == "quality_guard_fallback"
                                 else "ai fallback extend" if bleed_provider == "fallback"
                                 else "ai bleed extend"
                                 if bleed_mode == "ai_extend"
                                 else f"rgb({r:.2f}, {g:.2f}, {b:.2f})"),
        "layout_mode": "print_ready",
        "bleed_provider": bleed_provider,
        "content_upscaled": content_upscaled,
        "source_content_min_dpi": (round(source_content_min_dpi, 1) if source_content_min_dpi is not None else None),
    }


def process_page_home_paper(pdf, page, page_index, output_mode,
                            bg_color=None, bg_style="auto", flip=False,
                            source_fitz_doc=None, source_page_index=None,
                            print_rescue_quality="balanced",
                            lowres_upscale=False,
                            lowres_threshold_dpi=240,
                            force_upscale=False):
    """Re-compose a page onto 11x8.5 with white margins and rounded art area."""
    cfg = get_home_layout(output_mode)
    media_w = cfg["paper_w_in"] * PTS_PER_INCH
    media_h = cfg["paper_h_in"] * PTS_PER_INCH
    art_x = cfg["margin_lr_in"] * PTS_PER_INCH
    art_y = cfg["margin_tb_in"] * PTS_PER_INCH
    art_w = media_w - 2 * art_x
    art_h = media_h - 2 * art_y
    corner_r = HOME_CORNER_RADIUS_IN * PTS_PER_INCH

    orig_x0, orig_y0, orig_w, orig_h = get_page_dimensions(page)
    rotation = int(page.obj.get(Name.Rotate, 0)) % 360
    if rotation in (90, 270):
        orig_w, orig_h = orig_h, orig_w

    scale = min(art_w / orig_w, art_h / orig_h)
    scaled_w = orig_w * scale
    scaled_h = orig_h * scale
    x_off = art_x + (art_w - scaled_w) / 2.0
    y_off = art_y + (art_h - scaled_h) / 2.0

    if bg_color is None:
        bg_color = detect_background_color(page)

    orig_page_obj = page.obj
    if Name.Contents in orig_page_obj:
        contents = orig_page_obj[Name.Contents]
        if isinstance(contents, pikepdf.Array):
            content_data = b""
            for stream_ref in contents:
                content_data += stream_ref.read_bytes() + b"\n"
        else:
            content_data = contents.read_bytes()
    else:
        content_data = b""

    orig_resources = orig_page_obj.get(Name.Resources, Dictionary())
    form_xobj = pikepdf.Stream(pdf, content_data)
    form_xobj[Name.Type] = Name.XObject
    form_xobj[Name.Subtype] = Name.Form
    form_xobj[Name.BBox] = Array([orig_x0, orig_y0, orig_x0 + orig_w, orig_y0 + orig_h])
    form_xobj[Name.Resources] = orig_resources
    form_xobj[Name.Matrix] = Array([1, 0, 0, 1, -orig_x0, -orig_y0])

    form_name = Name("/OrigPage")
    new_resources = Dictionary()
    xobj_dict = Dictionary()
    xobj_dict[form_name] = form_xobj
    new_resources[Name.XObject] = xobj_dict
    new_resources[Name.ProcSet] = Array([
        Name.PDF, Name.Text, Name.ImageB, Name.ImageC, Name.ImageI,
    ])

    rounded_path = build_rounded_rect_path(art_x, art_y, art_w, art_h, corner_r)
    cl = []
    if flip:
        cl.append("q")
        cl.append(f"-1 0 0 -1 {media_w:.4f} {media_h:.4f} cm")

    # Keep paper white, then paint rounded artwork area.
    use_rescue_fill = False
    if (bg_style == "print_rescue_fill" and source_fitz_doc is not None
            and source_page_index is not None and source_page_index >= 0):
        src_page = source_fitz_doc.load_page(int(source_page_index))
        rescue_dpi = _print_rescue_quality_to_dpi(print_rescue_quality)
        bg_bytes, bg_w_px, bg_h_px, _ = _build_fallback_bleed_background_jpeg(
            src_page,
            art_w,
            art_h,
            x_off - art_x,
            y_off - art_y,
            scaled_w,
            scaled_h,
            provider="print_rescue_local",
            render_dpi=rescue_dpi,
        )
        allow_page_fallback = (os.getenv("PRINT_RESCUE_ENABLE_PAGE_FALLBACK", "true").strip().lower() == "true")
        if allow_page_fallback:
            try:
                bg_img = Image.open(io.BytesIO(bg_bytes)).convert("RGB")
                ok = _bleed_quality_ok(
                    bg_img,
                    int(round(x_off - art_x)),
                    int(round(y_off - art_y)),
                    int(round(x_off - art_x + scaled_w)),
                    int(round(y_off - art_y + scaled_h)),
                )
                use_rescue_fill = bool(ok)
            except Exception:
                use_rescue_fill = False
        else:
            use_rescue_fill = True
        if use_rescue_fill:
            bg_stream = pikepdf.Stream(pdf, bg_bytes)
            bg_name = Name("/HomeRescueBg")
            bg_stream[Name.Type] = Name.XObject
            bg_stream[Name.Subtype] = Name.Image
            bg_stream[Name.Width] = int(bg_w_px)
            bg_stream[Name.Height] = int(bg_h_px)
            bg_stream[Name.ColorSpace] = Name.DeviceRGB
            bg_stream[Name.BitsPerComponent] = 8
            bg_stream[Name.Filter] = Name.DCTDecode
            xobj_dict[bg_name] = bg_stream
            cl.append("q")
            cl.append(rounded_path)
            cl.append("W n")
            cl.append(f"{art_w:.4f} 0 0 {art_h:.4f} {art_x:.4f} {art_y:.4f} cm")
            cl.append(f"{bg_name} Do")
            cl.append("Q")
            r, g, b = bg_color
    if not use_rescue_fill and bg_style == "blue_gradient":
        shading_obj, shading_name = build_blue_gradient_resources(pdf, media_w, media_h)
        shading_dict = Dictionary()
        shading_dict[shading_name] = shading_obj
        new_resources[Name("/Shading")] = shading_dict
        cl.append("q")
        cl.append(rounded_path)
        cl.append("W n")
        cl.append(f"{shading_name} sh")
        cl.append("Q")
        r, g, b = (0.04, 0.06, 0.14)
    elif not use_rescue_fill:
        r, g, b = bg_color
        cl.append("q")
        cl.append(f"{r:.6f} {g:.6f} {b:.6f} rg")
        cl.append(rounded_path)
        cl.append("f")
        cl.append("Q")

    # Clip placed slide to rounded corners so the white paper margins show.
    content_upscaled = False
    source_content_min_dpi = None
    src_page_obj = None
    if source_fitz_doc is not None and source_page_index is not None and source_page_index >= 0:
        try:
            src_page_obj = source_fitz_doc.load_page(int(source_page_index))
        except Exception:
            src_page_obj = None
    if (lowres_upscale or force_upscale) and src_page_obj is not None:
        source_content_min_dpi = _fitz_page_content_min_dpi(src_page_obj)
    cl.append("q")
    cl.append(rounded_path)
    cl.append("W n")
    should_upscale = bool(force_upscale)
    if (not should_upscale and lowres_upscale and src_page_obj is not None
            and source_content_min_dpi is not None and source_content_min_dpi < float(lowres_threshold_dpi)):
        should_upscale = True
    if should_upscale and src_page_obj is not None:
        base_dpi = float(source_content_min_dpi) if source_content_min_dpi is not None else float(lowres_threshold_dpi)
        target_dpi = int(max(float(lowres_threshold_dpi), min(480.0, base_dpi * 2.0)))
        target_w_px = max(8, int(round((scaled_w / 72.0) * target_dpi)))
        target_h_px = max(8, int(round((scaled_h / 72.0) * target_dpi)))
        raster_slide = _fitz_page_to_pil(src_page_obj, target_w_px, target_h_px).convert("RGB")
        if os.getenv("LOWRES_UPSCALE_USE_MODEL", "true").strip().lower() == "true":
            raster_slide = _enhance_with_print_rescue_upscaler(raster_slide)
        slide_out = io.BytesIO()
        raster_slide.save(slide_out, format="JPEG", quality=92, optimize=True, progressive=True)
        slide_stream = pikepdf.Stream(pdf, slide_out.getvalue())
        slide_name = Name("/UpscaledPage")
        slide_stream[Name.Type] = Name.XObject
        slide_stream[Name.Subtype] = Name.Image
        slide_stream[Name.Width] = int(raster_slide.width)
        slide_stream[Name.Height] = int(raster_slide.height)
        slide_stream[Name.ColorSpace] = Name.DeviceRGB
        slide_stream[Name.BitsPerComponent] = 8
        slide_stream[Name.Filter] = Name.DCTDecode
        xobj_dict[slide_name] = slide_stream
        cl.append(f"{scaled_w:.4f} 0 0 {scaled_h:.4f} {x_off:.4f} {y_off:.4f} cm")
        cl.append(f"{slide_name} Do")
        content_upscaled = True
    else:
        cl.append(f"{scale:.6f} 0 0 {scale:.6f} {x_off:.4f} {y_off:.4f} cm")
        cl.append(f"{form_name} Do")
    cl.append("Q")

    if flip:
        cl.append("Q")

    new_content = "\n".join(cl).encode("latin-1")
    orig_page_obj[Name.Contents] = pikepdf.Stream(pdf, new_content)
    orig_page_obj[Name.Resources] = new_resources
    orig_page_obj[Name.MediaBox] = Array([0, 0, media_w, media_h])
    orig_page_obj[Name.TrimBox] = Array([art_x, art_y, art_x + art_w, art_y + art_h])
    orig_page_obj[Name.BleedBox] = Array([art_x, art_y, art_x + art_w, art_y + art_h])
    if Name.CropBox in orig_page_obj:
        del orig_page_obj[Name.CropBox]
    if Name.Rotate in orig_page_obj:
        del orig_page_obj[Name.Rotate]

    return {
        "page": page_index + 1,
        "original_size": f"{orig_w/72:.3f}\" x {orig_h/72:.3f}\"",
        "trim_size": f"{art_w/72:.4f}\" x {art_h/72:.4f}\"",
        "media_size": f"{media_w/72:.4f}\" x {media_h/72:.4f}\"",
        "safe_area": f"{art_w/72:.4f}\" x {art_h/72:.4f}\"",
        "scale_factor": round(scale, 6),
        "flipped": flip,
        "background_color": "blue gradient" if bg_style == "blue_gradient"
                           else ("print rescue fill" if bg_style == "print_rescue_fill"
                                 else f"rgb({r:.2f}, {g:.2f}, {b:.2f})"),
        "layout_mode": cfg["layout_mode"],
        "content_upscaled": content_upscaled,
        "source_content_min_dpi": (round(source_content_min_dpi, 1) if source_content_min_dpi is not None else None),
    }


def add_blank_page(pdf, trim_width_in, trim_height_in, bleed_in, insert_at):
    """Insert a blank white page at the given index with proper boxes and crop marks."""
    bleed = bleed_in * PTS_PER_INCH
    trim_w = trim_width_in * PTS_PER_INCH
    trim_h = trim_height_in * PTS_PER_INCH
    media_w = trim_w + 2 * bleed
    media_h = trim_h + 2 * bleed

    # Create a temporary single-page PDF, then copy the page in
    blank_pdf = Pdf.new()
    blank_pdf.add_blank_page(page_size=(media_w, media_h))
    blank_page = blank_pdf.pages[0]

    cl = []
    cl.append("q")
    cl.append("1 1 1 rg")
    cl.append(f"0 0 {media_w:.4f} {media_h:.4f} re f")
    cl.append("Q")
    cl.append(build_trim_marks_stream(bleed, bleed, trim_w, trim_h))

    content = "\n".join(cl).encode("latin-1")
    blank_page.obj[Name.Contents] = blank_pdf.make_indirect(
        pikepdf.Stream(blank_pdf, content))
    blank_page.obj[Name.TrimBox] = Array([bleed, bleed,
                                           bleed + trim_w, bleed + trim_h])
    blank_page.obj[Name.BleedBox] = Array([0, 0, media_w, media_h])

    # Copy into target PDF
    pdf.pages.insert(insert_at, blank_page)
    # Now set boxes on the inserted page (in the target pdf)
    inserted = pdf.pages[insert_at]
    inserted.obj[Name.TrimBox] = Array([bleed, bleed,
                                         bleed + trim_w, bleed + trim_h])
    inserted.obj[Name.BleedBox] = Array([0, 0, media_w, media_h])

    blank_pdf.close()


def add_blank_page_home(pdf, insert_at, output_mode):
    cfg = get_home_layout(output_mode)
    media_w = cfg["paper_w_in"] * PTS_PER_INCH
    media_h = cfg["paper_h_in"] * PTS_PER_INCH
    art_x = cfg["margin_lr_in"] * PTS_PER_INCH
    art_y = cfg["margin_tb_in"] * PTS_PER_INCH
    art_w = media_w - 2 * art_x
    art_h = media_h - 2 * art_y

    blank_pdf = Pdf.new()
    blank_pdf.add_blank_page(page_size=(media_w, media_h))
    blank_page = blank_pdf.pages[0]
    blank_page.obj[Name.TrimBox] = Array([art_x, art_y, art_x + art_w, art_y + art_h])
    blank_page.obj[Name.BleedBox] = Array([art_x, art_y, art_x + art_w, art_y + art_h])

    pdf.pages.insert(insert_at, blank_page)
    inserted = pdf.pages[insert_at]
    inserted.obj[Name.TrimBox] = Array([art_x, art_y, art_x + art_w, art_y + art_h])
    inserted.obj[Name.BleedBox] = Array([art_x, art_y, art_x + art_w, art_y + art_h])
    blank_pdf.close()


def render_thumbnails(pdf_path, dpi=THUMBNAIL_DPI, progress_cb=None):
    """Return list of base64-encoded PNG thumbnails, one per page."""
    doc = fitz.open(pdf_path)
    thumbs = []
    total = len(doc)
    for i, page in enumerate(doc):
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        png_data = pix.tobytes("png")
        b64 = base64.b64encode(png_data).decode("ascii")
        thumbs.append(b64)
        if progress_cb is not None:
            progress_cb("thumbnails", i + 1, total)
    doc.close()
    return thumbs


def analyze_bitmap_dpi(pdf_path):
    """Estimate effective raster DPI for page content and generated bleed fill."""
    doc = fitz.open(pdf_path)
    out = []
    min_area_in2 = float(os.getenv("DPI_MIN_IMAGE_AREA_IN2", "0.04"))
    generated_names = {"AIBleedBg", "HomeRescueBg"}
    detail_limit = max(1, int(os.getenv("DPI_IMAGE_DETAIL_LIMIT", "8")))
    for page in doc:
        page_result = {
            "has_raster": False,
            "raster_count": 0,
            "raster_min_dpi": None,
            "raster_avg_dpi": None,
            "content_has_raster": False,
            "content_raster_count": 0,
            "content_raster_min_dpi": None,
            "content_raster_avg_dpi": None,
            "content_raster_dpi_samples": [],
            "content_images": [],
            "bleed_fill_raster_count": 0,
            "bleed_fill_raster_min_dpi": None,
            "bleed_fill_raster_avg_dpi": None,
            "has_vector_content": False,
            "content_type": "unknown",
        }
        dpis_all = []
        dpis_content = []
        dpis_fill = []
        seen_all = set()
        seen_content = set()
        seen_fill = set()
        try:
            images = page.get_images(full=True)
        except Exception:
            images = []
        for info in images:
            if len(info) < 8:
                continue
            xref = int(info[0])
            px_w = float(info[2] or 0)
            px_h = float(info[3] or 0)
            image_name = str(info[7] or "").lstrip("/")
            is_generated_fill = image_name in generated_names
            if xref <= 0 or px_w <= 0 or px_h <= 0:
                continue
            try:
                rects = page.get_image_rects(xref)
            except Exception:
                rects = []
            for rect in rects:
                w_pt = float(rect.width)
                h_pt = float(rect.height)
                if w_pt <= 0 or h_pt <= 0:
                    continue
                w_in = w_pt / 72.0
                h_in = h_pt / 72.0
                area = w_in * h_in
                if area < min_area_in2:
                    continue
                dpi_x = px_w / max(1e-6, w_in)
                dpi_y = px_h / max(1e-6, h_in)
                eff_dpi = float(min(dpi_x, dpi_y))
                # Deduplicate repeated same xref/size placements.
                sig = (xref, round(w_in, 4), round(h_in, 4))
                if sig in seen_all:
                    continue
                seen_all.add(sig)
                dpis_all.append(eff_dpi)

                if is_generated_fill:
                    if sig not in seen_fill:
                        seen_fill.add(sig)
                        dpis_fill.append(eff_dpi)
                else:
                    if sig not in seen_content:
                        seen_content.add(sig)
                        dpis_content.append(eff_dpi)
                        # Optional per-image details for UI expanders.
                        if len(page_result["content_images"]) < detail_limit:
                            thumb_b64 = None
                            px_w_i = int(round(px_w))
                            px_h_i = int(round(px_h))
                            try:
                                pix = fitz.Pixmap(doc, xref)
                                if pix.alpha or pix.n > 4:
                                    pix = fitz.Pixmap(fitz.csRGB, pix)
                                img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
                                max_side = 56
                                ratio = min(1.0, max_side / max(1, max(img.width, img.height)))
                                tw = max(12, int(round(img.width * ratio)))
                                th = max(12, int(round(img.height * ratio)))
                                img = img.resize((tw, th), Image.Resampling.LANCZOS)
                                buf = io.BytesIO()
                                img.save(buf, format="PNG")
                                thumb_b64 = base64.b64encode(buf.getvalue()).decode("ascii")
                            except Exception:
                                thumb_b64 = None
                            page_result["content_images"].append({
                                "index": len(page_result["content_images"]) + 1,
                                "dpi": round(eff_dpi, 1),
                                "width_px": px_w_i,
                                "height_px": px_h_i,
                                "thumb_b64": thumb_b64,
                            })
        if dpis_all:
            page_result["has_raster"] = True
            page_result["raster_count"] = len(dpis_all)
            page_result["raster_min_dpi"] = round(min(dpis_all), 1)
            page_result["raster_avg_dpi"] = round(float(sum(dpis_all) / len(dpis_all)), 1)
        if dpis_content:
            page_result["content_has_raster"] = True
            page_result["content_raster_count"] = len(dpis_content)
            page_result["content_raster_min_dpi"] = round(min(dpis_content), 1)
            page_result["content_raster_avg_dpi"] = round(float(sum(dpis_content) / len(dpis_content)), 1)
            samples = sorted(float(x) for x in dpis_content)
            page_result["content_raster_dpi_samples"] = [round(x, 1) for x in samples[:4]]
        # Detect vector content (text/paths) separately from raster images.
        try:
            has_text = bool(page.get_text("words"))
        except Exception:
            has_text = False
        try:
            has_drawings = bool(page.get_drawings())
        except Exception:
            has_drawings = False
        has_vector = bool(has_text or has_drawings)
        page_result["has_vector_content"] = has_vector
        if has_vector and page_result["content_has_raster"]:
            page_result["content_type"] = "mixed"
        elif has_vector:
            page_result["content_type"] = "vector_only"
        elif page_result["content_has_raster"]:
            page_result["content_type"] = "raster_only"
        else:
            page_result["content_type"] = "unknown"
        if dpis_fill:
            page_result["bleed_fill_raster_count"] = len(dpis_fill)
            page_result["bleed_fill_raster_min_dpi"] = round(min(dpis_fill), 1)
            page_result["bleed_fill_raster_avg_dpi"] = round(float(sum(dpis_fill) / len(dpis_fill)), 1)
        out.append(page_result)
    doc.close()
    return out


def render_flipbook_thumbnails(processed_path, pages_info, bleed_in,
                               flip_even_pages=False, dpi=150):
    """Render thumbnails cropped to trim area for flipbook viewing.

    The preview uses pages as they exist in the processed PDF. The UI then
    applies physical page-turn orientation (top page flipped over binding).
    """
    bleed = bleed_in * PTS_PER_INCH

    src = Pdf.open(processed_path)
    bound = Pdf.new()

    for i, src_page in enumerate(src.pages):
        bound.pages.append(src_page)
        page = bound.pages[-1]

        mbox = [float(v) for v in page.mediabox]
        layout_mode = (
            pages_info[i].get("layout_mode")
            if i < len(pages_info) and isinstance(pages_info[i], dict)
            else "print_ready"
        )
        if layout_mode in ("us_letter_rounded", "a4_rounded"):
            # Keep full page so user can see white printable margins.
            page.obj[Name.MediaBox] = Array([mbox[0], mbox[1], mbox[2], mbox[3]])
        else:
            page.obj[Name.MediaBox] = Array([
                mbox[0] + bleed, mbox[1] + bleed,
                mbox[2] - bleed, mbox[3] - bleed,
            ])

        for box_name in [Name.TrimBox, Name.BleedBox, Name.CropBox]:
            if box_name in page.obj:
                del page.obj[box_name]

        # Normalize odd-index pages when alternate-page flip is enabled so the
        # flipbook can display a readable top page in that mode.
        if flip_even_pages and (i % 2 == 1):
            page.obj[Name.Rotate] = 180

    temp_path = processed_path + ".flipbook.pdf"
    bound.save(temp_path)
    bound.close()
    src.close()

    thumbs = render_thumbnails(temp_path, dpi=dpi)
    try:
        os.remove(temp_path)
    except OSError:
        pass
    return thumbs


def verify_page(page, trim_width_in, trim_height_in, bleed_in, output_mode="print_ready"):
    bleed = bleed_in * PTS_PER_INCH
    mbox = [float(v) for v in page.mediabox]
    tbox = [float(v) for v in page.obj.get(Name.TrimBox, page.mediabox)]

    media_w = mbox[2] - mbox[0]
    media_h = mbox[3] - mbox[1]
    trim_w = tbox[2] - tbox[0]
    trim_h = tbox[3] - tbox[1]

    bleed_l = tbox[0] - mbox[0]
    bleed_r = mbox[2] - tbox[2]
    bleed_b = tbox[1] - mbox[1]
    bleed_t = mbox[3] - tbox[3]

    if output_mode in ("us_letter_rounded", "a4_rounded"):
        cfg = get_home_layout(output_mode)
        expected_w = cfg["paper_w_in"] * PTS_PER_INCH
        expected_h = cfg["paper_h_in"] * PTS_PER_INCH
        expected_art_x = cfg["margin_lr_in"] * PTS_PER_INCH
        expected_art_y = cfg["margin_tb_in"] * PTS_PER_INCH
        checks = [
            ("Media size = 11.000\" x 8.500\"",
             abs(media_w - expected_w) < 0.01 and abs(media_h - expected_h) < 0.01),
            ("Left margin = 0.250\"",
             abs((tbox[0] - mbox[0]) / 72 - cfg["margin_lr_in"]) < 0.01),
            ("Right margin = 0.250\"",
             abs((mbox[2] - tbox[2]) / 72 - cfg["margin_lr_in"]) < 0.01),
            ("Top margin = 0.500\"",
             abs((mbox[3] - tbox[3]) / 72 - cfg["margin_tb_in"]) < 0.01),
            ("Bottom margin = 0.500\"",
             abs((tbox[1] - mbox[1]) / 72 - cfg["margin_tb_in"]) < 0.01),
            ("TrimBox present", Name.TrimBox in page.obj),
            ("BleedBox present", Name.BleedBox in page.obj),
            ("TrimBox origin correct",
             abs(tbox[0] - expected_art_x) < 0.01 and abs(tbox[1] - expected_art_y) < 0.01),
        ]
        return [{"label": label, "pass": ok} for label, ok in checks]

    checks = [
        (f"Trim width = {trim_width_in:.3f}\"",
         abs(trim_w / 72 - trim_width_in) < 0.01),
        (f"Trim height = {trim_height_in:.3f}\"",
         abs(trim_h / 72 - trim_height_in) < 0.01),
        (f"Bleed Left = {bleed_in}\"",
         abs(bleed_l / 72 - bleed_in) < 0.001),
        (f"Bleed Right = {bleed_in}\"",
         abs(bleed_r / 72 - bleed_in) < 0.001),
        (f"Bleed Top = {bleed_in}\"",
         abs(bleed_t / 72 - bleed_in) < 0.001),
        (f"Bleed Bottom = {bleed_in}\"",
         abs(bleed_b / 72 - bleed_in) < 0.001),
        ("MediaBox width correct",
         abs(media_w - trim_w - 2 * bleed) < 0.01),
        ("MediaBox height correct",
         abs(media_h - trim_h - 2 * bleed) < 0.01),
        ("TrimBox present", Name.TrimBox in page.obj),
        ("BleedBox present", Name.BleedBox in page.obj),
    ]
    return [{"label": label, "pass": ok} for label, ok in checks]


def process_pdf_file(input_path, output_path, trim_width_in, bleed_in,
                     margins, bg_style="auto",
                     insert_blank_after_cover=False,
                     delete_pages=None,
                     page_order=None,
                     flip_even_pages=False,
                     progress_cb=None,
                     output_mode="print_ready",
                     bleed_mode="classic",
                     ai_provider="print_rescue_local",
                     print_rescue_quality="balanced",
                     lowres_upscale=False,
                     lowres_threshold_dpi=240,
                     forced_upscale_pages=None,
                     forced_upscale_images=None):
    """Process a PDF: auto-height, optional blank page, page deletion,
    alternating page flip for double-sided binding.

    Parameters
    ----------
    delete_pages : list[int] or None
        0-based page indices to remove *before* processing.
    flip_even_pages : bool
        If True, every other page (0-based index 1, 3, 5, …) in the
        *final* output is rotated 180° so crop marks register correctly
        for double-sided spiral/comb binding.
    """
    pdf = Pdf.open(input_path)

    # ── Build final source-page order (delete + optional reorder) ─────
    orig_count = len(pdf.pages)
    delete_set = set(int(x) for x in (delete_pages or []) if isinstance(x, int))
    final_source_indices = []
    if page_order:
        seen = set()
        for raw in page_order:
            try:
                idx = int(raw)
            except Exception:
                continue
            if idx in seen or idx in delete_set:
                continue
            if 0 <= idx < orig_count:
                seen.add(idx)
                final_source_indices.append(idx)
    else:
        final_source_indices = [i for i in range(orig_count) if i not in delete_set]

    if not final_source_indices:
        raise ValueError("No pages left after selection.")

    # Rebuild PDF in selected order.
    ordered_pdf = Pdf.new()
    for src_idx in final_source_indices:
        ordered_pdf.pages.append(pdf.pages[src_idx])
    pdf.close()
    pdf = ordered_pdf

    # ── Determine trim height from the first page's aspect ratio ─────
    first_page = pdf.pages[0]
    _, _, orig_w, orig_h = get_page_dimensions(first_page)
    rotation = int(first_page.obj.get(Name.Rotate, 0)) % 360
    if rotation in (90, 270):
        orig_w, orig_h = orig_h, orig_w
    if output_mode in ("us_letter_rounded", "a4_rounded"):
        cfg = get_home_layout(output_mode)
        trim_width_in = cfg["paper_w_in"] - 2 * cfg["margin_lr_in"]
        trim_height_in = cfg["paper_h_in"] - 2 * cfg["margin_tb_in"]
        bleed_in = 0.0
    else:
        trim_height_in = compute_trim_height(orig_w, orig_h, trim_width_in, margins)

    # ── Insert blank page after cover (index 0) ─────────────────────
    if insert_blank_after_cover and len(pdf.pages) >= 1:
        if output_mode in ("us_letter_rounded", "a4_rounded"):
            add_blank_page_home(pdf, insert_at=1, output_mode=output_mode)
        else:
            add_blank_page(pdf, trim_width_in, trim_height_in, bleed_in, insert_at=1)

    # ── Process each page ────────────────────────────────────────────
    total_pages = len(pdf.pages)
    forced_set = set(int(x) for x in (forced_upscale_pages or []) if isinstance(x, int))
    forced_image_map = {}
    for item in (forced_upscale_images or []):
        if not isinstance(item, dict):
            continue
        try:
            pnum = int(item.get("page", 0))
        except Exception:
            pnum = 0
        if pnum <= 0:
            continue
        imgs = item.get("images", [])
        if not isinstance(imgs, list):
            continue
        parsed = []
        for idx in imgs:
            try:
                v = int(idx)
            except Exception:
                continue
            if v > 0:
                parsed.append(v)
        if parsed:
            forced_image_map[pnum - 1] = sorted(set(parsed))
    source_doc = fitz.open(input_path)
    kept_indices = list(final_source_indices)
    ai_enabled = (bleed_mode in {"ai_extend", "print_rescue"} and output_mode == "print_ready" and bleed_in > 0)
    if progress_cb is not None:
        progress_cb("ai_bleed" if ai_enabled else "processing", 0, total_pages)
    pages_info = []
    try:
        for i, page in enumerate(pdf.pages):
            # Check if this is the inserted blank page
            is_blank = (insert_blank_after_cover and i == 1)
            if is_blank:
                # Already set up by add_blank_page; just need to maybe flip
                if flip_even_pages and (i % 2 == 1):
                    _flip_existing_page(pdf, page, trim_width_in, trim_height_in, bleed_in)
                pages_info.append({
                    "page": i + 1,
                    "original_size": "blank",
                    "trim_size": f"{trim_width_in:.4f}\" x {trim_height_in:.4f}\"",
                    "media_size": f"{trim_width_in + 2*bleed_in:.4f}\" x {trim_height_in + 2*bleed_in:.4f}\"",
                    "safe_area": "blank",
                    "scale_factor": 1.0,
                    "flipped": flip_even_pages and (i % 2 == 1),
                    "background_color": "rgb(1.00, 1.00, 1.00)",
                    "is_blank": True,
                    "layout_mode": output_mode,
                })
                if progress_cb is not None:
                    progress_cb("ai_bleed" if ai_enabled else "processing", i + 1, total_pages)
                continue

            flip = flip_even_pages and (i % 2 == 1)
            force_upscale = (i in forced_set)
            source_pos = i if not insert_blank_after_cover else (i if i == 0 else i - 1)
            source_page_index = kept_indices[source_pos] if 0 <= source_pos < len(kept_indices) else None
            page_forced_from_images = (i in forced_image_map and len(forced_image_map.get(i, [])) > 0)
            # Direct xref image replacement can break some PDFs (masked images,
            # unusual color spaces). Keep it opt-in; default to safer page-level
            # upscale path so images are never dropped.
            allow_direct_image_replace = os.getenv("ENABLE_DIRECT_IMAGE_REPLACE", "false").lower() in ("1", "true", "yes", "on")
            if allow_direct_image_replace and source_page_index is not None and i in forced_image_map:
                try:
                    src_page_obj = source_doc.load_page(int(source_page_index))
                    infos = src_page_obj.get_images(full=True)
                    xrefs = []
                    seen_xrefs = set()
                    for inf in infos:
                        if len(inf) < 1:
                            continue
                        xr = int(inf[0])
                        if xr <= 0 or xr in seen_xrefs:
                            continue
                        seen_xrefs.add(xr)
                        xrefs.append(xr)
                    for one_based_idx in forced_image_map[i]:
                        arr_idx = one_based_idx - 1
                        if 0 <= arr_idx < len(xrefs):
                            _upscale_source_image_xref(source_doc, xrefs[arr_idx], scale=2)
                except Exception:
                    pass
            # Reliability fallback: image-targeted upscale request always forces
            # an upscale render on this page even if direct image-object replace
            # is unsupported by the source PDF structure.
            force_upscale = force_upscale or page_forced_from_images
            if output_mode in ("us_letter_rounded", "a4_rounded"):
                info = process_page_home_paper(
                    pdf, page, i, output_mode=output_mode, bg_style=bg_style, flip=flip,
                    source_fitz_doc=source_doc, source_page_index=source_page_index,
                    print_rescue_quality=print_rescue_quality,
                    lowres_upscale=lowres_upscale,
                    lowres_threshold_dpi=lowres_threshold_dpi,
                    force_upscale=force_upscale,
                )
            else:
                info = process_page(pdf, page, i,
                                    trim_width_in, trim_height_in,
                                    bleed_in, margins,
                                    bg_style=bg_style,
                                    flip=flip,
                                    bleed_mode=bleed_mode,
                                    source_fitz_doc=source_doc,
                                    source_page_index=source_page_index,
                                    ai_provider=ai_provider,
                                    print_rescue_quality=print_rescue_quality,
                                    lowres_upscale=lowres_upscale,
                                    lowres_threshold_dpi=lowres_threshold_dpi,
                                    force_upscale=force_upscale)
            pages_info.append(info)
            if progress_cb is not None:
                progress_cb("ai_bleed" if ai_enabled else "processing", i + 1, total_pages)
    finally:
        source_doc.close()

    pdf.save(output_path, linearize=False)
    pdf.close()

    # ── Verify ───────────────────────────────────────────────────────
    pdf2 = Pdf.open(output_path)
    verification = []
    if progress_cb is not None:
        progress_cb("verifying", 0, len(pdf2.pages))
    for i, page in enumerate(pdf2.pages):
        checks = verify_page(
            page, trim_width_in, trim_height_in, bleed_in, output_mode=output_mode
        )
        all_pass = all(c["pass"] for c in checks)
        verification.append({"page": i + 1, "checks": checks, "all_pass": all_pass})
        if progress_cb is not None:
            progress_cb("verifying", i + 1, len(pdf2.pages))
    pdf2.close()

    # ── Thumbnails + raster DPI analysis ────────────────────────────
    thumbnails = render_thumbnails(output_path, progress_cb=progress_cb)
    dpi_report = analyze_bitmap_dpi(output_path)
    for i, info in enumerate(pages_info):
        if i < len(dpi_report):
            info.update(dpi_report[i])

    return {
        "pages": pages_info,
        "verification": verification,
        "thumbnails": thumbnails,
        "trim_width_in": trim_width_in,
        "trim_height_in": round(trim_height_in, 4),
    }


def _flip_existing_page(pdf, page, trim_width_in, trim_height_in, bleed_in):
    """Rotate an already-composed page 180° by wrapping its contents."""
    mbox = [float(v) for v in page.obj.get(Name.MediaBox, page.mediabox)]
    media_w = mbox[2] - mbox[0]
    media_h = mbox[3] - mbox[1]

    obj = page.obj
    # Read existing content
    if Name.Contents in obj:
        contents = obj[Name.Contents]
        if isinstance(contents, pikepdf.Array):
            existing = b""
            for ref in contents:
                existing += ref.read_bytes() + b"\n"
        else:
            existing = contents.read_bytes()
    else:
        existing = b""

    # Wrap in 180° rotation
    wrapped = (
        f"q\n-1 0 0 -1 {media_w:.4f} {media_h:.4f} cm\n".encode("latin-1")
        + existing
        + b"\nQ"
    )
    obj[Name.Contents] = pikepdf.Stream(pdf, wrapped)


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return send_from_directory("static", "index.html")


@app.route("/api/preview", methods=["POST"])
def api_preview():
    """Upload a PDF and return thumbnails of the *original* pages (before
    processing) so the user can choose which pages to delete / where to
    insert a blank."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    file = request.files["file"]
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Please upload a PDF file"}), 400

    job_id = str(uuid.uuid4())[:8]
    job_dir = os.path.join(UPLOAD_DIR, job_id)
    os.makedirs(job_dir, exist_ok=True)
    input_path = os.path.join(job_dir, "input.pdf")
    file.save(input_path)

    # Get page info
    pdf = Pdf.open(input_path)
    pages_info = []
    for i, page in enumerate(pdf.pages):
        _, _, w, h = get_page_dimensions(page)
        rot = int(page.obj.get(Name.Rotate, 0)) % 360
        if rot in (90, 270):
            w, h = h, w
        pages_info.append({
            "index": i,
            "width_in": round(w / 72, 3),
            "height_in": round(h / 72, 3),
        })
    pdf.close()

    thumbs = render_thumbnails(input_path, dpi=72)

    return jsonify({
        "job_id": job_id,
        "filename": file.filename,
        "page_count": len(pages_info),
        "pages": pages_info,
        "thumbnails": thumbs,
    })


@app.route("/api/process", methods=["POST"])
def api_process():
    """Process the PDF with all options."""
    job_id = request.form.get("job_id")
    if not job_id or not job_id.isalnum() or len(job_id) != 8:
        return jsonify({"error": "Invalid job ID. Upload a file first."}), 400

    input_path = os.path.join(UPLOAD_DIR, job_id, "input.pdf")
    if not os.path.isfile(input_path):
        return jsonify({"error": "File not found. Please re-upload."}), 404

    output_path = os.path.join(UPLOAD_DIR, job_id, "output.pdf")

    bleed = float(request.form.get("bleed", DEFAULT_BLEED_IN))
    trim_width = float(request.form.get("trim_width", DEFAULT_TRIM_WIDTH_IN))

    margins = {
        "left": float(request.form.get("margin_left", DEFAULT_MARGIN_LEFT)),
        "right": float(request.form.get("margin_right", DEFAULT_MARGIN_RIGHT)),
        "top": float(request.form.get("margin_top", DEFAULT_MARGIN_TOP)),
        "bottom": float(request.form.get("margin_bottom", DEFAULT_MARGIN_BOTTOM)),
    }

    bg_style = request.form.get("bg_style", "auto")
    if bg_style not in ("auto", "blue_gradient", "print_rescue_fill"):
        bg_style = "auto"
    bleed_mode = request.form.get("bleed_mode", "classic")
    if bleed_mode not in ("classic", "print_rescue", "ai_extend"):
        bleed_mode = "classic"
    # Replicate AI option is temporarily disabled in UI and backend.
    ai_provider = "print_rescue_local"
    if bleed_mode == "ai_extend":
        bleed_mode = "print_rescue"
    print_rescue_quality = (request.form.get("print_rescue_quality", "balanced") or "balanced").strip().lower()
    if print_rescue_quality not in ("balanced", "high", "ultra"):
        print_rescue_quality = "balanced"
    lowres_upscale = request.form.get("lowres_upscale", "false") == "true"
    lowres_threshold_dpi = float(request.form.get("lowres_threshold_dpi", "240") or "240")
    lowres_threshold_dpi = max(180.0, min(480.0, lowres_threshold_dpi))

    insert_blank = request.form.get("insert_blank_after_cover", "false") == "true"
    flip_even = request.form.get("flip_even_pages", "false") == "true"
    output_mode = request.form.get("output_mode", "print_ready")
    if output_mode not in ("print_ready", "us_letter_rounded", "a4_rounded"):
        output_mode = "print_ready"
    if output_mode in ("us_letter_rounded", "a4_rounded"):
        cfg = get_home_layout(output_mode)
        bleed = 0.0
        bleed_mode = "classic"
        margins = {
            "left": cfg["margin_lr_in"],
            "right": cfg["margin_lr_in"],
            "top": cfg["margin_tb_in"],
            "bottom": cfg["margin_tb_in"],
        }
    if bleed_mode == "print_rescue":
        ai_provider = "print_rescue_local"

    # Parse delete_pages JSON array
    delete_pages_raw = request.form.get("delete_pages", "[]")
    try:
        delete_pages = json.loads(delete_pages_raw)
        if not isinstance(delete_pages, list):
            delete_pages = []
        delete_pages = [int(x) for x in delete_pages]
    except (json.JSONDecodeError, ValueError):
        delete_pages = []

    page_order_raw = request.form.get("page_order", "[]")
    try:
        page_order = json.loads(page_order_raw)
        if not isinstance(page_order, list):
            page_order = []
        page_order = [int(x) for x in page_order]
    except (json.JSONDecodeError, ValueError):
        page_order = []

    forced_upscale_pages_raw = request.form.get("forced_upscale_pages", "[]")
    try:
        forced_upscale_pages = json.loads(forced_upscale_pages_raw)
        if not isinstance(forced_upscale_pages, list):
            forced_upscale_pages = []
        forced_upscale_pages = [int(x) for x in forced_upscale_pages]
    except (json.JSONDecodeError, ValueError):
        forced_upscale_pages = []

    forced_upscale_images_raw = request.form.get("forced_upscale_images", "[]")
    try:
        forced_upscale_images = json.loads(forced_upscale_images_raw)
        if not isinstance(forced_upscale_images, list):
            forced_upscale_images = []
    except (json.JSONDecodeError, ValueError):
        forced_upscale_images = []

    filename = request.form.get("filename", "output.pdf")

    _set_progress(
        job_id,
        phase="queued",
        status="Queued...",
        current=0,
        total=1,
        phase_started_at=time.time(),
        done=False,
        error=None,
    )

    def progress_cb(phase, current, total):
        phase_changed = False
        with JOB_PROGRESS_LOCK:
            prev = JOB_PROGRESS.get(job_id, {})
            phase_changed = prev.get("phase") != phase
        phase_started_at = time.time() if phase_changed else None
        status_map = {
            "ai_bleed": "Generating bleed...",
            "processing": "Processing pages...",
            "verifying": "Verifying output...",
            "thumbnails": "Rendering preview...",
        }
        fields = {
            "phase": phase,
            "current": int(current),
            "total": int(total),
            "status": status_map.get(phase, "Working..."),
        }
        if phase_started_at is not None:
            fields["phase_started_at"] = phase_started_at
        state = _set_progress(job_id, **fields)
        percent, eta_seconds = _compute_progress(state)
        _set_progress(job_id, percent=percent, eta_seconds=eta_seconds)

    try:
        result = process_pdf_file(
            input_path, output_path, trim_width, bleed, margins,
            bg_style=bg_style,
            insert_blank_after_cover=insert_blank,
            delete_pages=delete_pages,
            page_order=page_order,
            flip_even_pages=flip_even,
            progress_cb=progress_cb,
            output_mode=output_mode,
            bleed_mode=bleed_mode,
            ai_provider=ai_provider,
            print_rescue_quality=print_rescue_quality,
            lowres_upscale=lowres_upscale,
            lowres_threshold_dpi=lowres_threshold_dpi,
            forced_upscale_pages=forced_upscale_pages,
            forced_upscale_images=forced_upscale_images,
        )
        result["job_id"] = job_id
        result["filename"] = filename
        result["output_mode"] = output_mode
        result["bleed_mode"] = bleed_mode
        result["ai_provider"] = ai_provider
        result["print_rescue_quality"] = print_rescue_quality
        result["lowres_upscale"] = lowres_upscale
        result["lowres_threshold_dpi"] = lowres_threshold_dpi
        result["forced_upscale_pages"] = forced_upscale_pages
        result["forced_upscale_images"] = forced_upscale_images
        result["page_order"] = page_order

        # Save processing info for flipbook preview
        flipbook_info = {
            "pages": result["pages"],
            "bleed_in": bleed,
            "flip_even_pages": flip_even,
            "output_mode": output_mode,
            "bleed_mode": bleed_mode,
            "ai_provider": ai_provider,
            "print_rescue_quality": print_rescue_quality,
            "lowres_upscale": lowres_upscale,
            "lowres_threshold_dpi": lowres_threshold_dpi,
            "forced_upscale_pages": forced_upscale_pages,
            "forced_upscale_images": forced_upscale_images,
            "page_order": page_order,
        }
        info_path = os.path.join(UPLOAD_DIR, job_id, "job_info.json")
        with open(info_path, "w") as f:
            json.dump(flipbook_info, f)

        _set_progress(
            job_id,
            phase="complete",
            status="Complete",
            current=1,
            total=1,
            percent=100,
            eta_seconds=0,
            done=True,
        )

        return jsonify(result)
    except Exception as e:
        _set_progress(
            job_id,
            phase="error",
            status="Failed",
            done=True,
            error=str(e),
        )
        return jsonify({"error": str(e)}), 500


@app.route("/api/progress/<job_id>")
def api_progress(job_id):
    if not job_id or not job_id.isalnum() or len(job_id) != 8:
        return jsonify({"error": "Invalid job ID"}), 400

    with JOB_PROGRESS_LOCK:
        state = dict(JOB_PROGRESS.get(job_id, {}))
    if not state:
        return jsonify({
            "phase": "queued",
            "status": "Starting...",
            "percent": 0,
            "eta_seconds": None,
            "done": False,
        })

    percent = state.get("percent")
    eta_seconds = state.get("eta_seconds")
    if percent is None:
        percent, eta_seconds = _compute_progress(state)

    return jsonify({
        "phase": state.get("phase", "queued"),
        "status": state.get("status", "Working..."),
        "percent": int(percent),
        "eta_seconds": eta_seconds,
        "done": bool(state.get("done", False)),
        "error": state.get("error"),
    })


@app.route("/api/download/<job_id>")
def api_download(job_id):
    if not job_id.isalnum() or len(job_id) != 8:
        return jsonify({"error": "Invalid job ID"}), 400
    output_path = os.path.join(UPLOAD_DIR, job_id, "output.pdf")
    if not os.path.isfile(output_path):
        return jsonify({"error": "File not found"}), 404

    original_name = request.args.get("name", "output")
    base = os.path.splitext(original_name)[0]
    download_name = f"{base}_print_ready.pdf"
    return send_file(output_path, as_attachment=True, download_name=download_name)


@app.route("/api/flipbook/<job_id>")
def api_flipbook(job_id):
    """Return trimmed, un-flipped thumbnails for the flip-book preview."""
    if not job_id.isalnum() or len(job_id) != 8:
        return jsonify({"error": "Invalid job ID"}), 400

    output_path = os.path.join(UPLOAD_DIR, job_id, "output.pdf")
    info_path = os.path.join(UPLOAD_DIR, job_id, "job_info.json")

    if not os.path.isfile(output_path):
        return jsonify({"error": "File not found. Please process first."}), 404
    if not os.path.isfile(info_path):
        return jsonify({"error": "Processing info not found."}), 404

    try:
        with open(info_path) as f:
            info = json.load(f)
        flip_even_pages = info.get("flip_even_pages", False)
        thumbs = render_flipbook_thumbnails(
            output_path, info["pages"], info["bleed_in"],
            flip_even_pages=flip_even_pages)
        return jsonify({
            "thumbnails": thumbs,
            "page_count": len(thumbs),
            "double_sided": True,
            "flip_even_pages": flip_even_pages,
            "pages": info.get("pages", []),
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    print("PDF Print Production Web UI")
    print("Open http://localhost:5001")
    app.run(host="0.0.0.0", port=5001, debug=False, threaded=True)
