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
from flask import Flask, request, jsonify, send_file, send_from_directory

import pikepdf
from pikepdf import Pdf, Page, Name, Array, Dictionary, String
import fitz  # pymupdf – for thumbnail rendering

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

app = Flask(__name__, static_folder="static")

UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "pdf_print_uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)


# ── Helpers ───────────────────────────────────────────────────────────────────

def get_page_dimensions(page):
    mbox = page.mediabox
    x0, y0, x1, y1 = [float(v) for v in mbox]
    return x0, y0, x1 - x0, y1 - y0


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
                 flip=False):
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
    if bg_style == "blue_gradient":
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
    else:
        r, g, b = bg_color
        cl.append("q")
        cl.append(f"{r:.6f} {g:.6f} {b:.6f} rg")
        cl.append(f"0 0 {media_w:.4f} {media_h:.4f} re f")
        cl.append("Q")

    # 2. Place original slide
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
                           else f"rgb({r:.2f}, {g:.2f}, {b:.2f})",
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


def render_thumbnails(pdf_path, dpi=THUMBNAIL_DPI):
    """Return list of base64-encoded PNG thumbnails, one per page."""
    doc = fitz.open(pdf_path)
    thumbs = []
    for page in doc:
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        png_data = pix.tobytes("png")
        b64 = base64.b64encode(png_data).decode("ascii")
        thumbs.append(b64)
    doc.close()
    return thumbs


def render_flipbook_thumbnails(processed_path, pages_info, bleed_in,
                               flip_even_pages=True, dpi=150):
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


def verify_page(page, trim_width_in, trim_height_in, bleed_in):
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
                     flip_even_pages=True):
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

    # ── Delete pages first (in reverse so indices stay valid) ─────────
    if delete_pages:
        for idx in sorted(set(delete_pages), reverse=True):
            if 0 <= idx < len(pdf.pages):
                del pdf.pages[idx]

    if len(pdf.pages) == 0:
        raise ValueError("No pages left after deletion.")

    # ── Determine trim height from the first page's aspect ratio ─────
    first_page = pdf.pages[0]
    _, _, orig_w, orig_h = get_page_dimensions(first_page)
    rotation = int(first_page.obj.get(Name.Rotate, 0)) % 360
    if rotation in (90, 270):
        orig_w, orig_h = orig_h, orig_w
    trim_height_in = compute_trim_height(orig_w, orig_h, trim_width_in, margins)

    # ── Insert blank page after cover (index 0) ─────────────────────
    if insert_blank_after_cover and len(pdf.pages) >= 1:
        add_blank_page(pdf, trim_width_in, trim_height_in, bleed_in, insert_at=1)

    # ── Process each page ────────────────────────────────────────────
    pages_info = []
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
            })
            continue

        flip = flip_even_pages and (i % 2 == 1)
        info = process_page(pdf, page, i,
                            trim_width_in, trim_height_in,
                            bleed_in, margins,
                            bg_style=bg_style,
                            flip=flip)
        pages_info.append(info)

    pdf.save(output_path, linearize=False)
    pdf.close()

    # ── Verify ───────────────────────────────────────────────────────
    pdf2 = Pdf.open(output_path)
    verification = []
    for i, page in enumerate(pdf2.pages):
        checks = verify_page(page, trim_width_in, trim_height_in, bleed_in)
        all_pass = all(c["pass"] for c in checks)
        verification.append({"page": i + 1, "checks": checks, "all_pass": all_pass})
    pdf2.close()

    # ── Thumbnails ───────────────────────────────────────────────────
    thumbnails = render_thumbnails(output_path)

    return {
        "pages": pages_info,
        "verification": verification,
        "thumbnails": thumbnails,
        "trim_width_in": trim_width_in,
        "trim_height_in": round(trim_height_in, 4),
    }


def _flip_existing_page(pdf, page, trim_width_in, trim_height_in, bleed_in):
    """Rotate an already-composed page 180° by wrapping its contents."""
    bleed = bleed_in * PTS_PER_INCH
    trim_w = trim_width_in * PTS_PER_INCH
    trim_h = trim_height_in * PTS_PER_INCH
    media_w = trim_w + 2 * bleed
    media_h = trim_h + 2 * bleed

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
    if bg_style not in ("auto", "blue_gradient"):
        bg_style = "auto"

    insert_blank = request.form.get("insert_blank_after_cover", "false") == "true"
    flip_even = request.form.get("flip_even_pages", "true") == "true"

    # Parse delete_pages JSON array
    delete_pages_raw = request.form.get("delete_pages", "[]")
    try:
        delete_pages = json.loads(delete_pages_raw)
        if not isinstance(delete_pages, list):
            delete_pages = []
        delete_pages = [int(x) for x in delete_pages]
    except (json.JSONDecodeError, ValueError):
        delete_pages = []

    filename = request.form.get("filename", "output.pdf")

    try:
        result = process_pdf_file(
            input_path, output_path, trim_width, bleed, margins,
            bg_style=bg_style,
            insert_blank_after_cover=insert_blank,
            delete_pages=delete_pages,
            flip_even_pages=flip_even,
        )
        result["job_id"] = job_id
        result["filename"] = filename

        # Save processing info for flipbook preview
        flipbook_info = {
            "pages": result["pages"],
            "bleed_in": bleed,
            "flip_even_pages": flip_even,
        }
        info_path = os.path.join(UPLOAD_DIR, job_id, "job_info.json")
        with open(info_path, "w") as f:
            json.dump(flipbook_info, f)

        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


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
        flip_even_pages = info.get("flip_even_pages", True)
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
    app.run(host="0.0.0.0", port=5001, debug=False)
