#!/usr/bin/env python3
"""
Web UI for the PDF Print Production Tool.

Run:
    python app.py

Then open http://localhost:5001 in your browser.
"""

import os
import re
import uuid
import json
import tempfile
import shutil
from flask import Flask, request, jsonify, send_file, send_from_directory

import pikepdf
from pikepdf import Pdf, Page, Name, Array, Dictionary, String

# ── Constants (all in points: 1 inch = 72 pt) ────────────────────────────────
PTS_PER_INCH = 72
MM_PER_INCH = 25.4
DEFAULT_BLEED_IN = 0.125

# Safe margins (inches) — space between trim edge and slide content
DEFAULT_MARGIN_LEFT = 0.5
DEFAULT_MARGIN_RIGHT = 0.5
DEFAULT_MARGIN_TOP = 1.5      # extra room for binding
DEFAULT_MARGIN_BOTTOM = 0.5

# Paper size presets (landscape orientation: width > height, in inches)
PAPER_SIZES = {
    "letter": {
        "label": 'US Letter (11" x 8.5")',
        "width_in": 11.0,
        "height_in": 8.5,
        "description": "Standard US paper size",
    },
    "a4": {
        "label": "A4 (297mm x 210mm)",
        "width_in": 297 / MM_PER_INCH,
        "height_in": 210 / MM_PER_INCH,
        "description": "International standard (ISO 216)",
    },
    "a5": {
        "label": "A5 (210mm x 148mm)",
        "width_in": 210 / MM_PER_INCH,
        "height_in": 148 / MM_PER_INCH,
        "description": "Half of A4, common in Europe & Asia",
    },
    "b5_jis": {
        "label": "B5 JIS (257mm x 182mm)",
        "width_in": 257 / MM_PER_INCH,
        "height_in": 182 / MM_PER_INCH,
        "description": "Common in Japan for books & magazines",
    },
    "tabloid": {
        "label": 'Tabloid / Ledger (17" x 11")',
        "width_in": 17.0,
        "height_in": 11.0,
        "description": "Large US format, 2x Letter",
    },
}
DEFAULT_PAPER_SIZE = "letter"

MARK_OFFSET = 3
MARK_LENGTH = 18
WHITE_WEIGHT = 1.0
BLACK_WEIGHT = 0.75

app = Flask(__name__, static_folder="static")

UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "pdf_print_uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)


# ── PDF Processing ────────────────────────────────────────────────────────────

def get_page_dimensions(page):
    mbox = page.mediabox
    x0, y0, x1, y1 = [float(v) for v in mbox]
    return x0, y0, x1 - x0, y1 - y0


def detect_background_color(page):
    """Best-effort detection of a page's background fill color.

    Scans the beginning of the content stream for an RGB fill (rg) or
    grayscale fill (g) that appears before/near a full-page rectangle.
    Returns (r, g, b) floats in 0-1 range; defaults to white.
    """
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
        # Only look at the first ~3000 chars (background is always first)
        early = text[:3000]

        # RGB fill:  "R G B rg"
        m = re.search(r"([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+rg", early)
        if m:
            return (float(m.group(1)), float(m.group(2)), float(m.group(3)))

        # RGB via scn/sc (color-space fill):  "R G B scn" or "R G B sc"
        m = re.search(r"([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+scn?", early)
        if m:
            return (float(m.group(1)), float(m.group(2)), float(m.group(3)))

        # Grayscale fill:  "G g"
        m = re.search(r"([\d.]+)\s+g\b", early)
        if m:
            v = float(m.group(1))
            return (v, v, v)

    except Exception:
        pass

    return (1.0, 1.0, 1.0)  # white fallback


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
        hx_start = cx + MARK_OFFSET * hdir
        hx_end = cx + (MARK_OFFSET + MARK_LENGTH) * hdir
        lines.append(f"{hx_start:.4f} {cy:.4f} m {hx_end:.4f} {cy:.4f} l S")
        vy_start = cy + MARK_OFFSET * vdir
        vy_end = cy + (MARK_OFFSET + MARK_LENGTH) * vdir
        lines.append(f"{cx:.4f} {vy_start:.4f} m {cx:.4f} {vy_end:.4f} l S")
    # Black on top
    lines.append(f"{BLACK_WEIGHT} w")
    lines.append("0 0 0 RG")
    lines.append("2 J")
    for cx, cy, hdir, vdir in corners:
        hx_start = cx + MARK_OFFSET * hdir
        hx_end = cx + (MARK_OFFSET + MARK_LENGTH) * hdir
        lines.append(f"{hx_start:.4f} {cy:.4f} m {hx_end:.4f} {cy:.4f} l S")
        vy_start = cy + MARK_OFFSET * vdir
        vy_end = cy + (MARK_OFFSET + MARK_LENGTH) * vdir
        lines.append(f"{cx:.4f} {vy_start:.4f} m {cx:.4f} {vy_end:.4f} l S")
    lines.append("Q")
    return "\n".join(lines)


def build_blue_gradient_resources(pdf, media_w, media_h):
    """Create a PDF axial-shading resource for a vertical blue gradient.

    Returns (shading_dict, shading_name) to add to /Resources/Shading.
    Gradient runs bottom→top: dark navy → medium blue.
    """
    # Colour stops: bottom = dark navy, top = lighter blue
    c0 = (0.04, 0.06, 0.14)   # #0a0f24  (deep navy)
    c1 = (0.11, 0.20, 0.38)   # #1c3361  (medium blue)

    # Type 2 exponential interpolation function (N=1 → linear)
    fn = Dictionary()
    fn[Name("/FunctionType")] = 2
    fn[Name("/Domain")] = Array([0, 1])
    fn[Name("/C0")] = Array([c0[0], c0[1], c0[2]])
    fn[Name("/C1")] = Array([c1[0], c1[1], c1[2]])
    fn[Name("/N")] = 1

    shading = Dictionary()
    shading[Name("/ShadingType")] = 2          # axial
    shading[Name("/ColorSpace")] = Name.DeviceRGB
    shading[Name("/Coords")] = Array([0, 0, 0, media_h])  # bottom→top
    shading[Name("/Function")] = pdf.make_indirect(fn)
    shading[Name("/Extend")] = Array([True, True])

    shading_name = Name("/BlueGrad")
    return pdf.make_indirect(shading), shading_name


def process_page(pdf, page, page_index, trim_width_in, trim_height_in,
                 bleed_in, margins, bg_color=None, bg_style="auto"):
    """Re-compose a page with safe margins and background-only bleed.

    Parameters
    ----------
    trim_width_in, trim_height_in : float
        Target trim box size in inches.
    bleed_in : float
        Bleed amount in inches (all four sides).
    margins : dict
        Keys: left, right, top, bottom — safe margins in inches.
    bg_color : tuple or None
        (r, g, b) 0-1 for the page background.  Auto-detected if None.
    bg_style : str
        "auto" (default) – use detected/supplied solid colour.
        "blue_gradient" – vertical blue gradient background.
    """
    bleed = bleed_in * PTS_PER_INCH
    trim_w = trim_width_in * PTS_PER_INCH
    trim_h = trim_height_in * PTS_PER_INCH

    media_w = trim_w + 2 * bleed
    media_h = trim_h + 2 * bleed

    # ── Original page geometry ───────────────────────────────────────────
    orig_x0, orig_y0, orig_w, orig_h = get_page_dimensions(page)
    rotation = int(page.obj.get(Name.Rotate, 0)) % 360
    if rotation in (90, 270):
        orig_w, orig_h = orig_h, orig_w

    # ── Safe area inside the trim box ────────────────────────────────────
    ml = margins["left"] * PTS_PER_INCH
    mr = margins["right"] * PTS_PER_INCH
    mt = margins["top"] * PTS_PER_INCH
    mb = margins["bottom"] * PTS_PER_INCH

    safe_w = trim_w - ml - mr
    safe_h = trim_h - mt - mb
    if safe_w <= 0 or safe_h <= 0:
        raise ValueError(
            f"Margins ({margins}) are too large for the trim area "
            f"({trim_width_in:.3f}\" x {trim_height_in:.3f}\")."
        )

    # Uniform scale so full slide fits in safe area (no cropping)
    scale_x = safe_w / orig_w
    scale_y = safe_h / orig_h
    scale = min(scale_x, scale_y)

    scaled_w = orig_w * scale
    scaled_h = orig_h * scale

    # Centre the scaled slide within the safe area
    # Safe area origin in media-box coordinates:
    safe_x0 = bleed + ml
    safe_y0 = bleed + mb

    x_off = safe_x0 + (safe_w - scaled_w) / 2.0
    y_off = safe_y0 + (safe_h - scaled_h) / 2.0

    # ── Background colour ────────────────────────────────────────────────
    if bg_color is None:
        bg_color = detect_background_color(page)

    # ── Wrap original page as Form XObject ───────────────────────────────
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

    # ── Build the new content stream ─────────────────────────────────────
    cl = []

    # 1. Fill entire media box with background (extends into bleed)
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
        r, g, b = (0.04, 0.06, 0.14)   # report the dark end
    else:
        r, g, b = bg_color
        cl.append("q")
        cl.append(f"{r:.6f} {g:.6f} {b:.6f} rg")
        cl.append(f"0 0 {media_w:.4f} {media_h:.4f} re f")
        cl.append("Q")

    # 2. Place original slide scaled into the safe area
    cl.append("q")
    cl.append(f"{scale:.6f} 0 0 {scale:.6f} {x_off:.4f} {y_off:.4f} cm")
    cl.append(f"{form_name} Do")
    cl.append("Q")

    # 3. Trim marks
    cl.append(build_trim_marks_stream(bleed, bleed, trim_w, trim_h))

    new_content = "\n".join(cl).encode("latin-1")
    new_stream = pikepdf.Stream(pdf, new_content)

    # ── Set page boxes ───────────────────────────────────────────────────
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
        "margins": {
            "left": margins["left"],
            "right": margins["right"],
            "top": margins["top"],
            "bottom": margins["bottom"],
        },
        "bleed": {
            "left": round(bleed / 72, 4),
            "right": round(bleed / 72, 4),
            "top": round(bleed / 72, 4),
            "bottom": round(bleed / 72, 4),
        },
        "background_color": "blue gradient" if bg_style == "blue_gradient"
                           else f"rgb({r:.2f}, {g:.2f}, {b:.2f})",
    }


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


def process_pdf_file(input_path, output_path, trim_width_in, trim_height_in,
                     bleed_in, margins, bg_style="auto"):
    pdf = Pdf.open(input_path)
    pages_info = []
    for i, page in enumerate(pdf.pages):
        info = process_page(pdf, page, i,
                            trim_width_in, trim_height_in,
                            bleed_in, margins,
                            bg_style=bg_style)
        pages_info.append(info)
    pdf.save(output_path, linearize=False)
    pdf.close()

    # Verify
    pdf2 = Pdf.open(output_path)
    verification = []
    for i, page in enumerate(pdf2.pages):
        checks = verify_page(page, trim_width_in, trim_height_in, bleed_in)
        all_pass = all(c["pass"] for c in checks)
        verification.append({"page": i + 1, "checks": checks, "all_pass": all_pass})
    pdf2.close()

    return {"pages": pages_info, "verification": verification}


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return send_from_directory("static", "index.html")


@app.route("/api/paper-sizes")
def api_paper_sizes():
    bleed = float(request.args.get("bleed", DEFAULT_BLEED_IN))
    result = {}
    for key, ps in PAPER_SIZES.items():
        trim_w = ps["width_in"] - 2 * bleed
        trim_h = ps["height_in"] - 2 * bleed
        result[key] = {
            "label": ps["label"],
            "description": ps["description"],
            "paper_width_in": ps["width_in"],
            "paper_height_in": ps["height_in"],
            "trim_width_in": trim_w,
            "trim_height_in": trim_h,
        }
    return jsonify(result)


@app.route("/api/process", methods=["POST"])
def api_process():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Please upload a PDF file"}), 400

    bleed = float(request.form.get("bleed", DEFAULT_BLEED_IN))

    # Paper size → trim dimensions
    paper_size = request.form.get("paper_size", DEFAULT_PAPER_SIZE)
    if paper_size in PAPER_SIZES:
        ps = PAPER_SIZES[paper_size]
        trim_width = ps["width_in"] - 2 * bleed
        trim_height = ps["height_in"] - 2 * bleed
    else:
        # Fallback: manual trim dimensions
        trim_width = float(request.form.get("trim_width", 10.75))
        trim_height = float(request.form.get("trim_height", 8.25))

    # Safe margins
    margins = {
        "left":   float(request.form.get("margin_left",   DEFAULT_MARGIN_LEFT)),
        "right":  float(request.form.get("margin_right",  DEFAULT_MARGIN_RIGHT)),
        "top":    float(request.form.get("margin_top",     DEFAULT_MARGIN_TOP)),
        "bottom": float(request.form.get("margin_bottom",  DEFAULT_MARGIN_BOTTOM)),
    }

    bg_style = request.form.get("bg_style", "auto")
    if bg_style not in ("auto", "blue_gradient"):
        bg_style = "auto"

    job_id = str(uuid.uuid4())[:8]
    job_dir = os.path.join(UPLOAD_DIR, job_id)
    os.makedirs(job_dir, exist_ok=True)

    input_path = os.path.join(job_dir, "input.pdf")
    output_path = os.path.join(job_dir, "output.pdf")
    file.save(input_path)

    try:
        result = process_pdf_file(input_path, output_path,
                                  trim_width, trim_height,
                                  bleed, margins,
                                  bg_style=bg_style)
        result["job_id"] = job_id
        result["filename"] = file.filename
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/download/<job_id>")
def api_download(job_id):
    # Sanitize job_id to prevent path traversal
    if not job_id.isalnum() or len(job_id) != 8:
        return jsonify({"error": "Invalid job ID"}), 400
    output_path = os.path.join(UPLOAD_DIR, job_id, "output.pdf")
    if not os.path.isfile(output_path):
        return jsonify({"error": "File not found"}), 404

    original_name = request.args.get("name", "output")
    base = os.path.splitext(original_name)[0]
    download_name = f"{base}_print_ready.pdf"

    return send_file(output_path, as_attachment=True, download_name=download_name)


if __name__ == "__main__":
    print("PDF Print Production Web UI")
    print("Open http://localhost:5001")
    app.run(host="0.0.0.0", port=5001, debug=False)
