#!/usr/bin/env python3
"""
Web UI for the PDF Print Production Tool.

Run:
    python app.py

Then open http://localhost:5001 in your browser.
"""

import os
import uuid
import json
import tempfile
import shutil
from flask import Flask, request, jsonify, send_file, send_from_directory

import pikepdf
from pikepdf import Pdf, Page, Name, Array, Dictionary, String

# ── Constants (all in points: 1 inch = 72 pt) ────────────────────────────────
PTS_PER_INCH = 72
DEFAULT_TRIM_WIDTH_IN = 10.5
DEFAULT_BLEED_IN = 0.125

MARK_OFFSET = 3
MARK_LENGTH = 18
WHITE_WEIGHT = 1.0
BLACK_WEIGHT = 0.75

app = Flask(__name__, static_folder="static")

UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "pdf_print_uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)


# ── PDF Processing (adapted from pdf_print_production.py) ─────────────────────

def get_page_dimensions(page):
    mbox = page.mediabox
    x0, y0, x1, y1 = [float(v) for v in mbox]
    return x0, y0, x1 - x0, y1 - y0


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


def process_page(pdf, page, page_index, trim_width_in, bleed_in):
    bleed = bleed_in * PTS_PER_INCH
    trim_w = trim_width_in * PTS_PER_INCH

    orig_x0, orig_y0, orig_w, orig_h = get_page_dimensions(page)
    rotation = int(page.obj.get(Name.Rotate, 0)) % 360
    if rotation in (90, 270):
        orig_w, orig_h = orig_h, orig_w

    aspect = orig_h / orig_w
    trim_h = trim_w * aspect

    media_w = trim_w + 2 * bleed
    media_h = trim_h + 2 * bleed

    scale_x = media_w / orig_w
    scale_y = media_h / orig_h
    scale = max(scale_x, scale_y)
    scaled_w = orig_w * scale
    scaled_h = orig_h * scale
    x_off = (media_w - scaled_w) / 2.0
    y_off = (media_h - scaled_h) / 2.0

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
    new_resources[Name.ProcSet] = Array([Name.PDF, Name.Text, Name.ImageB, Name.ImageC, Name.ImageI])

    content_lines = []
    content_lines.append("q")
    content_lines.append(f"0 0 {media_w:.4f} {media_h:.4f} re W n")
    content_lines.append(f"{scale:.6f} 0 0 {scale:.6f} {x_off:.4f} {y_off:.4f} cm")
    content_lines.append(f"{form_name} Do")
    content_lines.append("Q")

    trim_marks = build_trim_marks_stream(bleed, bleed, trim_w, trim_h)
    content_lines.append(trim_marks)

    new_content = "\n".join(content_lines).encode("latin-1")
    new_stream = pikepdf.Stream(pdf, new_content)
    orig_page_obj[Name.Contents] = new_stream
    orig_page_obj[Name.Resources] = new_resources
    orig_page_obj[Name.MediaBox] = Array([0, 0, media_w, media_h])
    orig_page_obj[Name.TrimBox] = Array([bleed, bleed, bleed + trim_w, bleed + trim_h])
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
        "scale_factor": round(scale, 6),
        "bleed": {
            "left": round(bleed / 72, 4),
            "right": round((media_w - bleed - trim_w) / 72, 4),
            "top": round((media_h - bleed - trim_h) / 72, 4),
            "bottom": round(bleed / 72, 4),
        },
    }


def verify_page(page, bleed_in):
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
        ("Trim width = 10.5\"", abs(trim_w / 72 - 10.5) < 0.01),
        ("Bleed Left = 0.125\"", abs(bleed_l / 72 - bleed_in) < 0.001),
        ("Bleed Right = 0.125\"", abs(bleed_r / 72 - bleed_in) < 0.001),
        ("Bleed Top = 0.125\"", abs(bleed_t / 72 - bleed_in) < 0.001),
        ("Bleed Bottom = 0.125\"", abs(bleed_b / 72 - bleed_in) < 0.001),
        ("MediaBox width correct", abs(media_w - trim_w - 2 * bleed) < 0.01),
        ("MediaBox height correct", abs(media_h - trim_h - 2 * bleed) < 0.01),
        ("TrimBox present", Name.TrimBox in page.obj),
        ("BleedBox present", Name.BleedBox in page.obj),
    ]

    return [{"label": label, "pass": ok} for label, ok in checks]


def process_pdf_file(input_path, output_path, trim_width_in, bleed_in):
    pdf = Pdf.open(input_path)
    pages_info = []
    for i, page in enumerate(pdf.pages):
        info = process_page(pdf, page, i, trim_width_in, bleed_in)
        pages_info.append(info)
    pdf.save(output_path, linearize=False)
    pdf.close()

    # Verify
    pdf2 = Pdf.open(output_path)
    verification = []
    for i, page in enumerate(pdf2.pages):
        checks = verify_page(page, bleed_in)
        all_pass = all(c["pass"] for c in checks)
        verification.append({"page": i + 1, "checks": checks, "all_pass": all_pass})
    pdf2.close()

    return {"pages": pages_info, "verification": verification}


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return send_from_directory("static", "index.html")


@app.route("/api/process", methods=["POST"])
def api_process():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Please upload a PDF file"}), 400

    trim_width = float(request.form.get("trim_width", DEFAULT_TRIM_WIDTH_IN))
    bleed = float(request.form.get("bleed", DEFAULT_BLEED_IN))

    job_id = str(uuid.uuid4())[:8]
    job_dir = os.path.join(UPLOAD_DIR, job_id)
    os.makedirs(job_dir, exist_ok=True)

    input_path = os.path.join(job_dir, "input.pdf")
    output_path = os.path.join(job_dir, "output.pdf")
    file.save(input_path)

    try:
        result = process_pdf_file(input_path, output_path, trim_width, bleed)
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
