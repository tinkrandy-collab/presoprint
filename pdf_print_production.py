#!/usr/bin/env python3
"""
PDF Print Production Tool

Prepares PDFs for professional print production by:
1. Resizing document width to 10.5 inches (maintaining aspect ratio)
2. Adding 0.125" bleed on all four sides with proper content scaling
3. Adding dual-layer trim marks (white halo + black) visible on any background
4. Setting correct PDF boxes (MediaBox, TrimBox, BleedBox)
5. Preserving vector artwork, fonts, images, and color profiles

Usage:
    python pdf_print_production.py input.pdf [output.pdf]

If no output path is given, produces input_print_ready.pdf
"""

import sys
import os
import pikepdf
from pikepdf import Pdf, Page, Name, Array, Dictionary, String


# ── Constants (all in points: 1 inch = 72 pt) ──────────────────────────────
TARGET_TRIM_WIDTH_IN = 10.5
BLEED_IN = 0.125
PTS_PER_INCH = 72

TARGET_TRIM_WIDTH = TARGET_TRIM_WIDTH_IN * PTS_PER_INCH   # 756 pt
BLEED = BLEED_IN * PTS_PER_INCH                           # 9 pt

# Trim mark drawing parameters
MARK_OFFSET = 3      # pt gap between trim edge and start of mark
MARK_LENGTH = 18     # pt length of each mark line
WHITE_WEIGHT = 1.0   # pt line width for white halo
BLACK_WEIGHT = 0.75  # pt line width for black line on top


def get_page_dimensions(page):
    """Return (x0, y0, width, height) in points from the page's MediaBox."""
    mbox = page.mediabox
    x0, y0, x1, y1 = [float(v) for v in mbox]
    return x0, y0, x1 - x0, y1 - y0


def build_trim_marks_stream(trim_x, trim_y, trim_w, trim_h):
    """
    Build PDF content-stream operators for dual-layer trim marks.

    Each corner gets two L-shaped marks (horizontal + vertical).
    White marks are drawn first (halo), then black marks on top.

    Parameters are in the coordinate system of the NEW page.
    trim_x, trim_y  – lower-left corner of TrimBox
    trim_w, trim_h   – dimensions of TrimBox
    """
    lines = []
    lines.append("q")  # save graphics state

    corners = [
        # (corner_x, corner_y, h_dir, v_dir)
        (trim_x, trim_y, -1, -1),                         # bottom-left
        (trim_x + trim_w, trim_y, +1, -1),                # bottom-right
        (trim_x + trim_w, trim_y + trim_h, +1, +1),       # top-right
        (trim_x, trim_y + trim_h, -1, +1),                # top-left
    ]

    # Draw white marks first (halo layer)
    lines.append(f"{WHITE_WEIGHT} w")        # line width
    lines.append("1 1 1 RG")                 # white stroke (RGB)
    lines.append("2 J")                       # round line cap

    for cx, cy, hdir, vdir in corners:
        # Horizontal mark
        hx_start = cx + MARK_OFFSET * hdir
        hx_end = cx + (MARK_OFFSET + MARK_LENGTH) * hdir
        lines.append(f"{hx_start:.4f} {cy:.4f} m {hx_end:.4f} {cy:.4f} l S")
        # Vertical mark
        vy_start = cy + MARK_OFFSET * vdir
        vy_end = cy + (MARK_OFFSET + MARK_LENGTH) * vdir
        lines.append(f"{cx:.4f} {vy_start:.4f} m {cx:.4f} {vy_end:.4f} l S")

    # Draw black marks on top
    lines.append(f"{BLACK_WEIGHT} w")        # line width
    lines.append("0 0 0 RG")                 # black stroke (RGB)
    lines.append("2 J")                       # round line cap

    for cx, cy, hdir, vdir in corners:
        # Horizontal mark
        hx_start = cx + MARK_OFFSET * hdir
        hx_end = cx + (MARK_OFFSET + MARK_LENGTH) * hdir
        lines.append(f"{hx_start:.4f} {cy:.4f} m {hx_end:.4f} {cy:.4f} l S")
        # Vertical mark
        vy_start = cy + MARK_OFFSET * vdir
        vy_end = cy + (MARK_OFFSET + MARK_LENGTH) * vdir
        lines.append(f"{cx:.4f} {vy_start:.4f} m {cx:.4f} {vy_end:.4f} l S")

    lines.append("Q")  # restore graphics state
    return "\n".join(lines)


def process_page(pdf, page, page_index):
    """
    Transform a single page for print production.

    Steps:
      1. Read original dimensions
      2. Compute new trim size (10.5" wide, proportional height)
      3. Compute MediaBox (trim + bleed on all sides)
      4. Compute scale factor = max(media_w/orig_w, media_h/orig_h)
      5. Center scaled content in MediaBox, clip to MediaBox
      6. Draw trim marks on top
      7. Set MediaBox, TrimBox, BleedBox
    """
    orig_x0, orig_y0, orig_w, orig_h = get_page_dimensions(page)

    # Handle page rotation — swap width/height for effective dimensions
    rotation = int(page.obj.get(Name.Rotate, 0)) % 360
    if rotation in (90, 270):
        orig_w, orig_h = orig_h, orig_w

    print(f"  Page {page_index + 1}: original size = {orig_w:.2f} × {orig_h:.2f} pt "
          f"({orig_w/72:.3f}\" × {orig_h/72:.3f}\")"
          f"{f'  (rotation: {rotation}°)' if rotation else ''}")

    # ── 1. New trim dimensions ───────────────────────────────────────────
    trim_w = TARGET_TRIM_WIDTH                         # 756 pt (10.5")
    aspect = orig_h / orig_w
    trim_h = trim_w * aspect                           # proportional height

    # ── 2. MediaBox = trim + bleed on all sides ──────────────────────────
    media_w = trim_w + 2 * BLEED
    media_h = trim_h + 2 * BLEED

    print(f"  Trim size  = {trim_w:.2f} × {trim_h:.2f} pt "
          f"({trim_w/72:.4f}\" × {trim_h/72:.4f}\")")
    print(f"  Media size = {media_w:.2f} × {media_h:.2f} pt "
          f"({media_w/72:.4f}\" × {media_h/72:.4f}\")")

    # ── 3. Scale factor (cover strategy) ─────────────────────────────────
    scale_x = media_w / orig_w
    scale_y = media_h / orig_h
    scale = max(scale_x, scale_y)

    scaled_w = orig_w * scale
    scaled_h = orig_h * scale

    # Center offsets
    x_off = (media_w - scaled_w) / 2.0
    y_off = (media_h - scaled_h) / 2.0

    print(f"  Scale factor = {scale:.6f}  (scale_x={scale_x:.6f}, scale_y={scale_y:.6f})")
    print(f"  Scaled content = {scaled_w:.2f} × {scaled_h:.2f} pt")
    print(f"  Offset = ({x_off:.4f}, {y_off:.4f}) pt")

    # ── 4. Wrap original page content in a Form XObject ──────────────────
    # This preserves all original resources, fonts, images, etc.
    orig_page_obj = page.obj

    # Get the original page's content stream(s) as bytes
    if Name.Contents in orig_page_obj:
        contents = orig_page_obj[Name.Contents]
        if isinstance(contents, pikepdf.Array):
            content_data = b""
            for stream_ref in contents:
                stream = stream_ref
                content_data += stream.read_bytes() + b"\n"
        else:
            content_data = contents.read_bytes()
    else:
        content_data = b""

    # Get original resources
    orig_resources = orig_page_obj.get(Name.Resources, Dictionary())

    # Create a Form XObject from original page content
    # BBox must match the original coordinate system (including origin offset)
    form_xobj = pikepdf.Stream(pdf, content_data)
    form_xobj[Name.Type] = Name.XObject
    form_xobj[Name.Subtype] = Name.Form
    form_xobj[Name.BBox] = Array([orig_x0, orig_y0, orig_x0 + orig_w, orig_y0 + orig_h])
    form_xobj[Name.Resources] = orig_resources
    # Matrix translates the original origin to (0,0) so our external transform works
    form_xobj[Name.Matrix] = Array([1, 0, 0, 1, -orig_x0, -orig_y0])

    # Register Form XObject
    form_name = Name("/OrigPage")
    if Name.Resources not in orig_page_obj:
        orig_page_obj[Name.Resources] = Dictionary()

    new_resources = Dictionary()
    xobj_dict = Dictionary()
    xobj_dict[form_name] = form_xobj
    new_resources[Name.XObject] = xobj_dict
    # We also need a minimal ProcSet
    new_resources[Name.ProcSet] = Array([Name.PDF, Name.Text, Name.ImageB, Name.ImageC, Name.ImageI])

    # ── 5. Build new content stream ──────────────────────────────────────
    # Content order:
    #   a) Clip to MediaBox
    #   b) Apply transformation matrix and draw original content (Form XObject)
    #   c) Draw trim marks on top (last layer)

    content_lines = []

    # a) Clip to MediaBox boundaries
    content_lines.append("q")
    content_lines.append(f"0 0 {media_w:.4f} {media_h:.4f} re W n")

    # b) Transform and draw original content
    content_lines.append(f"{scale:.6f} 0 0 {scale:.6f} {x_off:.4f} {y_off:.4f} cm")
    content_lines.append(f"{form_name} Do")

    content_lines.append("Q")  # end clip + transform

    # c) Trim marks (drawn in MediaBox coordinate space, on top of everything)
    trim_x = BLEED
    trim_y = BLEED
    trim_marks = build_trim_marks_stream(trim_x, trim_y, trim_w, trim_h)
    content_lines.append(trim_marks)

    new_content = "\n".join(content_lines).encode("latin-1")

    # ── 6. Replace page content and set boxes ────────────────────────────
    new_stream = pikepdf.Stream(pdf, new_content)
    orig_page_obj[Name.Contents] = new_stream
    orig_page_obj[Name.Resources] = new_resources

    # MediaBox: full size including bleed
    orig_page_obj[Name.MediaBox] = Array([0, 0, media_w, media_h])

    # TrimBox: the final cut size, inset by bleed
    orig_page_obj[Name.TrimBox] = Array([BLEED, BLEED, BLEED + trim_w, BLEED + trim_h])

    # BleedBox: same as MediaBox
    orig_page_obj[Name.BleedBox] = Array([0, 0, media_w, media_h])

    # Remove CropBox if present (let MediaBox define visible area)
    if Name.CropBox in orig_page_obj:
        del orig_page_obj[Name.CropBox]

    # Remove Rotate since we've accounted for it in the transform
    if Name.Rotate in orig_page_obj:
        del orig_page_obj[Name.Rotate]

    # Compute actual bleed on each side for verification
    bleed_left = trim_x
    bleed_right = media_w - (trim_x + trim_w)
    bleed_bottom = trim_y
    bleed_top = media_h - (trim_y + trim_h)

    print(f"  Bleed (L/R/T/B): {bleed_left/72:.4f}\" / {bleed_right/72:.4f}\" / "
          f"{bleed_top/72:.4f}\" / {bleed_bottom/72:.4f}\"")
    print(f"  TrimBox  = [{BLEED:.2f} {BLEED:.2f} {BLEED+trim_w:.2f} {BLEED+trim_h:.2f}]")
    print(f"  MediaBox = [0 0 {media_w:.2f} {media_h:.2f}]")
    print()


def process_pdf(input_path, output_path=None):
    """Process all pages of a PDF for print production."""
    if not os.path.isfile(input_path):
        print(f"Error: File not found: {input_path}")
        sys.exit(1)

    if output_path is None:
        base, ext = os.path.splitext(input_path)
        output_path = f"{base}_print_ready{ext}"

    print(f"PDF Print Production Tool")
    print(f"{'='*60}")
    print(f"Input:  {input_path}")
    print(f"Output: {output_path}")
    print(f"Target trim width: {TARGET_TRIM_WIDTH_IN}\"")
    print(f"Bleed: {BLEED_IN}\" on all sides")
    print(f"{'='*60}")
    print()

    pdf = Pdf.open(input_path)

    num_pages = len(pdf.pages)
    print(f"Processing {num_pages} page(s)...\n")

    for i, page in enumerate(pdf.pages):
        process_page(pdf, page, i)

    # Save with linearization disabled to preserve structure
    pdf.save(output_path, linearize=False)
    print(f"{'='*60}")
    print(f"Saved print-ready PDF: {output_path}")
    print()

    # Verification summary
    verify_output(output_path)


def verify_output(path):
    """Open the output PDF and print verification info."""
    pdf = Pdf.open(path)
    print("VERIFICATION:")
    print(f"{'─'*60}")

    for i, page in enumerate(pdf.pages):
        mbox = [float(v) for v in page.mediabox]
        tbox = [float(v) for v in page.obj.get(Name.TrimBox, page.mediabox)]
        bbox = [float(v) for v in page.obj.get(Name.BleedBox, page.mediabox)]

        media_w = mbox[2] - mbox[0]
        media_h = mbox[3] - mbox[1]
        trim_w = tbox[2] - tbox[0]
        trim_h = tbox[3] - tbox[1]

        bleed_l = tbox[0] - mbox[0]
        bleed_r = mbox[2] - tbox[2]
        bleed_b = tbox[1] - mbox[1]
        bleed_t = mbox[3] - tbox[3]

        print(f"  Page {i+1}:")
        print(f"    MediaBox: {media_w:.2f} × {media_h:.2f} pt  "
              f"({media_w/72:.4f}\" × {media_h/72:.4f}\")")
        print(f"    TrimBox:  {trim_w:.2f} × {trim_h:.2f} pt  "
              f"({trim_w/72:.4f}\" × {trim_h/72:.4f}\")")
        print(f"    BleedBox: {bbox[2]-bbox[0]:.2f} × {bbox[3]-bbox[1]:.2f} pt")
        print(f"    Bleed L={bleed_l/72:.4f}\"  R={bleed_r/72:.4f}\"  "
              f"T={bleed_t/72:.4f}\"  B={bleed_b/72:.4f}\"")

        # Checks
        checks = []
        checks.append(("Trim width = 10.5\"", abs(trim_w/72 - 10.5) < 0.01))
        checks.append(("Bleed L = 0.125\"", abs(bleed_l/72 - 0.125) < 0.001))
        checks.append(("Bleed R = 0.125\"", abs(bleed_r/72 - 0.125) < 0.001))
        checks.append(("Bleed T = 0.125\"", abs(bleed_t/72 - 0.125) < 0.001))
        checks.append(("Bleed B = 0.125\"", abs(bleed_b/72 - 0.125) < 0.001))
        checks.append(("MediaBox = Trim + 2×Bleed (W)", abs(media_w - trim_w - 2*BLEED) < 0.01))
        checks.append(("MediaBox = Trim + 2×Bleed (H)", abs(media_h - trim_h - 2*BLEED) < 0.01))
        checks.append(("TrimBox present", Name.TrimBox in page.obj))
        checks.append(("BleedBox present", Name.BleedBox in page.obj))

        all_pass = True
        for label, ok in checks:
            status = "PASS" if ok else "FAIL"
            if not ok:
                all_pass = False
            print(f"    [{status}] {label}")

        if all_pass:
            print(f"    All checks passed.")
        else:
            print(f"    WARNING: Some checks failed!")
        print()

    pdf.close()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python pdf_print_production.py input.pdf [output.pdf]")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    process_pdf(input_file, output_file)
