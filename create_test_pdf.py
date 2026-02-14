#!/usr/bin/env python3
"""Create a test PDF with mixed backgrounds for validating the print production tool."""

from reportlab.lib.pagesizes import landscape
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor


def create_test_pdf(output_path="test_input.pdf"):
    """Create a 10x5.625 inch landscape PDF with varied backgrounds."""
    width = 10 * inch    # 720 pt
    height = 5.625 * inch  # 405 pt  (16:9 aspect)

    c = canvas.Canvas(output_path, pagesize=(width, height))

    # ── Page 1: Dark gradient-like background ────────────────────────────
    # Dark navy background
    c.setFillColor(HexColor("#1a2332"))
    c.rect(0, 0, width, height, fill=1, stroke=0)

    # A lighter band in the middle
    c.setFillColor(HexColor("#2d4a6f"))
    c.rect(0, height * 0.3, width, height * 0.4, fill=1, stroke=0)

    # Title text
    c.setFillColor(HexColor("#ffffff"))
    c.setFont("Helvetica-Bold", 36)
    c.drawCentredString(width / 2, height / 2 + 20, "Print Production Test")

    c.setFont("Helvetica", 18)
    c.drawCentredString(width / 2, height / 2 - 20, "Page 1 — Dark Background")

    # Corner markers (small circles to verify content extends to edges)
    c.setFillColor(HexColor("#ff3366"))
    for x, y in [(15, 15), (width - 15, 15), (15, height - 15), (width - 15, height - 15)]:
        c.circle(x, y, 8, fill=1, stroke=0)

    # Edge markers (to verify bleed extension)
    c.setFillColor(HexColor("#00ff88"))
    c.rect(0, height/2 - 5, 10, 10, fill=1, stroke=0)       # left edge
    c.rect(width - 10, height/2 - 5, 10, 10, fill=1, stroke=0) # right edge
    c.rect(width/2 - 5, 0, 10, 10, fill=1, stroke=0)         # bottom edge
    c.rect(width/2 - 5, height - 10, 10, 10, fill=1, stroke=0)  # top edge

    c.showPage()

    # ── Page 2: Light background ─────────────────────────────────────────
    c.setFillColor(HexColor("#f5f0e8"))
    c.rect(0, 0, width, height, fill=1, stroke=0)

    # Some vector artwork
    c.setFillColor(HexColor("#cc4444"))
    c.rect(50, 50, 200, 150, fill=1, stroke=0)

    c.setFillColor(HexColor("#4444cc"))
    c.circle(width - 150, height - 100, 80, fill=1, stroke=0)

    c.setFillColor(HexColor("#333333"))
    c.setFont("Helvetica-Bold", 36)
    c.drawCentredString(width / 2, height / 2 + 20, "Print Production Test")

    c.setFont("Helvetica", 18)
    c.drawCentredString(width / 2, height / 2 - 20, "Page 2 — Light Background")

    # Edge markers
    c.setFillColor(HexColor("#ff00ff"))
    for x, y in [(5, 5), (width - 5, 5), (5, height - 5), (width - 5, height - 5)]:
        c.circle(x, y, 5, fill=1, stroke=0)

    c.showPage()
    c.save()
    print(f"Created test PDF: {output_path} ({width/72:.1f}\" × {height/72:.1f}\")")


if __name__ == "__main__":
    create_test_pdf()
