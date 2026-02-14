#!/usr/bin/env python3
"""Verify the output PDF structure and content streams."""

import pikepdf
from pikepdf import Name
import sys


def verify(path):
    pdf = pikepdf.open(path)

    for i, page in enumerate(pdf.pages):
        print(f"=== Page {i+1} ===")

        # Check boxes
        mbox = [float(v) for v in page.mediabox]
        print(f"MediaBox: {mbox}")

        if Name.TrimBox in page.obj:
            tbox = [float(v) for v in page.obj[Name.TrimBox]]
            print(f"TrimBox:  {tbox}")
        else:
            print("TrimBox:  MISSING!")

        if Name.BleedBox in page.obj:
            bbox = [float(v) for v in page.obj[Name.BleedBox]]
            print(f"BleedBox: {bbox}")
        else:
            print("BleedBox: MISSING!")

        # Check resources
        res = page.obj.get(Name.Resources, {})
        print(f"Resources keys: {list(res.keys())}")

        if Name.XObject in res:
            xobjs = res[Name.XObject]
            print(f"XObjects: {list(xobjs.keys())}")
            for name, xobj in xobjs.items():
                print(f"  {name}: Type={xobj.get(Name.Type)}, "
                      f"Subtype={xobj.get(Name.Subtype)}, "
                      f"BBox={[float(v) for v in xobj.get(Name.BBox, [])]}")
                # Check that the form XObject has resources
                if Name.Resources in xobj:
                    form_res = xobj[Name.Resources]
                    print(f"    Form resources: {list(form_res.keys())}")

        # Check content stream
        contents = page.obj[Name.Contents]
        if isinstance(contents, pikepdf.Array):
            data = b""
            for s in contents:
                data += s.read_bytes()
        else:
            data = contents.read_bytes()

        content_str = data.decode("latin-1")
        print(f"\nContent stream ({len(content_str)} bytes):")
        # Print first 500 chars
        print(content_str[:800])
        print("...")
        # Print last 500 chars (trim marks area)
        if len(content_str) > 800:
            print(content_str[-600:])
        print()


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else "test_output.pdf"
    verify(path)
