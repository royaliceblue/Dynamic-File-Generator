#!/usr/bin/env python3
import argparse
import os
import io
import zipfile
from zipfile import ZipInfo, ZIP_STORED

# DOCX
from docx import Document
# XLSX
from openpyxl import Workbook
# PPTX
from pptx import Presentation
# PDF
from reportlab.pdfgen import canvas
# PST via Aspose
from aspose.email.storage.pst import PersonalStorage, FileFormatVersion

def parse_args():
    parser = argparse.ArgumentParser(
        description="Generate a dummy file of arbitrary size"
    )
    parser.add_argument(
        "-f", "--format",
        choices=["docx","xlsx","pptx","pdf","pst","zip"],
        default="docx",
        help="Which format to produce"
    )
    parser.add_argument(
        "size",
        type=int,
        help="Target file size in megabytes"
    )
    parser.add_argument(
        "output",
        nargs="?",
        help="Output filename (default: <format>_<size>.<ext>)"
    )

    # Use intermixed parser if available (allows mixing pos/opts freely)
    if hasattr(parser, "parse_intermixed_args"):
        return parser.parse_intermixed_args()
    else:
        return parser.parse_args()

def pad_file(path, size_mb):
    target = size_mb * 1024 * 1024
    current = os.path.getsize(path)
    if current > target:
        raise ValueError(f"{path!r} is already {current} bytes, exceeds {target}")
    with open(path, "ab") as f:
        f.write(b"\0" * (target - current))

def generate_docx(size_mb, output):
    doc = Document()
    doc.add_paragraph("")  
    doc.save(output)
    pad_file(output, size_mb)

def generate_xlsx(size_mb, output):
    wb = Workbook()
    wb.active.title = "Sheet1"
    wb.save(output)
    pad_file(output, size_mb)

def generate_pptx(size_mb, output):
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[0])
    prs.save(output)
    pad_file(output, size_mb)

def generate_pdf(size_mb, output):
    c = canvas.Canvas(output)
    c.drawString(100, 750, "")
    c.showPage()
    c.save()
    pad_file(output, size_mb)

def generate_pst(size_mb, output):
    with PersonalStorage.create(output, FileFormatVersion.UNICODE) as pst:
        pst.root_folder.add_sub_folder("Inbox")
    pad_file(output, size_mb)

def generate_zip(size_mb, output):
    # create minimal zip
    with zipfile.ZipFile(output, "w", compression=ZIP_STORED) as z:
        z.writestr("dummy.txt", "placeholder")
    # append pad.bin
    with zipfile.ZipFile(output, "a", compression=ZIP_STORED) as z:
        info = ZipInfo("pad.bin")
        info.compress_type = ZIP_STORED
        pad_bytes = size_mb*1024*1024 - os.path.getsize(output)
        chunk = 1024*1024
        full, rem = divmod(pad_bytes, chunk)
        with z.open(info, "w") as f:
            for _ in range(full):
                f.write(b"\0" * chunk)
            if rem:
                f.write(b"\0" * rem)

def main():
    args = parse_args()
    size_mb = args.size
    fmt     = args.format
    out     = args.output or f"{fmt}_{size_mb}.{fmt}"

    dispatch = {
        "docx": generate_docx,
        "xlsx": generate_xlsx,
        "pptx": generate_pptx,
        "pdf":  generate_pdf,
        "pst":  generate_pst,
        "zip":  generate_zip,
    }

    dispatch[fmt](size_mb, out)
    print(f"Generated {out!r} ({size_mb} MB, format={fmt})")

if __name__ == "__main__":
    main()
