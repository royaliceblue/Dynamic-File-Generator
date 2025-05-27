#!/usr/bin/env python3
"""
Generate large files in various formats (DOCX, XLSX, PPTX, PDF, PST, ZIP) with valid structure.

Dependencies (install via pip):
  python-docx, openpyxl, python-pptx, reportlab, Aspose.Email-for-Python-via-NET
"""
import os
import io
import argparse
import sys
import zipfile

# Document libraries
from docx import Document
from openpyxl import Workbook
from pptx import Presentation
from reportlab.pdfgen import canvas

# PST generation via Aspose.Email-for-Python-via-NET
try:
    from aspose.email.storage.pst import PersonalStorage, FileFormatVersion
except ImportError:
    PersonalStorage = None


def pad_stream(buf: io.BytesIO, target_bytes: int) -> bytes:
    data = buf.getvalue()
    pad_len = target_bytes - len(data)
    if pad_len < 0:
        raise ValueError(f"Minimum file size {len(data)} exceeds target {target_bytes} bytes")
    return data + (b"\0" * pad_len)


def generate_docx(size_mb: int, output: str):
    buf = io.BytesIO()
    doc = Document()
    doc.add_paragraph('Sample')
    doc.save(buf)
    data = pad_stream(buf, size_mb * 1024 * 1024)
    with open(output, 'wb') as f:
        f.write(data)
    print(f"DOCX written '{output}' (~{os.path.getsize(output)/(1024*1024):.2f} MB)")


def generate_xlsx(size_mb: int, output: str):
    buf = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws['A1'] = 'Sample'
    wb.save(buf)
    data = pad_stream(buf, size_mb * 1024 * 1024)
    with open(output, 'wb') as f:
        f.write(data)
    print(f"XLSX written '{output}' (~{os.path.getsize(output)/(1024*1024):.2f} MB)")


def generate_pptx(size_mb: int, output: str):
    buf = io.BytesIO()
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    prs.save(buf)
    data = pad_stream(buf, size_mb * 1024 * 1024)
    with open(output, 'wb') as f:
        f.write(data)
    print(f"PPTX written '{output}' (~{os.path.getsize(output)/(1024*1024):.2f} MB)")


def generate_pdf(size_mb: int, output: str):
    buf = io.BytesIO()
    c = canvas.Canvas(buf)
    c.drawString(100, 750, 'Sample')
    c.showPage()
    c.save()
    data = pad_stream(buf, size_mb * 1024 * 1024)
    with open(output, 'wb') as f:
        f.write(data)
    print(f"PDF written '{output}' (~{os.path.getsize(output)/(1024*1024):.2f} MB)")


def generate_pst(size_mb: int, output: str):
    if PersonalStorage is None:
        print("Error: Aspose.Email-for-Python-via-NET is required for PST generation", file=sys.stderr)
        sys.exit(1)
    # Create a new Unicode PST file
    with PersonalStorage.create(output, FileFormatVersion.UNICODE) as pst:
        pst.root_folder.add_sub_folder("Inbox")
    # Pad file to target size
    current = os.path.getsize(output)
    target = size_mb * 1024 * 1024
    pad_len = target - current
    if pad_len < 0:
        raise ValueError(f"Generated PST size {current} exceeds target {target}")
    with open(output, 'ab') as f:
        f.write(b"\0" * pad_len)
    print(f"PST written '{output}' (~{os.path.getsize(output)/(1024*1024):.2f} MB)")


def generate_zip(size_mb: int, output: str):
    """
    Generate a valid ZIP containing a dummy file and a large pad.bin entry to reach target size.
    """
    target = size_mb * 1024 * 1024
    # Create zip with a small dummy.txt
    with zipfile.ZipFile(output, 'w', compression=zipfile.ZIP_STORED) as z:
        z.writestr('dummy.txt', 'Sample')

    # Determine how many bytes to pad
    current = os.path.getsize(output)
    pad_bytes = target - current
    if pad_bytes < 0:
        raise ValueError(f"Generated ZIP size {current} exceeds target {target}")

    # Append a pad.bin entry stored without compression
    with zipfile.ZipFile(output, 'a', compression=zipfile.ZIP_STORED) as z:
        info = zipfile.ZipInfo('pad.bin')
        info.compress_type = zipfile.ZIP_STORED
        with z.open(info, 'w') as f:
            chunk = b"\0" * (1024 * 1024)
            full_chunks = pad_bytes // len(chunk)
            for _ in range(full_chunks):
                f.write(chunk)
            # write remaining bytes
            remainder = pad_bytes - full_chunks * len(chunk)
            if remainder:
                f.write(b"\0" * remainder)

    print(f"ZIP written '{output}' (~{os.path.getsize(output)/(1024*1024):.2f} MB)")


def main():
    parser = argparse.ArgumentParser(description="Generate large valid files of various formats.")
    parser.add_argument('size', type=int, help='Target size in megabytes')
    parser.add_argument('-f', '--format', choices=['docx','xlsx','pptx','pdf','pst','zip'], default='docx', help='File format')
    parser.add_argument('output', nargs='?', help='Output filename')
    args = parser.parse_args()
    out = args.output or f"output.{args.format}"

    if args.format == 'docx':
        generate_docx(args.size, out)
    elif args.format == 'xlsx':
        generate_xlsx(args.size, out)
    elif args.format == 'pptx':
        generate_pptx(args.size, out)
    elif args.format == 'pdf':
        generate_pdf(args.size, out)
    elif args.format == 'pst':
        generate_pst(args.size, out)
    elif args.format == 'zip':
        generate_zip(args.size, out)

if __name__ == '__main__':
    main()
