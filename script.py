#!/usr/bin/env python3
import argparse
import os
import re
import zipfile
from zipfile import ZipInfo, ZIP_STORED, ZIP_DEFLATED

# Office libraries
from docx import Document
from openpyxl import Workbook
from pptx import Presentation
from reportlab.pdfgen import canvas

# PST via Aspose.Email-for-Python-via-NET
from aspose.email.storage.pst import PersonalStorage, FileFormatVersion

_UNIT_RE = re.compile(r'^(?P<value>\d+(\.\d+)?)(?P<unit>KB|MB)?$', re.IGNORECASE)

def parse_args():
    parser = argparse.ArgumentParser(
        description="Generate a valid dummy file of arbitrary size (KB or MB)"
    )
    parser.add_argument(
        "-f", "--format",
        choices=["docx", "xlsx", "pptx", "pdf", "pst", "zip"],
        default="docx",
        help="Which format to produce"
    )
    parser.add_argument(
        "size",
        help="Target size (e.g. 150KB, 2.5MB, or 10 for 10MB)"
    )
    parser.add_argument(
        "output",
        nargs="?",
        help="Output filename (default: <format>_<size>.<ext>)"
    )
    if hasattr(parser, "parse_intermixed_args"):
        return parser.parse_intermixed_args()
    return parser.parse_args()

def parse_size(size_str: str) -> int:
    """
    Parse sizes like '150KB', '2.5MB', or '10' (MB default) into bytes.
    """
    m = _UNIT_RE.match(size_str.strip())
    if not m:
        raise argparse.ArgumentTypeError(
            f"Invalid size '{size_str}'. Use e.g. 150KB, 2.5MB, or 10"
        )
    val = float(m.group('value'))
    unit = (m.group('unit') or 'MB').upper()
    if unit == 'KB':
        return int(val * 1024)
    # MB
    return int(val * 1024 * 1024)

def _embed_pad_in_zip(zip_path: str, media_folder: str, target_bytes: int):
    """
    Patch [Content_Types].xml and add pad.bin so the total archive size == target_bytes.
    """
    current = os.path.getsize(zip_path)
    pad_bytes = target_bytes - current
    if pad_bytes < 0:
        raise ValueError(f"{zip_path!r} is already {current} bytes, exceeds target {target_bytes}")

    with zipfile.ZipFile(zip_path, "a", compression=ZIP_STORED) as z:
        # 1) patch content types to accept .bin
        xml = z.read("[Content_Types].xml").decode("utf-8")
        if 'Extension="bin"' not in xml:
            patched = xml.replace(
                "</Types>",
                '  <Default Extension="bin" ContentType="application/octet-stream"/>\n</Types>'
            )
            z.writestr("[Content_Types].xml", patched, compress_type=ZIP_DEFLATED)

        # 2) embed pad.bin
        zi = ZipInfo(f"{media_folder}/pad.bin")
        zi.compress_type = ZIP_STORED
        chunk = 1024 * 1024
        full, rem = divmod(pad_bytes, chunk)
        with z.open(zi, "w") as f:
            for _ in range(full):
                f.write(b"\0" * chunk)
            if rem:
                f.write(b"\0" * rem)

def pad_file_trailer(path: str, target_bytes: int):
    """
    Append raw zero‐bytes after EOF until file size == target_bytes.
    Used for PDF, PST, and plain ZIP fallback.
    """
    current = os.path.getsize(path)
    pad_bytes = target_bytes - current
    if pad_bytes < 0:
        raise ValueError(f"{path!r} is already {current} bytes, exceeds target {target_bytes}")
    with open(path, "ab") as f:
        f.write(b"\0" * pad_bytes)

def generate_docx(target_bytes: int, output: str):
    doc = Document()
    doc.add_paragraph(" ")
    doc.save(output)
    _embed_pad_in_zip(output, "word/media", target_bytes)

def generate_xlsx(target_bytes: int, output: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    wb.save(output)
    _embed_pad_in_zip(output, "xl/media", target_bytes)

def generate_pptx(target_bytes: int, output: str):
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[0])
    prs.save(output)
    _embed_pad_in_zip(output, "ppt/media", target_bytes)

def generate_pdf(target_bytes: int, output: str):
    c = canvas.Canvas(output)
    c.drawString(100, 750, "")
    c.showPage()
    c.save()
    pad_file_trailer(output, target_bytes)

def generate_pst(target_bytes: int, output: str):
    with PersonalStorage.create(output, FileFormatVersion.UNICODE) as pst:
        pst.root_folder.add_sub_folder("Inbox")
    pad_file_trailer(output, target_bytes)

def generate_zip(target_bytes: int, output: str):
    # create minimal ZIP
    with zipfile.ZipFile(output, "w", compression=ZIP_STORED) as z:
        z.writestr("dummy.txt", "placeholder")
    # embed padding inside ZIP root
    _embed_pad_in_zip(output, "", target_bytes)

def main():
    args = parse_args()
    target_bytes = parse_size(args.size)
    fmt = args.format
    out = args.output or f"{fmt}_{args.size.lower()}.{fmt}"

    generators = {
        "docx": generate_docx,
        "xlsx": generate_xlsx,
        "pptx": generate_pptx,
        "pdf":  generate_pdf,
        "pst":  generate_pst,
        "zip":  generate_zip,
    }
    generators[fmt](target_bytes, out)
    actual = os.path.getsize(out)
    print(f"✅ Generated '{out}' ({actual} bytes, ≈{actual/1024/1024:.2f} MB)")

if __name__ == "__main__":
    main()
