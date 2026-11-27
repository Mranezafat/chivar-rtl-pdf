import argparse
import subprocess
import sys
from pathlib import Path

def main():
    p = argparse.ArgumentParser(description="Convert Persian-first PDFs to editable DOCX.")
    p.add_argument("input", type=Path)
    p.add_argument("output", type=Path)
    p.add_argument("--pages", default="all")
    p.add_argument("--ocr", choices=["auto", "never"], default="auto")
    p.add_argument("--max-pages", type=int, default=50)
    p.add_argument("--max-ocr-pages", type=int, default=10)
    p.add_argument("--ocr-dpi", type=int, default=200)
    a = p.parse_args()
    if not a.input.is_file() or a.input.resolve() == a.output.resolve():
        p.error("Input must exist and differ from output")
    if a.output.exists():
        p.error("Output already exists")
    if a.max_pages < 1 or a.max_ocr_pages < 0 or not 72 <= a.ocr_dpi <= 600:
        p.error("Invalid page limits or OCR DPI")
    return subprocess.call([sys.executable, "-m", "chivar_rtl_pdf.worker",
        str(a.input), str(a.output), a.pages, "editable", a.ocr,
        str(a.max_pages), str(a.max_ocr_pages), str(a.ocr_dpi)])

if __name__ == "__main__":
    sys.exit(main())
