import subprocess
import sys
from docx import Document
from reportlab.pdfgen import canvas
from chivar_rtl_pdf import worker

def test_normalization():
    assert worker.normalize_text("علي كتاب") == "علی کتاب"
    assert worker.normalize_text("می‌روم") == "می‌روم"

def test_page_ranges():
    assert worker.parse_pages("1,3-4", 5) == [0, 2, 3]

def test_directional_runs():
    runs = worker.split_directional_runs("سلام ABC123")
    assert "".join(part for _, part in runs) == "سلام ABC123"
    assert {direction for direction, _ in runs} == {"rtl", "ltr"}

def test_native_pdf(tmp_path):
    source = tmp_path / "input.pdf"
    output = tmp_path / "output.docx"
    c = canvas.Canvas(str(source))
    c.drawString(72, 720, "A simple editable document for regression testing.")
    c.save()
    result = subprocess.run([sys.executable, "-m", "chivar_rtl_pdf.cli",
        str(source), str(output), "--ocr", "never"], capture_output=True, text=True, timeout=60)
    assert result.returncode == 0, result.stdout + result.stderr
    assert "simple editable document" in " ".join(p.text for p in Document(output).paragraphs)
