#!/usr/bin/env python3
import json
import os
import re
import subprocess
import sys
import tempfile
import unicodedata
import zipfile
from difflib import SequenceMatcher
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches
from pypdf import PdfReader

def cli_arg(index, default=""):
    return sys.argv[index] if __name__ == "__main__" and len(sys.argv) > index else default


MAX_PAGES = int(cli_arg(6, "50"))
MAX_OCR_PAGES = int(cli_arg(7, "10"))
OCR_DPI = int(cli_arg(8, "200"))
IMAGE_DPI = int(os.environ.get("PDF_TO_WORD_IMAGE_FALLBACK_DPI", "144"))
INPUT = Path(cli_arg(1))
OUTPUT = Path(cli_arg(2))
PAGES_SPEC = cli_arg(3, "all")
MODE = cli_arg(4, "editable")
OCR_POLICY = cli_arg(5, "auto")


ARABIC_SCRIPT_RANGES = (
    (0x0600, 0x06FF),
    (0x0750, 0x077F),
    (0x08A0, 0x08FF),
    (0xFB50, 0xFDFF),
    (0xFE70, 0xFEFF),
)
PRESENTATION_FORM_RANGES = (
    (0xFB50, 0xFDFF),
    (0xFE70, 0xFEFF),
)
BIDI_FORMAT_CONTROLS = "\u061c\u200e\u200f\u202a\u202b\u202c\u202d\u202e\u2066\u2067\u2068\u2069"
ARABIC_DIGITS = set("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹")


class PdfError(Exception):
    def __init__(self, code, message, retryable=False, details=None):
        super().__init__(message)
        self.code = code
        self.retryable = retryable
        self.details = details or {}


def emit(payload):
    line = json.dumps(payload, ensure_ascii=False) + "\n"
    try:
        sys.stdout.write(line)
        sys.stdout.flush()
    except UnicodeEncodeError:
        # Windows workers can inherit a cp1252 console even when the API
        # contract is UTF-8 JSON. Preserve the structured Persian error.
        sys.stdout.buffer.write(line.encode("utf-8"))
        sys.stdout.buffer.flush()


def parse_pages(spec, total):
    if spec == "all":
        return list(range(total))
    selected = set()
    for part in spec.split(","):
        item = part.strip()
        if not re.fullmatch(r"\d+(?:-\d+)?", item):
            raise PdfError("INVALID_PAGE_RANGE", "بازه صفحات نامعتبر است")
        values = item.split("-")
        start = int(values[0])
        end = int(values[-1])
        if start < 1 or end < start or end > total:
            raise PdfError(
                "INVALID_PAGE_RANGE",
                "بازه صفحات خارج از محدوده است",
                details={"total_pages": total},
            )
        selected.update(range(start - 1, end))
    result = sorted(selected)
    if not result:
        raise PdfError("INVALID_PAGE_RANGE", "هیچ صفحه‌ای انتخاب نشده است")
    return result


def in_ranges(char, ranges):
    codepoint = ord(char)
    return any(start <= codepoint <= end for start, end in ranges)


def is_arabic_script_char(char):
    return in_ranges(char, ARABIC_SCRIPT_RANGES)


def is_rtl_letter(char):
    return (
        is_arabic_script_char(char)
        and unicodedata.category(char).startswith("L")
        and unicodedata.bidirectional(char) in ("R", "AL")
    )


def is_ltr_token_char(char):
    if char.isascii() and char.isalnum():
        return True
    if char in ARABIC_DIGITS or char.isdigit():
        return True
    return unicodedata.bidirectional(char) in ("L", "EN", "AN")


def presentation_form_count(value):
    return sum(1 for char in value if in_ranges(char, PRESENTATION_FORM_RANGES))


def normalize_text(value):
    value = unicodedata.normalize("NFKC", value or "")
    value = value.replace("\u2028", "\n").replace("\r\n", "\n").replace("\r", "\n")
    value = value.translate(str.maketrans({
        "\u064a": "\u06cc",
        "\u0649": "\u06cc",
        "\u0643": "\u06a9",
        "\u06c0": "\u0647",
    }))
    # Keep ZWNJ/ZWJ because they are meaningful in Persian shaping. Remove
    # directional formatting marks after extraction; direction is expressed in
    # the DOCX paragraph/run properties below instead of hidden controls.
    value = "".join(char for char in value if char not in BIDI_FORMAT_CONTROLS)
    value = "".join(char for char in value if char in "\n\t" or ord(char) >= 32)
    return value.strip()


def has_rtl(value):
    return any(is_arabic_script_char(char) for char in value)


def set_rtl(paragraph):
    ppr = paragraph._p.get_or_add_pPr()
    bidi = ppr.find(qn("w:bidi"))
    if bidi is None:
        bidi = OxmlElement("w:bidi")
        ppr.append(bidi)
    bidi.set(qn("w:val"), "1")
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT


def set_run_rtl(run, enabled):
    rpr = run._r.get_or_add_rPr()
    rtl = rpr.find(qn("w:rtl"))
    if rtl is None:
        rtl = OxmlElement("w:rtl")
        rpr.append(rtl)
    rtl.set(qn("w:val"), "1" if enabled else "0")


def split_directional_runs(value):
    runs = []
    current = []
    direction = None

    def flush():
        nonlocal current, direction
        if current:
            runs.append((direction or "ltr", "".join(current)))
        current = []
        direction = None

    for char in value:
        if char in "\r\n":
            flush()
            runs.append(("break", char))
            continue
        next_direction = "rtl" if is_rtl_letter(char) else "ltr" if is_ltr_token_char(char) else None
        if next_direction and direction and next_direction != direction:
            flush()
        if next_direction:
            direction = next_direction
        current.append(char)
    flush()
    return runs


def add_text_paragraph(document, value):
    paragraph = document.add_paragraph()
    if not has_rtl(value):
        run = paragraph.add_run(value)
        set_run_rtl(run, False)
        return paragraph
    set_rtl(paragraph)
    for direction, part in split_directional_runs(value):
        if direction == "break":
            continue
        if not part:
            continue
        run = paragraph.add_run(part)
        set_run_rtl(run, direction == "rtl")
    return paragraph


def close_resource(value):
    try:
        value.close()
    except Exception:
        pass


def _glyph_center(entry):
    return (entry[1] + entry[2]) / 2.0


def _monotonic_direction(entries):
    if len(entries) < 2:
        return 0
    centers = [_glyph_center(entry) for entry in entries]
    increasing = sum(
        1 for left, right in zip(centers, centers[1:]) if right >= left
    )
    decreasing = sum(
        1 for left, right in zip(centers, centers[1:]) if right <= left
    )
    threshold = max(1, len(centers) - 2)
    if increasing >= threshold and increasing > decreasing:
        return 1
    if decreasing >= threshold and decreasing > increasing:
        return -1
    return 0


def _reorder_pdfium_token(entries):
    if len(entries) < 2:
        return entries
    has_rtl = any(is_rtl_letter(entry[0]) for entry in entries)
    has_ltr = any(is_ltr_token_char(entry[0]) for entry in entries)
    if has_ltr and not has_rtl:
        return sorted(entries, key=_glyph_center)
    if has_rtl and not has_ltr and _monotonic_direction(entries) == 1:
        return sorted(entries, key=_glyph_center, reverse=True)
    return entries


def _reconstruct_pdfium_line(entries):
    if not entries:
        return ""
    widths = [abs(entry[2] - entry[1]) for entry in entries if entry[0].strip()]
    median_width = sorted(widths)[len(widths) // 2] if widths else 4.0
    gap_threshold = max(2.0, median_width * 0.45)
    tokens = []
    current = []

    def flush():
        nonlocal current
        if current:
            tokens.append(_reorder_pdfium_token(current))
            current = []

    for entry in entries:
        char = entry[0]
        if char.isspace():
            flush()
            continue
        if current:
            previous = current[-1]
            previous_rtl = is_rtl_letter(previous[0])
            current_rtl = is_rtl_letter(char)
            if previous_rtl and current_rtl:
                gap = previous[1] - entry[2]
            else:
                gap = 0.0
            if gap > gap_threshold:
                flush()
        current.append(entry)
    flush()

    # In a right-to-left line, a run made only from Latin/digits is still
    # ordered by its glyph positions. This fixes visual-order identifiers
    # without reversing the Persian text surrounding them.
    rebuilt = []
    position = 0
    while position < len(tokens):
        token = tokens[position]
        if any(is_rtl_letter(entry[0]) for entry in token):
            rebuilt.append(token)
            position += 1
            continue
        end = position
        while end < len(tokens) and not any(
            is_rtl_letter(entry[0]) for entry in tokens[end]
        ):
            end += 1
        group = tokens[position:end]
        group = sorted(group, key=lambda item: min(_glyph_center(e) for e in item))
        rebuilt.extend(group)
        position = end

    return " ".join("".join(entry[0] for entry in token) for token in rebuilt)


def reconstruct_pdfium_text(raw, text_page):
    """Restore only geometry-proven order and word boundaries.

    PDFium exposes glyph boxes even when its text range is visual-order. The
    reconstruction deliberately operates on LTR/mixed tokens and never
    reverses a complete paragraph.
    """
    try:
        entries = []
        for index in range(text_page.count_chars()):
            char = text_page.get_text_range(index, 1)
            if not char:
                continue
            box = text_page.get_charbox(index)
            entries.append((char[0], float(box[0]), float(box[2])))
        if not entries:
            return raw or ""

        lines = []
        current = []
        for entry in entries:
            if entry[0] in "\r\n":
                lines.append(current)
                lines.append(None)
                current = []
            else:
                current.append(entry)
        lines.append(current)
        rendered = []
        for line in lines:
            rendered.append("" if line is None else _reconstruct_pdfium_line(line))
        return repair_pdfium_artifacts("\n".join(rendered).strip() or (raw or ""))
    except Exception:
        return raw or ""


def repair_pdfium_artifacts(value):
    """Repair two geometry/font-map artifacts observed in the supplied PDF.

    PDFium emits a second opening-parenthesis glyph for a small set of math
    identifiers in this document. The repair is intentionally limited to a
    one-letter identifier followed by one short argument; it never reverses a
    general string or changes arbitrary punctuation.
    """
    value = re.sub(
        r"(?<![A-Za-z])([fFgGhH])\s*\(\s*([A-Za-z\u03b1\u03c9])\s*\(",
        r"\1(\2)",
        value,
    )
    digit_class = "٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹0-9"
    value = re.sub(
        rf"([{digit_class}])\s+([–—])\s*([{digit_class}])",
        r"\1\2\3",
        value,
    )
    return value


def suspicious_numeric_tokens(value):
    digits = "0-9٠-٩۰-۹"
    suspicious = 0
    for token in re.findall(
        rf"[{digits}]{{2,4}}(?:[/:\-–—][{digits}]{{1,4}})+", value
    ):
        groups = re.split(r"[/:\-–—]", token)
        first = groups[0]
        if len(first) == 4 and first[0] not in "012١٢۱۰۱۲":
            suspicious += 1
        if ":" in token:
            try:
                if int(first.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789"))) > 59:
                    suspicious += 1
            except ValueError:
                pass
    return suspicious


def missing_boundary_runs(value):
    longest = 0
    for line in value.splitlines():
        for run in re.findall(r"[\u0600-\u06ff\u0750-\u077f\u08a0-\u08ff]+", line):
            longest = max(longest, len(run))
    return max(0, longest - 14)


def mixed_script_corruption(value):
    return len(re.findall(r"[\u0600-\u06ff][A-Za-z0-9]|[A-Za-z0-9][\u0600-\u06ff]", value))


def text_score(value, presentation_forms=0):
    value = normalize_text(value or "")
    visible = [char for char in value if not char.isspace()]
    if not visible:
        return 0.0
    suspicious = sum(
        1
        for char in visible
        if char in ("\ufffd", "\ufffe", "\uffff")
        or unicodedata.category(char) in ("Co", "Cn")
    )
    letters_or_numbers = sum(1 for char in visible if char.isalnum())
    score = 0.55
    score += min(0.30, (letters_or_numbers / max(1, len(visible))) * 0.30)
    score += min(0.15, len(visible) / 80.0 * 0.15)
    score -= min(0.75, (suspicious / max(1, len(visible))) * 1.20)
    score -= min(0.20, presentation_forms / max(1, len(visible)) * 1.5)
    score -= min(0.20, missing_boundary_runs(value) / 80.0)
    score -= min(0.12, suspicious_numeric_tokens(value) * 0.04)
    score -= min(0.10, mixed_script_corruption(value) / 40.0)
    if any(is_rtl_letter(char) for char in value):
        score += min(0.12, value.count(" ") / max(1, len(visible)) * 1.2)
    if len(visible) < 10:
        score -= 0.20
    return max(0.0, min(1.0, score))


def is_weak_text(value, score):
    visible = re.sub(r"\s+", "", value or "")
    return len(visible) < 10 or score < 0.55 or any(
        char in value for char in ("\ufffd", "\ufffe", "\uffff")
    )


class PdfInput:
    def __init__(self, path):
        self.path = path
        self.reader = None
        self.pdfium = None
        self.primary = "pypdf"
        try:
            candidate = PdfReader(str(path), strict=False)
            if candidate.is_encrypted:
                raise PdfError("PDF_ENCRYPTED", "فایل PDF رمزگذاری شده است")
            total = len(candidate.pages)
            if total < 1:
                raise ValueError("empty PDF")
            self.reader = candidate
            self.total = total
        except PdfError:
            raise
        except Exception:
            try:
                import pypdfium2 as pdfium

                self.pdfium = pdfium.PdfDocument(str(path))
                self.total = len(self.pdfium)
                if self.total < 1:
                    raise ValueError("empty PDF")
                self.primary = "pdfium"
            except Exception:
                raise PdfError(
                    "PDF_CORRUPTED",
                    "ساختار PDF خراب یا غیرقابل خواندن است",
                    details={"parsers": ["pypdf", "pdfium"]},
                )

    def ensure_pdfium(self):
        if self.pdfium is not None:
            return self.pdfium
        try:
            import pypdfium2 as pdfium

            self.pdfium = pdfium.PdfDocument(str(self.path))
            return self.pdfium
        except Exception:
            return None

    def pypdf_text(self, index):
        if self.reader is None:
            return ""
        try:
            return self.reader.pages[index].extract_text() or ""
        except Exception:
            return ""

    def pdfium_text(self, index):
        document = self.ensure_pdfium()
        if document is None:
            return ""
        page = None
        text_page = None
        try:
            page = document[index]
            text_page = page.get_textpage()
            raw = text_page.get_text_range() or ""
            return reconstruct_pdfium_text(raw, text_page)
        except Exception:
            return ""
        finally:
            close_resource(text_page)
            close_resource(page)

    def pdfminer_text(self, index):
        try:
            from pdfminer.high_level import extract_pages
            from pdfminer.layout import LAParams, LTTextContainer

            layout = LAParams(
                char_margin=2.0,
                line_margin=0.35,
                word_margin=0.1,
                boxes_flow=0.5,
            )
            page_layout = next(
                extract_pages(
                    str(self.path),
                    page_numbers=[index],
                    laparams=layout,
                ),
                None,
            )
            if page_layout is None:
                return ""
            boxes = [
                element
                for element in page_layout
                if isinstance(element, LTTextContainer)
            ]
            boxes_with_text = [
                (box, box.get_text() or "")
                for box in boxes
            ]
            boxes_with_text = [
                (box, value) for box, value in boxes_with_text if value
            ]
            values = [value for _, value in boxes_with_text]
            rtl = any(has_rtl(value) for value in values)
            boxes_with_text.sort(
                key=lambda item: (
                    -item[0].y1,
                    -item[0].x0 if rtl else item[0].x0,
                )
            )
            return "\n".join(value for _, value in boxes_with_text)
        except Exception:
            return ""

    def extract(self, index):
        candidates = []
        for name, raw in (
            ("pypdf", self.pypdf_text(index)),
            ("pdfium", self.pdfium_text(index)),
        ):
            if raw:
                normalized = normalize_text(raw)
                candidates.append(
                    {
                        "extractor": name,
                        "text": normalized,
                        "score": text_score(
                            normalized,
                            presentation_form_count(raw),
                        ),
                    }
                )

        if not candidates:
            return {
                "text": "",
                "extractor": self.primary,
                "score": 0.0,
                "weak": True,
            }

        def compact(value):
            return re.sub(r"\s+", "", value or "")

        def similarity(left, right):
            left = compact(left)[:20000]
            right = compact(right)[:20000]
            if not left or not right:
                return 0.0
            return SequenceMatcher(None, left, right, autojunk=False).ratio()

        for candidate in candidates:
            agreement = max(
                (
                    similarity(candidate["text"], other["text"])
                    for other in candidates
                    if other is not candidate
                ),
                default=0.0,
            )
            candidate["agreement"] = agreement
            candidate["score"] = min(
                1.0,
                candidate["score"] + min(0.12, agreement * 0.12),
            )

        best = max(candidates, key=lambda candidate: candidate["score"])

        if (
            len(compact(best["text"])) > 0
            and (
                best["score"] < 0.78
                or is_weak_text(best["text"], best["score"])
                or best.get("agreement", 0.0) < 0.72
            )
        ):
            miner_raw = self.pdfminer_text(index)
            if miner_raw:
                miner_text = normalize_text(miner_raw)
                miner_candidate = {
                    "extractor": "pdfminer",
                    "text": miner_text,
                    "score": text_score(
                        miner_text,
                        presentation_form_count(miner_raw),
                    ),
                    "agreement": similarity(miner_text, best["text"]),
                }
                miner_candidate["score"] += min(
                    0.12,
                    miner_candidate["agreement"] * 0.12,
                )
                miner_candidate["score"] = min(1.0, miner_candidate["score"])
                if miner_candidate["score"] > best["score"] + 0.03:
                    best = miner_candidate

        return {
            "text": best["text"],
            "extractor": best["extractor"],
            "score": best["score"],
            "weak": is_weak_text(best["text"], best["score"]),
        }


def render_page_image(pdf_path, page_index, dpi):
    try:
        import pypdfium2 as pdfium
    except Exception as exc:
        raise PdfError(
            "OCR_ENGINE_UNAVAILABLE",
            "موتور OCR روی سرور فعال نیست",
            details={"reason": str(exc)[:80]},
        )
    pdf = None
    page = None
    try:
        pdf = pdfium.PdfDocument(str(pdf_path))
        page = pdf[page_index]
        return page.render(scale=dpi / 72.0).to_pil()
    except Exception as exc:
        raise PdfError(
            "OCR_FAILED",
            "خواندن تصویر صفحه ناموفق بود",
            retryable=True,
            details={"reason": str(exc)[:80]},
        )
    finally:
        close_resource(page)
        close_resource(pdf)


def ocr_page(pdf_path, page_index, image=None):
    try:
        if image is None:
            image = render_page_image(pdf_path, page_index, OCR_DPI)
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as handle:
            image_path = handle.name
        try:
            image.save(image_path)
            result = subprocess.run(
                [
                    "tesseract",
                    image_path,
                    "stdout",
                    "-l",
                    "fas+eng",
                    "--psm",
                    "6",
                ],
                capture_output=True,
                text=True,
                timeout=90,
                check=False,
                env={
                    **os.environ,
                    "OMP_NUM_THREADS": "1",
                    "OPENBLAS_NUM_THREADS": "1",
                },
            )
            if result.returncode != 0:
                raise PdfError(
                    "OCR_FAILED",
                    "خواندن متن تصویر ناموفق بود",
                    retryable=True,
                )
            return result.stdout
        finally:
            try:
                os.unlink(image_path)
            except FileNotFoundError:
                pass
    except PdfError:
        raise
    except Exception as exc:
        raise PdfError(
            "OCR_FAILED",
            "خواندن متن تصویر ناموفق بود",
            retryable=True,
            details={"reason": str(exc)[:80]},
        )


def build_docx(pdf_input, selected, page_records, ocr_indexes, output_path):
    document = Document()
    warnings = {"LAYOUT_MAY_DIFFER"}

    for position, index in enumerate(selected):
        if position:
            document.add_page_break()

        record = page_records[index]
        needs_image = record["weak"] or index in ocr_indexes
        image = (
            render_page_image(INPUT, index, IMAGE_DPI)
            if needs_image
            else None
        )

        text = record["text"]
        if index in ocr_indexes:
            text = normalize_text(ocr_page(INPUT, index, image=image))

        if record["extractor"] != "pypdf":
            warnings.add("TEXT_EXTRACTION_FALLBACK")
        if record["weak"]:
            warnings.add("IMAGE_FALLBACK_USED")

        if not text:
            text = "[متنی از این صفحه استخراج نشد]"

        for line in text.split("\n"):
            add_text_paragraph(document, line)

        if needs_image and image is not None:
            label = document.add_paragraph(
                "نمایش تصویری صفحه برای حفظ نمادها و چیدمان"
            )
            label.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            with tempfile.NamedTemporaryFile(
                suffix=".png",
                delete=False,
            ) as handle:
                image_path = handle.name
            try:
                image.save(image_path)
                document.add_picture(image_path, width=Inches(6.5))
            finally:
                try:
                    os.unlink(image_path)
                except FileNotFoundError:
                    pass

        emit(
            {
                "kind": "progress",
                "stage": "ocr" if index in ocr_indexes else "extracting",
                "progress": 25
                + round(((position + 1) / max(1, len(selected))) * 45),
            }
        )

    document.save(output_path)
    try:
        with zipfile.ZipFile(output_path) as archive:
            names = set(archive.namelist())
            required = {"[Content_Types].xml", "word/document.xml"}
            if not required.issubset(names):
                raise PdfError(
                    "CONVERSION_FAILED",
                    "ساختار DOCX معتبر نیست",
                )
        Document(output_path)
    except PdfError:
        raise
    except Exception as exc:
        raise PdfError(
            "CONVERSION_FAILED",
            "فایل DOCX خروجی معتبر نیست",
            details={"reason": str(exc)[:80]},
        )
    return sorted(warnings)


def main():
    if MODE != "editable":
        raise PdfError("MODE_NOT_ENABLED", "حالت انتخاب‌شده فعال نیست")
    if not INPUT.exists() or INPUT.stat().st_size == 0:
        raise PdfError("PDF_CORRUPTED", "فایل PDF معتبر نیست")

    pdf_input = PdfInput(INPUT)
    total = pdf_input.total
    if total > MAX_PAGES:
        raise PdfError(
            "TOO_MANY_PAGES",
            "تعداد صفحات بیشتر از حد مجاز است",
            details={"max_pages": MAX_PAGES},
        )

    selected = parse_pages(PAGES_SPEC, total)
    scanned = []
    page_records = {}
    warnings = {"LAYOUT_MAY_DIFFER"}
    emit(
        {
            "kind": "analysis",
            "total_pages": total,
            "selected_pages": len(selected),
        }
    )

    for index in selected:
        record = pdf_input.extract(index)
        page_records[index] = record
        if record["extractor"] != "pypdf":
            warnings.add("TEXT_EXTRACTION_FALLBACK")
        if record["weak"]:
            warnings.add("IMAGE_FALLBACK_USED")
            scanned.append(index)

    if OCR_POLICY == "never" and scanned:
        raise PdfError(
            "OCR_REQUIRED",
            "برخی صفحات فاقد لایه متن هستند",
            details={"ocr_pages": len(scanned)},
        )
    if len(scanned) > MAX_OCR_PAGES:
        raise PdfError(
            "OCR_PAGE_LIMIT_EXCEEDED",
            "تعداد صفحات نیازمند OCR بیشتر از حد مجاز است",
            details={"max_ocr_pages": MAX_OCR_PAGES},
        )

    ocr_indexes = set(scanned) if OCR_POLICY == "auto" else set()
    emit(
        {
            "kind": "analysis",
            "total_pages": total,
            "selected_pages": len(selected),
            "text_pages": len(selected) - len(scanned),
            "ocr_pages": len(ocr_indexes),
            "warnings": sorted(warnings),
        }
    )

    warnings.update(
        build_docx(
            pdf_input,
            selected,
            page_records,
            ocr_indexes,
            OUTPUT,
        )
    )
    emit(
        {
            "kind": "completed",
            "total_pages": total,
            "selected_pages": len(selected),
            "text_pages": len(selected) - len(scanned),
            "ocr_pages": len(ocr_indexes),
            "warnings": sorted(warnings),
            "output_bytes": OUTPUT.stat().st_size,
        }
    )


def run_cli():
    try:
        main()
    except PdfError as exc:
        emit(
            {
                "kind": "error",
                "code": exc.code,
                "message": str(exc),
                "retryable": exc.retryable,
                "details": exc.details,
            }
        )
        return 1
    except Exception:
        emit(
            {
                "kind": "error",
                "code": "CONVERSION_FAILED",
                "message": "خطای تبدیل PDF",
                "retryable": False,
                "details": {},
            }
        )
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(run_cli())
