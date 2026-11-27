# Chivar RTL PDF

A Persian-first PDF-to-DOCX conversion engine extracted from the PDF processing component of [Chivar](https://chivar.ir). The standalone package is an early-stage project focused on editable document conversion.

## Why

PDF text extraction can lose word boundaries, reverse glyph order or damage mixed Persian and Latin identifiers. This engine compares extraction candidates and uses glyph geometry and directional DOCX runs to recover more usable editable text.

## Installation

Requires Python 3.10+. From a checkout:

```sh
python -m pip install .
chivar-rtl-pdf input.pdf output.docx --ocr never
```

For scanned documents, install Tesseract and its Persian (fas) and English (eng) language data, then use `--ocr auto` (default). On Debian/Ubuntu: `sudo apt install tesseract-ocr tesseract-ocr-fas tesseract-ocr-eng`.

Select pages with `--pages 1,3-5`. The CLI emits JSON progress and error records, some currently in Persian.

## Architecture

- pypdf and PDFium extraction with PDFMiner fallback
- Unicode normalization and Persian character mapping
- Geometry-informed token reconstruction
- Heuristic candidate scoring and agreement
- Persian/English OCR and image fallback
- DOCX paragraph and run direction properties

## Limitations

Editable output is not a pixel-perfect replica. Tables, formulas, complex layouts and mixed-direction text may need manual correction. Arabic script is recognized but Persian-specific normalization can alter Arabic spelling; Arabic fidelity is not established. Quality heuristics are not guarantees. Process only trusted PDFs in a resource-limited isolated environment; page limits alone do not prevent resource exhaustion. Do not expose this CLI directly as a public upload service.

## Development

```sh
python -m pip install -e ".[test]"
python -m pytest
```

The included regression tests are a starting point, not a comprehensive Persian benchmark.

## Release status

Version 0.1.0 is an early development version licensed under MIT. Third-party dependencies retain their respective licenses. Initial automated tests cover normalization, page selection, directional runs and native English PDF conversion; comprehensive Persian and OCR benchmarks remain planned.

## Roadmap

- Publish redistributable Persian and mixed-script regression fixtures.
- Measure character and word error rates against reference text.
- Improve OCR language selection and Arabic-specific normalization.
- Expand security and resource-limit testing.
- Prepare PyPI distribution.

