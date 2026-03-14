# Quality Reviewer (Source vs Processed XML)

This tool compares a **source document** (PDF/DOCX/TXT/CSV) against a **processed XML** and produces a **QC report** (CSV/Excel/JSON) plus an optional HTML summary.

## Features
- Configuration-driven field mapping
  - Source extraction via **regex** (for text/PDF/DOCX/TXT) or **csv_column** (for CSV)
  - XML extraction via simple XPath (compatible with `xml.etree.ElementTree`)
- Comparison modes: `exact`, `fuzzy`, `numeric` (with absolute/percent tolerance), `date` (with day tolerance), `enum`, `bool`
- Normalizations: `strip`, `lower`, `upper`, `collapse_whitespace`, `alnum_only`, `remove_punctuation`
- Reports: CSV, XLSX, JSON, and HTML summary with PASS/FAIL color-coding

## Quick Start
1. Place your files:
   - `source.pdf` (or `.docx`/`.txt`/`.csv`)
   - `processed.xml`
   - `sample_config.json` (edit to match your fields)
2. Run:
   ```bash
   python quality_reviewer.py --source source.pdf --xml processed.xml --config sample_config.json --outdir qc_output --report-format all --html-summary
   ```
3. Outputs are generated in `qc_output/`.

## Configuration Schema (`sample_config.json`)
```json
{
  "fields": [
    {
      "name": "case_id",
      "type": "string",
      "source": {"method": "regex", "pattern": "Case\\s*ID[:\\s]*([A-Za-z0-9_-]+)", "flags": ["IGNORECASE", "MULTILINE"]},
      "xml": {"xpath": ".//Case/ID"},
      "normalize": ["strip"],
      "comparison": {"mode": "exact"}
    }
  ]
}
```

### Notes
- **PDF extraction** uses `PyPDF2`. For scanned PDFs, add OCR before using the tool.
- **DOCX extraction** uses `python-docx`.
- **CSV extraction**: set `source.method = "csv_column"` and provide `csv_column` name.
- XML evaluation supports simple paths like `.//Case/ID`. If your config uses `.../text()`, the tool will handle it.

## Extensions (Roadmap)
- Batch mode (folder of pairs) with consolidated dashboard
- Schema validation (XSD) for XML
- Field-level custom Python hooks
- UI (Streamlit) for point-and-click reviews
- Audit trail and e-sign ready PDFs

## Troubleshooting
- If a field doesn't extract from PDF/DOCX, test your regex on the plain text (export text first).
- If XML fields are missing, verify the XPath using a minimal XML sample.
- For fuzzy matches, adjust `fuzzy_threshold` (0.0–1.0).

