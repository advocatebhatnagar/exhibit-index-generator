# Exhibit Index Generator

CLI tool for litigators — builds court-ready exhibit indexes from folders of PDFs in seconds.

## What it does

- Scans a directory of exhibit PDFs (recursive option)
- Extracts metadata: page count, file size, Bates numbers (from filename or PDF text layer)
- Auto-generates exhibit descriptions from filenames
- Outputs **Markdown**, **CSV**, or **formatted DOCX** (table with proper styling)
- Court-ready format with case metadata (name, number, court)

## Install

```bash
pip install -r requirements.txt
# or
pip install pymupdf python-docx
```

## Usage

```bash
# Basic — Markdown output
python exhibit_index.py /path/to/exhibits -o index.md

# With case metadata, CSV output
python exhibit_index.py /path/to/exhibits -o index.csv \
  --case "State v. Smith" --case-no "CR-2026-00123" --court "High Court of Delhi"

# DOCX output (formatted table)
python exhibit_index.py /path/to/exhibits -o index.docx \
  --case "ABC Corp v. XYZ Ltd" --court "Delhi High Court"

# Recursive scan (subfolders)
python exhibit_index.py /path/to/exhibits --recursive -o index.md

# Skip Bates extraction (faster)
python exhibit_index.py /path/to/exhibits -o index.md --no-bates
```

## Bates Detection

Detects Bates numbers from:
- **Filenames**: `EXH-00001_Contract.pdf`, `Exhibit 12-DOC_0045.pdf`, `PLAINTIFF_EXH_000123.pdf`
- **PDF text layer** (page 1): scans for patterns like `BATES: ABC-12345`, `Exhibit #: DOC_000456`

Supported patterns: `PREFIX-NNNNNN`, `PREFIX_NNNNNN`, `EXH-NNNNN`, `DOC_NNNNNN`, etc.

## Output Formats

| Format | Best for |
|--------|----------|
| `.md` / `.markdown` | Quick review, GitHub, Obsidian, copy-paste |
| `.csv` | Excel, case management software import |
| `.docx` | Court filing, printing, formal submission |

## Example Output (Markdown)

```markdown
# Exhibit Index

**Case:** State v. Smith
**Case No.:** CR-2026-00123
**Court:** High Court of Delhi
**Generated:** 2026-08-10 14:32
**Total Exhibits:** 12
**Total Pages:** 247

| Ex. # | Bates / ID | Description | Pages | Size |
|-------|------------|-------------|-------|------|
| 1 | EXH-00001 | Contract Agreement | 12 | 456.2 KB |
| 2 | EXH-00002 | Email Correspondence | 8 | 234.1 KB |
| 3 | — | Witness Statement | 5 | 112.3 KB |
```

## Requirements

- Python 3.8+
- `pymupdf` (fitz) — PDF parsing
- `python-docx` — DOCX output (optional, only for `.docx`)

## License

MIT — free for any use, including commercial.