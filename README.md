# Document Inspector (`get_file_info.py`)

Batch-inspect office/PDF files and print human-readable summaries: basic properties, metadata, link annotations, and embedded comments — with neat ASCII tables.

Supported formats:
- **PDF**: basic info, metadata, link annotations  
- **DOCX**: basic info (pages/words if available), core properties  
- **PPTX**: basic info (slide count), core properties  
- **XLSX**: basic info (sheets, dimensions)  
- (Detection helper for legacy **.doc** files is included)

> The script relies on local helpers in `libs/`:
> - `libs/pdf.py`: `get_pdf_basic_info`, `extract_metadata`, `extract_link_annotations`
> - `libs/doc.py`: `get_docx_basic_info`
> - `libs/ppt.py`: `is_pptx_file`, `get_pptx_basic_info`
> - `libs/xlsx.py`: `is_xlsx_file`, `get_xlsx_basic_info`
> - `libs/shared.py`: `human_readable_size`

---

## Features

- 🔎 **Quick overview** for each file (type, size, counts)  
- 🧾 **Core metadata** (title, author, created/modified, etc.)  
- 🔗 **PDF link annotations** extraction  
- 💬 **Comments** listing (when available in the given format)  
- 🧱 **ASCII table** output that’s easy to scan or paste into tickets

---

## Installation

```bash
git clone <your-repo-url>
cd <your-repo-name>
python -m venv .venv
source .venv/bin/activate  # on Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

Project layout (expected):
```
.
├── get_file_info.py
└── libs/
    ├── pdf.py
    ├── doc.py
    ├── ppt.py
    ├── xlsx.py
    └── shared.py
```

---

## Usage

Basic invocation:
```bash
python get_file_info.py <path-to-file>
```

**Examples**
```bash
python get_file_info.py samples/report.pdf
python get_file_info.py samples/brief.docx
python get_file_info.py slides/talk.pptx
python get_file_info.py sheets/data.xlsx
```

---

## Example Output

```
File: report.pdf
Type: PDF
Size: 1.2 MB

┌──────────────┬───────────────┬───────────────┬──────────────┐
│ Pages        │ Title         │ Author        │ Modified     │
├──────────────┼───────────────┼───────────────┼──────────────┤
│ 12           │ Q3 Summary    │ Jane Doe      │ 2025-10-28   │
└──────────────┴───────────────┴───────────────┴──────────────┘
```

---

## Roadmap

- Recursive folder processing
- JSON output
- CSV export
- File hashing
- Redaction checks

---

## License

GNU 3.0
