import re
from collections import Counter
from datetime import datetime
from PyPDF2 import PdfReader

PDF_META_FIELDS = [
    "/Title", "/Author", "/Subject", "/Keywords", "/Creator", "/Producer",
    "/CreationDate", "/ModDate", "/Trapped", "/Custom", "/Company", "/SourceModified",
    "/Category", "/ContentType", "/Language", "/Identifier", "/Format",
    "/LastModifiedBy", "/Revision", "/Description"
]

STANDARD_SIZES = {
    "A0": (2384, 3370),
    "A1": (1684, 2384),
    "A2": (1191, 1684),
    "A3": (842, 1191),
    "A4": (595, 842),
    "A5": (420, 595),
    "A6": (298, 420),
    "Letter": (612, 792),
    "Legal": (612, 1008),
    "Tabloid": (792, 1224)
}
TOLERANCE = 3

def match_standard_format(width, height):
    for name, (std_w, std_h) in STANDARD_SIZES.items():
        # Portrait
        if abs(width - std_w) <= TOLERANCE and abs(height - std_h) <= TOLERANCE:
            return name
        # Landscape
        if abs(width - std_h) <= TOLERANCE and abs(height - std_w) <= TOLERANCE:
            return name + " (landscape)"
    return ""

def parse_pdf_date(date_str):
    if not date_str:
        return ""
    date_str = date_str.strip()
    if date_str.startswith('D:'):
        date_str = date_str[2:]
    match = re.match(r"(\d{4})(\d{2})?(\d{2})?(\d{2})?(\d{2})?(\d{2})?", date_str)
    if not match:
        return date_str
    parts = match.groups(default="01")
    year = parts[0]
    month = parts[1] or "01"
    day = parts[2] or "01"
    hour = parts[3] or "00"
    minute = parts[4] or "00"
    second = parts[5] or "00"
    try:
        dt = datetime(int(year), int(month), int(day), int(hour), int(minute), int(second))
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return date_str

def extract_metadata(pdf_path):
    with open(pdf_path, "rb") as f:
        reader = PdfReader(f)
        metadata = reader.metadata
        meta = dict(metadata) if metadata else {}
        result = []
        for field in PDF_META_FIELDS:
            value = meta.get(field, "")
            if field.lower().endswith("date"):
                value = parse_pdf_date(value)
            result.append([field[1:], value])  # Remove leading /
        return result

def extract_link_annotations(pdf_path):
    urls = set()
    with open(pdf_path, "rb") as f:
        reader = PdfReader(f)
        for page_num, page in enumerate(reader.pages):
            if "/Annots" in page:
                for annot in page["/Annots"]:
                    obj = annot.get_object()
                    if obj.get("/Subtype") == "/Link":
                        if "/A" in obj and obj["/A"].get("/URI"):
                            url = obj["/A"]["/URI"]
                            urls.add(url)
    return sorted(urls)

def get_page_size_summary(pdf_file):
    with open(pdf_file, "rb") as f:
        reader = PdfReader(f)
        sizes = []
        for page in reader.pages:
            mediabox = page.mediabox
            width = round(float(mediabox.width))
            height = round(float(mediabox.height))
            sizes.append((width, height))
        count = Counter(sizes)
        page_size_strs = []
        for (w, h), v in count.items():
            fmt = match_standard_format(w, h)
            suffix = f" [{fmt}]" if fmt else ""
            count_str = f" ({v}x)" if v > 1 else ""
            page_size_strs.append(f"{w} x {h}{suffix}{count_str}")
        return ', '.join(page_size_strs)

def get_pdf_basic_info(pdf_file):
    # Returns dict: file_size_bytes, pdf_version, is_encrypted, num_pages, page_size
    from os.path import getsize
    file_size = getsize(pdf_file)
    with open(pdf_file, "rb") as f:
        reader = PdfReader(f)
        try:
            pdf_version = reader.pdf_header_version
        except AttributeError:
            pdf_version = "unknown"
        is_encrypted = reader.is_encrypted
        num_pages = len(reader.pages)
        page_size = get_page_size_summary(pdf_file)
    return {
        "file_size_bytes": file_size,
        "pdf_version": pdf_version,
        "is_encrypted": is_encrypted,
        "num_pages": num_pages,
        "page_size": page_size
    }

