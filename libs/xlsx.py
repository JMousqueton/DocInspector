
import os
from zipfile import ZipFile
import xml.etree.ElementTree as ET

def is_xlsx_file(filename):
    return filename.lower().endswith(".xlsx") and _has_zip_sig(filename)

def _has_zip_sig(filename):
    try:
        with open(filename, "rb") as f:
            return f.read(2) == b"PK"
    except Exception:
        return False

def _read_core_properties(xlsx_path):
    core = {}
    try:
        with ZipFile(xlsx_path) as zf:
            if "docProps/core.xml" in zf.namelist():
                xml = zf.read("docProps/core.xml")
                root = ET.fromstring(xml)
                ns = {
                    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
                    "dc": "http://purl.org/dc/elements/1.1/",
                    "dcterms": "http://purl.org/dc/terms/",
                }
                def g(tag):
                    el = root.find(tag, ns)
                    return el.text if el is not None else ""
                core = {
                    "title": g("dc:title"),
                    "subject": g("dc:subject"),
                    "creator": g("dc:creator"),
                    "description": g("dc:description"),
                    "keywords": g("cp:keywords"),
                    "lastModifiedBy": g("cp:lastModifiedBy"),
                    "revision": g("cp:revision"),
                    "created": g("dcterms:created"),
                    "modified": g("dcterms:modified"),
                }
    except Exception:
        pass
    return core

def _read_app_properties(xlsx_path):
    app = {}
    try:
        with ZipFile(xlsx_path) as zf:
            if "docProps/app.xml" in zf.namelist():
                xml = zf.read("docProps/app.xml")
                root = ET.fromstring(xml)
                ns = {"ap":"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"}
                def g(tag):
                    el = root.find(tag, ns)
                    return el.text if el is not None else ""
                app = {
                    "application": g("ap:Application"),
                    "appVersion": g("ap:AppVersion"),
                    "docSecurity": g("ap:DocSecurity"),
                    "company": g("ap:Company"),
                    "manager": g("ap:Manager"),
                }
    except Exception:
        pass
    return app

def _sheet_names_and_hyperlinks(xlsx_path):
    """
    Returns (sheet_names, hyperlinks) where:
      - sheet_names is a list of sheet names.
      - hyperlinks is a set of hyperlink targets found across sheets.
    Uses minimal parsing of xl/workbook.xml and xl/worksheets/*.xml
    """
    sheet_names = []
    hyperlinks = set()
    try:
        with ZipFile(xlsx_path) as zf:
            # Sheet names from xl/workbook.xml
            if "xl/workbook.xml" in zf.namelist():
                wb_xml = zf.read("xl/workbook.xml")
                root = ET.fromstring(wb_xml)
                ns = {"r":"http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
                for sh in root.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet"):
                    nm = sh.get("name", "")
                    if nm:
                        sheet_names.append(nm)

            # Hyperlinks (look across all worksheets)
            for name in zf.namelist():
                if name.startswith("xl/worksheets/") and name.endswith(".xml"):
                    xml = zf.read(name)
                    try:
                        r = ET.fromstring(xml)
                        for h in r.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}hyperlink"):
                            tgt = h.get("display") or h.get("ref") or h.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                            # relationships may hold external targets, but extracting those requires following rels;
                            # we'll also capture explicit 'location' or 'tooltip' if present
                            href = h.get("location") or h.get("tooltip") or tgt
                            if href:
                                hyperlinks.add(href)
                    except Exception:
                        continue
    except Exception:
        pass
    return sheet_names, sorted(hyperlinks)

def _images_list(xlsx_path):
    imgs = []
    try:
        with ZipFile(xlsx_path) as zf:
            for name in zf.namelist():
                if name.startswith("xl/media/"):
                    imgs.append(name)
    except Exception:
        pass
    return imgs

def _comments(xlsx_path):
    comments = []
    try:
        with ZipFile(xlsx_path) as zf:
            # Comments can be in xl/comments*.xml (legacy) or threadedComments
            for name in zf.namelist():
                if name.startswith("xl/comments") and name.endswith(".xml"):
                    xml = zf.read(name)
                    root = ET.fromstring(xml)
                    # legacy comments: commentList/comment with attributes authorId, ref
                    for cm in root.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}comment"):
                        text_parts = []
                        for t in cm.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"):
                            if t.text:
                                text_parts.append(t.text)
                        text = "".join(text_parts).strip()
                        author = cm.get("authorId", "")
                        ref = cm.get("ref", "")
                        comments.append({"author": author, "location": ref, "text": text})

                if name.startswith("xl/threadedComments") and name.endswith(".xml"):
                    xml = zf.read(name)
                    root = ET.fromstring(xml)
                    for tc in root.findall(".//{http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments}threadedComment"):
                        text = (tc.get("text") or "").strip()
                        author = tc.get("personId", "")
                        ref = tc.get("ref", "")
                        comments.append({"author": author, "location": ref, "text": text})
    except Exception:
        pass
    return comments

def has_vba_macros(xlsx_path):
    try:
        with ZipFile(xlsx_path) as zf:
            return "xl/vbaProject.bin" in zf.namelist()
    except Exception:
        return False


from libs.shared import human_readable_size

def get_xlsx_basic_info(xlsx_file):
    file_size = os.path.getsize(xlsx_file)
    core = _read_core_properties(xlsx_file)
    app = _read_app_properties(xlsx_file)
    sheet_names, links = _sheet_names_and_hyperlinks(xlsx_file)
    images = _images_list(xlsx_file)
    comments = _comments(xlsx_file)
    has_macros = has_vba_macros(xlsx_file)

    meta_pairs = []
    for k, v in core.items():
        meta_pairs.append([f"core_{k}", v])
    for k, v in app.items():
        meta_pairs.append([f"app_{k}", v])

    info = {
        "file_size_bytes": file_size,
        "file_size_human": human_readable_size(file_size),
        "meta": meta_pairs,
        "sheet_count": len(sheet_names),
        "sheet_names": sheet_names,
        "links": links,
        "images": images,
        "comments": comments,
        "has_vba_macros": has_macros,
    }
    return info

