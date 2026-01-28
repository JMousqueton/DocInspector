import sys
import os
from libs.pdf import get_pdf_basic_info, extract_metadata, extract_link_annotations
from libs.doc import get_docx_basic_info
from libs.ppt import is_pptx_file, get_pptx_basic_info
from libs.xlsx import is_xlsx_file, get_xlsx_basic_info
from libs.shared import human_readable_size

def print_ascii_table(array_table, headers):
    cols = len(headers)
    col_widths = [len(str(header)) for header in headers]
    for row in array_table:
        for idx, cell in enumerate(row):
            col_widths[idx] = max(col_widths[idx], len(str(cell)))
    sep_line = "┌" + "┬".join("─"*(w+2) for w in col_widths) + "┐"
    mid_line = "├" + "┼".join("─"*(w+2) for w in col_widths) + "┤"
    bot_line = "└" + "┴".join("─"*(w+2) for w in col_widths) + "┘"
    print(sep_line)
    header_line = "│ " + " │ ".join(headers[i].ljust(col_widths[i]) for i in range(cols)) + " │"
    print(header_line)
    print(mid_line)
    for row in array_table:
        print("│ " + " │ ".join(str(row[i]).ljust(col_widths[i]) for i in range(cols)) + " │")
    print(bot_line)

def is_pdf_file(filename):
    if not os.path.isfile(filename):
        return False
    try:
        with open(filename, "rb") as f:
            sig = f.read(5)
        return sig == b"%PDF-"
    except Exception:
        return False

def is_docx_file(filename):
    if not filename.lower().endswith('.docx'):
        return False
    if not os.path.isfile(filename):
        return False
    try:
        with open(filename, "rb") as f:
            sig = f.read(2)
        return sig == b'PK'
    except Exception:
        return False

def is_doc_file(filename):
    if not filename.lower().endswith('.doc'):
        return False
    if not os.path.isfile(filename):
        return False
    try:
        with open(filename, "rb") as f:
            sig = f.read(8)
        return sig == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
    except Exception:
        return False

if __name__ == "__main__":
    print('░███████                            ░██████                                                          ░██                        ')
    print('░██   ░██                             ░██                                                            ░██                        ')
    print('░██    ░██  ░███████   ░███████       ░██  ░████████   ░███████  ░████████   ░███████   ░███████  ░████████  ░███████  ░██░████ ')
    print('░██    ░██ ░██    ░██ ░██    ░██      ░██  ░██    ░██ ░██        ░██    ░██ ░██    ░██ ░██    ░██    ░██    ░██    ░██ ░███     ')
    print('░██    ░██ ░██    ░██ ░██             ░██  ░██    ░██  ░███████  ░██    ░██ ░█████████ ░██           ░██    ░██    ░██ ░██      ')
    print('░██   ░██  ░██    ░██ ░██    ░██      ░██  ░██    ░██        ░██ ░███   ░██ ░██        ░██    ░██    ░██    ░██    ░██ ░██      ')
    print('░███████    ░███████   ░███████     ░██████░██    ░██  ░███████  ░██░█████   ░███████   ░███████      ░████  ░███████  ░██      ')
    print('v1.1                                                             ░██                                                            ')
    print('                                                                 ░██                                                            ')
    print('                                                                                                                              ')
    if len(sys.argv) < 2:
        print("Usage: python extract.py <file>")
        sys.exit(1)
    filename = sys.argv[1]

    # Detect file type by signature
    if is_pdf_file(filename):
        filetype = "pdf"
    elif is_docx_file(filename):
        filetype = "docx"
    elif is_doc_file(filename):
        filetype = "doc"
    elif is_pptx_file(filename):
        filetype = "pptx"
    elif is_xlsx_file(filename):
        filetype = "xlsx"
    else:
        print("Not a supported file type (PDF, Word, PPTX, XLSX)")
        sys.exit(1)

    array_table = []

    if filetype == "pdf":
        info = get_pdf_basic_info(filename)
        array_table.append(["file_size_bytes", info["file_size_bytes"]])
        array_table.append(["file_size_human", human_readable_size(info["file_size_bytes"])])
        array_table.append(["pdf_version", info["pdf_version"]])
        array_table.append(["is_encrypted", info["is_encrypted"]])
        array_table.append(["num_pages", info["num_pages"]])
        array_table.append(["page_size", info["page_size"]])
        array_table += extract_metadata(filename)
        print_ascii_table(array_table, ["Property", "Value"])

        # Robust canarytoken/URL detection (raw scan)
        from libs.pdf import extract_urls_from_pdf_raw, detect_canarytokens
        all_urls = extract_urls_from_pdf_raw(filename)
        print("\nAll URLs found in PDF:")
        if all_urls:
            for url in all_urls:
                if "purl.org" in url.lower():
                    continue
                green_flag = ""
                if "microsoft.com" in url.lower():
                    green_flag += " [\033[92mMICROSOFT\033[0m]"
                if "adobe.com" in url.lower():
                    green_flag += " [\033[92mADOBE\033[0m]"
                if "w3.org" in url.lower():
                    green_flag += " [\033[92mW3 Org\033[0m]"
                if "canary" in url.lower():
                    green_flag += " [\033[91mCANARY\033[0m]"
                print(f"  - {url}{green_flag}")
        else:
            print("  (none found)")
        canarytokens = detect_canarytokens(all_urls)
        if canarytokens:
            print("\n\033[91mWARNING: Canarytoken(s) detected in PDF!\033[0m")
            for token in canarytokens:
                print(f"  Suspicious URL: {token}")
        else:
            print("\nNo canarytoken URLs detected in PDF.")

    elif filetype == "docx":
        info = get_docx_basic_info(filename)
        array_table.append(["file_size_bytes", info["file_size_bytes"]])
        array_table.append(["file_size_human", info["file_size_human"]])
        array_table.append(["num_pages", info.get("num_pages")])
        array_table.append(["has_revision_marks", info.get("has_revision_marks")])
        array_table.append(["num_paragraphs", info["num_paragraphs"]])
        array_table.append(["num_tables", info["num_tables"]])
        array_table += info["meta"]
        print_ascii_table(array_table, ["Property", "Value"])

        print("\nEmbedded URLs:")
        urls = info.get("links", [])
        if urls:
            for url in urls:
                print("  -", url)
        else:
            print("  (none found)")

        print("\nEmbedded Images:")
        images = info.get("images", [])
        if images:
            for img in images:
                print("  -", img)
        else:
            print("  (none found)")

        print("\nComments:")
        comments = info.get("comments", [])
        if comments:
            for c in comments:
                author = c.get("author", "")
                date = c.get("date", "")
                text = c.get("text", "")
                print(f"  - {author} ({date}): {text}")
        else:
            print("  (none found)")

    elif filetype == "doc":
        info = get_doc_basic_info(filename)
        array_table.append(["file_size_bytes", info["file_size_bytes"]])
        array_table.append(["file_size_human", info["file_size_human"]])
        array_table.append(["num_paragraphs", info["num_paragraphs"]])
        array_table.append(["num_tables", info["num_tables"]])
        array_table.append(["has_macros", info["has_vba_macros"]])
        array_table.append(["num_custom_xml", len(info["custom_xml_parts"])])
        array_table += info["meta"]
        print_ascii_table(array_table, ["Property", "Value"])
        print("\nEmbedded URLs:")
        urls = info.get("links", [])
        if urls:
            for url in urls:
                print("  -", url)
        else:
            print("  (none found)")
        print("\nEmbedded Images:")
        images = info.get("images", [])
        if images:
            for img in images:
                print("  -", img)
        else:
            print("  (none found)")
        print("\nComments:")
        comments = info.get("comments", [])
        if comments:
            for c in comments:
                author = c.get("author", "")
                date = c.get("date", "")
                text = c.get("text", "")
                print(f"  - {author} ({date}): {text}")
        else:
            print("  (none found)")

    elif filetype == "pptx":
        info = get_pptx_basic_info(filename)
        array_table.append(["file_size_bytes", info["file_size_bytes"]])
        array_table.append(["file_size_human", info["file_size_human"]])
        array_table.append(["num_slides", info["num_slides"]])
        array_table.append(["num_slides_with_notes", info["num_slides_with_notes"]])
        array_table.append(["has_macros", info["has_vba_macros"]])
        array_table.append(["num_custom_xml", len(info["custom_xml_parts"])])
        array_table += info["meta"]
        print_ascii_table(array_table, ["Property", "Value"])

        print("\nEmbedded URLs:")
        urls = info.get("links", [])
        if urls:
            for url in urls:
                print("  -", url)
        else:
            print("  (none found)")

        print("\nEmbedded Images:")
        images = info.get("images", [])
        if images:
            for img in images:
                print("  -", img)
        else:
            print("  (none found)")

        print("\nComments:")
        comments = info.get("comments", [])
        if comments:
            for c in comments:
                author = c.get("author", "")
                date = c.get("date", "")
                text = c.get("text", "")
                print(f"  - {author} ({date}): {text}")
        else:
            print("  (none found)")

    elif filetype == "xlsx":
        info = get_xlsx_basic_info(filename)
        array_table.append(["file_size_bytes", info["file_size_bytes"]])
        array_table.append(["file_size_human", info["file_size_human"]])
        array_table.append(["sheet_count", info.get("sheet_count", 0)])
        array_table += info.get("meta", [])
        print_ascii_table(array_table, ["Property", "Value"])

        print("\nSheets:")
        sheet_names = info.get("sheet_names", [])
        if sheet_names:
            for nm in sheet_names:
                print("  -", nm)
        else:
            print("  (none found)")

        print("\nEmbedded URLs:")
        urls = info.get("links", [])
        if urls:
            for url in urls:
                print("  -", url)
        else:
            print("  (none found)")

        print("\nEmbedded Images:")
        images = info.get("images", [])
        if images:
            for img in images:
                print("  -", img)
        else:
            print("  (none found)")

        print("\nComments:")
        comments = info.get("comments", [])
        if comments:
            for c in comments:
                author = c.get("author", "")
                loc = c.get("location", "")
                text = c.get("text", "")
                print(f"  - {author} @ {loc}: {text}")
        else:
            print("  (none found)")


def handle_xlsx(filepath):
    info = get_xlsx_basic_info(filepath)
    print("=== XLSX Report ===")
    print(f"File: {filepath}")
    print(f"Size: {info.get('file_size_human')} ({info.get('file_size_bytes')} bytes)")

    meta = info.get("meta", {})
    core = meta.get("core", {})
    app = meta.get("app", {})
    print("\nCore Properties:")
    for k, v in core.items():
        if v:
            print(f"  - {k}: {v}")
    print("\nApp Properties:")
    for k, v in app.items():
        if v:
            print(f"  - {k}: {v}")

    print(f"\nSheets ({info.get('sheet_count', 0)}):")
    for nm in info.get("sheet_names", []):
        print(f"  - {nm}")

    print("\nLinks:")
    links = info.get("links", [])
    if links:
        for l in links:
            print(f"  - {l}")
    else:
        print("  (none found)")

    print("\nImages:")
    imgs = info.get("images", [])
    if imgs:
        for i in imgs:
            print(f"  - {i}")
    else:
        print("  (none found)")

    print("\nComments:")
    comments = info.get("comments", [])
    if comments:
        for c in comments:
            author = c.get("author", "")
            loc = c.get("location", "")
            text = c.get("text", "")
            print(f"  - {author} @ {loc}: {text}")
    else:
        print("  (none found)")



