#!/usr/bin/env python3
import zipfile, xml.etree.ElementTree as ET, re, sys
from pathlib import Path

# --- Patterns for titles without chapters ---
RE_TITLE_LINE = re.compile(r'^TITLE\s+(\d+)\.\s*—\s*(.+?)\s*$', re.IGNORECASE)
# Match section number, title (up to em dash), and body text after em dash
RE_SECTION_INLINE = re.compile(r'^(\d+)\.\s*(.+?)\s*—\s*(.+)$')
RE_SECTION_HEADER = re.compile(r'^(\d+)\.\s')

def iter_docx_paragraph_text(docx_path: Path):
    """Yield plain text for each paragraph (<w:p>) in order."""
    with zipfile.ZipFile(docx_path) as z:
        with z.open("word/document.xml") as f:
            tree = ET.parse(f)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    for p in tree.findall(".//w:p", ns):
        texts = [t.text for t in p.findall(".//w:t", ns) if t.text]
        if texts:
            yield "".join(texts).replace("\xa0", " ").strip()

def normalize(s: str) -> str:
    """Normalize en dash → em dash; collapse whitespace runs to single spaces."""
    s = s.replace("–", "—")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def format_section_body(text: str) -> str:
    """Add spacing after subsection markers (a), (b), (1), (2), etc."""
    # Add space after patterns like "(a)", "(b)", "(1)", "(2)" if not already present
    text = re.sub(r'\(([a-z0-9]+)\)(?=[A-Z])', r'(\1) ', text)
    # Ensure single space after these markers
    text = re.sub(r'\(([a-z0-9]+)\)\s+', r'(\1) ', text)
    return text

def build_title_name(docx_path: Path, paragraphs):
    """Extract title number and display name from paragraphs."""
    title_display = None
    title_num = None
    for p in paragraphs:
        m = RE_TITLE_LINE.match(p)
        if m:
            title_num = m.group(1)
            title_text = m.group(2)
            # Clean up any TOC headers that may have bled in
            for cut_word in ["Chapter", "Sec."]:
                cut = title_text.find(cut_word)
                if cut != -1:
                    title_text = title_text[:cut].rstrip(" .")
            title_display = f"TITLE {title_num}.—{title_text.strip()}"
            break
    if not title_display:
        # Fallback to filename
        m = re.search(r'Title\s+(\d+)', docx_path.stem, re.IGNORECASE)
        title_num = m.group(1) if m else "?"
        title_display = f"TITLE {title_num}.—"
    return title_num, title_display

def to_xml_tree(docx_path: Path):
    """Convert DOCX to XML tree with Title -> Section structure (no chapters)."""
    paras = [normalize(p) for p in iter_docx_paragraph_text(docx_path) if p.strip()]

    title_num, title_display = build_title_name(docx_path, paras)

    root = ET.Element("USCode")
    title_el = ET.SubElement(root, "Title")
    title_el.set("name", title_display)

    # Skip introductory text until we find the first section
    i = 0
    while i < len(paras):
        line = paras[i]
        # Look for first section (starts with digit followed by period)
        if RE_SECTION_HEADER.match(line):
            break
        i += 1

    def add_section(num, title, body):
        """Add a Section element to the Title."""
        s = ET.SubElement(title_el, "Section")
        # Escape quotes in title and body
        title = title.strip()
        body = format_section_body(body.strip())
        s.set("name", f"{num}. {title}")
        s.text = body

    # Parse sections
    while i < len(paras):
        line = paras[i]

        # Try inline section format: "1. Section Title — Body text"
        sm = RE_SECTION_INLINE.match(line)
        if sm:
            num, sec_title, first_body = sm.groups()
            body_lines = [first_body] if first_body else []
            j = i + 1
            # Gather subsequent paragraphs until we hit another section
            while j < len(paras):
                nxt = paras[j]
                if RE_SECTION_HEADER.match(nxt):
                    break
                body_lines.append(nxt)
                j += 1
            add_section(num, sec_title, " ".join(body_lines))
            i = j
            continue

        i += 1

    # Pretty-print indentation
    def indent(elem, level=0):
        i = "\n" + level * "    "
        if len(elem):
            if not elem.text or not elem.text.strip():
                elem.text = i + "    "
            for e in elem:
                indent(e, level + 1)
            if not e.tail or not e.tail.strip():
                e.tail = i
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i

    indent(root)
    return ET.ElementTree(root), title_num

def main():
    here = Path(".")
    args = sys.argv[1:]
    target: Path | None = None

    # Allow passing either the exact filename or just the title number
    if args:
        a0 = args[0]
        p = Path(a0)
        if p.exists() and p.suffix.lower() == ".docx":
            target = p
        elif a0.isdigit():
            matches = list(here.glob(f"Title {a0}*.docx"))
            if not matches:
                matches = list(here.glob(f"Title*{a0}*.docx"))
            if not matches:
                print(f"No DOCX matching Title {a0}*.docx in {here.resolve()}")
                sys.exit(1)
            target = matches[0]
        else:
            print("Usage: python make_xml_from_docx_no_chapters.py [TitleNumber | DocxFilename]")
            sys.exit(1)
    else:
        matches = sorted(here.glob("Title *.docx"))
        if not matches:
            print("No 'Title *.docx' file found in this folder.")
            sys.exit(1)
        if len(matches) > 1:
            print("Multiple 'Title *.docx' files found. Specify one, e.g.:")
            print("  python make_xml_from_docx_no_chapters.py 4")
            print("  python make_xml_from_docx_no_chapters.py 'Title 4.docx'")
            sys.exit(1)
        target = matches[0]

    tree, num = to_xml_tree(target)
    out_name = f"Title_{num}.xml"
    tree.write(out_name, encoding="UTF-8", xml_declaration=True)
    print(f"✅ Created {out_name} (no chapters)")

if __name__ == "__main__":
    main()
