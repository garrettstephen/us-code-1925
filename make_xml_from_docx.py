#!/usr/bin/env python3
import zipfile, xml.etree.ElementTree as ET, re, sys
from pathlib import Path

# --- Patterns (robust to layout quirks) ---
RE_TITLE_LINE_MIN   = re.compile(r'^TITLE\s+(\d+)\.\s*—\s*(.+?)\s*$', re.IGNORECASE)
RE_CHAPTER_HEAD     = re.compile(r'^Chapter\s+(\d+)\.?\s*—\s*(.+?)\.?\s*$', re.IGNORECASE)  # dot optional
RE_SECTION_INLINE   = re.compile(r'^(\d+)\.\s*(.+?)\s*—\s*(.+)$')
RE_SECTION_HEADER   = re.compile(r'^(\d+)\.\s')  # to detect start of a new section block

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
    # Normalize en dash → em dash; collapse whitespace runs to single spaces.
    s = s.replace("–", "—")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def build_title_name(docx_path: Path, paragraphs):
    title_display = None
    title_num = None
    for p in paragraphs:
        m = RE_TITLE_LINE_MIN.match(p)
        if m:
            title_num = m.group(1)
            title_text = m.group(2)
            # If TOC headers like "Chapter Sec." bled into the same paragraph, cut them off.
            cut = title_text.find("Chapter")
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
    paras = [normalize(p) for p in iter_docx_paragraph_text(docx_path) if p.strip()]

    title_num, title_display = build_title_name(docx_path, paras)

    root = ET.Element("USCode")
    title_el = ET.SubElement(root, "Title")
    title_el.set("name", title_display)

    # Skip intros and any TOC-like row such as "Chapter   Sec."
    i = 0
    while i < len(paras):
        line = paras[i]
        # Skip TOC header lines that aren’t real chapter headers
        if line.lower().startswith("chapter") and "sec" in line.lower() and not RE_CHAPTER_HEAD.match(line):
            i += 1
            continue
        if RE_CHAPTER_HEAD.match(line):
            break
        i += 1

    current_chapter = None

    def start_chapter(num, name):
        nonlocal current_chapter
        current_chapter = ET.SubElement(title_el, "Chapter")
        current_chapter.set("name", f"Chapter {num}.—{name.strip().rstrip('.')}.")  # mirror sample format

    def add_section(num, title, body):
        if current_chapter is None:
            start_chapter("0", "UNSPECIFIED")
        s = ET.SubElement(current_chapter, "Section")
        s.set("name", f"{num}. {title.strip()}")
        s.text = body.strip()

    # Parse chapters + sections
    while i < len(paras):
        line = paras[i]

        chm = RE_CHAPTER_HEAD.match(line)
        if chm:
            start_chapter(chm.group(1), chm.group(2))
            i += 1
            continue

        sm = RE_SECTION_INLINE.match(line)
        if sm:
            add_section(sm.group(1), sm.group(2), sm.group(3))
            i += 1
            continue

        # Multi-paragraph section body starting at this line
        m = re.match(r'^(\d+)\.\s*(.+?)\s*—\s*(.*)$', line)
        if m:
            num, sec_title, first_body = m.groups()
            body_lines = [first_body] if first_body else []
            j = i + 1
            while j < len(paras):
                nxt = paras[j]
                if RE_CHAPTER_HEAD.match(nxt) or RE_SECTION_INLINE.match(nxt) or RE_SECTION_HEADER.match(nxt):
                    break
                body_lines.append(nxt)
                j += 1
            add_section(num, sec_title, "\n".join(body_lines))
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

    # Allow passing either the exact filename or just the title number.
    if args:
        a0 = args[0]
        p = Path(a0)
        if p.exists() and p.suffix.lower() == ".docx":
            target = p
        elif a0.isdigit():
            matches = list(here.glob(f"Title {a0}*.docx"))
            if not matches:
                print(f"No DOCX matching Title {a0}*.docx in {here.resolve()}")
                sys.exit(1)
            target = matches[0]
        else:
            print("Usage: python make_one_title_from_docx.py [TitleNumber | DocxFilename]")
            sys.exit(1)
    else:
        matches = sorted(here.glob("Title *.docx"))
        if not matches:
            print("No 'Title *.docx' file found in this folder.")
            sys.exit(1)
        if len(matches) > 1:
            print("Multiple 'Title *.docx' files found. Specify one, e.g.:")
            print("  python make_one_title_from_docx.py 3")
            print("  python make_one_title_from_docx.py 'Title 3 - The President.docx'")
            sys.exit(1)
        target = matches[0]

    tree, num = to_xml_tree(target)
    out_name = f"Title_{num}.xml"
    tree.write(out_name, encoding="utf-8", xml_declaration=True)
    print(f"✅ Created {out_name}")

if __name__ == "__main__":
    main()
