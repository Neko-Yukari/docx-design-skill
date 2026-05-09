#!/usr/bin/env python3
"""
Edit .docx files with robust text replacement and structural edits.
Usage: python edit_docx.py <input.docx> <output.docx> [--replace OLD NEW]... [--insert-after N TEXT] [--delete N]

Examples:
  # Simple find-and-replace
  python edit_docx.py in.docx out.docx --replace "old" "new"

  # Multiple replacements
  python edit_docx.py in.docx out.docx --replace "张三" "李四" --replace "2024" "2025"

  # Insert text after paragraph 3
  python edit_docx.py in.docx out.docx --insert-after 3 "This is a new paragraph."

  # Delete paragraph 5
  python edit_docx.py in.docx out.docx --delete 5

  # Combined: fill template blanks
  python edit_docx.py template.docx report.docx --replace "姓名：" "姓名：张三" --replace "学号：" "学号：2024001"
"""
import sys
import os
from docx import Document
from docx.oxml.ns import qn
from copy import deepcopy


# ---------------------------------------------------------------------------
# Core: robust text replace (handles split runs)
# ---------------------------------------------------------------------------

def set_paragraph_text(para, new_text):
    """Replace entire paragraph text, preserving formatting of the first run."""
    if not para.runs:
        para.add_run(new_text)
        return
    # Preserve first run's formatting
    ref_run = para.runs[0]
    # Clear all existing runs
    for run in para.runs[1:]:
        run._element.getparent().remove(run._element)
    ref_run.text = new_text


def replace_in_paragraph(para, old, new):
    """Replace text in a paragraph, handling text split across multiple runs."""
    # Reconstruct full paragraph text
    full_text = para.text
    if old not in full_text:
        return False

    # If simple case (text in a single run), do fast path
    for run in para.runs:
        if old in run.text:
            run.text = run.text.replace(old, new)
            return True

    # Complex case: text spans multiple runs
    # Strategy: set all text into first run, clear others
    set_paragraph_text(para, full_text.replace(old, new))
    return True


def replace_in_table_cell(cell, old, new):
    """Replace text in a table cell."""
    changed = False
    for para in cell.paragraphs:
        if replace_in_paragraph(para, old, new):
            changed = True
    return changed


def replace_all(doc, old, new):
    """Replace text across entire document: paragraphs, tables, headers, footers."""
    count = 0

    # Body paragraphs
    for para in doc.paragraphs:
        if replace_in_paragraph(para, old, new):
            count += 1

    # Tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if replace_in_table_cell(cell, old, new):
                    count += 1

    # Headers & Footers
    for section in doc.sections:
        for hf in [section.header, section.footer,
                   section.even_page_header, section.even_page_footer,
                   section.first_page_header, section.first_page_footer]:
            if hf is None:
                continue
            for para in hf.paragraphs:
                if replace_in_paragraph(para, old, new):
                    count += 1
            for table in hf.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if replace_in_table_cell(cell, old, new):
                            count += 1

    return count


# ---------------------------------------------------------------------------
# Structural edits
# ---------------------------------------------------------------------------

def insert_paragraph_after(doc, after_index, text, style=None):
    """Insert a paragraph after the given paragraph index (0-based)."""
    if after_index < 0 or after_index >= len(doc.paragraphs):
        print(f"Error: paragraph index {after_index} out of range (0-{len(doc.paragraphs)-1})")
        return

    target = doc.paragraphs[after_index]
    new_para = doc.add_paragraph(text)
    if style:
        new_para.style = style
    target._element.addnext(new_para._element)
    print(f"Inserted paragraph after index {after_index}")


def delete_paragraph(doc, index):
    """Delete paragraph at given index (0-based)."""
    if index < 0 or index >= len(doc.paragraphs):
        print(f"Error: paragraph index {index} out of range (0-{len(doc.paragraphs)-1})")
        return

    para = doc.paragraphs[index]
    para._element.getparent().remove(para._element)
    print(f"Deleted paragraph {index}: '{para.text[:50]}...'")


def insert_text_in_paragraph(para, position, text):
    """Insert text at a specific character position within a paragraph."""
    full_text = para.text
    new_text = full_text[:position] + text + full_text[position:]
    set_paragraph_text(para, new_text)


# ---------------------------------------------------------------------------
# Paragraph listing (for finding indices)
# ---------------------------------------------------------------------------

def list_paragraphs(doc, max_show=None):
    """Print all paragraphs with indices, useful for finding positions to edit."""
    total = len(doc.paragraphs)
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if max_show and i >= max_show:
            break
        style = para.style.name if para.style else '-'
        preview = text[:80] + ('...' if len(text) > 80 else '')
        print(f"[{i:3d}] [{style:20s}] {preview}")
    if max_show and total > max_show:
        print(f"... and {total - max_show} more paragraphs")


# ---------------------------------------------------------------------------
# Track Changes (OOXML level)
# ---------------------------------------------------------------------------

def add_tracked_insertion(para, text, author="Editor"):
    """Add text as a tracked insertion at the end of a paragraph."""
    from datetime import datetime
    from lxml import etree

    date_str = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")

    # Create <w:ins> element
    nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    ins = etree.SubElement(para._element, qn('w:ins'))
    ins.set(qn('w:id'), str(os.urandom(2)[0] % 1000))
    ins.set(qn('w:author'), author)
    ins.set(qn('w:date'), date_str)

    # Add a run inside
    r = etree.SubElement(ins, qn('w:r'))
    rPr = etree.SubElement(r, qn('w:rPr'))
    t = etree.SubElement(r, qn('w:t'))
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text


def add_tracked_deletion(para, author="Editor"):
    """Mark entire paragraph as tracked deletion."""
    from datetime import datetime
    from lxml import etree

    date_str = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")

    # Mark paragraph mark as deleted
    pPr = para._element.find(qn('w:pPr'))
    if pPr is None:
        pPr = etree.SubElement(para._element, qn('w:pPr'))
        para._element.insert(0, pPr)
    rPr = pPr.find(qn('w:rPr'))
    if rPr is None:
        rPr = etree.SubElement(pPr, qn('w:rPr'))
    del_mark = etree.SubElement(rPr, qn('w:del'))
    del_mark.set(qn('w:id'), str(os.urandom(2)[0] % 1000))
    del_mark.set(qn('w:author'), author)
    del_mark.set(qn('w:date'), date_str)

    # Wrap all runs in <w:del>
    for run_elem in list(para._element.findall(qn('w:r'))):
        del_elem = etree.Element(qn('w:del'))
        del_elem.set(qn('w:id'), str(os.urandom(2)[0] % 1000))
        del_elem.set(qn('w:author'), author)
        del_elem.set(qn('w:date'), date_str)

        para._element.replace(run_elem, del_elem)
        # Change <w:t> to <w:delText>
        for t_elem in run_elem.findall(qn('w:t')):
            t_elem.tag = qn('w:delText')
        del_elem.append(run_elem)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]
    args = sys.argv[3:]

    if not os.path.exists(input_path):
        print(f"Error: file not found: {input_path}")
        sys.exit(1)

    doc = Document(input_path)

    # Parse commands
    i = 0
    while i < len(args):
        cmd = args[i]

        if cmd == '--replace' and i + 2 < len(args):
            old = args[i + 1]
            new = args[i + 2]
            count = replace_all(doc, old, new)
            print(f"Replace '{old}' -> '{new}': {count} occurrences")
            i += 3

        elif cmd == '--insert-after' and i + 2 < len(args):
            idx = int(args[i + 1])
            text = args[i + 2]
            insert_paragraph_after(doc, idx, text)
            i += 3

        elif cmd == '--delete' and i + 1 < len(args):
            idx = int(args[i + 1])
            delete_paragraph(doc, idx)
            i += 2

        elif cmd == '--list':
            max_show = int(args[i + 1]) if i + 1 < len(args) and args[i + 1].isdigit() else None
            list_paragraphs(doc, max_show)
            i += 2 if max_show else 1

        elif cmd == '--fill':
            # Read replacements from a json file: {"old1": "new1", "old2": "new2"}
            import json
            if i + 1 >= len(args):
                print("Error: --fill requires a JSON file path")
                sys.exit(1)
            json_path = args[i + 1]
            with open(json_path, 'r', encoding='utf-8') as f:
                replacements = json.load(f)
            for old, new in replacements.items():
                count = replace_all(doc, old, new)
                print(f"Replace '{old}' -> '{new}': {count} occurrences")
            i += 2

        else:
            print(f"Unknown command: {cmd}")
            i += 1

    doc.save(output_path)
    print(f"\nSaved: {output_path}")


if __name__ == '__main__':
    main()
