#!/usr/bin/env python3
"""
Merge adjacent runs with identical formatting in a .docx file.
This eliminates the #1 bug in docx editing: text split across runs.

Workflow:
  1. python merge_runs.py input.docx merged/
  2. Edit the XML files in merged/word/
  3. Repack: use edit_docx.py or manually zip

Usage:
  python merge_runs.py input.docx output_dir/
"""
import sys
import os
import shutil
import zipfile
from pathlib import Path
from lxml import etree

NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def unpack(docx_path, output_dir):
    """Unpack .docx ZIP into directory."""
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir)
    with zipfile.ZipFile(docx_path, 'r') as z:
        z.extractall(output_dir)
    print('Unpacked: ' + docx_path)


def merge_runs_in_dir(unpacked_dir):
    """Merge adjacent runs with identical formatting in document.xml."""
    doc_xml = Path(unpacked_dir) / 'word' / 'document.xml'

    if not doc_xml.exists():
        print("Error: word/document.xml not found in " + unpacked_dir)
        return 0

    tree = etree.parse(str(doc_xml))
    root = tree.getroot()

    nsmap = {'w': NS}

    # Remove proofErr elements (spell/grammar markers block merging)
    for elem in root.findall('.//w:proofErr', nsmap):
        elem.getparent().remove(elem)

    # Strip rsid attributes from runs
    for run in root.findall('.//w:r', nsmap):
        for attr_name in list(run.attrib.keys()):
            if 'rsid' in attr_name.lower():
                del run.attrib[attr_name]

    # Find all containers that have runs (paragraphs, ins, del)
    containers = set()
    for run in root.findall('.//w:r', nsmap):
        containers.add(run.getparent())

    total_merged = 0
    for container in containers:
        total_merged += merge_in_container(container, nsmap)

    # Write back
    tree.write(str(doc_xml), xml_declaration=True, encoding='UTF-8', pretty_print=True)
    return total_merged


def merge_in_container(container, nsmap):
    """Merge adjacent runs within a single container element."""
    merged = 0
    children = list(container)

    i = 0
    while i < len(children):
        current = children[i]
        if not _is_run(current):
            i += 1
            continue

        # Look ahead for adjacent mergeable runs
        j = i + 1
        while j < len(children):
            next_elem = children[j]
            if _is_run(next_elem) and _same_format(current, next_elem):
                _merge_runs(current, next_elem)
                container.remove(next_elem)
                children = list(container)  # refresh after removal
                merged += 1
                # don't increment j, check next again
            elif _is_run(next_elem):
                break  # run but different format
            elif _is_text_node(next_elem) and (next_elem.tail or '').strip():
                break  # non-whitespace text between runs
            else:
                j += 1  # skip whitespace-only text nodes

        # Consolidate text elements within the run
        _consolidate_text(current)
        i += 1

    return merged


def _merge_runs(target, source):
    """Move all content children from source into target (skip rPr)."""
    for child in list(source):
        tag = etree.QName(child).localname if hasattr(child, 'tag') else ''
        if tag != 'rPr':
            target.append(child)


def _consolidate_text(run):
    """Merge adjacent <w:t> elements within a single run."""
    t_elements = [c for c in run if _is_tag(c, 't')]
    for k in range(len(t_elements) - 1, 0, -1):
        curr = t_elements[k]
        prev = t_elements[k - 1]
        if _adjacent(prev, curr, run):
            prev_text = prev.text or ''
            curr_text = curr.text or ''
            prev.text = prev_text + curr_text
            if prev.text.startswith(' ') or prev.text.endswith(' '):
                prev.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            run.remove(curr)


def _adjacent(elem1, elem2, parent):
    """Check if elem1 and elem2 are adjacent (no element or non-whitespace text between)."""
    between = False
    for child in parent:
        if child is elem1:
            between = True
        elif child is elem2:
            return between and True
        elif between:
            if not _is_element(child):
                continue
            return False
        if between and _is_element(child):
            return False
    return False

def _is_run(elem):
    return hasattr(elem, 'tag') and etree.QName(elem).localname == 'r'

def _is_tag(elem, tag):
    return hasattr(elem, 'tag') and etree.QName(elem).localname == tag

def _is_text_node(elem):
    return not hasattr(elem, 'tag')

def _is_element(elem):
    return hasattr(elem, 'tag')

def _same_format(run1, run2):
    """Check if two runs have identical rPr."""
    rpr1 = _find_child(run1, 'rPr')
    rpr2 = _find_child(run2, 'rPr')
    if (rpr1 is None) != (rpr2 is None):
        return False
    if rpr1 is None:
        return True
    return etree.tostring(rpr1) == etree.tostring(rpr2)

def _find_child(parent, tag):
    for child in parent:
        if hasattr(child, 'tag') and etree.QName(child).localname == tag:
            return child
    return None


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    input_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else input_path.replace('.docx', '_unpacked')

    unpack(input_path, output_dir)
    count = merge_runs_in_dir(output_dir)
    print('Merged %d runs' % count)
    print('Ready for editing: %s/word/document.xml' % output_dir)
