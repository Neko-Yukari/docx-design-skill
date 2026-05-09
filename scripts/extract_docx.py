#!/usr/bin/env python3
"""
Extract text from .docx files into structured Markdown.
Usage: python extract_docx.py <input.docx> [output.md]

If output is omitted, writes to <input>_extracted.md
"""
import sys
import os
from docx import Document


def extract_docx_to_markdown(docx_path, output_path):
    doc = Document(docx_path)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(f"# Source: {os.path.basename(docx_path)}\n\n")

        for para in doc.paragraphs:
            style = para.style.name if para.style else 'Normal'
            text = para.text.strip()
            if not text:
                f.write('\n')
                continue

            if style.startswith('Heading'):
                try:
                    level = int(style.replace('Heading ', ''))
                    f.write(f"{'#' * level} {text}\n\n")
                except ValueError:
                    f.write(f"# {text}\n\n")
            else:
                f.write(f"{text}\n\n")

        # Tables
        if doc.tables:
            f.write(f"\n---\n\n")
        for i, table in enumerate(doc.tables):
            f.write(f"\n**Table {i + 1}:**\n\n")
            for row in table.rows:
                cells = [cell.text.replace('\n', ' ').strip() for cell in row.cells]
                f.write('| ' + ' | '.join(cells) + ' |\n')
            f.write('\n')

    print(f"Extracted: {docx_path} -> {output_path}")
    print(f"  Paragraphs: {len(doc.paragraphs)}")
    print(f"  Tables: {len(doc.tables)}")


def main():
    if len(sys.argv) < 2:
        print("Usage: python extract_docx.py <input.docx> [output.md]")
        sys.exit(1)

    input_path = sys.argv[1]

    if len(sys.argv) > 2:
        output_path = sys.argv[2]
    else:
        base = os.path.splitext(input_path)[0]
        output_path = f"{base}_extracted.md"

    extract_docx_to_markdown(input_path, output_path)


if __name__ == '__main__':
    main()
