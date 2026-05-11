#!/usr/bin/env python3
"""
Convert .docx to .doc (legacy Word binary format).

Usage:
    python convert_docx_to_doc.py input.docx [output.doc]

If output is omitted, writes to <input>.doc

Windows: Uses MS Word COM (requires pywin32)
Cross-platform: Uses LibreOffice (requires soffice in PATH)
"""
import sys
import os
import subprocess
import platform


def convert_win32(docx_path, doc_path):
    """Windows: Convert via Microsoft Word COM."""
    import win32com.client

    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False

    # 0 = wdFormatDocument (binary .doc)
    doc = word.Documents.Open(os.path.abspath(docx_path))
    doc.SaveAs2(os.path.abspath(doc_path), FileFormat=0)
    doc.Close()
    word.Quit()
    print(f'Converted (MS Word): {docx_path} -> {doc_path}')


def convert_libreoffice(docx_path, doc_path):
    """Cross-platform: Convert via LibreOffice."""
    out_dir = os.path.dirname(doc_path) or '.'
    cmd = [
        'soffice',
        '--headless',
        '--convert-to', 'doc',
        '--outdir', out_dir,
        docx_path
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)

    if result.returncode != 0:
        print(f'Error: LibreOffice conversion failed')
        print(result.stderr)
        sys.exit(1)

    # LibreOffice names output as <basename>.doc in out_dir
    libre_out = os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + '.doc')
    if libre_out != doc_path and os.path.exists(libre_out):
        os.replace(libre_out, doc_path)
    print(f'Converted (LibreOffice): {docx_path} -> {doc_path}')


def convert(docx_path, doc_path):
    """Auto-detect conversion method based on platform and available tools."""
    if platform.system() == 'Windows':
        try:
            convert_win32(docx_path, doc_path)
            return
        except ImportError:
            print('pywin32 not installed, trying LibreOffice...')
        except Exception as e:
            print(f'MS Word COM failed ({e}), trying LibreOffice...')

    # Fall back to LibreOffice
    try:
        convert_libreoffice(docx_path, doc_path)
    except FileNotFoundError:
        print('Error: LibreOffice (soffice) not found in PATH.')
        print('Install LibreOffice: https://www.libreoffice.org/download/')
        print('Or on Windows with MS Word: pip install pywin32')
        sys.exit(1)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    input_path = sys.argv[1]
    if not os.path.exists(input_path):
        print(f'Error: file not found: {input_path}')
        sys.exit(1)

    if len(sys.argv) > 2:
        output_path = sys.argv[2]
    else:
        output_path = os.path.splitext(input_path)[0] + '.doc'

    convert(input_path, output_path)
