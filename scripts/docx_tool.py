#!/usr/bin/env python3
"""
docx_tool.py — One-stop Word document builder for opencode agents.

CLI: python docx_tool.py create --output report.docx --title "标题" --add-heading "1. 概述" --add-paragraph "正文"
API: from docx_tool import DocumentBuilder; DocumentBuilder().preset_chinese().add_heading("标题").save("out.docx")
"""
import argparse
import sys
import os
import copy
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
OMML_NS = M


def _remove_theme_refs(style):
    """Remove eastAsiaTheme/asciiTheme/hAnsiTheme/cstheme from a style's rFonts."""
    rPr = style.element.find(f'{{{W}}}rPr')
    if rPr is None:
        return
    rFonts = rPr.find(f'{{{W}}}rFonts')
    if rFonts is None:
        return
    for attr in [f'{{{W}}}eastAsiaTheme', f'{{{W}}}asciiTheme', f'{{{W}}}hAnsiTheme', f'{{{W}}}cstheme']:
        if attr in rFonts.attrib:
            del rFonts.attrib[attr]


def _set_style_color(style, hex_color='000000'):
    """Set font color on a style (hex without #)."""
    rPr = style.element.find(f'{{{W}}}rPr')
    if rPr is None:
        rPr = etree.SubElement(style.element, f'{{{W}}}rPr')
    color = rPr.find(f'{{{W}}}color')
    if color is None:
        color = etree.SubElement(rPr, f'{{{W}}}color')
    color.set(f'{{{W}}}val', hex_color)


def _set_style_east_asian_font(style, ea_font, latin_font='Times New Roman'):
    """Set East Asian and Latin fonts on a style."""
    rPr = style.element.find(f'{{{W}}}rPr')
    if rPr is None:
        rPr = etree.SubElement(style.element, f'{{{W}}}rPr')
    rFonts = rPr.find(f'{{{W}}}rFonts')
    if rFonts is None:
        rFonts = etree.SubElement(rPr, f'{{{W}}}rFonts')
    rFonts.set(f'{{{W}}}eastAsia', ea_font)
    rFonts.set(f'{{{W}}}ascii', latin_font)
    rFonts.set(f'{{{W}}}hAnsi', latin_font)
    style.font.name = latin_font


class DocumentBuilder:
    """Fluent API for building .docx documents with Chinese typography presets."""

    def __init__(self, template_path=None):
        """Load template or create blank document."""
        self.doc = Document(template_path) if template_path else Document()

    # ── Presets ────────────────────────────────────────────────

    def preset_chinese(self, body_font='SimHei', heading_font='SimSun',
                       body_size=10.5, h1_size=18, h2_size=14, h3_size=12,
                       color='000000'):
        """
        Apply Chinese document defaults:
        - Body: SimHei (黑体), 10.5pt
        - Heading 1/2/3: SimSun (宋体), bold
        - All text: black
        - Remove theme font references
        """
        # Normal
        style = self.doc.styles['Normal']
        style.font.size = Pt(body_size)
        _set_style_east_asian_font(style, body_font)
        _set_style_color(style, color)
        _remove_theme_refs(style)

        # Headings
        for lvl, sz in [(1, h1_size), (2, h2_size), (3, h3_size)]:
            try:
                s = self.doc.styles[f'Heading {lvl}']
                s.font.size = Pt(sz)
                s.font.bold = True
                _set_style_east_asian_font(s, heading_font)
                _set_style_color(s, color)
                _remove_theme_refs(s)
            except KeyError:
                pass
        return self

    def preset_english(self, body_font='Times New Roman', heading_font='Times New Roman',
                       body_size=12, h1_size=18, h2_size=14, h3_size=12,
                       color='000000'):
        """Apply English document defaults."""
        style = self.doc.styles['Normal']
        style.font.size = Pt(body_size)
        style.font.name = body_font
        _set_style_color(style, color)
        _remove_theme_refs(style)

        for lvl, sz in [(1, h1_size), (2, h2_size), (3, h3_size)]:
            try:
                s = self.doc.styles[f'Heading {lvl}']
                s.font.size = Pt(sz)
                s.font.bold = True
                s.font.name = heading_font
                _set_style_color(s, color)
                _remove_theme_refs(s)
            except KeyError:
                pass
        return self

    # ── Content building ───────────────────────────────────────

    def add_heading(self, text, level=1, alignment=None, font=None, color=None):
        """Add a heading paragraph."""
        p = self.doc.add_heading(text, level=level)
        if alignment is not None:
            p.alignment = alignment
        if font or color:
            for run in p.runs:
                if font:
                    run.font.name = font
                if color:
                    run.font.color.rgb = RGBColor(int(color[:2], 16), int(color[2:4], 16), int(color[4:6], 16))
        return self

    def add_paragraph(self, text, style='Normal', alignment=None, font=None, color=None):
        """Add a body paragraph."""
        p = self.doc.add_paragraph(text, style=style)
        if alignment is not None:
            p.alignment = alignment
        if font or color:
            for run in p.runs:
                if font:
                    run.font.name = font
                if color:
                    run.font.color.rgb = RGBColor(int(color[:2], 16), int(color[2:4], 16), int(color[4:6], 16))
        return self

    def add_formula(self, latex_str, display=True):
        """
        Insert a LaTeX formula as native OMML.
        display=True  → block-level (w:p > m:oMathPara)
        display=False → inline (w:r > m:oMath)
        """
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from latex2omml import latex_to_omml

        p = self.doc.add_paragraph()
        if display:
            # Display: insert oMathPara directly into paragraph
            omml_xml = latex_to_omml(latex_str, inline=False)
            omp_elem = etree.fromstring(omml_xml)
            p._element.append(omp_elem)
        else:
            # Inline: insert oMath into a run
            omml_xml = latex_to_omml(latex_str, inline=True)
            run = p.add_run('')
            run._element.append(etree.fromstring(omml_xml))
        return self

    def add_table(self, rows, cols, data=None, style='Table Grid'):
        """Add a table. data: list of lists, each inner list is a row."""
        table = self.doc.add_table(rows=rows, cols=cols, style=style)
        if data:
            for i, row_data in enumerate(data):
                if i >= rows:
                    break
                row = table.rows[i]
                for j, cell_text in enumerate(row_data):
                    if j >= cols:
                        break
                    row.cells[j].text = str(cell_text)
        return self

    def add_page_break(self):
        """Add a page break."""
        self.doc.add_page_break()
        return self

    def add_image(self, image_path, width=None, height=None):
        """Add an image. width/height in inches (use docx.shared.Inches)."""
        kwargs = {}
        if width:
            kwargs['width'] = width
        if height:
            kwargs['height'] = height
        self.doc.add_picture(image_path, **kwargs)
        return self

    # ── Output ─────────────────────────────────────────────────

    def save(self, path):
        """Save the document."""
        self.doc.save(path)
        return self

    # ── Read-only accessors ────────────────────────────────────

    @property
    def paragraph_count(self):
        return len(self.doc.paragraphs)

    @property
    def table_count(self):
        return len(self.doc.tables)


# ═══════════════════════════════════════════════════════════════
# CLI Entry Point
# ═══════════════════════════════════════════════════════════════

def _cli():
    parser = argparse.ArgumentParser(description='Build .docx documents from command line.')
    sub = parser.add_subparsers(dest='cmd', required=True)

    # ── create ─────────────────────────────────────────────────
    p_create = sub.add_parser('create', help='Create a new document')
    p_create.add_argument('--output', '-o', required=True, help='Output .docx path')
    p_create.add_argument('--preset', choices=['chinese', 'english'], default='chinese',
                          help='Typography preset (default: chinese)')
    p_create.add_argument('--title', help='Document title (Heading 1)')
    p_create.add_argument('--add-heading', action='append', default=[],
                          help='Add heading (can repeat). Format: "LEVEL:TEXT" e.g. "1:Overview"')
    p_create.add_argument('--add-paragraph', action='append', default=[],
                          help='Add paragraph (can repeat)')
    p_create.add_argument('--add-formula', action='append', default=[],
                          help='Add LaTeX formula as display block (can repeat)')
    p_create.add_argument('--add-inline-formula', action='append', default=[],
                          help='Add inline LaTeX formula (can repeat)')
    p_create.add_argument('--body-font', default=None, help='Override body font')
    p_create.add_argument('--heading-font', default=None, help='Override heading font')
    p_create.add_argument('--color', default='000000', help='Font color hex (default: 000000)')

    # ── modify ─────────────────────────────────────────────────
    p_modify = sub.add_parser('modify', help='Modify an existing .docx')
    p_modify.add_argument('--input', '-i', required=True, help='Input .docx path')
    p_modify.add_argument('--output', '-o', required=True, help='Output .docx path')
    p_modify.add_argument('--add-heading', action='append', default=[],
                          help='Add heading (can repeat). Format: "LEVEL:TEXT"')
    p_modify.add_argument('--add-paragraph', action='append', default=[],
                          help='Add paragraph (can repeat)')
    p_modify.add_argument('--add-formula', action='append', default=[],
                          help='Add LaTeX formula as display block')

    args = parser.parse_args()

    if args.cmd == 'create':
        builder = DocumentBuilder()
        if args.preset == 'chinese':
            kw = {}
            if args.body_font:
                kw['body_font'] = args.body_font
            if args.heading_font:
                kw['heading_font'] = args.heading_font
            if args.color:
                kw['color'] = args.color
            builder.preset_chinese(**kw)
        else:
            builder.preset_english()

        if args.title:
            builder.add_heading(args.title, level=1)

        for h in args.add_heading:
            if ':' in h:
                lvl, text = h.split(':', 1)
                builder.add_heading(text, level=int(lvl))
            else:
                builder.add_heading(h, level=1)

        for p in args.add_paragraph:
            builder.add_paragraph(p)

        for f in args.add_formula:
            builder.add_formula(f, display=True)

        for f in args.add_inline_formula:
            builder.add_formula(f, display=False)

        builder.save(args.output)
        print(f"Created: {args.output}")

    elif args.cmd == 'modify':
        builder = DocumentBuilder(template_path=args.input)
        for h in args.add_heading:
            if ':' in h:
                lvl, text = h.split(':', 1)
                builder.add_heading(text, level=int(lvl))
            else:
                builder.add_heading(h, level=1)
        for p in args.add_paragraph:
            builder.add_paragraph(p)
        for f in args.add_formula:
            builder.add_formula(f, display=True)
        builder.save(args.output)
        print(f"Modified: {args.output}")


if __name__ == '__main__':
    _cli()
