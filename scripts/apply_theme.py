#!/usr/bin/env python3
"""
Apply a format-spec theme to a python-docx Document via the STYLE SYSTEM.

Usage:
    from apply_theme import load_theme, apply_theme, themed

    theme = load_theme('format-spec/thesis-sjtu.md', degree='硕士')
    apply_theme(doc, theme)

    # Then add content using Word styles (no per-run formatting):
    doc.add_paragraph('第一章 引言', style='Heading 1')
    doc.add_paragraph('正文内容...', style='Normal')
    themed(doc).add_figure('img.jpg', '图1 系统架构')
"""
import re
from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree


# ============================================================
# Theme parser
# ============================================================

def load_theme(path, degree='硕士'):
    """Parse a format-spec/*.md theme file into nested dict."""
    cfg, section = {}, None
    with open(path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            m = re.match(r'^\[(\w+)\]$', line)
            if m:
                section = m.group(1)
                cfg[section] = {}
                continue
            m2 = re.match(r'^(\w+)\s*=\s*(.+)$', line)
            if m2 and section:
                cfg[section][m2.group(1)] = m2.group(2).strip().replace('{degree}', degree)
    return cfg


# ============================================================
# Value parsing
# ============================================================

def _size(s):
    s = s.strip()
    if s.endswith('pt'): return Pt(float(s[:-2]))
    if s.endswith('cm'): return Cm(float(s[:-2]))
    if s.endswith('mm'): return Cm(float(s[:-2])/10)
    if s.endswith('in'): return Inches(float(s[:-2]))
    return s

def _get(theme, section, key, default=None):
    return theme.get(section, {}).get(key, default)

def _bool(theme, section, key):
    return _get(theme, section, key, 'false') == 'true'

def _align(v):
    if v == 'center': return WD_ALIGN_PARAGRAPH.CENTER
    if v == 'right':  return WD_ALIGN_PARAGRAPH.RIGHT
    if v == 'left':   return WD_ALIGN_PARAGRAPH.LEFT
    return None


# ============================================================
# Style overrides (the fix for Problem 1 & 2)
# ============================================================

def _set_style_font(style, theme, section, prefix=''):
    """Override font in a Word style object. Handles eastAsia correctly."""
    font_name = _get(theme, section, prefix + 'font')
    size = _get(theme, section, prefix + 'size')
    bold = _bool(theme, section, prefix + 'bold')

    if font_name:
        style.font.name = font_name
        # Set eastAsia via XML (python-docx doesn't expose this on styles)
        rPr = style.element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.insert(0, rFonts)
        rFonts.set(qn('w:eastAsia'), font_name)
        rFonts.set(qn('w:ascii'), font_name)
        rFonts.set(qn('w:hAnsi'), font_name)
        # Remove theme font references (majorHAnsi/majorEastAsia) so our explicit fonts win
        for attr in ['w:asciiTheme', 'w:eastAsiaTheme', 'w:hAnsiTheme', 'w:cstheme']:
            qattr = qn(attr)
            if qattr in rFonts.attrib:
                del rFonts.attrib[qattr]

    if size:
        style.font.size = _size(size)
    if bold:
        style.font.bold = True

    # Color — theme must specify, no hardcoded default
    color = _get(theme, section, prefix + 'color')
    if not color:
        color = _get(theme, section, 'color')  # try without prefix (e.g. 'color' not 'heading_color')
    if color:
        from docx.shared import RGBColor
        style.font.color.rgb = RGBColor(
            int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
        )


def _set_style_paragraph(style, theme, section):
    """Apply paragraph formatting to a style."""
    for key, attr in [('before_spacing', 'space_before'), ('after_spacing', 'space_after')]:
        v = _get(theme, section, key)
        if v:
            setattr(style.paragraph_format, attr, _size(v))

    v = _get(theme, section, 'line_spacing')
    if v:
        style.paragraph_format.line_spacing = float(v)

    v = _get(theme, section, 'first_line_indent')
    if v:
        style.paragraph_format.first_line_indent = _size(v)

    v = _get(theme, section, 'align')
    if v:
        al = _align(v)
        if al: style.paragraph_format.alignment = al


def apply_theme(doc, theme):
    """Apply a theme to a Document. Overrides styles, page setup, headers/footers.
    After calling this, use style names for all content (no per-run formatting needed)."""

    # --- Page setup ---
    for sec in doc.sections:
        for key, attr in [('page_width','page_width'), ('page_height','page_height'),
                          ('margin_top','top_margin'), ('margin_bottom','bottom_margin'),
                          ('margin_left','left_margin'), ('margin_right','right_margin')]:
            v = _get(theme, 'document', key)
            if v: setattr(sec, attr, _size(v))

    # --- Override Normal (body text) ---
    normal = doc.styles['Normal']
    _set_style_font(normal, theme, 'body')
    _set_style_paragraph(normal, theme, 'body')

    # --- Override Heading 1/2/3 ---
    heading_map = {
        'Heading 1': 'heading1',
        'Heading 2': 'heading2',
        'Heading 3': 'heading3',
    }
    for style_name, section in heading_map.items():
        if section in theme:
            style = doc.styles[style_name]
            _set_style_font(style, theme, section)
            _set_style_paragraph(style, theme, section)

    # --- Header / Footer ---
    section = doc.sections[0]
    hdr_text = _get(theme, 'header', 'text')
    if hdr_text:
        header = section.header
        header.paragraphs[0].text = ''
        run = header.paragraphs[0].add_run(hdr_text)
        font = _get(theme, 'header', 'font')
        size = _get(theme, 'header', 'size')
        if font:
            run.font.name = font
            rPr = run._element.get_or_add_rPr()
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                rPr.insert(0, rFonts)
            rFonts.set(qn('w:eastAsia'), font)
        if size:
            run.font.size = _size(size)

    footer = section.footer
    footer.paragraphs[0].text = ''
    footer.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    ffont = _get(theme, 'footer', 'font')
    fsize = _get(theme, 'footer', 'size')
    if ffont or fsize:
        # Footer just gets page number (placeholder)
        run = footer.paragraphs[0].add_run(' ')
        if ffont: run.font.name = ffont
        if fsize: run.font.size = _size(fsize)

    return doc


# ============================================================
# Content helpers — use styles, not per-run formatting
# ============================================================

class ThemedDocument:
    """Wrapper that adds content using the style system."""

    def __init__(self, doc, theme):
        self.doc = doc
        self.theme = theme
        self._fig_num = 0
        self._tbl_num = 0

    def heading(self, text, level=1):
        style = f'Heading {level}'
        p = self.doc.add_paragraph(text, style=style)
        return p

    def body(self, text):
        p = self.doc.add_paragraph(text, style='Normal')
        return p

    def figure(self, image_path, caption_text):
        """Insert figure with centered numbered caption. Uses body (center) + custom label."""
        self._fig_num += 1
        prefix = _get(self.theme, 'figure', 'caption_prefix', '图')

        p_img = self.doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.add_run().add_picture(image_path, width=Cm(10))

        # Caption: bold label + normal text
        p_cap = self.doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER

        r_label = p_cap.add_run(f'{prefix}{self._fig_num} ')
        lbl_font = _get(self.theme, 'figure', 'label_font')
        lbl_size = _get(self.theme, 'figure', 'label_size')
        lbl_bold = _bool(self.theme, 'figure', 'label_bold')
        if lbl_font:
            r_label.font.name = lbl_font
            rPr = r_label._element.get_or_add_rPr()
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                rPr.insert(0, rFonts)
            rFonts.set(qn('w:eastAsia'), lbl_font)
        if lbl_size: r_label.font.size = _size(lbl_size)
        if lbl_bold: r_label.bold = True

        r_text = p_cap.add_run(caption_text)
        cap_font = _get(self.theme, 'figure', 'caption_font')
        cap_size = _get(self.theme, 'figure', 'caption_size')
        if cap_font:
            r_text.font.name = cap_font
            rPr = r_text._element.get_or_add_rPr()
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                rPr.insert(0, rFonts)
            rFonts.set(qn('w:eastAsia'), cap_font)
        if cap_size: r_text.font.size = _size(cap_size)

        p_cap.paragraph_format.space_after = Pt(12)
        return p_cap

    def table(self, headers, rows_data, caption_text):
        """Insert table with numbered caption above."""
        self._tbl_num += 1
        prefix = _get(self.theme, 'table', 'caption_prefix', '表')

        # Caption above
        p_cap = self.doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_label = p_cap.add_run(f'{prefix}{self._tbl_num} ')
        lbl_font = _get(self.theme, 'table', 'label_font')
        lbl_size = _get(self.theme, 'table', 'label_size')
        lbl_bold = _bool(self.theme, 'table', 'label_bold')
        if lbl_font:
            r_label.font.name = lbl_font
            rPr = r_label._element.get_or_add_rPr()
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                rPr.insert(0, rFonts)
            rFonts.set(qn('w:eastAsia'), lbl_font)
        if lbl_size: r_label.font.size = _size(lbl_size)
        if lbl_bold: r_label.bold = True

        r_text = p_cap.add_run(caption_text)
        cap_font = _get(self.theme, 'table', 'caption_font')
        cap_size = _get(self.theme, 'table', 'caption_size')
        if cap_font:
            r_text.font.name = cap_font
            rPr = r_text._element.get_or_add_rPr()
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                rPr.insert(0, rFonts)
            rFonts.set(qn('w:eastAsia'), cap_font)
        if cap_size: r_text.font.size = _size(cap_size)

        # Build table
        ncols, nrows = len(headers), len(rows_data) + 1
        tbl = self.doc.add_table(rows=nrows, cols=ncols, style='Light Grid Accent 1')

        cfont = _get(self.theme, 'table', 'content_font')
        csize = _get(self.theme, 'table', 'content_size')

        for i, h in enumerate(headers):
            cell = tbl.rows[0].cells[i]
            cell.text = h
            for para in cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in para.runs:
                    run.bold = True
                    if cfont: run.font.name = cfont
                    if csize: run.font.size = _size(csize)

        for ri, row_data in enumerate(rows_data):
            for ci, cell_text in enumerate(row_data):
                cell = tbl.rows[ri+1].cells[ci]
                cell.text = cell_text
                for para in cell.paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in para.runs:
                        if cfont: run.font.name = cfont
                        if csize: run.font.size = _size(csize)

        return tbl

    def formula(self, latex_str):
        """Insert LaTeX formula (centered)."""
        from latex2omml import latex_to_omml
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        omml = latex_to_omml(latex_str)
        run = p.add_run('')
        run._element.append(etree.fromstring(omml))
        return p

    def cover(self, title_text, info_lines):
        """Add cover page with title and info centered."""
        for _ in range(6):
            self.doc.add_paragraph('')

        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(title_text)
        cover_font = _get(self.theme, 'cover', 'title_font')
        cover_size = _get(self.theme, 'cover', 'title_size')
        cover_bold = _bool(self.theme, 'cover', 'title_bold')
        if cover_font:
            r.font.name = cover_font
            rPr = r._element.get_or_add_rPr()
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                rPr.insert(0, rFonts)
            rFonts.set(qn('w:eastAsia'), cover_font)
        if cover_size: r.font.size = _size(cover_size)
        if cover_bold: r.bold = True

        info_font = _get(self.theme, 'cover', 'info_font')
        info_size = _get(self.theme, 'cover', 'info_size')

        for line in info_lines:
            self.doc.add_paragraph('')
            p = self.doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(line)
            if info_font: r.font.name = info_font
            if info_size: r.font.size = _size(info_size)


def themed(doc, theme=None):
    """Quick access: themed(doc).heading('...', 1)"""
    return ThemedDocument(doc, theme)


# ============================================================
# Quick test
# ============================================================
if __name__ == '__main__':
    import sys
    sys.path.insert(0, '.')
    theme = load_theme('format-spec/thesis-sjtu.md', '硕士')

    doc = Document()
    apply_theme(doc, theme)
    t = themed(doc, theme)

    t.heading('第一章 测试', 1)
    t.body('这是正文，使用 Normal 样式，字体应为 SimSun 12pt，1.5 倍行距。' * 3)
    t.heading('1.1 小节', 2)
    t.body('小节正文。')

    doc.save('theme_test_output.docx')
    print('Saved: theme_test_output.docx — check Normal style font in Word')
