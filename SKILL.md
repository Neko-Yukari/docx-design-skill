---
name: docx
description: Read, create, and edit Word documents (.docx and .doc) on any platform. Supports LaTeX formulas, Chinese typesetting, modular themes, and Word style integration. Use when working with .docx/.doc files — reading, creating, editing, filling templates, generating thesis papers, or converting LaTeX to Word equations.
license: MIT
---

# DOCX / DOC Skill

Read, create, edit Word documents on Windows using `python-docx`.

---

## Critical Rule (Windows)

**NEVER run Python inline in PowerShell.** Powershell chokes on f-strings, Chinese chars, and nested quotes. Always: write `.py` file → `python script.py`.

---

## Bundled Scripts

| Script | Purpose |
|--------|---------|
| `scripts/extract_docx.py` | Read .docx → Markdown. Full extraction, structure outline, range/section/grep modes |
| `scripts/edit_docx.py` | CLI editing: replace, insert, delete, fill, paragraph-level ops, range-scoped edits |
| `scripts/apply_theme.py` | Apply format-spec themes to documents via Word style system |
| `scripts/latex2omml.py` | Convert LaTeX formulas to OMML XML for .docx embedding |
| `scripts/convert_to_doc.py` | Convert .docx to legacy .doc (MS Word COM or LibreOffice) |
| `scripts/merge_runs.py` | Unpack .docx and merge adjacent runs with identical formatting (fix split-run bug) |

**Script paths**: PowerShell does NOT expand `~`. Use full absolute path to bundled scripts, or copy them to working directory first.

---

## Quick Reference

| Task | Method | Section |
|------|--------|---------|
| Read .docx (full) | `python extract_docx.py file.docx` | [Reading](#reading-docx) |
| Get document outline | `python extract_docx.py file.docx --structure` | [Chunking & Search](#chunking--search-modes) |
| Extract paragraph range | `python extract_docx.py file.docx --range 10-25` | [Chunking & Search](#chunking--search-modes) |
| Extract by section name | `python extract_docx.py file.docx --section "Results"` | [Chunking & Search](#chunking--search-modes) |
| Search in document | `python extract_docx.py file.docx --grep "keyword" --context 2` | [Chunking & Search](#chunking--search-modes) |
| Read .doc text | Convert to .docx first, then read | [.doc Handling](#doc-handling) |
| Convert .docx → .doc | `python convert_to_doc.py in.docx [out.doc]` | [.doc Handling](#doc-handling) |
| Create new .docx | Write a .py script with `python-docx` | [Creating](#creating-docx) |
| Apply format theme | `from apply_theme import load_theme, apply_theme` | [Theme System](#theme-system) |
| Insert LaTeX formula | `from latex2omml import insert_formula` | [LaTeX Formulas](#latex-formulas) |
| Edit .docx (replace) | `python edit_docx.py in.docx out.docx --replace "old" "new"` | [Editing](#editing-docx) |
| Edit paragraphs | `python edit_docx.py in.docx out.docx --paragraph 3 "new text"` | [Paragraph Editing](#paragraph-level-editing) |
| Edit within range | `python edit_docx.py in.docx out.docx --range 10-25 --replace "a" "b"` | [Editing](#editing-docx) |
| Fill template form | `python edit_docx.py template.docx out.docx --fill map.json` | [Editing](#editing-docx) |
| Fix split-run bug | `python merge_runs.py in.docx unpacked/` | [Editing](#editing-docx) |
| Track changes | Unpack ZIP → edit XML → repack | [OOXML Reference](#ooxml-reference) |

---

## Reading .docx

### Full Extraction

```powershell
# Full path:
python C:\Users\81004\.config\opencode\skills\docx\scripts\extract_docx.py report.docx

# Or copy first:
copy C:\Users\81004\.config\opencode\skills\docx\scripts\extract_docx.py .
python extract_docx.py report.docx
```

Outputs `report_extracted.md` with headings and tables preserved.

### Chunking & Search Modes

For large documents (>30 pages), use chunking modes to avoid context blowup. Each mode targets only the relevant portion:

#### `--structure` — Document Outline

Outputs heading hierarchy with paragraph ranges and table associations. No body text — ~100-300 tokens.

```powershell
python extract_docx.py thesis.docx --structure
# Output:
# [document]  thesis.docx  (186 paragraphs, 12 tables)
# +-- Heading 1 (p0):  第一章 引言               [p0-p15,  table:1]
# |   +-- Heading 2 (p1): 1.1 研究背景            [p1-p5]
# |   +-- Heading 2 (p6): 1.2 国内外现状           [p6-p12]
# |   +-- Heading 2 (p13): 1.3 本文工作            [p13-p15]
# +-- Heading 1 (p16): 第二章 系统模型              [p16-p45, table:2]
```

Headings are detected by Word style (Heading 1/2/3). For documents without styles, add `--detect-freeform` to use regex-based heading detection (patterns like "第X章", "1.1 Title", etc.).

```powershell
python extract_docx.py report.docx --structure --detect-freeform
# Free-form headings show confidence scores: H1 H2 [conf:0.95]
```

`--detect-freeform` requires `--structure` (error if used alone).

#### `--range N-M` — Extract Paragraph Range

Only extracts paragraphs N through M (0-based indexing). Tables within the range are included.

```powershell
python extract_docx.py thesis.docx --range 16-45
# Output: thesis_range_16_45.md
```

Errors on invalid ranges (N < 1, N > M, M exceeds document length).

#### `--section "name"` — Extract by Section Name

Matches section heading by substring (case-insensitive). Extracts that section and all its sub-sections.

```powershell
python extract_docx.py thesis.docx --section "系统模型"
# Output: thesis_section_系统模型.md

# Exact match only:
python extract_docx.py thesis.docx --section "系统模型" --exact

# Non-existent section → clear error message
python extract_docx.py thesis.docx --section "nonexistent"
# Error: no section matching "nonexistent"
```

#### `--grep "pattern"` — Full-Text Search

Searches all paragraphs and returns matches with surrounding context.

```powershell
python extract_docx.py thesis.docx --grep "RFocus" --context 2
# Output:
# [p42] "...采用了RFocus技术中的majority voting算法..."
#   <- p40: "反向散射通信是物联网领域的..."
#   -> p43: "该算法通过迭代优化实现..."
#
# [p78] "...RFocus系统由3000根天线..."
```

For custom processing, write a .py script:

```python
from docx import Document

doc = Document('file.docx')
for para in doc.paragraphs:
    print(para.text)        # full text of each paragraph
for table in doc.tables:
    for row in table.rows:
        print(' | '.join(cell.text for cell in row.cells))
```

---

## Creating .docx

**Workflow**: Write a .py script → `python script.py`. Core building blocks:

```python
from docx import Document
from docx.shared import Cm, Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from lxml import etree

doc = Document()

# Page setup (A4)
section = doc.sections[0]
section.page_width  = Cm(21)
section.page_height = Cm(29.7)
section.left_margin = Cm(2.5)
section.right_margin = Cm(2.5)

# Heading
h = doc.add_heading('Title', level=1)

# Paragraph with formatting
p = doc.add_paragraph()
run = p.add_run('Bold text. ')
run.bold = True
run.font.size = Pt(12)
run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
p.add_run('Normal text.')

# Table
t = doc.add_table(rows=2, cols=2, style='Light Grid Accent 1')
t.rows[0].cells[0].text = 'Header 1'
t.rows[0].cells[1].text = 'Header 2'

# Image
doc.add_picture(r'C:\path\to\image.jpg', width=Cm(10))

# Chinese font (must set eastAsia separately!)
def set_chinese_font(run, font_name):
    run.font.name = font_name
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = etree.SubElement(rPr, qn('w:rFonts'))
    rFonts.set(qn('w:eastAsia'), font_name)

run2 = p.add_run(' 中文宋体 ')
set_chinese_font(run2, 'SimSun')

# Math formula — use latex2omml.py instead of raw OMML (see LaTeX Formulas section)
from latex2omml import insert_formula
insert_formula(p, r"E = mc^2")

doc.save('output.docx')
```

### Using the Theme System

For thesis/formal documents with strict formatting requirements, use the theme system instead of manual formatting:

```python
from apply_theme import load_theme, apply_theme, themed

# Load a theme (available: thesis-sjtu, thesis-generic)
theme = load_theme('format-spec/thesis-sjtu.md', degree='硕士')
apply_theme(doc, theme)

# Then add content using Word styles — formatting is automatic:
doc.add_paragraph('第一章 引言', style='Heading 1')
doc.add_paragraph('正文内容...', style='Normal')

# themed() context manager provides helper methods
with themed(doc, theme) as t:
    t.add_figure('chart.png', '图1 系统架构图')
    t.add_table(headers=['参数', '数值'], rows=[['频率', '2.4GHz'], ['功率', '10dBm']])
```

See [Theme System](#theme-system) for details on creating custom themes.

### Visual Verification (Render → Inspect → Fix)

Professional document editing requires checking layout visually. On Windows:

```powershell
# Option A: LibreOffice (if installed)
soffice --headless --convert-to pdf output.docx

# Option B: Open in Word, review layout, close
# Option C: Use the pdf-reader skill on the exported PDF
```

After each significant edit: render → inspect at 100% zoom → fix any layout breaks.

---

## Editing .docx

**Workflow**: Use bundled `edit_docx.py` for simple edits, write a script for complex ones.

### Prepare First: Eliminate Split-Run Problems at XML Level

The #1 editing bug is text split across multiple `<w:r>` elements. Instead of working around it in Python, fix it at the source:

```powershell
# Unpack + merge adjacent runs with identical formatting
python merge_runs.py input.docx unpacked/

# Now edit word/document.xml in the unpacked directory
# Then repack (the bundled edit_docx.py does this automatically)
```

This is the Anthropic approach — by merging runs before editing, text replacement becomes reliable without any special Python workaround.

### Fast Path (CLI)

```powershell
# Simple find-and-replace
python edit_docx.py in.docx out.docx --replace "旧文本" "新文本"

# Multiple replacements
python edit_docx.py in.docx out.docx --replace "张三" "李四" --replace "2024" "2025"

# List all paragraphs (quick overview)
python edit_docx.py in.docx out.docx --list

# Insert after paragraph 3
python edit_docx.py in.docx out.docx --insert-after 3 "新段落"

# Delete paragraph 5
python edit_docx.py in.docx out.docx --delete 5

# Fill template blanks from JSON
python edit_docx.py template.docx done.docx --fill replacements.json
# replacements.json: {"姓名：": "姓名：张三", "学号：": "学号：2024001"}

# Replace only within paragraph range 2-5 (0-based)
python edit_docx.py in.docx out.docx --range 2-5 --replace "天线" "antenna"
```

### Paragraph-Level Editing

Target individual paragraphs by index (0-based, from `--structure` or `--list`):

```powershell
# Replace entire paragraph (preserves formatting of first run)
python edit_docx.py in.docx out.docx --paragraph 2 "New paragraph text"

# Replace paragraph content from a file
python edit_docx.py in.docx out.docx --paragraph 2 --file content.md

# Append text to end of paragraph
python edit_docx.py in.docx out.docx --paragraph 3 --append " (updated 2026)"

# Prepend text to start of paragraph
python edit_docx.py in.docx out.docx --paragraph 3 --prepend "Note: "

# Set paragraph style (e.g., convert body text to heading)
python edit_docx.py in.docx out.docx --set-style 0 "Heading 1"
python edit_docx.py in.docx out.docx --set-style 0 "Heading 1" --set-style 2 "Heading 2"
```

All paragraph index operations validate bounds — out-of-range indices produce clear error messages.

### Script Path (Custom Logic)

**The #1 pitfall**: Word splits text across multiple XML runs. "你好世界" can be Run1="你好", Run2="世界". Naive per-run replace misses it. Always use this pattern:

```python
# Robust replace — always use this function, never iterate runs directly
def replace_in_para(para, old, new):
    if old not in para.text:
        return False
    for run in para.runs:
        if old in run.text:
            run.text = run.text.replace(old, new)
            return True
    # Text spans multiple runs — consolidate
    para.runs[0].text = para.text.replace(old, new)
    for run in para.runs[1:]:
        run._element.getparent().remove(run._element)
    return True

# Usage
doc = Document('template.docx')
for para in doc.paragraphs:
    replace_in_para(para, '姓名：', '姓名：张三')
# Also search tables and headers/footers
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                replace_in_para(para, '姓名：', '姓名：张三')
doc.save('filled.docx')
```

**Insert at position**:
```python
# AFTER paragraph 3
target = doc.paragraphs[3]
new_para = doc.add_paragraph('New text')
target._element.addnext(new_para._element)

# BEFORE paragraph 3
target = doc.paragraphs[3]
heading = doc.add_heading('New Section', level=2)
target._element.addprevious(heading._element)

# Delete paragraph 4
doc.paragraphs[4]._element.getparent().remove(doc.paragraphs[4]._element)
```

**Replace image** (only via OOXML — `python-docx` can't do this):
```python
import zipfile, shutil, os
def replace_image(docx_path, out_path, old_media_name, new_image_path):
    tmp = '_tmp'
    os.makedirs(tmp, exist_ok=True)
    with zipfile.ZipFile(docx_path) as z: z.extractall(tmp)
    shutil.copy2(new_image_path, os.path.join(tmp, 'word', 'media', old_media_name))
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for r,d,f in os.walk(tmp):
            for fn in f: z.write(os.path.join(r,fn), os.path.relpath(os.path.join(r,fn),tmp))
    shutil.rmtree(tmp)
```

---

## Theme System

For documents with strict formatting requirements (thesis, official reports), use the theme system to apply consistent formatting via Word styles — no per-run manual formatting.

### Available Themes

| Theme File | Description |
|------------|-------------|
| `format-spec/thesis-sjtu.md` | 上海交通大学学位论文 (SJTU thesis) |
| `format-spec/thesis-generic.md` | 通用中国学位论文 GB/T 7713 (generic Chinese thesis) |

### Usage

```python
from apply_theme import load_theme, apply_theme, themed

# Load theme (replace {degree} placeholder with "硕士"/"博士")
theme = load_theme('format-spec/thesis-sjtu.md', degree='硕士')
apply_theme(doc, theme)

# Then add content using Word styles — formatting is automatic:
doc.add_paragraph('第一章 引言', style='Heading 1')
doc.add_paragraph('正文内容...', style='Normal')

# themed() context manager adds helper methods
with themed(doc, theme) as t:
    t.add_figure('system.png', '图1 系统架构图')
    t.add_table(
        headers=['参数', '数值'],
        rows=[['频率', '2.4GHz'], ['功率', '10dBm']]
    )
```

### How It Works

`apply_theme()` configures Word's built-in styles (Heading 1/2/3, Normal, etc.) with the font, size, spacing, and alignment defined in the theme file. It handles:

- **Chinese fonts correctly** — sets both `w:ascii` and `w:eastAsia` attributes
- **Page setup** — margins, page size
- **Header/footer** — with `{degree}` placeholder substitution
- **All required sections** — cover, abstract, TOC, headings, body, header, footer
- **Optional sections** — figures, tables, references

### Creating Custom Themes

Theme files live under `format-spec/`. Full schema: `format-spec/THEME-SCHEMA.md`.

Minimal theme skeleton:

```ini
[document]
page_width     = 21cm
page_height    = 29.7cm
margin_top     = 2.5cm
margin_bottom  = 2.5cm
margin_left    = 3.0cm
margin_right   = 2.5cm

[cover]
title_font     = SimHei
title_size     = 22pt
title_color    = 000000
title_bold     = true
title_align    = center
info_font      = SimSun
info_size      = 14pt
info_color     = 000000
info_align     = center

[abstract]
heading_font        = SimHei
heading_size        = 18pt
heading_bold        = true
heading_align       = center
heading_text        = 摘要
body_font           = SimSun
body_size           = 12pt
body_line_spacing   = 1.5
keywords_label      = 关键词：
keywords_label_font = SimHei
keywords_font       = SimSun

[heading1]
font = SimHei
size = 16pt
color = 000000
bold = true
align = center
before_spacing = 12pt
after_spacing = 6pt

[heading2]
font = SimHei
size = 14pt
color = 000000
bold = true
align = left
before_spacing = 6pt
after_spacing = 3pt

[heading3]
font = SimHei
size = 12pt
color = 000000
bold = true
align = left
before_spacing = 3pt
after_spacing = 3pt

[body]
font = SimSun
size = 12pt
color = 000000
line_spacing = 1.5
first_line_indent = 0.74cm
before_spacing = 0pt
after_spacing = 0pt

[header]
font = SimSun
size = 9pt
color = 000000
text = {degree}学位论文

[footer]
font = SimSun
size = 9pt
color = 000000
page_number_align = center
```

---

## LaTeX Formulas

Convert LaTeX math expressions to OMML (Office Math Markup Language) for native .docx embedding — no MathType, no images, no copy-paste from Word's equation editor.

### Usage

```python
from latex2omml import latex_to_omml, insert_formula

# Get OMML XML string
omml = latex_to_omml(r"x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}")

# Or insert directly into a paragraph
from docx import Document
doc = Document()
p = doc.add_paragraph()
p.add_run('The quadratic formula: ')
insert_formula(p, r"x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}")
doc.save('math.docx')
```

### CLI Quick Test

```powershell
python latex2omml.py "E = mc^2"
# Outputs OMML XML to stdout
```

### Supported LaTeX Constructs

| LaTeX | Description | Example |
|-------|-------------|---------|
| `x^{n}` / `x_{n}` | Superscript / subscript | `a_i^2` |
| `\frac{a}{b}` | Fraction | `\frac{1}{2}` |
| `\sqrt{x}` / `\sqrt[n]{x}` | Square root / nth root | `\sqrt{a+b}` |
| `\sum`, `\int`, `\prod` | Large operators | `\sum_{i=1}^n` |
| `\pm`, `\cdot`, `\times` | Binary operators | `a \pm b` |
| `\alpha`, `\beta`, `\pi` | Greek letters | `\alpha + \beta` |
| `\infty`, `\partial`, `\nabla` | Special symbols | `\infty` |
| `\sin`, `\cos`, `\log` | Standard functions | `\sin(x)` |
| `\text{...}` | Text within math | `x \text{ where } x>0` |
| `\left( ... \right)` | Auto-sized parentheses | `\left(\frac{a}{b}\right)` |

If a LaTeX construct isn't supported, the converter raises an error with the unsupported element name. For complex formulas, write OMML XML directly (see [OOXML Reference > Math Formulas](#math-formulas-omml)).

---

## .doc Handling

`.doc` is binary — `python-docx` can't read it. Convert to `.docx` first:

```python
# Windows: MS Word COM (requires pip install pywin32)
import os, sys
def doc_to_docx(path):
    import win32com.client
    w = win32com.client.Dispatch('Word.Application')
    w.Visible = False
    out = os.path.splitext(path)[0] + '.docx'
    w.Documents.Open(os.path.abspath(path)).SaveAs2(os.path.abspath(out), FileFormat=16)
    w.Quit()
    return out
```

Or: `soffice --headless --convert-to docx input.doc --outdir .` (requires LibreOffice).

Or: ask user to open in Word → Save As `.docx`.

### Converting .docx → .doc (Legacy Format)

Use the bundled `convert_to_doc.py` script:

```powershell
python convert_to_doc.py report.docx
# Output: report.doc (MS Word COM on Windows, LibreOffice fallback)

python convert_to_doc.py report.docx legacy_report.doc
# Custom output path
```

The script auto-detects the best conversion method:
1. **Windows**: MS Word COM via `pywin32` (requires `pip install pywin32`)
2. **Cross-platform fallback**: LibreOffice `soffice` (must be in PATH)

---

## OOXML Reference

For features beyond `python-docx` (tracked changes, comments): unpack the .docx ZIP, edit XML, repack.

```python
import zipfile, shutil, os

def unpack(docx, out_dir):
    if os.path.exists(out_dir): shutil.rmtree(out_dir)
    os.makedirs(out_dir)
    with zipfile.ZipFile(docx) as z: z.extractall(out_dir)

def pack(in_dir, out_path):
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for r,d,f in os.walk(in_dir):
            for fn in f:
                fp = os.path.join(r,fn)
                z.write(fp, os.path.relpath(fp, in_dir))
```

Key XML files: `word/document.xml` (body), `word/comments.xml`, `word/media/` (images).

### Tracked Changes

```xml
<!-- Insertion -->
<w:ins w:id="1" w:author="Reviewer" w:date="2026-05-09T00:00:00Z">
  <w:r><w:rPr/><w:t>inserted text</w:t></w:r>
</w:ins>

<!-- Deletion -->
<w:del w:id="2" w:author="Reviewer" w:date="2026-05-09T00:00:00Z">
  <w:r><w:delText>deleted text</w:delText></w:r>
</w:del>
```

### Math Formulas (OMML)

```xml
<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:oMath>
    <m:f>                           <!-- fraction -->
      <m:num><m:r><m:t>a</m:t></m:r></m:num>
      <m:den><m:r><m:t>b</m:t></m:r></m:den>
    </m:f>
    <m:sSup>                        <!-- superscript -->
      <m:e><m:r><m:t>x</m:t></m:r></m:e>
      <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
    </m:sSup>
    <m:rad>                         <!-- radical -->
      <m:e><m:r><m:t>a+b</m:t></m:r></m:e>
    </m:rad>
  </m:oMath>
</m:oMathPara>
```

Element reference: `<m:f>`=fraction, `<m:sSup>`=superscript, `<m:sSub>`=subscript, `<m:rad>`=radical (√), `<m:nary>`=∑/∫/∏, `<m:r><m:t>`=text.

Unicode: `\u00B1`=±, `\u2211`=∑, `\u222B`=∫, `\u221E`=∞, `\u03C0`=π.

---

## Reference Tables

### Chinese Fonts

| Name | XML value | Name | XML value |
|------|-----------|------|-----------|
| 宋体 | `SimSun` | 仿宋 | `FangSong` |
| 黑体 | `SimHei` | 微软雅黑 | `Microsoft YaHei` |
| 楷体 | `KaiTi` | 等线 | `DengXian` |

### Common Formatting

| Task | Code |
|------|------|
| Bold | `run.bold = True` |
| Italic | `run.italic = True` |
| Underline | `run.underline = True` |
| Strikethrough | `run.font.strike = True` |
| Font size | `run.font.size = Pt(12)` |
| Font color | `run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)` |
| Highlight | `run.font.highlight_color = WD_COLOR_INDEX.YELLOW` |
| Page break | `doc.add_page_break()` |
| Alignment | `para.alignment = WD_ALIGN_PARAGRAPH.CENTER` |
| Line spacing | `para.paragraph_format.line_spacing = 1.5` |

### Dependencies

| Package | Install | Purpose |
|---------|---------|---------|
| `python-docx` | `pip install python-docx` | Core: read, create, edit .docx |
| `lxml` | comes with python-docx | XML parsing for OOXML/OMML |
| `latex2mathml` | `pip install latex2mathml` | LaTeX → MathML (used by latex2omml.py) |
| `mammoth` | `pip install mammoth` | Alternative: .docx→Markdown |
| `pywin32` | `pip install pywin32` | .doc→.docx / .docx→.doc via MS Word COM |
| LibreOffice | `winget install LibreOffice.LibreOffice` | .doc→.docx, PDF export, .docx→.doc |

---

## Quality Expectations

Professional document delivery standards (adapted from OpenAI's doc skill):

- **Render after each meaningful edit** — convert to PDF, inspect at 100% zoom
- **No clipped or overlapping text** — check table cells and multi-column layouts
- **Consistent typography** — use styles instead of manual formatting; set fonts via `<w:eastAsia>` for Chinese
- **Clean output** — remove temp files after delivery; use descriptive filenames
- **No Unicode dashes** — use ASCII hyphens in OOXML; use `&#x201C;`/`&#x201D;` for smart quotes
- **Citations must be human-readable** — never leave tool tokens or placeholder strings in final output

## Limitations

- **python-docx** cannot read: tracked changes, comments, text boxes, equations, SmartArt, charts
- **.doc** must be converted to .docx before reading
- For tracked changes / comments: use the OOXML workflow (unpack ZIP → edit XML → repack)
- **Bundled scripts** (in `~/.config/opencode/skills/docx/scripts/`):
  - `extract_docx.py` — read .docx → Markdown (full, structure, range, section, grep)
  - `edit_docx.py` — CLI editing (replace, insert, delete, fill, paragraph ops, range-scoped edits)
  - `apply_theme.py` — apply format-spec themes via Word style system
  - `latex2omml.py` — LaTeX → OMML formula conversion
  - `convert_to_doc.py` — .docx → legacy .doc conversion
  - `merge_runs.py` — unpack + merge adjacent runs (fix split-run at XML level)
