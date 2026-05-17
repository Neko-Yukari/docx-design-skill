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

## Quick Reference

| Task | Method | Section |
|------|--------|---------|
| **Create .docx (recommended)** | `python docx_tool.py create --output out.docx --title ...` | [Creating](#creating-docx) |
| Read .docx text | `python extract_docx.py file.docx` | [Reading](#reading-docx) |
| Read .doc text | Convert to .docx first, then read | [.doc Handling](#doc-handling) |
| Edit .docx (simple) | `python edit_docx.py in.docx out.docx --replace "old" "new"` | [Editing](#editing-docx) |
| Fill template form | `python edit_docx.py template.docx out.docx --fill map.json` | [Editing](#editing-docx) |
| Track changes | Unpack ZIP → edit XML → repack | [OOXML Reference](#ooxml-reference) |

**Script paths**: PowerShell does NOT expand `~`. Use full absolute path to bundled scripts, or copy them to working directory first.

---

## Recommended Workflow (docx_tool.py)

Instead of writing a one-off `generate_report.py` every time, use the bundled **`docx_tool.py`** — a reusable builder with presets for Chinese/English typography, automatic OMML formula insertion, and theme-font cleanup.

### CLI — one-liner document creation

```powershell
python "C:\Users\81004\.config\opencode\skills\docx\scripts\docx_tool.py" create `
  --output report.docx `
  --preset chinese `
  --title "项目报告" `
  --add-heading "1:项目概述" `
  --add-heading "2:技术方案" `
  --add-paragraph "本文使用黑体作为正文字体。" `
  --add-formula "x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}"
```

What `--preset chinese` does automatically:
- Normal → SimHei (黑体), 10.5pt, black
- Heading 1/2/3 → SimSun (宋体), bold, black
- Removes `eastAsiaTheme` / `asciiTheme` theme references
- Sets explicit `w:color w:val="000000"` on all styles

### Python API — fluent builder

```python
import sys
sys.path.insert(0, r"C:\Users\81004\.config\opencode\skills\docx\scripts")
from docx_tool import DocumentBuilder

# One-liner preset + chain build
DocumentBuilder() \
    .preset_chinese() \
    .add_heading("项目报告", level=1) \
    .add_heading("1. 概述", level=2) \
    .add_paragraph("使用黑体作为正文字体。") \
    .add_formula(r"x = \frac{1}{2}", display=True) \
    .add_paragraph("行内公式示例：") \
    .add_formula(r"E = mc^2", display=False) \
    .save("report.docx")
```

**Key difference from raw `python-docx`:**
- No manual `eastAsia` font XML manipulation
- No theme reference cleanup
- No OMML structure foot-guns (`add_formula()` automatically inserts at `w:p` level for display, `w:r` level for inline)
- Preset defaults eliminate repetitive setup

### When to use raw `python-docx`

Use `docx_tool.py` for 90% of cases. Drop down to raw `python-docx` only when:
- You need fine-grained run-level formatting (per-character fonts/colors)
- You need custom styles beyond Heading 1-3 + Normal
- You need to manipulate OOXML directly (tracked changes, comments, etc.)
- You need to read/edit existing documents (use `extract_docx.py` / `edit_docx.py` instead)

---

## Reading .docx

```powershell
# Full path:
python C:\Users\81004\.config\opencode\skills\docx\scripts\extract_docx.py report.docx

# Or copy first:
copy C:\Users\81004\.config\opencode\skills\docx\scripts\extract_docx.py .
python extract_docx.py report.docx
```

Outputs `report_extracted.md` with headings and tables preserved.

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

## Advanced: Raw python-docx

For cases where `docx_tool.py` presets are insufficient. **Workflow**: Write a .py script → `python script.py`.

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

# Math formula (OMML injection) — MUST use native OMML, NEVER images or plain text
# CRITICAL: oMathPara MUST be inserted at paragraph level (w:p), NOT inside w:r
omml = '''<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:oMath><m:r><m:t>E</m:t></m:r><m:r><m:t>=</m:t></m:r>
  <m:r><m:t>m</m:t></m:r><m:r><m:t>c</m:t></m:r>
  <m:sSup><m:e/><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>
</m:oMath></m:oMathPara>'''
p._element.append(etree.fromstring(omml))  # ← insert into paragraph, NOT run

doc.save('output.docx')
```

**Formula insertion rules (non-negotiable):**

| Rule | Correct | Wrong |
|------|---------|-------|
| Format | **OMML XML** (`m:oMath`, `m:oMathPara`) | Images, screenshots, MathType objects, plain text |
| Display formula parent | `w:p > m:oMathPara > m:oMath` | `w:r > m:oMathPara > m:oMath` |
| Inline formula parent | `w:r > m:oMath` | `w:r > m:oMathPara` |
| Insert method | `paragraph._element.append(omml)` for display | `run._element.append(omml)` for display formulas |

**Only OMML produces the "Equation Tools" tab in Word.** Images render visually but are not editable equations. Plain text with Unicode math characters is not a formula. If a formula must be inserted, it MUST be OMML — no exceptions.
```

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
python edit_docx.py in.docx out.docx --replace "旧文本" "新文本"
python edit_docx.py in.docx out.docx --list                        # list all paragraphs
python edit_docx.py in.docx out.docx --insert-after 3 "新段落"
python edit_docx.py in.docx out.docx --delete 5
python edit_docx.py template.docx done.docx --fill replacements.json
# replacements.json: {"姓名：": "姓名：张三", "学号：": "学号：2024001"}
```

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

**OMML is the ONLY acceptable format for formulas.** Images, plain text, MathType objects, or Unicode math characters are NOT formulas — they render but Word does not show Equation Tools and they are not editable.

**Display equation** (block-level): `w:p > m:oMathPara > m:oMath`

```xml
<w:p>
  <m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    <m:oMath>
      <m:f>                           <!-- fraction -->
        <m:fPr/>
        <m:num><m:r><m:t>a</m:t></m:r></m:num>
        <m:den><m:r><m:t>b</m:t></m:r></m:den>
      </m:f>
      <m:sSup>                        <!-- superscript -->
        <m:sSupPr/>
        <m:e><m:r><m:t>x</m:t></m:r></m:e>
        <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
      </m:sSup>
    </m:oMath>
  </m:oMathPara>
</w:p>
```

**Inline equation**: `w:r > m:oMath` (note: NO `oMathPara` wrapper)

```xml
<w:p>
  <w:r><w:t>The area is </w:t></w:r>
  <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    <m:r><m:t>π</m:t></m:r>
    <m:sSup>
      <m:sSupPr/>
      <m:e><m:r><m:t>r</m:t></m:r></m:e>
      <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
    </m:sSup>
  </m:oMath>
</w:p>
```

Element reference: `<m:f>`=fraction, `<m:sSup>`=superscript, `<m:sSub>`=subscript, `<m:rad>`=radical (√), `<m:nary>`=∑/∫/∏, `<m:r><m:t>`=text.

Unicode: `\u00B1`=±, `\u2211`=∑, `\u222B`=∫, `\u221E`=∞, `\u03C0`=π.

**Structural rules (from OOXML spec MS-OE376):**
- `m:oMathPara` MUST be a direct child of `w:p` — never inside `w:r`. Violating this causes Word to not show Equation Tools.
- `m:oMath` can be a direct child of `w:p` (inline) or inside `w:r` (inline) or inside `m:oMathPara` (display).
- `m:ctrlPr` is valid inside `m:fPr`, `m:num`, `m:den`, and other property/argument elements. Its presence in `m:num`/`m:den` (common in WPS exports) is spec-legal but unnecessary for Word rendering.

**Prohibited formats (will NOT produce editable formulas):**
- Images (.png, .jpg, .svg) of equations
- MathType embedded objects
- Plain text with Unicode math symbols (e.g., `∑`, `∫`, `²`) without OMML wrapper
- LaTeX source code rendered as text
- OLE objects or any non-OMML embedding

---

## Reference Tables

### Chinese Fonts

| Name | XML value | Name | XML value |
|------|-----------|------|-----------|
| 宋体 | `SimSun` | 仿宋 | `FangSong` |
| 黑体 | `SimHei` | 微软雅黑 | `Microsoft YaHei` |
| 楷体 | `KaiTi` | 等线 | `DengXian` |

### Style Defaults (fonts, size, color)

When setting up document styles (especially for Chinese documents), set **explicit font, size, and color** at the style level. Theme font references (`eastAsiaTheme`, `asciiTheme`, etc.) will override your explicit font and cause Word to fall back to system defaults (e.g. MS Gothic).

```python
from docx import Document
from docx.shared import Pt
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

doc = Document()
style = doc.styles['Heading 1']

# Set font properties explicitly
style.font.size = Pt(18)
style.font.name = 'Times New Roman'
style.font.color.rgb = RGBColor(0x00, 0x00, 0x00)  # black

# Also set East Asian font via XML (python-docx style.font.name only sets Latin)
rPr = style.element.find(f'{{{W}}}rPr')
if rPr is None:
    rPr = etree.SubElement(style.element, f'{{{W}}}rPr')
rFonts = rPr.find(f'{{{W}}}rFonts')
if rFonts is None:
    rFonts = etree.SubElement(rPr, f'{{{W}}}rFonts')
rFonts.set(f'{{{W}}}eastAsia', 'SimSun')          # Chinese font
rFonts.set(f'{{{W}}}ascii', 'Times New Roman')   # Latin font
rFonts.set(f'{{{W}}}hAnsi', 'Times New Roman')

# Remove theme font references so they don't override
for attr in [f'{{{W}}}eastAsiaTheme', f'{{{W}}}asciiTheme', f'{{{W}}}hAnsiTheme', f'{{{W}}}cstheme']:
    if attr in rFonts.attrib:
        del rFonts.attrib[attr]

# Set color at XML level (works for both Latin and East Asian)
color = rPr.find(f'{{{W}}}color')
if color is None:
    color = etree.SubElement(rPr, f'{{{W}}}color')
color.set(f'{{{W}}}val', '000000')  # black
```

**Key rule**: Always set `w:color w:val="000000"` in the style `rPr`. Without it, headings inherit the theme color (often blue or auto) instead of black.

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
| `mammoth` | `pip install mammoth` | Alternative: .docx→Markdown |
| `pywin32` | `pip install pywin32` | .doc→.docx via MS Word COM |
| LibreOffice | `winget install LibreOffice.LibreOffice` | .doc→.docx, PDF export |

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
  - `docx_tool.py` — **Recommended**: fluent document builder with Chinese/English presets, automatic OMML formula insertion, and theme-font cleanup
  - `extract_docx.py` — read .docx → Markdown
  - `edit_docx.py` — CLI editing (replace, insert, delete, fill)
  - `merge_runs.py` — unpack + merge adjacent runs (fix split-run at XML level)
  - `latex2omml.py` — LaTeX → OMML formula conversion (used by docx_tool.py)
