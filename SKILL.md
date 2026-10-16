---
name: docx
description: Use environment-installed Python dependencies to read, create, and edit Word documents (.docx and .doc) on any platform. Supports LaTeX formulas, Chinese typesetting, modular themes, and Word style integration. Use when working with .docx/.doc files — reading, creating, editing, filling templates, generating thesis papers, or converting LaTeX to Word equations.
license: MIT
---

# DOCX / DOC Skill

Read, create, edit Word documents on Windows using `python-docx`.

---

## Critical Rule (Windows)

**Use CLI for simple tasks, short Python scripts for complex generation.**

- **CLI**: quick edits, simple create (< 5 paragraphs), extract, search
- **Python script**: complex documents with multiple sections, tables, images, formulas (keep scripts ≤ 40 lines, import from `docx_tool.py`)
- **NEVER run Python inline in PowerShell** — f-strings, Chinese chars, and nested quotes break inline execution

---

## Quick Reference

**One tool does everything**: `python docx_tool.py <command> ...`

| Task | CLI | Python API |
|------|-----|------------|
| **Install dependencies** | `python -m pip install -r requirements.txt` | Use the scripts’ Python environment |
| **Create** new .docx | `docx_tool.py create --output out.docx --preset chinese ...` | `DocxBuilder().preset_chinese().add_heading(...).save(...)` |
| **Read/Extract** .docx | `docx_tool.py extract file.docx --output file.md` | `DocxReader(file).extract_to_markdown(...)` |
| **Search** in .docx | `docx_tool.py extract file.docx --grep "keyword"` | `DocxReader(file).grep(...)` |
| **Edit** .docx | `docx_tool.py edit in.docx out.docx --replace "old" "new"` | `DocxEditor(in_file).replace(...).save(out_file)` |
| **Merge runs** (fix split-run) | `docx_tool.py merge-runs in.docx out.docx` | `DocxUtil.merge_runs(in_file, out_file)` |
| **Convert** .docx → .doc | `docx_tool.py convert-to-doc in.docx out.doc` | `DocxUtil.convert_to_doc(in_file, out_file)` |
| **Apply theme** | `docx_tool.py apply-theme in.docx theme.yaml out.docx` | `DocxUtil.apply_theme(in_file, theme, out_file)` |
| **Set fonts** (existing doc → 宋体/黑体) | `docx_tool.py set-fonts in.docx out.docx --body 宋体 --heading 黑体` | `DocxUtil.set_fonts(in_file, out_file, body_font="宋体", heading_font="黑体")` |
| Read .doc text | Convert to .docx first, then extract | [.doc Handling](#doc-handling) |
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
- **Auto heading numbering** — uses Word built-in multilevel list (1, 1.1, 1.1.1).
  Call `add_heading("引言", level=1)` **without** the digit in the text.
  Setting ``heading_numbering=False`` disables this.

### Python API — for complex documents

**Rule: CLI for simple edits, Python API for complex generation.**

| Scenario | Method | Example |
|----------|--------|---------|
| **Quick edits** (replace, insert, delete, extract) | **CLI** | `python docx_tool.py edit ...` |
| **Simple create** (< 5 paragraphs, no tables/images) | **CLI** | `python docx_tool.py create ...` |
| **Complex documents** (multiple sections, tables, images, formulas, loops) | **Python API** | 20-40 line script importing `DocxBuilder` |

**When using Python API, write SHORT scripts (≤ 40 lines)** that import from `docx_tool.py`. Do NOT reinvent formula insertion or style setup — use `DocxBuilder().preset_chinese()`.

```python
import sys
sys.path.insert(0, r"C:\Users\81004\.config\opencode\skills\docx\scripts")
from docx_tool import DocxBuilder
from docx.shared import Inches

builder = DocxBuilder().preset_chinese()
builder.add_heading("项目报告", level=1)
builder.add_heading("概述", level=2)            # ← NO digit in text; Word auto-numbers
builder.add_paragraph("本文使用黑体作为正文字体。")
builder.add_formula(r"x = \frac{1}{2}", display=True)
builder.add_table(rows=3, cols=2, data=[['A', 'B'], ['1', '2']])
builder.add_image('chart.png', width=Inches(4))
builder.save("report.docx")
```

**Built-in heading numbering** (since 2026-06-22):
- `preset_chinese()` now enables Word multilevel list by default:
  - Level 1 → `1`、`2`、…  Level 2 → `1.1`、`1.2`、…  Level 3 → `1.1.1`、`1.1.2`、…
- `add_heading()` text should **not** include the digit — Word inserts it.
- To disable: `DocxBuilder().preset_chinese(heading_numbering=False)`.
- To set manually: `builder.set_heading_numbering(levels=3)`.

**Key difference from raw `python-docx`:**
- No manual `eastAsia` font XML manipulation
- No theme reference cleanup
- No OMML structure foot-guns (`add_formula()` handles display vs inline automatically)
- Preset defaults eliminate repetitive setup

---

## Unified Workflow (docx_tool.py)

All operations go through `docx_tool.py`. **CLI for quick tasks, short scripts (≤ 40 lines) for complex documents.**

### Create

```powershell
# CLI
python docx_tool.py create --output report.docx --preset chinese --title "报告" --add-heading "1:概述" --add-paragraph "正文" --add-formula "x = \frac{1}{2}"

# Python API
from docx_tool import DocxBuilder
DocxBuilder().preset_chinese().add_heading("报告", level=1).add_paragraph("正文").save("report.docx")
```

### Read / Extract

```powershell
# Full extraction to Markdown
python docx_tool.py extract report.docx --output report.md

# Document structure (heading tree)
python docx_tool.py extract report.docx --structure

# Grep search with context
python docx_tool.py extract report.docx --grep "关键词" --context 2

# Extract paragraph range
python docx_tool.py extract report.docx --range 10-25 --output section.md

# Extract by section heading
python docx_tool.py extract report.docx --section "Results" --output results.md
```

### Edit

```powershell
# Replace text
python docx_tool.py edit in.docx out.docx --replace "old" "new"

# Insert paragraph after index N
python docx_tool.py edit in.docx out.docx --insert-after 3 "新段落"

# Delete paragraph
python docx_tool.py edit in.docx out.docx --delete 5

# Set paragraph style
python docx_tool.py edit in.docx out.docx --set-style 2 "Heading 1"

# Replace entire paragraph
python docx_tool.py edit in.docx out.docx --paragraph 2 "新内容"

# Fill template from JSON
python docx_tool.py edit template.docx out.docx --fill replacements.json
```

### Utils

```powershell
# Merge adjacent runs (fix split-run bug)
python docx_tool.py merge-runs in.docx out.docx

# Convert .docx → .doc
python docx_tool.py convert-to-doc in.docx out.doc

# Apply format theme
python docx_tool.py apply-theme in.docx theme.yaml out.docx

# Retarget fonts on an existing document (Chinese thesis: 宋体 body + 黑体 headings)
python docx_tool.py set-fonts in.docx out.docx --body 宋体 --heading 黑体
```

### set-fonts vs apply-theme

`apply-theme` builds a **thesis-shaped** document (cover / abstract / TOC / acknowledgements are mandatory sections), so it is the wrong tool when the document already has content and only its typography must change.

Use **`set-fonts`** for that case. It rewrites `docDefaults` plus every style in `styles.xml` and `stylesWithEffects.xml`, matching heading styles by style-id/name (`Heading*`, `Title`, `Subtitle`) and treating everything else as body. It also:

- removes `asciiTheme` / `eastAsiaTheme` / … — otherwise Word keeps resolving the theme font and silently ignores the explicit one;
- rewrites run-level fonts in `document.xml` while keeping monospace runs (Consolas, Courier…) monospaced — pass `--styles-only` to leave runs untouched;
- defaults to `--body 宋体 --heading 黑体 --body-latin "Times New Roman" --heading-latin "Times New Roman"`.

```python
from docx_tool import DocxUtil
DocxUtil.set_fonts("in.docx", "out.docx", body_font="宋体", heading_font="黑体")
```

**Font name tip (Windows):** Chinese fonts are installed under localized family names, so `InstalledFontCollection` returns `宋体` / `黑体` rather than `SimSun` / `SimHei`. Both spellings resolve in Word; prefer the localized names when targeting a Chinese-locale machine. Word substitutes silently when a font is missing (e.g. `Noto Sans CJK SC`), which is what produces garbled or missing glyphs.

### Prefer a curated template over cloning someone else's file

Cloning a real document — load it with `python-docx`, delete every body element **except** `w:sectPr`, write new content with its styles — reproduces its typography exactly, but it also reproduces its **defects**. You then have to discover them after generating (clipped titles, bloated files) and patch the output. Only do this when you genuinely must match a third-party file.

For your own deliverables, start from a curated template instead. `templates/chinese-report.docx` ships with this skill and already fixes the common defects:

- body 宋体 + headings 黑体 + Times New Roman for Latin, set in `docDefaults` and every style, with theme-font references removed;
- **all line spacing as a multiple** (`w:lineRule="auto"`), so a larger font never gets clipped;
- page size, margins, heading sizes and table style defined once;
- no embedded media, so outputs stay small.

```python
from docx import Document
TEMPLATE = r"...\skills\docx\templates\chinese-report.docx"
doc = Document(TEMPLATE)          # styles are already correct
# ... keep w:sectPr, clear the body, write content ...
doc.save("out.docx")              # no post-processing needed
```

Regenerate the template with `scripts/`-adjacent tooling only when you want to change the shared look; never hand-patch individual outputs.

If you do have to clone, three things break:

- **Never remove `w:sectPr`.** It must stay as the last body child: python-docx inserts new paragraphs/tables before it, and `Document.add_table` reads it for the text-column width. Removing it raises `IndexError` in `document.py:_block_width`.
- **Unreferenced media is carried over.** Images embedded in the template stay in the package even if the new body never references them, so each output silently gains megabytes. Drop every `word/media/*` whose relationship id (`r:embed` / `r:link`) does not appear in `document.xml`, and remove those `<Relationship>` entries in `word/_rels/document.xml.rels`.
- **Fixed line spacing clips anything taller than the line.** `w:lineRule="exact"` is not only an image hazard — inline pictures get cut to a sliver, and so does any larger text: a 22 pt title or a 15 pt heading under a 16 pt exact line loses its top half. Use multiple/auto spacing on headings so the value scales with the font size, and override spacing on image paragraphs (`w:line="240" w:lineRule="auto"`, plus `w:keepNext` and centered alignment).

---

## Advanced: Python API (for complex documents only)

**For CLI-insufficient cases: multi-section documents, tables, images, formula loops. Keep scripts ≤ 40 lines.**

### DocxBuilder — create documents

```python
from docx_tool import DocxBuilder
from docx.shared import Inches

# Custom preset beyond built-in chinese/english
builder = DocxBuilder()
builder.doc.styles['Normal'].font.size = Pt(11)
builder.add_heading('Custom Report', level=1)
builder.add_paragraph('Body text with custom formatting.')
builder.add_table(rows=3, cols=2, data=[['A', 'B'], ['1', '2'], ['3', '4']])
builder.add_image('chart.png', width=Inches(4))
builder.save('output.docx')
```

#### set_heading_numbering(levels=3)

Enables Word built-in multilevel list numbering for headings (1, 1.1, 1.1.1).
`preset_chinese()` calls this by default.

```python
builder = DocxBuilder().preset_chinese(heading_numbering=True)  # default
# or set manually:
builder.set_heading_numbering(levels=3)

# Now add_heading text should NOT include digits
builder.add_heading("引言", level=1)      # → "1  引言"
builder.add_heading("研究背景", level=2)   # → "1.1  研究背景"
builder.add_heading("数据来源", level=2)   # → "1.2  数据来源"
```

Disable: `DocxBuilder().preset_chinese(heading_numbering=False)`.

#### add_numbered_list(items, *, format="decimal", lvl_text=None, start=1, hanging=360, left_indent=720)

Adds a Word auto-numbered (or bulleted) list. **Do NOT** type digits or bullet characters into the text — Word inserts them.

Supports **flat lists** (all same level) and **multi-level lists** via `(text, level)` tuples.

```python
# Flat list with default "1." format
builder.add_numbered_list([
    "在人均消费维度，美国……",
    "在用药质量维度，中国……",
])

# Chinese paren format: （1）（2）（3）
builder.add_numbered_list([
    "第一条结论",
    "第二条结论",
], format="decimalParen")

# Multi-level: 1 → 1.1 → 1.1.1
builder.add_numbered_list([
    ("一级条目", 0),
    ("  二级子条目", 1),
    ("    三级子条目", 2),
    ("一级条目", 0),
], format="decimal")

# Chinese counting: 一、二、三
builder.add_numbered_list([
    "引言部分", "方法部分", "结果部分",
], format="chineseCounting")

# Bullet list (convenience method)
builder.add_bullet_list(["项目A", "项目B", "项目C"])
```

| `format` key         | numFmt             | lvlText            | Renders as          |
| -------------------- | ------------------ | ------------------ | ------------------- |
| `"decimal"`            | `decimal`            | `%1.`                | 1. 2. 3.            |
| `"decimalParen"`       | `decimal`            | `（%1）`             | （1）（2）（3）      |
| `"decimalParen2"`      | `decimal`            | `%1)`                | 1) 2) 3)            |
| `"chineseCounting"`    | `chineseCounting`    | `一、`               | 一、二、三、        |
| `"chineseParen"`       | `chineseCounting`    | `（%1）`             | （一）（二）（三）  |
| `"bullet"`             | `bullet`             | `\u2022`             | • • •               |
| `"upperRoman"`         | `upperRoman`         | `%1.`                | I. II. III.         |
| `"upperLetter"`        | `upperLetter`        | `%1.`                | A. B. C.            |
| `"lowerLetter"`        | `lowerLetter`        | `%1.`                | a. b. c.            |

Pass a raw `numFmt` value (e.g. `"chineseLegalSimplified"`) to `format` for formats not in the table.
Pass `lvl_text="第%1条"` to override the default template string.
Multi-level lists auto-generate `%1.` / `%1.%2.` / `%1.%2.%3.` unless `lvl_text` is provided.

#### add_bullet_list(items)

Convenience: equivalent to `add_numbered_list(items, format="bullet")`.

### DocxEditor — edit existing documents

```python
from docx_tool import DocxEditor

editor = DocxEditor('input.docx')
editor.replace('old_text', 'new_text')           # global replace
editor.insert_after(5, 'Inserted paragraph.')    # insert after paragraph 5
editor.delete(3)                                 # delete paragraph 3
editor.set_style(2, 'Heading 2')                 # change style
editor.save('output.docx')
```

### DocxReader — extract and search

```python
from docx_tool import DocxReader

reader = DocxReader('report.docx')
reader.extract_to_markdown('report.md')          # full extraction
reader.print_structure()                          # heading tree
matches = reader.grep('keyword', context=2)       # search with context
reader.extract_range(10, 25, 'section.md')       # paragraph range
reader.extract_section('Results', 'results.md')  # by heading name
```

### DocxUtil — utilities

```python
from docx_tool import DocxUtil

# Fix split-run bug
DocxUtil.merge_runs('input.docx', 'output.docx')

# Convert .docx → .doc
DocxUtil.convert_to_doc('input.docx', 'output.doc')

# Apply theme
DocxUtil.apply_theme('input.docx', 'theme.yaml', 'output.docx')
```

### Raw python-docx (DO NOT USE — reference only)

**Agents: Never write raw python-docx code.** For features beyond `docx_tool.py` (tracked changes, comments), see [OOXML Reference](#ooxml-reference).

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

## Advanced: Raw python-docx (Reference Only)

For cases where `docx_tool.py` presets are insufficient. Use `DocxBuilder` / `DocxEditor` first; drop down to raw `python-docx` only for:
- Fine-grained run-level formatting (per-character fonts/colors)
- Custom styles beyond Heading 1-3 + Normal
- Direct OOXML manipulation (tracked changes, comments, etc.)

**Formula insertion rules (non-negotiable):**

| Rule | Correct | Wrong |
|------|---------|-------|
| Format | **OMML XML** (`m:oMath`, `m:oMathPara`) | Images, screenshots, MathType objects, plain text |
| Display formula parent | `w:p > m:oMathPara > m:oMath` | `w:r > m:oMathPara > m:oMath` |
| Inline formula parent | `w:r > m:oMath` | `w:r > m:oMathPara` |
| Insert method | `paragraph._element.append(omml)` for display | `run._element.append(omml)` for display formulas |

**Only OMML produces the "Equation Tools" tab in Word.** Images render visually but are not editable equations. Plain text with Unicode math characters is not a formula. If a formula must be inserted, it MUST be OMML — no exceptions.

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

**Binomial coefficients**: `\binom{a}{b}` arrives from latex2mathml as `<mfrac linethickness="0">`; the converter maps it to `<m:f><m:fPr><m:type m:val="noBar"/>` — a stacked layout with **no fraction bar**. Ordinary `\frac` keeps the default bar. Inline conversion produces exactly one `m:oMath` (never nested `oMath` inside `oMath` — Word rejects that as corruption).

**N-ary operators (∫ ∑ ∏ ∮)**: the integrand/summand is attached as the `m:nary` operand (`m:e`), eliminating a detached gap between the operator and its operand. Boundary policy is conservative and all-or-nothing:
- `\int/\oint`: consumes up to and including the first top-level differential (`d<var>`, or a braced group containing one). If a relation operator (`=`, `<`, `,` …) or the row ends first, nothing is attached.
- `\sum/\prod`: consumes exactly one following sibling term — `= C` and adjacent terms are never swallowed.
- No-limit `\int x\,dx`: promoted to a real `m:nary` with empty limits.
- Nested operators bind right-to-left, so inner integrals keep their own operands.
- Known limitation: an integral whose integrand has no explicit `d<var>` differential (e.g. `\int_0^1 f(x) = C`) intentionally keeps the old detached rendering rather than guessing the operand extent.

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

Install dependencies in the Python environment that runs the scripts. Copied `site-packages/` directories and platform-specific launchers are not used.

| Package | Location | Install once? | Purpose |
|---------|----------|---------------|---------|
| `python-docx` | Selected Python environment | Yes | Core: read, create, edit .docx |
| `latex2mathml` | Selected Python environment | Yes | LaTeX → MathML conversion |
| `lxml` | System Python | **Yes** (`pip install lxml`) | XML parsing for OOXML/OMML |
| `mammoth` | System Python | Optional | Alternative: .docx→Markdown |
| `pywin32` | System Python | Optional | .doc→.docx via MS Word COM |

**One-time setup** (run from the skill directory):
```bash
python -m pip install -r requirements.txt
```

Run all scripts with this same Python interpreter. A virtual environment may be used; do not copy its installed packages into the skill.

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
- **Unified tool** (in `~/.config/opencode/skills/docx/scripts/`):
  - `docx_tool.py` — **One tool for everything**: create, read, edit, merge-runs, convert, apply-theme. Self-contained, no one-off scripts needed.
