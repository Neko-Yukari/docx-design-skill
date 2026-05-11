# DOCX Design Skill

> Read, create, and edit Word documents with LaTeX formulas, Chinese typesetting, and modular design themes.

A standalone Python toolkit for programmatic Word document generation. Works as an OpenCode skill, a Cursor/Claude Code context module, or a plain Python library.

## ✨ Features

| Category | Capability |
|----------|-----------|
| **Read** | Extract text/tables from .docx to Markdown |
| **Create** | Generate .docx from scratch with headings, tables, images, formulas |
| **Edit** | Find-replace, insert/delete paragraphs, fill templates |
| **Formulas** | Write LaTeX → auto-convert to native OMML equations |
| **Themes** | Modular format specs (SJTU thesis, GB/T standard, extensible) |
| **Style System** | Override Word built-in styles (Normal, Heading 1/2/3) for proper nav pane & TOC |
| **Chinese** | Full CJK font support via `eastAsia` attribute (SimSun, SimHei, KaiTi, etc.) |
| **.doc** | Convert legacy .doc to .docx (Windows: MS Word COM; cross-platform: LibreOffice) |

## 📦 Installation

### AI Agent Integration

**OpenCode**: copy to `~/.config/opencode/skills/docx/`

**Claude Code**: copy to `~/.claude/skills/docx/`

**Cursor / Copilot**: add as context file or project module

### Python Environment

```powershell
pip install python-docx lxml mammoth latex2mathml
```

| Package | Required | Purpose |
|---------|----------|---------|
| `python-docx` | ✅ | Core read/create/edit |
| `lxml` | ✅ | OOXML/OMML XML manipulation |
| `latex2mathml` | ✅ | LaTeX formula parsing |
| `mammoth` | ⚠️ | Alternative .docx→Markdown |

### Optional: .doc Support

| Platform | Method | Install |
|----------|--------|---------|
| Windows | MS Word COM | `pip install pywin32` |
| macOS/Linux | LibreOffice | `brew install libreoffice` / `apt install libreoffice` |

## 🚀 Quick Start

### As Python Library

```python
import sys
sys.path.insert(0, '/path/to/docx-design-skill/scripts')

from docx import Document
from apply_theme import load_theme, apply_theme, themed

theme = load_theme('format-spec/thesis-sjtu.md', degree='硕士')
doc = Document()
apply_theme(doc, theme)
t = themed(doc, theme)

t.heading('第一章 引言', 1)
t.body('这是正文。')
t.formula(r'E = mc^2')
doc.save('output.docx')
```

### Command Line

```powershell
# Read a document (full text)
python scripts/extract_docx.py report.docx

# Get document outline (structure only, ~300 tokens)
python scripts/extract_docx.py thesis.docx --structure

# Extract a specific section
python scripts/extract_docx.py thesis.docx --section "Results"

# Search within a document
python scripts/extract_docx.py thesis.docx --grep "keyword" --context 2

# Edit a document
python scripts/edit_docx.py template.docx output.docx --replace "旧" "新"

# Edit a specific paragraph
python scripts/edit_docx.py in.docx out.docx --paragraph 3 "New paragraph text"

# Convert LaTeX to OMML
python scripts/latex2omml.py "x^2 + y^2 = z^2"
```

## 📐 Theme System

Every aspect of formatting is defined in `.md` theme files — no hardcoded defaults.

### Built-in themes

| Theme | Description |
|-------|-------------|
| `format-spec/thesis-sjtu.md` | 上海交通大学学位论文 |
| `format-spec/thesis-generic.md` | GB/T 7713 通用学位论文 |

### Theme structure

```ini
[document]
page_width  = 21cm
margin_left = 3.0cm

[heading1]
font  = SimHei
size  = 16pt
color = 000000
bold  = true
align = center
```

Schema: `format-spec/THEME-SCHEMA.md` defines all required and optional fields. Create custom themes by copying and modifying an existing one.

## 🛠️ Modules

| Module | Purpose |
|--------|---------|
| `extract_docx.py` | Read .docx → Markdown (full, structure outline, range, section, grep search) |
| `edit_docx.py` | CLI editing (replace, insert, delete, paragraph ops, range-scoped edits, fill templates) |
| `apply_theme.py` | Parse theme → override Word styles → generate content |
| `latex2omml.py` | LaTeX → MathML → OMML (robust formula embedding) |
| `convert_to_doc.py` | Convert .docx to legacy .doc format |
| `merge_runs.py` | XML-level run merging (fixes Word's split-run rendering) |

## 🔧 Compatibility

- **Python**: 3.9+
- **OS**: Windows • macOS • Linux
- **Word Rendering**: Microsoft Word or LibreOffice
- **.doc Conversion**: Windows: MS Word COM • macOS/Linux: LibreOffice

## 📄 License

MIT — see [LICENSE](LICENSE).

## 🙏 Architecture

Design patterns adopted from Anthropic's [docx skill](https://github.com/anthropics/skills) (OOXML workflows, merge_runs, XML references) and OpenAI's [doc skill](https://github.com/openai/skills) (visual render loop, quality specifications). Built for practical use on Windows with cross-platform compatibility.