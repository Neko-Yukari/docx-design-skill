# DOCX Design Skill

> 读写、创建、编辑 Word 文档，支持 LaTeX 公式、中文排版与模块化主题设计。

这是一个独立的 Python 工具包，用于程序化生成 Word 文档。既可作为 OpenCode 的 Skill 使用，也可作为 Cursor / Claude Code 的上下文模块或普通 Python 库引入。

---

## ✨ 功能概览

| 类别 | 能力 |
|------|------|
| **读取** | 将 .docx 中的文字、表格提取为 Markdown |
| **创建** | 从零生成 .docx：标题、表格、图片、公式 |
| **编辑** | 查找替换、插入/删除段落、模板填充 |
| **公式** | 写入 LaTeX → 自动转换为原生 OMML 公式 |
| **主题** | 模块化格式规范（上海交大学位论文、GB/T 标准，可扩展） |
| **样式系统** | 覆盖 Word 内置样式（Normal、Heading 1/2/3），支持导航窗格与目录 |
| **中文排版** | 通过 `eastAsia` 属性完整支持 CJK 字体（宋体、黑体、楷体等） |
| **.doc 兼容** | 将旧版 .doc 转换为 .docx（Windows 用 MS Word COM；跨平台用 LibreOffice） |

---

## 📦 安装

### AI Agent 集成

| 平台 | 安装路径 |
|------|----------|
| **OpenCode** | 复制到 `~/.config/opencode/skills/docx/` |
| **Claude Code** | 复制到 `~/.claude/skills/docx/` |
| **Cursor / Copilot** | 作为上下文文件或项目模块引入 |

### Python 环境

```powershell
pip install python-docx lxml mammoth latex2mathml
```

| 包 | 必需 | 用途 |
|----|------|------|
| `python-docx` | ✅ | 核心读写/创建/编辑 |
| `lxml` | ✅ | OOXML/OMML XML 操作 |
| `latex2mathml` | ✅ | LaTeX 公式解析 |
| `mammoth` | ⚠️ | 可选的 .docx→Markdown 替代方案 |

### 可选：.doc 支持

| 平台 | 方式 | 安装 |
|------|------|------|
| Windows | MS Word COM | `pip install pywin32` |
| macOS/Linux | LibreOffice | `brew install libreoffice` / `apt install libreoffice` |

---

## 🚀 快速上手

### 作为 Python 库

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

### 命令行使用

```powershell
# 读取文档（全文）
python scripts/extract_docx.py report.docx

# 获取文档大纲（仅结构，约 300 tokens）
python scripts/extract_docx.py thesis.docx --structure

# 按章节名称提取
python scripts/extract_docx.py thesis.docx --section "结果分析"

# 在文档中搜索关键词
python scripts/extract_docx.py thesis.docx --grep "关键词" --context 2

# 编辑文档
python scripts/edit_docx.py template.docx output.docx --replace "旧文字" "新文字"

# 编辑指定段落
python scripts/edit_docx.py in.docx out.docx --paragraph 3 "新的段落内容"

# 转换 LaTeX 为 OMML
python scripts/latex2omml.py "x^2 + y^2 = z^2"
```

---

## 📐 主题系统

格式的每个方面都由 `.md` 主题文件定义——没有硬编码的默认值。

### 内置主题

| 主题文件 | 说明 |
|----------|------|
| `format-spec/thesis-sjtu.md` | 上海交通大学学位论文 |
| `format-spec/thesis-generic.md` | GB/T 7713 通用学位论文 |

### 主题结构

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

详细 Schema 见 `format-spec/THEME-SCHEMA.md`（已使用中文编写），其中定义了所有必填及可选字段。复制并修改现有主题即可创建自定义主题。

---

## 🛠️ 模块说明

| 模块 | 用途 |
|------|------|
| `extract_docx.py` | 读取 .docx → Markdown（全文、大纲结构、段落范围、章节提取、关键词搜索） |
| `edit_docx.py` | CLI 编辑（替换、插入、删除、段落操作、范围限定编辑、模板填充） |
| `apply_theme.py` | 解析主题 → 覆盖 Word 样式 → 生成内容 |
| `latex2omml.py` | LaTeX → MathML → OMML（稳健的公式嵌入） |
| `convert_to_doc.py` | 将 .docx 转换为旧版 .doc 格式 |
| `merge_runs.py` | XML 级别 run 合并（修复 Word 的分段渲染问题） |

---

## 🔧 兼容性

- **Python**：3.9+
- **操作系统**：Windows · macOS · Linux
- **Word 渲染**：Microsoft Word 或 LibreOffice
- **.doc 转换**：Windows 用 MS Word COM · macOS/Linux 用 LibreOffice

---

## 📄 许可证

MIT — 详见 [LICENSE](LICENSE)。

---

## 🙏 架构参考

本文档的设计模式借鉴了 Anthropic 的 [docx skill](https://github.com/anthropics/skills)（OOXML 工作流、merge_runs、XML 参考）和 OpenAI 的 [doc skill](https://github.com/openai/skills)（可视化渲染循环、质量标准）。在保证 Windows 实用性的基础上兼顾跨平台兼容。
