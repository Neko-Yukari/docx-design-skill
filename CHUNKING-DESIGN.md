# 文档分段操作能力 — 接口设计

## 问题

当前所有工具都是"全量操作"：提取全文、编辑全文。大文档（>30页）会导致AI上下文膨胀，且无法聚焦到具体章节/段落。

## 目标

让AI能像操作代码一样操作文档——先看结构（outline），定位到要改的部分（range/section），再定点编辑（paragraph）。

## 新增接口

### 1. `extract_docx.py --structure`

**用途**：输出文档大纲（标题层级 + 每节段数），不输出正文。给AI做"地图"。

```
$ python extract_docx.py thesis.docx --structure

[document]  thesis.docx  (186 paragraphs, 12 tables)
├── Heading 1 (p0):  第一章 引言               [段落: p0-p15,  表格: 1]
│   ├── Heading 2 (p1): 1.1 研究背景            [段落: p1-p5]
│   ├── Heading 2 (p6): 1.2 国内外现状           [段落: p6-p12]
│   └── Heading 2 (p13): 1.3 本文工作            [段落: p13-p15]
├── Heading 1 (p16): 第二章 系统模型              [段落: p16-p45, 表格: 2]
│   ├── ...
...
```

**token消耗**：~100-300 tokens（只看结构，不看正文）。

### 2. `extract_docx.py --range N-M`

**用途**：只提取指定段落范围。

```
$ python extract_docx.py thesis.docx --range 16-45

# 第二章 系统模型 (段落 16-45)
[提取的正文内容，含表格]
```

### 3. `extract_docx.py --section "关键词"`

**用途**：按章节名匹配提取。精确匹配Heading文本。

```
$ python extract_docx.py thesis.docx --section "系统模型"
# 匹配 "第二章 系统模型" 或其子节
[提取该节及所有子节的完整内容]
```

支持模糊匹配：`--section "模型"` 匹配所有标题含"模型"的章节。

### 4. `edit_docx.py --paragraph N "new text"`

**用途**：替换指定段落的内容（保留其格式风格）。

```
$ python edit_docx.py thesis.docx output.docx --paragraph 42 "新的段落内容"
```

支持：
- `--paragraph N --file content.md` 从文件读取新内容（长文本场景）
- `--paragraph N --append "追加内容"` 在段落末尾追加
- `--paragraph N --prepend "前缀内容"` 在段落开头插入

### 5. `edit_docx.py --range N-M --replace "old" "new"`

**用途**：仅在指定段落范围内做替换。

```
$ python edit_docx.py thesis.docx output.docx --range 16-45 --replace "5G" "6G"
```

## AI 工作流示例

### 场景：修改硕士论文第三章的数据表格

```
Step 1: 看结构 (token: ~100)
  python extract_docx.py thesis.docx --structure
  → 发现第三章在段落90-140

Step 2: 提取目标段 (token: ~3,000)
  python extract_docx.py thesis.docx --range 90-140
  → 只看第三章内容

Step 3: AI分析后定点编辑 (token: ~500)
  python edit_docx.py thesis.docx thesis_v2.docx --paragraph 105 "更新后的段落"

Step 4: 验证
  python extract_docx.py thesis_v2.docx --range 104-106
  → 确认修改正确
```

全流程token消耗：~4,000（vs 全量提取的15,000+）。

## 实现清单

| 任务 | 改动文件 | 行数估计 |
|------|----------|----------|
| `--structure` 模式 | `extract_docx.py` | +30行 |
| `--range N-M` 模式 | `extract_docx.py` | +15行 |
| `--section "name"` 模式 | `extract_docx.py` | +25行 |
| `--paragraph N` 定点编辑 | `edit_docx.py` | +30行 |
| `--range --replace` 范围替换 | `edit_docx.py` | +10行 |

总计约110行，不改动现有逻辑。
