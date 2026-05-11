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

支持模糊匹配：`--section "模型"` 匹配所有标题含"模型"的章节（子串匹配，大小写不敏感）。也支持精确匹配：`--section "第三章 系统模型" --exact`。

### 4. `extract_docx.py --grep "pattern"`

**用途**：快速搜索全文，返回匹配段落索引 + 前后N段上下文。不提取全文，只返回命中的片段。

```
$ python extract_docx.py thesis.docx --grep "RFocus" --context 2

[段落42] "...采用了RFocus技术中的majority voting算法..."
  ← p40: "反向散射通信是物联网领域的..."
  → p43: "该算法通过迭代优化实现..."
  → p44: "仿真结果表明..."

[段落78] "...RFocus系统由3000根天线..."
  ← p76: "天线阵列的配置需要考虑..."
  → p79: "每根天线的反射状态..."

[段落156] "...与传统方案相比，RFocus的优势..."
```

**token消耗**：3个匹配位 + 上下文 = ~500 tokens（vs 全文30K+）。

### 5. `edit_docx.py --paragraph N "new text"`

**用途**：替换指定段落的内容（保留其格式风格）。

```
$ python edit_docx.py thesis.docx output.docx --paragraph 42 "新的段落内容"
```

支持：
- `--paragraph N --file content.md` 从文件读取新内容（长文本场景）
- `--paragraph N --append "追加内容"` 在段落末尾追加
- `--paragraph N --prepend "前缀内容"` 在段落开头插入

### 6. `edit_docx.py --range N-M --replace "old" "new"`

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
| `--grep "pattern"` 搜索 | `extract_docx.py` | +20行 |
| `--paragraph N` 定点编辑（含 `--file`/`--append`/`--prepend`） | `edit_docx.py` | +40行 |
| `--range --replace` 范围替换 | `edit_docx.py` | +10行 |

总计约140行，不改动现有逻辑。

## QA 验证场景

完成实现后，每个功能需通过以下验证：

### `--structure`
```
$ python extract_docx.py thesis_sjtu_v2.docx --structure
```
**预期**：输出树形结构，每个 Heading 节点显示段落范围，根节点显示总段落数和表格数。无正文内容。

### `--range`
```
$ python extract_docx.py thesis_sjtu_v2.docx --range 1-5
```
**预期**：仅输出段落1到5的内容。对无 style 的空段落，按 index 正常返回。

### `--section`
```
$ python extract_docx.py thesis_sjtu_v2.docx --section "系统模型"
```
**预期**：匹配标题含"系统模型"的章节及所有子节（"1.1 系统模型"匹配）。`--section "不存在的标题"` 应报清晰错误。

### `--grep`
```
$ python extract_docx.py thesis_sjtu_v2.docx --grep "天线" --context 1
```
**预期**：返回所有含"天线"的段落索引 + 前后各1段。不存在时输出"No matches"。

### `--paragraph`
```
$ python edit_docx.py thesis_sjtu_v2.docx out.docx --paragraph 0 "新标题"
```
**预期**：段落0内容替换为"新标题"，其他段落不变，格式保持。`--paragraph 999` 应报段落不存在。
$ python edit_docx.py thesis_sjtu_v2.docx out.docx --paragraph 3 --append "（补）"
```
**预期**：段落3末尾追加"（补）"。

### `--range --replace`
```
$ python edit_docx.py thesis_sjtu_v2.docx out.docx --range 2-5 --replace "天线" "antenna"
```
**预期**：仅段落2-5中的"天线"被替换，段落1及其他保持不变。
