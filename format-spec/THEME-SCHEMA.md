# Theme Schema — 格式主题的结构约束

主题文件是**可选**的格式预定义。不指定主题时，采用自由模式：AI 读取现有文档的格式风格，自行延续。

当指定主题时，文件需包含以下必填 section。缺失必填 section/field 的主题文件无效。

---

## SECTION: [document] — 必填

| Field | 类型 | 说明 | 示例 |
|-------|------|------|------|
| page_width | 尺寸 | 页面宽度 | `21cm` |
| page_height | 尺寸 | 页面高度 | `29.7cm` |
| margin_top | 尺寸 | 上边距 | `2.5cm` |
| margin_bottom | 尺寸 | 下边距 | `2.5cm` |
| margin_left | 尺寸 | 左边距（装订边） | `3.0cm` |
| margin_right | 尺寸 | 右边距 | `2.5cm` |

## SECTION: [cover] — 必填

| Field | 类型 | 说明 |
|-------|------|------|
| title_font | 字体名 | 标题字体 |
| title_size | 字号 | 标题字号 |
| title_color | hex颜色 | 如 `000000`（黑色）。**注意：所有有字体的section必须指定color，不设默认值** |
| title_bold | bool | 是否加粗 |
| title_align | left/center/right | 对齐 |
| info_font | 字体名 | 作者信息字体 |
| info_size | 字号 | 作者信息字号 |
| info_color | hex颜色 | |
| info_align | left/center/right | 对齐 |

## SECTION: [abstract] — 必填

| Field | 类型 | 说明 |
|-------|------|------|
| heading_font | 字体名 | |
| heading_size | 字号 | |
| heading_bold | bool | |
| heading_align | left/center/right | |
| heading_text | 字符串 | 标题文字，如 `摘要` |
| body_font | 字体名 | |
| body_size | 字号 | |
| body_line_spacing | float | 行距倍数 |
| keywords_label | 字符串 | 如 `关键词：` |
| keywords_label_font | 字体名 | |
| keywords_font | 字体名 | |

## SECTION: [heading1] / [heading2] / [heading3] — 必填（全部三个）

| Field | 类型 | 说明 |
|-------|------|------|
| font | 字体名 | |
| size | 字号 | |
| color | hex颜色 | 必填。如 `000000` |
| bold | bool | |
| align | left/center/right | |
| before_spacing | 尺寸 | 段前间距 |
| after_spacing | 尺寸 | 段后间距 |

## SECTION: [body] — 必填

| Field | 类型 | 说明 |
|-------|------|------|
| font | 字体名 | |
| size | 字号 | |
| color | hex颜色 | 必填。如 `000000` |
| line_spacing | float | 行距倍数 |
| first_line_indent | 尺寸 | 首行缩进，如 `0.74cm` |
| before_spacing | 尺寸 | |
| after_spacing | 尺寸 | |

## SECTION: [header] — 必填

| Field | 类型 | 说明 |
|-------|------|------|
| font | 字体名 | |
| size | 字号 | |
| text | 字符串 | 页眉文字，可用 `{degree}` 占位，如 `{degree}学位论文` |

## SECTION: [footer] — 必填

| Field | 类型 | 说明 |
|-------|------|------|
| font | 字体名 | |
| size | 字号 | |
| page_number_align | left/center/right | 页码对齐 |

---

## 可选 Section

### [toc] — 目录格式

| Field | 类型 | 说明 |
|-------|------|------|
| heading_font | 字体名 | |
| heading_size | 字号 | |
| heading_bold | bool | |
| heading_align | left/center/right | |
| heading_text | 字符串 | 如 `目录` |
| entry_font | 字体名 | |
| entry_size | 字号 | |
| entry_line_spacing | float | |

### [figure] — 图格式

| Field | 类型 | 说明 |
|-------|------|------|
| caption_font | 字体名 | |
| caption_size | 字号 | |
| caption_align | left/center | |
| caption_prefix | 字符串 | 前缀，如 `图` |
| label_font | 字体名 | 编号字体 |
| label_bold | bool | |

### [table] — 表格式

| Field | 类型 | 说明 |
|-------|------|------|
| caption_font | 字体名 | |
| caption_size | 字号 | |
| caption_align | left/center | |
| caption_prefix | 字符串 | 前缀，如 `表` |
| label_font | 字体名 | |
| label_bold | bool | |
| content_font | 字体名 | 表内文字 |
| content_size | 字号 | |

### [references] — 参考文献

| Field | 类型 | 说明 |
|-------|------|------|
| heading_font | 字体名 | |
| heading_size | 字号 | |
| heading_bold | bool | |
| heading_align | left/center/right | |
| heading_text | 字符串 | 如 `参考文献` |
| entry_font | 字体名 | |
| entry_size | 字号 | |
| entry_line_spacing | float | |

### [acknowledgements] — 致谢

| Field | 类型 | 说明 |
|-------|------|------|
| heading_font | 字体名 | |
| heading_size | 字号 | |
| heading_bold | bool | |
| heading_align | left/center/right | |
| heading_text | 字符串 | 如 `致谢` |
| body_font | 字体名 | |
| body_size | 字号 | |
| body_line_spacing | float | |

---

## 值类型说明

- **字体名**: `SimSun`(宋体), `SimHei`(黑体), `KaiTi`(楷体), `FangSong`(仿宋), `Microsoft YaHei`(微软雅黑), `Arial`, `Times New Roman` 等
- **字号**: `42pt`(初号), `36pt`(小初), `26pt`(一号), `24pt`(小一), `22pt`(二号), `18pt`(小二), `16pt`(三号), `15pt`(小三), `14pt`(四号), `12pt`(小四), `10.5pt`(五号), `9pt`(小五)
- **尺寸**: `Ncm`, `Nmm`, `Nin`, `Npt`
- **bool**: `true` 或 `false`
- **字符串**: 支持 `{degree}` 占位符，生成时替换为 `本科`/`硕士`/`博士`
