# 预设 JSON Schema 说明

每个预设是一个 JSON 文件，描述一套完整的公文格式标准。

## 顶层字段

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `preset_name` | string | ✅ | 预设显示名称 |
| `description` | string | | 预设描述（不在提示栏显示） |
| `font_mode_default` | string | | 默认字体方案：`bundled`/`mac`/`standard` |
| `western_font` | string | | 西文/数字字体名（如 `Times New Roman`） |
| `page` | object | ✅ | 页面设置 |
| `line_spacing_pt` | number | | 固定行距（磅），默认 28.8 |
| `page_number` | object | | 页码设置 |
| `letterhead` | object | | 红头设置 |
| `footer_line` | object | | 页脚红线设置 |
| `auto_match_roles` | array | | 自动角色识别顺序 |
| `roles` | object | ✅ | 各角色排版参数 |
| `table_style` | object | | 表格样式 |

## page 字段

```json
{
  "size": "A4",
  "width_mm": 210,
  "height_mm": 297,
  "margin_top_mm": 37,
  "margin_bottom_mm": 35,
  "margin_left_mm": 28,
  "margin_right_mm": 26
}
```

## roles 字段

每个角色（title/h1/h2/h3/h4/body/table/figure/fig_caption/tbl_caption/
doc_number/recipient/signoff）包含：

| 字段 | 类型 | 说明 |
|------|------|------|
| `label` | string | 角色中文名 |
| `font` | string | 标准字体名 |
| `mac_font` | string | Mac 替代字体名 |
| `bundled_font` | string | 内置字体名（通常同 font） |
| `size_pt` | number | 字号（磅） |
| `bold` | bool | 是否加粗 |
| `color` | string | 字体颜色（hex，如 `000000`） |
| `align` | string | 对齐：`center`/`left`/`right`/`justify` |
| `first_line_indent_chars` | number | 首行缩进字数 |
| `space_before_pt` | number | 段前间距（磅） |
| `space_after_pt` | number | 段后间距（磅） |
| `numbering` | string | 序号形式（如 `一、`/`（一）`/`1.`/`（1）`） |
| `auto_number` | string | 自动序号样式：`cn_dun`/`cn_paren`/`digit_dot`/`digit_paren` |
| `match_regex` | string | 自动识别正则（用于图题/表题/发文字号/主送/落款） |

## letterhead 字段（红头）

```json
{
  "enabled": true,
  "org_name": "机关名称",
  "font": "方正大标宋简体",
  "mac_font": "STSong",
  "bundled_font": "方正大标宋简体",
  "size_pt": 24,
  "bold": false,
  "color": "FF0000",
  "align": "center",
  "space_before_pt": 10,
  "space_after_pt": 2,
  "line": {
    "enabled": true,
    "color": "FF0000",
    "width_pt": 3,
    "space_before_pt": 0,
    "space_after_pt": 12
  }
}
```

## page_number 字段

```json
{
  "enabled": true,
  "font": "宋体",
  "mac_font": "STSong",
  "bundled_font": "宋体",
  "size_pt": 14,
  "format": "— {n} —",
  "odd_align": "right",
  "even_align": "left"
}
```

## table_style 字段

```json
{
  "three_line": true,
  "header_bold": true,
  "header_align": "center",
  "outer_line_pt": 1.5,
  "header_line_pt": 0.75,
  "inner_lines": false
}
```

## auto_match_roles 字段

角色自动识别顺序数组。引擎仅对 `role == body` 的段落按此顺序匹配 `match_regex`。
常见值：`["fig_caption", "tbl_caption", "doc_number", "recipient", "signoff"]`

## 自定义预设存储

- 内置预设：`assets/presets/*.json`（只读）
- 项目预设：项目根 `presets/*.json`（只读）
- 用户自定义：`~/.gongwen_format_tool_presets/custom_*.json`（可读写）
