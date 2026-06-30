# 公文格式排版 Skill (gongwen-formatter)

> 党政机关公文格式一键排版工具。将 Word(.docx) 转换为符合 GB/T 9704-2012 标准的公文版式。
> 适用于 CodeBuddy / Claude Code 等 AI 编程助手。

## 安装

### 方式一：AI 一键安装

对任意 AI 助手下达：
> "帮我在 GitHub 上安装 717484950-max/gongwen-formatter skill"

AI 会自动执行：
```bash
git clone https://github.com/717484950-max/gongwen-formatter.git ~/.codebuddy/skills/gongwen-formatter
python ~/.codebuddy/skills/gongwen-formatter/scripts/setup.py
```

### 方式二：手动安装

```bash
# 克隆到 CodeBuddy skills 目录
git clone https://github.com/717484950-max/gongwen-formatter.git ~/.codebuddy/skills/gongwen-formatter

# 安装依赖和字体
python ~/.codebuddy/skills/gongwen-formatter/scripts/setup.py
```

### 方式三：快速安装脚本

```bash
curl -fsSL https://raw.githubusercontent.com/717484950-max/gongwen-formatter/main/install.sh | bash
```

## 功能

- 📄 `.docx` / `.wps` / `.doc` → 标准公文格式
- 📐 GB/T 9704-2012 国标对齐（行距 28 磅 / 标题二号方正小标宋 / 正文三号仿宋）
- 🏷️ 红头公文 + 标准公文双预设
- 📊 自动三线表、图片居中、图表题注
- 🔤 内置 20 个公文字体包（方正小标宋简体/GBK、方正大标宋简体/简繁、仿宋/仿宋_GB2312、楷体/楷体_GB2312、黑体、方正各系列、宋体、Times New Roman），首次使用一键安装
- 🔍 自动识别：发文字号、主送机关、落款署名与日期
- 🔒 全程本地处理，不联网，不修改原稿
- 💾 输出文档自动嵌入字体，跨电脑显示一致

## 使用

在 CodeBuddy 对话中触发：
- 发送 `.docx` 文件 + "请排版为公文格式"
- 说"排版公文"、"红头文件"、"GB/T 9704"、"公文标准格式"

Skill 会自动预检文档结构 → 确认格式 → 执行排版 → 输出 `原名-公文版.docx`。

## 默认格式

| 要素 | 字体 | 字号 |
|---|---|---|
| 红头机关名 | 方正大标宋简体 | 24pt 红色 |
| 大标题 | 方正小标宋简体 | 22pt（二号） |
| 一级标题 | 黑体 | 16pt（三号） |
| 二级标题 | 楷体_GB2312 | 16pt（三号） |
| 正文 | 仿宋_GB2312 | 16pt（三号） |
| 全文行距 | — | 28 磅 |

> 默认机关名：深圳市腾讯计算机系统有限公司（可在排版时修改）

## 项目结构

```
├── SKILL.md              # Skill 主指令
├── README.md             # 本文件
├── install.sh            # 一键安装脚本
├── scripts/
│   ├── setup.py          # 首次部署（依赖+字体）
│   ├── format_engine.py  # 核心排版引擎
│   ├── font_manager.py   # 字体检查/安装
│   ├── doc_analyzer.py   # 文档结构预检
│   ├── preset_manager.py # 预设管理
│   ├── file_converter.py # 格式转换
│   └── selftest.py       # 环境自检
├── assets/
│   ├── presets/          # 格式预设 JSON
│   └── fonts/            # 标准公文字体
└── references/           # 国标速查 + Schema 文档
```

## 许可

仅限个人及内部使用。字体版权归各自所有者。
