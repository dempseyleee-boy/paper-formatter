# Paper Formatter

[![PyPI version](https://badge.fury.io/py/paper-formatter.svg)](https://badge.fury.io/py/paper-formatter)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

一个可以将 Markdown 或 Word 文档一键转换为符合中国硕士毕业论文格式的 Word 文档的工具。

[English](./README.md) | 简体中文

## ✨ 功能特点

- 📄 **多格式输入**: 支持 Markdown (.md) 和 Word (.docx) 输入
- 📑 **内置模板**: 符合中国硕士毕业论文格式要求
- 🎨 **自定义模板**: 支持 JSON 格式自定义模板
- 🔍 **自动识别**: 智能识别标题、列表、代码块、表格等元素
- 📑 **自动目录**: 自动生成目录

## 📋 格式说明

### 中国硕士毕业论文默认格式

| 元素 | 格式 |
|------|------|
| 纸张 | A4 |
| 页边距 | 上2.5cm，下2.5cm，左3.0cm，右2.0cm |
| 一级标题 | 黑体，三号，居中 |
| 二级标题 | 黑体，四号，居左 |
| 三级标题 | 宋体，小四号，居左 |
| 正文 | 宋体，小四号，两端对齐，行距1.5倍 |
| 代码 | Consolas，五号，带边框和底纹 |

## 🚀 快速开始

### 安装

```bash
pip install python-docx markdown
```

或者克隆项目：

```bash
git clone https://github.com/yourusername/paper_formatter.git
cd paper_formatter
pip install -r requirements.txt
```

### 基本用法

```bash
# Markdown 转 Word
python paper_formatter.py input.md output.docx

# Word 转 Word (会自动提取格式)
python paper_formatter.py input.docx output.docx

# 使用自定义模板
python paper_formatter.py input.md output.docx --template templates/china_master.json

# 列出可用模板
python paper_formatter.py --list-templates

# 提取输入文件的格式信息
python paper_formatter.py --extract input.md
```

## 📖 模板配置

### 创建自定义模板

创建 `my_template.json`：

```json
{
  "name": "自定义模板",
  "page": {
    "paper_size": "A4",
    "margin": {
      "top": 2.5,
      "bottom": 2.5,
      "left": 3.0,
      "right": 2.0
    }
  },
  "title": {
    "font_name": "黑体",
    "font_size": 22,
    "bold": true,
    "alignment": "center"
  },
  "heading1": {
    "font_name": "黑体",
    "font_size": 16,
    "bold": true,
    "alignment": "center",
    "before_spacing": 24,
    "after_spacing": 18
  },
  "heading2": {
    "font_name": "黑体",
    "font_size": 15,
    "bold": true,
    "alignment": "left"
  },
  "body": {
    "font_name": "宋体",
    "font_size": 12,
    "line_spacing": 1.5,
    "alignment": "justify"
  },
  "code": {
    "font_name": "Consolas",
    "font_size": 10,
    "border": true,
    "shading": "#F0F0F0"
  },
  "toc": {
    "include": true,
    "title": "目 录",
    "font_name": "黑体",
    "font_size": 14
  }
}
```

### 模板参数说明

| 参数 | 说明 |
|------|------|
| `font_name` | 字体名称，如 "宋体"、"黑体"、"Consolas" |
| `font_size` | 字号（磅） |
| `bold` | 是否加粗 (true/false) |
| `alignment` | 对齐方式 (left/center/right/justify) |
| `before_spacing` | 段前间距（磅） |
| `after_spacing` | 段后间距（磅） |
| `line_spacing` | 行距倍数 (1.5, 2.0 等) |

## 📁 项目结构

```
paper_formatter/
├── paper_formatter.py    # 主程序
├── templates/            # 模板目录
│   └── china_master.json # 默认模板
├── examples/             # 示例文件
│   ├── 论文示例.md
│   └── 论文示例.docx
├── requirements.txt      # Python依赖
├── .gitignore
├── LICENSE
└── README.md
```

## 🔧 依赖

- Python 3.7+
- python-docx >= 0.8.11
- markdown >= 3.3.0
- pandoc (用于 Word 输入转换，可选)

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

## 📝 许可证

MIT License - 查看 [LICENSE](LICENSE) 文件

## 🏆 感谢

- [python-docx](https://python-docx.readthedocs.io/) - Word 文档处理
- [Pandoc](https://pandoc.org/) - 文档格式转换
