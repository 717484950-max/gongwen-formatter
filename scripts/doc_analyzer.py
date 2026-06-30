#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文档预检分析
============
在排版前分析 .docx 文档结构，检测标题层级是否正确标注，
为 AI 提供结构化信息以决定是否需要警告用户。

CLI 用法：
  python doc_analyzer.py 文档.docx
"""

import os
import re
import sys
import json

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph


# Word 内置标题样式名 → 角色映射
_STYLE_MAP = {
    "title": "title", "标题": "title",
    "heading 1": "h1", "标题 1": "h1", "标题1": "h1",
    "heading 2": "h2", "标题 2": "h2", "标题2": "h2",
    "heading 3": "h3", "标题 3": "h3", "标题3": "h3",
    "heading 4": "h4", "标题 4": "h4", "标题4": "h4",
}


def _classify(paragraph):
    name = (paragraph.style.name or "").strip().lower()
    return _STYLE_MAP.get(name, "body")


def _is_image(paragraph):
    p = paragraph._p
    return bool(
        p.findall(".//" + qn("w:drawing"))
        or p.findall(".//" + qn("w:object"))
        or p.findall(".//" + qn("w:pict"))
    )


def _iter_blocks(doc):
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)


# 疑似标题的正文段落模式（用于预检警告）
_SUSPECT_PATTERNS = [
    (r"^\s*[一二三四五六七八九十]+\s*[、.．]\s*\S", "一级标题（一、）"),
    (r"^\s*[（(][一二三四五六七八九十]+[)）]\s*\S", "二级标题（（一））"),
    (r"^\s*\d+\s*[、.．]\s*\S", "三级标题（1.）"),
    (r"^\s*[（(]\d+[)）]\s*\S", "四级标题（（1））"),
]


def analyze_document(path):
    """分析文档结构，返回结构化预检结果。

    返回 dict:
      file: str              — 文件名
      total_paragraphs: int  — 段落总数（不含空段）
      total_tables: int      — 表格总数
      headings: dict         — 各级标题数量 {title, h1, h2, h3, h4}
      body_count: int        — 正文段落数
      figures: int           — 图片数
      suspected_unstyled: list — 疑似未标样式的标题 [(段落序号, 文本前30字, 疑似级别)]
      has_heading_styles: bool — 是否检测到任何标题样式
      summary: str           — 可读摘要
    """
    doc = Document(path)
    headings = {"title": 0, "h1": 0, "h2": 0, "h3": 0, "h4": 0}
    body_count = 0
    figures = 0
    tables = 0
    suspected = []
    para_idx = 0

    for block in _iter_blocks(doc):
        if isinstance(block, Table):
            tables += 1
            continue
        p = block
        if _is_image(p):
            figures += 1
            continue
        text = (p.text or "").strip()
        if not text:
            continue
        para_idx += 1
        role = _classify(p)
        if role in headings:
            headings[role] += 1
        else:
            body_count += 1
            # 检查疑似未标注的标题
            for pat, level in _SUSPECT_PATTERNS:
                if re.match(pat, text):
                    suspected.append((para_idx, text[:40], level))
                    break

    has_styles = sum(headings.values()) > 0
    fname = os.path.basename(path)

    # 生成可读摘要
    lines = [f"📄 文档结构分析：{fname}", ""]
    lines.append(f"段落总数：{para_idx}（含正文 {body_count}）")
    lines.append(f"表格数量：{tables}")
    lines.append(f"图片数量：{figures}")
    lines.append("")
    lines.append("标题层级分布：")
    lines.append(f"  大标题（标题样式）：{headings['title']}")
    lines.append(f"  一级标题（标题1）：{headings['h1']}")
    lines.append(f"  二级标题（标题2）：{headings['h2']}")
    lines.append(f"  三级标题（标题3）：{headings['h3']}")
    lines.append(f"  四级标题（标题4）：{headings['h4']}")
    if not has_styles:
        lines.append("")
        lines.append("⚠ 警告：未检测到任何标题样式！")
        lines.append("  工具将无法识别层级，全部内容会被当作正文处理。")
        lines.append("  请先在 Word/WPS 中将标题设为对应的“标题/标题1/2/3/4”样式。")
    if suspected:
        lines.append("")
        lines.append(f"⚠ 发现 {len(suspected)} 个疑似未标注层级的标题：")
        for idx, text, level in suspected[:10]:
            lines.append(f"  第{idx}段“{text}”→ 疑似{level}")
        if len(suspected) > 10:
            lines.append(f"  ...还有 {len(suspected) - 10} 个")
    summary = "\n".join(lines)

    return {
        "file": fname,
        "total_paragraphs": para_idx,
        "total_tables": tables,
        "headings": headings,
        "body_count": body_count,
        "figures": figures,
        "suspected_unstyled": suspected,
        "has_heading_styles": has_styles,
        "summary": summary,
    }


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法：python doc_analyzer.py 文档.docx")
        sys.exit(1)
    path = sys.argv[1]
    if not os.path.exists(path):
        print(f"错误：文件不存在 — {path}")
        sys.exit(1)
    if not path.lower().endswith(".docx"):
        print(f"错误：只支持 .docx 格式 — {path}")
        sys.exit(1)
    result = analyze_document(path)
    print(result["summary"])
    # 同时输出 JSON 供 AI 解析
    print("\n--- JSON ---")
    json_result = {k: v for k, v in result.items() if k != "summary"}
    print(json.dumps(json_result, ensure_ascii=False, indent=2))
