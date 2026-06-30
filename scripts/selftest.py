#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Skill 自检脚本
===============
验证 Skill 环境是否正常：Python 依赖、字体、预设、排版引擎。

CLI 用法：
  python selftest.py                   # 基础自检
  python selftest.py 文档.docx         # 带文档排版测试
"""

import os
import sys

# 把同目录加入搜索路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def run_selftest(test_doc=None):
    """执行自检，返回 (success: bool, report: str)。"""
    lines = []
    lines.append("=" * 50)
    lines.append("🔧 公文格式工具 Skill · 自检")
    lines.append("=" * 50)
    lines.append("")

    all_ok = True

    # 1. Python 依赖检查
    lines.append("1. Python 依赖检查")
    try:
        import docx  # noqa: F401
        lines.append("   ✅ python-docx 已安装")
    except ImportError:
        lines.append("   ❌ python-docx 未安装 — 请运行: pip install python-docx")
        all_ok = False
    try:
        import lxml  # noqa: F401
        lines.append("   ✅ lxml 已安装")
    except ImportError:
        lines.append("   ❌ lxml 未安装 — 请运行: pip install lxml")
        all_ok = False
    lines.append("")

    # 2. 字体检查
    lines.append("2. 字体检查")
    try:
        from font_manager import check_font_status, format_status_text  # noqa: E402
        status = check_font_status()
        if status["installed"]:
            lines.append(f"   ✅ 全部 {status['total']} 个字体已安装")
        else:
            lines.append(f"   ⚠ 共 {status['total']} 个字体，缺失 {len(status['missing'])} 个")
            lines.append("   运行以下命令安装：")
            lines.append(f"     python {os.path.abspath(__file__).replace('selftest', 'font_manager')} install")
            all_ok = False
    except Exception as e:
        lines.append(f"   ❌ 字体检查失败：{e}")
        all_ok = False
    lines.append("")

    # 3. 预设检查
    lines.append("3. 预设检查")
    try:
        from preset_manager import list_presets  # noqa: E402
        presets = list_presets()
        if presets:
            lines.append(f"   ✅ 找到 {len(presets)} 个预设：")
            for name, path in presets:
                lines.append(f"      · {name}")
        else:
            lines.append("   ❌ 未找到任何预设")
            all_ok = False
    except Exception as e:
        lines.append(f"   ❌ 预设检查失败：{e}")
        all_ok = False
    lines.append("")

    # 4. 排版引擎检查
    lines.append("4. 排版引擎检查")
    try:
        from format_engine import format_document, report_to_text  # noqa: E402
        from preset_manager import list_presets, load_preset  # noqa: E402
        lines.append("   ✅ 排版引擎可导入")
        if test_doc and os.path.exists(test_doc):
            presets = list_presets()
            if presets:
                preset = load_preset(presets[0][1])
                out, report = format_document(test_doc, preset, font_mode="bundled")
                lines.append(f"   ✅ 排版测试通过 — 输出: {out}")
                lines.append("")
                lines.append("   排版报告：")
                for rline in report_to_text(report).split("\n"):
                    lines.append("     " + rline)
        elif test_doc:
            lines.append(f"   ⚠ 测试文档不存在: {test_doc}")
    except Exception as e:
        lines.append(f"   ❌ 排版引擎测试失败：{e}")
        all_ok = False
    lines.append("")

    # 5. 文件转换能力
    lines.append("5. 文件转换能力")
    try:
        from file_converter import can_convert, find_libreoffice  # noqa: E402
        if can_convert():
            lines.append(f"   ✅ LibreOffice 可用 — 支持 .wps/.doc 自动转换")
            lines.append(f"      路径: {find_libreoffice()}")
        else:
            lines.append("   ⚠ LibreOffice 未安装 — .wps/.doc 需手动转换")
    except Exception as e:
        lines.append(f"   ❌ 文件转换检查失败：{e}")
    lines.append("")

    # 总结
    lines.append("=" * 50)
    if all_ok:
        lines.append("✅ 自检通过 — Skill 已就绪")
    else:
        lines.append("⚠ 自检发现问题 — 请按上述提示修复")
    lines.append("=" * 50)

    return all_ok, "\n".join(lines)


if __name__ == "__main__":
    test_doc = sys.argv[1] if len(sys.argv) > 1 else None
    ok, report = run_selftest(test_doc)
    print(report)
    sys.exit(0 if ok else 1)
