#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件格式转换
============
检测 LibreOffice 是否可用，并将 .wps / .doc 转换为 .docx。

CLI 用法：
  python file_converter.py check              # 检查 LibreOffice 是否可用
  python file_converter.py convert 文件.wps   # 转换为 .docx
"""

import os
import shutil
import subprocess
import sys


def find_libreoffice():
    """查找 LibreOffice 可执行文件路径。返回路径或 None。"""
    # macOS 常见路径
    mac_paths = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for p in mac_paths:
        if os.path.isfile(p):
            return p
    # Windows / Linux: 查 PATH
    for name in ("soffice", "libreoffice", "soffice.exe"):
        found = shutil.which(name)
        if found:
            return found
    return None


def can_convert():
    """是否可以进行 .wps/.doc → .docx 转换。"""
    return find_libreoffice() is not None


def convert_to_docx(input_path, output_dir=None):
    """将 .wps / .doc 转换为 .docx。返回输出路径或 None。

    依赖 LibreOffice 的 headless 模式。
    """
    soffice = find_libreoffice()
    if not soffice:
        return None
    if not os.path.exists(input_path):
        return None
    out_dir = output_dir or os.path.dirname(os.path.abspath(input_path))
    try:
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "docx",
             "--outdir", out_dir, input_path],
            capture_output=True, text=True, timeout=60,
        )
        if result.returncode != 0:
            return None
        base = os.path.splitext(os.path.basename(input_path))[0]
        out_path = os.path.join(out_dir, base + ".docx")
        if os.path.exists(out_path):
            return out_path
    except Exception:
        return None
    return None


def check_and_report():
    """检查转换能力并返回可读报告。"""
    soffice = find_libreoffice()
    lines = ["=" * 50, "📂 文件格式转换能力检查", "=" * 50, ""]
    if soffice:
        lines.append("✅ LibreOffice 已安装")
        lines.append(f"   路径：{soffice}")
        lines.append("")
        lines.append("支持的输入格式：")
        lines.append("   · .docx  — 直接处理（无需转换）")
        lines.append("   · .wps   — 自动转换为 .docx")
        lines.append("   · .doc   — 自动转换为 .docx")
    else:
        lines.append("⚠ LibreOffice 未安装")
        lines.append("")
        lines.append("支持的输入格式：")
        lines.append("   · .docx  — 直接处理（无需转换）")
        lines.append("   · .wps   — ❌ 需手动另存为 .docx")
        lines.append("   · .doc   — ❌ 需手动另存为 .docx")
        lines.append("")
        lines.append("如需自动转换 .wps / .doc，请安装 LibreOffice：")
        lines.append("   macOS:  brew install --cask libreoffice")
        lines.append("   或访问: https://www.libreoffice.org/download/")
    return "\n".join(lines)


def ensure_docx(path):
    """确保文件是 .docx 格式。如需转换则自动转换。

    返回 (docx_path, was_converted, message)
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return (path, False, "已是 .docx 格式，无需转换")
    if ext in (".wps", ".doc"):
        if not can_convert():
            return (None, False, f"⚠ {ext} 格式需要 LibreOffice 才能自动转换。\n"
                    "请用 WPS/Word 打开后另存为 .docx，或安装 LibreOffice。")
        out = convert_to_docx(path)
        if out:
            return (out, True, f"✅ 已将 {ext} 转换为 .docx：{out}")
        return (None, False, f"⚠ {ext} 转换失败，请手动另存为 .docx")
    return (None, False, f"⚠ 不支持的格式 {ext}，仅支持 .docx / .wps / .doc")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法：python file_converter.py check|convert <文件>")
        sys.exit(1)
    cmd = sys.argv[1].lower()
    if cmd == "check":
        print(check_and_report())
    elif cmd == "convert":
        if len(sys.argv) < 3:
            print("用法：python file_converter.py convert <文件.wps>")
            sys.exit(1)
        path = sys.argv[2]
        docx_path, converted, msg = ensure_docx(path)
        print(msg)
        if docx_path:
            print(f"\n输出文件：{docx_path}")
    else:
        print(f"未知命令：{cmd}")
        sys.exit(1)
