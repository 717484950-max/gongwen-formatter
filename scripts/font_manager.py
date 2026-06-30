#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
公文字体管理
============
检查/安装/打开内置公文字体包。
支持 macOS / Windows / Linux。

CLI 用法：
  python font_manager.py check      # 检查字体安装状态（输出详细清单）
  python font_manager.py install    # 一键安装字体到系统字体目录
  python font_manager.py open       # 打开内置字体文件夹
"""

import os
import shutil
import subprocess
import sys


def _find_font_dir():
    """查找内置字体目录，逻辑与 format_engine.py 一致。"""
    candidates = []
    env_dir = os.environ.get("GW_FONT_DIR")
    if env_dir:
        candidates.append(env_dir)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    skill_root = os.path.dirname(script_dir)
    # scripts/ → gongwen-formatter/ → skills/ → .codebuddy/ → 项目根（4 级）
    project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(script_dir))))
    candidates.append(os.path.join(project_root, "fonts", "embedded"))
    candidates.append(os.path.join(skill_root, "assets", "fonts"))
    for d in candidates:
        if d and os.path.isdir(d):
            return d
    return candidates[0] if candidates else ""


EMBEDDED_FONT_DIR = _find_font_dir()

# 公文字体用途说明（用于向用户展示）
FONT_PURPOSES = {
    "方正小标宋简体": "大标题（发文标题）",
    "方正小标宋_GBK": "大标题（备用）",
    "方正大标宋简体": "红头机关名",
    "方正大标宋简繁": "红头机关名（备用）",
    "仿宋_GB2312": "正文 / 各级标题",
    "仿宋": "正文（备用）",
    "方正仿宋简体": "正文（备用）",
    "方正仿宋_GBK": "正文（备用）",
    "楷体_GB2312": "二级标题",
    "楷体": "二级标题（备用）",
    "方正楷体简体": "二级标题（备用）",
    "方正楷体_GBK": "二级标题（备用）",
    "黑体": "一级标题 / 图表题注",
    "方正黑体简体": "一级标题（备用）",
    "方正黑体_GBK": "一级标题（备用）",
    "宋体": "页码",
    "Times New Roman": "西文 / 数字",
}


def font_dest_dir():
    """返回系统字体目录。"""
    if sys.platform == "darwin":
        return os.path.expanduser("~/Library/Fonts")
    if sys.platform.startswith("win"):
        return os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")
    return os.path.expanduser("~/.fonts")


def iter_embedded_fonts():
    """遍历内置字体目录中的所有字体文件。"""
    if not EMBEDDED_FONT_DIR or not os.path.isdir(EMBEDDED_FONT_DIR):
        return
    for root, _, files in os.walk(EMBEDDED_FONT_DIR):
        for fn in files:
            if fn.lower().endswith((".ttf", ".otf", ".ttc")):
                yield os.path.join(root, fn)


def _font_display_name(filename):
    """从文件名推断字体显示名（用于状态清单）。"""
    base = os.path.splitext(filename)[0]
    return FONT_PURPOSES.get(base, "")


def fonts_installed():
    """检查所有内置字体是否已安装到系统字体目录。"""
    dest = font_dest_dir()
    needed = [os.path.basename(p) for p in iter_embedded_fonts()]
    return bool(needed) and all(os.path.exists(os.path.join(dest, n)) for n in needed)


def check_font_status():
    """返回详细的字体安装状态清单，供 AI 展示给用户。

    返回 dict:
      installed: bool  — 是否全部已装
      total: int       — 内置字体总数
      missing: list    — 缺失字体 [(文件名, 用途), ...]
      present: list    — 已装字体 [(文件名, 用途), ...]
      font_dir: str    — 内置字体目录路径
      dest_dir: str    — 系统字体目录路径
    """
    dest = font_dest_dir()
    all_fonts = list(iter_embedded_fonts())
    missing = []
    present = []
    for path in all_fonts:
        fn = os.path.basename(path)
        purpose = _font_display_name(fn)
        if os.path.exists(os.path.join(dest, fn)):
            present.append((fn, purpose))
        else:
            missing.append((fn, purpose))
    return {
        "installed": len(missing) == 0,
        "total": len(all_fonts),
        "missing": missing,
        "present": present,
        "font_dir": EMBEDDED_FONT_DIR,
        "dest_dir": dest,
    }


def install_fonts():
    """把内置字体复制到系统字体目录。返回新安装的数量。"""
    dest = font_dest_dir()
    os.makedirs(dest, exist_ok=True)
    cnt = 0
    for path in iter_embedded_fonts():
        target = os.path.join(dest, os.path.basename(path))
        if not os.path.exists(target):
            try:
                shutil.copy(path, target)
                cnt += 1
            except Exception:
                pass
    return cnt


def open_font_folder():
    """在系统文件管理器中打开内置字体文件夹。"""
    if not EMBEDDED_FONT_DIR or not os.path.isdir(EMBEDDED_FONT_DIR):
        return False
    try:
        if sys.platform == "darwin":
            subprocess.run(["open", EMBEDDED_FONT_DIR], check=False)
        elif sys.platform.startswith("win"):
            os.startfile(EMBEDDED_FONT_DIR)  # type: ignore[attr-defined]
        else:
            subprocess.run(["xdg-open", EMBEDDED_FONT_DIR], check=False)
        return True
    except Exception:
        return False


def format_status_text():
    """把字体状态格式化为可读文本（供 AI 直接展示）。"""
    status = check_font_status()
    lines = ["=" * 50, "🔤 公文字体安装状态", "=" * 50, ""]
    if status["installed"]:
        lines.append(f"✅ 全部 {status['total']} 个字体已安装")
    else:
        lines.append(f"⚠ 共 {status['total']} 个字体，缺失 {len(status['missing'])} 个")
    lines.append("")
    if status["missing"]:
        lines.append("❌ 缺失字体：")
        for fn, purpose in status["missing"]:
            lines.append(f"   · {fn}" + (f"（{purpose}）" if purpose else ""))
        lines.append("")
        lines.append("请运行以下命令一键安装：")
        lines.append(f"  python {os.path.abspath(__file__)} install")
        lines.append("或在文件管理器中打开字体文件夹手动安装：")
        lines.append(f"  python {os.path.abspath(__file__)} open")
        lines.append("")
        lines.append("⚠ 安装后请重启 Word/WPS 才能生效")
    if status["present"]:
        lines.append("")
        lines.append("✅ 已安装字体：")
        for fn, purpose in status["present"]:
            lines.append(f"   · {fn}" + (f"（{purpose}）" if purpose else ""))
    lines.append("")
    lines.append(f"内置字体目录：{status['font_dir']}")
    lines.append(f"系统字体目录：{status['dest_dir']}")
    return "\n".join(lines)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法：python font_manager.py check|install|open")
        sys.exit(1)
    cmd = sys.argv[1].lower()
    if cmd == "check":
        print(format_status_text())
    elif cmd == "install":
        print("正在安装公文字体包...")
        cnt = install_fonts()
        if cnt > 0:
            print(f"✅ 已安装 {cnt} 个字体到 {font_dest_dir()}")
        else:
            print("所有字体已存在于系统字体目录，无需重复安装。")
        print()
        print(format_status_text())
        print()
        print("⚠ 重要：请重启电脑（至少重启 Word/WPS）后字体才会完全生效。")
        print("如自动安装未生效，可打开字体文件夹双击手动安装：")
        print(f"  python {os.path.abspath(__file__)} open")
        # 安装后自动打开文件夹
        open_font_folder()
        print(f"已打开字体文件夹：{EMBEDDED_FONT_DIR}")
    elif cmd == "open":
        if open_font_folder():
            print(f"已打开字体文件夹：{EMBEDDED_FONT_DIR}")
            print("双击字体文件 →“安装字体”即可手动安装。")
            print("安装后请重启 Word/WPS。")
        else:
            print("⚠ 无法打开字体文件夹，请手动访问：", EMBEDDED_FONT_DIR)
    else:
        print(f"未知命令：{cmd}")
        print("用法：python font_manager.py check|install|open")
        sys.exit(1)
