#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Skill 首次部署脚本
==================
自动安装 Python 依赖、将内置公文字体安装到系统。
首次调用 Skill 时由 AI 自动执行，仅需用户授权即可。
所有字体已随 Skill 内置打包，无需联网下载。

CLI 用法：
  python setup.py              # 完整部署（依赖 + 字体安装）
  python setup.py deps         # 仅安装 Python 依赖
  python setup.py fonts        # 仅安装字体到系统
  python setup.py check        # 检查环境状态（不安装）
"""

import os
import shutil
import subprocess
import sys

PIPS = ["python-docx", "lxml"]


def _find_font_dir():
    """查找内置字体目录。优先级：GW_FONT_DIR 环境变量 > skill assets/fonts。"""
    candidates = []
    env_dir = os.environ.get("GW_FONT_DIR")
    if env_dir:
        candidates.append(env_dir)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    skill_root = os.path.dirname(script_dir)
    # 全局安装时直接用 skill 自带的 assets/fonts
    candidates.append(os.path.join(skill_root, "assets", "fonts"))
    # 项目内开发时上溯到项目根 fonts/embedded
    project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(script_dir))))
    candidates.append(os.path.join(project_root, "fonts", "embedded"))
    for d in candidates:
        if d and os.path.isdir(d):
            return d
    return candidates[0] if candidates else ""


def font_dest_dir():
    """系统字体目录（跨平台）。"""
    if sys.platform == "darwin":
        return os.path.expanduser("~/Library/Fonts")
    if sys.platform.startswith("win"):
        return os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")
    return os.path.expanduser("~/.fonts")


def install_deps():
    """安装 Python 依赖。返回 (成功数, 失败列表)。"""
    ok = 0
    failed = []
    for pkg in PIPS:
        try:
            __import__(pkg.replace("-", "_").split(".")[0])
            ok += 1
        except ImportError:
            print(f"  正在安装 {pkg}...")
            result = subprocess.run(
                [sys.executable, "-m", "pip", "install", pkg],
                capture_output=True, text=True,
            )
            if result.returncode == 0:
                print(f"  ✅ {pkg} 安装成功")
                ok += 1
            else:
                print(f"  ❌ {pkg} 安装失败: {result.stderr[:200]}")
                failed.append(pkg)
    return ok, failed


def iter_bundled_fonts():
    """遍历内置字体目录中的所有字体文件。"""
    font_dir = _find_font_dir()
    if not font_dir or not os.path.isdir(font_dir):
        return
    for root, _, files in os.walk(font_dir):
        for fn in files:
            if fn.lower().endswith((".ttf", ".otf", ".ttc")):
                yield os.path.join(root, fn)


def install_fonts_to_system():
    """将内置字体安装到系统字体目录。返回 (安装数, 跳过数)。"""
    font_dir = _find_font_dir()
    dest = font_dest_dir()
    os.makedirs(dest, exist_ok=True)
    installed = 0
    skipped = 0
    for path in iter_bundled_fonts():
        target = os.path.join(dest, os.path.basename(path))
        if os.path.exists(target):
            skipped += 1
        else:
            try:
                shutil.copy(path, target)
                installed += 1
            except Exception:
                pass
    return installed, skipped


def check_environment():
    """检查环境状态，返回可读报告。"""
    lines = ["=" * 50, "🔍 环境检查", "=" * 50, ""]

    # Python 依赖
    lines.append("Python 依赖：")
    for pkg in PIPS:
        mod = pkg.replace("-", "_").split(".")[0]
        try:
            __import__(mod)
            lines.append(f"  ✅ {pkg}")
        except ImportError:
            lines.append(f"  ❌ {pkg} — 未安装")

    # 字体
    lines.append("")
    lines.append("公文字体：")
    font_dir = _find_font_dir()
    dest = font_dest_dir()
    all_fonts = list(iter_bundled_fonts())
    installed_count = 0
    for path in all_fonts:
        if os.path.exists(os.path.join(dest, os.path.basename(path))):
            installed_count += 1
    lines.append(f"  内置字体文件：{len(all_fonts)} 个")
    lines.append(f"  已安装到系统：{installed_count} 个")
    if installed_count < len(all_fonts):
        lines.append(f"  ⚠ {len(all_fonts) - installed_count} 个字体尚未安装到系统")

    lines.append("")
    lines.append(f"字体目录：{font_dir}")
    lines.append(f"系统字体目录：{dest}")
    return "\n".join(lines)


def full_setup():
    """完整部署：安装依赖 + 安装内置字体到系统。"""
    print("=" * 50)
    print("🚀 公文格式工具 Skill · 首次部署")
    print("=" * 50)
    print()

    # 1. Python 依赖
    print("【1/2】安装 Python 依赖...")
    ok, failed = install_deps()
    if failed:
        print(f"  ⚠ {len(failed)} 个依赖安装失败: {', '.join(failed)}")
    else:
        print(f"  ✅ Python 依赖全部就绪（{ok} 个）")
    print()

    # 2. 安装字体到系统
    print("【2/2】安装公文字体到系统...")
    font_count = len(list(iter_bundled_fonts()))
    print(f"  内置字体包：{font_count} 个字体文件")
    installed, skipped = install_fonts_to_system()
    if installed > 0:
        print(f"  ✅ 新安装 {installed} 个字体到系统")
    if skipped > 0:
        print(f"  （{skipped} 个字体已存在，跳过）")
    print()

    print("=" * 50)
    print("✅ 部署完成！")
    print()
    print("⚠ 重要：请重启 Word/WPS 后字体才会完全生效。")
    print("  macOS：重启 Word/WPS 即可")
    print("  Windows：重启 Word/WPS，或重启电脑")
    print("=" * 50)
    return 0


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(full_setup())
    cmd = sys.argv[1].lower()
    if cmd == "deps":
        ok, failed = install_deps()
        sys.exit(0 if not failed else 1)
    elif cmd == "fonts":
        installed, skipped = install_fonts_to_system()
        print(f"新安装 {installed} 个，已存在 {skipped} 个")
        print()
        print("⚠ 重要：请重启 Word/WPS 后字体才会完全生效。")
        sys.exit(0)
    elif cmd == "check":
        print(check_environment())
    else:
        print(f"未知命令: {cmd}")
        print("用法: python setup.py [deps|fonts|check]")
        sys.exit(1)
