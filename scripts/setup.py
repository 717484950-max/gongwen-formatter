#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Skill 首次部署脚本
==================
自动安装 Python 依赖、下载并安装公文字体。
首次调用 Skill 时由 AI 自动执行，仅需用户授权即可。

CLI 用法：
  python setup.py              # 完整部署（依赖 + 字体下载 + 安装）
  python setup.py deps         # 仅安装 Python 依赖
  python setup.py fonts        # 仅下载并安装字体
  python setup.py check        # 检查环境状态（不安装）
"""

import os
import shutil
import subprocess
import sys
import urllib.request

# ---- 字体下载源（GitHub 公开仓库） ----
FONT_BASE_URL = "https://raw.githubusercontent.com/wuhongyi/fonts/master/"
# 可从 GitHub 下载的公共字体
DOWNLOADABLE_FONTS = {
    "仿宋_GB2312.ttf": "仿宋_GB2312.ttf",
    "楷体_GB2312.ttf": "楷体_GB2312.ttf",
    "黑体.ttf": "黑体.ttf",
    "宋体.ttc": "宋体.ttc",
}
# Times New Roman 子目录
TNR_BASE_URL = "https://raw.githubusercontent.com/wuhongyi/fonts/master/TimesNewRoman/"
DOWNLOADABLE_TNR = {
    "times.ttf": "times.ttf",
    "timesbd.ttf": "timesbd.ttf",
    "timesi.ttf": "timesi.ttf",
    "timesbi.ttf": "timesbi.ttf",
}
# 方正系列为商业字体，无法从公开源下载，需用户自行获取
COMMERCIAL_FONTS = [
    "方正小标宋简体.ttf",
    "方正小标宋_GBK.ttf",
    "方正大标宋简体.ttf",
    "方正大标宋简繁.ttf",
    "方正仿宋简体.ttf",
    "方正仿宋_GBK.ttf",
    "方正楷体简体.ttf",
    "方正楷体_GBK.ttf",
    "方正黑体简体.ttf",
    "方正黑体_GBK.ttf",
    "仿宋.ttf",
    "楷体.ttf",
]

PIPS = ["python-docx", "lxml"]


def _find_font_dir():
    """查找或创建字体存放目录。"""
    candidates = []
    env_dir = os.environ.get("GW_FONT_DIR")
    if env_dir:
        candidates.append(env_dir)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    skill_root = os.path.dirname(script_dir)
    project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(script_dir))))
    candidates.append(os.path.join(project_root, "fonts", "embedded"))
    candidates.append(os.path.join(skill_root, "assets", "fonts"))
    for d in candidates:
        if d:
            os.makedirs(d, exist_ok=True)
            return d
    return candidates[0] if candidates else ""


def font_dest_dir():
    """系统字体目录。"""
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


def download_fonts():
    """从 GitHub 下载公共字体到本地字体目录。返回 (下载数, 失败列表)。"""
    font_dir = _find_font_dir()
    downloaded = 0
    failed = []

    for filename, remote_name in DOWNLOADABLE_FONTS.items():
        local_path = os.path.join(font_dir, filename)
        if os.path.exists(local_path):
            continue
        url = FONT_BASE_URL + remote_name
        print(f"  下载 {filename}...")
        try:
            urllib.request.urlretrieve(url, local_path)
            downloaded += 1
        except Exception as e:
            print(f"  ❌ 下载失败: {filename} — {e}")
            failed.append(filename)

    # Times New Roman 子目录
    tnr_dir = os.path.join(font_dir, "TimesNewRoman")
    os.makedirs(tnr_dir, exist_ok=True)
    for filename, remote_name in DOWNLOADABLE_TNR.items():
        local_path = os.path.join(tnr_dir, filename)
        if os.path.exists(local_path):
            continue
        url = TNR_BASE_URL + remote_name
        print(f"  下载 TimesNewRoman/{filename}...")
        try:
            urllib.request.urlretrieve(url, local_path)
            downloaded += 1
        except Exception as e:
            print(f"  ❌ 下载失败: TimesNewRoman/{filename} — {e}")
            failed.append(f"TimesNewRoman/{filename}")

    return downloaded, failed


def install_fonts_to_system():
    """将本地字体目录中的字体安装到系统字体目录。返回安装数。"""
    font_dir = _find_font_dir()
    dest = font_dest_dir()
    os.makedirs(dest, exist_ok=True)
    cnt = 0
    for root, _, files in os.walk(font_dir):
        for fn in files:
            if fn.lower().endswith((".ttf", ".otf", ".ttc")):
                src = os.path.join(root, fn)
                target = os.path.join(dest, fn)
                if not os.path.exists(target):
                    try:
                        shutil.copy(src, target)
                        cnt += 1
                    except Exception:
                        pass
    return cnt


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
    local_count = 0
    installed_count = 0
    for root, _, files in os.walk(font_dir):
        for fn in files:
            if fn.lower().endswith((".ttf", ".otf", ".ttc")):
                local_count += 1
                if os.path.exists(os.path.join(dest, fn)):
                    installed_count += 1
    lines.append(f"  本地字体文件：{local_count} 个")
    lines.append(f"  已安装到系统：{installed_count} 个")

    # 方正字体提醒
    missing_commercial = []
    for fn in COMMERCIAL_FONTS:
        if not os.path.exists(os.path.join(font_dir, fn)):
            missing_commercial.append(fn)
    if missing_commercial:
        lines.append(f"  ⚠ 商业字体未获取（{len(missing_commercial)} 个）：")
        for fn in missing_commercial[:5]:
            lines.append(f"    · {fn}")
        if len(missing_commercial) > 5:
            lines.append(f"    ...等 {len(missing_commercial)} 个")
        lines.append("  （排版时会使用系统替代字体，不影响基本功能）")

    lines.append("")
    lines.append(f"字体目录：{font_dir}")
    lines.append(f"系统字体目录：{dest}")
    return "\n".join(lines)


def full_setup():
    """完整部署：安装依赖 + 下载字体 + 安装字体到系统。"""
    print("=" * 50)
    print("🚀 公文格式工具 Skill · 首次部署")
    print("=" * 50)
    print()

    # 1. Python 依赖
    print("【1/3】安装 Python 依赖...")
    ok, failed = install_deps()
    if failed:
        print(f"  ⚠ {len(failed)} 个依赖安装失败: {', '.join(failed)}")
    else:
        print(f"  ✅ Python 依赖全部就绪（{ok} 个）")
    print()

    # 2. 下载字体
    print("【2/3】下载公文字体...")
    font_dir = _find_font_dir()
    # 先检查已有字体
    existing = []
    for root, _, files in os.walk(font_dir):
        for fn in files:
            if fn.lower().endswith((".ttf", ".otf", ".ttc")):
                existing.append(fn)
    if len(existing) >= 8:
        print(f"  字体目录已有 {len(existing)} 个字体，跳过下载")
    else:
        downloaded, dl_failed = download_fonts()
        if dl_failed:
            print(f"  ⚠ {len(dl_failed)} 个字体下载失败")
        else:
            print(f"  ✅ 已下载 {downloaded} 个字体")
    print()

    # 3. 安装字体到系统
    print("【3/3】安装字体到系统...")
    installed = install_fonts_to_system()
    if installed > 0:
        print(f"  ✅ 已安装 {installed} 个字体到系统")
    else:
        print("  字体已在系统目录中，无需重复安装")
    print()

    # 检查方正字体
    font_dir = _find_font_dir()
    missing_commercial = [fn for fn in COMMERCIAL_FONTS if not os.path.exists(os.path.join(font_dir, fn))]
    if missing_commercial:
        print("⚠ 商业字体提醒：")
        print(f"  以下 {len(missing_commercial)} 个方正系列字体无法自动下载：")
        for fn in missing_commercial[:5]:
            print(f"    · {fn}")
        print("  排版时会使用系统替代字体。")
        print("  如需精确显示，请自行获取这些字体并放入：")
        print(f"    {font_dir}")
        print()

    print("=" * 50)
    print("✅ 部署完成！")
    print()
    print("⚠ 重要：请重启 Word/WPS 后字体才会完全生效。")
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
        downloaded, failed = download_fonts()
        installed = install_fonts_to_system()
        print(f"下载 {downloaded} 个，安装 {installed} 个到系统")
        sys.exit(0 if not failed else 1)
    elif cmd == "check":
        print(check_environment())
    else:
        print(f"未知命令: {cmd}")
        print("用法: python setup.py [deps|fonts|check]")
        sys.exit(1)
