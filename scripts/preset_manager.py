#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
预设模板管理
============
列出 / 加载 / 保存 / 删除 / 描述公文格式预设。

CLI 用法：
  python preset_manager.py list                    # 列出所有预设
  python preset_manager.py describe <预设名或路径>  # 显示预设详情
  python preset_manager.py delete <自定义预设名>    # 删除自定义预设
"""

import glob
import json
import os
import re
import sys

# 把同目录的 format_engine.py 加入搜索路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from format_engine import describe_preset, pick_font  # noqa: E402


def _find_preset_dirs():
    """返回预设搜索目录列表（按优先级）。"""
    dirs = []
    script_dir = os.path.dirname(os.path.abspath(__file__))
    skill_root = os.path.dirname(script_dir)
    # scripts/ → gongwen-formatter/ → skills/ → .codebuddy/ → 项目根（4 级）
    project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(script_dir))))
    # 1. Skill 自带的预设
    dirs.append(os.path.join(skill_root, "assets", "presets"))
    # 2. 项目预设
    dirs.append(os.path.join(project_root, "presets"))
    # 3. 用户自定义预设
    dirs.append(os.path.expanduser("~/.gongwen_format_tool_presets"))
    return [d for d in dirs if os.path.isdir(d)]


def list_presets():
    """列出所有可用预设。返回 [(显示名, 文件路径), ...]
    
    按搜索优先级去重：同一文件名只保留第一个找到的。
    """
    items = []
    seen_names = set()  # 按预设名去重
    seen_files = set()  # 按文件名去重
    for d in _find_preset_dirs():
        for path in sorted(glob.glob(os.path.join(d, "*.json"))):
            fname = os.path.basename(path)
            if fname in seen_files:
                continue
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                name = data.get("preset_name", fname)
            except Exception:
                name = fname
            if name in seen_names:
                continue
            seen_names.add(name)
            seen_files.add(fname)
            items.append((name, path))
    return items


def find_preset(name_or_path):
    """按名称或路径查找预设。返回文件路径或 None。"""
    # 直接是路径
    if os.path.isfile(name_or_path):
        return name_or_path
    # 按名称匹配
    for display_name, path in list_presets():
        if display_name == name_or_path:
            return path
    # 模糊匹配
    for display_name, path in list_presets():
        if name_or_path in display_name:
            return path
    return None


def load_preset(name_or_path):
    """加载预设。name_or_path 可以是显示名或文件路径。"""
    path = find_preset(name_or_path)
    if not path:
        raise FileNotFoundError(f"找不到预设：{name_or_path}")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_preset(preset, name):
    """保存预设到用户目录。返回保存路径。"""
    safe = re.sub(r'[\\/:*?"<>|]+', "_", name).strip() or "custom"
    user_dir = os.path.expanduser("~/.gongwen_format_tool_presets")
    os.makedirs(user_dir, exist_ok=True)
    path = os.path.join(user_dir, "custom_" + safe + ".json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(preset, f, ensure_ascii=False, indent=2)
    return path


def delete_preset(name):
    """删除自定义预设。仅允许删除用户目录中的 custom_* 文件。"""
    user_dir = os.path.expanduser("~/.gongwen_format_tool_presets")
    safe = re.sub(r'[\\/:*?"<>|]+', "_", name).strip() or "custom"
    path = os.path.join(user_dir, "custom_" + safe + ".json")
    if os.path.exists(path):
        os.remove(path)
        return True
    # 尝试按显示名匹配
    for display_name, p in list_presets():
        if display_name == name and "custom_" in os.path.basename(p):
            os.remove(p)
            return True
    return False


def format_list_text():
    """格式化预设列表为可读文本。"""
    presets = list_presets()
    if not presets:
        return "⚠ 未找到任何格式预设。"
    lines = ["📋 可用公文格式预设", "=" * 50, ""]
    for i, (name, path) in enumerate(presets, 1):
        is_custom = "custom_" in os.path.basename(path)
        tag = " [自定义]" if is_custom else ""
        lines.append(f"{i}. {name}{tag}")
        lines.append(f"   文件：{path}")
    lines.append("")
    lines.append(f"共 {len(presets)} 个预设")
    return "\n".join(lines)


def format_describe_text(name_or_path):
    """格式化预设详情为可读文本。"""
    preset = load_preset(name_or_path)
    mode = preset.get("font_mode_default", "bundled")
    text = describe_preset(preset, mode)
    lines = ["=" * 50, text, "=" * 50, ""]
    # 额外显示文件路径
    path = find_preset(name_or_path)
    if path:
        lines.append(f"预设文件：{path}")
    return "\n".join(lines)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法：python preset_manager.py list|describe <名称>|delete <名称>")
        sys.exit(1)
    cmd = sys.argv[1].lower()
    if cmd == "list":
        print(format_list_text())
    elif cmd == "describe":
        if len(sys.argv) < 3:
            print("用法：python preset_manager.py describe <预设名或路径>")
            sys.exit(1)
        print(format_describe_text(sys.argv[2]))
    elif cmd == "delete":
        if len(sys.argv) < 3:
            print("用法：python preset_manager.py delete <自定义预设名>")
            sys.exit(1)
        name = sys.argv[2]
        if delete_preset(name):
            print(f"✅ 已删除自定义预设：{name}")
        else:
            print(f"⚠ 找不到可删除的自定义预设：{name}")
            print("（仅可删除用户目录中的自定义预设，内置预设不可删除）")
    else:
        print(f"未知命令：{cmd}")
        sys.exit(1)
