#!/bin/bash
# 公文格式工具一键安装脚本
# 用法：curl -fsSL <raw-url> | bash

set -e

echo "📦 公文格式排版 Skill 安装中..."

SKILL_DIR="$HOME/.codebuddy/skills/gongwen-formatter"
REPO_URL="https://github.com/717484950-max/gongwen-formatter.git"

if [ -d "$SKILL_DIR" ]; then
    echo "📁 目录已存在，更新中..."
    cd "$SKILL_DIR"
    git pull
else
    echo "📥 克隆仓库..."
    mkdir -p "$(dirname "$SKILL_DIR")"
    git clone "$REPO_URL" "$SKILL_DIR"
fi

echo "🔧 安装 Python 依赖和字体..."
cd "$SKILL_DIR"
python scripts/setup.py

echo ""
echo "✅ 安装完成！"
echo "   现在可以在 CodeBuddy 中对 AI 说「请排版公文格式」即可使用。"
echo "   💡 请重启 Word/WPS 后字体完全生效。"
