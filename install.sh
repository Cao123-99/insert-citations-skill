#!/usr/bin/env bash
set -euo pipefail

# ============================================================
# insert-citations skill 安装脚本
# ============================================================
# 用法：
#   chmod +x install.sh && ./install.sh
#
# 也可以通过 URL 一键安装：
#   curl -fsSL <此脚本URL> | bash
# ============================================================

SKILL_DIR="$(cd "$(dirname "$0")" && pwd)"
DEST_DIR="$HOME/.claude/skills"

echo "📦 安装 insert-citations skill …"
echo "   安装来源: $SKILL_DIR"
echo "   安装目标: $DEST_DIR"
echo ""

# 1. 创建 skills 目录
mkdir -p "$DEST_DIR"

# 2. 安装主 skill（用户调用）
rm -rf "$DEST_DIR/insert-citations"
ln -sfn "$SKILL_DIR" "$DEST_DIR/insert-citations"
echo "  ✅ insert-citations  →  $DEST_DIR/insert-citations"

# 3. 安装子 skill（模型调用）
rm -rf "$DEST_DIR/citation-hunting"
ln -sfn "$SKILL_DIR/citation-hunting" "$DEST_DIR/citation-hunting"
echo "  ✅ citation-hunting   →  $DEST_DIR/citation-hunting"

# 4. 安装 Python 依赖
echo ""
echo "📦 安装 Python 依赖 (python-docx) …"
pip3 install python-docx --quiet 2>&1 | tail -1 || echo "  ⚠️  pip3 install 失败，请手动执行: pip3 install python-docx"

echo ""
echo "═══════════════════════════════════════════"
echo "  ✅ 安装完成！"
echo ""
echo "  现在在对话中说："
echo "    \"给论文.docx 添加交叉引用\""
echo "    \"按出现顺序排序参考文献\""
echo "    \"帮我找这篇论文的真实参考文献\""
echo ""
echo "  或直接运行脚本："
echo "    python3 $SKILL_DIR/scripts/insert_citations.py 论文.docx"
echo "═══════════════════════════════════════════"
