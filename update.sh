#!/usr/bin/env bash
set -euo pipefail

# ============================================================
# insert-citations skill 更新脚本
# ============================================================
# 用法：
#   ./update.sh              # 手动更新
#   ./update.sh --schedule   # 设置每天自动检查更新（cron）
#   ./update.sh --unschedule  # 取消自动检查
# ============================================================

SKILL_DIR="$(cd "$(dirname "$0")" && pwd)"

echo "🔍 insert-citations skill 更新检查"
echo "   目录: $SKILL_DIR"
echo ""

# ── 检查是否是 git 仓库 ──
if ! git -C "$SKILL_DIR" rev-parse --git-dir &>/dev/null; then
    echo "❌ $SKILL_DIR 不是 git 仓库，无法自动更新"
    echo "   请重新 git clone 安装"
    exit 1
fi

# ── 检查是否有 remote ──
REMOTE_URL=$(git -C "$SKILL_DIR" remote get-url origin 2>/dev/null || echo "")
if [ -z "$REMOTE_URL" ]; then
    echo "❌ 未配置 remote origin，请手动设置"
    exit 1
fi
echo "   remote: $REMOTE_URL"

# ── 处理调度参数 ──
case "${1:-}" in
    --schedule)
        CRON_LINE="0 9 * * * cd '$SKILL_DIR' && git fetch origin -q && git diff --quiet main origin/main || '$SKILL_DIR/update.sh' --quiet"
        # 检查是否已存在
        if crontab -l 2>/dev/null | grep -qF "insert-citations.*update.sh"; then
            echo "   ⚠️  自动更新已存在"
            exit 0
        fi
        (crontab -l 2>/dev/null || true; echo "$CRON_LINE") | crontab -
        echo "   ✅ 已设置每天 9:00 自动检查更新"
        echo "   取消: $0 --unschedule"
        exit 0
        ;;
    --unschedule)
        if crontab -l 2>/dev/null | grep -vF "insert-citations" | crontab - 2>/dev/null; then
            echo "   ✅ 已取消自动更新"
        else
            echo "   ⚠️  未找到自动更新任务"
        fi
        exit 0
        ;;
esac

# ── 拉取更新 ──
echo ""
echo "📥 拉取远程更新..."
git -C "$SKILL_DIR" fetch origin

LOCAL=$(git -C "$SKILL_DIR" rev-parse main)
REMOTE=$(git -C "$SKILL_DIR" rev-parse origin/main)

if [ "$LOCAL" = "$REMOTE" ]; then
    echo "   ✅ 已是最新版本 ($LOCAL)"
    exit 0
fi

echo "   发现新版本！"
echo "   $LOCAL → $REMOTE"
git -C "$SKILL_DIR" pull origin main

echo ""
echo "═══════════════════════════════════════════"
echo "  ✅ 已更新到最新版本"
echo ""
echo "  更新日志 (最近 5 条):"
echo "───────────────────────────────────────────"
git -C "$SKILL_DIR" log -5 --oneline --no-decorate
echo "═══════════════════════════════════════════"

# 如果以静默模式运行，无输出
if [ "${2:-}" = "--quiet" ]; then
    exit 0
fi
