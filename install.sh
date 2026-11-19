#!/usr/bin/env bash
# 把仓库里的 5 个 BioID skill 安装到 agent 的技能目录。
#   ./install.sh                   # 默认链接到 ~/.dsh/skills 与 ~/.codex/skills
#   ./install.sh --copy            # 复制而不是链接（想脱离本仓库独立使用）
#   ./install.sh --target DIR      # 只装到一个目录（可重复）
#   ./install.sh --force           # 覆盖同名条目（原条目备份为 <名字>.bak-<时间戳>）
#   ./install.sh --dry-run         # 只打印将要做什么
set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SRC="${ROOT}/skills"
SKILLS=(bioid bioid-batch-qc bioid-analysis bioid-repeat bioid-trend)
MODE="link"; FORCE=0; DRY=0; TARGETS=()

while [ $# -gt 0 ]; do
  case "$1" in
    --copy) MODE="copy" ;;
    --force) FORCE=1 ;;
    --dry-run) DRY=1 ;;
    --target) TARGETS+=("$2"); shift ;;
    -h|--help) sed -n '2,9p' "$0"; exit 0 ;;
    *) echo "未知参数：$1" >&2; exit 2 ;;
  esac
  shift
done

[ -d "$SRC" ] || { echo "找不到 ${SRC}" >&2; exit 1; }
if [ "${#TARGETS[@]}" -eq 0 ]; then
  TARGETS=("$HOME/.dsh/skills" "$HOME/.codex/skills")
fi

stamp="$(date +%Y%m%d-%H%M%S)"
for target in "${TARGETS[@]}"; do
  echo "→ $target"
  if [ ! -d "$target" ]; then
    if [ "$DRY" = "1" ]; then echo "   （dry-run）would mkdir -p $target"; continue; fi
    mkdir -p "$target"
  fi
  for skill in "${SKILLS[@]}"; do
    src="${SRC}/${skill}"; dst="${target}/${skill}"
    [ -d "$src" ] || { echo "   跳过 ${skill}（源缺失）"; continue; }
    if [ -L "$dst" ] && [ "$(readlink "$dst")" = "$src" ]; then
      echo "   = ${skill}（已指向本仓库）"; continue
    fi
    if [ -e "$dst" ] || [ -L "$dst" ]; then
      if [ "$FORCE" != "1" ]; then
        echo "   ! ${skill} 已存在，跳过（要覆盖加 --force）"; continue
      fi
      if [ "$DRY" = "1" ]; then echo "   （dry-run）would mv $dst $dst.bak-$stamp"; else
        mv "$dst" "${dst}.bak-${stamp}"; echo "   备份 ${skill} → $(basename "$dst").bak-$stamp"
      fi
    fi
    if [ "$DRY" = "1" ]; then
      echo "   （dry-run）would ${MODE} ${skill}"
    elif [ "$MODE" = "copy" ]; then
      cp -R "$src" "$dst"; echo "   + 复制 ${skill}"
    else
      ln -s "$src" "$dst"; echo "   + 链接 ${skill}"
    fi
  done
done

echo
echo "装好了。检查一下： bash doctor.sh"
echo "用法：对 agent 说“用 BioID 套件分析 /abs/path/raw”，或直接跑 README 里的命令。"
echo "skill 名：bioid / bioid-batch-qc / bioid-analysis / bioid-repeat / bioid-trend"
