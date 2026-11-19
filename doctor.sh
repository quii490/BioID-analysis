#!/usr/bin/env bash
# BioID 套件环境自检：找一个能跑 Python 阶段的解释器，检查 R 与可选工具。
# 用法： bash doctor.sh
set -u

ROOT="$(cd "$(dirname "$0")" && pwd)"
SKILLS="$ROOT/skills"
fail=0

echo "== 目录 =="
for s in bioid bioid-batch-qc bioid-analysis bioid-repeat bioid-trend; do
  if [ -d "${SKILLS}/${s}" ]; then
    echo "  ok   skills/${s}"
  else
    echo "  MISS skills/${s}"
    fail=1
  fi
done

echo
echo "== Python（analysis / batch-qc / repeat / trend 需要 numpy + openpyxl） =="
PY_FOUND=""
for cand in ${PYTHON:-} python3 /usr/bin/python3 /usr/local/bin/python3 python3.13 python3.12 python3.11; do
  [ -n "$cand" ] || continue
  command -v "$cand" >/dev/null 2>&1 || continue
  if "$cand" -c "import numpy, openpyxl" >/dev/null 2>&1; then
    PY_FOUND="$cand"
    "$cand" - <<'PY'
import sys, numpy, openpyxl
print(f"  ok   {sys.executable}  python {sys.version.split()[0]}  numpy {numpy.__version__}  openpyxl {openpyxl.__version__}")
PY
    break
  fi
done
if [ -z "$PY_FOUND" ]; then
  echo "  MISS 没有找到带 numpy+openpyxl 的 python"
  echo "       安装： python3 -m pip install --user numpy openpyxl"
  fail=1
fi

echo
echo "== R（预处理阶段需要 Rscript + readxl + openxlsx） =="
if command -v Rscript >/dev/null 2>&1; then
  echo "  ok   $(command -v Rscript)  $(Rscript --version 2>&1 | head -1)"
  if Rscript -e 'quit(status = if (requireNamespace("readxl", quietly=TRUE) && requireNamespace("openxlsx", quietly=TRUE)) 0 else 1)' >/dev/null 2>&1; then
    echo "  ok   R 包 readxl + openxlsx"
  else
    echo "  MISS R 包 readxl/openxlsx；安装： Rscript -e 'install.packages(c(\"readxl\",\"openxlsx\"))'"
    fail=1
  fi
else
  echo "  MISS Rscript（只跑 QC / 分析 / 比较 / 画图时可以不要 R，但要先有处理好的工作簿）"
fi

echo
echo "== 可选工具 =="
command -v pdftoppm >/dev/null 2>&1 && echo "  ok   poppler（PDF 转图自检）" || echo "  --   没有 poppler（可选： brew install poppler）"

echo
if [ -n "$PY_FOUND" ]; then
  echo "建议在 shell 里设："
  echo "  export PY=\"$PY_FOUND\""
  echo "  export BIOID_SKILLS=\"$SKILLS\""
fi
if [ "$fail" -ne 0 ]; then
  echo
  echo "有依赖未满足，先按上面的提示安装。"
  exit 1
fi
echo "环境检查通过。下一步：bash install.sh 安装，或读 README.md 的“新数据来了怎么说”一节。"
