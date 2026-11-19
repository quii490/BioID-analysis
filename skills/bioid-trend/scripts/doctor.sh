#!/usr/bin/env bash
# 找一台机器上能跑 bioid-trend 的 python：需要 matplotlib + openpyxl + numpy。
# 用法： bash scripts/doctor.sh
set -u

candidates="${PYTHON:-} python3 /usr/bin/python3 /usr/local/bin/python3 python3.11 python3.12 python3.13"
seen=""
for cand in $candidates; do
  [ -n "$cand" ] || continue
  case " $seen " in *" $cand "*) continue ;; esac
  seen="$seen $cand"
  command -v "$cand" >/dev/null 2>&1 || continue
  if "$cand" -c "import matplotlib, openpyxl, numpy" >/dev/null 2>&1; then
    "$cand" - <<'PY'
import sys, matplotlib, openpyxl, numpy
print(f"可用: {sys.executable}")
print(f"  matplotlib {matplotlib.__version__}  openpyxl {openpyxl.__version__}  numpy {numpy.__version__}")
PY
    echo "建议: export PY=\"${cand}\"   （或在命令里直接用 ${cand}）"
    command -v pdftoppm >/dev/null 2>&1 || echo "提示: 没装 poppler，无法把 PDF 转图自检（brew install poppler）"
    exit 0
  fi
done

echo "没有找到带 matplotlib 的 python。" >&2
echo "装依赖（任选其一）：" >&2
echo "  python3 -m pip install --user matplotlib openpyxl numpy" >&2
echo "  /usr/bin/python3 -m pip install --user matplotlib openpyxl numpy" >&2
exit 1
