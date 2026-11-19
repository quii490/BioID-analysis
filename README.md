# BioID analysis suite

Five skills that turn protein-level BioID mass-spec exports into a reviewable interactor list, a
batch-QC workbook, a self-contained HTML report, pairwise repeat comparisons, and per-protein trend
plots.

| Skill | What it does |
|---|---|
| `bioid` | Preprocessing: explicit bait→EV channel map, contaminant removal, log2FC, 5th-percentile floor |
| `bioid-batch-qc` | Cross-bait background flags, sample similarity / carryover screening, bait recovery and transfer |
| `bioid-analysis` | Load-corrected fold changes, SAINT-style score + calibrated q, gates, candidate table + HTML report |
| `bioid-repeat` | Compare the same bait across two experiments (Spearman/Pearson r, one-sided prey) |
| `bioid-trend` | One or a few proteins across many BioIDs (grouped x bands, mean line, missing at 0, vector PDF) |

`skills/bioid/scripts/bioid_run.py` runs the main chain in one command (preprocess → batch QC → bait
confirmation → scoring → verification) and refuses to produce a report without batch QC.

## Install

```bash
git clone https://github.com/quii490/BioID-analysis.git
cd BioID-analysis
bash doctor.sh      # checks R / Python / optional poppler
bash install.sh     # links the five skills into ~/.dsh/skills and ~/.codex/skills
```

Requirements: `Rscript` with `readxl` + `openxlsx` (preprocessing only); Python >= 3.9 with `numpy` +
`openpyxl`. Every skill is a plain directory (`SKILL.md` + `scripts/` + optional `assets/`, `tests/`),
so the scripts also run standalone without a skill loader.

## Quick start

```bash
python3 skills/bioid/scripts/bioid_run.py /abs/path/raw --out /abs/path/run
# prints the inferred channel map (and, for construct names, the label->gene table) and waits for --yes
python3 skills/bioid/scripts/bioid_run.py /abs/path/raw --out /abs/path/run \
    --resume --yes --bait-alias /abs/path/run/00_bait_aliases.tsv
```

The full operating manual (data contract, confirmation gates, every flag, interpretation rules, FAQ)
is the Chinese SOP below. Release history: [`CHANGELOG.md`](CHANGELOG.md).

---

## 中文 SOP（完整手册）

这套东西把「一堆蛋白级质谱导出」变成「可复核的候选互作蛋白列表 + 批次 QC + 图」。共 5 个 skill，
其中 3 个组成主链、2 个是只读的下游视图：

```text
raw 导出（每个诱饵一个 .xlsx）
        │
        ├─ bioid            预处理：通道映射、去污染、log2FC、+5% 下限
        │        ↓  *_processed.xlsx
        ├─ bioid-batch-qc   批次 QC：跨诱饵背景频率、交叉污染/串样、诱饵回收
        │        ↓  flagged/*_flagged.xlsx（追加 14 列，4 列可见）+ BioID_batch_QC.xlsx
        └─ bioid-analysis   打分：载荷校正、SAINT 风格分数/q、门控、候选表 + HTML 报告

下游只读视图（吃上面任意一种工作簿）：
   bioid-repeat   任意两份工作簿两两比较（重复性 / 细胞系差异）
   bioid-trend    同一（组）蛋白在多个 BioID 里的 log2FC 折线图
```

一条命令跑完主链的编排器：`skills/bioid/scripts/bioid_run.py`。

---

## 0. 快速开始（3 分钟）

```bash
bash doctor.sh          # 检查 R / Python / 依赖，并打印可用的 PY

# 全链：预处理 → 批次 QC → 诱饵确认 → 打分报告 → 验收
python3 skills/bioid/scripts/bioid_run.py /abs/path/raw --out /abs/path/runs/2026-11-20
# 它会先打印每份工作簿推断出的诱饵 + 证据，确认无误后加 --yes 重跑（已完成的阶段用 --resume 跳过）
```

---

## 版本说明（v1.0.0，2026-11-20）

这是 5 个 skill 的统一发行版。相对旧的单 skill 包：

- 五个 skill 齐全：预处理 `bioid`、批次 QC `bioid-batch-qc`、打分 `bioid-analysis`、
  两两比较 `bioid-repeat`、多 BioID 趋势 `bioid-trend`。
- 新增编排器 `skills/bioid/scripts/bioid_run.py`：一条命令跑完主链并做验收；没有它，批次 QC 很容易被跳过。
- `bioid-analysis` **默认要求 batch-QC 输出**（`--allow-no-qc` 才放行），缺 QC 时控制台明确告警；
  `flagged/` 覆盖不全时点名缺失条件。
- 诱饵推断失败不再静默：控制台 + Overview 都有 WARNING，并给出最富集的候选。
- 修了 batch-qc 的 Bait Transfer 在 ≥13 条件时明细表头覆盖矩阵末行、可达性检查、对照列全空崩溃、
  诱饵方向自检等问题。
- 两两比较（repeat）与多 BioID 折线图（trend）都有完整文档和示例（见第 5.4/5.5 节）。

---

## 1. 这个目录是什么、和线上 skill 的关系

| 路径 | 内容 |
|---|---|
| `skills/<名字>/` | 5 个 skill 的分发副本（每个都有自己的 `SKILL.md`、`scripts/`、`assets/`、`tests/`） |
| `doctor.sh` | 环境自检（R 包、Python 依赖、可选 poppler） |
| `install.sh` | 把 5 个 skill 装到 `~/.dsh/skills` / `~/.codex/skills`（`--copy` 可复制） |
| `VERSION` / `CHANGELOG.md` | 版本号与变更记录 |

agent 实际加载的是 `~/.dsh/skills/<名字>/`（或 `$CODEX_HOME/skills/`）。克隆后安装：

```bash
bash install.sh                                      # 链接到 ~/.dsh/skills 与 ~/.codex/skills
bash install.sh --copy --target ~/.dsh/skills        # 复制到指定目录，脱离本仓库独立使用
bash install.sh --force                              # 覆盖同名旧版本（旧的备份为 <名字>.bak-<时间戳>）
```

> 也可以不安装：直接让 agent 用本仓库 `skills/` 里的脚本；只是 skill 目录里的自动发现
> （描述、示例、注意事项）不会被加载。

---

## 2. 数据契约（开始前必须满足）

### 2.1 原始导出

- 每个诱饵一个 `.xlsx`，蛋白级（不是肽段级）。
- 必须有至少两列原始丰度，表头形如 `Abundance: F1: Sample` / `Abundance: F9: Control`。
  不要用 `Abundances (Normalized):` 列——三个下游 skill 都只认原始 `Abundance:` 列。
- 常见列：`Accession`（UniProt）、`GN`、`Description`、`Coverage [%]`、`# Peptides`、`# PSMs`、
  `# Unique Peptides`。

### 2.2 channel_map.tsv（人来决定，但不用手写）

**`bioid_run.py` 会自动生成草稿并在同一个闸门里让你确认**：输入目录里没有 `channel_map.tsv`、
也没给 `--channel-map` 时，它读每份工作簿的 `Abundance:` 列、按文件名找到诱饵蛋白、看它在哪个通道
最富集，写出 `<out>/00_channel_map.tsv`（每条规则上方注释里是证据），然后停下：

```text
推断的通道映射（草稿，尚未运行）：
  文件              诱饵  EV   置信     证据
  ATOSB-N.xlsx      F1    F9   medium  ATOSB 390x vs F9
  ATOSB-C.xlsx      F2    F9   medium  ATOSB 1525x vs F9
  ...
草稿已写到 /abs/path/run/00_channel_map.tsv（可直接编辑；确认无误后加 --yes 重跑）
```

- 置信度：`high` = 标签在 GN 里精确命中；`medium` = 只命中前缀（`ATOSB-C` → `ATOSB`，构造名常见）；
  `low` = 文件名里没有基因名，按"出现文件数最少的通道"推断（对照会在多份文件里重复，所以通常仍对）。
- 如果标签蛋白在"跨文件重复出现的通道"里反而更高，草稿会判定疑似列反：选该文件独有的通道并标注，
  证据里会写"请确认是不是诱饵/对照列反了"。
- **草稿永远不会被静默使用**：确认后加 `--yes` 重跑；想改就先编辑 `00_channel_map.tsv` 再跑
  （编辑内容会被沿用）；改了输入想重新推断，删掉这个文件即可。
- **构造名会单独问你**：文件名不是标准基因符号时（`ATOSB-C`、`FAM118B-L32R`、`EEF1AKMT4-Y39W41A+F204206A`
  或旧名），闸门会多写一张 `<out>/00_bait_aliases.tsv`（每行 `<标签> <基因>`）并打印
  “需要你确认的 构造名 → 基因名”，等你确认/修改后再 `--yes`。这张表会同时给 QC（诱饵回收、转运、
  方向自检）和分析（诱饵排除/标记）——没有它，QC 会把每个构造都报成 “Bait recovery QC failure”。
  纯基因名的文件名不会触发这个问题、也不会生成别名表。
- 你已经有一张手写的 map（输入目录里的 `channel_map.tsv`，或 `--channel-map` 指定的文件）时，
  以它为准，不弹草稿闸门。

手写时的格式（一行一个诱饵通道）：

```text
# 诱饵通道   EV 通道[|EV2...]   文件名
F4          F1                 1AKAP8L.xlsx
F8          F6                 8NCOA3.xlsx
F12         F11|F13            12BRD4.xlsx      # 两个 EV → 每个 EV 各出一份工作簿
```

- 通道号就是表头里 `Abundance: F<n>:` 的 `<n>`。
- **不要靠猜**：搞反了不会报错，只会得到一张干净但方向错误的表。确认方法是：诱饵蛋白应该只在
  自己文件里高、对照通道应该出现在很多文件里。
- `bioid` 跑完会打印 bait→EV 表，先看这张表。

### 2.3 命名约定（决定条件名与诱饵名）

- `bioid` 产物：`<序号><基因>_processed.xlsx`（如 `1SPO16_processed.xlsx`）。一个诱饵配多个 EV 时
  用不同序号前缀、同一基因名（`1SPO16` / `2SPO16`），这样它们会归成同一个 family。
- 下游 `bioid-batch-qc` 的产物是 `<原名>_flagged.xlsx`。
- 文件名里的基因名会被当作诱饵；**构造名/旧名要单独映射**：
  - `--bait-alias`（analysis、repeat 用）：`ATOSB-F405K → ATOSB`、`FAM214B → ATOSB`；
  - `--bait-map`（analysis）：直接按文件名覆盖，`ATOSB-C.xlsx → ATOSB`。

### 2.4 诱饵确认（不能跳过）

```bash
python3 skills/bioid-analysis/scripts/bioid_analysis.py /abs/path/batch --list-baits
```

每份工作簿打印：推断的诱饵标签、两个通道、该标签是否真的检出（带 unique 肽），以及诱饵通道里
最富集的几个蛋白作为证据。确认后才会正式分析。`bioid_run.py` 把这一步做成了闸门（不加 `--yes` 就停）。

---

## 3. 能做哪些分析（决策表）

| 你想要的 | 用哪个 | 入口 |
|---|---|---|
| 新数据从 0 到候选表 + 报告 | `bioid_run.py` 全链 | `skills/bioid/scripts/bioid_run.py` |
| 只把原始导出变成可分析工作簿 | `bioid` | `skills/bioid/scripts/bioid_process.R` |
| 看这批有没有串样/污染、谁背景高 | `bioid-batch-qc` | `skills/bioid-batch-qc/scripts/bioid_batch_review.py` |
| 出一份可交互的候选互作蛋白报告 | `bioid-analysis` | `skills/bioid-analysis/scripts/bioid_analysis.py` |
| 同一个诱饵在两次实验里重复得好不好 | `bioid-repeat` | `skills/bioid-repeat/scripts/bioid_repeat.py` |
| 某几个蛋白在多个 BioID 里怎么变 | `bioid-trend` | `skills/bioid-trend/scripts/bioid_trend.py` |

---

## 4. 新数据来了，怎么跟 agent 说（可直接复制）

> 下面每段都可以整段粘给 agent。把 `/abs/path/...` 换成你的真实绝对路径。agent 应该先读对应
> `SKILL.md`，再按里面的流程跑；如果它跳过了诱饵确认或 QC，直接把模板里的要求贴回去。

### 4.1 新的一批原始数据，跑完整条链（最常用）

```text
用 BioID 套件分析这批新数据：
- 原始导出目录：/abs/path/raw
- 输出目录：/abs/path/runs/2026-11-20
- 我还没写 channel_map，让 bioid_run.py 自动推断草稿，把推断出的通道映射和证据表给我确认

要求：
1. 先停在通道映射/诱饵闸门，把每个文件推断的诱饵通道、EV 通道、置信度和证据给我看；
   如果文件名是构造名/旧名，把“构造名 → 基因名”的候选映射也列出来问我确认，不要直接跑；
2. 我确认（或改完别名表）后加 --yes 重跑；
3. 如果诱饵标签是构造名（如 ATOSB-F405K）或旧名（FAM214B），用 --bait-alias 映射到数据库符号；
4. 跑完把验收结果（条件数、QC 是否接上、每个条件是否都标到诱饵）、报告路径、所有 WARNING 贴给我。
```

agent 应做的事：写/检查 channel_map → `bioid_run.py` → 停在诱饵闸门 → 你确认 → `--resume --yes` →
验收。若你还没写 channel_map，让它先按表头 `Abundance: F<n>` 和你给的对照通道生成草稿，再人工核对。

### 4.2 数据已经预处理过，只做批次 QC

```text
这批工作簿已经用 bioid 处理过了，只做批次 QC：
目录：/abs/path/processed（每份一个条件，文件名是 <序号><基因>_processed.xlsx）
用 bioid-batch-qc，输出到默认的 flagged/。
跑完告诉我：条件数、橙色/黄色背景旗各多少、哪些条件诱饵未回收、相似度最高的几对、
以及阈值可达性警告（批次太小时会有）。
```

### 4.3 QC 已跑完，只要候选表和报告

```text
QC 已经跑完：/abs/path/batch（flagged/ 在里面）。
用 bioid-analysis 出候选表和 HTML 报告，标准参数（1.5×、≥2 unique 肽）。
条件名是构造名，诱饵别名表在 /abs/path/bait_aliases.tsv。
先用 --list-baits 给我看诱饵，再正式跑；报告出来告诉我每个条件的候选数、top 候选、所有 WARNING。
```

### 4.4 两两比较（重复性 / 两个批次 / 两种细胞系）

```text
帮我比较这几对实验的重复性（bioid-repeat）。配对表写在 /abs/path/pairs.tsv：
每行是 <标签>\t<文件A>\t<文件B>，标签里用 "vs" 分隔两边（例如 "ATOSB: 293 vs MCF7"）。

先用 --list 把每对能匹配多少蛋白、用了哪一列、标签是否一致给我看；我确认后再生成报告。
报告里要能看到 Spearman r、Pearson r、拟合斜率，以及只在单侧检出的蛋白。
```

没有配对表也可以直接说，例如：

```text
用 bioid-repeat 比较这两个工作簿（同一个诱饵的两次实验）：
A=/abs/path/batch1/flagged/1AKAP8L_flagged.xlsx
B=/abs/path/batch4/flagged/9AKAP8L_processed_flagged.xlsx
先 --list，再出报告，并把 r 和只在一侧出现的蛋白数告诉我。
```

### 4.5 多个 BioID 一起看某几个蛋白怎么变

```text
我想看这几个蛋白在不同 BioID 里的标准化 log2FC 变化（bioid-trend）：
- PPP2R1A、PPP2R1B、PPP2CA、PPP2CB、PPP2R2A、PPP2R2D，按 A scaffold / C catalytic / B55 regulatory 分组
- BioID 按这个顺序：FL、N、C、F285W/W286A、F405K、F458A、R488/R490E
- 其中 FL 在 /abs/path/BioID-2nd/flagged/21FAM214B_flagged.xlsx，其余在 /abs/path/BioID-5th/analysis/bioid/flagged/
- FL 单独一组，N/C 是 truncation 一组，四个突变体是 mutation 一组（背景色区分）
- 未检出的按 0 画空心点，线照常连；再加一条跨蛋白平均线（--mean）

先 --check 把数值矩阵和 NA 给我看，确认后出 PDF（矢量），数值表也给我。
```

如果只是临时看两三个文件，可以更短：

```text
用 bioid-trend 画：文件 A、B、C 三个 BioID，看基因 PPP2R1A 和 PPP2CA 的变化，先 --check 再出图。
```

---

## 5. 每个 skill 的功能与参数

### 5.1 `bioid`（预处理，R）

```bash
Rscript skills/bioid/scripts/bioid_process.R <输入目录|文件> <输出目录> \
  [--map channel_map.tsv] [--sheet Proteins] [--contaminants 清单.tsv] [--force]
```

- 作用：按 `channel_map.tsv` 取诱饵/对照两列 → 去污染 → 算 `LOG2FC` → 对照缺失时用对照通道 5 分位
  做下限 → 写 `*_processed.xlsx`（`+5%` / `Proteins` / `Raw` / `Removed` / `Methods`）。
- `--force`：输入本身已含 `LOG2FC`（旧版流程产物）时允许重跑；原始两列仍在才安全。
- 退出码：`0` 干净；`2` 有文件被跳过（会列出原因）；`3` 输出包结构有问题。
- 不做：批次 QC、打分、报告。

### 5.2 `bioid-batch-qc`（批次 QC，Python）

```bash
python3 skills/bioid-batch-qc/scripts/bioid_batch_review.py <folder> \
  [--out DIR] [--exempt-genes "GENE1,GENE2"]
```

- 输入：`bioid` 产物（每个条件一个工作簿）。
- 输出：`<folder>/flagged/*_flagged.xlsx`（原表 + 14 列；可见的 4 列是 `Norm_LOG2FC`、
  `Combined_score`、`Batch_BG_flag`、`Cross_sample_QC`）+ `BioID_batch_QC.xlsx`（Batch Overview、
  Background Candidates、Similarity Matrix、Pairwise Similarity、Bait Transfer、Family 对比、Methods）。
- 背景旗规则（硬编码，固定阈值）：
  - 橙 `High background suspect`：family 内 combined 排名前 2% 出现在 ≥4 个 bait family；
  - 黄 `Review background suspect`：非橙，且 family 内 combined 前 5% ≥4 个 family，或富集前 5% ≥5 个。
  - 批次 <4 个 family 时规则不可达，会明确警告——**空旗不等于干净**。
- 交叉样本：相似度、bait transfer（诱饵转运）、串样/混样分类；共享同一个 EV 的成对比较被排除在
  阈值标定外（否则是循环论证）。
- 诱饵回收失败、诱饵方向疑似反了（诱饵在对照侧高 >2 倍）、对照列整列为空，都会告警/跳过并在
  QC 工作簿的 Methods 里留记录。

### 5.3 `bioid-analysis`（打分 + 报告，Python）

```bash
python3 skills/bioid-analysis/scripts/bioid_analysis.py <folder> \
  [--from-qc] [--no-qc] [--allow-no-qc] [--allow-partial-qc] \
  [--bait-map 映射.tsv] [--bait-alias 别名.tsv] [--list-baits] \
  [--min-fold 1.5] [--min-unique-peptides 2] [--normalize median|robust|total|none] \
  [--out DIR] [--smoke] [--sheet Proteins]
```

- 输入：`bioid` 产物；若目录里有 `flagged/`（或 `--from-qc`）则自动用 QC 产物。
- **默认要求 QC**：没有 batch-QC 输出会直接报错（`--allow-no-qc` 放行；`--allow-partial-qc`
  放行覆盖不全的 `flagged/`）。这是为了防止「没有 Background flag 的报告被当成完整结果」。
- 处理：载荷校正（median/robust/total）→ 对照缺失给 5 分位下限（标 `FC_lower_bound`，不参与 p/q）
  → CRAPome 过滤 → SAINT 风格分数 + BH q 值 → 门控（≥1.5×、≥2 unique 肽、非污染物、非诱饵）。
- 输出：`<folder>/bioid_analysis/BioID_analysis_summary.xlsx`（Overview、Candidates、每条件全表、
  Methods）+ `BioID_analysis_report.html`（自包含，可交互散点、搜索标记、每个点 hover 出基因/
  蛋白名/FC/肽段/q/通过原因）+ `analysis_payload.json`。
- `--smoke`：只分析第一个条件，迭代报告模板用；默认写到 `bioid_analysis_smoke/`，不会覆盖正式报告。
- 按门控通过的候选是**排序**，不是显著性结论；q 值只在单次运行内部校准。

### 5.4 `bioid-repeat`（两两比较，Python）

```bash
python3 skills/bioid-repeat/scripts/bioid_repeat.py --pairs pairs.tsv --list   # 先看配对
python3 skills/bioid-repeat/scripts/bioid_repeat.py --pairs pairs.tsv [--out DIR] \
  [--value norm_log2fc|log2fc|fc] [--label-top 8]
# 或临时一对： --pair FILE_A FILE_B（可重复）
```

- 配对表 `pairs.tsv`：每行 `<标签>\t<文件A>\t<文件B>`，`#` 注释；标签里用 `vs` 分隔两侧，
  `vs` 前面的文字会被当作诱饵（画紫色菱形）。
- 输入可以是 `*_processed.xlsx` 或 `*_flagged.xlsx`；值列默认 `Norm_LOG2FC`，也可 `LOG2FC`/`FC`。
- 输出 `<首个文件目录>/bioid_repeat/bioid_repeat_report.html`：每对一节，匹配蛋白数、Spearman r、
  Pearson r、拟合斜率/截距、截尾（对照未检出）点、只在单侧检出的蛋白表。
- 读法：r ≥0.8 好、0.5–0.8 中等、<0.5 弱；r 是重复性不是 p 值；两次实验各一次运行没有重复结构；
  跨批次不能比绝对倍数（各自批内中心化）。

### 5.5 `bioid-trend`（多 BioID 看蛋白变化，Python）

```bash
# 1) 先看矩阵（不写文件）
python3 skills/bioid-trend/scripts/bioid_trend.py --samples samples.tsv --proteins proteins.tsv --check

# 2) 出图
python3 skills/bioid-trend/scripts/bioid_trend.py --samples samples.tsv --proteins proteins.tsv \
  --mean --out out/ --name atosb_pp2a --xlabel "ATOSB construct" --title "..." \
  --legend-title "PP2A groups / proteins"
```

`samples.tsv`（行序 = x 轴顺序）：

```text
label	file	xgroup
FL	../../BioID-2nd/flagged/21FAM214B_flagged.xlsx	FL
N	../analysis/bioid/flagged/ATOSB-N_processed_flagged.xlsx	truncation
C	../analysis/bioid/flagged/ATOSB-C_processed_flagged.xlsx	truncation
F405K	../analysis/bioid/flagged/ATOSB-F405K_processed_flagged.xlsx	mutation
```

- `label` 是刻度；`file` 相对 samples 表所在目录；`sheet`（可选）用于 analysis summary 的某个条件页；
  `column` 指定确切数值列；`xgroup` 决定背景色带。
- 值列 `--value auto`：`Norm_LOG2FC` → `log2FC (EV-corrected)` → `LOG2FC`，逐文件打印实际列。

`proteins.tsv`：

```text
gene	group	label	style	color	accession
PPP2R1A	A scaffold
PPP2R1B	A scaffold
PPP2CA	C catalytic
```

- 颜色=蛋白，marker 形状=group，线一律实线；未检出画在 y=0 的空心点，**接触未检出点的线段变虚线**。
- `--mean`：每个 x 上跨蛋白平均（未检出按 0 参与，图注会写明这点）。
- 输出 `<out>/<stem>/<stem>.pdf`（矢量）+ `.png` + `data/<stem>_values.tsv`（含 `NA` 与 `MEAN` 行）
  + `data/<stem>_sources.tsv`（每个文件用了哪一列）。
- 其它常用：`--ylim "0,10"`、`--yticks`、`--fig-width/--fig-height`、`--rotate-labels auto|0|30`、
  `--no-caption`、`--keep-all-na`、`--flat`（不建子目录）。
- 快捷：`--point FL=a.xlsx --point N=b.xlsx --genes PPP2R1A,PPP2CA`，不用写表。
- `--inspect 文件 --sheet 条件名` 可以先看某个工作簿的表头与实际取值列。

---

## 6. 产物在哪、先看什么

全链（`bioid_run.py`）的输出：

```text
<run>/
  01_processed/      *_processed.xlsx（bioid 产物）
  02_qc/             *_flagged.xlsx + BioID_batch_QC.xlsx
  03_analysis/       BioID_analysis_summary.xlsx / BioID_analysis_report.html / analysis_payload.json
  logs/              01_preprocess.log 02_qc.log 03_bait_check.log 04_analysis.log
  run_manifest.json  每步命令、耗时、验收结果
```

阅读顺序：

1. `03_analysis/BioID_analysis_report.html` 的 Overview（条件数、诱饵回收、载荷校正、警告）；
2. 每个条件的散点 + 候选表（先看 `Passes gates` 和 `q (BH)`，再看 `FC type` 是否下界）；
3. `02_qc/BioID_batch_QC.xlsx` 的 `Background Candidates`（橙/黄名单）、`Similarity Matrix`
   （串样）、`Bait Transfer`（诱饵转运）、`Methods`（阈值可达性、方向告警）；
4. 需要跨实验比较 → `bioid-repeat`；需要跨条件看具体蛋白 → `bioid-trend`。

---

## 7. 读结果的硬口径（别跳过）

- **单次运行没有重复**：一个诱饵对一个 EV，没法学 SAINTexpress 的 p 值；`q (BH)` 只在本批的负半
  分布内校准，介于 0.05 的条目是"排序靠前"，不是"已重复验证的互作"。
- **FC 下界**：对照未检出时用 5 分位下限算出的倍数标 `FC_lower_bound`；它是下界，不参与 p/q，
  也不能当测量值。
- **背景旗是复核优先级**：橙/黄不等于假阳性；共用定位、真实共同邻居都会反复出现。
- **跨批次不可比绝对值**：`Norm_LOG2FC` 在各自批内中心化，只能比同一蛋白沿条件的相对变化。
- **诱饵未检出 ≠ 实验失败**：可能是库/物种/表达问题；先看 `--list-baits` 的证据，再怀疑实验。
- **"Pass all gates" 不是显著性**：它只是 1.5× + ≥2 unique 肽 + 非污染物 + 非诱饵；单次运行里
  1.5× 门本身有 20–40% 的偶然通过率（报告里有这张卡）。

---

## 8. 常见问题

| 现象 | 原因 / 处理 |
|---|---|
| `没有找到 batch-QC 输出…`（analysis 直接退出） | 先跑 `bioid-batch-qc`。直接跑 analysis 时确实只要打分就加 `--allow-no-qc`（报告会标注缺 QC）；用 `bioid_run.py` 时它强制先跑 QC |
| `batch-QC 输出缺这些条件…` | QC 之后又加了工作簿；重跑 QC，或 `--allow-partial-qc` 只分析已 QC 的部分 |
| `bait 'XXX' was not found…` | 文件名里的诱饵名不是数据库符号：用 `--bait-alias`（构造名→基因）或 `--bait-map`（文件名→基因） |
| `bait … higher in the control channel — the two Abundance columns may be reversed` | 诱饵/对照列可能反了；检查 channel_map |
| `At least two raw 'Abundance:' columns …` | 输入不是原始导出，或旧 openxlsx 把 dimension 写成 A1；用 `bioid --force` 重跑生成 |
| `no usable workbook`（batch-qc） | 所有输入都被跳过；看日志里每份的原因 |
| 中文标签在图中变成空白 | 本套件的 `bioid_trend` 会自动切 CJK 字体；其它图请用英文标签 |
| `python3` 缺 numpy/openpyxl | `bash doctor.sh` 找一个可用的解释器，或 `python3 -m pip install --user numpy openpyxl` |
| 报告 HTML 很大（7–10 MB） | plotly 内联，属正常；`--plotly cdn` 可减小但需要联网 |

退出码（`bioid_run.py`）：`0` 通过 / `2` 用法或预检失败 / `3` 等诱饵确认（看完证据加 `--yes`）/
`4` 某阶段失败 / `5` 验收失败（条件数、QC 接没接上、诱饵标没标到）。

---

## 9. 维护

- **改 skill**：改 `~/.dsh/skills/<名字>/`，跑该 skill 的测试（`bioid-analysis/tests/`、
  `bioid-batch-qc/tests/`），再把改动同步回本目录 `skills/`。
- **旧版本**：进 `archive/<YYYYMMDD>/`，不删。
- **打包给别人**：直接压本目录（`skills/` + `README.md` + `doctor.sh`）。
- 各 skill 的详细规则以各自的 `SKILL.md` 为准；本 README 是入口和使用说明，不替代 `SKILL.md`。
