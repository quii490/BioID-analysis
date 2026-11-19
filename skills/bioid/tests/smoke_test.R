#!/usr/bin/env Rscript
# bioid/tests/smoke_test.R —— 端到端冒烟测试（只用 base R + openxlsx）
#
# 造一个合成工作簿，覆盖脚本里几处容易悄悄坏掉的约定：
#   * 输入里叠着肽段级行（Master 为空 / 第二套表头），必须被当作非蛋白行剔除；
#   * 丰度列顺序故意打乱（EV 在前、诱饵在中间），输出必须重排成「诱饵、本组合EV、其余」；
#   * 一个诱饵对两个 EV → 两份工作簿；
#   * Keratin 行、Reverse 行、无 GN 行被删且出现在 Removed 表里；
#   * GN 的分号格式（"GN=DDD; PE=1"）不能带出分号；
#   * Unique 行的 +5% 表用对照组第 5 百分位填补。
#
# 用法: Rscript smoke_test.R

suppressMessages(library(openxlsx))

script <- normalizePath(file.path(dirname(sub("^--file=", "",
  grep("^--file=", commandArgs(FALSE), value = TRUE)[1])), "..", "scripts", "bioid_process.R"))
work <- file.path(tempdir(), "bioid_smoke")
unlink(work, recursive = TRUE)
dir.create(file.path(work, "in"), recursive = TRUE)

# ── 合成输入 ────────────────────────────────────────────────────────────────
# 诱饵 = F5；EV = F1 与 F2。输入列序刻意把 EV 放前面。
# 行计划（刻意把隐藏行放在中间，让「全局行号」与「可见段位置」不重合）：
#   1 P1 正常        2 P9 诱饵无值(隐藏)   3 P2 唯一(诱饵有值/EV无值)
#   4 P3 Keratin     5 P4 抽不到 GN        6 P5 分号格式 GN
#   7 P6 Reverse 标记 8 P7 肽段级行(Master=NA)  9 P8 肽段级行(Master=NA)
df <- data.frame(
  Master = c("Master Protein", "Master Protein", "Master Protein", "Master Protein",
             "Master Protein", "Master Protein", "Master Protein", NA, NA),
  Accession = c("P1", "P9", "P2", "P3", "P4", "P5", "P6", "P7", "P8"),
  Description = c(
    "Protein one OS=Homo sapiens GN=AAA PE=1 SV=1",
    "Protein nine OS=Homo sapiens GN=GGG PE=1 SV=1",
    "Protein two OS=Homo sapiens GN=BBB PE=1 SV=1",
    "Keratin, type I cytoskeletal 10 OS=Homo sapiens GN=KRT10 PE=1 SV=1",
    "Protein four with no gene name OS=Homo sapiens PE=1 SV=1",
    "RecName: Full=Protein five; GN=DDD; PE=1; SV=2",
    "Protein six OS=Homo sapiens GN=FFF PE=1 SV=1",
    "Confidence",
    "High"
  ),
  # 故意混入一个文本值：真实导出就是这么把 Coverage/# Peptides 整列变成文本的
  `# Unique Peptides` = c("5", "2", "3", "n/a", "2", "4", "2", NA, NA),
  `Abundance: F1: Sample` = c(100, 150, NA, 50, 70, 400, 60, NA, NA),
  `Abundance: F5: Sample` = c(1000, NA, 800, 500, 700, 1600, 65, NA, NA),
  `Abundance: F2: Sample` = c(200, 175, NA, 60, 80, 500, 70, NA, NA),
  Reverse = c(NA, NA, NA, NA, NA, NA, "+", NA, NA),
  check.names = FALSE, stringsAsFactors = FALSE
)
openxlsx::write.xlsx(df, file.path(work, "in", "synthetic.xlsx"), overwrite = TRUE)

writeLines(c("F5\tF1|F2"), file.path(work, "in", "channel_map.tsv"))

# ── 运行 ────────────────────────────────────────────────────────────────────
status <- system2("Rscript", c(script, file.path(work, "in"), file.path(work, "out")),
                  stdout = TRUE, stderr = TRUE)
cat(paste(status, collapse = "\n"), "\n\n")
code <- attr(status, "status")
if (is.null(code)) code <- 0L

# 数一数某张表里有多少行带浅绿底纹（Unique 标记），用于回归「排序后底纹错位」。
# 表名 → 部件按写入顺序固定：+5% = sheet1, Proteins = sheet2。
# 单元格的 s= 指向 cellXfs 条目，需经 fillId 才能落到 <fills>，不能直接用 s 当填充索引。
green_row_numbers <- function(path, part) {
  styles <- paste(readLines(unz(path, "xl/styles.xml"), warn = FALSE), collapse = "")
  fb <- regmatches(styles, regexec("<fills[^>]*>(.*?)</fills>", styles))[[1]]
  if (length(fb) < 2L) return(0L)
  fills <- regmatches(fb[[2]], gregexpr("<fill>.*?</fill>", fb[[2]]))[[1]]
  green_fills <- which(vapply(fills, function(f) grepl("C6EFCE", f, ignore.case = TRUE),
                              logical(1))) - 1L
  if (!length(green_fills)) return(0L)
  xb <- regmatches(styles, regexec("<cellXfs[^>]*>(.*?)</cellXfs>", styles))[[1]]
  if (length(xb) < 2L) return(0L)
  xfs <- regmatches(xb[[2]], gregexpr("<xf [^>]*/?>", xb[[2]]))[[1]]
  fill_id <- vapply(xfs, function(x) {
    m <- regmatches(x, regexec('fillId="([0-9]+)"', x))[[1]]
    if (length(m) >= 2L) as.integer(m[[2]]) else 0L
  }, integer(1))
  green_xf <- which(fill_id %in% green_fills) - 1L
  x <- paste(readLines(unz(path, part), warn = FALSE), collapse = "")
  rows <- regmatches(x, gregexpr("<row [^>]*>.*?</row>", x))[[1]]
  hit <- vapply(rows, function(r) {
    cells <- regmatches(r, gregexpr("<c [^>]*>", r))[[1]]
    any(vapply(cells, function(cc) {
      m <- regmatches(cc, regexec('s="([0-9]+)"', cc))[[1]]
      length(m) >= 2L && as.integer(m[[2]]) %in% green_xf
    }, logical(1)))
  }, logical(1))
  as.integer(sub('.*<row r="([0-9]+)".*', "\\1", rows[hit]))
}

green_rows <- function(path, part) length(green_row_numbers(path, part))

# 绿色底纹必须落在「诱饵有值、EV 无值」的那些行上（按工作表行号核对 GN）。
green_gn_values <- function(path, part, sheet_name, gn_col_name = "GN") {
  nums <- green_row_numbers(path, part)
  if (!length(nums)) return(character(0))
  df <- openxlsx::read.xlsx(path, sheet = sheet_name, colNames = TRUE, check.names = FALSE)
  df[[gn_col_name]][nums - 1L]
}

# 数一数某张表里有多少行被标记为 hidden（诱饵无值的行会分组隐藏）。
count_hidden_rows <- function(path, part) {
  x <- paste(readLines(unz(path, part), warn = FALSE), collapse = "")
  m <- regmatches(x, gregexpr('<row [^>]*hidden="1"[^>]*>', x))[[1]]
  length(m)
}

fails <- character(0)
check <- function(ok, label) {
  cat(sprintf("  %-4s %s\n", if (isTRUE(ok)) "PASS" else "FAIL", label))
  if (!isTRUE(ok)) fails <<- c(fails, label)
}

cat("断言:\n")
check(code == 0L, "退出码为 0（含包结构自检通过）")
outs <- sort(list.files(file.path(work, "out"), pattern = "_processed\\.xlsx$"))
check(length(outs) == 2L, sprintf("一个诱饵两个 EV → 2 份工作簿（实际 %d）", length(outs)))
check(identical(outs, c("synthetic_vsF1_processed.xlsx", "synthetic_vsF2_processed.xlsx")),
      "输出命名为 <stem>_vs<EV>_processed.xlsx")

for (o in outs) {
  ev <- sub(".*_vs(F[0-9]+)_processed\\.xlsx", "\\1", o)
  wb_path <- file.path(work, "out", o)
  wb <- openxlsx::loadWorkbook(wb_path)
  sheets <- names(wb)
  check(all(c("+5%", "Proteins", "Raw", "Removed", "Methods") %in% sheets),
        sprintf("%s: 五张工作表齐全", o))

  prot <- openxlsx::read.xlsx(wb_path, sheet = "Proteins", check.names = FALSE)
  abund <- grep("^Abundance", names(prot), value = TRUE)
  check(length(abund) == 3L && grepl("F5", abund[1]) && grepl(ev, abund[2]),
        sprintf("%s: 前两个丰度列为 诱饵F5、EV%s（实际 %s）", o, ev,
                paste(abund[1:2], collapse = " / ")))
  check(identical(sort(prot$Accession), c("P1", "P2", "P5", "P9")),
        sprintf("%s: 只留下 P1/P2/P5/P9（实际 %s）", o, paste(sort(prot$Accession), collapse = ",")))
  check(identical(prot$GN[prot$Accession == "P5"], "DDD"),
        sprintf("%s: 分号格式描述抽出 GN=DDD（实际 %s）", o, prot$GN[prot$Accession == "P5"]))
  check(!any(is.na(prot$GN)) && !any(prot$GN == ""), sprintf("%s: GN 列无空值", o))

  row_p1 <- prot[prot$Accession == "P1", ]
  f1 <- if (ev == "F1") 100 else 200
  check(abs(row_p1$LOG2FC - log2(1000 / f1)) < 1e-9,
        sprintf("%s: LOG2FC = log2(诱饵/%s)", o, ev))
  check(is.na(prot$LOG2FC[prot$Accession == "P2"]),
        sprintf("%s: Unique 行在 Proteins 表里 LOG2FC 为空", o))

  plus5 <- openxlsx::read.xlsx(wb_path, sheet = "+5%", check.names = FALSE)
  ev_vals <- if (ev == "F1") c(100, 400, 150) else c(200, 500, 175)
  p5 <- as.numeric(quantile(ev_vals, 0.05, type = 7, names = FALSE))
  ev_col <- grep("^Abundance", names(plus5))[2]
  check(abs(plus5[[ev_col]][plus5$Accession == "P2"] - p5) < 1e-6,
        sprintf("%s: +5%% 表里 Unique 行的对照值 = 第5百分位 %.3g", o, p5))

  raw <- openxlsx::read.xlsx(wb_path, sheet = "Raw", check.names = FALSE)
  check(nrow(raw) == nrow(df), sprintf("%s: Raw 表保留全部 %d 行输入", o, nrow(df)))

  rem <- openxlsx::read.xlsx(wb_path, sheet = "Removed", check.names = FALSE)
  reasons <- paste(rem$Reason, collapse = " | ")
  check(all(vapply(c("非蛋白级行", "污染物", "抽不到 GN", "Reverse"),
                   function(k) any(grepl(k, rem$Reason, fixed = TRUE)), logical(1))) &&
          nrow(rem) == 5L,
        sprintf("%s: Removed 记录 5 行（肽段级×2/污染物/无GN/Reverse，实际 %d 行: %s）",
                o, nrow(rem), substr(reasons, 1, 60)))
  check(all(!is.na(rem$RawSheetRow)), sprintf("%s: Removed 每行都有原始行号", o))

  # 测量列（被文本污染过的）必须被归一到数值型。
  # 注意：openxlsx 的读取器会把 "# Unique Peptides" 改写成 "#.Unique.Peptides"，
  # 所以按关键字匹配列名，而不是按精确名。
  up_col <- grep("Unique", names(prot), value = TRUE)[1]
  check(is.numeric(prot[[up_col]]) && is.numeric(plus5[[up_col]]),
        sprintf("%s: 被文本污染的 peptide 列已转成数值型（实际 %s / %s）", o,
                class(prot[[up_col]])[1], class(plus5[[up_col]])[1]))

  # 诱饵/EV 两列必须写成数值（Excel 里可排序），而不是从导出继承的文本
  apos <- match(abund, names(prot))
  check(is.numeric(prot[[apos[1]]]) && is.numeric(prot[[apos[2]]]),
        sprintf("%s: Proteins 表的诱饵/EV 两列是数值型（实际 %s）", o,
                paste(vapply(prot[apos[1:2]], function(z) class(z)[1], character(1)), collapse = "/")))
  ppos <- match(abund, names(plus5))
  check(is.numeric(plus5[[ppos[2]]]),
        sprintf("%s: +5%% 表的对照列是数值型（实际 %s）", o, class(plus5[[ppos[2]]])[1]))

  # Unique 底纹必须跟着排序走：+5% 表按补值后 LOG2FC 排序，标记也应映射到同一行
  check(green_rows(wb_path, "xl/worksheets/sheet1.xml") == 1L &&
          identical(green_gn_values(wb_path, "xl/worksheets/sheet1.xml", "+5%"), "BBB"),
        sprintf("%s: +5%% 表的浅绿底纹落在 Unique 行 BBB 上（实际 %s）", o,
                paste(green_gn_values(wb_path, "xl/worksheets/sheet1.xml", "+5%"), collapse = ",")))
  check(identical(green_gn_values(wb_path, "xl/worksheets/sheet2.xml", "Proteins"), "BBB"),
        sprintf("%s: Proteins 表的浅绿底纹落在 Unique 行 BBB 上（实际 %s）", o,
                paste(green_gn_values(wb_path, "xl/worksheets/sheet2.xml", "Proteins"), collapse = ",")))
  check(count_hidden_rows(wb_path, "xl/worksheets/sheet2.xml") == 1L,
        sprintf("%s: Proteins 表恰有 1 行被隐藏（诱饵无值）（实际 %d）", o,
                count_hidden_rows(wb_path, "xl/worksheets/sheet2.xml")))

}

cat(sprintf("\n%s（%d 项断言失败）\n", if (length(fails)) "失败" else "全部通过", length(fails)))
if (length(fails)) {
  for (f in fails) cat("  -", f, "\n")
  quit(status = 1L, save = "no")
}
