#!/usr/bin/env Rscript
# ============================================================================
# bioid / scripts/bioid_process.R
#
# 原始蛋白级质谱导出（Proteome Discoverer / MaxQuant 风格 .xlsx）
#   → 每个「诱饵 × EV」组合一份可复核工作簿
#     工作表：+5% / Proteins / Raw / Removed / Methods
#
# 与旧版 local-skills/BioID.R 的关键区别
#   1. 实验组/对照组不再按列位置猜，必须由 channel_map.tsv 人为指定；
#      同一诱饵对多个 EV 时，每个 EV 各出一份独立工作簿。
#   2. 显式蛋白级过滤（Master / Reverse / Potential contaminant），不再依赖
#      「抽不到 GN」顺手删掉肽段行；每类删除分开计数并记入 Removed 表。
#   3. 输出表把丰度列规范化为「诱饵在前、本组合的 EV 其次」，因此按位置取列的
#      下游工具（bioid-analysis / bioid-batch-qc）拿到的是正确的通道对。
#   4. Raw 表 = 输入工作表原样（未过滤、列序不变）；被删行全部写入 Removed 表。
#   5. readxl 读取用 guess_max = Inf；丰度列匹配 ^Abundance(?!s)，排除
#      "Abundances (Normalized)"。
#   6. GN 抽取修正三种失败形态：尾随分号、GN 后不是 PE=、一个描述里多个 GN。
#   7. 工作表已含 LOG2FC 列（已处理产物）时默认跳过，--force 可覆盖。
#   8. 修掉 openxlsx 两个会让 openpyxl（下游两个 skill 的读取器）读错或读不了的缺陷：
#      每个工作表被写上指向不存在部件的 drawing relationship，dimension 又被写成
#      "A1"（read_only 模式只看到一列）。保存前打补丁，保存后自检包结构。
#
# 用法：
#   Rscript bioid_process.R <输入目录|文件> [输出目录] [选项]
#
# 选项：
#   --map <channel_map.tsv>   诱饵→EV 通道映射；默认 <输入目录>/channel_map.tsv
#   --sheet <名称>            指定工作表；默认 Proteins，否则第一个工作表
#   --contaminants <文件>     污染物清单；默认 skill 的 assets/contaminants.tsv
#   --force                   允许处理已含 LOG2FC 的工作表
#   -h, --help                显示帮助
#
# 依赖：readxl, openxlsx
# ============================================================================

# ── 依赖 ────────────────────────────────────────────────────────────────────
require_pkg <- function(pkg, why) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    stop(sprintf("缺少 R 包 %s（%s）。请先运行：install.packages(\"%s\")",
                 pkg, why, pkg), call. = FALSE)
  }
}

# ── 样式常量 ────────────────────────────────────────────────────────────────
COLOR_HEADER <- "#A7CDF0"
COLOR_NORMAL <- "#DDEBF7"
COLOR_GREEN  <- "#C6EFCE"
COLOR_BORDER <- "#B4C6E7"

create_styles <- function() {
  list(
    header = openxlsx::createStyle(
      fontName = "Arial", fontSize = 10, textDecoration = "bold",
      fgFill = COLOR_HEADER, halign = "center", valign = "center",
      wrapText = TRUE, border = "TopBottomLeftRight",
      borderColour = COLOR_BORDER, borderStyle = "thin"
    ),
    normal = openxlsx::createStyle(
      fontName = "Arial", fontSize = 10, fgFill = COLOR_NORMAL,
      border = "TopBottomLeftRight", borderColour = COLOR_BORDER,
      borderStyle = "thin"
    ),
    unique = openxlsx::createStyle(
      fontName = "Arial", fontSize = 10, fgFill = COLOR_GREEN,
      border = "TopBottomLeftRight", borderColour = COLOR_BORDER,
      borderStyle = "thin"
    )
  )
}

# 找不到 assets/contaminants.tsv 时使用的内置清单。
DEFAULT_CONTAMINANTS <- data.frame(
  field   = c("description", "description", "gn"),
  pattern = c("Keratin", "Keratin-associated", "^KRT"),
  stringsAsFactors = FALSE
)

# ── 通用工具 ────────────────────────────────────────────────────────────────
`%||%` <- function(a, b) if (is.null(a)) b else a

script_dir <- function() {
  hit <- grep("^--file=", commandArgs(trailingOnly = FALSE), value = TRUE)
  if (!length(hit)) return(getwd())
  normalizePath(dirname(sub("^--file=", "", hit[[1]])), mustWork = FALSE)
}

find_col <- function(headers, pattern, ignore.case = TRUE) {
  idx <- which(!is.na(headers) & grepl(pattern, headers, ignore.case = ignore.case, perl = TRUE))
  if (length(idx)) idx[[1]] else NA_integer_
}

to_numeric <- function(x) suppressWarnings(as.numeric(as.character(x)))

replace_nan_with_na <- function(df) {
  for (j in seq_along(df)) {
    if (is.numeric(df[[j]])) df[[j]][is.nan(df[[j]])] <- NA_real_
  }
  df
}

glob_to_regex <- function(g) {
  out <- vapply(strsplit(g, "")[[1]], function(ch) {
    switch(ch, "*" = ".*", "?" = ".",
           "." = "\\.", "^" = "\\^", "$" = "\\$", "+" = "\\+",
           "(" = "\\(", ")" = "\\)", "[" = "\\[", "]" = "\\]",
           "{" = "\\{", "}" = "\\}", "|" = "\\|", "\\" = "\\\\", ch)
  }, character(1))
  paste0("^", paste(out, collapse = ""), "$")
}

# ── 与 openpyxl 的互操作 ────────────────────────────────────────────────────
# openxlsx 4.2.8.1 会给每个工作表写一条指向 ../drawings/drawingN.xml 的
# relationship，但从不写该部件，openpyxl 遇到悬空引用直接抛
# KeyError: "There is no item named 'xl/drawings/drawing1.xml'"，
# 于是 bioid-analysis / bioid-batch-qc 都读不了本脚本的输出。
# 这两个下游工具都用 openpyxl，所以这里在保存前把该关系去掉（只保留确实存在的
# printerSettings 关系），并在保存后自检包结构。
# 注意：本补丁会移除 drawing 关系，因此本脚本不支持插入图片的工作表
#（本脚本也从不插图），不要在同一 R 进程里用 wb 插图。
patch_openxlsx_compat <- function() {
  ns <- asNamespace("openxlsx")
  if (!exists("genBaseSheetRels", envir = ns, inherits = FALSE)) return(FALSE)
  ok <- tryCatch({
    unlockBinding("genBaseSheetRels", ns)
    on.exit(lockBinding("genBaseSheetRels", ns), add = TRUE)
    assign("genBaseSheetRels", function(sheetInd) {
      sprintf(paste0('<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/',
                     'officeDocument/2006/relationships/printerSettings" ',
                     'Target="../printerSettings/printerSettings%s.bin"/>'), sheetInd)
    }, envir = ns)
    TRUE
  }, error = function(e) FALSE)
  ok
}

# 解 zip 内的相对路径（处理 ./ 与 ../），用于校验 relationship 目标是否存在。
resolve_part_path <- function(base_dir, target) {
  parts <- strsplit(paste0(base_dir, target), "/", fixed = TRUE)[[1]]
  out <- character(0)
  for (p in parts) {
    if (!nzchar(p) || p == ".") next
    if (p == "..") {
      if (length(out)) out <- out[-length(out)]
      next
    }
    out <- c(out, p)
  }
  paste(out, collapse = "/")
}

# 返回包结构问题的清单（空 = 自洽，下游 openpyxl 可正常读取）。
# 两类问题都会让下游读错或读不了：
#   1. relationship 指向不存在的部件（openxlsx 的 drawing 关系缺陷）；
#   2. 工作表 dimension 被写成 "A1" —— read_only 模式只看到 1 列，
#      bioid-analysis 因此报「至少需要两个 Abundance: 列」。
check_xlsx_integrity <- function(path) {
  members <- tryCatch(utils::unzip(path, list = TRUE)$Name, error = function(e) character(0))
  if (!length(members)) return(sprintf("%s 不是可读的 xlsx 包", basename(path)))
  bad <- character(0)

  for (rel in grep("_rels/.*\\.rels$", members, value = TRUE)) {
    txt <- paste(readLines(unz(path, rel), warn = FALSE), collapse = "")
    targets <- regmatches(txt, gregexpr('Target="[^"]+"', txt))[[1]]
    if (!length(targets)) next
    targets <- sub('"$', "", sub('^Target="', "", targets))
    targets <- targets[!grepl("^https?://", targets)]
    base_dir <- sub("_rels/[^/]*$", "", rel)
    for (t in targets) {
      resolved <- resolve_part_path(base_dir, t)
      if (!(resolved %in% members)) bad <- c(bad, sprintf("悬空引用 %s -> %s", rel, resolved))
    }
  }

  for (part in grep("^xl/worksheets/sheet[0-9]*\\.xml$", members, value = TRUE)) {
    head_txt <- tryCatch(read_zip_head(path, part), error = function(e) "")
    if (!nzchar(head_txt)) next
    m <- regmatches(head_txt, regexec('<dimension ref="([A-Z]+)[0-9]+(?::([A-Z]+)[0-9]+)?"', head_txt))[[1]]
    end_col <- if (length(m) >= 3L && nzchar(m[[3]])) m[[3]] else if (length(m) >= 2L) m[[2]] else NA_character_
    row1 <- regmatches(head_txt, regexec("<row r=\"1\"[^>]*>(.*?)</row>", head_txt))[[1]]
    n_cells <- if (length(row1) >= 2L) {
      length(regmatches(row1[[2]], gregexpr("<c ", row1[[2]], fixed = TRUE))[[1]])
    } else 0L
    need <- col_letters_to_int(end_col)
    if (is.na(need) || need < n_cells) {
      bad <- c(bad, sprintf("dimension 过小 %s: ref=%s，第 1 行却有 %d 个单元格",
                            part, if (is.na(end_col)) "(缺失)" else end_col, n_cells))
    }
  }
  bad
}

col_letters_to_int <- function(x) {
  if (length(x) != 1L || is.na(x) || !nzchar(x)) return(NA_integer_)
  v <- utf8ToInt(toupper(x)) - utf8ToInt("A") + 1L
  if (any(v < 1L | v > 26L)) return(NA_integer_)
  sum(v * 26L ^ (rev(seq_along(v)) - 1L))
}

# 读取 zip 内某个部件开头的一段文本（connection 在本函数帧内关闭，避免泄漏）。
read_zip_head <- function(path, part, n = 200000L) {
  con <- unz(path, part, open = "rb")
  on.exit(close(con))
  rawToChar(suppressWarnings(readBin(con, "raw", n)))
}

# openxlsx 把每个工作表的 dimension 写成 "A1"，read_only 读取器只能看到一列；
# 保存前按真实使用范围写回正确的 dimension 元素。
set_sheet_dimension <- function(wb, sheet_name, n_rows, n_cols) {
  i <- match(sheet_name, names(wb))
  if (is.na(i)) return(invisible(NULL))
  wb$worksheets[[i]]$dimension <- sprintf(
    '<dimension ref="A1:%s%d"/>', openxlsx::int2col(max(1L, n_cols)), max(1L, n_rows))
  invisible(NULL)
}



# 描述里取基因名。GN= 之后取到第一个空白或分号为止，再去掉尾部分号/逗号：
#   "... GN=ACACA PE=1 SV=2"   -> ACACA
#   "... GN=ABC SV=2"          -> ABC   （旧版此处返回空并删除整行）
#   "...; GN=ABC; PE=1; SV=2"  -> ABC   （旧版返回 "ABC;"）
#   "... GN=A GN=B PE=1"       -> A     （旧版返回最后一个 B）
extract_gn <- function(description) {
  x <- as.character(description)
  x[is.na(x)] <- ""
  matched <- regmatches(x, regexec("GN=([^;\\s]+)", x, perl = TRUE))
  vapply(matched, function(z) {
    if (length(z) >= 2L) sub("[;,]+$", "", z[[2]]) else ""
  }, FUN.VALUE = character(1))
}

# ── 参数 ────────────────────────────────────────────────────────────────────
usage <- function() {
  cat("用法: Rscript bioid_process.R <输入目录|文件> [输出目录] [选项]\n\n")
  cat("  --map <channel_map.tsv>   诱饵→EV 通道映射（默认 <输入目录>/channel_map.tsv）\n")
  cat("  --sheet <名称>            指定工作表（默认 Proteins，否则第一个工作表）\n")
  cat("  --contaminants <文件>     污染物清单（默认 skill 的 assets/contaminants.tsv）\n")
  cat("  --force                   允许处理已含 LOG2FC 的工作表\n")
  cat("  -h, --help                显示本帮助\n\n")
  cat("示例:\n")
  cat("  Rscript bioid_process.R \"~/BioID-4th/MS261493-Protein+ID\" ./processed\n")
}

parse_args <- function(argv) {
  opt <- list(input = NULL, out = NULL, map = NULL, sheet = NULL,
              contaminants = NULL, force = FALSE, help = FALSE)
  pos <- character(0)
  i <- 1L
  while (i <= length(argv)) {
    a <- argv[[i]]
    if (a %in% c("-h", "--help")) {
      opt$help <- TRUE; i <- i + 1L; next
    }
    if (a == "--force") {
      opt$force <- TRUE; i <- i + 1L; next
    }
    if (a %in% c("--map", "--sheet", "--contaminants", "--out")) {
      if (i + 1L > length(argv)) stop(sprintf("%s 需要一个取值", a), call. = FALSE)
      opt[[sub("^--", "", a)]] <- argv[[i + 1L]]
      i <- i + 2L; next
    }
    if (startsWith(a, "--")) stop(sprintf("未知选项: %s（用 --help 查看用法）", a), call. = FALSE)
    pos <- c(pos, a); i <- i + 1L
  }
  if (length(pos) >= 1L) opt$input <- pos[[1]]
  if (length(pos) >= 2L) opt$out <- pos[[2]]
  if (is.null(opt$out)) opt$out <- "./processed"
  opt
}

# ── 通道映射 ────────────────────────────────────────────────────────────────
# 每行一条规则：<诱饵通道>  <EV 通道>[|<EV 通道>...]  [文件名 glob]
# 例如：F4	F1|F2      （诱饵 F4 分别与 EV F1、F2 比较，各出一份工作簿）
# 带文件名 glob 的规则优先于通用规则。
read_channel_map <- function(path) {
  if (is.null(path) || !file.exists(path)) {
    stop(sprintf(paste0("找不到通道映射文件: %s\n",
                        "  复制 skill 的 assets/channel_map.tsv 到输入目录并填写诱饵→EV，\n",
                        "  或用 --map 指定路径。映射必须人为给定，脚本不再按列位置猜测。"),
                 if (is.null(path)) "<未指定>" else path), call. = FALSE)
  }
  lines <- readLines(path, warn = FALSE)
  rules <- list()
  for (i in seq_along(lines)) {
    ln <- trimws(lines[[i]])
    if (!nzchar(ln) || startsWith(ln, "#")) next
    f <- strsplit(ln, "[[:space:]]+")[[1]]
    f <- f[nzchar(f)]
    if (length(f) < 2L) {
      stop(sprintf("%s 第 %d 行无法解析（需要：诱饵通道 EV通道[|EV通道...] [文件名glob]）: %s",
                   basename(path), i, ln), call. = FALSE)
    }
    ev <- trimws(strsplit(f[[2]], "[|,]")[[1]])
    ev <- ev[nzchar(ev)]
    if (!length(ev)) stop(sprintf("%s 第 %d 行没有 EV 通道: %s", basename(path), i, ln), call. = FALSE)
    rules[[length(rules) + 1L]] <- list(
      bait = f[[1]], ev = ev,
      file = if (length(f) >= 3L) f[[3]] else NA_character_,
      raw = ln
    )
  }
  if (!length(rules)) stop(sprintf("%s 里没有任何规则", basename(path)), call. = FALSE)

  baits <- vapply(rules, function(r) r$bait, character(1))
  evs <- unique(unlist(lapply(rules, function(r) r$ev)))
  clash <- intersect(unique(baits), evs)
  if (length(clash)) {
    warning(sprintf("通道 %s 同时被当作诱饵和 EV，请确认 channel_map 是否写错",
                    paste(clash, collapse = ", ")), call. = FALSE)
  }
  dup <- unique(baits[duplicated(baits)])
  if (length(dup)) {
    warning(sprintf("通道 %s 出现在多条诱饵规则里，只有第一条会生效",
                    paste(dup, collapse = ", ")), call. = FALSE)
  }
  rules
}

read_contaminants <- function(path) {
  if (is.null(path) || !file.exists(path)) return(DEFAULT_CONTAMINANTS)
  lines <- readLines(path, warn = FALSE)
  lines <- trimws(lines)
  lines <- lines[nzchar(lines) & !startsWith(lines, "#")]
  parsed <- lapply(lines, function(ln) {
    f <- strsplit(ln, "[[:space:]]+")[[1]]
    if (length(f) < 2L) return(NULL)
    data.frame(field = tolower(f[[1]]), pattern = paste(f[-1], collapse = " "),
               stringsAsFactors = FALSE)
  })
  parsed <- do.call(rbind, parsed)
  if (is.null(parsed) || !nrow(parsed)) return(DEFAULT_CONTAMINANTS)
  parsed
}

# ── 丰度列与比较组合 ────────────────────────────────────────────────────────
# ^Abundance(?!s) 只匹配 "Abundance:" 开头的列，排除 "Abundances (Normalized):"。
abundance_cols <- function(headers) {
  idx <- which(!is.na(headers) & grepl("^Abundance(?!s)", headers, ignore.case = TRUE, perl = TRUE))
  if (!length(idx)) {
    return(data.frame(idx = integer(0), header = character(0), token = character(0),
                      stringsAsFactors = FALSE))
  }
  data.frame(idx = idx, header = as.character(headers[idx]),
             token = vapply(as.character(headers[idx]), channel_token, character(1)),
             stringsAsFactors = FALSE)
}

channel_token <- function(header) {
  h <- as.character(header)
  if (is.na(h) || !nzchar(h)) return(NA_character_)
  m <- regmatches(h, regexec("^Abundance[^:]*:\\s*([^:\\s]+)", h,
                             ignore.case = TRUE, perl = TRUE))[[1]]
  if (length(m) >= 2L) return(m[[2]])
  m <- regmatches(h, regexec("^Abundance\\s*(.+)$", h, ignore.case = TRUE, perl = TRUE))[[1]]
  if (length(m) >= 2L) return(trimws(m[[2]]))
  NA_character_
}

# 返回该文件需要产出的所有 (诱饵, EV) 组合。
pick_comparisons <- function(abund, rules, file_name) {
  has_glob <- vapply(rules, function(r) !is.na(r$file), logical(1))
  ordered <- rules[order(has_glob)]        # 带文件名 glob 的规则优先
  out <- list()
  seen <- character(0)
  for (r in ordered) {
    if (!is.na(r$file) && !grepl(glob_to_regex(r$file), file_name, ignore.case = TRUE)) next
    bi <- match(r$bait, abund$token)
    if (is.na(bi)) next
    ei <- match(r$ev, abund$token)
    if (any(is.na(ei))) next
    for (k in seq_along(r$ev)) {
      key <- paste(r$bait, r$ev[[k]], sep = "->")
      if (key %in% seen) next
      seen <- c(seen, key)
      out[[length(out) + 1L]] <- list(
        bait_idx = abund$idx[[bi]], bait_name = abund$header[[bi]], bait_token = r$bait,
        ev_idx = abund$idx[[ei[[k]]]], ev_name = abund$header[[ei[[k]]]], ev_token = r$ev[[k]],
        rule = r$raw
      )
    }
  }
  out
}

# ── 读取 ────────────────────────────────────────────────────────────────────
read_sheet_quiet <- function(path, sheet) {
  df <- withCallingHandlers(
    readxl::read_excel(path, sheet = sheet, col_names = TRUE,
                       na = c("", "NA", "N/A"), guess_max = Inf),
    message = function(cond) {
      # 列名修复提示（...17）与 guess_max 说明都是噪声，其余消息照常输出。
      if (grepl("New names|guess_max", conditionMessage(cond))) invokeRestart("muffleMessage")
    },
    warning = function(cond) {
      if (grepl("guess_max", conditionMessage(cond))) invokeRestart("muffleWarning")
    }
  )
  as.data.frame(df, stringsAsFactors = FALSE, check.names = FALSE)
}

choose_sheet <- function(sheets, requested) {
  if (!is.null(requested)) return(if (requested %in% sheets) requested else NA_character_)
  hit <- grep("^Proteins$", sheets, ignore.case = TRUE, value = TRUE)
  if (length(hit)) return(hit[[1]])
  sheets[[1]]
}

# ── 行过滤（显式、可追溯） ──────────────────────────────────────────────────
# 返回 list(df, gn, removed, stats, desc_idx, acc_idx, n_input, n_unnamed)
prepare_table <- function(df, headers, contaminants) {
  n <- nrow(df)
  desc_idx <- find_col(headers, "^Description$")
  if (is.na(desc_idx)) desc_idx <- find_col(headers, "Description")
  acc_idx <- find_col(headers, "^Accession$")
  if (is.na(acc_idx)) acc_idx <- find_col(headers, "Accession")

  desc <- if (is.na(desc_idx)) rep("", n) else as.character(df[[desc_idx]])
  desc[is.na(desc)] <- ""
  acc <- if (is.na(acc_idx)) rep(NA_character_, n) else as.character(df[[acc_idx]])
  gn <- if (is.na(desc_idx)) rep("", n) else extract_gn(df[[desc_idx]])

  keep <- rep(TRUE, n)
  removed <- list()
  stats <- list()
  add_removed <- function(idx, reason) {
    if (!length(idx)) return(invisible(NULL))
    removed[[length(removed) + 1L]] <<- data.frame(
      RawSheetRow = idx + 1L, Accession = acc[idx], Description = substr(desc[idx], 1, 300),
      GN = gn[idx], Reason = reason, stringsAsFactors = FALSE
    )
    invisible(NULL)
  }
  drop_rows <- function(flag, reason, label) {
    hit <- which(keep & flag)
    if (!length(hit)) return(invisible(NULL))
    stats[[label]] <<- length(hit)
    add_removed(hit, reason)
    keep[hit] <<- FALSE
    invisible(NULL)
  }

  # 1) 蛋白级行：肽段级行（PD 导出常把肽段表叠在下面）在这里被明确剔除。
  master_idx <- find_col(headers, "^Master$")
  if (!is.na(master_idx)) {
    lab <- tolower(trimws(as.character(df[[master_idx]])))
    is_prot <- !is.na(lab) & lab == "master protein"
    if (any(is_prot)) {
      drop_rows(!is_prot, "非蛋白级行（Master 不是 'Master Protein'）", "peptide_level")
    } else {
      warning("有 Master 列但没有任何 'Master Protein' 行，跳过蛋白级过滤", call. = FALSE)
    }
  }
  for (spec in list(list("^Reverse$", "Reverse 命中"),
                    list("^Potential contaminant$", "潜在污染物标记"))) {
    j <- find_col(headers, spec[[1]])
    if (is.na(j)) next
    v <- toupper(trimws(as.character(df[[j]])))
    drop_rows(v %in% c("+", "TRUE", "YES", "1", "Y"), spec[[2]], spec[[2]])
  }

  # 2) Description 为空
  if (!is.na(desc_idx)) {
    drop_rows(!nzchar(trimws(desc)), "Description 为空", "desc_empty")
  }

  # 3) 污染物
  if (nrow(contaminants)) {
    hit <- rep(FALSE, n)
    why <- rep(NA_character_, n)
    for (k in seq_len(nrow(contaminants))) {
      target <- if (contaminants$field[[k]] == "gn") gn else desc
      m <- grepl(contaminants$pattern[[k]], target, ignore.case = TRUE)
      fresh <- m & !hit
      why[fresh] <- sprintf("污染物 %s 匹配 /%s/", contaminants$field[[k]], contaminants$pattern[[k]])
      hit <- hit | m
    }
    idx <- which(keep & hit)
    if (length(idx)) {
      stats$contaminant <- length(idx)
      removed[[length(removed) + 1L]] <- data.frame(
        RawSheetRow = idx + 1L, Accession = acc[idx], Description = substr(desc[idx], 1, 300),
        GN = gn[idx], Reason = why[idx], stringsAsFactors = FALSE
      )
      keep[idx] <- FALSE
    }
  }

  # 4) 抽不到 GN
  drop_rows(!nzchar(trimws(gn)), "抽不到 GN", "no_gn")

  kept <- which(keep)
  out_df <- df[kept, , drop = FALSE]

  # 5) 测量列归一为数值。原始导出常把注释文字混进 Coverage / # Peptides 这类列
  #    （叠在下面的肽段表把文本写进了同样的列位置），readxl 于是把整列读成文本，
  #    写出的工作簿在 Excel 里就无法排序。判据只看保留下来的行：其中 ≥90% 能解析成
  #    数值的列才转换，Description / Modifications / Accession 之类不受影响。
  for (j in seq_along(out_df)) {
    v <- out_df[[j]]
    if (!is.character(v)) next
    nz <- v[!is.na(v) & nzchar(trimws(v))]
    if (!length(nz)) next
    if (mean(!is.na(suppressWarnings(as.numeric(nz)))) >= 0.9) {
      out_df[[j]] <- to_numeric(v)
    }
  }

  list(
    df = out_df,
    gn = gn[kept],
    removed = if (length(removed)) do.call(rbind, removed) else
      data.frame(RawSheetRow = integer(0), Accession = character(0), Description = character(0),
                 GN = character(0), Reason = character(0), stringsAsFactors = FALSE),
    stats = stats,
    desc_idx = desc_idx,
    acc_idx = acc_idx,
    n_input = n,
    n_unnamed = sum(grepl("^\\.\\.\\.[0-9]+$", headers))
  )
}

# ── 丰度计算 ────────────────────────────────────────────────────────────────
compute_metrics <- function(df, bait_idx, ev_idx) {
  bait <- to_numeric(df[[bait_idx]])
  ev <- to_numeric(df[[ev_idx]])
  has_bait <- !is.na(bait) & bait > 0
  has_ev <- !is.na(ev) & ev > 0
  is_unique <- has_bait & !has_ev

  log2fc <- rep(NA_real_, length(bait))
  both <- has_bait & has_ev
  log2fc[both] <- log2(bait[both] / ev[both])

  valid_ev <- ev[has_ev]
  p5 <- if (length(valid_ev)) {
    as.numeric(stats::quantile(valid_ev, probs = 0.05, na.rm = TRUE,
                               names = FALSE, type = 7))
  } else {
    0
  }
  ev_plus5 <- ev
  ev_plus5[is_unique] <- p5
  log2fc_plus5 <- rep(NA_real_, length(bait))
  v <- has_bait & !is.na(ev_plus5) & ev_plus5 > 0
  log2fc_plus5[v] <- log2(bait[v] / ev_plus5[v])

  if (sum(is_unique) > 0L && (!is.finite(p5) || p5 <= 0)) {
    warning(sprintf("对照组没有可用的非零值（第5百分位=%.3g），%d 个 Unique 行在 +5%% 表里 LOG2FC 为空",
                    p5, sum(is_unique)), call. = FALSE)
  }

  list(bait = bait, ev = ev, has_bait = has_bait, has_ev = has_ev,
       is_unique = is_unique, log2fc = log2fc, ev_plus5 = ev_plus5,
       log2fc_plus5 = log2fc_plus5, p5 = p5)
}

# ── 输出布局 ────────────────────────────────────────────────────────────────
# 丰度列规范化为「诱饵、本组合 EV、其余原序」，其余列保持原相对顺序；
# GN 插到 Description 之后，LOG2FC 追加在最后一列。
make_layout <- function(headers, abund_idx, desc_idx, acc_idx, bait_idx, ev_idx) {
  n <- length(headers)
  order <- seq_len(n)
  slots <- match(abund_idx, order)
  order[slots] <- c(bait_idx, ev_idx, setdiff(abund_idx, c(bait_idx, ev_idx)))
  desc_pos <- which(order == desc_idx)
  final_pos <- function(oi) {
    if (is.na(oi)) return(NA_integer_)
    p <- which(order == oi)
    if (!length(p)) return(NA_integer_)
    p + if (p > desc_pos) 1L else 0L
  }
  list(
    order = order,
    n_cols = n,
    n_out = n + 2L,
    desc_pre = desc_pos,
    desc_pos = desc_pos,
    gn_pos = desc_pos + 1L,
    acc_pos = final_pos(acc_idx),
    bait_pos = final_pos(bait_idx),
    ev_pos = final_pos(ev_idx),
    log2fc_pos = n + 2L,
    ev_pre = which(order == ev_idx),
    bait_pre = which(order == bait_idx)
  )
}

build_sheet_df <- function(df_rows, layout, headers, gn, ev_values, log2fc) {
  out <- df_rows[, layout$order, drop = FALSE]
  # 输入导出常把空丰度写成空串或 "Filtered" 之类，readxl 于是把整列读成文本；
  # 这里把本组合用到的诱饵/EV 两列转成数值，便于 Excel 排序，也让下游拿到数字。
  # 原始文本形态保留在 Raw 表里。
  out[[layout$bait_pre]] <- to_numeric(out[[layout$bait_pre]])
  if (!is.null(ev_values)) out[[layout$ev_pre]] <- ev_values else
    out[[layout$ev_pre]] <- to_numeric(out[[layout$ev_pre]])

  left <- out[, seq_len(layout$desc_pre), drop = FALSE]
  right <- if (layout$desc_pre < ncol(out)) {
    out[, seq.int(layout$desc_pre + 1L, ncol(out)), drop = FALSE]
  } else {
    NULL
  }
  gn_df <- data.frame(GN = gn, stringsAsFactors = FALSE, check.names = FALSE)

  final <- if (is.null(right)) {
    cbind(left, gn_df, LOG2FC = log2fc)
  } else {
    cbind(left, gn_df, right, LOG2FC = log2fc)
  }

  ordered_headers <- headers[layout$order]
  names(final) <- c(
    ordered_headers[seq_len(layout$desc_pre)], "GN",
    if (layout$desc_pre < length(ordered_headers)) ordered_headers[seq.int(layout$desc_pre + 1L, length(ordered_headers))],
    "LOG2FC"
  )
  replace_nan_with_na(final)
}

# ── 写表 ────────────────────────────────────────────────────────────────────
write_header_row <- function(wb, sheet, headers, styles) {
  openxlsx::writeData(wb, sheet = sheet, x = matrix(headers, nrow = 1L),
                      startRow = 1L, startCol = 1L, colNames = FALSE, rowNames = FALSE,
                      keepNA = FALSE)
  openxlsx::addStyle(wb, sheet = sheet, style = styles$header, rows = 1L,
                     cols = seq_along(headers), gridExpand = TRUE, stack = TRUE)
}

write_processed_sheet <- function(wb, sheet_name, df_visible, df_hidden, layout,
                                  headers, gn_visible, gn_hidden, styles,
                                  ev_values = NULL, log2fc, is_unique = NULL) {
  openxlsx::addWorksheet(wb, sheet_name)
  n_cols <- layout$n_out
  n_hidden <- nrow(df_hidden)

  out_headers <- names(build_sheet_df(df_hidden[0, , drop = FALSE], layout, headers,
                                      gn_hidden[0], NULL, numeric(0)))
  write_header_row(wb, sheet_name, out_headers, styles)

  # 隐藏行先写、可见行（已排序）随后，与原脚本一致。
  all_rows <- rbind(df_hidden, df_visible)
  all_gn <- c(gn_hidden, gn_visible)
  # 隐藏行（诱饵无值）的 EV 与 LOG2FC 本来就是 NA，按位置补齐即可。
  ev_all <- if (is.null(ev_values)) NULL else c(rep(NA_real_, n_hidden), ev_values)
  fc_all <- c(rep(NA_real_, n_hidden), log2fc)
  out <- build_sheet_df(all_rows, layout, headers, all_gn, ev_all, fc_all)

  if (nrow(out) > 0L) {
    openxlsx::writeData(wb, sheet = sheet_name, x = out, startRow = 2L, startCol = 1L,
                        colNames = FALSE, rowNames = FALSE, keepNA = FALSE)
    last_row <- nrow(out) + 1L
    openxlsx::addStyle(wb, sheet = sheet_name, style = styles$normal,
                       rows = 2L:last_row, cols = seq_len(n_cols),
                       gridExpand = TRUE, stack = TRUE)

    if (!is.null(is_unique) && any(is_unique) && nrow(df_visible) > 0L) {
      pos <- which(is_unique)
      openxlsx::addStyle(wb, sheet = sheet_name, style = styles$unique,
                         rows = 1L + n_hidden + pos, cols = seq_len(n_cols),
                         gridExpand = TRUE, stack = TRUE)
    }
    if (n_hidden > 0L) {
      openxlsx::groupRows(wb, sheet = sheet_name, rows = 2L:(n_hidden + 1L), hidden = TRUE)
    }
    openxlsx::conditionalFormatting(wb, sheet = sheet_name, cols = layout$log2fc_pos,
                                    rows = 2L:last_row, type = "colourScale",
                                    style = c("#63BE7B", "#FFEB84", "#F8696B"))
  }

  visible_cols <- unique(stats::na.omit(c(layout$acc_pos, layout$desc_pos, layout$gn_pos,
                                          layout$bait_pos, layout$ev_pos, layout$log2fc_pos)))
  hidden_cols <- setdiff(seq_len(n_cols), visible_cols)
  if (length(hidden_cols)) {
    openxlsx::setColWidths(wb, sheet = sheet_name, cols = hidden_cols, widths = 2, hidden = TRUE)
  }
  if (!is.na(layout$acc_pos)) {
    openxlsx::setColWidths(wb, sheet = sheet_name, cols = layout$acc_pos, widths = 12)
  }
  openxlsx::setColWidths(wb, sheet = sheet_name, cols = layout$desc_pos, widths = 60)
  openxlsx::setColWidths(wb, sheet = sheet_name, cols = layout$gn_pos, widths = 12)
  openxlsx::setColWidths(wb, sheet = sheet_name, cols = layout$bait_pos, widths = 20)
  openxlsx::setColWidths(wb, sheet = sheet_name, cols = layout$ev_pos, widths = 20)
  openxlsx::setColWidths(wb, sheet = sheet_name, cols = layout$log2fc_pos, widths = 12)
  openxlsx::freezePane(wb, sheet = sheet_name, firstRow = TRUE)
  set_sheet_dimension(wb, sheet_name, nrow(out) + 1L, n_cols)
  invisible(nrow(out))
}

write_plain_sheet <- function(wb, sheet_name, df, styles, widths = 18) {
  openxlsx::addWorksheet(wb, sheet_name)
  openxlsx::writeData(wb, sheet = sheet_name, x = df, startRow = 1L, startCol = 1L)
  n_cols <- max(1L, ncol(df))
  openxlsx::addStyle(wb, sheet = sheet_name, style = styles$header, rows = 1L,
                     cols = seq_len(n_cols), gridExpand = TRUE, stack = TRUE)
  if (nrow(df) > 0L) {
    openxlsx::addStyle(wb, sheet = sheet_name, style = styles$normal, rows = 2L:(nrow(df) + 1L),
                       cols = seq_len(n_cols), gridExpand = TRUE, stack = TRUE)
  }
  openxlsx::setColWidths(wb, sheet = sheet_name, cols = seq_len(n_cols), widths = widths)
  openxlsx::freezePane(wb, sheet = sheet_name, firstRow = TRUE)
  set_sheet_dimension(wb, sheet_name, nrow(df) + 1L, n_cols)
}

# ── 单文件处理 ──────────────────────────────────────────────────────────────
process_file <- function(input_path, out_dir, opts, rules, contaminants, styles) {
  file_name <- basename(input_path)
  cat(sprintf("  %s\n", file_name))

  sheets <- tryCatch(readxl::excel_sheets(input_path), error = function(e) character(0))
  if (!length(sheets)) {
    cat("    ✗ 读不到工作表，跳过\n")
    return(list(skipped = file_name, reason = "读不到工作表"))
  }
  sheet <- choose_sheet(sheets, opts$sheet)
  if (is.na(sheet)) {
    cat(sprintf("    ✗ 工作表 '%s' 不存在（可用：%s），跳过\n",
                opts$sheet, paste(sheets, collapse = ", ")))
    return(list(skipped = file_name, reason = sprintf("工作表 '%s' 不存在", opts$sheet)))
  }
  if (length(sheets) > 1L && !identical(sheet, sheets[[1]])) {
    cat(sprintf("    使用工作表: %s（另有 %s）\n", sheet,
                paste(setdiff(sheets, sheet), collapse = ", ")))
  }

  df <- tryCatch(read_sheet_quiet(input_path, sheet), error = function(e) e)
  if (inherits(df, "error")) {
    cat(sprintf("    ✗ 读取失败: %s\n", conditionMessage(df)))
    return(list(skipped = file_name, reason = conditionMessage(df)))
  }
  headers <- names(df)

  # 已处理产物防护：含 LOG2FC 说明这份表已经被处理过。
  if (!opts$force && any(grepl("^LOG2FC$", headers, ignore.case = TRUE))) {
    cat("    ✗ 该工作表已含 LOG2FC 列（已处理产物）。请把原始导出作为输入；如确需重跑加 --force\n")
    return(list(skipped = file_name, reason = "已含 LOG2FC 列（已处理产物）"))
  }

  if (is.na(find_col(headers, "Description"))) {
    cat("    ✗ 找不到 Description 列，跳过\n")
    return(list(skipped = file_name, reason = "找不到 Description 列"))
  }

  abund <- abundance_cols(headers)
  if (nrow(abund) < 1L) {
    cat("    ✗ 找不到以 'Abundance:' 开头的丰度列，跳过\n")
    return(list(skipped = file_name, reason = "无 Abundance: 列"))
  }
  comps <- pick_comparisons(abund, rules, file_name)
  if (!length(comps)) {
    cat(sprintf("    ✗ channel_map 没有为该文件匹配到诱饵→EV（文件里的通道：%s）\n",
                paste(sprintf("%s=%s", abund$token, abund$header), collapse = ", ")))
    return(list(skipped = file_name,
                reason = sprintf("channel_map 未覆盖（通道：%s）", paste(abund$token, collapse = ", "))))
  }

  prepared <- prepare_table(df, headers, contaminants)
  if (!nrow(prepared$df)) {
    cat("    ✗ 过滤后没有剩余行，跳过\n")
    return(list(skipped = file_name, reason = "过滤后无剩余行"))
  }
  cat(sprintf("    输入 %d 行 → 保留 %d 行（%.1f%% 是肽段级/非蛋白行）\n",
              prepared$n_input, nrow(prepared$df),
              100 * (prepared$stats$peptide_level %||% 0) / max(1L, prepared$n_input)))
  if (length(prepared$stats)) {
    cat(sprintf("    删除明细: %s\n", paste(
      sprintf("%s=%d", names(prepared$stats), unlist(prepared$stats)), collapse = ", ")))
  }
  if (prepared$n_unnamed > 0L) {
    cat(sprintf("    提示: 输入表有 %d 个无表头列（来自叠在下面的肽段表宽度）\n", prepared$n_unnamed))
  }

  records <- list()
  for (cmp in comps) {
    for (idx in c(cmp$bait_idx, cmp$ev_idx)) {
      if (is.logical(df[[idx]])) {
        cat(sprintf("    ✗ 通道列 '%s' 被读成逻辑型（该列前 1000 行可能全空），跳过该文件\n",
                    headers[[idx]]))
        return(list(skipped = file_name, reason = sprintf("通道列 '%s' 读成逻辑型", headers[[idx]])))
      }
    }
    layout <- make_layout(headers, abund$idx, prepared$desc_idx, prepared$acc_idx,
                          cmp$bait_idx, cmp$ev_idx)
    m <- compute_metrics(prepared$df, cmp$bait_idx, cmp$ev_idx)

    visible <- which(m$has_bait)
    hidden <- which(!m$has_bait)
    uniq <- which(m$is_unique)

    # 两通道的载量比：>1 表示对照上样更多。下游会据此做载量校正，这里先给出可读的数。
    bait_total <- sum(m$bait, na.rm = TRUE)
    ev_total <- sum(m$ev, na.rm = TRUE)
    load_ratio <- if (bait_total > 0) ev_total / bait_total else NA_real_

    df_visible <- prepared$df[visible, , drop = FALSE]
    df_hidden <- prepared$df[hidden, , drop = FALSE]
    gn_visible <- prepared$gn[visible]
    gn_hidden <- prepared$gn[hidden]
    bait_v <- m$bait[visible]

    # Proteins：Unique 在前按诱饵丰度降序，其余按 LOG2FC 降序。
    ord_u <- if (length(uniq)) order(bait_v[uniq], decreasing = TRUE, na.last = TRUE) else integer(0)
    u_in_vis <- match(uniq, visible)
    rest <- setdiff(seq_along(visible), u_in_vis)
    ord_r <- if (length(rest)) order(m$log2fc[visible][rest], decreasing = TRUE, na.last = TRUE) else integer(0)
    p_order <- c(u_in_vis[ord_u], rest[ord_r])
    is_unique_sorted <- seq_along(visible) %in% seq_along(u_in_vis[ord_u])

    # +5%：按补值后的 LOG2FC 降序。
    ord5 <- order(m$log2fc_plus5[visible], decreasing = TRUE, na.last = TRUE)

    n_cmp <- length(comps)
    stem <- tools::file_path_sans_ext(file_name)
    out_name <- if (n_cmp == 1L) {
      sprintf("%s_processed.xlsx", stem)
    } else {
      sprintf("%s_vs%s_processed.xlsx", stem, cmp$ev_token)
    }
    out_path <- file.path(out_dir, out_name)

    wb <- openxlsx::createWorkbook()

    write_processed_sheet(
      wb, "+5%",
      df_visible = df_visible[ord5, , drop = FALSE],
      df_hidden = df_hidden, layout = layout, headers = headers,
      gn_visible = gn_visible[ord5], gn_hidden = gn_hidden, styles = styles,
      ev_values = m$ev_plus5[visible][ord5],
      log2fc = m$log2fc_plus5[visible][ord5],
      # 注意：uniq 是 prepared$df 里的全局行号，可见行要用 u_in_vis（可见段内的位置），
      # 否则有隐藏行时底纹会标到错误的行上。
      is_unique = (seq_along(visible) %in% u_in_vis)[ord5]
    )
    write_processed_sheet(
      wb, "Proteins",
      df_visible = df_visible[p_order, , drop = FALSE],
      df_hidden = df_hidden, layout = layout, headers = headers,
      gn_visible = gn_visible[p_order], gn_hidden = gn_hidden, styles = styles,
      log2fc = m$log2fc[visible][p_order],
      is_unique = is_unique_sorted
    )
    write_plain_sheet(wb, "Raw", df, styles)
    write_plain_sheet(wb, "Removed",
                      if (nrow(prepared$removed)) prepared$removed else
                        data.frame(RawSheetRow = NA_integer_, Accession = NA_character_,
                                   Description = NA_character_, GN = NA_character_,
                                   Reason = "无删除行（输入行全部保留）", stringsAsFactors = FALSE),
                      styles, widths = c(9, 14, 60, 12, 40))
    write_plain_sheet(wb, "Methods", rbind(
      methods_table(file_name, sheet, cmp, prepared, nrow(df_visible), nrow(df_hidden),
                    length(uniq), m$p5, out_name, opts),
      data.frame(
        Item = c("诱饵通道总信号", "对照通道总信号", "对照/诱饵 总信号比"),
        Value = c(sprintf("%.6g", bait_total), sprintf("%.6g", ev_total),
                  sprintf("%.3f", load_ratio)),
        Note = c("本工作簿内诱饵通道所有非零值之和", "同上，对照通道",
                 ">1 表示对照上样更多；下游 bioid-analysis 会做载量校正"),
        stringsAsFactors = FALSE)), styles, widths = c(26, 46, 52))

    openxlsx::saveWorkbook(wb, file = out_path, overwrite = TRUE)

    pkg_issues <- check_xlsx_integrity(out_path)
    if (length(pkg_issues)) {
      cat("    ⚠ 输出包结构有问题（下游 openpyxl 会读错或读不了）:\n")
      for (d in pkg_issues) cat(sprintf("        %s\n", d))
    }

    cat(sprintf("    ✓ %s  [诱饵 %s ← 对照 %s]\n", out_name, cmp$bait_token, cmp$ev_token))
    cat(sprintf("      Unique %d | 正常 %d | 隐藏 %d | 对照第5百分位 %.4g | 对照/诱饵信号比 %.2f%s\n",
                length(uniq), sum(m$has_bait & m$has_ev), sum(!m$has_bait), m$p5, load_ratio,
                if (load_ratio > 2 || load_ratio < 0.5) "  ⚠ 两通道上样量差异较大" else ""))

    records[[length(records) + 1L]] <- list(
      file = file_name, sheet = sheet, bait = cmp$bait_name, bait_token = cmp$bait_token,
      ev = cmp$ev_name, ev_token = cmp$ev_token, out = out_name,
      n_input = prepared$n_input, n_kept = nrow(prepared$df),
      unique = length(uniq), normal = sum(m$has_bait & m$has_ev), hidden = sum(!m$has_bait),
      p5 = m$p5, stats = prepared$stats, removed = prepared$removed,
      pkg_issues = pkg_issues
    )
  }

  list(records = records, skipped = NULL)
}

methods_table <- function(file_name, sheet, cmp, prepared, n_visible, n_hidden,
                          n_unique, p5, out_name, opts) {
  rows <- list(
    c("输入文件", file_name, ""),
    c("输入工作表", sheet, ""),
    c("诱饵通道", cmp$bait_name, sprintf("channel_map 规则: %s", cmp$rule)),
    c("对照(EV)通道", cmp$ev_name, "该组合的 EV；丰度列已按「诱饵、EV、其余」重排"),
    c("输出文件", out_name, ""),
    c("输入行数", prepared$n_input, "行号见 Removed 表的 RawSheetRow（等于输入工作表行号）"),
    c("保留行数", prepared$n_input - sum(unlist(prepared$stats)), "= 蛋白级 且 有 Description 且 非污染物 且 有 GN"),
    c("肽段级/非蛋白行", prepared$stats$peptide_level %||% 0, "Master 不是 'Master Protein' 的行，见 Removed 表"),
    c("污染物删除", prepared$stats$contaminant %||% 0, "见 Removed 表；清单可在 --contaminants 指定"),
    c("抽不到 GN 删除", prepared$stats$no_gn %||% 0, "见 Removed 表"),
    c("Description 为空删除", prepared$stats$desc_empty %||% 0, "见 Removed 表"),
    c("可见行(诱饵有值)", n_visible, ""),
    c("隐藏行(诱饵无值)", n_hidden, "分组隐藏，仍保留在表中"),
    c("Unique 行", n_unique, "诱饵有值、EV 无值；浅绿底色"),
    c("对照组第5百分位数", sprintf("%.15g", p5), "用于填补 Unique 行的对照组值（+5% 表）"),
    c("LOG2FC", "log2(诱饵 / 对照)", "未做上样量归一化；归一化与打分由 bioid-analysis 负责"),
    c("值域", NA_character_, NA_character_)
  )
  df <- as.data.frame(do.call(rbind, lapply(rows, function(r) {
    data.frame(Item = r[[1]], Value = as.character(r[[2]]),
               Note = as.character(r[[3]]), stringsAsFactors = FALSE)
  })), stringsAsFactors = FALSE)
  df
}

# ── 主程序 ──────────────────────────────────────────────────────────────────
main <- function() {
  opt <- parse_args(commandArgs(trailingOnly = TRUE))
  if (isTRUE(opt$help) || is.null(opt$input)) {
    usage()
    quit(status = if (is.null(opt$input)) 1L else 0L)
  }
  require_pkg("readxl", "读取 .xlsx")
  require_pkg("openxlsx", "写出 .xlsx")

  if (!dir.exists(opt$out)) dir.create(opt$out, recursive = TRUE, showWarnings = FALSE)

  map_path <- opt$map
  if (is.null(map_path)) {
    base <- if (dir.exists(opt$input)) opt$input else dirname(opt$input)
    map_path <- file.path(base, "channel_map.tsv")
  }
  rules <- read_channel_map(map_path)

  cont_path <- opt$contaminants %||% file.path(script_dir(), "..", "assets", "contaminants.tsv")
  contaminants <- read_contaminants(cont_path)

  if (dir.exists(opt$input)) {
    files <- list.files(opt$input, pattern = "\\.xlsx$", full.names = TRUE, ignore.case = TRUE)
    files <- sort(files)
    files <- files[!grepl("^~\\$", basename(files)) &
                     !grepl("_processed|_flagged", basename(files), ignore.case = TRUE)]
  } else if (file.exists(opt$input)) {
    files <- opt$input
  } else {
    stop(sprintf("输入路径不存在: %s", opt$input), call. = FALSE)
  }
  if (!length(files)) stop("没有找到待处理的 .xlsx 文件", call. = FALSE)

  styles <- create_styles()
  compat_ok <- patch_openxlsx_compat()
  if (!compat_ok) {
    cat(paste0("⚠ openxlsx 版本与已知的悬空 drawing 关系缺陷不匹配，未打补丁。\n",
               "  保存后会自动校验包结构；若出现悬空引用，openpyxl 将读不了这些文件。\n"))
  }
  log_lines <- c(
    sprintf("bioid_process.R 运行记录  %s", format(Sys.time(), "%Y-%m-%d %H:%M:%S")),
    sprintf("输入: %s", opt$input),
    sprintf("输出: %s", opt$out),
    sprintf("通道映射: %s", normalizePath(map_path, mustWork = FALSE)),
    sprintf("污染物清单: %s", normalizePath(cont_path, mustWork = FALSE)),
    sprintf("openxlsx %s，openpyxl 兼容补丁: %s", as.character(utils::packageVersion("openxlsx")),
            if (compat_ok) "已应用" else "未应用"),
    sprintf("待处理文件: %d", length(files)),
    ""
  )
  cat(sprintf("bioid 批处理：%d 个文件\n通道映射: %s\n\n", length(files),
              normalizePath(map_path, mustWork = FALSE)))

  all_records <- list()
  skipped <- list()
  for (f in files) {
    res <- tryCatch(process_file(f, opt$out, opt, rules, contaminants, styles),
                    error = function(e) {
                      cat(sprintf("    ✗ 处理失败: %s\n", conditionMessage(e)))
                      list(skipped = basename(f), reason = conditionMessage(e))
                    })
    if (!is.null(res$records)) all_records <- c(all_records, res$records)
    if (!is.null(res$skipped)) skipped[[length(skipped) + 1L]] <- res
  }

  cat(sprintf("\n%s\n", paste(rep("=", 68), collapse = "")))
  cat(sprintf("完成: 产出 %d 份工作簿，跳过 %d 个文件\n", length(all_records), length(skipped)))
  if (length(all_records)) {
    cat("\n诱饵 → EV 一览（请核对是否与实验设计一致）:\n")
    for (r in all_records) {
      cat(sprintf("  %-42s 诱饵 %-4s ← EV %-4s → %s\n", r$file, r$bait_token, r$ev_token, r$out))
      log_lines <- c(log_lines, sprintf("文件=%s 工作表=%s 诱饵=%s(%s) EV=%s(%s) 输出=%s 输入行=%d 保留行=%d 肽段级=%d 污染物=%d 无GN=%d Unique=%d 正常=%d 隐藏=%d EV第5百分位=%.15g",
                                        r$file, r$sheet, r$bait_token, r$bait,
                                        r$ev_token, r$ev, r$out, r$n_input, r$n_kept,
                                        r$stats$peptide_level %||% 0, r$stats$contaminant %||% 0,
                                        r$stats$no_gn %||% 0, r$unique, r$normal, r$hidden, r$p5))
    }
    if (length(unique(vapply(all_records, function(r) r$ev_token, character(1)))) > 1L) {
      cat("\n注意: 本批使用了多个不同的 EV 通道；bioid-analysis 里 EV 共享时按一个对照样本处理。\n")
    }
  }
  if (length(skipped)) {
    cat("\n跳过:\n")
    for (s in skipped) cat(sprintf("  %s —— %s\n", s$skipped, s$reason))
    log_lines <- c(log_lines, "", "跳过:", sprintf("  %s —— %s", vapply(skipped, function(s) s$skipped, character(1)),
                                                       vapply(skipped, function(s) s$reason, character(1))))
  }

  pkg_issues <- unlist(lapply(all_records, function(r) r$pkg_issues))
  if (length(pkg_issues)) {
    cat(paste0("\n⚠ 以下输出存在包结构问题，openpyxl（bioid-analysis / bioid-batch-qc）\n",
               "  会读错或读不了。补救：在 Excel 里打开并另存一次，或改用 openxlsx2 重写。\n"))
    for (d in pkg_issues) cat(sprintf("  %s\n", d))
    log_lines <- c(log_lines, "", "包结构问题:", pkg_issues)
  }

  log_path <- file.path(opt$out, "bioid_run_log.txt")
  writeLines(c(log_lines, "", "sessionInfo:", capture.output(utils::sessionInfo())), log_path)
  cat(sprintf("\n运行记录: %s\n", log_path))
  if (length(pkg_issues)) quit(status = 3L, save = "no")
  if (length(skipped)) quit(status = 2L, save = "no")
}

if (sys.nframe() == 0L) main()
