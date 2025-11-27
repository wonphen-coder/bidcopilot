# ==== 安装必需的R包
options(repos = c(CRAN = "https://mirrors.tuna.tsinghua.edu.cn/CRAN/"))

# 在应用启动前添加
check_dependencies <- function() {
  required_packages <- c("shiny", "bslib", "DT", "officer", "pdftools", 
                         "docxtractr", "tidyllm", "dplyr", "stringr", "purrr", 
                         "openxlsx", "glue", "tools", "cli", "lubridate")
  
  missing_packages <- required_packages[!sapply(required_packages, requireNamespace, quietly = TRUE)]
  
  if (length(missing_packages) > 0) {
    stop("缺少必要的R包: ", paste(missing_packages, collapse = ", "),
         "\n请运行: install.packages(c('", paste(missing_packages, collapse = "', '"), "'))")
  }
}

if (!require("pacman")) install.packages("pacman")
pacman::p_load(
  shiny, bslib, DT, officer, pdftools, docxtractr, tidyllm, dplyr, stringr,
  openxlsx, glue, tools, cli, lubridate, fontawesome, magrittr, purrr
)

# ==== 辅助工具 ----------------------
# （注：以下为原代码的核心函数，直接复用无需修改，确保解析逻辑一致）

# 1. 配置参数函数
get_bid_config <- function() {
  list(
    supported_extensions = c("docx", "doc", "pdf", "txt", "odt"),
    # 投标文件格式模板
    template_docx = "./20251030WORD样式模板.docx",
    # 最大上传文件大小：100MB
    max_file_size = 100 * 1024 * 1024,
    timeout = 300,
    # 采购需求章节
    procurement_pattern = "(?:需求|技术规格)",
    # 投标文件格式章节
    bid_format_pattern = "(?:文件[的]*格式|投标格式|投标文件组成|投标文件编制|格式附件)",
    # 合同章节
    contract_pattern = "(?:合同)",
    # 资格性与符合性审查
    audit_pattern = "(?:审查资料|审查内容|资格要求|符合性)",
    # 评分标准
    scoring_pattern = "(?:评分|分值|得分|商务部分|技术部分|价格部分)",
    audit_keywords = "(?:无效|废标|作废|实质性)",
    core_para_keywords = "截图|证书|测试报告|★|▲|☆|🔺",
    package_pattern = "包[1-9一二三四五六七八九十]?[ 、：]",
    default_model = "ollama/qwen2.5:7b",
    # Web应用使用临时目录，避免权限问题
    output_dir = tempdir(),
    # A4宽度（英寸）
    pg_width = 5.77,
    # A4高度（英寸）
    pg_height = 9.69,
    default_font = "宋体",
    default_size = 12,
    heading_font = "宋体",
    line_spacing = 1.5
  )
}

# 2. 辅助函数（目录创建、包检查、文件检测）
ensure_dir <- function(dir) {
  if (is.null(dir) || dir == "")
    return(invisible(FALSE))
  if (!dir.exists(dir)) {
    dir.create(dir, recursive = TRUE, showWarnings = FALSE)
  }
  invisible(dir.exists(dir))
}

clean_text_punctuation <- function(text) {
  text <- trimws(text)
  if (text == "") return(text)
  
  # 处理换行符前后的空格
  text <- stringr::str_replace_all(text, "\\s*\\n\\s*", "\n")
  
  # 清理中文文本之间的空格
  text <- clean_spaces(text)
  
  # 替换英文标点为中文标点
  punctuation_map <- c(
    "," = "，",
    ";" = "；",
    ":" = "：",
    "\\?" = "？",
    "!" = "！",
    "\\(" = "（",
    "\\)" = "）",
    "\\.{3}" = "……"
  )
  
  for (i in seq_along(punctuation_map)) {
    text <- stringr::str_replace_all(text, names(punctuation_map)[i], punctuation_map[i])
  }
  
  # 处理英文句号
  # text <- stringr::str_replace_all(text, "(?<![0-9])\\.(\\s|$)", "。\\1")
  
  # 处理连续标点符号：如果有两个标点结尾，保留最后一个
  text <- clean_ending_punctuation(text)
  
  # 清理多余空格但保留换行符
  text <- stringr::str_replace_all(text, "[ ]+", " ")
  text <- stringr::str_trim(text)
  
  # 处理错误的换行
  text <- fix_line_breaks(text)
  
  return(text)
}

# 清理结尾多余的标点符号
clean_ending_punctuation <- function(text) {
  # 定义中文标点符号
  chinese_punctuation <- "，。！？；："
  
  # 匹配结尾的连续标点
  while (stringr::str_detect(text, paste0("[", chinese_punctuation, "]{2,}$"))) {
    # 保留最后一个标点，移除前面的多余标点
    text <- stringr::str_replace(text,
                                 paste0("[", chinese_punctuation, "]+$"),
                                 stringr::str_sub(text, -1, -1))
  }
  
  return(text)
}

# 清理多余的空格
clean_spaces <- function(text) {
  
  chinese_chars <- "\\p{Han}"
  digits <- "0-9"
  punctuation <- "_-，。、！？；：\\."
  brackets <- "（）【】《》"
  
  all_chars <- paste0("[", chinese_chars, digits, punctuation, brackets, "]")
  pattern <- paste0("(", all_chars, ")[\\t ]+(", all_chars, ")")
  
  old_text <- ""
  while (text != old_text) {
    old_text <- text
    text <- stringr::str_replace_all(text, pattern, "\\1\\2")
  }
  
  # 在大写字母和小写字母之间插入空格（驼峰命名）
  # text <- stringr::str_replace_all(text, "([a-z])([A-Z])", "\\1 \\2")
  
  return(text)
}

# 处理错误换行的函数
fix_line_breaks <- function(text) {
  # === 1. 定义字符集 ===
  # 匹配各种序号模式并在前面插入换行符
  is_list <- paste0(
    "(?<!\\n)", # 匹配的位置前面不是换行符
    "(",
    "(?:\\s\\d{1,2}\\.)+\\d{1,2}|",             # 1.2.3、5.1、10.5
    "\\s[一二三四五六七八九十]+[、．\\.]|",     # “一、”、“十二.”
    "\\s\\d{1,2}[）、\\.．\\)]|",               # 1)、3）、7.、8、
    "（\\d{1,2}）|",                            # （1）、（12）
    "（[一二三四五六七八九十]{1,2}）|",         # （一）、（十二）
    "\\s第[一二三四五六七八九十0-9]{1,2}(:?章|部分|篇|节|条)|",
    "\\s附件[一二三四五六七八九十0-9]{1,2}：|",
    "\\s注：",
    ")"
  )
  
  # 可连接字符（不含句尾标点）
  connectable_chars <- "[\u4e00-\u9fff0-9a-zA-Z_\\-，、；：（）【】《》\\·]"
  
  # 句尾标点（不应再连接下一行）
  sentence_end <- "[。！？]$"
  
  # 标题序号模式：用于识别以章节/条目编号开头的行
  heading_start_pat <- c(
    "^[\\s]*[•●▪-]",
    # 无序符号：•, ●, ▪, -
    "^[\\s]*\\d{1,2}[\\.、\\)）]",
    # 1. 2、 3) 4）
    "^[\\s]*[\\(（]\\d{1,2}[\\)）]",
    # (1) （2）
    "^[\\s]*[一二三四五六七八九十]+[、\\.．]",
    # 一、 二. 三．
    "^[\\s]*[a-zA-Z][\\.\\)]",
    # a. b) A. B)
    "^[\\s]*第[一二三四五六七八九十0-9]{1,2}(:?章|部分|篇|节|条)",
    "^[\\s]*附件\\d*",
    "^[\\s]*注：?",
    "^\\s*[\\(（][一二三四五六七八九十]{1,2}[\\)）]",
    # (一)
    "^(?:\\d{1,2}\\.)+\\d{1,2}",
    # 多级数字序号如2.2.6, 2.2.2.26
    "："
  )
  
  heading_regex <- paste0("(", paste(heading_start_pat, collapse = "|"), ")")
  
  # 按行处理：避免把以标题序号结尾的行与下一行合并
  lines <- unlist(strsplit(text, "\\r?\\n")) |> trimws()
  lines <- lines[lines != ""]
  
  if (length(lines) <= 1) {
    lines <- stringr::str_replace_all(lines, is_list, "\n\\1")
    return(lines)
  }
    
  result_lines <- character()
  i <- 1
  
  while (i <= length(lines)) {
    current <- stringr::str_replace_all(lines[i], is_list, "\n\\1")
    
    # 判断当前行是否为标题开头，标点结尾
    is_heading_now <- stringr::str_detect(current, heading_regex)
    ends_now <- stringr::str_detect(current, sentence_end)
    
    # 初始化合并内容
    merged <- current
    j <- i
    
    # 即使是标题，只要没结束，且下一行不是新标题，就尝试继续合并
    while (j < length(lines)) {
      next_line <- stringr::str_replace_all(lines[j + 1], is_list, "\n\\1")
      next_is_heading <- stringr::str_detect(next_line, heading_regex)
      
      # 如果下一行是新标题/列表项 → 停止合并
      if (next_is_heading) break
      
      # 检查当前 merged 是否“未结束”
      current_ends <- stringr::str_detect(merged, sentence_end)
      if (current_ends) break  # 已结束，不再合并
      
      # 检查当前 merged 是否“未结束”
      current_ends <- stringr::str_detect(merged, sentence_end)
      if (current_ends) break  # 已结束，不再合并
      
      # 检查连接性
      tail_ok <- any(stringr::str_detect(merged, paste0(connectable_chars, "$")))
      head_ok <- any(stringr::str_detect(next_line, paste0("^", connectable_chars)))
      digit_join <- any(stringr::str_detect(merged, "[0-9]$")) &&
        any(stringr::str_detect(next_line, "^[0-9]"))

      if ((tail_ok && head_ok) || digit_join) {
        merged <- paste0(merged, next_line)
        j <- j + 1
      } else {
        break
      }
    }
    
    result_lines <- c(result_lines, merged)
    i <- j + 1
  }
  # === 3. 后处理：压缩多余空行 ===
  final_text <- paste(result_lines, collapse = "\n")

  return(final_text)
}

infer_style_from_text <- function(full_text) {
  # 1. 将全文字符串切分为段落向量
  paragraphs <- full_text |> 
    stringr::str_split("\n") |> 
    unlist() |> 
    stringr::str_trim() |> 
    # 移除空段落
    purrr::keep(.p = function(x) x != "")
  
  # 检查一级标题数量，以 "# " 开头
  n_headings <- sum(stringr::str_count(paragraphs, "^# "))

  # 如果数量超过10，说明原文档中样式混乱，则清除所有 "# " 标记（仅移除开头的 "# "）
  if (n_headings > 10) {
    paragraphs <- stringr::str_remove(paragraphs, "^# ")
  }

  # 2. 预计算排除状态，避免重复计算
  exclude_status <- sapply(paragraphs, function(p) {
    should_exclude_as_heading(p) || detect_date_format(p)
  })
  
  # 3. 检测文档中是否存在"第?章"或"第1章"标题
  chapter_heading_exists <- stringr::str_detect(full_text, "第[一二三四五六七八九十0-9]{1,2}(:?章|部分|篇)")
   
  # chapter_heading_exists <- any(sapply(paragraphs[!exclude_status], function(p) {
  #   any(sapply(chapter_patterns, function(pattern) stringr::str_detect(p, pattern)))
  # }))
  
  # 4. 预编译正则表达式模式，提高性能
  number_patterns <- vector("list", 9)
  for (level in 1:9) {
    dots_needed <- level - 1
    if (dots_needed == 0) {
      number_patterns[[level]] <- "^\\d+\\s*"
    } else {
      number_patterns[[level]] <- paste0("^\\d+(\\.\\d+){", dots_needed, "}\\s*")
    }
  }
  
  # 5. 根据是否存在章标题来设置级别映射
  if (chapter_heading_exists) {
    # 模式1：有章标题的情况（章→节→一、→二、→数字标题）
    level_mapping <- list(
      "^第[一二三四五六七八九十0-9]{1,2}(:?章|部分|篇)" = 1,  # 章为1级
      # "^第\\d{1,2}(:?章|部分|篇)" = 1,                      # 第1章格式
      "^第[一二三四五六七八九十万0-9]{1,2}节" = 2,            # 节为2级  
      "^[一二三四五六七八九十]{1,2}[、\\.]" = 2,              # 一、为2级
      "^Section|^Chapter" = 1,
      "default" = NA  # 非标题段落返回NA
    )
  } else {
    # 模式2：无章标题的情况（一、→二、→数字标题）
    level_mapping <- list(
      "^[一二三四五六七八九十]{1,2}[、\\.]" = 1,              # 一、升为1级
      "^Section|^Chapter" = 1,
      "default" = NA  # 非标题段落返回NA
    )
  }
  
  # 预编译中文和英文标题模式
  chinese_patterns <- names(level_mapping)[names(level_mapping) != "default"]
  
  # 6. 处理每个段落，添加标题标记
  processed_paragraphs <- character(length(paragraphs))
  
  for (i in seq_along(paragraphs)) {
    paragraph <- paragraphs[i]
    
    # 如果被排除，直接返回原段落
    if (exclude_status[i]) {
      processed_paragraphs[i] <- paragraph
      next
    }
    
    # 检测是否为标题段落
    if (!detect_section_numbering(paragraph)) {
      processed_paragraphs[i] <- paragraph
      next
    }
    
    # 确定标题级别
    level <- NA
    
    # 检查数字标题（优先级最高）
    for (lvl in 1:9) {
      if (stringr::str_detect(paragraph, number_patterns[[lvl]])) {
        # 按点数量计算级别：1个点(如1.1)为2级，2个点(如1.1.1)为3级，依此类推
        # 不管什么情况，1.1都应该是二级标题
        dots_count <- lvl - 1  # lvl=1时dots=0(如"1")，lvl=2时dots=1(如"1.1")
        
        if (dots_count == 0) {
          # 单数字(如"1")：有章标题时为2级，无章标题时为1级
          level <- if (chapter_heading_exists) NA else 1
        } else {
          # 带点的数字(如"1.1", "1.1.1")：级别 = 点数 + 1
          # 1.1 (1个点) -> 2级，1.1.1 (2个点) -> 3级
          level <- dots_count + 1
        }
        
        # 确保级别在1-6范围内
        level <- min(max(level, 1), 9)
        break
      }
    }
    
    # 如果数字标题未匹配，检查中文和英文标题
    if (is.na(level)) {
      for (pattern in chinese_patterns) {
        if (stringr::str_detect(paragraph, pattern)) {
          level <- level_mapping[[pattern]]
          break
        }
      }
    }
    
    # 如果仍未确定级别，使用默认值
    if (is.na(level)) {
      level <- level_mapping[["default"]]
    }
    
    # 如果是标题，添加相应数量的#号
    if (!is.na(level) && level >= 1 && level <= 6) {
      hashes <- paste(rep("#", level), collapse = "")
      processed_paragraphs[i] <- paste(hashes, paragraph)
    } else {
      processed_paragraphs[i] <- paragraph
    }
  }
  
  # 7. 将处理后的段落重新组合为字符串
  result_text <- paste(processed_paragraphs, collapse = "\n")
  return(result_text)
}

# 优化的检测函数（使用预编译模式）
detect_section_numbering <- function(text) {
  # 使用预编译的模式列表
  patterns <- c(
    "^第[一二三四五六七八九十0-9]+(章|部分|篇|节|条)",
    "^\\d{1,2}(\\.\\d{1,2})+\\.?\\s*",
    "^\\d{1,2}(\\.\\d{1,2})+\\.?$",
    "^[一二三四五六七八九十]+[、.]",
    "^(Section|Chapter)\\s+\\d+"
  )
  
  any(vapply(patterns, function(pattern) 
    stringr::str_detect(text, pattern), logical(1)))
}

# 其他辅助函数保持不变
should_exclude_as_heading <- function(text) {
  # 1. 长度检测
  if (stringr::str_length(text) > 30) {
    return(TRUE)
  }
  
  # 2. 标点符号密度检测
  punctuation_count <- stringr::str_count(text, "[，。；：！？、]")
  
  # 3、如果标点数量超过3，可能不是标题
  if (punctuation_count > 3) {
    return(TRUE)
  }
  
  # 4、检测是否以数字加顿号开头（可能是列表项）
  if (stringr::str_detect(text, "^\\d+、")) {
    return(TRUE)
  }
  
  # 5. 检测是否以句号结尾（标题通常不以句号结尾）
  if (stringr::str_detect(text, "。$")) {
    return(TRUE)
  }
  
  # 6. 检测是否包含引号（可能是指示性文字）
  if (stringr::str_detect(text, "[《》\"“”]")) {
    # 但如果是书名号包裹的短文本，可能是标题
    if (stringr::str_detect(text, "^《[^》]{1,20}》$")) {
      return(FALSE)  # 《XXX》格式可能是标题
    }
    return(TRUE)
  }
  
  # 7. 检测是否包含特殊符号（如@、#、$等）
  if (stringr::str_detect(text, "[@#$%^&*=]")) {
    return(TRUE)
  }
  
  # 8. 检测是否包含URL或邮箱
  if (stringr::str_detect(text, "(http|www|\\.com|\\.cn|@)")) {
    return(TRUE)
  }
  
  return(FALSE)
}

detect_date_format <- function(text) {
  date_patterns <- c(
    "^\\d{4}年\\d{1,2}月\\d{0,2}日?$",
    "^\\d{4}-\\d{1,2}-\\d{1,2}$",
    "^\\d{4}/\\d{1,2}/\\d{1,2}$",
    "^\\d{4}\\.\\d{1,2}\\.\\d{1,2}$",
    "^\\d{1,2}月\\d{1,2}日$",
    "^\\d{4}年度?$",
    "^第\\d{1,2}季度$",
    "^[一二三四五六七八九十零百千万]+年"
  )
  
  any(vapply(date_patterns, function(pattern)
    stringr::str_detect(text, pattern), logical(1)))
}

# 检查并安装Java函数
check_and_install_java <- function() {
  # 检查Java是否已安装
  java_check <- tryCatch({
    system2("java", args = c("-version"), stdout = FALSE, stderr = TRUE)
    return("java")
  }, error = function(e) {
    return(NULL)
  })

  if (!is.null(java_check)) {
    cat("  - Java已安装，使用默认路径\n")
    return("java")
  }

  cat("  - Java未安装，正在安装...\n")

  # 在Ubuntu上安装OpenJDK
  install_cmd <- "sudo apt-get update && sudo apt-get install -y openjdk-11-jre-headless"

  cat("  - 执行安装命令:", install_cmd, "\n")
  res <- tryCatch({
    system(install_cmd, intern = TRUE, ignore.stdout = FALSE)
    return(0)
  }, error = function(e) {
    warning("安装Java失败: ", e$message)
    return(1)
  })

  if (res != 0) {
    stop("Java安装失败，请手动安装: sudo apt-get install openjdk-11-jre-headless")
  }

  cat("  - Java安装成功\n")

  # 再次检查Java
  java_check <- tryCatch({
    system2("java", args = c("-version"), stdout = FALSE, stderr = TRUE)
    return("java")
  }, error = function(e) {
    stop("Java安装后仍无法找到，请检查PATH环境变量")
  })

  return(java_check)
}

# 移除页眉页脚
detect_common_headers_footers <- function(pages, top_n = 2, bottom_n = 2, min_fraction = 0.6) {
  # pages: character vector, 每个元素为一页的完整文本（含换行）
  # 返回 list(header_candidates, footer_candidates)
  page_lines_list <- lapply(pages, function(p) unlist(strsplit(p, "\r?\n")))
  npages <- length(page_lines_list)
  normalize_line <- function(l) {
    l2 <- gsub("[ \\t]+", " ", trimws(l))
    # 如果仅包含非可见字符，返回空字符串
    if (nchar(l2) == 0) return("")
    return(l2)
  }

  # 收集每页的前 top_n 行和后 bottom_n 行
  top_lines <- unlist(lapply(page_lines_list, function(lines) {
    n <- length(lines)
    if (n == 0) return(character(0))
    idx <- seq_len(min(top_n, n))
    sapply(lines[idx], normalize_line, USE.NAMES = FALSE)
  }))
  bottom_lines <- unlist(lapply(page_lines_list, function(lines) {
    n <- length(lines)
    if (n == 0) return(character(0))
    idx <- seq.int(from = max(1, n - bottom_n + 1), to = n)
    sapply(lines[idx], normalize_line, USE.NAMES = FALSE)
  }))

  # 统计出现频率（排除空串）
  freq_top <- sort(table(top_lines[top_lines != ""]), decreasing = TRUE)
  freq_bottom <- sort(table(bottom_lines[bottom_lines != ""]), decreasing = TRUE)

  thr <- ceiling(min_fraction * npages)
  header_candidates <- names(freq_top[freq_top >= thr])
  footer_candidates <- names(freq_bottom[freq_bottom >= thr])

  # 额外找寻可能的页码模式（纯数字，Page N，N / M，中文“第N页”）
  page_num_patterns <- c("^\\d+$", "^Page[ ]+\\d+$", "^\\d+\\s*/\\s*\\d+$", "^第[0-9一二三四五六七八九十百零]+页$")
  # 如果某些候选行匹配页码模式，把它们加入 footer_candidates
  all_bottom_unique <- unique(bottom_lines[bottom_lines != ""])
  for (ln in all_bottom_unique) {
    for (p in page_num_patterns) {
      if (grepl(p, ln, perl = TRUE)) {
        footer_candidates <- unique(c(footer_candidates, ln))
      }
    }
  }

  return(list(header = header_candidates, footer = footer_candidates))
}

remove_headers_footers <- function(pages, header_pattern = NULL, footer_pattern = NULL,
                                   top_n = 2, bottom_n = 2, min_fraction = 0.6) {
  # pages: character vector, 每个元素为一页的完整文本（含换行）。
  # 如果 header_pattern / footer_pattern 为 NULL 则自动检测（跨页重复）。
  if (!is.character(pages) || length(pages) == 0) return(pages)

  # 先做自动检测（当用户未提供模式）
  detected <- list(header = character(0), footer = character(0))
  if (is.null(header_pattern) || is.null(footer_pattern)) {
    detected <- detect_common_headers_footers(pages, top_n = top_n, bottom_n = bottom_n, min_fraction = min_fraction)
  }

  # 将可能的检测到的行转换为正则（匹配时做严格的 trim + 多空格宽松匹配）
  make_line_regex <- function(s) {
    # escape regex metachars, 但把空格变为 \s+以允许读取时多空格差异
    esc <- gsub("([\\^$.|?*+()\\[\\]{}\\\\])", "\\\\\\1", s, perl = TRUE)
    esc <- gsub(" ", "\\\\s+", esc)
    paste0("^\\s*", esc, "\\s*$")
  }

  header_regexes <- character(0)
  footer_regexes <- character(0)
  if (!is.null(header_pattern)) header_regexes <- c(header_regexes, header_pattern)
  if (!is.null(footer_pattern)) footer_regexes <- c(footer_regexes, footer_pattern)
  if (length(detected$header) > 0) header_regexes <- c(header_regexes, sapply(detected$header, make_line_regex, USE.NAMES = FALSE))
  if (length(detected$footer) > 0) footer_regexes <- c(footer_regexes, sapply(detected$footer, make_line_regex, USE.NAMES = FALSE))

  # 处理每页：去掉匹配 header_regexes 的开头行（通常 1-2 行）和匹配 footer_regexes 的结尾行（通常 1-2 行）
  cleaned_pages <- lapply(pages, function(page_text) {
    lines <- unlist(strsplit(page_text, "\r?\n"))
    if (length(lines) == 0) return("")
    # 规范化单行用于匹配：trim + collapse 中间多空格
    norm_line <- function(x) gsub("[ \\t]+", " ", trimws(x))

    # 移除开头
    while (length(lines) > 0 && length(header_regexes) > 0) {
      nl <- norm_line(lines[1])
      matched <- any(sapply(header_regexes, function(p) grepl(p, nl, perl = TRUE)))
      if (matched) lines <- lines[-1] else break
    }
    # 移除结尾
    while (length(lines) > 0 && length(footer_regexes) > 0) {
      nl <- norm_line(lines[length(lines)])
      matched <- any(sapply(footer_regexes, function(p) grepl(p, nl, perl = TRUE)))
      if (matched) lines <- lines[-length(lines)] else break
    }

    if (length(lines) == 0) return("")
    return(paste(lines, collapse = "\n"))
  })

  return(unlist(cleaned_pages))
}

# 读取PDF文本内容
remove_footnotes_pages <- function(pages, footnote_pattern = NULL, handle_brackets = TRUE) {
  # 支持两种模式：
  # 1) footnote_pattern 为 NULL：使用自动检测（行首编号 + URL，或编号行与下一行 URL 的跨行脚注）
  # 2) footnote_pattern 为正则字符串：按该正则删除匹配行
  if (!is.character(pages) || length(pages) == 0) return(pages)

  is_url_like <- function(s) {
    s <- trimws(s)
    if (s == "") return(FALSE)
    grepl("^(?:https?[:：]?//|www\\.|[A-Za-z0-9-]+\\.(?:com|org|net|cn|io|gov|edu|info)(?:/|$))", s, perl = TRUE, ignore.case = TRUE)
  }

  # 如果用户提供了自定义的脚注正则，按该正则删除包含匹配的行
  if (!is.null(footnote_pattern)) {
    clean_page_user <- function(p) {
      lines <- unlist(strsplit(p, "\r?\n"))
      if (length(lines) == 0) return("")
      keep <- !vapply(lines, function(ln) {
        # 匹配整行或行内包含模式均视为脚注
        res <- FALSE
        try({ res <- grepl(footnote_pattern, ln, perl = TRUE, ignore.case = TRUE) }, silent = TRUE)
        return(res)
      }, logical(1))
      kept <- lines[keep]
      if (length(kept) == 0) return("")
      return(paste(kept, collapse = "\n"))
    }
    return(vapply(pages, clean_page_user, FUN.VALUE = ""))
  }

  # 自动检测逻辑
  clean_page_auto <- function(p) {
    lines <- unlist(strsplit(p, "\r?\n"))
    if (length(lines) == 0) return("")
    keep <- rep(TRUE, length(lines))

    for (i in seq_along(lines)) {
      ln <- lines[i]
      # 1) 行首为编号且后面接 URL（或类似 URL 的片段），则删除该行
      m <- regexec("^\\s*(\\[\\s*\\d+\\s*\\]|\\d+)\\s*(.*)$", ln, perl = TRUE)
      parts <- regmatches(ln, m)[[1]]
      if (length(parts) >= 3) {
        rest <- parts[3]
        if (is_url_like(rest) || grepl("https?://", rest, perl = TRUE, ignore.case = TRUE)) {
          keep[i] <- FALSE
          next
        }
      }

      # 2) 如果当前行是 URL-like（单独一行），并且上一行仅为编号或[编号]，则删除上一行和当前行（跨行脚注）
      if (is_url_like(ln) || grepl("^\\s*https?://", ln, perl = TRUE, ignore.case = TRUE)) {
        if (i > 1) {
          prev <- lines[i - 1]
          if (grepl("^\\s*(\\[\\s*\\d+\\s*\\]|\\d+)\\s*$", prev, perl = TRUE)) {
            keep[i] <- FALSE
            keep[i - 1] <- FALSE
            next
          }
        }
        # 若上一行非编号，但上一行末尾为数字并与 URL 连续（极少见），可扩展处理，这里暂不处理
      }
    }

    kept <- lines[keep]
    if (length(kept) == 0) return("")
    return(paste(kept, collapse = "\n"))
  }

  return(vapply(pages, clean_page_auto, FUN.VALUE = ""))
}

# 合并相似的表格（通常是跨页的表格）
merge_similar_tables <- function(tables) {
  if (length(tables) <= 1) return(tables)
  
  # 辅助函数：检查两个表格是否相似
  are_tables_similar <- function(t1, t2) {
    # 检查列数是否相同
    if (ncol(t1) != ncol(t2)) return(FALSE)
    
    # 检查列名格式是否相同
    if (!all(names(t1) == names(t2))) return(FALSE)
    
    # 检查数据类型是否一致（可选）
    col_types1 <- sapply(t1, class)
    col_types2 <- sapply(t2, class)
    if (!all(col_types1 == col_types2)) return(FALSE)
    
    return(TRUE)
  }
  
  # 初始化结果列表
  merged_tables <- list()
  i <- 1
  
  while (i <= length(tables)) {
    current_table <- tables[[i]]
    
    # 如果是最后一个表格，直接添加
    if (i == length(tables)) {
      merged_tables[[length(merged_tables) + 1]] <- current_table
      break
    }
    
    # 检查下一个表格是否相似
    next_table <- tables[[i + 1]]
    if (are_tables_similar(current_table, next_table)) {
      # 合并表格
      merged <- rbind(current_table, next_table)
      merged_tables[[length(merged_tables) + 1]] <- merged
      i <- i + 2  # 跳过已合并的两个表格
    } else {
      # 如果不相似，只添加当前表格
      merged_tables[[length(merged_tables) + 1]] <- current_table
      i <- i + 1
    }
  }
  
  return(merged_tables)
}

# 过滤掉少于指定行数和列数的表格
filter_tables <- function(tables, min_rows = 3, min_cols = 2) {
  filtered <- list()
  for (table in tables) {
    if (nrow(table) >= min_rows && ncol(table) >= min_cols) {
      filtered[[length(filtered) + 1]] <- table
    }
  }
  return(filtered)
}

# 辅助函数：处理ODT元素
process_odt_elements <- function(df) {
  df <- df[order(df$doc_index), ]
  rownames(df) <- NULL
  df$text <- ifelse(is.na(df$text) |
                      trimws(df$text) == "", "", trimws(df$text))

  output_lines <- character()
  i <- 1
  n <- nrow(df)

  while (i <= n) {
    row <- df[i, ]

    # 处理标题
    if (row$content_type == "paragraph" &&
        !is.na(row$style_name) &&
        grepl("^heading \\d+$", row$style_name)) {
      level <- as.numeric(sub("heading ", "", row$style_name))
      if (is.na(level))
        level <- 1
      output_lines <- c(output_lines, paste0(strrep("#", level), " ", row$text))
      i <- i + 1
      next
    }

    # 处理表格
    if (row$content_type == "table cell") {
      table_block <- data.frame()
      while (i <= n && df$content_type[i] == "table cell") {
        table_block <- rbind(table_block, df[i, ])
        i <- i + 1
      }

      row_ids <- unique(table_block$row_id[!is.na(table_block$row_id)])
      if (length(row_ids) > 0) {
        rows_list <- lapply(row_ids, function(rid)
          as.character(table_block[table_block$row_id == rid, "text"]))
        max_cols <- max(lengths(rows_list))
        mat <- do.call(rbind, lapply(rows_list, function(r) {
          length(r) <- max_cols
          r[is.na(r)] <- ""
          r
        }))
        if (ncol(mat) >= 2) {
          headers <- mat[, 1]
          data_matrix <- mat[, -1, drop = FALSE]
          md_header <- paste0("| ", paste(headers, collapse = " | "), " |")
          md_sep <- paste0("| ", paste(rep("---", length(headers)), collapse = " | "), " |")
          md_body <- apply(data_matrix, 1, function(r)
            paste0("| ", paste(r, collapse = " | "), " |"))
          output_lines <- c(output_lines, md_header, md_sep, md_body)
        }
      }
      next
    }

    # 处理普通段落
    if (row$content_type == "paragraph" && row$text != "") {
      if (!stringr::str_detect(row$text, "PAGEREF_Toc")) {
        output_lines <- c(output_lines, row$text)
      }
      i <- i + 1
      next
    }

    i <- i + 1
  }

  return(output_lines)
}

# 辅助函数：通过LibreOffice转换处理
process_odt_via_conversion <- function(odt_path) {
  cat("步骤4: 检查LibreOffice\n")

  # 使用LibreOffice转换ODT到DOCX
  soffice <- Sys.which("soffice")
  if (soffice == "") {
    stop(
      "系统未找到LibreOffice (soffice)。请安装：\n",
      "  sudo apt-get update && sudo apt-get install -y libreoffice\n",
      "或者将文件另存为.docx后重试。"
    )
  }

  cat("步骤5: 创建临时目录\n")
  outdir <- tempfile("odt2docx_out")
  dir.create(outdir, recursive = TRUE)

  cat("步骤6: 转换ODT到DOCX\n")
  args <- c(
    "--headless",
    "--convert-to",
    "docx",
    "--outdir",
    outdir,
    normalizePath(odt_path)
  )

  res <- tryCatch({
    system2(soffice, args = args, stdout = TRUE, stderr = TRUE)
  }, warning = function(w) {
    warning("LibreOffice转换警告: ", w$message)
    return(as.character(w))
  }, error = function(e) {
    stop("转换ODT到DOCX失败: ", e$message)
  })

  docx_name <- paste0(tools::file_path_sans_ext(basename(odt_path)), ".docx")
  docx_path <- file.path(outdir, docx_name)

  if (!file.exists(docx_path)) {
    stop(
      "转换失败：生成的DOCX文件不存在。\n",
      "LibreOffice输出：\n",
      paste(res, collapse = "\n")
    )
  }

  cat("步骤7: 使用processDOCX处理\n")
  text_result <- tryCatch({
    processDOCX(docx_path)
  }, error = function(e) {
    stop("处理DOCX时出错: ", e$message)
  })

  try(unlink(outdir, recursive = TRUE), silent = TRUE)
  cat("步骤8: 转换完成\n")
  return(text_result)
}

# ==== 解析工具 ----------------------
processDOCX <- function(file_path) {
  cat("[DOCX日志] 文件路径:", file_path, "\n"); flush(stdout())

  # 检查文件
  if (!file.exists(file_path)) {
    cat("[DOCX错误] 文件不存在!\n"); flush(stdout())
    stop("文件不存在：", file_path)
  }

  # 设置超时和内存保护
  cat("[DOCX日志] 步骤1/5: 设置超时参数...\n"); flush(stdout())
  old_timeout <- getOption("timeout")
  options(timeout = 300)  # 5分钟超时
  on.exit(options(timeout = old_timeout))
  cat("[DOCX日志] ✅ 超时设置完成:", getOption("timeout"), "秒\n"); flush(stdout())

  # 清理环境
  cat("[DOCX日志] 步骤2/5: 清理内存...\n"); flush(stdout())
  gc(); flush(stdout())
  cat("[DOCX日志] ✅ 内存清理完成\n"); flush(stdout())

  # 读取DOCX
  cat("[DOCX日志] 步骤3/5: 读取DOCX文件...\n"); flush(stdout())
  cat("[DOCX日志] 正在调用officer::read_docx()...\n"); flush(stdout())

  doc <- tryCatch({officer::read_docx(file_path)}, error = function(e) {
    cat("[DOCX错误] read_docx调用失败!\n"); flush(stdout())
    stop("读取DOCX文件失败: ", e$message)
  })

  df <- tryCatch({officer::docx_summary(doc)}, error = function(e) {
    cat("[DOCX错误] docx_summary调用失败!\n")
    stop("提取文档摘要失败: ", e$message)
  })
  
  # 检查是否成功获取数据
  if (is.null(df) || nrow(df) == 0) {
    cat("[DOCX警告] 文档内容为空\n")
    return("")  # 返回空字符串而不是NULL
  }
  
  df <- df |>
    dplyr::mutate(text = trimws(text),
                  style_name = ifelse(is.na(style_name), "段落", style_name)) |>
    # 剔除目录
    dplyr::filter(!stringr::str_detect(style_name, "toc|Toc|TOC"))
  
  cat("[DOCX日志] 文档元素数量:", nrow(df), "\n")

  output_lines <- character(0)
  i <- 1
  n_total <- nrow(df)
  
  while (i <= n_total) {
    # 进度
    if (i %% 100 == 0 || i == n_total) {
      cat("[DOCX日志] 处理进度:", i, "/", n_total, sprintf("(%.1f%%)\n", i / n_total * 100))
    }
    
    # --- 表格处理 ---
    if (df$content_type[i] == "table cell") {
      cat("[DOCX日志] 在第", i, "行检测到表格\n")
      
      # 找出连续的表格行（officer 通常连续输出表格单元格）
      start_i <- i
      while (i <= n_total && df$content_type[i] == "table cell") {
        i <- i + 1
      }
      end_i <- i - 1
      
      table_block <- df[start_i:end_i, , drop = FALSE]
      
      # 重建表格
      row_ids <- sort(unique(table_block$row_id[!is.na(table_block$row_id)]))
      if (length(row_ids) > 0) {
        rows_list <- lapply(row_ids, function(rid) {
          as.character(table_block[table_block$row_id == rid, "text"])
        })
        
        max_cols <- max(lengths(rows_list))
        if (max_cols >= 2) {
          mat <- do.call(rbind, lapply(rows_list, function(r) {
            length(r) <- max_cols
            r[is.na(r) | r == ""] <- ""
            r
          }))
          
          headers <- mat[1, ]
          data_rows <- mat[-1, , drop = FALSE]
          
          md_header <- paste0("| ", paste(headers, collapse = " | "), " |")
          md_sep <- paste0("| ", paste(rep("---", length(headers)), collapse = " | "), " |")
          md_body <- apply(data_rows, 1, function(r) paste0("| ", paste(r, collapse = " | "), " |"))
          
          output_lines <- c(output_lines, md_header, md_sep, md_body, "")
          cat("[DOCX日志] ✅ 表格处理完成（行", start_i, "-", end_i, "）\n")
        }
      }
      # i 已经指向表格后下一行，继续循环
      next
    }
    
    # --- 段落/标题处理 ---
    row <- df[i, ]
    if (row$content_type == "paragraph") {
      if (!is.na(row$style_name) && grepl("^heading \\d+$", row$style_name)) {
        level <- as.numeric(sub("heading ", "", row$style_name))
        if (is.na(level)) level <- 1
        output_lines <- c(output_lines, paste0(strrep("#", level), " ", row$text))
      } else if (nchar(trimws(row$text)) > 0) {
        output_lines <- c(output_lines, row$text)
      }
    }
    # 忽略其他类型（或按需扩展）
    
    i <- i + 1
  }

  cat("[DOCX日志] ✅ processDOCX函数完成\n")
  
  # 确保返回字符串
  result_text <- paste(output_lines, collapse = "\n")
  # 从文本中推断标题级别
  result_text <- infer_style_from_text(result_text)
  cat("[DOCX日志] 最终文本长度:", nchar(result_text), "字符\n")
  cat("[DOCX日志] === processDOCX函数结束 ===\n\n")

  # 检查结果
  if (is.null(result_text) || result_text == "") {
    cat("[DOCX警告] 返回空结果\n")
  }

  return(result_text)
}

processDOC <- function(file_path) {
  # Ubuntu 24 环境下处理 DOC 文件：转换为 DOCX 后使用 processDOCX 处理
  cat("步骤1: 检查 LibreOffice\n")
  # 使用 LibreOffice 转换 DOC 到 DOCX（在 Ubuntu 24 上推荐）
  soffice <- Sys.which("soffice")
  if (soffice == "") {
    stop(
      "系统未找到 LibreOffice (soffice)。请安装：\n",
      "sudo apt-get update && sudo apt-get install -y libreoffice\n",
      "或者将文件另存为 .docx 后重试。"
    )
  }

  cat("步骤2: 创建临时目录\n")
  # 创建临时目录存储转换后的文件
  outdir <- tempfile("doc2docx_out")
  dir.create(outdir, recursive = TRUE)

  cat("步骤3: 转换 DOC 到 DOCX\n")
  # 使用 LibreOffice 将 DOC 转换为 DOCX
  args <- c(
    "--headless",      # 无头模式（无GUI）
    "--writer",        # 指定使用Writer组件
    "--convert-to", "docx",  # 转换为 DOCX
    "--outdir", outdir,      # 指定输出目录
    normalizePath(file_path, mustWork = TRUE) # 输入文件路径
  )

  # 执行转换
  res <- tryCatch({
    system2(
      command = soffice,
      args = args,
      stdout = TRUE,
      stderr = TRUE,
      timeout = 30  # 超时时间（秒），防止卡死
    )
  }, warning = function(w) {
    # 记录警告但继续
    warning("LibreOffice 转换警告: ", w$message)
    return(as.character(w))
  }, error = function(e) {
    stop("转换 DOC 到 DOCX 失败: ", e$message)
  })

  # 检查转换结果文件名（处理可能的重名后缀，如原文件已存在时LibreOffice会加(1)）
  base_name <- tools::file_path_sans_ext(basename(file_path))
  possible_files <- list.files(
    path = outdir,
    pattern = paste0("^", gsub("\\.", "\\\\.", base_name), ".*\\.docx$"),
    full.names = TRUE
  )
  if (length(possible_files) == 0) {
    stop(
      "转换失败：未生成任何 DOCX 文件。\n",
      "LibreOffice 输出：\n",
      paste(res, collapse = "\n")
    )
  }
  docx_path <- possible_files[1]  # 取第一个匹配文件（通常是目标文件）

  if (!file.exists(docx_path) || !has_document_xml(docx_path)) {
    stop(
      "转换失败：生成的 DOCX 文件无效或不存在。\n",
      "LibreOffice 输出：\n",
      paste(res, collapse = "\n"),
      "\n请检查文件是否损坏或 LibreOffice 是否正确安装"
    )
  }

  cat("步骤4: 使用 processDOCX 处理 DOCX\n")
  # 使用 processDOCX 函数处理转换后的 DOCX
  text_result <- tryCatch({
    processDOCX(docx_path)
  }, error = function(e) {
    stop("处理 DOCX 时出错: ", e$message)
  })

  # 清理临时文件（可选）
  try(unlink(outdir, recursive = TRUE), silent = TRUE)

  cat("步骤5: 转换完成\n")
  return(text_result)
}

# 解析pdf并自动去除页眉页脚
processPDF <- function(file_path, remove_headers = TRUE, header_pattern = NULL, footer_pattern = NULL,
                       remove_footnotes = FALSE, footnote_pattern = NULL) {
  cat("\n=== [PDF日志] processPDF函数开始 ===\n"); flush(stdout())
  cat("[PDF日志] 文件路径:", file_path, "\n"); flush(stdout())
  cat("[PDF日志] 当前工作目录:", getwd(), "\n"); flush(stdout())
  cat("[PDF日志] 当前时间:", Sys.time(), "\n"); flush(stdout())

  # 检查文件大小
  file_size <- file.size(file_path)
  cat("[PDF日志] 文件大小:", round(file_size/1024/1024, 2), "MB\n"); flush(stdout())

  # 设置超时和内存保护
  cat("[PDF日志] 步骤1/4: 设置超时参数...\n"); flush(stdout())
  old_timeout <- getOption("timeout")
  options(timeout = 300)  # 5分钟超时
  cat("[PDF日志] ✅ 超时设置完成:", getOption("timeout"), "秒\n"); flush(stdout())

  # 清理环境
  cat("[PDF日志] 步骤2/4: 清理内存...\n"); flush(stdout())
  gc(); flush(stdout())
  cat("[PDF日志] ✅ 内存清理完成\n"); flush(stdout())

  # 使用简单可靠的pdf_text，避免pdf_data导致内存问题
  cat("[PDF日志] 步骤3/4: 开始调用pdf_text...\n"); flush(stdout())
  cat("[PDF日志] 预计耗时: 大文件可能需要1-3分钟，请耐心等待...\n"); flush(stdout())

  result <- tryCatch({
    # 调用pdf_text
    pdftools::pdf_text(file_path)
  }, error = function(e) {
    cat("[PDF错误] pdf_text调用失败!\n")
    cat("[PDF错误] 错误信息:", e$message, "\n")
    print(sys.calls())
    stop("无法读取PDF文件: ", e$message, "\n\n详细日志:\n", e$message, "\n\n建议：\n1. 检查文件是否损坏\n2. 如果是扫描版PDF，请先OCR识别\n3. 尝试使用文本型PDF\n4. 确保有足够的内存和磁盘空间")
  })
  
  # 检查结果
  if (is.null(result) || length(result) == 0) {
    stop("pdf_text返回NULL或空结果！")
  } else {
    cat("[PDF日志] ✅ pdf_text调用成功\n")
    cat("[PDF日志] 返回页数:", length(result), "\n")
    cat("[PDF日志] 总文本长度:", sum(nchar(result)), "字符\n")
  }
  
  # 移除页眉页脚（可选）
  cat("[PDF日志] 步骤4/4: 处理页眉页脚...\n")
  if (remove_headers) {
    cat("[PDF日志] 正在移除页眉页脚...\n")
    result <- remove_headers_footers(result, header_pattern, footer_pattern)
    cat("[PDF日志] ✅ 页眉页脚处理完成\n")
  }

  # 移除脚注（可选）
  if (remove_footnotes) {
    cat("[PDF日志] 正在移除脚注...\n")
    result <- remove_footnotes_pages(result, footnote_pattern)
    cat("[PDF日志] ✅ 脚注处理完成\n")
  }
  
  # 按页合并
  result <- paste0(result, collapse = "")

  # 合并每个段落内的行（用空格连接）
  fixed_paragraphs <- fix_line_breaks(result)
  # 从文本中推断标题级别
  full_text <- infer_style_from_text(fixed_paragraphs)
  cat("[PDF日志] ✅ processPDF函数完成\n")
  cat("[PDF日志] 最终文本长度:", sum(nchar(full_text)), "字符\n")
  cat("[PDF日志] === processPDF函数结束 ===\n\n")
  # 返回 single character string
  return(full_text)  
}

# 另一种解析pdf并去除页眉页脚的方法
processPDF2 <- function(file_path,
                        header_margin = 50,
                        footer_margin = 50) {
  # （完整复用原代码中的 processPDF2 函数，负责解析PDF并去除页眉页脚）
  message("正在解析PDF: ", basename(file_path))
  pages <- try(pdf_text(file_path), silent = TRUE)
  if (inherits(pages, "try-error")) {
    stop("无法解析PDF文件（可能是扫描版、加密或损坏）")
  }
  info <- pdf_pagesize(file_path)
  data_list <- pdf_data(file_path)
  clean_pages <- character()
  
  for (i in seq_along(data_list)) {
    page_data <- data_list[[i]]
    page_height <- info$height[i]
    if (nrow(page_data) == 0) {
      clean_pages <- c(clean_pages, "")
      next
    }
    in_body <- subset(page_data,
                      y > footer_margin &
                        y < (page_height - header_margin))
    in_body$x <- in_body$x / info$width[i] * 100
    in_body$y <- in_body$y / page_height * 100
    in_body$line <- as.integer(cut(in_body$y, breaks = seq(0, 100, by = 4)))
    in_body <- in_body[order(in_body$line, in_body$x), ]
    in_body <- aggregate(
      text ~ line,
      data = in_body,
      FUN = function(col)
        paste(col, collapse = " ")
    )
    clean_page <- paste(in_body$text, collapse = "\n")
    clean_pages <- c(clean_pages, clean_page)
    clean_pages <- trimws(gsub(" +", " ", clean_pages))
  }
  message("PDF解析完成！共 ", length(pages), " 页。")
  return(paste(clean_pages, collapse = "\n"))
}

processODT <- function(file_path) {
  # 处理 ODT (Open Document Text) 文件
  file_size <- file.size(file_path)
  if (file_size > 100 * 1024 * 1024) {
    stop(paste("文件过大（", round(file_size/1024/1024, 2), "MB），请使用小于100MB的文件"))
  }

  # 设置超时
  old_timeout <- getOption("timeout")
  options(timeout = 300)  # 5分钟超时
  on.exit(options(timeout = old_timeout))

  cat("步骤1: 尝试使用officer包直接读取\n")

  # 尝试使用officer包读取ODT（如果支持）
  tryCatch({
    doc <- officer::read_docx(file_path)
    df <- officer::docx_summary(doc)

    # 转换为与processDOCX相同的输出格式
    output_lines <- process_odt_elements(df)
    cat("步骤2: 直接读取成功\n")
    return(paste(output_lines, collapse = "\n"))
  }, error = function(e) {
    # 如果officer不支持ODT，使用LibreOffice转换
    cat("步骤3: 转换为DOCX\n")
    return(process_odt_via_conversion(file_path))
  })
}

# 处理 TXT 文件：读取后直接返回文本
processTXT <- function(file_path) {
  file_size <- file.size(file_path)
  if (file_size > 100 * 1024 * 1024) {
    stop(paste("文件过大（", round(file_size/1024/1024, 2), "MB），请使用小于100MB的文件"))
  }

  cat("步骤1: 读取TXT文件\n")
  text <- tryCatch({
    paste(readLines(file_path, warn = FALSE), collapse = "\n")
  }, error = function(e) {
    stop("读取TXT文件失败: ", e$message)
  })
  
  # 从文本中推断标题级别
  full_text <- infer_style_from_text(text)
  
  cat("步骤2: 读取完成\n")
  return(full_text)
}



# ==== 抽取工具 ----------------------
# 招标关键信息提取函数，返回数据框
fun_extract_tender <- function(txt) {
  regex_list <- list(
    项目名称     = "项目名称\\s*[：:]\\s*(.+?)(?=\\n|$)",
    项目编号     = "(?:招标|项目)编号\\s*[：:]\\s*(.+?)(?=\\n|$)",
    采购人       = "[采购人单位\\s]{3,}[：:]\\s*(.+?)(?=\\n|$)",
    招标代理机构 = "(?:招标|采购)代理[机构\\s]*[：:]\\s*(.+?)(?=\\n|$)",
    采购内容     = "(?:采购内容|项目内容|采购需求|招标内容)\\s*[：:\\|]\\s*(.+?)(?=\\n|$)",
    `采购预算/限价` = "(?:最高限价|控制价|预算金额|采购预算)\\s*[：:\\|]*\\s*(.+?)(?=\\n|$)",
    项目属性     = "项目属性\\s*[：:\\|]\\s*(.+?)(?=\\n|$)",
    投标保证金   = "(?:投标保证金|保证金)\\s*[：:\\|]\\s*(.+?)(?=\\n|$)",
    合同履行期限 = "(?:合同履行期限|工期|交货期)\\s*[：:\\|]\\s*(.+?)(?=\\n|$)",
    开标时间     = "(?:截止|开标)时间\\s*[：:]\\s*(.+?)(?=\\n|$)",
    投标有效期   = "(?:投标有效期)\\s*[：:]\\s*(.+?)(?=\\n|$)"
  )
  
  find_valid_match <- function(pattern, text) {
    matches <- str_match_all(text, pattern)
    
    # 情况1: 完全无匹配 → matches 是 0 行矩阵
    if (is.null(matches) || length(matches) == 0) {
      return(NA_character_)
    }
    
    mat <- matches[[1]]
    
    # 情况2: 匹配结果为空矩阵
    if (nrow(mat) == 0 || ncol(mat) < 2) {
      return(NA_character_)
    }
    
    captures <- mat[, 2]
    valid_captures <- captures[!is.na(captures) & nchar(str_trim(captures)) > 0]
    
    if (length(valid_captures) == 0) {
      return(NA_character_)
    }
    
    for (val in valid_captures) {
      cleaned <- val |> 
        # 清除括号里的内容
        str_remove("（[^）]*）") |> 
        # 清除行首行尾的标点
        str_remove("^[。，：！？,!?\\|]+") |> 
        str_remove("[。，：！？,!?\\|]+$") |>
        str_squish()
      if (!str_detect(cleaned, "见|偏离") && cleaned != "") {
        return(cleaned)
      }
    }
    
    return(NA_character_)
  }
  
  res <- vapply(regex_list, function(p) {
    find_valid_match(p, txt)
  }, FUN.VALUE = character(1), USE.NAMES = TRUE)
  
  df <- data.frame(
    `信息类型` = names(regex_list),
    `提取结果` = as.character(res),
     check.names = FALSE
  )
  
  # 提取总包数
  # 从文本中使用 gregexpr + regmatches 提取所有 "包数字"
  matches <- tryCatch({
    regmatches(txt, gregexpr(config$package_pattern, txt))
  }, error = function(e) {
    max_packages <- 1
  })
  # 如果有匹配项，提取数字部分
  if (length(matches) > 0) {
    numbers <- try(matches |>
                     purrr::map(stringr::str_extract, "\\d+") |>
                     unlist() |>
                     purrr::discard(
                       .p = function(x)
                         is.na(x)
                     ) |>
                     as.numeric() |>
                     max(),
                   silent = TRUE)
    
    # 如果没有有效的数字，使用默认值
    max_packages <- ifelse(numbers > 1, numbers, 1)
  }
  
  df <- bind_rows(df, data.frame(
    `信息类型` = "总包数",
    `提取结果` = as.character(max_packages)
  ))
  
  return(df)
}

# 按章节拆分，返回章节名和章节内容数据框
fun_split_by_chapter <- function(txt) {
  ## 参数检查
  if (length(txt) != 1L || !is.character(txt))
    stop("'txt' 必须是单个字符串")
  
  if (nchar(trimws(txt)) == 0) {
    warning("输入文本为空，返回空数据框")
    return(data.frame(
      title = character(),
      content = character(),
      stringsAsFactors = FALSE
    ))
  }
  
  ## 1. 按换行拆行（支持 Windows/Unix 换行）
  lines <- strsplit(txt, "\r?\n")[[1]]
  pat <- tryCatch({
    lines |> 
      # 从完整文本中提取所有标题
      stringr::str_extract_all("^#\\s.*") |> 
      unlist()
  }, error = function(e)
    stop(conditionMessage(e)))

  if (!is.character(pat) || anyNA(pat) || length(pat) == 0L)
    stop("'pat' 必须是长度≥1 的字符向量且不含 NA！")
  
  ## 2. 找到整行完全匹配的行号
  idx <- which(lines %in% pat)
  if (length(idx) == 0L) {
    warning("未匹配到任何标题行，返回空数据框。")
    return(data.frame(
      title = character(),
      content = character(),
      stringsAsFactors = FALSE
    ))
  }
  
  ## 3. 计算每个区间
  n <- length(lines)
  # 内容开始行
  from <- idx + 1
  # 内容结束行
  to   <- c(idx[-1] - 1, n)
  # 处理“标题在最后一行”的边界
  from[from > n] <- NA
  to  [to   < 1] <- NA
  
  ## 4. 提取并拼接
  titles   <- lines[idx]
  contents <- mapply(function(s, e) {
    if (is.na(s))
      return("")
    paste(lines[s:e], collapse = "\n")
  }, from, to, USE.NAMES = FALSE)
  
  ## 5. 返回
  # data.frame(
  #   title = titles,
  #   content = trimws(contents),
  #   stringsAsFactors = FALSE
  # )
  res <- data.frame(
    title = titles,
    content = trimws(contents),
    stringsAsFactors = FALSE
  )
  # 合并内容长度小于1000的章节
  if (nrow(res) > 1) {
    i <- 2
    while (i <= nrow(res)) {
      if (nchar(res$content[i], type = "chars") < 1000) {
        sep <- if (nzchar(res$content[i - 1]) && nzchar(res$content[i])) "\n" else ""
        res$content[i - 1] <- paste0(res$content[i - 1], sep, res$content[i])
        res <- res[-i, , drop = FALSE]
      } else {
        i <- i + 1
      }
    }
  }
  return(res)
}

# 本工具实现对根据章节名匹配提取对应章节内容
fun_extract_chapter <- function(chapters,
                                pattern,
                                return_mode = c("longest", "last", "first")) {
  return_mode <- match.arg(return_mode)
  
  # 输入验证
  if (!is.data.frame(chapters) ||
      !all(c("title", "content") %in% names(chapters))) {
    stop("chapters 必须是包含 'title' 和 'content' 列的数据框")
  }
  
  if (nrow(chapters) == 0) {
    warning("输入章节数据框为空")
    return(
      data.frame(
        title = NA_character_,
        content = NA_character_,
        stringsAsFactors = FALSE
      )
    )
  }
  
  tryCatch({
    chapters_filtered <- chapters |>
      dplyr::mutate(
        full_length = nchar(content),
        match_score = stringr::str_count(title, pattern)
      ) |>
      dplyr::filter(match_score > 0) |>
      dplyr::arrange(desc(match_score))
    
    if (nrow(chapters_filtered) == 0) {
      message("未找到匹配模式 '", pattern, "' 的章节")
      return(
        data.frame(
          title = NA_character_,
          content = NA_character_,
          stringsAsFactors = FALSE
        )
      )
    }
    
    # 根据模式选择
    if (return_mode == "longest") {
      selected <- chapters_filtered |>
        dplyr::filter(full_length == max(full_length, na.rm = TRUE)) |>
        dplyr::slice(1)
    } else if (return_mode == "last") {
      selected <- chapters_filtered |>
        dplyr::slice(n())
    } else {
      selected <- chapters_filtered |>
        dplyr::slice(1)
    }
    
    # 输出信息
    cat("✅ 成功提取章节：", selected$title[1], "\n")
    cat("📊 章节字符数：", selected$full_length[1], "\n")
    cat("🔍 内容预览：",
        stringr::str_sub(selected$content[1], 1, 200),
        "...\n\n")
    
    return(as.data.frame(selected))
    
  }, error = function(e) {
    warning("提取章节时出错: ", e$message)
    return(
      data.frame(
        title = NA_character_,
        content = NA_character_,
        stringsAsFactors = FALSE
      )
    )
  })
}

# 表格清洗
clean_tbl <- function(tbl) {
  if (!is.data.frame(tbl) || nrow(tbl) == 0 || ncol(tbl) == 0) return(NULL)
  # 删除空行
  row_keep <- apply(tbl, 1, function(row) {
    any(!is.na(row) & trimws(as.character(row)) != "")
  })
  tbl <- tbl[row_keep, , drop = FALSE]
  if (nrow(tbl) == 0) return(NULL)
  # 删除空列
  col_keep <- apply(tbl, 2, function(col) {
    any(!is.na(col) & trimws(as.character(col)) != "")
  })
  tbl <- tbl[, col_keep, drop = FALSE]
  if (ncol(tbl) <= 1 || nrow(tbl) <= 1) return(NULL) # 只剩表头或单列
  tbl
}

# 从 DOCX 文件中提取所有表格
fun_extract_tables_docx <- function(docx_path) {
  tables <- docxtractr::read_docx(docx_path) |>
    docxtractr::docx_extract_all_tbls() |>
    purrr::map(clean_tbl) |>
    purrr::compact()  # 丢掉 NULL
  if (length(tables) == 0) {
    warning("文档中未找到任何表格")
    list()
  } else {
    tables
  }
}

# 从 TXT 文件中提取表格（以多个空格分割）
#'
#' @param txt 文本字符串
#' @param min_cols 最少列数（默认：2）
#' @return 数据框列表，每个数据框代表一个表格
fun_extract_tables_txt <- function(txt, min_cols = 2) {
  # 将文本按行分割
  lines <- unlist(strsplit(txt, "\n", fixed = TRUE))
  lines <- trimws(lines)

  tables_list <- list()

  for (i in seq_along(lines)) {
    line <- lines[i]

    # 使用多个空格或制表符分割行
    cells <- strsplit(line, "\\s{2,}")[[1]]
    cells <- trimws(cells)

    # 过滤空单元格
    cells <- cells[cells != ""]

    # 检查列数是否满足要求
    if (length(cells) >= min_cols) {
      # 检查下一行是否也是表格内容（通过相同的列数判断）
      next_line_idx <- grep(paste0("^", line, "$"), lines) + 1
      if (next_line_idx <= length(lines)) {
        next_line <- lines[next_line_idx]
        next_cells <- strsplit(next_line, "\\s{2,}")[[1]]
        next_cells <- trimws(next_cells)
        next_cells <- next_cells[next_cells != ""]

        # 如果下一行也是表格内容，扩展表格
        if (length(next_cells) >= min_cols && all(next_cells != cells)) {
          # 找到表格块的开始和结束
          table_block <- list(cells)
          start_idx <- grep(paste0("^", line, "$"), lines)

          # 读取连续的行，直到列数变化
          for (j in start_idx:(start_idx + 100)) {
            if (j > length(lines)) break
            row <- trimws(lines[j])
            if (row == "") break

            row_cells <- strsplit(row, "\\s{2,}")[[1]]
            row_cells <- trimws(row_cells)
            row_cells <- row_cells[row_cells != ""]

            # 如果列数匹配，添加到表格块
            if (length(row_cells) == length(cells)) {
              table_block[[length(table_block) + 1]] <- row_cells
            } else {
              break
            }
          }

          # 转换为数据框
          if (length(table_block) >= 2) {
            table_df <- as.data.frame(do.call(rbind, table_block), stringsAsFactors = FALSE)
            colnames(table_df) <- table_df[1, ]
            table_df <- table_df[-1, ]
            rownames(table_df) <- NULL

            tables_list[[length(tables_list) + 1]] <- table_df
          }
        } else {
          # 单行表格，转换为数据框
          table_df <- as.data.frame(matrix(cells, nrow = 1, byrow = FALSE), stringsAsFactors = FALSE)
          colnames(table_df) <- paste0("col", seq_along(cells))
          rownames(table_df) <- NULL

          tables_list[[length(tables_list) + 1]] <- table_df
        }
      } else {
        # 单行表格，转换为数据框
        table_df <- as.data.frame(matrix(cells, nrow = 1, byrow = FALSE), stringsAsFactors = FALSE)
        colnames(table_df) <- paste0("col", seq_along(cells))
        rownames(table_df) <- NULL

        tables_list[[length(tables_list) + 1]] <- table_df
      }
    }
  }

  # 简化表格查找：基于分值/评分关键词
  # 如果没有找到表格，使用简化方法
  if (length(tables_list) == 0) {
    # 查找包含数字和项目的行
    for (line in lines) {
      if (trimws(line) == "" || nchar(line) < 10) next

      # 检查行是否包含数字和文本（可能是表格行）
      if (grepl("\\d+", line) && grepl("[^0-9\\s]+", line)) {
        # 使用多个空格分割
        cells <- strsplit(line, "\\s{2,}")[[1]]
        cells <- trimws(cells)
        cells <- cells[cells != ""]

        # 过滤纯数字行或纯文本行
        if (length(cells) >= min_cols && !all(grepl("^[0-9]+$", cells))) {
          table_df <- as.data.frame(matrix(cells, nrow = 1, byrow = FALSE), stringsAsFactors = FALSE)
          colnames(table_df) <- paste0("col", seq_along(cells))
          rownames(table_df) <- NULL

          tables_list[[length(tables_list) + 1]] <- table_df
        }
      }
    }
  }

  return(tables_list)
}

# 从pdf文件中提取表格
fun_extract_tables_pdf <- function(file_path, extra_args = NULL) {
  cat("步骤1: 检测和安装Java\n")
  java_path <- check_and_install_java()

  cat("步骤2: 检查tabula.jar\n")
  # 更安全地查找tabula.jar
  app_dir <- dirname(normalizePath("app.R"))
  if (!file.exists(file.path(app_dir, "app.R"))) {
    # 如果当前目录没有app.R，尝试在工作目录查找
    app_dir <- getwd()
  }
  
  tabula_jar <- list.files(
    path = app_dir,
    pattern = "^tabula.*\\.jar$",
    full.names = TRUE
  )
  
  if (length(tabula_jar) == 0) {
    stop("tabula.jar 未找到。请将tabula.jar放在app.R所在目录下")
  }
  
  # 取最后一个匹配，一般最新版
  tabula_jar <- tabula_jar[length(tabula_jar)]  

  cat("步骤3: 创建临时输出文件\n")
  out_file <- tempfile(fileext = ".csv")
  on.exit({
    if (file.exists(out_file)) {
      unlink(out_file)
    }
  })
  
  # 尝试不同的提取参数组合
  param_sets <- list(
    # 方案A：使用格子检测（推荐用于有明确边框的表格）
    c("-l", "-p", "all", "-f", "CSV", "-o", shQuote(out_file), shQuote(normalizePath(file_path))),
    
    # 方案B：使用流模式（推荐用于无边框表格）
    c("-f", "CSV", "-p", "all", "-o", shQuote(out_file), shQuote(normalizePath(file_path))),
    
    # 方案C：指定区域（自动检测）
    c("-f", "CSV", "-p", "all", "-a", "0,0,1000,1000", "-o", shQuote(out_file), shQuote(normalizePath(file_path)))
  )
  
  tables <- list()
  
  for (i in seq_along(param_sets)) {
    cat("尝试参数方案", i, "\n")
    
    args <- c("-jar", shQuote(tabula_jar), param_sets[[i]])
    if (!is.null(extra_args)) args <- c(args, extra_args)
    
    cat("执行命令: java", paste(args, collapse = " "), "\n")
    
    # 执行Tabula
    res <- system2(
      command = java_path,
      args = args,
      stdout = TRUE,
      stderr = TRUE,
      wait = TRUE
    )
    
    status <- attr(res, "status")
    if (!is.null(status) && status != 0) {
      cat("方案", i, "失败，状态码:", status, "\n")
      next
    }
    
    if (!file.exists(out_file) || file.info(out_file)$size == 0) {
      cat("方案", i, "未生成有效输出\n")
      next
    }
    
    cat("输出文件大小:", file.info(out_file)$size, "字节\n")
    
    # 解析CSV文件
    current_tables <- tryCatch({
      parse_tabula_output(out_file)
    }, error = function(e) {
      cat("方案", i, "解析失败:", e$message, "\n")
      list()
    })
    
    if (length(current_tables) > 0) {
      cat("方案", i, "成功提取", length(current_tables), "个表格\n")
      
      # 检查表格质量（列数大于1才认为是有效表格）
      valid_tables <- current_tables[sapply(current_tables, function(tbl) ncol(tbl) > 1)]
      if (length(valid_tables) > 0) {
        tables <- valid_tables
        cat("找到有效表格，使用方案", i, "\n")
        break
      } else {
        cat("方案", i, "提取的表格列数不足\n")
      }
    }
    
    # 清理临时文件，为下一次尝试做准备
    if (file.exists(out_file)) unlink(out_file)
  }
  
  # 如果所有方案都失败，尝试基于文本的解析
  if (length(tables) == 0) {
    cat("所有Tabula方案失败，尝试基于文本解析\n")
    tables <- extract_tables_from_text(file_path)
  }
  
  cat("最终提取", length(tables), "个表格\n")
  return(tables)
}

# CSV解析函数
parse_tabula_output <- function(csv_file) {
  lines <- readLines(csv_file, warn = FALSE, encoding = "UTF-8")
  if (length(lines) == 0) return(list())
  
  tables <- list()
  
  # 检测表格分隔（多个连续逗号可能表示表格边界）
  comma_counts <- sapply(strsplit(lines, ","), length)
  avg_commas <- mean(comma_counts)
  
  # 如果平均逗号数很少，可能是单列数据
  if (avg_commas < 2) {
    cat("检测到单列数据，尝试重新解析\n")
    return(parse_single_column_data(lines))
  }
  
  # 按空行分割表格
  is_blank <- grepl("^\\s*$", lines)
  if (all(!is_blank)) {
    blocks <- list(lines)
  } else {
    blocks <- split(lines[!is_blank], cumsum(is_blank)[!is_blank])
  }
  
  for (blk in blocks) {
    if (length(blk) == 0) next
    
    # 尝试读取CSV
    df <- tryCatch({
      read.csv(
        text = paste(blk, collapse = "\n"),
        header = FALSE,
        stringsAsFactors = FALSE,
        fill = TRUE,
        blank.lines.skip = FALSE
      )
    }, error = function(e) {
      NULL
    })
    
    if (!is.null(df) && nrow(df) > 0 && ncol(df) > 0) {
      # 清理数据
      df <- clean_dataframe(df)
      if (ncol(df) > 1) {  # 只保留多列表格
        colnames(df) <- paste0("Col", seq_len(ncol(df)))
        tables[[length(tables) + 1]] <- df
        cat("  表格: ", nrow(df), "行 x", ncol(df), "列\n")
      }
    }
  }
  
  return(tables)
}

# 处理单列数据的函数
parse_single_column_data <- function(lines) {
  tables <- list()
  current_table <- NULL
  current_block <- character()
  
  for (line in lines) {
    line_trim <- trimws(line)
    
    if (line_trim == "") {
      # 空行可能表示表格边界
      if (length(current_block) > 0) {
        df <- try_create_table_from_text(current_block)
        if (!is.null(df) && ncol(df) > 1) {
          tables[[length(tables) + 1]] <- df
        }
        current_block <- character()
      }
    } else {
      current_block <- c(current_block, line_trim)
    }
  }
  
  # 处理最后一个块
  if (length(current_block) > 0) {
    df <- try_create_table_from_text(current_block)
    if (!is.null(df) && ncol(df) > 1) {
      tables[[length(tables) + 1]] <- df
    }
  }
  
  return(tables)
}

# 从文本块尝试创建表格
try_create_table_from_text <- function(text_block) {
  if (length(text_block) < 2) return(NULL)
  
  # 尝试按多个空格分割
  rows <- lapply(text_block, function(line) {
    parts <- strsplit(trimws(line), "\\s{2,}")[[1]]
    parts[parts != ""]
  })
  
  # 检查是否所有行都有相同的列数
  col_counts <- sapply(rows, length)
  if (length(unique(col_counts)) == 1 && unique(col_counts) > 1) {
    df <- as.data.frame(do.call(rbind, rows), stringsAsFactors = FALSE)
    colnames(df) <- paste0("Col", seq_len(ncol(df)))
    return(df)
  }
  
  return(NULL)
}

# 清理数据框
clean_dataframe <- function(df) {
  # 移除全空列
  non_empty_cols <- sapply(df, function(col) any(!is.na(col) & col != ""))
  df <- df[, non_empty_cols, drop = FALSE]
  
  # 移除全空行
  non_empty_rows <- apply(df, 1, function(row) any(!is.na(row) & row != ""))
  df <- df[non_empty_rows, , drop = FALSE]
  
  return(df)
}

# 基于文本的备选方案
extract_tables_from_text <- function(file_path) {
  cat("使用pdftools提取文本并解析表格\n")
  
  tryCatch({
    text <- pdftools::pdf_text(file_path)
    all_tables <- list()
    
    for (page_text in text) {
      lines <- strsplit(page_text, "\n")[[1]]
      tables <- parse_text_tables(lines)
      all_tables <- c(all_tables, tables)
    }
    
    return(all_tables)
    
  }, error = function(e) {
    cat("文本提取失败:", e$message, "\n")
    return(list())
  })
}

# 解析文本表格
parse_text_tables <- function(lines) {
  tables <- list()
  current_table <- NULL
  in_table <- FALSE
  
  for (line in lines) {
    line_trim <- trimws(line)
    
    # 简单的表格检测逻辑
    if (is_potential_table_row(line_trim)) {
      if (!in_table) {
        in_table <- TRUE
        current_table <- character()
      }
      current_table <- c(current_table, line_trim)
    } else {
      if (in_table && length(current_table) >= 2) {
        # 尝试将当前块转换为表格
        df <- try_create_table_from_text(current_table)
        if (!is.null(df)) {
          tables[[length(tables) + 1]] <- df
        }
        in_table <- FALSE
      }
    }
  }
  
  return(tables)
}

# 判断是否为可能的表格行
is_potential_table_row <- function(line) {
  if (nchar(line) < 10) return(FALSE)
  
  # 包含数字和文字的混合
  has_digits <- grepl("\\d", line)
  has_text <- grepl("[a-zA-Z\\u4e00-\\u9fff]", line)
  
  return(has_digits && has_text)
}

# 基于表头信息，从多个表格中提取需求的表格
fun_extract_from_tables <- function(tables, pattern = NULL) {
  if (length(tables) == 0)
    return(list())

  matched_tables <- lapply(tables, function(tbl) {
    # 检查列名是否匹配
    titles <- names(tbl) |> str_remove_all("\\s")
    col_match <- any(grepl(pattern, titles, ignore.case = TRUE))

    # 检查首行是否匹配
    # 使用 tbl[1, ] 获取第一行，然后转为字符向量
    first_row <- as.character(tbl[1, ]) |> str_remove_all("\\s")
    first_match <- any(grepl(pattern, first_row, ignore.case = TRUE))

    # 如果列名匹配，返回表格
    if (col_match && nrow(tbl) > 0) {
      names(tbl) <- as.character(titles)
      return(tbl)
    } else if(first_match && nrow(tbl) > 0) {
      # 使用第一行作为列名
      # new_names <- as.character(tbl[1, ]) |> str_remove_all("\\s")
      # names(tbl) <- new_names
      # tbl <- tbl[-1, ]  # 删除第一行
      return(tbl)
    } else {
      return(NULL)
    }
  })
  
  # 移除NULL元素
  matched_tables <- matched_tables[!sapply(matched_tables, is.null)]
  
  cat("找到", length(matched_tables), "个匹配 “", pattern, "” 的表格。\n")
  return(matched_tables)
}

# 从字符串中提取包含关键词的行，返回字符向量
fun_extract_sentences <- function(text, keywords) {
  if (is.null(text) || trimws(text) == "") 
    return("未提取到相关项！")
  # \\R是匹配任何类型的换行符
  sentences <- unlist(strsplit(text, "\\R", perl = TRUE)) |>
    stringr::str_squish() |>
    purrr::keep(.p = function(x) x != "")
  
  if (length(sentences) == 0) 
    return("未提取到相关项！")
  
  matched <- sentences[stringr::str_detect(sentences, keywords)]
  matched <- stringr::str_remove_all(matched, "^\\|+|\\|+$") |>
    # 清除行首行尾的标点
    stringr::str_remove_all("^\\d") |> 
    stringr::str_remove_all("^[\\| ]") |> 
    stringr::str_remove_all("[\\| ]$") |> 
    stringr::str_squish() |>
    purrr::keep(.p = function(x) x != "")
  
  if (length(matched) > 0) matched else "未提取到相关项！"
}

# 从多个表格中提取条款
fun_extract_items_from_tables <- function(tables_list) {
  cat("[表格] 开始提取，表格列表长度:", length(tables_list), "\n")

  if (length(tables_list) == 0) {
    cat("[表格] 表格列表为空，返回空结果\n")
    return(character(0))
  }

  items <- character(0)
  for (tbl_idx in seq_along(tables_list)) {
    tbl <- tables_list[[tbl_idx]]
    cat("[表格] 处理表格", tbl_idx, "，维度:", nrow(tbl), "行 x", ncol(tbl), "列\n")

    if (nrow(tbl) > 0) {
      for (i in 1:nrow(tbl)) {
        row_text <- paste(as.character(tbl[i, ]), collapse = " | ")
        # 放宽长度限制，至少3个字符
        if (nchar(row_text) >= 3) {
          items <- c(items, row_text)
        }
      }
    }
  }

  cat("[条款] 提取到", length(items), "条内容\n")
  return(items)
}

# 检查DOCX文件是否有效
has_document_xml <- function(docx_path) {
  tryCatch({
    zip_info <- utils::unzip(docx_path, list = TRUE)
    any(grepl("word/document\\.xml", zip_info$Name, ignore.case = TRUE))
  }, error = function(e) {
    FALSE
  })
}

# 高亮关键词函数（可配置关键字，红色加粗显示）
#'
#' @param text 要处理的文本字符串
#' @param keywords 关键字向量，用于匹配和替换
#' @param color 高亮颜色（默认：红色 "red"）
#' @param bold 是否加粗（默认：TRUE）
#' @return 高亮处理后的HTML字符串
#' @examples
#' fun_bold("这是无效标书和废标条款", keywords = c("无效", "废标"))
fun_bold <- function(text,
                     keywords,
                     color = "red",
                     bold = TRUE) {
  if (is.null(text) || is.na(text) || text == "") {
    return(text)
  }
  
  if (is.null(keywords) || length(keywords) == 0) {
    return(text)
  }
  
  # 处理每个关键字
  for (keyword in keywords) {
    if (is.null(keyword) || keyword == "")
      next
    
    # 获取关键字长度
    keyword_len <- nchar(keyword)
    
    # 构建HTML标签
    font_weight <- if (bold)
      "font-weight: bold;"
    else
      ""
    html_tag <- sprintf('<span style="color: %s; %s">%s</span>',
                        color,
                        font_weight,
                        keyword)
    
    # 替换所有匹配的关键字（区分大小写）
    text <- gsub(keyword, html_tag, text, fixed = TRUE)
  }
  
  return(text)
}

# 合并项目信息、评分办法、无效条款等所有提取内容
fun_extract_all_audit_terms <- function(txt, bid_info, audit_tables, score_tables, package_no = 1) {
  # 提取各部分内容
  if (length(audit_tables) > 0) {
    audit_content <- fun_extract_items_from_tables(audit_tables)
    cat("[审计] 提取到资格评审内容", length(audit_content), "条\n")
  } else {
    audit_content <- NA_character_
  }
  
  # 将评分标准转换为评分条款
  if (length(score_tables) > 0) {
    score_item <- fun_extract_items_from_tables(score_tables)
    cat("[审计] 提取到评分条款", length(score_item), "条\n")
  } else {
    score_item <- NA_character_
  }
  
  # 提取无效条款
  void_term <- try({
    result <- fun_extract_sentences(txt, keywords = config$audit_keywords)
    cat("[审计] 提取到废标条款", length(result), "条\n")
    result
  }, silent = TRUE
  )
  
  # 合并所有提取结果
  all_terms <- c(
    if(nrow(bid_info) > 0) apply(bid_info, 1, function(row) paste(row, collapse = "：")),
    audit_content,
    if(length(void_term) != 0 || void_term != "未提取到废标项！") void_term else character(0),
    score_item
  )

  cat("[审计] 合并前共有", length(all_terms), "条内容\n")

  # 清理和过滤
  audit_term <- all_terms |>
    # 去除空值
    purrr::keep(.p = function(x) !is.na(x) & x != "") |> 
    # 去除行首的标点符号和数字
    stringr::str_remove_all("^[[:punct:][:space:]一二三四五六七八九十\\|]+") |>
    # 去除行尾的标点符号
    stringr::str_remove_all("[[:punct:][:space:]\\|]+$") |>
    # 过滤极短的字符串（小于6个字符）
    purrr::keep(.p = function(x) nchar(x) >= 6) |>
    # 只去除明显是章节标题的行（包含"第X章"且后面是空白或标点）
    purrr::keep(.p = function(x) !str_detect(x, "^第[一二三四五六七八九十百零0-9]+章\\s*$")) |>
    unique()

  cat("[审计] 过滤后剩余", length(audit_term), "条评审条款\n")

  # 如果过滤后为空，返回至少一个提示信息
  if (length(audit_term) == 0) {
    cat("[警告] 所有评审条款都被过滤掉了，返回提示信息\n")
    return("未提取到评审条款（请检查文件格式或手动添加）")
  }

  return(audit_term)
}


# ==== 生成工具 ----------------------
call_llm <- function(model,
                     prompt,
                     temperature = 0.7,
                     textfile = NULL,
                     timeout = 60,
                     stream = FALSE,
                     show_progress = TRUE,
                     top_p = 0.9,
                     max_tries = 3,
                     verbose = FALSE,
                     json_schema = NULL,
                     tools = NULL,
                     seed = NULL,
                     stop = NULL,
                     frequency_penalty = 0,
                     presence_penalty = 0) {
  # 参数验证
  if (missing(model) || missing(prompt)) {
    stop("model 和 prompt 参数是必需的")
  }
  
  # 模型名称格式验证
  if (!grepl(".+/.+", model)) {
    stop("模型名称格式应为 'provider/model_name'，例如 'ollama/llama3.2:3b'")
  }
  parts <- strsplit(model, "/")[[1]]
  provider <- parts[1]
  model_name <- parts[2]
  
  # 映射 provider 到 tidyllm 支持的后端（tidyllm 使用相同命名）
  supported_providers <- c("ollama", "qwen", "deepseek", "kimi", "openai", "ctyun")
  if (!(provider %in% supported_providers)) {
    stop(glue(
      "不支持的模型提供商: {provider}。支持: {paste(supported_providers, collapse = ', ')}"
    ))
  }
  
  # 显示进度
  if (show_progress) {
    cli_alert_info(glue("正在调用模型: {model} ..."))
    start_time <- Sys.time()
  }
  
  # 获取 API Key（优先环境变量）
  get_api_key <- function(env_var) {
    key <- Sys.getenv(env_var, unset = "")
    if (key == "")
      stop(glue("请设置环境变量 {env_var}"))
    return(key)
  }
  
  # 根据provider选择调用方式
  select_provider <- switch(
    provider,
    "ollama" = ollama(),
    "kimi" = openai(.api_url = "https://api.moonshot.cn/v1"),
    "ctyun" = openai(.api_url = "https://wishub-x1.ctyun.cn/v1"),
    "qwen" = openai(.api_url = "https://dashscope.aliyuncs.com/compatible-mode/v1"),
    "deepseek" = deepseek(),
    "openai" = chatgpt()
  )
  # 调用 tidyllm::llm()
  # tidyllm 会自动：
  # - 从环境变量读取 API key（如 QWEN_API_KEY）
  # - 调用对应后端
  # - 处理 Ollama 本地请求
  tryCatch({
    result_text <- tidyllm::llm_message(.prompt = prompt, .textfile = textfile) |>
      tidyllm::chat(
        .provider = select_provider,
        .model = model_name,
        .timeout = timeout,
        .stream = stream,
        .temperature = temperature,
        .top_p = top_p,
        .max_tries = max_tries,
        .verbose = verbose,
        .json_schema = json_schema,
        .tools = tools,
        .seed = seed,
        .stop = stop,
        .frequency_penalty = frequency_penalty,
        .presence_penalty = presence_penalty
      ) |>
      tidyllm::get_reply()
  }, error = function(e) {
    if (show_progress)
      cli_alert_danger("模型调用失败！")
    stop("模型调用错误: ", e$message)
  })
  
  if (show_progress) {
    elapsed <- round(difftime(Sys.time(), start_time, units = "secs"), 2)
    cli_alert_success(glue("模型调用成功（耗时 {elapsed} 秒）"))
  }
  
  return(result_text)
}

#' 从评分文本中提取结构化评分规则
#'
#' @param score_item 字符向量，每个元素表示一个评分条目（或多行组成的条目）
#' @param model 字符串，所要调用的模型（传递给 call_llm）
#' @param ... 透传给 call_llm 的其他参数（temperature, timeout, json_schema 等）
#' @return 字符向量或列表，每个元素为 LLM 返回的结构化 JSON 字符串（后续需要 fromJSON 解析）
#' @examples
#' # rules <- fun_extract_scoring_rules(c("技术评分 30分：..."), model = "ollama/qwen2.5:7b")
fun_extract_scoring_rules <- function(score_item,
                                     model,
                                     textfile = NULL,
                                     timeout = 60,
                                     stream = FALSE,
                                     show_progress = TRUE,
                                     temperature = 0.7,
                                     top_p = 0.9,
                                     max_tries = 3,
                                     verbose = FALSE,
                                     json_schema = NULL,
                                     tools = NULL,
                                     seed = NULL,
                                     stop = NULL,
                                     frequency_penalty = 0,
                                     presence_penalty = 0) {
  output_schema <- tidyllm_schema(
    `评分因素` = field_chr("如“商务资质”、“价格评分”、“技术评分”、“类似案例”、“企业资质”、“实施方案”等"),
    `评分标准` = field_chr("如“按提供方案的合理性、符合性、完整性评分”等，保留原文关键信息，可适当精简"),
    `分值` = field_dbl("满分值，如 2、3、10"),
    `评审类型` = field_fct("“主观分”或“客观分”", .levels = c("主观分", "客观分"))
  )
  
  score_rules <- vector()
  # Step 1: 准备文本
  for (text_to_parse in score_item) {
    # text_to_parse <- paste(score_item, collapse = "\n")
    # Step 2: 构造 Prompt
    prompt <- glue::glue(
      "
      你是一名招投标专家，请从以下招标文件文本中，提取所有用于评标打分的评审条目，并以结构化表格形式输出，每条包含四个字段：
      评分因素, 评分标准, 分值, 评审类型。
      要求：
      忽略计分办法、排名规则、流程说明、政策优惠等非打分条目；
      每个打分项单独一行，不要合并；
      分值必须为数字，若为区间（如“1-2分”）则取上限或按原文明确值；
      以扁平化表格形式输出，确保每行对应一个评分条目，列名为：评分因素, 评分标准, 分值, 评审类型，列之前以“|”分隔；
      如果未提取到“分值”列，返回空表格。

      评分细则文本：
      {text_to_parse}
      "
    )
    
    # 调用 call_llm，支持自动读环境变量）
    response <- call_llm(
      model = model,
      prompt = prompt,
      textfile = textfile,
      timeout = timeout,
      stream = stream,
      show_progress = show_progress,
      temperature = temperature,
      top_p = top_p,
      max_tries = max_tries,
      verbose = verbose,
      json_schema = output_schema,
      tools = tools,
      seed = seed,
      stop = stop,
      frequency_penalty = frequency_penalty,
      presence_penalty = presence_penalty
    )
    
    score_rules <- c(score_rules, response)
  }
  return(score_rules)
}

fun_generate_rules <- function(score_items, model = "ollama/qwen2.5:7b", config = config) {
  # 生成评分细则：仅在存在评分条目时调用 LLM
  score_rules <- character(0)
  if (length(score_items) == 0) {
    message("未提取到评分条目，跳过评分细则生成。")
  } else {
    score_rules <- tryCatch(
      fun_extract_scoring_rules(
        score_items,
        model = model,
        textfile = NULL,
        timeout = 120,
        temperature = 0.3,
        show_progress = TRUE,
        stream = FALSE,
        json_schema = NULL,
        top_p = NULL,
        max_tries = NULL,
        verbose = NULL,
        tools = NULL,
        seed = NULL,
        stop = NULL,
        frequency_penalty = NULL,
        presence_penalty = NULL
      ),
      error = function(e) {
        warning("调用 LLM 生成评分细则失败：", e$message)
        character(0)
      }
    )
  }
  
  # 解析 LLM 返回的 JSON（如果有）并合并为数据框
  df_rules <- data.frame()
  if (length(score_rules) > 0) {
    parsed_list <- lapply(score_rules, function(js) {
      tryCatch({
        fromJSON(js, simplifyDataFrame = TRUE) |> as.data.frame()
      }, error = function(e) {
        warning("解析 JSON 失败，跳过该条：", e$message)
        NULL
      })
    })
    # 移除解析失败的 NULL
    parsed_list <- parsed_list[!sapply(parsed_list, is.null)]
    
    if (length(parsed_list) > 0) {
      df_rules <- tryCatch({
        do.call(rbind, parsed_list)
      }, error = function(e) {
        warning("合并评分规则列表失败：", e$message)
        data.frame()
      })
    }
  }
  
  # 仅在存在列名为 `分值` 且有行时过滤并写出
  if (nrow(df_rules) > 0 && "分值" %in% names(df_rules)) {
    df_rules <- dplyr::filter(df_rules, `分值` > 0)
    cat("✅ 成功生成评分细则：\n")
    print(df_rules)

    write.csv(
      df_rules,
      file.path(config$output_dir, "extracted_score_rules.csv"),
      row.names = FALSE,
      fileEncoding = "GB18030"
    )
  } else {
    message("未生成有效的评分细则（没有可用的分值列或结果为空），跳过写入 CSV。")
  }
  
  cat("✅ 文件处理完成\n")
  return(df_rules)
}

# ==== Shiny应用代码 ----------------------
## ui ----
ui <- page_sidebar(
  # 页面基础配置
  title = "智能投标助手 - BidCopilot",
  theme = bs_theme(
    version = 5,
    bg = "#ffffff",
    # 背景色：纯白（提高对比度）
    fg = "#000000",
    # 文字色：纯黑（最高对比度）
    primary = "#007bff",
    # 主色调：标准蓝色（更易识别）
    secondary = "#6c757d",
    # 辅助色：深灰（中等对比度）
    base_font = font_google("Noto Sans SC") # 中文友好字体（适配招标文档常见字体）
  ),
  # 侧边栏区域（上传文件和配置参数）
  sidebar = sidebar(
    width = 350,
    # 1. 文件上传组件
    div(
      style = "margin-bottom: 15px;",
      tags$p(style = "font-weight: bold; margin-bottom: 10px; color: #007bff;", "📄 点击或拖拽上传招标文件"),
      tags$p(
        style = "font-size: 12px; color: #6c757d; margin-bottom: 10px;",
        "支持文件格式：DOCX、DOC、ODT、PDF、TXT",
        br(),
        style = "font-size: 12px; color: #6c757d; margin-bottom: 10px;",
        "单文件大小：<=100MB"
      ),
      tags$style(
        HTML(
          "
      .shiny-input-container:has(#upload_file) {
        padding-top: 0 !important;
      }
      /* 增加文件上传区域的高度，所有元素统一高度 */
      .shiny-input-container input[type=file] {
        height: 120px !important;
      }
      .shiny-input-container .btn-file {
        height: 120px !important;
        line-height: 120px !important;
        font-size: 16px !important;
        padding: 0 30px !important;
      }
      .shiny-input-container .input-group {
        height: 120px !important;
      }
      .shiny-input-container .input-group-btn {
        height: 120px !important;
      }
      .shiny-input-container .form-control {
        height: 120px !important;
        line-height: 120px !important;
        font-size: 16px !important;
        padding: 40px 15px !important;
      }
      /* 确保所有表格列宽设置正确 */
      #file_basic_info_table th {
        background-color: #f0f8ff !important;
      }
      /* 标签自适应换行 */
      #upload_file label {
      white-space: normal;
      line-height: 1.4;
    }
    "
        )
      ),
      fileInput(
        inputId = "upload_file",
        label = NULL,
        multiple = FALSE, # 是否允许一次选择多个文件
        accept = c(
          ".docx",
          ".doc",
          ".odt",
          ".pdf",
          ".txt",
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          "application/msword",
          "application/vnd.oasis.opendocument.text",
          "application/pdf",
          "text/plain"
        ),
        buttonLabel = "点击选择文件",
        placeholder = "或拖拽文件到此处",
        width = "100%"
      )
    ),
    br(),
    # 2. 解析参数配置（折叠面板，默认收起，保持界面简洁）
    accordion(accordion_panel(
      title = "高级配置",
      # 包号选择（针对多包招标文档）
      fluidRow(
        column(
          width = 6,
          selectInput(
            inputId = "package_no",
            label = "投标包号",
            choices = 1:5,
            selected = 1,
            width = "100%"
          )
        ),
        column(
          width = 6,
          # 总包数将显示为只读文本（从文档提取后自动设置）
          div(
            style = "margin-bottom: 25px;",
            tags$label("总包数", class = "control-label", style = "font-weight: bold; color: #333; margin-bottom: 5px; display: block;"),
            verbatimTextOutput("total_packages_display", placeholder = TRUE)
            # tags$p("说明：总包数从文档中自动提取，不可修改", style = "font-size: 12px; color: #666; margin-top: 2px;")
          )
        )
      ),
      br(),
      # LLM模型选择（默认使用本地Ollama模型，避免API依赖）
      textInput(
        inputId = "llm_model",
        label = "LLM模型名称",
        value = "ollama/qwen2.5:7b",
        placeholder = "格式：provider/model_name（如 openai/gpt-3.5-turbo）",
        width = "100%"
      ),
      tags$style(
        HTML(
          "
      /* 调整模型名称输入框高度 */
      #llm_model {
        height: 50px !important;
        line-height: 50px !important;
        padding: 10px !important;
        font-size: 14px !important;
      }
    "
        )
      )
    )),
    br(),
    # 3. 解析按钮（突出显示，引导用户操作）
    actionButton(
      inputId = "parse_btn",
      label = "开始解析文件",
      class = "btn-primary btn-lg",
      width = "100%",
      icon = icon("play-circle") # 增加图标，提升视觉引导
    ),
    # 4. 解析状态提示（初始隐藏，解析时显示）
    uiOutput("parse_status")
  ),
  
  # 主内容区域（分标签页展示不同解析结果，避免信息混乱）
  tabsetPanel(
    # ==== 标签1：解析概览 ====================
    # 基础信息，优先展示
    tabPanel(
      title = "解析概览",
      icon = icon("info-circle"),
      # 标签1专用的CSS样式
      tags$head(tags$style(
        HTML(
          "
          /* 解析概览页面的自定义样式 */
          #file_basic_info_table th {
            background-color: #f0f8ff;
          }

          /* 解析概览表格列宽设置 */
          #file_basic_info_table td:nth-child(1),
          #file_basic_info_table th:nth-child(1) {
            width: 40% !important;
            min-width: 40% !important;
            max-width: 40% !important;
          }
          #file_basic_info_table td:nth-child(2),
          #file_basic_info_table th:nth-child(2) {
            width: 20% !important;
            min-width: 20% !important;
            max-width: 20% !important;
          }
          #file_basic_info_table td:nth-child(3),
          #file_basic_info_table th:nth-child(3) {
            width: 10% !important;
            min-width: 10% !important;
            max-width: 10% !important;
          }
          #file_basic_info_table td:nth-child(4),
          #file_basic_info_table th:nth-child(4) {
            width: 30% !important;
            min-width: 30% !important;
            max-width: 30% !important;
          }
        "
        )
      )),
      br(),
      # 文档基础信息（文件名称、大小、格式、解析时间）
      card(card_header("文档基础信息"), tableOutput("file_basic_info")),
      br(),
      # 招标核心信息（项目名称、编号、采购人等）
      tags$head(tags$style(
        HTML(
          "
          /* 招标核心信息表格列宽设置 */
          #tender_core_info_table td:nth-child(1),
          #tender_core_info_table th:nth-child(1) {
            width: 20% !important;
            min-width: 20% !important;
            max-width: 20% !important;
            text-align: center !important;
          }
          #tender_core_info_table td:nth-child(2),
          #tender_core_info_table th:nth-child(2) {
            width: 80% !important;
            min-width: 80% !important;
            max-width: 80% !important;
          }
        "
        )
      )),
      # 招标核心信息（项目名称、编号、采购人等）
      card(card_header("招标核心信息"), tableOutput("tender_core_info"))
    ),
    
    # ==== 标签2：章节概要 ====================
    # 文档章节结构预览，供用户核对
    tabPanel(
      title = "章节结构",
      icon = icon("file-alt"),
      # 标签2专用的CSS样式
      tags$head(tags$style(
        HTML("
          /* 章节概要页面的文本样式 */
          #raw_text_preview {
            line-height: 1.6;
          }
        ")
      )),
      br(),
      card(
        card_header("文档章节结构预览"),
        # 滚动容器，避免页面过长
        div(style = "height: 500px; overflow-y: auto; white-space: pre-wrap;",
            textOutput("raw_text_preview"))
      )
    ),
    
    # ==== 标签3：核心参数 ====================
    tabPanel(
      title = "核心参数",
      icon = icon("info-circle"),
      # 标签4专用的CSS样式
      tags$head(tags$style(
        HTML(
          "
          /* 页面的表格样式 */
          #audit_terms_table th {
            background-color: #fff5f5;
          }
          /* 表格列宽设置 */
          #audit_terms_table td:nth-child(1),
          #audit_terms_table th:nth-child(1) {
            width: 80px !important;
            min-width: 80px !important;
            max-width: 80px !important;
            text-align: center !important;
          }
          #audit_terms_table td:nth-child(2),
          #audit_terms_table th:nth-child(2) {
            width: calc(100% - 80px) !important;
            min-width: calc(100% - 80px) !important;
            max-width: calc(100% - 80px) !important;
          }
          /* 关键词高亮样式 */
          .audit-keyword {
            color: #d9534f;
            font-weight: bold;
          }
        "
        )
      )),
      br(),
      card(
        card_header("核心招标参数"),
        # 标红关键词，提升可读性
        div(style = "overflow-x: auto;", uiOutput("core_parameters"))
      )
    ), # ==== 标签4：评分标准 ====================
    # 核心功能之一，结构化展示
    tabPanel(
      title = "评分标准",
      icon = icon("list-ol"),
      # 标签3专用的CSS样式
      tags$head(tags$style(
        HTML(
          "
          /* 表头背景色 */
          #score_items_table th {
            background-color: #e6f7e6 !important;
          }
          /* 第一列居中 */
          #score_items_table td:nth-child(1),
          #score_items_table th:nth-child(1) {
            text-align: center !important;
          }
        "
        )
      )),
      br(),
      card(
        card_header("对应包号的评分标准"),
        # 支持分页、搜索的表格（只支持单个表格）
        # DTOutput("score_items_table")
        tableOutput("score_items_table")
      ),
      br()
      # 结构化评分规则（LLM生成结果，可选）
      # uiOutput("structured_score_rules")
    ), # ==== 标签5：评审条款 ====================
    # 废标/无效条款，重点标注
    tabPanel(
      title = "评审条款",
      icon = icon("check-circle"),
      # 标签4专用的CSS样式
      tags$head(tags$style(
        HTML(
          "
          /* 评审条款页面的表格样式 */
          #audit_terms_table th {
            background-color: #fff5f5;
          }
          /* 评审条款表格列宽设置 */
          #audit_terms_table td:nth-child(1),
          #audit_terms_table th:nth-child(1) {
            width: 80px !important;
            min-width: 80px !important;
            max-width: 80px !important;
            text-align: center !important;
          }
          #audit_terms_table td:nth-child(2),
          #audit_terms_table th:nth-child(2) {
            width: calc(100% - 80px) !important;
            min-width: calc(100% - 80px) !important;
            max-width: calc(100% - 80px) !important;
          }
          /* 评审条款的关键词高亮样式 */
          .audit-keyword {
            color: #d9534f;
            font-weight: bold;
          }
        "
        )
      )),
      br(),
      card(
        card_header("投标文件评审依据"),
        # 标红关键词，提升可读性
        div(style = "overflow-x: auto;", uiOutput("audit_terms"))
      ),
      br(),
      # 审核依据导出（Excel格式，方便用户后续使用）
      fluidRow(
        column(
          width = 6,
          downloadButton(
            outputId = "download_audit",
            label = "导出评审依据（Excel）",
            class = "btn-secondary",
            icon = icon("file-excel"),
            style = "width: 100%;"
          )
        ),
        column(
          width = 6,
          downloadButton(
            outputId = "download_bid_format",
            label = "导出投标文件格式（Word）",
            class = "btn-primary",
            icon = icon("file-word"),
            style = "width: 100%;"
          )
        )
      )
    )
  ), # 页面底部：版权与说明（简洁，不干扰主功能）
  footer = tags$footer(style = "text-align: center; padding: 20px; color: #7f8c8d;", tags$p("智能投标助手 - BidCopilot ©2025 | 基于 R Shiny 开发"))
)

# Server逻辑
server <- function(input, output, session) {
  # ---- 1. 响应式变量 ----------------
  parse_results <- reactiveValues(
    file_info        = data.frame(),  # 初始化空数据框
    tender_info      = data.frame(),
    all_score_tables = list(),
    score_tables     = list(),
    audit_terms      = list(),        # 初始化空列表
    raw_text         = character(0),  # 初始化空字符向量
    core_parameters  = data.frame(),
    chapters         = data.frame(),
    structured_rules = data.frame()
  )

  # 应用状态（存储从文档提取的总包数）
  app_state <- reactiveValues(
    total_packages = 1  # 默认为1
  )
  
  # ---- 2. 上传文件后提示 ----------------
  observeEvent(input$upload_file, {
    if (is.null(input$upload_file)) {
      output$parse_status <- renderUI(NULL)
    } else {
      output$parse_status <- renderUI(
        tags$div(
          class = "alert alert-success",
          style = "padding:10px;margin-top:10px;",
          tags$strong("✅ 文件已上传"),
          " - 点击下方按钮开始解析"
        )
      )
    }
  }, ignoreInit = TRUE)
  
  # ---- 3. 解析按钮 ----------------
  observeEvent(input$parse_btn, {
    # 在解析开始前强制垃圾回收
    gc()
    cat("\n\n")
    cat("╔══════════════════════════════════════════════════════════════╗\n")
    cat("║                开始解析文档 - 调试模式                       ║\n")
    cat("╚══════════════════════════════════════════════════════════════╝\n")
    cat("[主日志] 解析开始时间:", format(Sys.time(), "%Y-%m-%d %H:%M:%S"), "\n")
    cat("[主日志] 文件信息:", input$upload_file$name, "\n")
    cat("[主日志] 文件大小:", round(input$upload_file$size/1024/1024, 2), "MB\n")

    # 检查是否上传文件
    if (is.null(input$upload_file)) {
      showModal(modalDialog(title = "提示", "请先上传招标文档", footer = modalButton("确定")))
      return()
    }

    path <- input$upload_file$datapath
    cat("[主日志] 文件路径:", path, "\n")

    if (!file.exists(path)) {
      showModal(modalDialog(
        title = "错误",
        "文件不存在或路径错误，请重新上传",
        footer = modalButton("确定")
      ))
      return()
    }
    
    withProgress(message = "正在解析文档...", value = 0, {
      tryCatch({
        ## 1：获取文件信息----
        incProgress(0.1, detail = "获取文件信息...")
        parse_results$file_info <- data.frame(
          `文件名` = input$upload_file$name,
          `大小`   = paste0(round(input$upload_file$size / 1024, 2), " KB"),
          `格式`   = toupper(tools::file_ext(input$upload_file$name)),
          `解析时间` = format(Sys.time(), "%Y-%m-%d %H:%M:%S"),
          check.names = FALSE
        )
        
        # 2：读取文件----
        incProgress(0.3, detail = "读取文件...")
        ext <- tolower(tools::file_ext(path))
        
        if (!ext %in% config$supported_extensions) {
          stop("不支持的文件格式，请使用 DOCX、DOC、ODT、PDF 或 TXT 格式")
        }

        # 验证文件大小
        file_size <- file.size(path)
        if (file_size > config$max_file_size) {
          stop(paste("文件过大（", round(file_size/1024/1024, 2), "MB），请使用小于", config$max_file_size /1024/1024, "MB的文件！"))
        }

        # 开始读取文件
        text <- tryCatch({
          # 根据文件格式选择处理方法
          switch(ext,
            docx = processDOCX(path),
            doc = processDOC(path),
            pdf = processPDF(path),
            txt = processTXT(path),
            odt = processODT(path),
            stop("不支持的文件格式")
          )
        }, error = function(e) {
          stop(paste("读取", ext, "文件失败:", e$message))
        })
        parse_results$raw_text <- text
        
        # 3：提取招标信息----
        incProgress(0.5, detail = "提取招标信息...")
        bid_info <- tryCatch({
          fun_extract_tender(text)
        }, error = function(e) {
          stop(paste("提取招标信息失败:", e$message))
        })
        parse_results$tender_info <- bid_info
        
        # 保存解析后的文本到.md文件（优先执行）
        cat("[保存] 开始保存解析文本...\n")
        save_result <- tryCatch({
          # 创建data目录
          data_dir <- file.path(getwd(), "data")
          if (!dir.exists(data_dir)) {
            dir.create(data_dir, recursive = TRUE, showWarnings = FALSE)
            cat("[保存] 创建data目录:", data_dir, "\n")
          }

          # 生成文件名（使用项目名称或默认名称）
          if (!is.null(bid_info) && nrow(bid_info) > 0) {
            project_name <- bid_info$`提取结果`[bid_info$信息类型 == "项目名称"]
            if (!is.na(project_name) && nchar(project_name) > 0) {
              # 清理文件名，移除非法字符
              safe_name <- stringr::str_replace_all(project_name, "[^\\w\\u4e00-\\u9fa5_-]", "_")
              md_filename <- paste0(format(Sys.time(), "%Y%m%d_%H%M%S_"), safe_name, ".md")
            } else {
              md_filename <- paste0(format(Sys.time(), "%Y%m%d_%H%M%S"), "_document.md")
            }
          } else {
            md_filename <- paste0(format(Sys.time(), "%Y%m%d_%H%M%S"), "_document.md")
          }

          md_path <- file.path(data_dir, md_filename)
          cat("[保存] 文件路径:", md_path, "\n")

          # 将原始文本写入.md文件（使用UTF-8编码）
          con <- file(md_path, "w", encoding = "UTF-8")
          on.exit({
            if (exists("con") && isOpen(con)) {
              close(con)
            }
          }, add = TRUE)
          writeLines(text, con, useBytes = TRUE)
          close(con)

          cat("[保存] 解析文本已成功保存到:", md_path, "\n")
          cat("[保存] 文件大小:", file.info(md_path)$size, "字节\n")
          TRUE
        }, error = function(e) {
          cat("[错误] 保存.md文件失败:", e$message, "\n")
          FALSE
        })

        # 4：提取章节----
        incProgress(0.55, detail = "提取章节...")
        tryCatch({
          parse_results$chapters <- fun_split_by_chapter(text)
        }, error = function(e) {
          warning(paste("提取章节失败:", e$message))
        })
        
        # 5：提取所有表格----
        incProgress(0.7, detail = "提取所有表格...")
        tryCatch({
          if (ext == "docx") {
            # 提取所有表格
            tabs <- fun_extract_tables_docx(path)
          } else if (ext %in% c("doc", "odt")) {
            # DOC/ODT文件：转换为DOCX后提取表格
            output_dir <- tempfile("tables_extract_")
            dir.create(output_dir, recursive = TRUE)
            on.exit(try(unlink(output_dir, recursive = TRUE), silent = TRUE))

            docx_name <- paste0(tools::file_path_sans_ext(basename(path)), ".docx")
            temp_docx <- file.path(output_dir, docx_name)

            # 如果是doc文件，转换为DOCX
            if (ext == "doc") {
              soffice <- Sys.which("soffice")
              if (soffice != "") {
                args <- c("--headless", "--convert-to", "docx", "--outdir", output_dir, path)
                system2(soffice, args = args, stdout = FALSE, stderr = FALSE)
                if (file.exists(temp_docx)) {
                  # 提取所有表格
                  tabs <- fun_extract_tables_docx(temp_docx)
                }
              }
            } else if (ext == "odt") {
              # 如果是odt文件，转换为DOCX
              soffice <- Sys.which("soffice")
              if (soffice != "") {
                args <- c("--headless", "--convert-to", "docx", "--outdir", output_dir, path)
                system2(soffice, args = args, stdout = FALSE, stderr = FALSE)
                if (file.exists(temp_docx)) {
                  tabs <- fun_extract_tables_docx(temp_docx)
                }
              }
            }
          } else if (ext == "pdf") {
            # PDF文件：使用PDF表格提取函数
            tabs <- fun_extract_tables_pdf(path)
          } else if (ext == "txt") {
            # TXT文件：使用TXT表格提取函数
            tabs <- fun_extract_tables_txt(text, min_cols = 2)
          }
        }, error = function(e) {
          warning(paste("提取表格失败:", e$message))
        })

        # 6：提取资格审核表格----
        incProgress(0.8, detail = "提取资格审查条款...")
        # 提取资格评审项
        audit_tables <- fun_extract_from_tables(tabs, pattern = config$audit_pattern)
        if (length(audit_tables) == 0) {
            cat("[警告] 未提取到资格评审表格\n")
        }

        # 7：提取评分标准表格----
        incProgress(0.85, detail = "提取评分标准...")
        # 提取评分标准
        score_tables <- fun_extract_from_tables(tabs, pattern = config$scoring_pattern)
        # 检查评分标准表格的数量
        num <- length(score_tables)
        cat("[主日志] 总共提取到", num, "个评分标准表格\n")
        
        for (i in 1:num) {
          if (!is.null(score_tables[[i]]) && nrow(score_tables[[i]]) > 0) {
            tbl <- score_tables[[i]]
            # # 为表格添加行序号列
            # tbl_with_index <- cbind(`序号` = seq_len(nrow(tbl)), tbl)
            parse_results$all_score_tables[[i]] <- tbl
            cat("[主日志] 表格", i, "：", nrow(tbl), "行 x", ncol(tbl), "列\n")
          } else {
            cat("[警告] 表格", i, "为空，跳过\n")
          }
        }
        
        cat("[主日志] 提取到有效评分标准表格数量:", num, "\n")

        # 多个表格时，选择用户指定的包号对应的表格
        package_no <- isolate(input$package_no)
        max_packages <- as.numeric(bid_info$`提取结果`[bid_info$信息类型 == "总包数"])

        cat("[主日志] 本项目共有", max_packages, "个包\n")
        
        # 如果只有一个包，返回所有表格
        if (max_packages == 1 && num >= 1) {
          parse_results$score_tables <- parse_results$all_score_tables
        } else if (max_packages > 1 && max_packages >= num) {
          parse_results$score_tables <- parse_results$all_score_tables[[package_no]]
        } else if (max_packages > 1 && max_packages < num) {
          n <- ceiling(length(parse_results$all_score_tables) / max_packages)
          start <- (package_no - 1) * n + 1
          end   <- min(package_no * n, num)
          parse_results$score_tables <- parse_results$all_score_tables[start:end]
        } else {
          cat("[错误] 未提取到任何评分标准表格\n")
          parse_results$score_tables <- list()
        }

        # 8：提取评审条款----
        incProgress(0.90, detail = "提取评审条款...")
        # 只从相关的章节中查找
        content <- try(parse_results$chapters |> 
          # 不包含投标文件格式章节
          dplyr::filter(!str_detect(title, config$bid_format_pattern)) |> 
          # 不包含合同模板章节
          dplyr::filter(!str_detect(title, config$contract_pattern)) |> 
          dplyr::pull(content) |> 
          stringr::str_c(collapse = "\n"), silent = TRUE
        )
        content <- ifelse(is.na(content), text, content)
        audit_terms <- fun_extract_all_audit_terms(content, bid_info, audit_tables, score_tables, input$package_no)
        parse_results$audit_terms <- audit_terms
        
        # 9：提取核心参数----
        incProgress(0.95, detail = "提取核心参数...")
        # 只从相关的章节中查找
        content <- try(parse_results$chapters |> 
          dplyr::filter(str_detect(title, config$procurement_pattern)) |> 
          dplyr::pull(content) |> 
          stringr::str_c(collapse = "\n"), silent = TRUE
        )
        content <- ifelse(is.na(content), text, content)
        core_parameters <- fun_extract_sentences(content, keywords = config$core_para_keywords)
        parse_results$core_parameters <- core_parameters

        # 完成
        incProgress(1, detail = "完成！")

        output$parse_status <- renderUI(
          tags$div(
            class = "alert alert-success",
            style = "padding: 10px; margin-top: 10px;",
            tags$strong("解析状态："),
            "解析完成！",
            # if (save_result) {
            #   tags$span(
            #     style = "display: block; margin-top: 5px; color: #28a745; font-size: 13px;",
            #     "✓ 解析文本已自动保存到 data 文件夹"
            #   )
            # } else {
            #   tags$span(
            #     style = "display: block; margin-top: 5px; color: #dc3545; font-size: 13px;",
            #     "✗ 保存解析文本失败"
            #   )
            # }
          )
        )
        on.exit({
          # 清除大对象，释放内存
          for (obj in c("text", "bid_info", "result", "tabs", "st", "audit_data", "wb")) {
            if (exists(obj, envir = environment())) rm(list = obj, envir = environment())
          }
          gc()  # 强制垃圾回收
        })
      }, error = function(e) {
        # 确保进度条能够完成并显示错误
        incProgress(1, detail = paste("错误:", e$message))
        output$parse_status <- renderUI(
          tags$div(
            class = "alert alert-danger",
            style = "padding: 10px; margin-top: 10px;",
            tags$strong("解析失败："),
            tags$p(style = "color: red; margin-top: 5px;", e$message),
            tags$p(style = "color: #666; font-size: 12px; margin-top: 10px;", "请检查文件格式是否正确，或联系技术支持")
          )
        )
        return(invisible(NULL))
      })
    })
  }, ignoreInit = TRUE)
  
  # ---- 4. 渲染总包数显示 ----------------
  output$total_packages_display <- renderText({
    as.character(app_state$total_packages)
  })

  # ---- 5. 渲染文件信息表格 ----------------
  output$file_basic_info <- renderTable({
    if (is.null(parse_results$file_info) || nrow(parse_results$file_info) == 0) {
      return(data.frame(提示 = "请上传文件并解析！"))
    }
    parse_results$file_info
  }, rownames = FALSE, striped = TRUE, hover = TRUE, bordered = TRUE)
  output$tender_core_info <- renderTable({
    if (is.null(parse_results$tender_info) || nrow(parse_results$tender_info) == 0) {
      return(data.frame(提示 = "请上传文件并解析！"))
    }
    parse_results$tender_info
  }, rownames = FALSE, striped = TRUE, hover = TRUE, bordered = TRUE)

  # ---- 5. 核心参数 ----------------
  output$core_parameters <- renderUI({
    if (is.null(parse_results$core_parameters) ||
        length(parse_results$core_parameters) == 0)
      return(tags$p("请上传文件并解析！", style = "text-align: center; color: #666; font-size: 16px; padding: 20px;"))

    rows <- lapply(seq_along(parse_results$core_parameters), function(i) {
      term <- fun_bold(
        parse_results$core_parameters[[i]],
        keywords = config$core_para_keywords,
        color = "red",
        bold = TRUE
      )
      tags$tr(
        tags$td(style = "text-align:center;font-weight:bold;", i),
        tags$td(style = "text-align:left;", HTML(term))
      )
    })
    tags$table(class = "table table-striped table-hover", tags$thead(tags$tr(tags$th("序号"), tags$th("核心参数"))), tags$tbody(rows))
  })
  
  # ---- 6. 评分标准（支持多表格显示） ----------------
  output$score_items_table <- renderUI({
    # 添加对package_no的依赖，确保包号变化时重新渲染
    req(input$package_no)

    package_no <- as.integer(input$package_no)
    all_score_tables <- parse_results$all_score_tables

    # 获取总包数
    max_packages <- if (!is.null(parse_results$tender_info) && nrow(parse_results$tender_info) > 0) {
      max_pkg_row <- parse_results$tender_info[parse_results$tender_info$`信息类型` == "总包数", ]
      if (nrow(max_pkg_row) > 0) {
        as.numeric(max_pkg_row$`提取结果`)
      } else {
        1
      }
    } else {
      1
    }

    # 当总包数为1时，显示所有表格
    if (max_packages == 1 && !is.null(all_score_tables) && length(all_score_tables) > 0) {
      # 多个表格：创建一个垂直布局显示所有表格
      tables_html <- list()
      tables_html[[length(tables_html) + 1]] <- tags$div(
        class = "alert alert-info",
        style = "padding: 10px; margin-bottom: 15px;",
        tags$strong("评分标准表格列表（共", length(all_score_tables), "个表格）："),
        tags$span("本项目仅1个包，显示所有评分标准")
      )

      # 显示每个表格
      for (i in seq_along(all_score_tables)) {
        tbl <- all_score_tables[[i]]
        # 为当前表格生成HTML行
        rows <- lapply(1:nrow(tbl), function(r) {
          tags$tr(lapply(tbl[r, ], function(x) {
            tags$td(style = "padding: 5px;", as.character(x))
          }))
        })
        tables_html[[length(tables_html) + 1]] <- tags$div(
          style = "margin-bottom: 15px; border: 1px solid #ddd; padding: 15px; border-radius: 5px;",
          tags$h4(paste("表格", i, "（", nrow(tbl), "行 x", ncol(tbl), "列）"),
                  style = "margin-bottom: 10px; color: #337ab7;"),
          tags$table(class = "table table-striped table-hover", style = "width: 100%;",
            tags$thead(
              tags$tr(lapply(colnames(tbl), function(col) tags$th(style = "padding: 5px;", col)))
            ),
            tags$tbody(rows)
          )
        )
      }
      return(tags$div(tables_html))
    }

    # 当总包数大于1时，显示对应包号的表格
    if (!is.null(all_score_tables) && length(all_score_tables) > 0 && package_no <= length(all_score_tables)) {
      # 如果有多个表格且当前包号有效，显示对应包号的表格
      tbl <- all_score_tables[[package_no]]
      if (!is.null(tbl) && nrow(tbl) > 0) {
        # 生成HTML表格
        rows <- lapply(1:nrow(tbl), function(i) {
          tags$tr(lapply(tbl[i, ], function(x) {
            tags$td(style = "padding: 5px;", as.character(x))
          }))
        })
        return(
          tags$div(
            tags$div(
              class = "alert alert-info",
              style = "padding: 10px; margin-bottom: 15px;",
              tags$strong(paste("包", package_no, "的评分标准："))
            ),
            tags$table(class = "table table-striped table-hover", style = "width: 100%;",
              tags$thead(
                tags$tr(lapply(colnames(tbl), function(col) tags$th(style = "padding: 5px;", col)))
              ),
              tags$tbody(rows)
            )
          )
        )
      }
    }

    return(tags$p("请上传文件并解析！", style = "text-align: center; color: #666; font-size: 16px; padding: 20px;"))
  })
  
  # ---- 7. 审核条款 ----------------
  output$audit_terms <- renderUI({
    # 添加对package_no的依赖，确保包号变化时重新渲染
    req(input$package_no)
    
    # 获取当前选择的包号
    package_no <- as.integer(input$package_no)
    
    # 使用parse_results中的评审条款
    audit_terms <- parse_results$audit_terms
    
    # 定义高亮关键词
    kw <- c(
      "无效",
      "废标",
      "资格",
      "符合",
      "评审标准",
      "项目名称",
      "项目编号",
      "采购人",
      "招标代理机构",
      "采购内容",
      "采购预算/限价",
      "项目属性",
      "投标保证金",
      "合同履行期限",
      "开标时间"
    )
    
    if (!is.null(audit_terms) && length(audit_terms) > 0) {
      # 生成表格行
      rows <- lapply(seq_along(audit_terms), function(i) {
        term <- fun_bold(
          audit_terms[[i]],
          keywords = kw,
          color = "red",
          bold = TRUE
        )
        tags$tr(
          tags$td(style = "text-align:center;font-weight:bold;", i),
          tags$td(style = "text-align:left;", HTML(term))
        )
      })
      
      # 返回包含包号标识的完整HTML内容
      return(tags$div(
        tags$div(
          class = "alert alert-info",
          style = "padding: 10px; margin-bottom: 15px;",
          tags$strong(paste("包", package_no, "的评审条款："))
        ),
        tags$table(class = "table table-striped table-hover", 
          tags$thead(tags$tr(tags$th("序号"), tags$th("评审条款"))), 
          tags$tbody(rows)
        )
      ))
    }

    return(tags$p("请上传文件并解析！", style = "text-align: center; color: #666; font-size: 16px; padding: 20px;"))
  })
  
  # ---- 8. 结构化评分规则（预览） ----------------
  output$structured_score_rules <- renderUI({
    df <- parse_results$score_tables
    if (is.null(df) || nrow(df) == 0)
      return(tags$p("未提取到评分项目"))
    tags$div(tags$h5("结构化评分规则（预览）"), tags$p("评分条目共 ", nrow(df), " 项"))
  })
  
  # ---- 9. 下载评审依据 ----------------
  output$download_audit <- downloadHandler(
    filename = function() {
      # 检查是否有解析结果
      if (is.null(parse_results$audit_terms) || length(parse_results$audit_terms) == 0) {
        # 返回一个虚拟文件名，实际不会下载
        return("请先上传文件并解析.txt")
      }

      if (!is.null(parse_results$tender_info) && !is.na(parse_results$tender_info[1, 2])) {
        tmp <- parse_results$tender_info[1, 2]
        nm <- stringr::str_replace_all(tmp, "[^\\w\\u4e00-\\u9fa5]", "_")
      } else {
        nm = NA_character_
        cat("[注意] 未提取到项目名称！")
      }
      paste0(format(Sys.time(), "%Y%m%d%H%M%S"), nm, "_评审内容", ".xlsx")
    },
    content = function(file) {
      # 检查是否有解析结果
      if (is.null(parse_results$audit_terms) || length(parse_results$audit_terms) == 0) {
        # 创建一个包含提示信息的文本文件
        writeLines("请先上传文件并解析！", file)
        return()
      }
      audit_data <- data.frame(
        `序号`   = seq_along(parse_results$audit_terms),
        `评审条款` = parse_results$audit_terms,
        `关键词` = sapply(parse_results$audit_terms, function(x) {
          if (grepl("无效", x))
            return("无效")
          if (grepl("废标", x))
            return("废标")
          "其他"
        }),
        `评审状态` = "",
        check.names = FALSE
      )
      wb <- createWorkbook()
      addWorksheet(wb, "评审条目")
      writeData(wb,
                "评审条目",
                audit_data,
                headerStyle = createStyle(textDecoration = "bold"))
      setColWidths(wb,
                   "评审条目",
                   cols = 1:4,
                   widths = c(8, 62, 10, 20))
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )

  # ---- 10. 下载投标文件格式（Word） ----------------
  output$download_bid_format <- downloadHandler(
    filename = function() {
      # 检查是否有解析结果
      if (is.null(parse_results$chapters) || nrow(parse_results$chapters) == 0) {
        # 返回一个虚拟文件名，实际不会下载
        return("请先上传文件并解析.txt")
      }

      if (!is.null(parse_results$tender_info) &&
          !is.na(parse_results$tender_info[1, 2])) {
        tmp <- parse_results$tender_info[1, 2]
        nm <- stringr::str_replace_all(tmp, "[^\\w\\u4e00-\\u9fa5]", "_")
      } else {
        nm = NA_character_
      }
      paste0(format(Sys.time(), "%Y%m%d%H%M%S"), nm, "_投标文件格式", ".docx")
    },
    content = function(file) {
      # 检查是否有解析结果
      if (is.null(parse_results$chapters) || nrow(parse_results$chapters) == 0) {
        # 创建一个包含提示信息的文本文件
        writeLines("请先上传文件并解析！", file)
        return()
      }

      tryCatch({
        # 详细错误日志
        cat("[下载] 开始生成Word文档...\n")
        
        # 提取投标文件格式相关章节
        cat("[下载] 提取投标文件格式章节...\n")
        
        # 提取投标文件格式相关章节
        bid_format_chapter <- fun_extract_chapter(
          chapters = parse_results$chapters,
          pattern = config$bid_format_pattern,
          return_mode = "last"
        )
        
        if (is.null(bid_format_chapter) ||
            nrow(bid_format_chapter) == 0 ||
            is.na(bid_format_chapter$title[1]) ||
            is.na(bid_format_chapter$content[1])) {
          stop("未找到投标文件格式相关章节，请确保文档中包含'投标文件格式'、'投标格式'等章节")
        }
        
        cat("[下载] 找到章节内容，长度:", nchar(bid_format_chapter$content[1]), "\n")
        
        # 获取模板文件路径
        template_path <- config$template_docx
        cat("[下载] 模板文件路径:", template_path, "\n")
        
        # 检查模板文件是否存在
        if (!file.exists(template_path)) {
          # 如果模板文件不存在，创建一个空的文档
          cat("[下载] 警告: 模板文件不存在，创建空文档\n")
          doc <- officer::read_docx()
        } else {
          # 使用模板文件
          cat("[下载] 使用模板文件\n")
          doc <- officer::read_docx(template_path)
        }
        
        # 按换行切分内容
        content_lines <- strsplit(bid_format_chapter$content[1], "\n")[[1]]
        
        # 逐行处理，根据#个数确定标题级别
        for (line_idx in seq_along(content_lines)) {
          line <- content_lines[line_idx]
          cat(sprintf("[下载] 处理第%d行: %s\n", line_idx, substr(line, 1, 50)))
          
          # 跳过空行
          if (trimws(line) == "") {
            doc <- officer::body_add_par(doc, value = "", style = "Normal")
            next
          }
          
          # 计算行首#的个数
          hash_match <- stringr::str_extract(line, "^#+")
          hash_count <- ifelse(is.na(hash_match), 0, nchar(hash_match))
          
          # 确定样式
          if (hash_count > 0) {
            # 有#，根据#个数确定标题级别
            level <- min(hash_count, 3)  # 限制最多3级标题，避免样式不存在
            content_text <- stringr::str_remove(line, "^#+") |> 
              stringr::str_trim()
            
            if (nchar(content_text) > 0) {
              style_name <- paste0("heading ", level)
              cat(sprintf("[下载] 添加%d级标题: %s\n", level, substr(content_text, 1, 30)))
              
              # 安全地添加标题
              tryCatch({
                doc <- officer::body_add_par(doc, value = content_text, style = style_name)
              }, error = function(e) {
                # 如果指定样式不存在，使用默认样式
                cat(sprintf("[下载] 样式%s不存在，使用Normal样式\n", style_name))
                doc <- officer::body_add_par(doc, value = content_text, style = "Normal")
              })
            }
          } else {
            # 无#，作为正文段落
            # 使用标准样式名称，避免中文样式名
            cat(sprintf("[下载] 添加正文: %s\n", substr(line, 1, 30)))
            tryCatch({
              doc <- officer::body_add_par(doc, value = line, style = "Normal")
            }, error = function(e) {
              # 如果Normal样式也不存在，使用空样式
              cat("[下载] Normal样式不存在，使用默认样式\n")
              doc <- officer::body_add_par(doc, value = line)
            })
          }
        }
        
        # 添加项目基本信息
        if (!is.null(parse_results$tender_info) && 
            nrow(parse_results$tender_info) > 0) {
          cat("[下载] 添加项目基本信息\n")
          
          doc <- officer::body_add_par(doc, value = "", style = "Normal")
          doc <- officer::body_add_par(doc, value = "项目基本信息", style = "heading 2")
          
          tender_text <- paste(apply(parse_results$tender_info, 1, function(row) {
            paste0(row[1], "：", row[2])
          }), collapse = "\n")
          
          doc <- officer::body_add_par(doc, value = tender_text, style = "Normal")
        }
        
        # 保存文档
        cat("[下载] 开始保存文档...\n")
        temp_file <- tempfile(fileext = ".docx")
        print(doc, target = temp_file)
        
        # 检查文件是否成功生成
        if (file.exists(temp_file) && file.size(temp_file) > 0) {
          # 复制到目标文件
          file.copy(temp_file, file)
          cat("[下载] 文档保存成功，文件大小:", file.size(file), "字节\n")
          # 清理临时文件
          unlink(temp_file)
        } else {
          # 使用友好的错误处理，避免直接抛出错误
          cat("[下载错误] 文档生成失败，输出文件为空或不存在\n")
          # 创建一个简单的错误文档
          error_doc <- officer::read_docx()
          error_doc <- officer::body_add_par(error_doc, "文档生成失败", style = "heading 1")
          error_doc <- officer::body_add_par(error_doc, "输出文件为空或不存在，请重试。", style = "Normal")
          print(error_doc, target = file)
          return()
        }
        
        cat("[下载] Word文档生成完成\n")
        
      }, error = function(e) {
        # 详细的错误处理，但使用友好的方式
        error_msg <- paste("生成Word文档失败:", e$message)
        cat("[下载错误]", error_msg, "\n")
        cat("[下载错误] 调用堆栈:\n")
        print(sys.calls())
        
        # 创建一个包含错误信息的文档，而不是抛出错误
        tryCatch({
          error_doc <- officer::read_docx()
          error_doc <- officer::body_add_par(error_doc, "文档生成过程中遇到问题", style = "heading 1")
          error_doc <- officer::body_add_par(error_doc, "请检查输入文档并重试。", style = "Normal")
          error_doc <- officer::body_add_par(error_doc, paste0("错误信息: ", substr(e$message, 1, 200)), style = "Normal")
          print(error_doc, target = file)
        }, error = function(inner_e) {
          cat("[下载错误] 甚至创建错误文档也失败了:", inner_e$message, "\n")
        })
      })
    }
  )
  
  # ---- 11. 原始文本预览 ----------------
  output$raw_text_preview <- renderText({
    if (is.null(parse_results$raw_text) || is.null(parse_results$chapters) ||
        nrow(parse_results$chapters) == 0)
      return("暂无解析文本（请先上传并解析文档）")

    # 检查chapters是否为有效的data.frame
    if (!is.data.frame(parse_results$chapters)) {
      return("暂无解析文本（请先上传并解析文档）")
    }

    txt <- parse_results$chapters |>
      dplyr::mutate(
        preview = stringr::str_sub(content, 1, 400),
        # 添加章节分隔线和更清晰的格式
        info    = paste0(
          "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n",
          "📄 章节：", title, "\n",
          "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n",
          preview, "\n\n"
        )
      ) |>
      dplyr::pull(info) |> paste(collapse = "")
    if (nchar(txt) > 5000)
      txt <- paste0(substr(txt, 1, 5000), "\n\n...（文本过长，已截取前5000字符）")
    txt
  })
  
  # ---- 12. 重新上传时清空 ----------------
  observeEvent(input$upload_file, {
    parse_results$file_info        <- NULL
    parse_results$tender_info      <- NULL
    parse_results$score_tables     <- NULL
    parse_results$all_score_tables <- NULL
    parse_results$audit_terms      <- NULL
    parse_results$raw_text         <- NULL
    parse_results$core_parameters  <- NULL
    parse_results$structured_rules <- NULL
    parse_results$chapters         <- NULL
  })
  
  # ---- 13. 实时更新包号参数 ----------------
  # 监听包号输入变化，触发内容更新
  observeEvent(input$package_no, {
    package_no <- as.integer(input$package_no)

    # 当包号变化时，记录日志
    cat("[包号更新] 当前选择包号:", package_no, "\n")

    # 使用isolate避免循环触发
    isolate({
      # 如果已有解析结果，根据包号更新显示的内容
      if (!is.null(parse_results$all_score_tables) && length(parse_results$all_score_tables) > 0) {
        if (package_no <= length(parse_results$all_score_tables)) {
          parse_results$score_tables <- parse_results$all_score_tables[[package_no]]
        } else {
          parse_results$score_tables <- parse_results$all_score_tables[[1]]
        }
      }

      # 重新计算当前包号对应的评审条款
      if (!is.null(parse_results$chapters) && nrow(parse_results$chapters) > 0) {
        tryCatch({
          cat("[包号更新] 重新计算评审条款，包号:", package_no, "\n")

          # 提取相关章节内容（排除投标文件格式和合同章节）
          content <- parse_results$chapters |>
            # 不包含投标文件格式章节
            dplyr::filter(!stringr::str_detect(title, config$bid_format_pattern)) |>
            # 不包含合同模板章节
            dplyr::filter(!stringr::str_detect(title, config$contract_pattern)) |>
            dplyr::pull(content) |>
            stringr::str_c(collapse = "\n")

          # 如果提取的内容为空，使用原始文本
          if (is.na(content) || content == "") {
            content <- parse_results$raw_text
          }

          # 获取当前包号对应的score_tables
          current_score_tables <- if (!is.null(parse_results$all_score_tables) &&
                                      length(parse_results$all_score_tables) > 0 &&
                                      package_no <= length(parse_results$all_score_tables)) {
            list(parse_results$all_score_tables[[package_no]])
          } else {
            list()
          }

          # 重新生成审计条款
          if (length(current_score_tables) > 0) {
            new_audit_terms <- fun_extract_all_audit_terms(
              content,
              parse_results$tender_info,
              list(),  # audit_tables为空，因为主要从文本中提取
              current_score_tables,
              package_no
            )
            parse_results$audit_terms <- new_audit_terms
            cat("[包号更新] 评审条款更新完成，共", length(new_audit_terms), "条\n")
          }
        }, error = function(e) {
          cat("[包号更新] 重新计算评审条款失败:", e$message, "\n")
        })
      }
    })
  })
  
  # ---- 14. 监听总包数变化，更新包号可选范围 ----------------
  # 当app_state中的总包数变化时，更新投标包号的选项范围
  observeEvent(app_state$total_packages, {
    req(app_state$total_packages)

    current_max <- as.integer(app_state$total_packages)
    current_package <- as.integer(input$package_no)

    # 验证总包数有效性
    if (!is.na(current_max) && current_max >= 1) {
      # 更新包号的选项范围
      updateSelectInput(session, "package_no", choices = 1:current_max)

      # 确保当前选择的包号在新范围内
      if (!is.na(current_package) && current_package > current_max) {
        updateSelectInput(session, "package_no", selected = current_max)
      }

      cat("[总包数更新] 当前总包数:", current_max, "\n")
    }
  }, ignoreInit = FALSE)

  # ---- 15. 基于解析结果更新总包数 ----------------
  # 当解析完成后，根据bid_info中的max_packages更新总包数
  observe({
    req(parse_results$tender_info)

    # 检查是否从文件中提取到了总包数信息
    bid_info <- parse_results$tender_info
    if (!is.null(bid_info) && nrow(bid_info) > 0) {
      # 查找总包数信息行
      max_pkg_row <- bid_info[bid_info$`信息类型` == "总包数", ]

      if (nrow(max_pkg_row) > 0) {
        extracted_max <- as.numeric(max_pkg_row$`提取结果`)

        if (!is.na(extracted_max) && extracted_max >= 1) {
          # 只有当提取到的总包数有效时才更新
          if (extracted_max != app_state$total_packages) {
            # 更新总包数（显示为只读）
            app_state$total_packages <- extracted_max
            # 更新package_no的可选范围为1到提取到的总包数之间
            updateSelectInput(session, "package_no", choices = 1:extracted_max, selected = 1)

            cat("[总包数提取] 从文档中提取到总包数:", extracted_max, "\n")
          }
        }
      }
    }
  })

  # ---- 17. 重置参数到默认值 ----------------
  # 添加一个重置按钮（可选，可以在UI中添加）
  observeEvent(input$upload_file, {
    # 当上传新文件时，重置包号和总包数为默认值
    app_state$total_packages <- 1  # 重置总包数为1
    updateSelectInput(session, "package_no", choices = 1, selected = 1)
  })
}

# 启动Web应用
# ==== 配置Shiny应用参数 =====
# 设置全局选项防止断开
# options(shiny.autoreload = FALSE)
options(shiny.reactlog = FALSE)
options(shiny.suppressMissingContextError = TRUE)
# 增加会话超时到10分钟，避免长时间PDF处理导致断开
options(shiny.maxRequestSize = 100 * 1024^2)  # 100MB最大文件
options(shiny.timeout = 600)  # 10分钟超时

config <- get_bid_config()
# 在shinyApp调用前检查
check_dependencies()
# 启动应用
shinyApp(ui = ui, server = server)
