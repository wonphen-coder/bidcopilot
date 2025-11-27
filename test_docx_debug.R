# =============================================================================
# DOCX调试测试脚本
# 目的：独立测试DOCX处理函数，定位报错问题
# =============================================================================

cat("╔══════════════════════════════════════════════════════════════╗\n"); flush(stdout())
cat("║              DOCX处理调试测试脚本                             ║\n"); flush(stdout())
cat("╚══════════════════════════════════════════════════════════════╝\n\n"); flush(stdout())

# 1. 加载必要的包
cat("[步骤1] 加载包...\n"); flush(stdout())
required_packages <- c("officer", "stringr")

for (pkg in required_packages) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    cat("❌ 包", pkg, "未安装，正在安装...\n"); flush(stdout())
    install.packages(pkg, repos = "https://mirrors.tuna.tsinghua.edu.cn/CRAN/")
  }
}

library(officer)
library(stringr)
cat("✅ 包加载完成\n\n"); flush(stdout())

# 2. 内存状态检查
cat("[步骤2] 初始内存状态:\n"); flush(stdout())
print(gc())
cat("\n"); flush(stdout())

# 3. 列出测试目录中的DOCX文件
cat("[步骤3] 查找DOCX文件...\n"); flush(stdout()); flush(stdout())
docx_files <- list.files(pattern = "\\.docx$", full.names = TRUE)

if (length(docx_files) == 0) {
  cat("❌ 未找到DOCX文件\n"); flush(stdout()); flush(stdout())
  cat("请将DOCX文件放在当前目录下再运行此脚本\n"); flush(stdout()); flush(stdout())
  stop("没有可测试的DOCX文件")
}

cat("✅ 找到", length(docx_files)); flush(stdout()); cat("个DOCX文件:\n"); flush(stdout())
for (i in seq_along(docx_files)) {
  cat("  ", i, ". ", docx_files[i], "\n"); flush(stdout()); flush(stdout())
}
cat("\n"); flush(stdout()); flush(stdout())

# 4. 逐个测试DOCX文件
for (i in seq_along(docx_files)) {
  docx_file <- docx_files[i]
  cat("\n"); flush(stdout())
  cat("╔══════════════════════════════════════════════════════════════╗\n"); flush(stdout())
  cat("║ 测试文件", i, ":", basename(docx_file; flush(stdout()), "\n"))
  cat("╚══════════════════════════════════════════════════════════════╝\n"); flush(stdout())

  file_size <- file.size(docx_file)
  cat("[文件信息] 大小:", round(file_size/1024/1024, 2; flush(stdout()), "MB\n"))
  cat("[文件信息] 修改时间:", file.info(docx_file; flush(stdout())$mtime, "\n\n"))

  # 检查文件是否可读
  cat("[测试] 检查文件是否可读...\n"; flush(stdout()))
  if (!file.exists(docx_file)) {
    cat("❌ 文件不存在\n\n"; flush(stdout()))
    next
  }
  cat("✅ 文件存在\n\n"; flush(stdout()))

  # 开始DOCX处理测试
  cat("[测试] 开始调用processDOCX(; flush(stdout())...\n"); flush(stdout())
  cat("注意: 这可能需要几十秒时间，请耐心等待...\n\n"; flush(stdout()); flush(stdout()))

  start_time <- Sys.time()
  cat("[开始时间]", start_time, "\n"; flush(stdout()); flush(stdout()))

  result <- tryCatch({
    # 分步骤执行
    cat("\n[执行] 步骤1/3 - 调用officer::read_docx(; flush(stdout())...\n"); flush(stdout())
    doc <- officer::read_docx(docx_file)

    cat("\n[执行] 步骤2/3 - 调用officer::docx_summary(; flush(stdout())...\n"); flush(stdout())
    df <- officer::docx_summary(doc)

    cat("\n[执行] 步骤3/3 - 处理文档内容...\n"); flush(stdout()); flush(stdout())
    # 简单模拟processDOCX的逻辑
    output_lines <- character()
    for (j in 1:nrow(df)) {
      row <- df[j, ]
      if (row$content_type == "paragraph" && !is.na(row$text) && trimws(row$text) != "") {
        if (!stringr::str_detect(row$text, "PAGEREF_Toc")) {
          output_lines <- c(output_lines, row$text)
        }
      }
    }

    cat("返回行数:", nrow(df; flush(stdout()), "\n")); flush(stdout())
    cat("处理段落数:", length(output_lines; flush(stdout()), "\n")); flush(stdout())
    cat("提取文本长度:", nchar(paste(output_lines, collapse = "\n"); flush(stdout())), "字符\n"); flush(stdout())

    return(paste(output_lines, collapse = "\n"))

  }, error = function(e) {
    cat("\n❌ 错误发生!\n"); flush(stdout()); flush(stdout())
    cat("[错误类型]", class(e)); flush(stdout()), "\n"); flush(stdout())
    cat("[错误信息]", e$message, "\n"); flush(stdout()); flush(stdout())
    cat("[错误时间]", Sys.time(; flush(stdout()), "\n")); flush(stdout())
    cat("\n[调用栈]\n"); flush(stdout()); flush(stdout())
    print(sys.calls())
    cat("\n[内存状态]\n"); flush(stdout()); flush(stdout())
    print(gc())
    return(NULL)
  })

  end_time <- Sys.time()
  cat("\n[结束时间]", end_time, "\n"; flush(stdout()); flush(stdout())
  cat("[耗时]", round(difftime(end_time, start_time, units = "secs"; flush(stdout()), 2), "秒\n"); flush(stdout())

  if (!is.null(result)) {
    cat("✅ DOCX处理成功!\n\n"; flush(stdout()); flush(stdout())
    cat("前200字符预览:\n"; flush(stdout()); flush(stdout())
    cat(substr(result, 1, 200; flush(stdout()), "...\n\n"); flush(stdout())
  } else {
    cat("❌ DOCX处理失败!\n\n"; flush(stdout()); flush(stdout())
  }

  cat("当前内存状态:\n"; flush(stdout()); flush(stdout())
  print(gc())
  cat("\n"); flush(stdout()); flush(stdout())
}

# 5. 总结报告
cat("\n"); flush(stdout())
cat("╔══════════════════════════════════════════════════════════════╗\n"; flush(stdout())
cat("║                      测试总结                                ║\n"; flush(stdout())
cat("╚══════════════════════════════════════════════════════════════╝\n"; flush(stdout())

cat("测试文件数:", length(docx_files; flush(stdout()), "\n")
cat("测试时间:", Sys.time(; flush(stdout()), "\n")
cat("最终内存状态:\n"; flush(stdout())
print(gc())

cat("\n"); flush(stdout())
cat("建议:\n"; flush(stdout())
cat("1. 如果所有DOCX都失败，检查officer包和依赖\n"; flush(stdout())
cat("2. 如果某些DOCX失败，可能是文件格式问题\n"; flush(stdout())
cat("3. 确保officer包版本 >= 0.6.0\n"; flush(stdout())
cat("4. 尝试用Word重新保存文件\n"; flush(stdout())
cat("\n"); flush(stdout())
