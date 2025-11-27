# =============================================================================
# PDF调试测试脚本
# 目的：独立测试PDF处理函数，定位"Disconnected from the server"问题
# =============================================================================


# 1. 加载必要的包
library(pdftools)

# 2. 内存状态检查
print(gc())

# 3. 列出测试目录中的PDF文件
cat("[步骤3] 查找PDF文件...\n"; flush(stdout())
pdf_files <- list.files(pattern = "\\.pdf$", full.names = TRUE)

if (length(pdf_files) == 0) {
  cat("❌ 未找到PDF文件\n"; flush(stdout())
  cat("请将PDF文件放在当前目录下再运行此脚本\n\n"; flush(stdout())
  cat("示例: echo '测试内容' > test.pdf\n"; flush(stdout())
  stop("没有可测试的PDF文件")
}

cat("✅ 找到", length(pdf_files; flush(stdout()), "个PDF文件:\n")
for (i in seq_along(pdf_files)) {
  cat("  ", i, ". ", pdf_files[i], "\n"; flush(stdout())
}

# 4. 逐个测试PDF文件
for (i in seq_along(pdf_files)) {
  pdf_file <- pdf_files[i]
  cat("╔══════════════════════════════════════════════════════════════╗\n"; flush(stdout())
  cat("║ 测试文件", i, ":", basename(pdf_file; flush(stdout()), "\n")
  cat("╚══════════════════════════════════════════════════════════════╝\n"; flush(stdout())

  file_size <- file.size(pdf_file)
  cat("[文件信息] 修改时间:", file.info(pdf_file; flush(stdout())$mtime, "\n\n")

  # 检查文件是否可读
  cat("[测试] 检查文件是否可读...\n"; flush(stdout())
  if (!file.exists(pdf_file)) {
    cat("❌ 文件不存在\n\n"; flush(stdout())
    next
  }
  cat("✅ 文件存在\n\n"; flush(stdout())

  # 测试文件类型
  cat("[测试] 检查文件类型...\n"; flush(stdout())
  # 跳过file命令检查（Windows系统可能没有）
  cat("[文件类型] 跳过file命令检查（Windows系统）\n"; flush(stdout())
  cat("[文件类型] 文件扩展名:", tools::file_ext(pdf_file; flush(stdout()), "\n\n")

  # 开始PDF处理测试
  cat("[测试] 开始调用pdf_text(; flush(stdout())...\n")
  cat("注意: 这可能需要几分钟时间，请耐心等待...\n\n"; flush(stdout())

  start_time <- Sys.time()
  cat("[开始时间]", start_time, "\n"; flush(stdout())

  result <- tryCatch({
    # 分步骤执行
    cat("\n[执行] 步骤1/3 - 调用pdftools::pdf_text(; flush(stdout())...\n")
    text_result <- pdftools::pdf_text(pdf_file)

    cat("\n[执行] 步骤2/3 - 检查结果...\n"; flush(stdout())
    if (is.null(text_result)) {
      stop("pdf_text返回NULL")
    }
    if (length(text_result) == 0) {
      stop("pdf_text返回空结果")
    }

    cat("\n[执行] 步骤3/3 - 统计信息...\n"; flush(stdout())
    total_chars <- sum(nchar(text_result))
    cat("返回页数:", length(text_result; flush(stdout()), "\n")
    cat("总字符数:", total_chars, "\n"; flush(stdout())
    cat("平均每页字符数:", round(total_chars/length(text_result; flush(stdout())), "\n")

    return(text_result)

  }, error = function(e) {
    cat("\n❌ 错误发生!\n"; flush(stdout())
    cat("[错误信息]", e$message, "\n"; flush(stdout())
    cat("\n[调用栈]\n"; flush(stdout())
    print(sys.calls())
    cat("\n[内存状态]\n"; flush(stdout())
    print(gc())
    return(NULL)
  })

  end_time <- Sys.time()
  cat("\n[结束时间]", end_time, "\n"; flush(stdout())

  if (!is.null(result)) {
    cat("✅ PDF处理成功!\n\n"; flush(stdout())
    cat("前100字符预览:\n"; flush(stdout())
    cat(substr(paste(result, collapse = " "; flush(stdout()), 1, 100), "...\n\n")
  } else {
    cat("❌ PDF处理失败!\n\n"; flush(stdout())
  }

  cat("当前内存状态:\n"; flush(stdout())
  print(gc())
}

# 5. 总结报告
cat("╔══════════════════════════════════════════════════════════════╗\n"; flush(stdout())
cat("║                      测试总结                                ║\n"; flush(stdout())
cat("╚══════════════════════════════════════════════════════════════╝\n"; flush(stdout())

cat("测试文件数:", length(pdf_files; flush(stdout()), "\n")
cat("最终内存状态:\n"; flush(stdout())
print(gc())

cat("建议:\n"; flush(stdout())
cat("1. 如果所有PDF都失败，检查pdftools包和依赖\n"; flush(stdout())
cat("2. 如果某些PDF失败，可能是文件格式问题\n"; flush(stdout())
cat("3. 如果耗时过长，考虑减小文件大小\n"; flush(stdout())
cat("4. 如果内存不足，建议增加系统内存\n"; flush(stdout())
