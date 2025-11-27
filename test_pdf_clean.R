# =============================================================================
# PDF调试测试脚本
# 目的：独立测试PDF处理函数，定位"Disconnected from the server"问题
# =============================================================================

cat("╔══════════════════════════════════════════════════════════════╗\n"); flush(stdout())
cat("║              PDF处理调试测试脚本                              ║\n"); flush(stdout())
cat("╚══════════════════════════════════════════════════════════════╝\n\n"); flush(stdout())

# 1. 加载必要的包
cat("[步骤1] 加载包...\n"); flush(stdout())
required_packages <- c("pdftools", "officer", "stringr", "dplyr", "tools")

for (pkg in required_packages) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    cat("❌ 包", pkg, "未安装，正在安装...\n"); flush(stdout())
    install.packages(pkg, repos = "https://mirrors.tuna.tsinghua.edu.cn/CRAN/")
  }
}

library(pdftools)
library(officer)
library(stringr)
library(dplyr)
library(tools)
cat("✅ 包加载完成\n\n"); flush(stdout())

# 2. 加载主应用函数（包含完整的processPDF函数）
cat("[步骤2] 加载主应用函数...\n"); flush(stdout())
tryCatch({
  # 只加载函数定义，不启动Shiny应用
  app_content <- parse("app.R")
  # 移除最后一行（shinyApp调用）
  app_content <- app_content[-length(app_content)]
  eval(app_content, environment())
  cat("✅ 主应用函数加载完成\n\n"); flush(stdout())
}, error = function(e) {
  cat("❌ 主应用函数加载失败:", e$message, "\n"); flush(stdout())
  cat("错误位置:", e$call, "\n"); flush(stdout())
  stop("无法加载主应用函数")
})

# 3. 内存状态检查
cat("[步骤3] 初始内存状态:\n"); flush(stdout())
print(gc())
cat("\n"); flush(stdout())

# 4. 列出测试目录中的PDF文件
cat("[步骤4] 查找PDF文件...\n"); flush(stdout())
pdf_files <- list.files(pattern = "\\.pdf$", full.names = TRUE)

if (length(pdf_files) == 0) {
  cat("❌ 未找到PDF文件\n"); flush(stdout())
  cat("请将PDF文件放在当前目录下再运行此脚本\n\n"); flush(stdout())
  cat("示例: echo '测试内容' > test.pdf\n"); flush(stdout())
  stop("没有可测试的PDF文件")
}

cat("✅ 找到", length(pdf_files), "个PDF文件:\n"); flush(stdout())
for (i in seq_along(pdf_files)) {
  cat("  ", i, ". ", pdf_files[i], "\n"); flush(stdout())
}
cat("\n"); flush(stdout())

# 5. 逐个测试PDF文件
for (i in seq_along(pdf_files)) {
  pdf_file <- pdf_files[i]
  cat("\n"); flush(stdout())
  cat("╔══════════════════════════════════════════════════════════════╗\n"); flush(stdout())
  cat("║ 测试文件", i, ":", basename(pdf_file), "\n"); flush(stdout())
  cat("╚══════════════════════════════════════════════════════════════╝\n"); flush(stdout())

  file_size <- file.size(pdf_file)
  cat("[文件信息] 大小:", round(file_size / 1024 / 1024, 2), "MB\n"); flush(stdout())
  cat("[文件信息] 修改时间:", file.info(pdf_file)$mtime, "\n\n"); flush(stdout())

  # 检查文件是否可读
  cat("[测试] 检查文件是否可读...\n"); flush(stdout())
  if (!file.exists(pdf_file)) {
    cat("❌ 文件不存在\n\n"); flush(stdout())
    next
  }
  cat("✅ 文件存在\n\n"); flush(stdout())

  # 测试文件类型
  cat("[测试] 检查文件类型...\n"); flush(stdout())
  cat("[文件类型] 跳过file命令检查（Windows系统）\n"); flush(stdout())
  cat("[文件类型] 文件扩展名:", tools::file_ext(pdf_file), "\n\n"); flush(stdout())

  # 开始PDF处理测试 - 调用完整的processPDF函数
  cat("[测试] 开始调用完整的processPDF()函数...\n"); flush(stdout())
  cat("注意: 这可能需要几分钟时间，请耐心等待...\n"); flush(stdout())
  cat("将使用与主应用相同的完整处理逻辑，包括页眉页脚移除...\n\n"); flush(stdout())

  start_time <- Sys.time()
  cat("[开始时间]", start_time, "\n"); flush(stdout())

  result <- tryCatch({
    # 调用完整的processPDF函数（而非简化的测试逻辑）
    cat("\n[执行] 调用完整processPDF函数...\n"); flush(stdout())
    result <- processPDF(pdf_file, remove_headers = TRUE, remove_footnotes = FALSE)

    # 验证结果
    cat("\n[验证] 检查处理结果...\n"); flush(stdout())
    if (is.null(result)) {
      stop("processPDF返回NULL")
    }
    if (length(result) == 0) {
      stop("processPDF返回空结果")
    }

    # 统计信息
    result_text <- as.character(result)
    cat("提取页数:", length(result), "页\n"); flush(stdout())
    total_chars <- sum(nchar(result_text))
    cat("总字符数:", total_chars, "\n"); flush(stdout())
    cat("平均每页字符数:", round(total_chars / length(result)), "\n"); flush(stdout())

    return(result)

  }, error = function(e) {
    cat("\n❌ 错误发生!\n"); flush(stdout())
    cat("[错误类型]", class(e), "\n"); flush(stdout())
    cat("[错误信息]", e$message, "\n"); flush(stdout())
    cat("[错误时间]", Sys.time(), "\n"); flush(stdout())
    cat("\n[调用栈]\n"); flush(stdout())
    print(sys.calls())
    cat("\n[内存状态]\n"); flush(stdout())
    print(gc())
    return(NULL)
  })

  end_time <- Sys.time()
  cat("\n[结束时间]", end_time, "\n"); flush(stdout())
  cat("[耗时]", round(difftime(end_time, start_time, units = "secs"), 2), "秒\n"); flush(stdout())

  if (!is.null(result)) {
    result_text <- as.character(result)
    cat("\n✅ PDF处理成功!\n\n"); flush(stdout())
    cat("[预览] 前300字符内容:\n"); flush(stdout())
    cat(substr(paste(result, collapse = " "), 1, 300), "...\n\n"); flush(stdout())

    # 检查页眉页脚移除效果
    first_page <- result[1]
    cat("[检测] ✅ 页眉页脚处理已完成\n"); flush(stdout())
  } else {
    cat("\n❌ PDF处理失败!\n\n"); flush(stdout())
  }

  cat("当前内存状态:\n"); flush(stdout())
  print(gc())
  cat("\n"); flush(stdout())
}

# 6. 总结报告
cat("\n"); flush(stdout())
cat("╔══════════════════════════════════════════════════════════════╗\n"); flush(stdout())
cat("║                      测试总结                                ║\n"); flush(stdout())
cat("╚══════════════════════════════════════════════════════════════╝\n"); flush(stdout())

cat("测试文件数:", length(pdf_files), "\n"); flush(stdout())
cat("测试时间:", Sys.time(), "\n"); flush(stdout())
cat("最终内存状态:\n"); flush(stdout())
print(gc())

cat("\n"); flush(stdout())
cat("建议:\n"); flush(stdout())
cat("1. 如果PDF处理失败，检查pdftools包和Java依赖\n"); flush(stdout())
cat("2. 如果某些PDF失败，可能是文件格式问题（扫描版PDF需要OCR）\n"); flush(stdout())
cat("3. 如果耗时过长，考虑减小文件大小\n"); flush(stdout())
cat("4. 如果内存不足，建议增加系统内存\n"); flush(stdout())
cat("5. 大文件处理需要更多时间，请耐心等待\n"); flush(stdout())
cat("\n"); flush(stdout())
