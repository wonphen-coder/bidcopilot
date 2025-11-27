# =============================================================================
# DOCX调试测试脚本
# 目的：独立测试DOCX处理函数，定位报错问题
# =============================================================================

cat("╔══════════════════════════════════════════════════════════════╗\n"); flush(stdout())
cat("║              DOCX处理调试测试脚本                             ║\n"); flush(stdout())
cat("╚══════════════════════════════════════════════════════════════╝\n\n"); flush(stdout())

# 1. 加载必要的包和函数
cat("[步骤1] 加载包...\n"); flush(stdout())
required_packages <- c("officer", "stringr", "pdftools", "dplyr", "tidyllm", "jsonlite", "openxlsx", "cli", "lubridate", "tools")

for (pkg in required_packages) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    cat("❌ 包", pkg, "未安装，正在安装...\n"); flush(stdout())
    install.packages(pkg, repos = "https://mirrors.tuna.tsinghua.edu.cn/CRAN/")
  }
}

library(officer)
library(stringr)
library(pdftools)
library(dplyr)
library(tidyllm)
library(jsonlite)
library(openxlsx)
library(cli)
library(lubridate)
library(tools)
cat("✅ 包加载完成\n\n"); flush(stdout())

# 2. 加载主应用函数（包含完整的processDOCX函数）
cat("[步骤2] 加载主应用函数...\n"); flush(stdout())
tryCatch({
  # 只加载函数定义，不启动Shiny应用
  # 使用parse和eval而不是source，避免执行shinyApp
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

# 4. 列出测试目录中的DOCX文件
cat("[步骤4] 查找DOCX文件...\n"); flush(stdout())
docx_files <- list.files(pattern = "\\.docx$", full.names = TRUE)

if (length(docx_files) == 0) {
  cat("❌ 未找到DOCX文件\n"); flush(stdout())
  cat("请将DOCX文件放在当前目录下再运行此脚本\n"); flush(stdout())
  stop("没有可测试的DOCX文件")
}

cat("✅ 找到", length(docx_files), "个DOCX文件:\n"); flush(stdout())
for (i in seq_along(docx_files)) {
  cat("  ", i, ". ", docx_files[i], "\n"); flush(stdout())
}
cat("\n"); flush(stdout())

# 5. 逐个测试DOCX文件
for (i in seq_along(docx_files)) {
  docx_file <- docx_files[i]
  cat("\n"); flush(stdout())
  cat("╔══════════════════════════════════════════════════════════════╗\n"); flush(stdout())
  cat("║ 测试文件", i, ":", basename(docx_file), "\n"); flush(stdout())
  cat("╚══════════════════════════════════════════════════════════════╝\n"); flush(stdout())

  file_size <- file.size(docx_file)
  cat("[文件信息] 大小:", round(file_size / 1024 / 1024, 2), "MB\n"); flush(stdout())
  cat("[文件信息] 修改时间:", file.info(docx_file)$mtime, "\n\n"); flush(stdout())

  # 检查文件是否可读
  cat("[测试] 检查文件是否可读...\n"); flush(stdout())
  if (!file.exists(docx_file)) {
    cat("❌ 文件不存在\n\n"); flush(stdout())
    next
  }
  cat("✅ 文件存在\n\n"); flush(stdout())

  # 开始DOCX处理测试 - 调用完整的processDOCX函数
  cat("[测试] 开始调用完整的processDOCX()函数...\n"); flush(stdout())
  cat("注意: 这可能需要几十秒时间，请耐心等待...\n"); flush(stdout())
  cat("将使用与主应用相同的完整处理逻辑，包括表格处理...\n\n"); flush(stdout())

  start_time <- Sys.time()
  cat("[开始时间]", start_time, "\n"); flush(stdout())

  result <- tryCatch({
    # 调用完整的processDOCX函数（而非简化的测试逻辑）
    cat("\n[执行] 调用完整processDOCX函数...\n"); flush(stdout())
    result <- processDOCX(docx_file)

    # 验证结果
    cat("\n[验证] 检查处理结果...\n"); flush(stdout())
    result_text <- as.character(result)
    if (is.null(result_text) || length(result_text) == 0 || all(nchar(result_text) == 0)) {
      stop("processDOCX返回NULL或空字符串")
    }

    # 统计信息
    cat("提取文本长度:", sum(nchar(result_text)), "字符\n"); flush(stdout())
    cat("文本行数:", length(strsplit(result_text, "\n")[[1]]), "行\n"); flush(stdout())
    cat("是否包含表格:", ifelse(any(grepl("\\|", result_text)), "是", "否"), "\n"); flush(stdout())

    return(result)

  }, error = function(e) {
    cat("\n❌ 错误发生!\n"); flush(stdout())
    cat("[错误类型]", class(e), "\n"); flush(stdout())
    cat("[错误信息]", e$message, "\n"); flush(stdout())
    cat("[错误位置] ", as.character(e$call[1]), "\n"); flush(stdout())
    cat("[错误时间]", Sys.time(), "\n"); flush(stdout())
    cat("\n[调用栈]\n"); flush(stdout())
    print(sys.calls())
    cat("\n[内存状态]\n"); flush(stdout())
    print(gc())
    cat("\n[建议] 检查文件格式和依赖包\n"); flush(stdout())
    return(NULL)
  })

  end_time <- Sys.time()
  cat("\n[结束时间]", end_time, "\n"); flush(stdout())
  cat("[耗时]", round(difftime(end_time, start_time, units = "secs"), 2), "秒\n"); flush(stdout())

  if (!is.null(result)) {
    result_text <- as.character(result)
    cat("\n✅ DOCX处理成功!\n\n"); flush(stdout())
    cat("[预览] 前300字符内容:\n"); flush(stdout())
    cat(substr(result_text, 1, 300), "...\n\n"); flush(stdout())

    # 检查是否包含表格
    if (any(grepl("\\|", result_text))) {
      cat("[检测] ✅ 表格数据已成功提取\n\n"); flush(stdout())
    } else {
      cat("[检测] ⚠️ 未检测到表格数据（可能文档中无表格）\n\n"); flush(stdout())
    }
  } else {
    cat("\n❌ DOCX处理失败!\n\n"); flush(stdout())
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

cat("测试文件数:", length(docx_files), "\n"); flush(stdout())
cat("测试时间:", Sys.time(), "\n"); flush(stdout())
cat("最终内存状态:\n"); flush(stdout())
print(gc())

cat("\n"); flush(stdout())
cat("建议:\n"); flush(stdout())
cat("1. 如果DOCX处理失败，检查officer包和依赖\n"); flush(stdout())
cat("2. 如果提取内容不完整，可能是文档格式复杂\n"); flush(stdout())
cat("3. 确保officer包版本 >= 0.6.0\n"); flush(stdout())
cat("4. 尝试用Word重新保存文件\n"); flush(stdout())
cat("5. 大文件处理需要更多时间，请耐心等待\n"); flush(stdout())
cat("\n"); flush(stdout())
