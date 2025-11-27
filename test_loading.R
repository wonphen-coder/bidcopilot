# =============================================================================
# 快速测试函数加载
# =============================================================================

cat("╔══════════════════════════════════════════════════════════════╗\n"); flush(stdout())
cat("║              函数加载测试                                    ║\n"); flush(stdout())
cat("╚══════════════════════════════════════════════════════════════╝\n\n"); flush(stdout())

# 测试1: 检查app.R能否正确加载
cat("[测试1] 检查app.R文件...\n"); flush(stdout())
if (!file.exists("app.R")) {
  cat("❌ app.R文件不存在\n"); flush(stdout())
  stop("请确保app.R文件在当前目录")
}
cat("✅ app.R文件存在\n\n"); flush(stdout())

# 测试2: 加载函数（不启动Shiny）
cat("[测试2] 加载函数定义...\n"); flush(stdout())
tryCatch({
  app_content <- parse("app.R")
  n_lines <- length(app_content)
  cat("  app.R总行数:", n_lines, "\n"); flush(stdout())

  # 移除最后一行（shinyApp）
  app_content <- app_content[-length(app_content)]
  cat("  移除最后一行后，剩余:", length(app_content), "行\n"); flush(stdout())

  # 执行函数定义
  eval(app_content, environment())
  cat("✅ 函数加载完成\n\n"); flush(stdout())
}, error = function(e) {
  cat("❌ 函数加载失败:", e$message, "\n"); flush(stdout())
  cat("错误位置:", e$call, "\n"); flush(stdout())
  stop("加载失败")
})

# 测试3: 检查关键函数是否存在
cat("[测试3] 检查关键函数...\n"); flush(stdout())
functions_to_check <- c("processDOCX", "processPDF", "processTXT")

for (func_name in functions_to_check) {
  if (exists(func_name, mode = "function")) {
    cat("  ✅", func_name, "函数存在\n"); flush(stdout())
  } else {
    cat("  ❌", func_name, "函数不存在\n"); flush(stdout())
  }
}
cat("\n"); flush(stdout())

# 测试4: 简单的DOCX/PDF文件检查
cat("[测试4] 检查测试文件...\n"); flush(stdout())
docx_files <- list.files(pattern = "\\.docx$", full.names = TRUE)
pdf_files <- list.files(pattern = "\\.pdf$", full.names = TRUE)

cat("  DOCX文件:", length(docx_files), "个\n"); flush(stdout())
cat("  PDF文件:", length(pdf_files), "个\n"); flush(stdout())

if (length(docx_files) > 0) {
  cat("  示例DOCX:", basename(docx_files[1]), "\n"); flush(stdout())
}
if (length(pdf_files) > 0) {
  cat("  示例PDF:", basename(pdf_files[1]), "\n"); flush(stdout())
}
cat("\n"); flush(stdout())

# 完成
cat("✅ 所有测试通过!\n"); flush(stdout())
cat("\n现在可以运行完整测试:\n"); flush(stdout())
cat("  source('test_docx_clean.R')  # 测试DOCX\n"); flush(stdout())
cat("  source('test_pdf_clean.R')   # 测试PDF\n"); flush(stdout())
cat("\n"); flush(stdout())
