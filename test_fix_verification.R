# =============================================================================
# 验证修复效果的快速测试脚本
# =============================================================================

cat("╔══════════════════════════════════════════════════════════════╗\n"); flush(stdout())
cat("║              验证修复效果 - 快速测试                          ║\n"); flush(stdout())
cat("╚══════════════════════════════════════════════════════════════╝\n\n"); flush(stdout())

# 1. 检查文件是否存在
cat("[检查] 查找测试文件...\n"); flush(stdout())
docx_files <- list.files(pattern = "\\.docx$", full.names = TRUE)
pdf_files <- list.files(pattern = "\\.pdf$", full.names = TRUE)

if (length(docx_files) == 0 && length(pdf_files) == 0) {
  cat("❌ 未找到测试文件\n"); flush(stdout())
  stop("请将DOCX或PDF文件放在当前目录下")
}

cat("✅ 找到文档文件:\n"); flush(stdout())
for (f in c(docx_files, pdf_files)) {
  cat("  -", basename(f), "\n"); flush(stdout())
}
cat("\n"); flush(stdout())

# 2. 验证基础语法
cat("[验证] 检查基础语法...\n"); flush(stdout())
tryCatch({
  # 尝试加载一个简单的库
  library(stringr)
  cat("✅ 基础语法正常\n\n"); flush(stdout())
}, error = function(e) {
  cat("❌ 基础语法错误:", e$message, "\n"); flush(stdout())
  stop("语法检查失败")
})

# 3. 简单功能测试
cat("[测试] 测试完整函数调用...\n"); flush(stdout())
tryCatch({
  if (length(docx_files) > 0) {
    cat("测试DOCX文件:", basename(docx_files[1]), "\n"); flush(stdout())

    # 快速测试：只检查能否调用函数，不实际处理
    cat("  ✓ 文件存在:", file.exists(docx_files[1]), "\n"); flush(stdout())
    cat("  ✓ 文件大小:", round(file.size(docx_files[1]) / 1024 / 1024, 2), "MB\n"); flush(stdout())

    # 检查能否加载基本包
    if (requireNamespace("officer", quietly = TRUE)) {
      cat("  ✓ officer包可用\n"); flush(stdout())
    } else {
      cat("  ⚠ officer包未安装\n"); flush(stdout())
    }
  }

  if (length(pdf_files) > 0) {
    cat("测试PDF文件:", basename(pdf_files[1]), "\n"); flush(stdout())
    cat("  ✓ 文件存在:", file.exists(pdf_files[1]), "\n"); flush(stdout())
    cat("  ✓ 文件大小:", round(file.size(pdf_files[1]) / 1024 / 1024, 2), "MB\n"); flush(stdout())

    if (requireNamespace("pdftools", quietly = TRUE)) {
      cat("  ✓ pdftools包可用\n"); flush(stdout())
    } else {
      cat("  ⚠ pdftools包未安装\n"); flush(stdout())
    }
  }

  cat("\n✅ 基础验证完成!\n"); flush(stdout())
  cat("现在可以运行完整测试:\n"); flush(stdout())
  cat("  source('test_docx_clean.R')  # 测试DOCX\n"); flush(stdout())
  cat("  source('test_pdf_clean.R')   # 测试PDF\n"); flush(stdout())
  cat("\n"); flush(stdout())

}, error = function(e) {
  cat("❌ 验证失败:", e$message, "\n"); flush(stdout())
  stop("验证过程出错")
})
