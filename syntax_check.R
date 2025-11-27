# =============================================================================
# 快速语法检查脚本
# =============================================================================

cat("╔══════════════════════════════════════════════════════════════╗\n"); flush(stdout())
cat("║              语法检查脚本                                    ║\n"); flush(stdout())
cat("╚══════════════════════════════════════════════════════════════╝\n\n"); flush(stdout())

# 检查DOCX脚本语法
cat("[检查] test_docx_clean.R 语法...\n"); flush(stdout())
tryCatch({
  # 尝试解析但不执行
  parse("test_docx_clean.R")
  cat("✅ DOCX脚本语法正确\n\n"); flush(stdout())
}, error = function(e) {
  cat("❌ DOCX脚本语法错误:", e$message, "\n\n"); flush(stdout())
})

# 检查PDF脚本语法
cat("[检查] test_pdf_clean.R 语法...\n"); flush(stdout())
tryCatch({
  parse("test_pdf_clean.R")
  cat("✅ PDF脚本语法正确\n\n"); flush(stdout())
}, error = function(e) {
  cat("❌ PDF脚本语法错误:", e$message, "\n\n"); flush(stdout())
})

cat("✅ 语法检查完成!\n"); flush(stdout())
cat("\n现在可以运行:\n"); flush(stdout())
cat("  source('test_docx_clean.R')  # 测试DOCX\n"); flush(stdout())
cat("  source('test_pdf_clean.R')   # 测试PDF\n"); flush(stdout())
cat("\n"); flush(stdout())
