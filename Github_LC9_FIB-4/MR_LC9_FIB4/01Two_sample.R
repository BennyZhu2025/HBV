# 安装并加载必要包
packages <- c("TwoSampleMR", "ieugwasr", "dplyr")
install_if_missing <- function(pkg) {
  if (!requireNamespace(pkg, quietly = TRUE)) install.packages(pkg)
  suppressPackageStartupMessages(library(pkg, character.only = TRUE))
}
invisible(lapply(packages, install_if_missing))

# 设置暴露（exposure）和结局（outcome）的 GWAS ID（可替换）
exposure_gwas_id <- "ebi-a-GCST90096893"
outcome_gwas_id <- "finn-b-K11_FIBROCHIRLIV"

# 提取并过滤强效工具变量（不包含条件语句或错误处理）
get_strong_instruments <- function(gwas_id) {
  instruments <- extract_instruments(outcomes = gwas_id)
  instruments <- clump_data(instruments)
  instruments$F_stat <- (instruments$beta.exposure / instruments$se.exposure)^2
  dplyr::filter(instruments, F_stat > 10)
}

# 执行双向MR分析
run_bidirectional_mr <- function(exposure_id, outcome_id) {
  # 正向 MR：暴露 -> 结局
  exposure_instruments <- get_strong_instruments(exposure_id)
  outcome_data_fwd <- extract_outcome_data(snps = exposure_instruments$SNP, outcomes = outcome_id)
  harmonised_fwd <- harmonise_data(exposure_dat = exposure_instruments, outcome_dat = outcome_data_fwd)
  mr_result_fwd <- mr(harmonised_fwd, method_list = "mr_ivw")
  odds_ratios_fwd <- generate_odds_ratios(mr_result_fwd)

  # 反向 MR：结局 -> 暴露
  outcome_instruments <- get_strong_instruments(outcome_id)
  exposure_data_rev <- extract_outcome_data(snps = outcome_instruments$SNP, outcomes = exposure_id)
  harmonised_rev <- harmonise_data(exposure_dat = outcome_instruments, outcome_dat = exposure_data_rev)
  mr_result_rev <- mr(harmonised_rev, method_list = "mr_ivw")
  odds_ratios_rev <- generate_odds_ratios(mr_result_rev)

  list(
    forward_mr = odds_ratios_fwd,
    reverse_mr = odds_ratios_rev
  )
}

# 运行分析并打印结果
mr_results <- run_bidirectional_mr(exposure_gwas_id, outcome_gwas_id)
print(mr_results)
