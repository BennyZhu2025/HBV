# ------------------ 1. 加载必要的 R 包 ------------------
library(MendelianRandomization)
library(TwoSampleMR)
library(data.table)
library(ieugwasr) # 从 IEU OpenGWAS 获取 GWAS 数据
library(dplyr) # 数据处理
library(tibble) # 创建数据框
library(openxlsx) # 读写 Excel 文件
library(future.apply) # 并行处理支持

# ------------------ 2. 定义多变量 MR 分析函数 ------------------
run_multivariable_mr <- function(primary_exposure_id, covariate_ids, outcome_id) {
  tryCatch(
    {
      # 构建所有暴露 ID 列表（主暴露 + 协变量）
      all_exposures <- c(primary_exposure_id, covariate_ids)
      # 提取暴露数据
      if (file.exists(paste0("data_multiplemr/exposure_", primary_exposure_id, ".csv"))) {
        exposure_data <- read.csv(file = paste0("data_multiplemr/exposure_", primary_exposure_id, ".csv"), header = TRUE, fileEncoding = "UTF-8")
      } else {
        exposure_data <- mv_extract_exposures(id_exposure = all_exposures, pval_threshold = 5e-06)
        write.xlsx(exposure_data, file = paste0("data_multiplemr/exposure_", primary_exposure_id, ".csv"))
      }
      # 提取结局数据
      if (file.exists(paste0("data_multiplemr/outcome_", primary_exposure_id, "_", outcome_id, ".csv"))) {
        outcome_data <- read.csv(file = paste0("data_multiplemr/outcome_", primary_exposure_id, "_", outcome_id, ".csv"), header = TRUE, fileEncoding = "UTF-8")
      } else {
        outcome_data <- extract_outcome_data(snps = unique(exposure_data$SNP), outcomes = outcome_id)
        write.xlsx(outcome_data, file = paste0("data_multiplemr/outcome_", primary_exposure_id, "_", outcome_id, ".csv"))
        message("结局提取完成：", outcome_data)
      }
      # 协调暴露与结局数据，使其位点一致
      if (file.exists(paste0("data_multiplemr/harmonised_", primary_exposure_id, "_", outcome_id, ".csv"))) {
        harmonised_data <- read.csv(file = paste0("data_multiplemr/harmonised_", primary_exposure_id, "_", outcome_id, ".csv"), header = TRUE, fileEncoding = "UTF-8")
      } else {
        harmonised_data <- mv_harmonise_data(exposure_data, outcome_data)
        write.xlsx(outcome_data, file = paste0("data_multiplemr/harmonised_", primary_exposure_id, "_", outcome_id, ".csv"))
      }
      # 运行多变量 MR 分析
      mv_result <- mv_multiple(harmonised_data)
      # 计算 OR（odds ratio）
      mv_or <- generate_odds_ratios(mv_result$result)
      # 输出列名以便调试
      print(colnames(mv_or))
      # 添加标记信息：主暴露、结局、是否为主暴露
      mv_or <- mv_or %>%
        mutate(
          primary_exposure = primary_exposure_id,
          outcome = outcome_id,
          is_primary = id.exposure == primary_exposure_id
        )
      return(mv_or)
    },
    error = function(e) {
      message("❌ MR 分析失败：", primary_exposure_id, " -> ", outcome_id, " : ", e$message)
      return(NULL)
    }
  )
}

# ------------------ 3. 串行运行多个暴露变量并保存结果 ------------------
run_mvmr_parallel <- function(exposure_ids, covariates_group, outcome_id, results_file) {
  if (file.exists(results_file)) {
    completed <- read.xlsx(results_file)
  } else {
    completed <- tibble()
    write.xlsx(completed, results_file, overwrite = TRUE)
  }

  already_done <- unique(completed$primary_exposure)
  remaining <- setdiff(exposure_ids, already_done)

  if (length(remaining) == 0) {
    message("✅ 所有主暴露分析已完成。")
    return(completed)
  }

  message("🚀 分析 ", length(remaining), " 个主暴露变量...")
  plan(multisession)
  results_list <- future_lapply(remaining, function(exposure_id) {
    message("➡️ 当前主暴露：", exposure_id)
    result <- tryCatch(
      run_multivariable_mr(primary_exposure_id = exposure_id, covariate_ids = covariates_group, outcome_id = outcome_id),
      error = function(e) {
        message("   ❌ 分析失败：", e$message)
        return(NULL)
      }
    )
  }, future.seed = TRUE)
  write.xlsx(completed, results_file, overwrite = TRUE)
  return(completed)
}

# ------------------ 4. 设置分析参数 ------------------
exposure_ids <- c(
  "ebi-a-GCST90096893", "ebi-a-GCST90093322", "ieu-b-4860", "ebi-a-GCST90029014",
  "ukb-a-224", "ukb-a-9", "ieu-a-1088", "ukb-b-19953", "ebi-a-GCST90029007",
  "ieu-b-109", "ebi-a-GCST90018956", "ebi-a-GCST90002232", "ebi-a-GCST90018955",
  "finn-b-I9_HYPTENS", "ukb-d-I9_HYPTENS", "ebi-a-GCST90018833", "ukb-d-F5_DEPRESSIO"
)

# 协变量
covariates_group <- c("finn-b-NAFLD") # "ebi-a-GCST90018804") #

# 结局变量 ID
outcome_id <- "finn-b-K11_FIBROCHIRLIV"

# 结果保存路径
results_file_path <- "02mvmr_results_HBV.xlsx"

# ------------------ 5. 启动主循环，重复运行分析任务 ------------------
for (i in 1:10000000) {
  message("🔁 第 ", i, " 次运行开始...")
  result <- run_mvmr_parallel(exposure_ids, covariates_group, outcome_id, results_file_path)
  Sys.sleep(1)
}
