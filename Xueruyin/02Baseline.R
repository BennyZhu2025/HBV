# ========================================================
# 1. 加载分析所需的 R 包
# ========================================================
library(tidyverse)     # 数据处理与可视化（包括 dplyr, ggplot2 等）
library(gtsummary)     # 医学统计表格生成
library(flextable)     # 表格美化（配合 gtsummary 使用）
library(officer)       # Word/PPT 等文档导出支持
library(readxl)        # 读取 Excel 文件
library(nnet)          # 多项式逻辑回归（multinom 函数）

# 清空环境变量，释放内存
rm(list = ls())
gc()
options(error = recover)  # 遇到错误时启动调试

# ========================================================
# 2. 设置工作目录（可选）
# ========================================================
work_dir <- "D:/OneDrive/2Jinlin_Hou/01尹雪如/20250516宝安妇幼数据分析/analysis"
if (dir.exists(work_dir)) {
  setwd(work_dir)
} else {
  warning("工作目录不存在，请检查路径。")
}

# ========================================================
# 3. 数据读取与预处理
# ========================================================
# 读取原始编码数据
data <- read_excel("01data_encoding.xlsx")

# 删除不必要的时间类与辅助列
drop_vars <- c("Last menstrual time", "First test pregnancy week",
               "Late pregnancy test week", "Time of first DNA test",
               "Gestation week of first DNA test", "Gestation week of last DNA test",
               "HBsAg at first", "Anti-HBs at first", "HBeAg at first",
               "Anti-HBe at first", "Anti-HBc at first",
               "HBsAg at delivery", "Anti-HBs at delivery", "HBeAg at delivery",
               "Anti-HBe at delivery", "Anti-HBc at delivery")
data <- data[ , !(names(data) %in% drop_vars)]

# ========================================================
# 4. 按变量功能分组整理数据（便于建模与描述分析）
# ========================================================

## 4.1 人口学特征
demographics <- data[c("Ethnicity", "Age", "Marital status", "Husband age",
                       "Education", "Occupation", "Local", "Guangdong")]

## 4.2 孕产史与分娩信息
obstetrics <- data[c("Gravida", "Parity", "First check pregnancy week", "Delivery pregnancy week",  "Number of fetuses", "Delivery method", "Delivery hospital")]

## 4.3 初筛乙肝五项
hbv_early <- data[c()]

## 4.4 分娩时乙肝五项
hbv_late <- data[c()]

## 4.5 病毒载量与肝功能检测
dna_test <- data[c("Whether DNA test was conducted", "First DNA test result",
                   "High viral load at first test", "Whether Last DNA test was conducted",
                   "Last DNA test result", "High viral load in last test", "ALT", "AST", "TBIL", "Viral decrease")]

## 4.6 新生儿基本信息
neonatal <- data[c("Baby gender", "Baby weight", "Baby height", "Baby head circumference (cm)",
                   "Apgar1", "Apgar5", "Baby HBsAg after birth", "Baby HBsAb after birth")]

## 4.7 妊娠期并发症
maternal_complications <- data[c("Gestational diabetes", "Gestational hypertension")]

## 4.8 HBV 状态分类变量
hbv_status <- data[c("HBV status at first", "HBV status at delivery")]

## 4.9 抗病毒用药信息
antiviral <- data[c("Drug", "Drug starting week", "Drug start stage")]

## 4.10 母婴结局变量
outcomes <- data[c("Baby weight Category", "Maternal adverse outcome",
                   "Baby adverse outcome", "Preterm birth", "LBW",
                   "Birth defect", "Perinatal death", "Stillbirth")]

# -------------------------------
# 6. 定义基线表函数
# -------------------------------
# # 生成基线表函数
create_and_save_table <- function(data, group_var, filename) {
  tbl <- data %>%
    tbl_summary(
      by = all_of(group_var),
      type = list(all_continuous() ~ "continuous2","Baby head circumference (cm)" ~ "continuous2"),
      statistic = all_continuous() ~ "{mean} ± {sd}",
      missing_text = 'missing',
      digits = all_continuous() ~ 2
    ) %>%
    add_p(
      pvalue_fun = ~style_pvalue(.x, digits = 2),
      test = list(
        all_continuous() ~ "t.test",
        all_categorical() ~ "chisq.test"
      )
    ) %>%
    add_overall() %>%
    add_stat_label() %>%
    as_flex_table()

  # 导出为 Word 文档
  read_docx() %>%
    body_add_flextable(tbl) %>%
    print(target = filename)
}

# -------------------------------
# 7. 基线表
# -------------------------------
# 组合所需分析变量
analysis_data <- bind_cols(
  demographics, obstetrics, hbv_early, hbv_late,
  dna_test, neonatal, maternal_complications, hbv_status, antiviral,
  outcomes
)

# 出生缺陷
create_and_save_table(
  data = analysis_data,
  group_var = "Birth defect",
  filename = "02baseline_birth_defect.docx"
)
# 早产
create_and_save_table(
  data = analysis_data,
  group_var = "Preterm birth",
  filename = "02baseline_preterm_birth.docx"
)
# 低出生体重
create_and_save_table(
  data = analysis_data,
  group_var = "LBW",
  filename = "02baseline_LBW.docx"
)
# 新生儿不良结局
create_and_save_table(
  data = analysis_data,
  group_var = "Baby weight Category",
  filename = "02baseline_baby_weight.docx"
)
# 围产期死亡
create_and_save_table(
  data = analysis_data,
  group_var = "Perinatal death",
  filename = "02baseline_perinatal_death.docx"
)
# 产妇不良结局
create_and_save_table(
  data = analysis_data,
  group_var = "Maternal adverse outcome",
  filename = "02baseline_maternal_adverse_outcome.docx"
)

# 用药时间
create_and_save_table(
  data = analysis_data,
  group_var = "Drug start stage",
  filename = "02baseline_drug_start_stage_outcome.docx"
)
