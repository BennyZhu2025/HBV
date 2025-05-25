# 0. 清理环境 & 设置选项
rm(list = ls()); gc()
options(error = recover)

# 1. 加载所需包
library(tidyverse)
library(readxl)
library(openxlsx)
library(glue)
library(nnet)
library(purrr)

# 2. 设置工作目录
work_dir <- "D:/OneDrive/2Jinlin_Hou/01尹雪如/20250516宝安妇幼数据分析/analysis"
setwd(work_dir)

# 3. 读数据 & 预处理
data_raw <- read_excel("01data_encoding.xlsx")

drop_vars <- c(
  "Last menstrual time", "First test pregnancy week", "Late pregnancy test week",
  "Time of first DNA test", "Gestation week of first DNA test", "Gestation week of last DNA test",
  "LGA", "Baby adverse outcome", "Amniotic fluid embolism",
  "Eclampsia", "Placenta accreta", "Placenta previa", "Placenta retention",
  "Placental abruption", "Postpartum hemorrhage", "Premature rupture of membranes",
  "Puerperal infection", "Threatened eclampsia", "Uterine rupture",
  "Perinatal death", "Stillbirth","Birth defect", "LBW", "Preterm birth", "Maternal adverse outcome")

data <- data_raw %>% select(-all_of(drop_vars))

outcomes <- c("Viral decrease")

pairs_list <- list(c("Drug starting week","Viral decrease"), c("Drug start stage","Viral decrease"))

# 定义线性函数
lm_regression <- function(df, outcome, predictors) {
  out_safe  <- paste0("`", outcome, "`")
  preds_safe <- paste0("`", predictors, "`")
  form <- as.formula(glue("{out_safe} ~ {paste(preds_safe, collapse = ' + ')}"))

  model <- tryCatch(
    lm(form, data = df),
    error = function(e) {
      message(glue("  拟合失败：{conditionMessage(e)}"))
      return(NULL)})

  if (is.null(model)) return(NULL)

  # 用 broom 提取系数、SE、CI、p
  tab <- broom::tidy(model, conf.int = TRUE) %>%
    rename(
      Predictor     = term,
      Estimate      = estimate,
      `Std. Error`  = std.error,
      `95% CI Lower`= conf.low,
      `95% CI Upper`= conf.high,
      P_value       = p.value
    ) %>%
    # 再做数值格式化
    mutate(
      Outcome     = outcome,
      Estimate    = round(Estimate, 2),
      `Std. Error`= round(`Std. Error`, 2),
      `95% CI Lower` = round(`95% CI Lower`, 2),
      `95% CI Upper` = round(`95% CI Upper`, 2),
      P_value    = round(P_value,2)) %>%
    select(Outcome, Predictor, Estimate, `Std. Error`, `95% CI Lower`, `95% CI Upper`, P_value)
  # 提取模型层面指标：R2 & Adjusted R2
  stats <- broom::glance(model)
  r2      <- round(stats$r.squared,   2)
  adj_r2  <- round(stats$adj.r.squared, 2)
  # 在每一行都加上这两个值
  tab <- tab %>%
    mutate(
      R_squared       = r2,
      Adjusted_R2     = adj_r2
    )
  return(tab)}

# 单变量线性回归
all_lm <- list()
for (pair in pairs_list) {
  predictor <- pair[1]
  outcome   <- pair[2]
  message(glue("正在分析：{predictor} → {outcome}"))

  df_i <- data %>% select(all_of(c(predictor, outcome))) %>% drop_na()

  if (n_distinct(df_i[[predictor]]) < 2 || n_distinct(df_i[[outcome]]) < 2) {
    message(glue("  跳过：{predictor} 或 {outcome} 无足够变异"))
    next
  }

  mod <- lm_regression(df_i, outcome, predictor)
  mod$Y<-outcome
  mod$X<-predictor
  if (!is.null(mod)) {all_lm[[length(all_lm) + 1]] <- mod}}
univariate_lm <- bind_rows(all_lm)
write.xlsx(univariate_lm,'03lm_univaraite.xlsx')

# 提取显著协变量
covariate_lm <- list()
vars <- setdiff(colnames(data), c(outcomes, "First DNA test result", "High viral load at first test",
                                  "Last DNA test result", "High viral load in last test",
                                  "Baby HBsAg after birth", "Baby HBsAb after birth",
                                  "HBsAg at first","Anti-HBs at first","HBeAg at first","Anti-HBe at first","Anti-HBc at first","First test pregnancy week",
                                  "HBsAg at delivery","Anti-HBs at delivery", "HBeAg at delivery", "Anti-HBe at delivery", "Anti-HBc at delivery","Late pregnancy test week"))
for (var in vars) {
  for (outcome in outcomes) {
    df_i <- data %>% select(all_of(c(var, outcome))) %>% drop_na()
    if (n_distinct(df_i[[var]]) < 2 || n_distinct(df_i[[outcome]]) < 2) next
    mod <- lm_regression(df_i, outcome, var)
    mod$Covariate <- var
    mod$Outcome <- outcome
    covariate_lm[[length(covariate_lm) + 1]] <- mod
  }
}
covariate_lm_df <- bind_rows(covariate_lm)

# 处理covariate_lr_df
sig_vars <- covariate_lm_df %>%
  filter(P_value < 0.05) %>%               # 过滤P值小于0.05的行
  filter(!grepl("Intercept", Predictor))   # 排除包含Intercept的Predictor

# 多变量线性回归
multiple_lm <- list()
for (pair in pairs_list) {
  predictor <- pair[1]
  outcome   <- pair[2]
  message(glue("正在分析（多变量）：{predictor} → {outcome}"))

  # 获取指定Outcome对应的唯一Covariates
  covars <- sig_vars %>%
    filter(Outcome == outcome) %>%
    pull(Covariate) %>%
    unique()  # 保证Covariates唯一

  if (is.null(covars)) covars <- character(0)
  predictors <- unique(c(covars, predictor))

  # 安全选择列
  selected_vars <- intersect(c(predictors, outcome), colnames(data))
  df_i <- data %>% select(all_of(selected_vars)) %>% drop_na()

  # 拿掉只有一个水平的 predictor
  valid_predictors <- keep(predictors, ~ n_distinct(na.omit(df_i[[.x]])) >= 2)

  # 检查 outcome 有变异，样本量足够
  if (n_distinct(df_i[[outcome]]) < 2 || nrow(df_i) < 10) {
    message(glue("  跳过：{predictor} → {outcome} 无足够变异或样本数过少"))
    next
  }

  # predictor 必须有效
  if (!(predictor %in% valid_predictors)) {
    message(glue("  跳过：{predictor} 在中无有效变异"))
    next
  }

  final_predictors <- unique(c(predictor, setdiff(valid_predictors, predictor)))

  # 回归分析
  mod <- lm_regression(df_i, outcome, final_predictors)
  mod$Y<-outcome
  mod$X<-predictor
  if (!is.null(mod)) {multiple_lm[[length(multiple_lm) + 1]] <- mod}}

multiple_lm_df <- bind_rows(multiple_lm)
write.xlsx(multiple_lm_df,'03lm_multiple.xlsx')

# 亚组分析
sub_lm <- list()
group_vars <- c("HBV status at first", "HBV status at delivery")

for (group_var in group_vars) {
  group_values <- unique(na.omit(data[[group_var]]))

  for (group in group_values) {
    message(glue("正在进行亚组分析：{group_var} = {group}"))

    data_sub <- data%>% filter(!!sym(group_var) == group)%>% drop_na()

    for (pair in pairs_list) {
      predictor <- pair[1]
      outcome   <- pair[2]

      covars <- sig_vars[[outcome]]
      if (is.null(covars)) covars <- character(0)
      predictors <- unique(c(covars, predictor))

      # 拿掉只有一个水平的 predictor
      valid_predictors <- predictors[sapply(predictors, function(var) {
        n_vals <- n_distinct(na.omit(data_sub[[var]]))
        return(n_vals >= 2)
      })]

      # outcome 也要有两个水平
      if (n_distinct(na.omit(data_sub[[outcome]])) < 2) {
        message(glue("  跳过：{predictor} → {outcome} 在子组 {group_var}={group} 中 outcome 无变异"))
        next
      }

      # 主 predictor 至少出现一次
      if (!(predictor %in% valid_predictors)) {
        message(glue("  跳过：{predictor} 在子组 {group_var}={group} 中无有效变异"))
        next
      }

      final_predictors <- unique(c(predictor, setdiff(valid_predictors, predictor)))

      df_i <- data_sub %>% select(all_of(c(final_predictors, outcome))) %>% drop_na()
      if (nrow(df_i) < 10) {
        message(glue("  跳过：样本数不足（<10） in subgroup {group_var} = {group}"))
        next
      }

      mod <- lm_regression(df_i, outcome, final_predictors)
      if (!is.null(mod)) {
        mod$Subgroup <- group
        mod$GroupVar <- group_var
        mod$Y<-outcome
        mod$X<-predictor
        sub_lm[[length(sub_lm) + 1]] <- mod}
    }
  }
}
sub_lm_df <- bind_rows(sub_lm)
write.xlsx(sub_lm_df,'03lm_subgroups.xlsx')
