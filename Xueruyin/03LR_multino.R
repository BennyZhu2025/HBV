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

outcomes <- c("Baby weight Category")

pairs_list <- list(
  c("Drug", "Baby weight Category"),
  c("Drug starting week", "Baby weight Category"),
  c("Drug start stage", "Baby weight Category")
)

# 定义多分类函数
lr_regression <- function(df, outcome, predictors) {
  out_safe  <- paste0("`", outcome, "`")
  preds_safe <- paste0("`", predictors, "`")
  form <- as.formula(glue("{out_safe} ~ {paste(preds_safe, collapse = ' + ')}"))

  model <- tryCatch(
    nnet::multinom(form, data = df, trace = FALSE),
    error = function(e) {
      message(glue("  拟合失败：{conditionMessage(e)}"))
      return(NULL)})

  if (is.null(model)) return(NULL)

  sm <- summary(model)
  coefs <- sm$coefficients
  ses   <- sm$standard.errors
  if (!is.matrix(coefs) || nrow(coefs) == 0 || !is.matrix(ses)) {
    message("  拟合结果为空（无系数），跳过")
    return(NULL)
  }

  # 构建输出
  results <- list()
  outcome_levels <- rownames(coefs)

  for (i in seq_len(nrow(coefs))) {
    est <- coefs[i, ]
    se  <- ses[i, ]
    or  <- round(exp(est), 2)
    ci_l <- round(exp(est - 1.96 * se), 2)
    ci_u <- round(exp(est + 1.96 * se), 2)
    z <- est / se
    p_val <- round(2 * (1 - pnorm(abs(z))), 3)

    res <- data.frame(
      OutcomeLevel = outcome_levels[i],
      Predictor = names(est),
      OR = or,
      `95% CI Lower` = ci_l,
      `95% CI Upper` = ci_u,
      P_value = p_val,
      stringsAsFactors = FALSE
    )
    results[[i]] <- res
  }

  result_df <- bind_rows(results)
  result_df$Outcome <- outcome
  return(result_df)
}

# 单变量线性回归
all_lr <- list()
for (pair in pairs_list) {
  predictor <- pair[1]
  outcome   <- pair[2]
  message(glue("正在分析：{predictor} → {outcome}"))

  df_i <- data %>% select(all_of(c(predictor, outcome))) %>% drop_na()

  if (n_distinct(df_i[[predictor]]) < 2 || n_distinct(df_i[[outcome]]) < 2) {
    message(glue("  跳过：{predictor} 或 {outcome} 无足够变异"))
    next
  }

  mod <- lr_regression(df_i, outcome, predictor)
  mod$Y<-outcome
  mod$X<-predictor
  if (!is.null(mod)) {all_lr[[length(all_lr) + 1]] <- mod}}
univariate_lr <- bind_rows(all_lr)
# 格式化 P 值：保留两位小数，小于 0.01 显示为 "<0.01"
univariate_lr <- univariate_lr %>%
  mutate(P_value = ifelse(P_value < 0.01, "<0.01", sprintf("%.2f", P_value)))
write.xlsx(univariate_lr,'03lr_multino_univaraite.xlsx')

# 提取显著协变量
covariate_lr <- list()
vars <- setdiff(colnames(data), c(outcomes, "First DNA test result", "High viral load at first test",
                                  "Last DNA test result", "High viral load in last test",
                                  "Baby HBsAg after birth", "Baby HBsAb after birth",
                                  "HBsAg at first","Anti-HBs at first","HBeAg at first","Anti-HBe at first","Anti-HBc at first","First test pregnancy week",
                                  "HBsAg at delivery","Anti-HBs at delivery", "HBeAg at delivery", "Anti-HBe at delivery", "Anti-HBc at delivery","Late pregnancy test week"))
for (var in vars) {
  for (outcome in outcomes) {
    df_i <- data %>% select(all_of(c(var, outcome))) %>% drop_na()
    if (n_distinct(df_i[[var]]) < 2 || n_distinct(df_i[[outcome]]) < 2) next
    mod <- lr_regression(df_i, outcome, var)
    mod$Covariate <- var
    mod$Outcome <- outcome
    covariate_lr[[length(covariate_lr) + 1]] <- mod
  }
}
covariate_lr_df <- bind_rows(covariate_lr)

# 处理covariate_lr_df
sig_vars <- covariate_lr_df %>%
  filter(P_value < 0.05) %>%               # 过滤P值小于0.05的行
  filter(!grepl("Intercept", Predictor))   # 排除包含Intercept的Predictor


# 多变量线性回归
multiple_lr <- list()
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
  mod <- lr_regression(df_i, outcome, final_predictors)
  mod$Y<-outcome
  mod$X<-predictor
  if (!is.null(mod)) {multiple_lr[[length(multiple_lr) + 1]] <- mod}}

multiple_lr_df <- bind_rows(multiple_lr)
# 格式化 P 值：保留两位小数，小于 0.01 显示为 "<0.01"
multiple_lr_df <- multiple_lr_df %>%
  mutate(P_value = ifelse(P_value < 0.01, "<0.01", sprintf("%.2f", P_value)))
write.xlsx(multiple_lr_df,'03lr_multino_multiple.xlsx')

# 亚组分析
sub_lr <- list()
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

      mod <- lr_regression(df_i, outcome, final_predictors)
      if (!is.null(mod)) {
      mod$Subgroup <- group
      mod$GroupVar <- group_var
      mod$Y<-outcome
      mod$X<-predictor
      sub_lr[[length(sub_lr) + 1]] <- mod}
    }
  }
}
sub_lr_df <- bind_rows(sub_lr)
write.xlsx(sub_lr_df,'03lr_multino_subgroups.xlsx')
