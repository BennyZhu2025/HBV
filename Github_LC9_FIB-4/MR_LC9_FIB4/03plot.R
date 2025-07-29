# 加载必要包
library(grid)
library(forestploter)
library(readxl)
library(dplyr)

# 读取 Excel 文件并筛选 Method 为 IVW 的行
df <- read_excel("01森林图.xlsx", sheet = 1) %>%
  filter(Method == "IVW") %>%
  mutate(across(
    c(
      OR1, OR2, OR3, OR4,
      or1_lci95, or2_lci95, or3_lci95, or4_lci95,
      or1_uci95, or2_uci95, or3_uci95, or4_uci95
    ),
    as.numeric
  ))

# 清洗特定列的缺失值并转为字符型
df[[1]] <- as.character(df[[1]])
df[[1]][is.na(df[[1]])] <- ""
df[[3]] <- as.character(df[[3]])
df[[3]][is.na(df[[3]])] <- ""
df[[15]] <- as.character(df[[15]])
df[[15]][is.na(df[[15]])] <- ""
colnames(df)[c(3, 15)] <- ""
colnames(df)[c(11)] <- "1st OR (96%CI)"
colnames(df)[c(12)] <- "1st p"
colnames(df)[c(13)] <- "2nd OR (96%CI)"
colnames(df)[c(14)] <- "2nd p"
colnames(df)[c(16)] <- "1st OR (96%CI)"
colnames(df)[c(17)] <- "1st p"
colnames(df)[c(18)] <- "2nd OR (96%CI)"
colnames(df)[c(19)] <- "2nd p"

# 提取表格显示的列
cols_to_format <- c(12, 14, 17, 19)
df[cols_to_format] <- lapply(df[cols_to_format], function(x) {
  if (is.numeric(x)) formatC(x, format = "f", digits = 2) else x
})
table_columns <- df[, c(1, 3, 12, 14, 15, 17, 19, 10, 26)]

# 绘制森林图
p <- forest(
  data = table_columns,
  est = list(df$OR1, df$OR2, df$OR3, df$OR4),
  lower = list(df$or1_lci95, df$or2_lci95, df$or3_lci95, df$or4_lci95),
  upper = list(df$or1_uci95, df$or2_uci95, df$or3_uci95, df$or4_uci95),
  ci_column = c(2, 5),
  ref_line = 1,
  colgap = unit(20, "mm"),
  xlim = list(c(0, 2), c(0, 2)),
  xticks = c(0, 1.0, 2.0),
  nudge_y = 0.2,
  theme = forest_theme(
    base_size = 8,
    refline_gp = gpar(lty = "solid"),
    footnote_gp = gpar(col = "blue"),
    ci_pch = c(15, 18),
    ci_col = c("#4485C7", "#629C35"),
    legend_name = "Group",
    legend_value = c("1st ID", "2nd ID"),
    vertline_lty = c("dashed", "dotted"),
    vertline_col = c("#d6604d", "#d6604d")
  ),
  plotwidth = unit(10, "cm"), colgap = unit(2, "mm"),
  col_rel_width = c(2, 10, 2, 2, 10, 2, 2, 3, 3),
  panel_size = list(unit(200, "mm"), unit(200, "mm"))
)

# 显示图形
print(p)
