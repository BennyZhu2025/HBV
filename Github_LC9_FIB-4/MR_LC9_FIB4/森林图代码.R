rm(list = ls())
# 加载必要包
library(grid)
library(forestploter)
library(readxl)
library(dplyr)

# 读取数据
dt <- read_excel("01森林图.xlsx", sheet = 1) %>%
  filter(Method == "IVW") %>%
  mutate(across(
    c(
      OR1, OR2, OR3, OR4,
      or1_lci95, or2_lci95, or3_lci95, or4_lci95,
      or1_uci95, or2_uci95, or3_uci95, or4_uci95
    ),
    as.numeric
  ))

# 更改列名
# 清洗特定列的缺失值并转为字符型
dt[[1]] <- as.character(dt[[1]])
dt[[1]][is.na(dt[[1]])] <- ""
dt[[3]] <- as.character(dt[[3]])
dt[[3]][is.na(dt[[3]])] <- ""
dt[[15]] <- as.character(dt[[15]])
dt[[15]][is.na(dt[[15]])] <- ""
colnames(dt)[c(3, 15)] <- ""
colnames(dt)[c(11)] <- "OR (95%CI)"
colnames(dt)[c(12)] <- "1st p"
colnames(dt)[c(13)] <- "OR (95%CI)"
colnames(dt)[c(14)] <- "2nd p"
colnames(dt)[c(16)] <- "OR (95%CI)"
colnames(dt)[c(17)] <- "1st p"
colnames(dt)[c(18)] <- "OR (95%CI)"
colnames(dt)[c(19)] <- "2nd p"
#
cols_to_format <- c(12, 14, 17, 19)
dt[cols_to_format] <- lapply(dt[cols_to_format], function(x) {
  if (is.numeric(x)) formatC(x, format = "f", digits = 2) else x
})
# 创建一个主题对象，用于定制森林图的外观
tm <- forest_theme(
  base_size = 15, # 字体大小为15
  base_family = "serif", # 字体为Times New Roman

  ci_pch = 16, # 置信区间点的形状
  ci_fill = c("#BF5960", "#6F99AD"), # 置信区间点的颜色
  ci_col = c("#BF5960", "#6F99AD"), # 置信区间线条的颜色
  ci_lwd = 2, # 置信区间线的线宽为2
  # ci_Theight = 0.3,        #置信区间两端竖线高度为0.3
  # ci_lty = "dashed",       #置信区间的线形

  legend_name = "ID Group",
  legend_value = c(" 1st Exposure ID", " 2nd Exposure ID"), # 图例名称

  refline_col = "black", # 参考线的颜色
  # refline_lty = "solid",             #参考线的线形为实线
  refline_lwd = 2, # 参考线的线宽

  arrow_type = c("closed"), # 箭头类型为闭合式箭头

  core = list(
    padding = unit(c(15, 10), "mm"), # 做图区域高度
    bg_params = list(fill = c("white"))
  ), # 背景为白色

  # 标题对齐方式(左、居中、居中、居中、居中、居中、居中、居中、居中)
  colhead = list(fg_params = list(
    hjust = c(0, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0, 0),
    x = c(0.05, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.05, 0.05)
  ))
)
# 创建一个森林图对象，并赋值给p
p <- forest(dt[, c(1, 3, 12, 14, 15, 17, 19, 10, 26)], # 使用dt数据框的哪几列用于作图，可以通过View(dt)查看
  est = list(dt$OR1, dt$OR2, dt$OR3, dt$OR4), # 使用dt数据框中的值作为点估计值
  lower = list(dt$or1_lci95, dt$or2_lci95, dt$or3_lci95, dt$or4_lci95), # 置信区间的下限
  upper = list(dt$or1_uci95, dt$or2_uci95, dt$or3_uci95, dt$or4_uci95), # 置信区间的上限
  sizes = 0.6, # 点的大小
  ci_column = c(2, 5), # 置信区间显示在第2列
  ref_line = 1, # 在X=1处添加垂直参考线
  xlim = list(c(0, 3), c(0.9, 1.1)), # 设置X轴的范围
  ticks_at = list(c(0, 1, 2, 3), c(0.9, 1, 1.1)), # X轴的刻度分布
  nudge_y = 0.3, # 置信区间线条之间的距离
  theme = tm
) # 使用之前创建的tm主题
plot(p)

# 顶上方插入标题
g <- insert_text(p, text = "Forward MR", part = "header", col = 3:4, gp = gpar(fontface = "bold", fontsize = 15, fontfamily = "serif"))
g <- add_text(g, text = "Reverse MR", part = "header", col = 6:7, row = 1, gp = gpar(fontface = "bold", fontsize = 15, fontfamily = "serif"))

# 在标题部分的底部添加下划线
g <- add_border(g, part = "header", row = 1, where = "top", gp = gpar(lwd = 2))
g <- add_border(g, part = "header", row = 2, where = "bottom", gp = gpar(lwd = 2))
g <- add_border(g, part = "header", row = 1, col = c(3:4, 6:7), where = "bottom", gp = gpar(lwd = 2))

# 居中对齐，row和col指定行和列
g <- edit_plot(g, col = c(3:4, 6:7), which = "text", hjust = unit(0.5, "npc"), x = unit(0.5, "npc"))
g
# 保存为pdf文件，后续可使用PS/AI进一步美化
