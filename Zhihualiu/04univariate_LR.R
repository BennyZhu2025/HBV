# Load necessary libraries
library(dplyr)       # Data manipulation
library(tableone)    # For creating Table 1
library(openxlsx)    # For writing results to an Excel file
library(broom)       # For tidying logistic regression results
library(MASS)        # For confidence intervals

# Clear the environment
rm(list = ls())

# Set working directory
setwd("D:\\OneDrive\\2Jinlin_Hou\\02刘志华\\20240929Stage\\analysis_second")

# Read the data
df <- read.csv("02data_impute_missing.csv", header = TRUE, stringsAsFactors = FALSE)

# Ensure column names are syntactically valid to avoid issues with special characters
colnames(df) <- make.names(colnames(df))

# Define columns related to adverse pregnancy outcomes (APO)
apo_columns <- c('LGA', 'SGA', 'Stillbirth', 'Birth.defects',
                 'Neonatal.asphyxia', 'Intrauterine.distress',
                 'PROM', 'Preterm.birth', 'Postpartum.hemorrhage',
                 'Low.birth.weight.infant')

# Create the 'event' column indicating if any APO occurred
df$event <- as.integer(rowSums(df[, apo_columns]) > 0)

# Convert columns with fewer than 10 unique values to factors
df <- df %>%
  mutate(across(where(~ length(unique(.)) < 10), as.factor))

# Exclude APO-related columns and 'event' from variables for Table 1 and regression
variables <- setdiff(names(df), c(apo_columns, "event"))

# Store the results of univariate logistic regressions
results_list <- list()

# Loop over each outcome
for (outcome in c(apo_columns, "event")) {

  # Loop over each variable in variables
  for (var in variables) {

    # Create a formula for logistic regression
    formula <- as.formula(paste(outcome, "~", var))

    # Fit the logistic regression model
    model <- glm(formula, data = df, family = binomial)

    # Extract the coefficients, p-values, and confidence intervals
    summary_model <- broom::tidy(model, conf.int = TRUE, conf.level = 0.95)

    # Perform calculations for OR and CIs
    summary_model <- summary_model %>%
      filter(term != "(Intercept)") %>% # Exclude the intercept
      mutate(
        OR = round(exp(estimate), 2),
        lower_95_CI = round(exp(conf.low), 2),
        upper_95_CI = round(exp(conf.high), 2),
        p.value = round(p.value, 2),
        outcome = outcome,
        predictor = var,
        `OR (95% CI)` = paste0(OR, " (", lower_95_CI, ", ", upper_95_CI, ")")
      )

    # Append results to the list
    results_list[[paste(outcome, var, sep = "_")]] <- summary_model
  }
}

# Combine all results into a single data frame
results_df <- bind_rows(results_list)

# Filter the results for specific terms
final_results <- results_df %>%
  filter(term %in% c("Stage2", "Stage3", "Stage4", "Stage5")) %>%
  transmute(outcome, term, `OR (95% CI)`, p.value)

# Write results to an Excel file
write.xlsx(results_df, "04results_univariateLR.xlsx", rowNames = FALSE)

# Write the filtered results to an Excel file
write.xlsx(final_results, "04results_filtered_univariateLR.xlsx", rowNames = FALSE)
