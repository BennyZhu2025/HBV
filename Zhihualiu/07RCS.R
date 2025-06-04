# Load necessary libraries
library(dplyr)       # Data manipulation
library(tableone)    # For creating Table 1
library(openxlsx)    # For writing results to an Excel file
library(broom)       # For tidying logistic regression results
library(MASS)        # For confidence intervals

# Clear the environment
rm(list = ls())

# Set working directory
setwd("D:/OneDrive/8.Jinlin hou/刘志华/20240929Stage/20241119news folder")

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

# Store the significant predictors for each outcome
significant_predictors <- list()

# Perform univariate logistic regression for each outcome and variable
for (outcome in c(apo_columns, "event")) {
  significant_vars <- c()  # List to store significant variables for this outcome

  for (var in variables) {
    # Create a formula for logistic regression
    formula <- as.formula(paste(outcome, "~", var))

    # Fit the logistic regression model
    model <- glm(formula, data = df, family = binomial)

    # Extract p-value
    p_value <- summary(model)$coefficients[2, 4]  # p-value for the predictor

    # Check if the variable is significant (p < 0.05)
    if (p_value < 0.05) {
      significant_vars <- c(significant_vars, var)
    }
  }
  # Ensure 'Stage' is included in the significant variables
  if (!("Stage" %in% significant_vars)) {
    significant_vars <- c(significant_vars, "Stage")
  }
  # Store the significant variables for this outcome
  significant_predictors[[outcome]] <- significant_vars
}

# Perform multivariable logistic regression for each outcome using significant predictors
multi_results <- list()

for (outcome in names(significant_predictors)) {
  if (length(significant_predictors[[outcome]]) > 0) {
    # Create formula for multivariable logistic regression
    formula <- as.formula(paste(outcome, "~", paste(significant_predictors[[outcome]], collapse = " + ")))

    # Fit the multivariable logistic regression model
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
        p.value = round(p.value, 3),
        outcome = outcome,
        `OR (95% CI)` = paste0(OR, " (", lower_95_CI, ", ", upper_95_CI, ")")
      )

    # Store the results
    multi_results[[outcome]] <- summary_model
  }
}
