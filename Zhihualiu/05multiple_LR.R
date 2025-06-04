# Load necessary libraries
library(dplyr)       # For data manipulation
library(tableone)    # For creating Table 1
library(openxlsx)    # For writing results to an Excel file
library(broom)       # For tidying logistic regression results
library(MASS)        # For confidence intervals

# Clear the environment
rm(list = ls())

# Set working directory
setwd("D:/OneDrive/2Jinlin_Hou/02刘志华/20240929Stage/analysis_second")

# Read the dataset
df <- read.csv("02data_impute_missing.csv", header = TRUE, stringsAsFactors = FALSE)

# Ensure column names are syntactically valid to avoid issues with special characters
colnames(df) <- make.names(colnames(df))

# Define columns related to adverse pregnancy outcomes (APO)
apo_columns <- c(
  'LGA', 'SGA', 'Stillbirth', 'Birth.defects',
  'Neonatal.asphyxia', 'Intrauterine.distress',
  'PROM', 'Preterm.birth', 'Postpartum.hemorrhage',
  'Low.birth.weight.infant'
)

# Create the 'event' column indicating if any APO occurred
df <- df %>%
  mutate(
    event = as.integer(rowSums(across(all_of(apo_columns))) > 0),
    across(where(~ length(unique(.)) < 10), as.factor) # Convert columns with <10 unique values to factors
  )

# Exclude APO-related columns and 'event' from variables for Table 1 and regression
variables <- setdiff(names(df), c(apo_columns, "event"))

# Read the significant variables from the Excel file
significant_vars <- read.xlsx("05results_significant.xlsx", sheet = 1)

# Extract outcome variable names
outcomes <- c(apo_columns,"event")#colnames(significant_vars)

# Initialize list to store results from logistic regression
results_list <- list()

# Perform multivariate logistic regression for each outcome
for (outcome in outcomes) {
  # Get predictor variables for the current outcome
  predictors <-variables #na.omit(significant_vars[[outcome]])

  if (length(predictors) > 0) {  # Proceed only if predictors exist
    # Create the formula for logistic regression
    formula <- as.formula(paste(outcome, "~", paste(predictors, collapse = " + ")))

    # Fit the logistic regression model
    model <- glm(formula, data = df, family = binomial)

    # Extract and format model results
    summary_model <- broom::tidy(model, conf.int = TRUE, conf.level = 0.95) %>%
      filter(term != "(Intercept)") %>% # Exclude intercept
      mutate(
        OR = round(exp(estimate), 2),
        lower_95_CI = round(exp(conf.low), 2),
        upper_95_CI = round(exp(conf.high), 2),
        p.value = round(p.value, 2),
        outcome = outcome
      )

    # Append results to the list
    results_list[[outcome]] <- summary_model
  }
}

# Combine all results into a single data frame
results_df <- bind_rows(results_list, .id = "Outcome")

# Save the results to an Excel file
write.xlsx(results_df, "05results_multivariateLR.xlsx", rowNames = TRUE)
