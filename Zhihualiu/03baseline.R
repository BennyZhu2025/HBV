# Load only necessary packages
library(dplyr)       # Data manipulation
library(tableone)    # For creating Table 1
library(openxlsx)    # For writing multiple sheets in an Excel file

# Clear the environment
rm(list = ls())

# Set working directory
setwd("D:/OneDrive/2Jinlin_Hou/02刘志华/20240929Stage/analysis_second")

# Read the data
df <- read.csv("02data_impute_missing.csv", header = TRUE, stringsAsFactors = FALSE)

# Ensure column names are syntactically valid to avoid issues with special characters
colnames(df) <- make.names(colnames(df))

# Define columns related to adverse pregnancy outcomes (APO)
# Adjust column names to match the modified names after make.names()
apo_columns <- c('LGA', 'SGA', 'Stillbirth', 'Birth.defects',
                 'Neonatal.asphyxia', 'Intrauterine.distress',
                 'PROM', 'Preterm.birth', 'Postpartum.hemorrhage',
                 'Low.birth.weight.infant')

# Create the 'event' column indicating if any APO occurred
df$event <- as.integer(rowSums(df[, apo_columns]) > 0)

# Convert columns with fewer than 10 unique values to factors
df <- df %>%
  mutate(across(where(~ length(unique(.)) < 10), as.factor))

# Exclude APO-related columns and 'event' from variables for Table 1
variables <- setdiff(names(df), c(apo_columns, "event"))

# Function to create Table 1 for each stratification column and append to the results
create_and_append_table_one <- function(strata_column) {
  # Create Table 1 for the current strata
  tab1 <- CreateTableOne(vars = variables, strata = strata_column, data = df, test = TRUE)

  # Convert the table to a data frame
  table1 <- as.data.frame(print(tab1,printToggle=FALSE, showAllLevels = TRUE, smd = TRUE))
  table1 <- cbind(RowName = rownames(table1), table1)

  # Return the updated table
  return(table1)
}

# Initialize an Excel workbook to save results
wb <- createWorkbook()

# Loop through APO columns and the 'event' column to generate tables and save as sheets
for (strata in c(apo_columns, "event")) {
  table <- create_and_append_table_one(strata)

  # Add the table to a new sheet in the workbook
  addWorksheet(wb, sheetName = strata)
  writeData(wb, sheet = strata, table, rowNames = FALSE)
}

# Save the workbook to an Excel file
saveWorkbook(wb, file = "03results_baseline.xlsx", overwrite = TRUE)
