# ---------------------------
# Installing libraries
# ---------------------------
install.packages("here")
install.packages("tidyverse")
install.packages("dplyr")
install.packages("readxl")
install.packages("lubridate")
install.packages("skimr")
install.packages("scales")
install.packages("stringr")
install.packages("ggcorrplot")
install.packages("caret")
install.packages("nnet")
install.packages("randomForest")
install.packages("xgboost")
install.packages("janitor")
install.packages("e1071")
install.packages("purrr")
install.packages("car")
install.packages("tibble")

# ---------------------------
# LOAD ALL LIBRARIES
# ---------------------------
library(here)         # For building file paths robustly
library(tidyverse)    # Main data science suite (includes dplyr, ggplot2, etc.)
library(dplyr)        # Data wrangling
library(readxl)       # Reading Excel files
library(lubridate)    # For working with dates
library(skimr)        # For quick, rich data summaries
library(scales)       # For axis formatting, etc.
library(stringr)      # String handling (cleaning, searching)
library(ggcorrplot)   # For correlation heatmaps
library(caret)        # For machine learning: splitting, model building
library(nnet)         # Neural networks
library(randomForest) # Random forest models
library(xgboost)      # Extreme gradient boosting
library(janitor)      # Clean up column names
library(e1071)        # For skewness, SVM, etc.
library(purrr)        # Functional programming (map, etc.)
library(car)          # For checking collinearity (VIF)
library(tibble)       # For nice tibbles (data frames)


# ---------------------------
# Load Data
# ---------------------------
ast_dtls <- read_excel(here("data", "capstn_prj_dt_set.xlsx"), sheet = 1)
wo_cost <- read_excel(here("data", "capstn_prj_dt_set.xlsx"), sheet = 2)
miles <- read_excel(here("data", "capstn_prj_dt_set.xlsx"), sheet = 3)

glimpse(ast_dtls)
glimpse(wo_cost)
glimpse(miles)

# Load necessary packages
library(dplyr)
library(lubridate)
library(janitor)
library(stringr)

# ----------------------------------------
#  Cleaning Asset Details (ast_dtls)
# ----------------------------------------
ast_dtls_clean <- ast_dtls %>%
  clean_names() %>%
  mutate(
    unit_no = str_trim(unit_number),
    in_service_date = as.Date(in_service_date),
    disposal_date = as.Date(disposal_date),
    unit_age = round(as.numeric(difftime(Sys.Date(), in_service_date, units = "days")) / 365.25, 1),
    service_duration_days = as.numeric(difftime(disposal_date, in_service_date, units = "days")),
    is_active = if_else(unit_status == "IN SERVICE", 1, 0),
    year = as.numeric(year)
  ) %>%
  select(company, unit_no, unit_status, category, category_class_desc, year, make, model,
         in_service_date, disposal_date, meter, unit_age, service_duration_days, is_active)

# ----------------------------------------
# Cleaning Work Orders (wo_cost)
# ----------------------------------------
wo_cost_clean <- wo_cost %>%
  clean_names() %>%
  mutate(
    unit_no = str_trim(unit_no),
    open_date = mdy_hm(open_date),
    wo_completed_date = mdy_hm(wo_completed_date),
    wo_duration_days = as.numeric(difftime(wo_completed_date, open_date, units = "days")),
    wo_status = factor(wo_status, levels = c("O", "C", "X"), ordered = TRUE),
    total_labor_cost = coalesce(external_labor_cost, 0) + coalesce(labor_amount, 0) + coalesce(labor_tax, 0),
    total_parts_cost = coalesce(part_amount, 0) + coalesce(parts_tax, 0),
    line_total = coalesce(line_total, total_labor_cost + total_parts_cost),
    total_cost = total_labor_cost + total_parts_cost
  ) %>%
  select(fiscal_period, unit_no, company, category_class, cat_class_desc,
         wo_nbr, wo_reason_desc, open_date, wo_completed_date, wo_duration_days,
         wo_status, job, job_description, job_reason_description,
         total_labor_cost, total_parts_cost, total_cost, line_total)

# ----------------------------------------
#  Cleaning Utilization (miles)
# ----------------------------------------
miles_clean <- miles %>%
  clean_names() %>%
  mutate(
    unit_no = str_trim(unit_no),
    annual_mileage = coalesce(annualized_mileage, miles_driven),
    average_miles_per_month = round(annual_mileage / coalesce(months_with_mileage, 12), 2),
    utilized_days_per_month = utilized_days_per_month
  ) %>%
  select(unit_no, company, asset_type_asset_type, model_year, miles_driven,
         months_with_mileage, annual_mileage, average_miles_per_month, utilized_days_per_month)

library(dplyr)
library(e1071)
library(purrr)

# Function to calculate skewness for numeric columns
get_skewness <- function(df, sheet_name) {
  df %>%
    select(where(is.numeric)) %>%
    map_df(~ tibble(skewness = skewness(., na.rm = TRUE)), .id = "column") %>%
    mutate(sheet = sheet_name, .before = 1)
}

# Run skewness analysis for each sheet
skew_ast_dtls <- get_skewness(ast_dtls_clean, "Asset_Details")
skew_wo_cost  <- get_skewness(wo_cost_clean, "Work_Orders")
skew_miles    <- get_skewness(miles_clean, "Mileage")

# Combine all into one summary
skew_summary <- bind_rows(skew_ast_dtls, skew_wo_cost, skew_miles)

# Print result
print(skew_summary)

# -------------------------------
#          Asset Details
# -------------------------------
# Check missing values
sapply(ast_dtls_clean, function(x) sum(is.na(x)))


# Handle missing values 
median_service_date <- median(ast_dtls_clean$in_service_date, na.rm = TRUE)
median_meter <- median(ast_dtls_clean$meter, na.rm = TRUE)

ast_dtls_clean <- ast_dtls_clean %>%
  mutate(
    in_service_date = if_else(is.na(in_service_date), median_service_date, in_service_date),
    unit_age = round(as.numeric(difftime(Sys.Date(), in_service_date, units = "days")) / 365.25, 1),
    service_duration_days = if_else(is.na(service_duration_days), 0, service_duration_days),
    meter = if_else(is.na(meter), median_meter, meter)
  )

# -------------------------------
#            Work Orders
# -------------------------------
# Check missing values
sapply(wo_cost_clean, function(x) sum(is.na(x)))


## Simplify job_description using keywords
wo_cost_clean <- wo_cost_clean %>%
  mutate(
    job_group = case_when(
      str_detect(tolower(job_description), "inspect|pm|prevent") ~ "INSPECTION",
      str_detect(tolower(job_description), "repair|replace|fix") ~ "REPAIR",
      str_detect(tolower(job_description), "clean") ~ "CLEANING",
      str_detect(tolower(job_description), "install|mount") ~ "INSTALLATION",
      str_detect(tolower(job_description), "adjust|tighten|align") ~ "ADJUSTMENT",
      str_detect(tolower(job_description), "test|check|verify") ~ "TESTING",
      str_detect(tolower(job_description), "oil|grease|fluid|lubricate") ~ "LUBRICATION",
      TRUE ~ "OTHER"
    ),
    
    job_reason_group = case_when(
      str_detect(tolower(job_reason_description), "technician|found|reported") ~ "TECH FOUND ISSUE",
      str_detect(tolower(job_reason_description), "routine|prevent|pm") ~ "PREVENTIVE MAINTENANCE",
      str_detect(tolower(job_reason_description), "damage|accident") ~ "DAMAGE",
      str_detect(tolower(job_reason_description), "customer|reported|driver") ~ "REPORTED BY USER",
      TRUE ~ "OTHER"
    )
  )


# -------------------------------
#          Utilization
# -------------------------------
# Check missing values
sapply(miles_clean, function(x) sum(is.na(x)))

# Calculate median utilized days per month
median_utilized_days <- median(miles_clean$utilized_days_per_month, na.rm = TRUE)

# Impute missing values
miles_clean <- miles_clean %>%
  mutate(
    months_with_mileage = ifelse(is.na(months_with_mileage), 12, months_with_mileage),
    utilized_days_per_month = ifelse(is.na(utilized_days_per_month), median_utilized_days, utilized_days_per_month)
  )

# -------------------------------
#  Aggregating Work Order Data
# -------------------------------
wo_unit_summary <- wo_cost_clean %>%
  group_by(unit_no) %>%
  summarise(
    total_jobs = n_distinct(wo_nbr),
    total_line_cost = sum(line_total, na.rm = TRUE),
    avg_line_cost = mean(line_total, na.rm = TRUE),
    total_cost = sum(total_cost, na.rm = TRUE),
    unplanned_jobs = sum(str_detect(tolower(wo_reason_desc), "unplanned"), na.rm = TRUE),
    first_repair_date = min(open_date, na.rm = TRUE),
    last_repair_date = max(wo_completed_date, na.rm = TRUE),
    .groups = "drop"
  )

# -------------------------------
# Aggregating Utilization Data
# -------------------------------
miles_unit_summary <- miles_clean %>%
  group_by(unit_no) %>%
  summarise(
    annual_mileage = mean(annual_mileage, na.rm = TRUE),
    average_miles_per_month = mean(average_miles_per_month, na.rm = TRUE),
    months_with_mileage = mean(months_with_mileage, na.rm = TRUE),
    utilized_days_per_month = mean(utilized_days_per_month, na.rm = TRUE),
    total_miles_driven = sum(miles_driven, na.rm = TRUE),
    .groups = "drop"
  )

# -------------------------------
#  Joining All Cleaned Data Sets
# -------------------------------
fleet_joined <- ast_dtls_clean %>%
  inner_join(wo_unit_summary, by = "unit_no") %>%
  inner_join(miles_unit_summary, by = "unit_no")

# -------------------------------
#  Inspecting Final Data set
# -------------------------------
glimpse(fleet_joined)

summary(fleet_joined)

# -------------------------------------
# Preparing Data for Linear Regression
# -------------------------------------
# Select features and remove rows with any remaining NAs
model_data <- fleet_joined %>%
  select(total_line_cost, unit_age, meter, annual_mileage,
         average_miles_per_month, months_with_mileage,
         utilized_days_per_month, total_jobs, avg_line_cost, is_active) %>%
  na.omit()

# Split into training and testing sets
set.seed(9999999)
train_index <- createDataPartition(model_data$total_line_cost, p = 0.8, list = FALSE)
train_data <- model_data[train_index, ]
test_data <- model_data[-train_index, ]

# -------------------------------------
# Fittng Linear Regression Model
# -------------------------------------
lm_model <- lm(total_line_cost ~ ., data = train_data)

# Predict on test set
lm_pred <- predict(lm_model, test_data)

# Evaluate performance
lm_rmse <- RMSE(lm_pred, test_data$total_line_cost)
lm_r2 <- R2(lm_pred, test_data$total_line_cost)

# Print results
cat(" Linear Regression Results\n")

# --- Diagnostics for Linear Model ---
par(mfrow = c(2, 2))
plot(lm_model)

# Check for multicollinearity
library(car)
vif(lm_model)

cat("RMSE:", round(lm_rmse, 2), "\n")

cat("R-squared:", round(lm_r2, 4), "\n")

#  Show summary
summary(lm_model)

# -------------------------------------
#  Preparing Data for Log-Linear Regression
# -------------------------------------
model_data_log <- fleet_joined %>%
  select(total_line_cost, unit_age, meter, annual_mileage,
         average_miles_per_month, months_with_mileage,
         utilized_days_per_month, total_jobs, avg_line_cost, is_active) %>%
  filter(total_line_cost > 0) %>%              # Remove zero-cost rows (log(0) is undefined)
  mutate(log_total_line_cost = log(total_line_cost)) %>%
  na.omit()

# Train-test split
set.seed(9999999)
train_index <- createDataPartition(model_data_log$log_total_line_cost, p = 0.8, list = FALSE)
train_data <- model_data_log[train_index, ]
test_data <- model_data_log[-train_index, ]

# -------------------------------------
#  Fitting Log-Linear Regression Model
# -------------------------------------
log_lm_model <- lm(log_total_line_cost ~ unit_age + meter + annual_mileage +
                     average_miles_per_month + months_with_mileage +
                     utilized_days_per_month + total_jobs + avg_line_cost + is_active,
                   data = train_data)

# Predict log-cost on test set
log_pred <- predict(log_lm_model, test_data)

# Back-transform predictions to dollar scale
cost_pred <- exp(log_pred)

# Evaluate model
log_rmse <- RMSE(cost_pred, test_data$total_line_cost)
log_r2 <- R2(cost_pred, test_data$total_line_cost)

# -------------------------------------
#  Printing Evaluation
# -------------------------------------
cat(" Log-Linear Regression Results\n")

cat("RMSE (Back-transformed):", round(log_rmse, 2), "\n")

cat("R-squared (Back-transformed):", round(log_r2, 4), "\n")

# View model summary
summary(log_lm_model)

# -------------------------------------
#    Preparing data
# -------------------------------------
model_data_rf <- fleet_joined %>%
  select(total_line_cost, unit_age, meter, annual_mileage,
         average_miles_per_month, months_with_mileage,
         utilized_days_per_month, total_jobs, avg_line_cost, is_active) %>%
  na.omit()

set.seed(9999999)
train_index <- createDataPartition(model_data_rf$total_line_cost, p = 0.8, list = FALSE)
train_data <- model_data_rf[train_index, ]
test_data  <- model_data_rf[-train_index, ]

# -------------------------------------
#  Tune and Train Random Forest
# -------------------------------------
control <- trainControl(method = "cv", number = 5)
tunegrid <- expand.grid(.mtry = c(2, 3, 4, 5, 6))

rf_tuned <- train(
  total_line_cost ~ .,
  data = train_data,
  method = "rf",
  metric = "RMSE",
  tuneGrid = tunegrid,
  trControl = control,
  ntree = 500,
  importance = TRUE
)

# Best mtry
print(rf_tuned)

best_mtry <- rf_tuned$bestTune$mtry

# -------------------------------------
#  Predict and Evaluate
# -------------------------------------
rf_pred <- predict(rf_tuned, test_data)
rf_rmse <- RMSE(rf_pred, test_data$total_line_cost)
rf_r2   <- R2(rf_pred, test_data$total_line_cost)

cat("Random Forest Regression Results\n")

cat("Best mtry:", best_mtry, "\n")

cat("RMSE:", round(rf_rmse, 2), "\n")

cat("R-squared:", round(rf_r2, 4), "\n")


# -------------------------------------
#  Feature Importance Plot
# -------------------------------------
varImpPlot(rf_tuned$finalModel, main = "Variable Importance - Random Forest")

# -------------------------------------
#   Add Composite Variables
# -------------------------------------
fleet_enhanced <- fleet_joined %>%
  mutate(
    jobs_per_year = total_jobs / (unit_age + 0.1),  # Avoid division by zero
    mileage_per_age = annual_mileage / (unit_age + 0.1),
    age_x_jobs = unit_age * total_jobs,
    age_x_mileage = unit_age * annual_mileage
  )

# -------------------------------------
#  Preparing Data
# -------------------------------------
model_data_enh <- fleet_enhanced %>%
  select(total_line_cost, unit_age, meter, annual_mileage, average_miles_per_month,
         months_with_mileage, utilized_days_per_month, total_jobs, avg_line_cost, is_active,
         jobs_per_year, mileage_per_age, age_x_jobs, age_x_mileage) %>%
  na.omit()

set.seed(9999999)
train_index <- createDataPartition(model_data_enh$total_line_cost, p = 0.8, list = FALSE)
train_data <- model_data_enh[train_index, ]
test_data <- model_data_enh[-train_index, ]

# -------------------------------------
#  Tune and Train Random Forest
# -------------------------------------
control <- trainControl(method = "cv", number = 5)
tunegrid <- expand.grid(.mtry = c(3, 4, 5, 6, 7))

rf_enhanced <- train(
  total_line_cost ~ .,
  data = train_data,
  method = "rf",
  metric = "RMSE",
  tuneGrid = tunegrid,
  trControl = control,
  ntree = 500,
  importance = TRUE
)

# -------------------------------------
#  Predict and Evaluate
# -------------------------------------
rf_pred_enh <- predict(rf_enhanced, test_data)
rf_rmse_enh <- RMSE(rf_pred_enh, test_data$total_line_cost)
rf_r2_enh <- R2(rf_pred_enh, test_data$total_line_cost)

cat("Enhanced Random Forest Regression Results\n")

cat("Best mtry:", rf_enhanced$bestTune$mtry, "\n")

cat("RMSE:", round(rf_rmse_enh, 2), "\n")

cat("R-squared:", round(rf_r2_enh, 4), "\n")

# -------------------------------------
#  Plot Feature Importance
# -------------------------------------
varImpPlot(rf_enhanced$finalModel, main = "Enhanced Variable Importance - Random Forest")

# -------------------------------------
#  Prepare & Scale Data
# -------------------------------------
model_data_nn <- fleet_enhanced %>%
  select(total_line_cost, unit_age, meter, annual_mileage, average_miles_per_month,
         months_with_mileage, utilized_days_per_month, total_jobs, avg_line_cost, is_active,
         jobs_per_year, mileage_per_age, age_x_jobs, age_x_mileage) %>%
  na.omit()

# Normalize numeric variables
predictors <- names(model_data_nn)[names(model_data_nn) != "total_line_cost"]
model_data_nn[predictors] <- scale(model_data_nn[predictors])

# Train/test split
set.seed(123)
train_index <- createDataPartition(model_data_nn$total_line_cost, p = 0.8, list = FALSE)
train_data <- model_data_nn[train_index, ]
test_data <- model_data_nn[-train_index, ]

# -------------------------------------
#  Train Neural Network Model
# -------------------------------------
set.seed(9999999)
nn_model <- nnet(
  total_line_cost ~ .,
  data = train_data,
  size = 5,         # hidden neurons
  decay = 0.01,     # regularization
  linout = TRUE,    # for regression
  trace = FALSE
)

# -------------------------------------
#  Predict and Evaluate
# -------------------------------------
nn_pred <- predict(nn_model, test_data)
nn_rmse <- RMSE(nn_pred, test_data$total_line_cost)
nn_r2 <- R2(nn_pred, test_data$total_line_cost)

cat("Neural Network Regression Results\n")

cat("RMSE:", round(nn_rmse, 2), "\n")

cat("R-squared:", round(nn_r2, 4), "\n")

# -- Model Commentary --
cat("Note: Neural Network R² (0.62) is significantly lower than tree-based models.
It is retained here for benchmarking purposes only and is not recommended for production.")

# -------------------------------------
# Creating Risk Flag (Top 10% Cost)
# -------------------------------------
model_data_clf <- fleet_enhanced %>%
  select(total_line_cost, unit_age, meter, annual_mileage, average_miles_per_month,
         months_with_mileage, utilized_days_per_month, total_jobs, avg_line_cost, is_active,
         jobs_per_year, mileage_per_age, age_x_jobs, age_x_mileage) %>%
  na.omit()

# Define high-risk threshold (90th percentile)
cost_threshold <- quantile(model_data_clf$total_line_cost, 0.90)

# Create binary risk flag
model_data_clf <- model_data_clf %>%
  mutate(risk_flag = ifelse(total_line_cost >= cost_threshold, 1, 0)) %>%
  select(-total_line_cost)  # drop continuous target

# -------------------------------------
#  Train/Test Split
# -------------------------------------
set.seed(9999999)
train_index <- createDataPartition(model_data_clf$risk_flag, p = 0.8, list = FALSE)
train_data <- model_data_clf[train_index, ]
test_data <- model_data_clf[-train_index, ]

# -------------------------------------
#  Train Logistic Regression
# -------------------------------------
logit_model <- glm(risk_flag ~ ., data = train_data, family = "binomial")

# Predict probabilities and convert to class
logit_prob <- predict(logit_model, test_data, type = "response")
logit_pred <- ifelse(logit_prob > 0.5, 1, 0)

# -------------------------------------
#  Evaluate Model
# -------------------------------------
conf_mat <- confusionMatrix(as.factor(logit_pred), as.factor(test_data$risk_flag))

cat(" Logistic Regression Classification Results\n")

print(conf_mat)

# -------------------------------------
#  Prepare Data (reuse risk_flag data set)
# -------------------------------------
model_data_rf_clf <- fleet_enhanced %>%
  select(total_line_cost, unit_age, meter, annual_mileage, average_miles_per_month,
         months_with_mileage, utilized_days_per_month, total_jobs, avg_line_cost, is_active,
         jobs_per_year, mileage_per_age, age_x_jobs, age_x_mileage) %>%
  na.omit()

# Define high-risk threshold (top 10%)
cost_threshold <- quantile(model_data_rf_clf$total_line_cost, 0.90)

# Add binary flag
model_data_rf_clf <- model_data_rf_clf %>%
  mutate(risk_flag = ifelse(total_line_cost >= cost_threshold, 1, 0)) %>%
  select(-total_line_cost)

# -------------------------------------
#  Train/Test Split
# -------------------------------------
set.seed(123)
train_index <- createDataPartition(model_data_rf_clf$risk_flag, p = 0.8, list = FALSE)
train_data <- model_data_rf_clf[train_index, ]
test_data <- model_data_rf_clf[-train_index, ]

# -------------------------------------
#  Train Random Forest Classifier
# -------------------------------------
set.seed(9999999)
rf_classifier <- randomForest(
  as.factor(risk_flag) ~ .,
  data = train_data,
  ntree = 500,
  mtry = 4,
  importance = TRUE
)

# -------------------------------------
#  Predict and Evaluate
# -------------------------------------
rf_pred <- predict(rf_classifier, test_data)
conf_mat_rf <- confusionMatrix(rf_pred, as.factor(test_data$risk_flag))

cat("Random Forest Classifier Results\n")

print(conf_mat_rf)

# -------------------------------------
#  Feature Importance Plot
# -------------------------------------
varImpPlot(rf_classifier, main = "Variable Importance - Random Forest Classifier")

# -------------------------------------
#  Prepare Data
# -------------------------------------
model_data_xgb <- fleet_enhanced %>%
  select(total_line_cost, unit_age, meter, annual_mileage, average_miles_per_month,
         months_with_mileage, utilized_days_per_month, total_jobs, avg_line_cost, is_active,
         jobs_per_year, mileage_per_age, age_x_jobs, age_x_mileage) %>%
  na.omit()

set.seed(123)
train_index <- createDataPartition(model_data_xgb$total_line_cost, p = 0.8, list = FALSE)
train_data <- model_data_xgb[train_index, ]
test_data <- model_data_xgb[-train_index, ]

# -------------------------------------
#  Train XGBoost using caret
# -------------------------------------
xgb_grid <- expand.grid(
  nrounds = 200,
  max_depth = 6,
  eta = 0.1,
  gamma = 0,
  colsample_bytree = 0.8,
  min_child_weight = 1,
  subsample = 0.8
)

xgb_control <- trainControl(method = "cv", number = 5, verboseIter = FALSE)

set.seed(9999999)
xgb_model <- train(
  total_line_cost ~ .,
  data = train_data,
  method = "xgbTree",
  trControl = xgb_control,
  tuneGrid = xgb_grid,
  metric = "RMSE"
)

# -------------------------------------
#  Predict and Evaluate
# -------------------------------------
xgb_pred <- predict(xgb_model, test_data)
xgb_rmse <- RMSE(xgb_pred, test_data$total_line_cost)
xgb_r2 <- R2(xgb_pred, test_data$total_line_cost)

cat(" XGBoost Regression Results\n")

cat("RMSE:", round(xgb_rmse, 2), "\n")

cat("R-squared:", round(xgb_r2, 4), "\n")

# Extract trained xgboost model from caret
xgb_final <- xgb_model$finalModel

# Get and plot feature importance
importance_matrix <- xgb.importance(model = xgb_final)

# Plot top 10 important features
xgb.plot.importance(importance_matrix,
                    top_n = 10,
                    rel_to_first = TRUE,
                    xlab = "Relative Importance",
                    main = "XGBoost Feature Importance")

# --- Summary Table of All Models ---
model_performance <- tibble::tibble(
  Model = c("Linear Regression", "Log-Linear Regression", "Random Forest", "Enhanced RF", "XGBoost", "Neural Net"),
  RMSE = c(lm_rmse, log_rmse, rf_rmse, rf_rmse_enh, xgb_rmse, nn_rmse),
  R2   = c(lm_r2, log_r2, rf_r2, rf_r2_enh, xgb_r2, nn_r2)
)
print(model_performance)

# --- Final Predictions Export for Tableau ---

final_predictions <- fleet_enhanced %>%
  mutate(
    xgb_predicted_cost = predict(xgb_model, newdata = .),
    rf_risk_flag = predict(rf_classifier, newdata = ., type = "response")
  )

write.csv(final_predictions, "fleet_predictions_export.csv", row.names = FALSE) 






