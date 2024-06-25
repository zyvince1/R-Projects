library(openxlsx)
library(tis)
library(airutilities)
library(rjdhf)
library(sautilities)
library(dplyr)

setwd("your directory")

# Function to read and preprocess the data
read_preprocess_data <- function(filepath, sheet) {
  df <- openxlsx::read.xlsx(filepath, sheet = sheet)
  na.omit(df)
}

# Function to create tis objects
create_tis <- function(data, start_date) {
  tis(data, start = start_date, tif = "wsaturday")
}

# Function to generate holiday regressors
generate_holiday_regressors <- function(week_tis, year_tis) {
  holidays <- list(
    newyear = c(8, 1, c(0, 0, 0, 0, 0, 0, 1, 1)),
    mlk = c(1, 1, 1),
    president = c(1, 1, 1),
    easter = c(8, 8, c(1, 0, 0, 0, 0, 0, 0, 0)),
    memorial = c(1, 1, 1),
    july4 = c(1, 1, 1),
    labor = c(2, 2, c(0, 1)),
    columbus = c(1, 1, 1),
    veteran = c(1, 1, 1),
    thanksgiving = c(1, 1, 1)
  )
  
  holiday_regressors <- lapply(names(holidays), function(holiday) {
    params <- holidays[[holiday]]
    airutilities::gen_movereg_holiday(
      hol_n = params[1],
      hol_index = params[2],
      hol_wt = ifelse(length(params) == 3, params[3], array(params[3], dim = 1)),
      hol_type = holiday,
      this_week = week_tis,
      this_year = year_tis
    )
  })
  
  do.call(cbind, holiday_regressors)
}

# Function to generate special holiday regressors
generate_special_holiday_regressors <- function(week_tis) {
  july4_wed <- airutilities::match_month_day(week_tis, "0707")
  xmas_w53 <- airutilities::match_week(week_tis, 53)
  xmas_fri <- airutilities::match_month_day(week_tis, "1226")
  
  cbind(july4_wed, xmas_w53, xmas_fri)
}

# Function to generate the final regression matrix
generate_regression_matrix <- function(holiday_matrix, week_tis, year_tis) {
  tc_date <- matrix(c(13, 2020), ncol = 2, byrow = TRUE)
  tc_matrix <- airutilities::gen_tc_outlier_matrix(tc_date, week_tis, year_tis, 0)
  
  final_matrix <- cbind(holiday_matrix, tc_matrix)
  colnames(final_matrix) <- c(colnames(holiday_matrix), colnames(tc_matrix))
  
  final_matrix
}

# Main processing function
process_series <- function(data, start_date, end_date, holiday_matrix, week_tis, year_tis) {
  tis_data <- create_tis(data, start_date)
  tis_final <- window(tis_data, start = start_date, end = end_date)
  
  ljung_cv <- sautilities::set_critical_value(length(tis_final), cv_alpha = 0.005)
  est <- rjdhf::fractionalAirlineEstimation(tis_final, periods = c(365.25 / 7), x = holiday_matrix, outliers = c("ao", "ls"), criticalValue = ljung_cv)
  decomp <- rjdhf::fractionalAirlineDecomposition(est$model$linearized, 365.25 / 7, stde = TRUE)
  
  model_matrix <- airutilities::gen_air_model_matrix(est, xreg_names = colnames(holiday_matrix), this_week = week_tis, this_year = year_tis)
  row_names <- rownames(model_matrix)
  otl_index <- sort(c(grep("ao", tolower(substr(row_names, 1, 2))), grep("ls", tolower(substr(row_names, 1, 2))), grep("tc", tolower(substr(row_names, 1, 2)))))
  xtype <- c(rep("hol", ncol(holiday_matrix)), tolower(substr(row_names, 1, 2))[otl_index])
  
  comp <- airutilities::gen_air_components(est, decomp, this_xtype = xtype, this_log = FALSE, this_stde = TRUE)
  comp_tis <- lapply(comp, function(x) try(tis::tis(x, start = start_date, tif = "wsaturday")))
  
  as.numeric(comp_tis$sa)
}

# Main script
ic_df <- read_preprocess_data("icstate_raw.xlsx", "df")
uihis_start <- c(1987, 1)
this_start <- c(1987, 1)
### manually change this_end to latest!
this_end <- c(2024, 9)

ic_week_tis <- create_tis(ic_df$week, uihis_start)
ic_year_tis <- create_tis(ic_df$year, uihis_start)
holiday_matrix <- generate_holiday_regressors(ic_week_tis, ic_year_tis)
special_holiday_matrix <- generate_special_holiday_regressors(ic_week_tis)
holiday_matrix <- cbind(holiday_matrix, special_holiday_matrix)
colnames(holiday_matrix) <- c("ny", "mlk", "president", "easter", "memorial", "july4", "labor", "columbus", "veteran", "thanksgiving", "july4_wed", "xmas_w53", "xmas_fri")

regression_matrix <- generate_regression_matrix(holiday_matrix, ic_week_tis, ic_year_tis)

df <- data.frame(ic_df[, 1])
colnames(df)[1] <- "Dates"

for (i in 4:ncol(ic_df)) {
  sa_values <- process_series(ic_df[, i], uihis_start, this_start, this_end, regression_matrix, ic_week_tis, ic_year_tis)
  df1 <- data_frame(sa_values)
  colnames(df1)[1] <- colnames(ic_df)[i]
  df <- cbind(df, df1)
}

write.xlsx(df, file = "icstate_sa.xlsx")
