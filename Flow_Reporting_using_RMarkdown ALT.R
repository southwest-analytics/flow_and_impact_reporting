# ═════════════════════════════════════════════
# * 0. Load libraries and define functions ----
# ═════════════════════════════════════════════
library(tidyverse)
library(readxl)
library(ini)
library(flextable)
library(officer)
library(officedown)
library(conflicted)

# Define parameters
year_start <- '2024-04-01'
current_month <- '2024-05-01'

# Flextable borders
std_border <- fp_border(color = "gray")

# Glenday Palette
palGlenday <- scales::col_factor(
  palette = c('Green' = 'darkgreen', 'Yellow' = 'gold', 'Blue' = 'royalblue', 'Red' = 'red3'),
  levels = c('Green', 'Yellow', 'Blue', 'Red'))

# Glenday Sieve
fnGlenday <- function(df, var.x, var.y){
  df <- df %>%
    dplyr::filter(get(var.y) > 0) %>%
    arrange(desc(get(var.y)), get(var.x)) %>%
    mutate(pct = get(var.y)/ sum(get(var.y)),
           cumpct = cumsum(pct),
           lag_cumpct = dplyr::lag(cumpct),
           glenday = if_else(cumpct <= 0.5 | (lag_cumpct < 0.5 & cumpct > 0.5),
                             'Green',
                             if_else(cumpct <= 0.95 | (lag_cumpct < 0.95 & cumpct > 0.95),
                                     'Yellow',
                                     if_else(cumpct <= 0.99 | (lag_cumpct < 0.99 & cumpct > 0.99),
                                             'Blue','Red')))) %>%
    select(-lag_cumpct) %>%
    # Replace any NA values in glenday as Green
    replace_na(replace = list(glenday = 'Green'))
  return(df)
}

# New clients supported - Ignore those that are in the exclusion list i.e "Declined", "No engagement"
fnNewClientsSupported <- function(df, period_start, period_end){
  df_report <- df_caseload_tracker %>% 
    dplyr::filter(ct_start >= period_start & ct_start < period_end) %>%
    dplyr::filter(!(ct_closure %in% unlist(unname(ini_file_settings$caseload_tracker_exclusions))))
  return(df_report)  
}

# All clients supported - Ignore those that are in the exclusion list i.e "Declined", "No engagement"
fnClientsSupported <- function(df, period_start, period_end){
  df_report <- df_caseload_tracker %>% 
    dplyr::filter(ct_start < period_end & (ct_end >= period_start | is.na(ct_end))) %>%
    dplyr::filter(!(ct_closure %in% unlist(unname(ini_file_settings$caseload_tracker_exclusions))))
  return(df_report)  
}

# Total new clients - simple count of new clients (ignoring exclusions)
fnTotalNewClients <- function(df, period_start, period_end){
  df_report <- df_caseload_tracker %>% 
    dplyr::filter(ct_start >= period_start & ct_start < period_end)
  return(df_report)
}

# Caseload - simple count of clients open at the end of the period
fnCaseload <- function(df, period_start, period_end){
  df_report <- df_caseload_tracker %>% 
    dplyr::filter(ct_start <= period_end & ( is.na(ct_end) | (ct_end >= period_end)))
  return(df_report)
}

# Clients at end of support - Ignore those that are in the exclusion list i.e "Declined", "No engagement"
fnClientsSupportedEnd <- function(df, period_start, period_end){
  df_report <- df_caseload_tracker %>% 
    dplyr::filter(ct_end >= period_start & ct_end < period_end) %>%
    dplyr::filter(!(ct_closure %in% unlist(unname(ini_file_settings$caseload_tracker_exclusions))))
  return(df_report)
}

# Read in the initialisation file settings
ini_file_settings <- read.ini('flow_reporting.ini')

# Time periods
dt_year_start <- as.Date(year_start)
dt_current_month <- as.Date(current_month)

# ────────────────────────────────
# * 0.1. Import the workbooks ----
# ────────────────────────────────

# Initialise data frames to receive the different sections of the workbooks
df_data_points <- data.frame()
df_support_and_referrals <- data.frame()
df_outputs <- data.frame()
df_outcomes <- data.frame()

# Loop through each of the reporting workbooks
for(f in ini_file_settings$reporting_workbooks){
  # ────────────────────────────────────────────────
  # * * 0.1.1. Import the data points worksheet ----
  # ────────────────────────────────────────────────
  
  # Only import the first 3 columns of the sheet and set the variable type
  df_tmp <- read_excel(path = f, 
                       sheet = 'Data Points', 
                       range = cell_cols('A:C'), 
                       col_types = c('date','text','numeric')) %>%
    # Rename the column names as previously these have overwritten by caseworkers
    rename_with(.fn = ~c('month', 'metric', 'value')) %>% 
    # Remove any rows that have no month or no value
    dplyr::filter(!is.na(month) & !is.na(value)) %>%
    # Add the source filename to the data
    mutate(source = basename(file.path(f)))
  
  # Bind to the combined data frame
  df_data_points <- df_data_points %>% bind_rows(df_tmp)
  
  # ──────────────────────────────────────────────────────────
  # * * 0.1.2. Import the support and referrals worksheet ----
  # ──────────────────────────────────────────────────────────
  
  # Only import the first 4 columns of the sheet and set the variable type
  df_tmp <- read_excel(path = f, 
                       sheet = 'Support and Referrals', 
                       range = cell_cols('A:D'), 
                       col_types = c('text','date','text','text')) %>%
    # Rename the column names as previously these have overwritten by caseworkers
    rename_with(.fn = ~c('client_id', 'month', 'section', 'support')) %>% 
    # Remove any rows that have an NA in column
    dplyr::filter( !(is.na(client_id) | is.na(month) | is.na(section) | is.na(support))) %>%
    # Add the source filename to the data
    mutate(source = basename(file.path(f)))
  
  # Bind to the combined data frame
  df_support_and_referrals <- df_support_and_referrals %>% bind_rows(df_tmp)
  
  # ────────────────────────────────────────────
  # * * 0.1.3. Import the outputs worksheet ----
  # ────────────────────────────────────────────
  
  # Only import the first 4 columns of the sheet and set the variable type
  df_tmp <- read_excel(path = f, 
                       sheet = 'Outputs', 
                       range = cell_cols('A:D'), 
                       col_types = c('text','date','text','text')) %>%
    # Rename the column names as previously these have overwritten by caseworkers
    rename_with(.fn = ~c('client_id', 'month', 'section', 'output')) %>% 
    # Remove any rows that have an NA in column
    dplyr::filter( !(is.na(client_id) | is.na(month) | is.na(section) | is.na(output))) %>%
    # Add the source filename to the data
    mutate(source = basename(file.path(f)))
  
  # Bind to the combined data frame
  df_outputs <- df_outputs %>% bind_rows(df_tmp)
  
  # ─────────────────────────────────────────────
  # * * 0.1.4. Import the outcomes worksheet ----
  # ─────────────────────────────────────────────
  
  # Only import the first 4 columns of the sheet and set the variable type
  df_tmp <- read_excel(path = f, 
                       sheet = 'Outcomes', 
                       range = cell_cols('A:D'), 
                       col_types = c('text','date','text','text')) %>%
    # Rename the column names as previously these have overwritten by caseworkers
    rename_with(.fn = ~c('client_id', 'month', 'section', 'outcome')) %>% 
    # Remove any rows that have an NA in column
    dplyr::filter( !(is.na(client_id) | is.na(month) | is.na(section) | is.na(outcome))) %>%
    # Add the source filename to the data
    mutate(source = basename(file.path(f)))
  
  # Bind to the combined data frame
  df_outcomes <- df_outcomes %>% bind_rows(df_tmp)
}

# ──────────────────────────────────────────────
# * 0.1. Import caseload tracker worksheets ----
# ──────────────────────────────────────────────

# Initialise the data frame to receive the caseload tracker data
df_caseload_tracker <- data.frame()

# Get the worksheet names from the ini file settings
caseload_tracker_worksheets <- unname(unlist(ini_file_settings$caseload_tracker_sheets))

# Get the required field numbers from the ini file
caseload_tracker_field_numbers <- c(
  unlist(as.integer(unname(ini_file_settings$caseload_tracker_demographics))),
  unlist(as.integer(unname(ini_file_settings$caseload_tracker_numbers_supported))),
  # We will no longer use the caseload tracker for the activity numbers but will use the 
  # ED|EM|AMB data sheet supplied by RDUH BI Team
  # unlist(as.integer(unname(ini_file_settings$caseload_tracker_activity))),
  unlist(as.integer(unname(ini_file_settings$caseload_tracker_wemwbs_goals))))

# Get the required field names from the ini file
caseload_tracker_field_names <- c(
  unlist(names(ini_file_settings$caseload_tracker_demographics)),
  unlist(names(ini_file_settings$caseload_tracker_numbers_supported)),
  # We will no longer use the caseload tracker for the activity numbers but will use the 
  # ED|EM|AMB data sheet supplied by RDUH BI Team
  # unlist(names(ini_file_settings$caseload_tracker_activity)),
  unlist(names(ini_file_settings$caseload_tracker_wemwbs_goals)))


# Loop through each of the caseload tracker worksheets
for(s in caseload_tracker_worksheets){
  # Read in all bar the first two lines of the sheets (the headers)
  df_tmp <- read_excel(path = ini_file_settings$caseload_tracker$caseload_tracker, 
                       sheet = s,
                       col_type = 'text', 
                       col_names = FALSE,
                       skip = 3,
                       ) %>% 
    # Select the required fields and rename them
    select(all_of(caseload_tracker_field_numbers)) %>%
    rename_with(.fn = ~caseload_tracker_field_names) %>%
    # Ignore any rows that don't have an ID
    dplyr::filter(!is.na(ct_id))
  
  # Bind to the combined data frame
  df_caseload_tracker <- df_caseload_tracker %>% bind_rows(df_tmp)
}

# Format the columns to required format
df_caseload_tracker <- df_caseload_tracker %>% 
  mutate(ct_age = as.integer(ct_age),
         ct_status = as.factor(ct_status),
         ct_closure = as.factor(ct_closure),
         ct_start = as.Date(as.integer(ct_start), origin = '1899-12-30'),
         ct_end = as.Date(as.integer(ct_end), origin = '1899-12-30'))
  # We no longer need this as the ED|EM|AMB activity will come from the RDUH BI team sheet 
  # %>%
  # mutate(across(.cols = all_of(c(unlist(names(ini_file_settings$caseload_tracker_activity)),
  #                         unlist(names(ini_file_settings$caseload_tracker_wemwbs_goals)))),
  #               .fns = as.integer))

save(list = c('df_data_points', 'df_support_and_referrals', 'df_outputs', 'df_outcomes'),
     file = paste0('data_', format(dt_current_month, '%Y%m%d'), '.RObj'))

# ═════════════════════════════════
# 1. Process markdown document ----
# ═════════════════════════════════

# ─────────────────────────────
# * 1.1. Headlines section ----
# ─────────────────────────────
df_RPT_01_headlines <- data.frame(metric = c('New Clients Supported in Period',
                                             'Clients Supported in Period',
                                             'Previous 12 Month ED Attendances for Clients Supported in Period',
                                             'Previous 12 Month Emergency Admissions for Clients Supported in Period',
                                             'Previous 12 Month Ambulance Conveyances for Clients Supported in Period',
                                             'Previous 3 Month ED Attendances for Clients Supported in Period',
                                             'Previous 3 Month Emergency Admissions for Clients Supported in Period',
                                             'Previous 3 Month Ambulance Conveyances for Clients Supported in Period'),
                                  current_month = as.integer(rep(NA, 8)),
                                  q1 = as.integer(rep(NA, 8)),
                                  q2 = as.integer(rep(NA, 8)),
                                  q3 = as.integer(rep(NA, 8)),
                                  q4 = as.integer(rep(NA, 8)),
                                  ytd = as.integer(rep(NA, 8)))

# * * 1.1.1. Number of new people supported in the period ----
df_RPT_01_headlines[1, 2] <- fnNewClientsSupported(df = df_caseload_tracker, 
                                              period_start = dt_current_month, 
                                              period_end = dt_current_month + months(1)) %>% NROW()
df_RPT_01_headlines[1, 3] <- ifelse(dt_current_month >= dt_year_start,
                                    fnNewClientsSupported(df = df_caseload_tracker, 
                                                          period_start = dt_year_start, 
                                                          period_end = dt_year_start + months(3)) %>% NROW(),
                                    NA)
df_RPT_01_headlines[1, 4] <- ifelse(dt_current_month >= dt_year_start + months(3),
                                    fnNewClientsSupported(df = df_caseload_tracker, 
                                                          period_start = dt_year_start + months(3), 
                                                          period_end = dt_year_start + months(6)) %>% NROW(),
                                    NA)
df_RPT_01_headlines[1, 5] <- ifelse(dt_current_month >= dt_year_start + months(6),
                                    fnNewClientsSupported(df = df_caseload_tracker, 
                                                          period_start = dt_year_start + months(6), 
                                                          period_end = dt_year_start + months(9)) %>% NROW(),
                                    NA)
df_RPT_01_headlines[1, 6] <- ifelse(dt_current_month >= dt_year_start + months(9),
                                    fnNewClientsSupported(df = df_caseload_tracker, 
                                                          period_start = dt_year_start + months(9), 
                                                          period_end = dt_year_start + months(12)) %>% NROW(),
                                    NA)
df_RPT_01_headlines[1, 7] <-  fnNewClientsSupported(df = df_caseload_tracker, 
                                                    period_start = dt_year_start, 
                                                    period_end = dt_year_start + months(12)) %>% NROW()

# * * 1.1.2. Number of people supported in the period ----
df_RPT_01_headlines[2, 2] <- fnClientsSupported(df = df_caseload_tracker, 
                                                period_start = dt_current_month, 
                                                period_end = dt_current_month + months(1)) %>% NROW()
df_RPT_01_headlines[2, 3] <- ifelse(dt_current_month >= dt_year_start,
                                    fnClientsSupported(df = df_caseload_tracker, 
                                                       period_start = dt_year_start, 
                                                       period_end = dt_year_start + months(3)) %>% NROW(),
                                    NA)
df_RPT_01_headlines[2, 4] <- ifelse(dt_current_month >= dt_year_start + months(3),
                                    fnClientsSupported(df = df_caseload_tracker, 
                                                       period_start = dt_year_start + months(3), 
                                                       period_end = dt_year_start + months(6)) %>% NROW(),
                                    NA)
df_RPT_01_headlines[2, 5] <- ifelse(dt_current_month >= dt_year_start + months(6),
                                    fnClientsSupported(df = df_caseload_tracker, 
                                                       period_start = dt_year_start + months(6), 
                                                       period_end = dt_year_start + months(9)) %>% NROW(),
                                    NA)
df_RPT_01_headlines[2, 6] <- ifelse(dt_current_month >= dt_year_start + months(9),
                                    fnClientsSupported(df = df_caseload_tracker, 
                                                       period_start = dt_year_start + months(9), 
                                                       period_end = dt_year_start + months(12)) %>% NROW(),
                                    NA)
df_RPT_01_headlines[2, 7] <- fnClientsSupported(df = df_caseload_tracker, 
                                                period_start = dt_year_start, 
                                                period_end = dt_year_start + months(12)) %>% NROW()


# We will replace this section with the data from the RDUH BI Team data sheet
# # * * 1.1.3. Emergency department, Emergency Admission and Ambulance conveyances count of people supported in the period ----
# df_tmp <- fnClientsSupported(df = df_caseload_tracker, 
#                              period_start = dt_current_month, 
#                              period_end = dt_current_month + months(1))
# df_RPT_01_headlines[3, 2] <- sum(df_tmp$ct_ed_12m, na.rm = TRUE)
# df_RPT_01_headlines[4, 2] <- sum(df_tmp$ct_em_12m, na.rm = TRUE)
# df_RPT_01_headlines[5, 2] <- sum(df_tmp$ct_amb_12m, na.rm = TRUE)
# df_RPT_01_headlines[6, 2] <- sum(df_tmp$ct_ed_3m, na.rm = TRUE)
# df_RPT_01_headlines[7, 2] <- sum(df_tmp$ct_em_3m, na.rm = TRUE)
# df_RPT_01_headlines[8, 2] <- sum(df_tmp$ct_amb_3m, na.rm = TRUE)
# 
# if(dt_current_month >= dt_year_start){
#   df_tmp <- fnClientsSupported(df = df_caseload_tracker, 
#                                period_start = dt_year_start, 
#                                period_end = dt_year_start + months(3))
#   df_RPT_01_headlines[3, 3] <- sum(df_tmp$ct_ed_12m, na.rm = TRUE)
#   df_RPT_01_headlines[4, 3] <- sum(df_tmp$ct_em_12m, na.rm = TRUE)
#   df_RPT_01_headlines[5, 3] <- sum(df_tmp$ct_amb_12m, na.rm = TRUE)
#   df_RPT_01_headlines[6, 3] <- sum(df_tmp$ct_ed_3m, na.rm = TRUE)
#   df_RPT_01_headlines[7, 3] <- sum(df_tmp$ct_em_3m, na.rm = TRUE)
#   df_RPT_01_headlines[8, 3] <- sum(df_tmp$ct_amb_3m, na.rm = TRUE)
# }
# 
# if(dt_current_month >= dt_year_start + months(3)){
#   df_tmp <- fnClientsSupported(df = df_caseload_tracker, 
#                                period_start = dt_year_start + months(3), 
#                                period_end = dt_year_start + months(6))
#   df_RPT_01_headlines[3, 4] <- sum(df_tmp$ct_ed_12m, na.rm = TRUE)
#   df_RPT_01_headlines[4, 4] <- sum(df_tmp$ct_em_12m, na.rm = TRUE)
#   df_RPT_01_headlines[5, 4] <- sum(df_tmp$ct_amb_12m, na.rm = TRUE)
#   df_RPT_01_headlines[6, 4] <- sum(df_tmp$ct_ed_3m, na.rm = TRUE)
#   df_RPT_01_headlines[7, 4] <- sum(df_tmp$ct_em_3m, na.rm = TRUE)
#   df_RPT_01_headlines[8, 4] <- sum(df_tmp$ct_amb_3m, na.rm = TRUE)
# }
# 
# if(dt_current_month >= dt_year_start + months(6)){
#   df_tmp <- fnClientsSupported(df = df_caseload_tracker, 
#                                period_start = dt_year_start + months(6), 
#                                period_end = dt_year_start + months(9))
#   df_RPT_01_headlines[3, 5] <- sum(df_tmp$ct_ed_12m, na.rm = TRUE)
#   df_RPT_01_headlines[4, 5] <- sum(df_tmp$ct_em_12m, na.rm = TRUE)
#   df_RPT_01_headlines[5, 5] <- sum(df_tmp$ct_amb_12m, na.rm = TRUE)
#   df_RPT_01_headlines[6, 5] <- sum(df_tmp$ct_ed_3m, na.rm = TRUE)
#   df_RPT_01_headlines[7, 5] <- sum(df_tmp$ct_em_3m, na.rm = TRUE)
#   df_RPT_01_headlines[8, 5] <- sum(df_tmp$ct_amb_3m, na.rm = TRUE)
# }
# 
# if(dt_current_month >= dt_year_start + months(9)){
#   df_tmp <- fnClientsSupported(df = df_caseload_tracker, 
#                                period_start = dt_year_start + months(9), 
#                                period_end = dt_year_start + months(12))
#   df_RPT_01_headlines[3, 6] <- sum(df_tmp$ct_ed_12m, na.rm = TRUE)
#   df_RPT_01_headlines[4, 6] <- sum(df_tmp$ct_em_12m, na.rm = TRUE)
#   df_RPT_01_headlines[5, 6] <- sum(df_tmp$ct_amb_12m, na.rm = TRUE)
#   df_RPT_01_headlines[6, 6] <- sum(df_tmp$ct_ed_3m, na.rm = TRUE)
#   df_RPT_01_headlines[7, 6] <- sum(df_tmp$ct_em_3m, na.rm = TRUE)
#   df_RPT_01_headlines[8, 6] <- sum(df_tmp$ct_amb_3m, na.rm = TRUE)
# }
# 
# number_YTD <- fnClientsSupported(df = df_caseload_tracker, 
#                                  period_start = dt_year_start, 
#                                  period_end = dt_year_start + months(12)) %>% NROW()
# 
# df_tmp <- fnClientsSupported(df = df_caseload_tracker, 
#                              period_start = dt_year_start, 
#                              period_end = dt_year_start + months(12))
# df_RPT_01_headlines[3, 7] <- sum(df_tmp$ct_ed_12m, na.rm = TRUE)
# df_RPT_01_headlines[4, 7] <- sum(df_tmp$ct_em_12m, na.rm = TRUE)
# df_RPT_01_headlines[5, 7] <- sum(df_tmp$ct_amb_12m, na.rm = TRUE)
# df_RPT_01_headlines[6, 7] <- sum(df_tmp$ct_ed_3m, na.rm = TRUE)
# df_RPT_01_headlines[7, 7] <- sum(df_tmp$ct_em_3m, na.rm = TRUE)
# df_RPT_01_headlines[8, 7] <- sum(df_tmp$ct_amb_3m, na.rm = TRUE)

df_RPT_01_headlines

# ──────────────────────────────────────────────────────────────────
# * 1.2. Changes in activity / tracking - NHSE/OND KPIs section ----
# ──────────────────────────────────────────────────────────────────

# This first section will need to be re-written when the ED|EM|AMB data sheet is
# provided by RDUH BI team
df_RPT_02_changes_in_activity <- df_RPT_01_headlines %>% dplyr::filter(metric=='New Clients Supported in Period') %>%
  bind_rows(
    data.frame(metric = c('Reduction in ED attendances starting 3 months from intervention beginning (NHSE Target 40%)',
                          'Reduction in non elective admissions 3 months from intervention beginning (NHSE Target 40%)',
                          'Reduction in ambulance conveyances 3 months from intervention beginning (No NHSE Target)',
                          'Reduction in ED attendances 12 month prior to 12 months post (OND Target 40%)',
                          'Reduction in Non elective admissions 12 months prior to 12 months post. (OND Target 40%)',
                          'Reduction in ambulance conveyances 12 months from intervention beginning (No NHSE Target)',
                          'Reduction in people experiencing loneliness at the end of support (NHSE Target 66%)',
                          'Clients Ending Support', 
                          'Clients ending Support and Experiencing Improved Wellbeing',
                          'People report a positive experience from our support (NHSE Target 80%)',
                          'People progress at least one goal (OND Target 90%)'),
               current_month = rep(NA, 11),
               q1 = rep(NA, 11),
               q2 = rep(NA, 11),
               q3 = rep(NA, 11),
               q4 = rep(NA, 11),
               ytd = rep(NA, 11)))

# * * 1.2.1 Clients at end of support ----
df_tmp <- fnClientsSupportedEnd(df = df_caseload_tracker, 
                                period_start = dt_current_month, 
                                period_end = dt_current_month + months(1))
df_RPT_02_changes_in_activity[9, 2] <- df_tmp %>% NROW()
df_RPT_02_changes_in_activity[10, 2] <- sum((df_tmp$ct_wemwbs_score_out > df_tmp$ct_wemwbs_score_in ) * 1, na.rm = TRUE)
df_RPT_02_changes_in_activity[12, 2] <- sum(
  (
    (df_tmp$ct_goals_mh_out > df_tmp$ct_goals_mh_in) | 
    (df_tmp$ct_goals_ph_out > df_tmp$ct_goals_ph_in) | 
    (df_tmp$ct_goals_housing_out > df_tmp$ct_goals_housing_in) | 
    (df_tmp$ct_goals_diet_out > df_tmp$ct_goals_diet_in) | 
    (df_tmp$ct_goals_relations_out > df_tmp$ct_goals_relations_in) | 
    (df_tmp$ct_goals_money_out > df_tmp$ct_goals_money_in) | 
    (df_tmp$ct_goals_growth_out > df_tmp$ct_goals_growth_in) | 
    (df_tmp$ct_goals_community_out > df_tmp$ct_goals_community_in) | 
    (df_tmp$ct_goals_abuse_out > df_tmp$ct_goals_abuse_in) | 
    (df_tmp$ct_goals_children_out > df_tmp$ct_goals_children_in)
  ) * 1,
  na.rm = TRUE)

if(dt_current_month >= dt_year_start){
  df_tmp <- fnClientsSupportedEnd(df = df_caseload_tracker, 
                                  period_start = dt_year_start, 
                                  period_end = dt_year_start + months(3))
  df_RPT_02_changes_in_activity[9, 3] <- df_tmp %>% NROW()
  df_RPT_02_changes_in_activity[10, 3] <- sum((df_tmp$ct_wemwbs_score_out > df_tmp$ct_wemwbs_score_in ) * 1, na.rm = TRUE)
  df_RPT_02_changes_in_activity[12, 3] <- sum((
    (df_tmp$ct_goals_mh_out > df_tmp$ct_goals_mh_in) | (df_tmp$ct_goals_ph_out > df_tmp$ct_goals_ph_in) | (df_tmp$ct_goals_housing_out > df_tmp$ct_goals_housing_in) | 
      (df_tmp$ct_goals_diet_out > df_tmp$ct_goals_diet_in) | (df_tmp$ct_goals_relations_out > df_tmp$ct_goals_relations_in) | (df_tmp$ct_goals_money_out > df_tmp$ct_goals_money_in) | 
      (df_tmp$ct_goals_growth_out > df_tmp$ct_goals_growth_in) | (df_tmp$ct_goals_community_out > df_tmp$ct_goals_community_in) | (df_tmp$ct_goals_abuse_out > df_tmp$ct_goals_abuse_in) | 
      (df_tmp$ct_goals_children_out > df_tmp$ct_goals_children_in)) * 1, na.rm = TRUE)
}

if(dt_current_month >= dt_year_start + months(3)){
  df_tmp <- fnClientsSupportedEnd(df = df_caseload_tracker, 
                                  period_start = dt_year_start + months(3), 
                                  period_end = dt_year_start + months(6))
  df_RPT_02_changes_in_activity[9, 4] <- df_tmp %>% NROW()
  df_RPT_02_changes_in_activity[10, 4] <- sum((df_tmp$ct_wemwbs_score_out > df_tmp$ct_wemwbs_score_in ) * 1, na.rm = TRUE)
  df_RPT_02_changes_in_activity[12, 4] <- sum((
    (df_tmp$ct_goals_mh_out > df_tmp$ct_goals_mh_in) | (df_tmp$ct_goals_ph_out > df_tmp$ct_goals_ph_in) | (df_tmp$ct_goals_housing_out > df_tmp$ct_goals_housing_in) | 
      (df_tmp$ct_goals_diet_out > df_tmp$ct_goals_diet_in) | (df_tmp$ct_goals_relations_out > df_tmp$ct_goals_relations_in) | (df_tmp$ct_goals_money_out > df_tmp$ct_goals_money_in) | 
      (df_tmp$ct_goals_growth_out > df_tmp$ct_goals_growth_in) | (df_tmp$ct_goals_community_out > df_tmp$ct_goals_community_in) | (df_tmp$ct_goals_abuse_out > df_tmp$ct_goals_abuse_in) | 
      (df_tmp$ct_goals_children_out > df_tmp$ct_goals_children_in)) * 1, na.rm = TRUE)
}

if(dt_current_month >= dt_year_start + months(6)){
  df_tmp <- fnClientsSupportedEnd(df = df_caseload_tracker, 
                                  period_start = dt_year_start + months(6), 
                                  period_end = dt_year_start + months(9))
  df_RPT_02_changes_in_activity[9, 5] <- df_tmp %>% NROW()
  df_RPT_02_changes_in_activity[10, 5] <- sum((df_tmp$ct_wemwbs_score_out > df_tmp$ct_wemwbs_score_in ) * 1, na.rm = TRUE)
  df_RPT_02_changes_in_activity[12, 5] <- sum((
    (df_tmp$ct_goals_mh_out > df_tmp$ct_goals_mh_in) | (df_tmp$ct_goals_ph_out > df_tmp$ct_goals_ph_in) | (df_tmp$ct_goals_housing_out > df_tmp$ct_goals_housing_in) | 
      (df_tmp$ct_goals_diet_out > df_tmp$ct_goals_diet_in) | (df_tmp$ct_goals_relations_out > df_tmp$ct_goals_relations_in) | (df_tmp$ct_goals_money_out > df_tmp$ct_goals_money_in) | 
      (df_tmp$ct_goals_growth_out > df_tmp$ct_goals_growth_in) | (df_tmp$ct_goals_community_out > df_tmp$ct_goals_community_in) | (df_tmp$ct_goals_abuse_out > df_tmp$ct_goals_abuse_in) | 
      (df_tmp$ct_goals_children_out > df_tmp$ct_goals_children_in)) * 1, na.rm = TRUE)
}

if(dt_current_month >= dt_year_start + months(9)){
  df_tmp <- fnClientsSupportedEnd(df = df_caseload_tracker, 
                                  period_start = dt_year_start + months(9), 
                                  period_end = dt_year_start + months(12))
  df_RPT_02_changes_in_activity[9, 6] <- df_tmp %>% NROW()
  df_RPT_02_changes_in_activity[10, 6] <- sum((df_tmp$ct_wemwbs_score_out > df_tmp$ct_wemwbs_score_in ) * 1, na.rm = TRUE)
  df_RPT_02_changes_in_activity[12, 6] <- sum((
    (df_tmp$ct_goals_mh_out > df_tmp$ct_goals_mh_in) | (df_tmp$ct_goals_ph_out > df_tmp$ct_goals_ph_in) | (df_tmp$ct_goals_housing_out > df_tmp$ct_goals_housing_in) | 
      (df_tmp$ct_goals_diet_out > df_tmp$ct_goals_diet_in) | (df_tmp$ct_goals_relations_out > df_tmp$ct_goals_relations_in) | (df_tmp$ct_goals_money_out > df_tmp$ct_goals_money_in) | 
      (df_tmp$ct_goals_growth_out > df_tmp$ct_goals_growth_in) | (df_tmp$ct_goals_community_out > df_tmp$ct_goals_community_in) | (df_tmp$ct_goals_abuse_out > df_tmp$ct_goals_abuse_in) | 
      (df_tmp$ct_goals_children_out > df_tmp$ct_goals_children_in)) * 1, na.rm = TRUE)
}

df_tmp <- fnClientsSupportedEnd(df = df_caseload_tracker, 
                                period_start = dt_year_start, 
                                period_end = dt_year_start + months(12))
df_RPT_02_changes_in_activity[9, 7] <- df_tmp %>% NROW()
df_RPT_02_changes_in_activity[10, 7] <- sum((df_tmp$ct_wemwbs_score_out > df_tmp$ct_wemwbs_score_in ) * 1, na.rm = TRUE)
df_RPT_02_changes_in_activity[12, 7] <- sum((
  (df_tmp$ct_goals_mh_out > df_tmp$ct_goals_mh_in) | (df_tmp$ct_goals_ph_out > df_tmp$ct_goals_ph_in) | (df_tmp$ct_goals_housing_out > df_tmp$ct_goals_housing_in) | 
    (df_tmp$ct_goals_diet_out > df_tmp$ct_goals_diet_in) | (df_tmp$ct_goals_relations_out > df_tmp$ct_goals_relations_in) | (df_tmp$ct_goals_money_out > df_tmp$ct_goals_money_in) | 
    (df_tmp$ct_goals_growth_out > df_tmp$ct_goals_growth_in) | (df_tmp$ct_goals_community_out > df_tmp$ct_goals_community_in) | (df_tmp$ct_goals_abuse_out > df_tmp$ct_goals_abuse_in) | 
    (df_tmp$ct_goals_children_out > df_tmp$ct_goals_children_in)) * 1, na.rm = TRUE)

df_RPT_02_changes_in_activity

# ────────────────────────────────
# * 1.3. Process KPIS section ----
# ────────────────────────────────

df_RPT_03_process_kpis <- data.frame(metric = c('80% of new clients has a wemwebs baseline score',
                                                '80% exited client has a wemwbs exit score',
                                                '80% of clients have a Y/N loneliness question',
                                                '80% client has a loneliness exit score'),
                                     current_month = rep(NA, 4),
                                     q1 = rep(NA, 4),
                                     q2 = rep(NA, 4),
                                     q3 = rep(NA, 4),
                                     q4 = rep(NA, 4),
                                     ytd = rep(NA, 4)) 

# Baseline questions
# Current Month
df_tmp <- fnNewClientsSupported(df = df_caseload_tracker, 
                                period_start = dt_current_month, 
                                period_end = dt_current_month + months(1))
df_RPT_03_process_kpis[1, 2] <- sprintf('%.1f%% (%d/%d)', 
                                        (df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_in)) %>% NROW())/(df_tmp %>% NROW()) * 100,
                                        (df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_in)) %>% NROW()),
                                        (df_tmp %>% NROW()))
df_RPT_03_process_kpis[3, 2] <- sprintf('%.1f%% (%d/%d)', 
                                        (df_tmp %>% dplyr::filter(ct_loneliness %in% c('Yes','No')) %>% NROW())/(df_tmp %>% NROW()) * 100,
                                        (df_tmp %>% dplyr::filter(ct_loneliness %in% c('Yes','No')) %>% NROW()),
                                        (df_tmp %>% NROW()))
if(dt_current_month >= dt_year_start){
  # Quarter 1
  df_tmp <- fnNewClientsSupported(df = df_caseload_tracker, 
                                  period_start = dt_year_start, 
                                  period_end = dt_year_start + months(3))
  df_RPT_03_process_kpis[1, 3] <- sprintf('%.1f%% (%d/%d)', 
                                          (df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_in)) %>% NROW())/(df_tmp %>% NROW()) * 100,
                                          (df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_in)) %>% NROW()),
                                          (df_tmp %>% NROW()))
  df_RPT_03_process_kpis[3, 3] <- sprintf('%.1f%% (%d/%d)', 
                                          (df_tmp %>% dplyr::filter(ct_loneliness %in% c('Yes','No')) %>% NROW())/(df_tmp %>% NROW()) * 100,
                                          (df_tmp %>% dplyr::filter(ct_loneliness %in% c('Yes','No')) %>% NROW()),
                                          (df_tmp %>% NROW()))
  # Quarter 2
  df_tmp <- fnNewClientsSupported(df = df_caseload_tracker, 
                                  period_start = dt_year_start + months(3), 
                                  period_end = dt_year_start + months(6))
  df_RPT_03_process_kpis[1, 4] <- sprintf('%.1f%% (%d/%d)', 
                                          (df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_in)) %>% NROW())/(df_tmp %>% NROW()) * 100,
                                          (df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_in)) %>% NROW()),
                                          (df_tmp %>% NROW()))
  df_RPT_03_process_kpis[3, 4] <- sprintf('%.1f%% (%d/%d)', 
                                          (df_tmp %>% dplyr::filter(ct_loneliness %in% c('Yes','No')) %>% NROW())/(df_tmp %>% NROW()) * 100,
                                          (df_tmp %>% dplyr::filter(ct_loneliness %in% c('Yes','No')) %>% NROW()),
                                          (df_tmp %>% NROW()))
  # Quarter 3
  df_tmp <- fnNewClientsSupported(df = df_caseload_tracker, 
                                  period_start = dt_year_start + months(6), 
                                  period_end = dt_year_start + months(9))
  df_RPT_03_process_kpis[1, 5] <- sprintf('%.1f%% (%d/%d)', 
                                          (df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_in)) %>% NROW())/(df_tmp %>% NROW()) * 100,
                                          (df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_in)) %>% NROW()),
                                          (df_tmp %>% NROW()))
  df_RPT_03_process_kpis[3, 5] <- sprintf('%.1f%% (%d/%d)', 
                                          (df_tmp %>% dplyr::filter(ct_loneliness %in% c('Yes','No')) %>% NROW())/(df_tmp %>% NROW()) * 100,
                                          (df_tmp %>% dplyr::filter(ct_loneliness %in% c('Yes','No')) %>% NROW()),
                                          (df_tmp %>% NROW()))
  # Quarter 4
  df_tmp <- fnNewClientsSupported(df = df_caseload_tracker, 
                                  period_start = dt_year_start + months(9), 
                                  period_end = dt_year_start + months(12))
  df_RPT_03_process_kpis[1, 6] <- sprintf('%.1f%% (%d/%d)', 
                                          (df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_in)) %>% NROW())/(df_tmp %>% NROW()) * 100,
                                          (df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_in)) %>% NROW()),
                                          (df_tmp %>% NROW()))
  df_RPT_03_process_kpis[3, 6] <- sprintf('%.1f%% (%d/%d)', 
                                          (df_tmp %>% dplyr::filter(ct_loneliness %in% c('Yes','No')) %>% NROW())/(df_tmp %>% NROW()) * 100,
                                          (df_tmp %>% dplyr::filter(ct_loneliness %in% c('Yes','No')) %>% NROW()),
                                          (df_tmp %>% NROW()))
}

# Year to date
df_tmp <- fnNewClientsSupported(df = df_caseload_tracker, 
                                period_start = dt_year_start, 
                                period_end = dt_year_start + months(12))
df_RPT_03_process_kpis[1, 7] <- sprintf('%.1f%% (%d/%d)', 
                                        (df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_in)) %>% NROW())/(df_tmp %>% NROW()) * 100,
                                        (df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_in)) %>% NROW()),
                                        (df_tmp %>% NROW()))
df_RPT_03_process_kpis[3, 7] <- sprintf('%.1f%% (%d/%d)', 
                                        (df_tmp %>% dplyr::filter(ct_loneliness %in% c('Yes','No')) %>% NROW())/(df_tmp %>% NROW()) * 100,
                                        (df_tmp %>% dplyr::filter(ct_loneliness %in% c('Yes','No')) %>% NROW()),
                                        (df_tmp %>% NROW()))

# Exit questions
# Current Month
df_tmp <- fnClientsSupportedEnd(df = df_caseload_tracker, 
                                period_start = dt_current_month, 
                                period_end = dt_current_month + months(1))
df_RPT_03_process_kpis[2, 2] <- sprintf('%.1f%% (%d/%d)', 
                                        (df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_out)) %>% NROW())/(df_tmp %>% NROW()) * 100,
                                        (df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_out)) %>% NROW()),
                                        (df_tmp %>% NROW()))
df_RPT_03_process_kpis[4, 2] <- NA
if(dt_current_month >= dt_year_start){
  # Quarter 1
  df_tmp <- fnClientsSupportedEnd(df = df_caseload_tracker, 
                                  period_start = dt_year_start, 
                                  period_end = dt_year_start + months(3))
  df_RPT_03_process_kpis[2, 3] <- sprintf('%.1f%% (%d/%d)', 
                                          (df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_out)) %>% NROW())/(df_tmp %>% NROW()) * 100,
                                          (df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_out)) %>% NROW()),
                                          (df_tmp %>% NROW()))
  df_RPT_03_process_kpis[4, 3] <- NA
  # Quarter 2
  df_tmp <- fnClientsSupportedEnd(df = df_caseload_tracker, 
                                  period_start = dt_year_start + months(3), 
                                  period_end = dt_year_start + months(6))
  df_RPT_03_process_kpis[2, 4] <- sprintf('%.1f%% (%d/%d)', 
                                          (df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_out)) %>% NROW())/(df_tmp %>% NROW()) * 100,
                                          (df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_out)) %>% NROW()),
                                          (df_tmp %>% NROW()))
  df_RPT_03_process_kpis[4, 4] <- NA
  # Quarter 3
  df_tmp <- fnClientsSupportedEnd(df = df_caseload_tracker, 
                                  period_start = dt_year_start + months(6), 
                                  period_end = dt_year_start + months(9))
  df_RPT_03_process_kpis[2, 5] <- sprintf('%.1f%% (%d/%d)', 
                                          (df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_out)) %>% NROW())/(df_tmp %>% NROW()) * 100,
                                          (df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_out)) %>% NROW()),
                                          (df_tmp %>% NROW()))
  df_RPT_03_process_kpis[4, 5] <- NA
  # Quarter 4
  df_tmp <- fnClientsSupportedEnd(df = df_caseload_tracker, 
                                  period_start = dt_year_start + months(9), 
                                  period_end = dt_year_start + months(12))
  df_RPT_03_process_kpis[2, 6] <- sprintf('%.1f%% (%d/%d)', 
                                          (df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_out)) %>% NROW())/(df_tmp %>% NROW()) * 100,
                                          (df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_out)) %>% NROW()),
                                          (df_tmp %>% NROW()))
  df_RPT_03_process_kpis[4, 6] <- NA
}
# Year to date
df_tmp <- fnClientsSupportedEnd(df = df_caseload_tracker, 
                                period_start = dt_year_start, 
                                period_end = dt_year_start + months(12))
df_RPT_03_process_kpis[2, 7] <- sprintf('%.1f%% (%d/%d)', 
                                        (df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_out)) %>% NROW())/(df_tmp %>% NROW()) * 100,
                                        (df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_out)) %>% NROW()),
                                        (df_tmp %>% NROW()))
df_RPT_03_process_kpis[4, 7] <- NA

df_RPT_03_process_kpis

# ───────────────────────
# * 1.4. Data points ----
# ───────────────────────

# Process the Total current open cases (aka current caseload as at end of period)
caseload_current_month <- fnCaseload(df = df_caseload_tracker, 
                                     period_start = dt_current_month, 
                                     period_end = dt_current_month + months(1)) %>% NROW()

caseload_Q1 <- ifelse(dt_current_month >= dt_year_start,
                      fnCaseload(df = df_caseload_tracker, 
                                 period_start = dt_year_start, 
                                 period_end = dt_year_start + months(3)) %>% NROW(), NA)
caseload_Q2 <- ifelse(dt_current_month >= dt_year_start + months(3),
                      fnCaseload(df = df_caseload_tracker, 
                                 period_start = dt_year_start + months(3), 
                                 period_end = dt_year_start + months(6)) %>% NROW(), NA)
caseload_Q3 <- ifelse(dt_current_month >= dt_year_start + months(6),
                      fnCaseload(df = df_caseload_tracker, 
                                 period_start = dt_year_start + months(6), 
                                 period_end = dt_year_start + months(9)) %>% NROW(), NA)
caseload_Q4 <- ifelse(dt_current_month >= dt_year_start + months(9),
                      fnCaseload(df = df_caseload_tracker, 
                                 period_start = dt_year_start + months(9), 
                                 period_end = dt_year_start + months(12)) %>% NROW(), NA)
caseload_YTD <- fnCaseload(df = df_caseload_tracker, 
                           period_start = dt_year_start, 
                           period_end = dt_year_start + months(12)) %>% NROW()

df_tmp <- data.frame(
  metric = 'Current caseload', 
  current_month = caseload_current_month,
  q1 = caseload_Q1,
  q2 = caseload_Q2,
  q3 = caseload_Q3,
  q4 = caseload_Q4,
  ytd = caseload_YTD)

metrics <- c('Number of wider beneficiaries', 'Clients who declined', 
             'Case concluded successfully', 'Closed cases due to disengagement', 'Closed cases due to death',
             'Closed cases (other reasons, ie moving out of area)', 'Number of contacts/interventions with clients')
df_metrics <- data.frame(metric = factor(metrics, levels = metrics))

df_RPT_04_data_points <- df_metrics %>% 
  left_join(
    df_data_points %>%
      dplyr::filter(as.Date(month) == dt_current_month) %>%
      group_by(metric) %>%
      summarise(current_month = sum(value, na.rm = TRUE)) %>%
      ungroup(),
    by = 'metric') %>%
  replace_na(replace = list('current_month' = 0)) %>%
  left_join(
    df_data_points %>%
      dplyr::filter(as.Date(month) >= dt_year_start & as.Date(month) < (dt_year_start + months(3))) %>%
      group_by(metric) %>%
      summarise(q1 = sum(value, na.rm = TRUE)) %>%
      ungroup(),
    by = 'metric') %>%
  replace_na(replace = list('q1' = 0)) %>%
  left_join(
    df_data_points %>%
      dplyr::filter(as.Date(month) >= (dt_year_start + months(3)) & as.Date(month) < (dt_year_start + months(6))) %>%
      group_by(metric) %>%
      summarise(q2 = sum(value, na.rm = TRUE)) %>%
      ungroup(),
    by = 'metric') %>%
  replace_na(replace = list('q2' = 0)) %>%
  left_join(
    df_data_points %>%
      dplyr::filter(as.Date(month) >= (dt_year_start + months(6)) & as.Date(month) < (dt_year_start + months(9))) %>%
      group_by(metric) %>%
      summarise(q3 = sum(value, na.rm = TRUE)) %>%
      ungroup(),
    by = 'metric') %>%
  replace_na(replace = list('q3' = 0)) %>%
  left_join(
    df_data_points %>%
      dplyr::filter(as.Date(month) >= (dt_year_start + months(9)) & as.Date(month) < (dt_year_start + months(12))) %>%
      group_by(metric) %>%
      summarise(q4 = sum(value, na.rm = TRUE)) %>%
      ungroup(),
    by = 'metric') %>%
  replace_na(replace = list('q4' = 0)) %>%
  left_join(
    df_data_points %>%
      dplyr::filter(as.Date(month) >= dt_year_start & as.Date(month) < (dt_year_start + months(12))) %>%
      group_by(metric) %>%
      summarise(ytd = sum(value, na.rm = TRUE)) %>%
      ungroup(),
    by = 'metric') %>%
  replace_na(replace = list('ytd' = 0)) %>%
  bind_rows(df_tmp)

df_RPT_04_data_points
  
# ────────────────────────────
# * 1.5. Support provided ----
# ────────────────────────────

metrics <- c('Team Around the Person meeting conducted', 'Flow meeting with FC & Lead Professional',
                'One-to-one work with clients (per client) number of individual one to one interactions with client',
                'Continued ongoing contacts with professionals (total number of seperate contacts)',
                'Caseworker research undertaken to find solutions for clients', 'Caseworker support to access Personal Health Budget',
                'Caseworker support with Form filling', 'Caseworker support with IT incl. virtual meetings, emails etc',
                'Caseworker support to meet aspirations', 'Client involved in coproduction work (total number of seperate contacts)')
df_metrics <- data.frame(metric = factor(metrics, levels = metrics))

df_RPT_05_support_provided <- df_metrics %>% 
  left_join(
    df_support_and_referrals %>%
      dplyr::filter(as.Date(month) == dt_current_month) %>%
      group_by(support) %>%
      summarise(current_month = n()) %>%
      ungroup(),
    by = c('metric' = 'support')) %>%
  replace_na(replace = list('current_month' = 0)) %>%
  left_join(
    df_support_and_referrals %>%
      dplyr::filter(as.Date(month) >= dt_year_start & as.Date(month) < (dt_year_start + months(3))) %>%
      group_by(support) %>%
      summarise(q1 = n()) %>%
      ungroup(),
    by = c('metric' = 'support')) %>%
  replace_na(replace = list('q1' = 0)) %>%
  left_join(
    df_support_and_referrals %>%
      dplyr::filter(as.Date(month) >= (dt_year_start + months(3)) & as.Date(month) < (dt_year_start + months(6))) %>%
      group_by(support) %>%
      summarise(q2 = n()) %>%
      ungroup(),
    by = c('metric' = 'support')) %>%
  replace_na(replace = list('q2' = 0)) %>%
  left_join(
    df_support_and_referrals %>%
      dplyr::filter(as.Date(month) >= (dt_year_start + months(6)) & as.Date(month) < (dt_year_start + months(9))) %>%
      group_by(support) %>%
      summarise(q3 = n()) %>%
      ungroup(),
    by = c('metric' = 'support')) %>%
  replace_na(replace = list('q3' = 0)) %>%
  left_join(
    df_support_and_referrals %>%
      dplyr::filter(as.Date(month) >= (dt_year_start + months(9)) & as.Date(month) < (dt_year_start + months(12))) %>%
      group_by(support) %>%
      summarise(q4 = n()) %>%
      ungroup(),
    by = c('metric' = 'support')) %>%
  replace_na(replace = list('q4' = 0)) %>%
  left_join(
    df_support_and_referrals %>%
      dplyr::filter(as.Date(month) >= dt_year_start & as.Date(month) < (dt_year_start + months(12))) %>%
      group_by(support) %>%
      summarise(ytd = n()) %>%
      ungroup(),
    by = c('metric' = 'support')) %>%
  replace_na(replace = list('ytd' = 0)) 

df_RPT_05_support_provided

# ───────────────────────────
# * 1.6. Outputs section ----
# ───────────────────────────

# For the visuals we will use YTD and past 3 months
df_outputs_sections_prev_3m <- df_outputs %>% 
  dplyr::filter(
    as.Date(month) >= (dt_current_month - months(2)) & 
      as.Date(month) < (dt_current_month + months(1)) &
      as.Date(month) >= dt_year_start) %>%
  group_by(section) %>%
  summarise(volume = n()) %>%
  ungroup()

df_RPT_06_outputs_sections_prev_3m <- fnGlenday(df_outputs_sections_prev_3m, var.x = 'section', var.y = 'volume')
df_RPT_06_outputs_sections_prev_3m

df_outputs_sections_ytd <- df_outputs %>% 
  dplyr::filter(
    as.Date(month) >= dt_year_start & 
    as.Date(month) < (dt_current_month + months(1)) &
    as.Date(month) < (dt_year_start + months(12))
  ) %>%
  group_by(section) %>%
  summarise(volume = n()) %>%
  ungroup()

df_RPT_07_outputs_sections_ytd <- fnGlenday(df_outputs_sections_ytd, var.x = 'section', var.y = 'volume')
df_RPT_07_outputs_sections_ytd

# ─────────────────────────────────────
# * * 1.6.1. Output detail section ----
# ─────────────────────────────────────

# For the visuals we will use YTD and past 3 months
df_outputs_sections_and_outputs_prev_3m <- df_outputs %>% 
  dplyr::filter(
    as.Date(month) >= (dt_current_month - months(2)) & 
      as.Date(month) < (dt_current_month + months(1)) &
      as.Date(month) >= dt_year_start) %>%
  group_by(section, output) %>%
  summarise(volume = n(),
            .groups = 'keep') %>%
  ungroup()

df_RPT_08_outputs_sections_and_outputs_prev_3m <- fnGlenday(df_outputs_sections_and_outputs_prev_3m, var.x = 'output', var.y = 'volume') 
df_RPT_08_outputs_sections_and_outputs_prev_3m

df_outputs_sections_and_outputs_ytd <- df_outputs %>% 
  dplyr::filter(
    as.Date(month) >= dt_year_start & 
    as.Date(month) < (dt_current_month + months(1)) &
    as.Date(month) < (dt_year_start + months(12))
  ) %>%
  group_by(section, output) %>%
  summarise(volume = n(),
            .groups = 'keep') %>%
  ungroup()

df_RPT_09_outputs_sections_and_outputs_ytd <- fnGlenday(df_outputs_sections_and_outputs_ytd, var.x = 'output', var.y = 'volume') 
df_RPT_09_outputs_sections_and_outputs_ytd

# ────────────────────────────
# * 1.7. Outcomes section ----
# ────────────────────────────

# For the visuals we will use YTD and past 3 months
df_outcomes_sections_prev_3m <- df_outcomes %>% 
  dplyr::filter(
    as.Date(month) >= (dt_current_month - months(2)) & 
      as.Date(month) < (dt_current_month + months(1)) &
      as.Date(month) >= dt_year_start) %>%
  group_by(section) %>%
  summarise(volume = n()) %>%
  ungroup()

df_RPT_10_outcomes_sections_prev_3m <- fnGlenday(df_outcomes_sections_prev_3m, var.x = 'section', var.y = 'volume')
df_RPT_10_outcomes_sections_prev_3m

df_outcomes_sections_ytd <- df_outcomes %>% 
  dplyr::filter(as.Date(month) >= dt_year_start) %>%
  group_by(section) %>%
  summarise(volume = n()) %>%
  ungroup()

df_RPT_11_outcomes_sections_ytd <- fnGlenday(df_outcomes_sections_ytd, var.x = 'section', var.y = 'volume')
df_RPT_11_outcomes_sections_ytd

# ───────────────────────────────────────
# * * 1.7.1. Outcome details section ----
# ───────────────────────────────────────

# For the visuals we will use YTD and past 3 months
df_outcomes_sections_and_outcomes_prev_3m <- df_outcomes %>% 
  dplyr::filter(
    as.Date(month) >= (dt_current_month - months(2)) & 
      as.Date(month) < (dt_current_month + months(1)) &
      as.Date(month) >= dt_year_start) %>%
  group_by(section, outcome) %>%
  summarise(volume = n(),
            .groups = 'keep') %>%
  ungroup()

df_RPT_12_outcomes_sections_and_outcomes_prev_3m <- fnGlenday(df_outcomes_sections_and_outcomes_prev_3m, var.x = 'outcome', var.y = 'volume') 
df_RPT_12_outcomes_sections_and_outcomes_prev_3m

df_outcomes_sections_and_outcomes_ytd <- df_outcomes %>% 
  dplyr::filter(as.Date(month) >= dt_year_start) %>%
  group_by(section, outcome) %>%
  summarise(volume = n(),
            .groups = 'keep') %>%
  ungroup()
df_RPT_13_outcomes_sections_and_outcomes_ytd <- fnGlenday(df_outcomes_sections_and_outcomes_ytd, var.x = 'outcome', var.y = 'volume') 
df_RPT_13_outcomes_sections_and_outcomes_ytd

# ════════════════════
# 2. Display data ----
# ════════════════════

# ─────────────────────────────
# * 2.1. Headlines display ----
# ─────────────────────────────

tbl <- flextable(df_RPT_01_headlines) %>%
  set_header_labels(
    metric = 'Metric',
    current_month = format(dt_current_month, '%b-%y'),
    q1 = 'Q1',
    q2 = 'Q2',
    q3 = 'Q3',
    q4 = 'Q4',
    ytd = 'YTD') %>%
  set_table_properties(layout = "autofit", width = .9) %>%
  hline(part = "body", border = std_border) %>%
  vline(part = "all", border = std_border) %>%
  bold(bold = TRUE, part = "header")
tbl
#block_section(prop_section(page_size = page_size(orient = 'portrait')))

# ───────────────────────────────────────
# * 2.2. Changes in activity display ----
# ───────────────────────────────────────

tbl <- flextable(df_RPT_02_changes_in_activity) %>%
  set_header_labels(
    metric = 'Metric',
    current_month = format(dt_current_month, '%b-%y'),
    q1 = 'Q1',
    q2 = 'Q2',
    q3 = 'Q3',
    q4 = 'Q4',
    ytd = 'YTD') %>%
  set_table_properties(layout = "autofit", width = .9) %>%
  hline(part = "body", border = std_border) %>%
  vline(part = "all", border = std_border) %>%
  bold(bold = TRUE, part = "header")
tbl
#block_section(prop_section(page_size = page_size(orient = 'portrait')))

# ────────────────────────────────
# * 2.3. Process KPIs display ----
# ────────────────────────────────

tbl <- flextable(df_RPT_03_process_kpis) %>%
  set_header_labels(
    metric = 'Metric',
    current_month = format(dt_current_month, '%b-%y'),
    q1 = 'Q1',
    q2 = 'Q2',
    q3 = 'Q3',
    q4 = 'Q4',
    ytd = 'YTD') %>%
  set_table_properties(layout = "autofit", width = .9) %>%
  hline(part = "body", border = std_border) %>%
  vline(part = "all", border = std_border) %>%
  bold(bold = TRUE, part = "header")
tbl
#block_section(prop_section(page_size = page_size(orient = 'portrait')))

# ───────────────────────────────
# * 2.4. Data points display ----
# ───────────────────────────────
tbl <- flextable(df_RPT_04_data_points) %>%
  set_header_labels(
    metric = 'Metric',
    current_month = format(dt_current_month, '%b-%y'),
    q1 = 'Q1',
    q2 = 'Q2',
    q3 = 'Q3',
    q4 = 'Q4',
    ytd = 'YTD') %>%
  set_table_properties(layout = "autofit", width = .9) %>%
  hline(part = "body", border = std_border) %>%
  vline(part = "all", border = std_border) %>%
  bold(bold = TRUE, part = "header")
tbl
#block_section(prop_section(page_size = page_size(orient = 'portrait')))

# ────────────────────────────────────
# * 2.5. Support provided display ----
# ────────────────────────────────────
tbl <- flextable(df_RPT_05_support_provided) %>%
  set_header_labels(
    metric = 'Metric',
    current_month = format(dt_current_month, '%b-%y'),
    q1 = 'Q1',
    q2 = 'Q2',
    q3 = 'Q3',
    q4 = 'Q4',
    ytd = 'YTD') %>%
  set_table_properties(layout = "autofit", width = .9) %>%
  hline(part = "body", border = std_border) %>%
  vline(part = "all", border = std_border) %>%
  bold(bold = TRUE, part = "header")
tbl
#block_section(prop_section(type = 'continuous'))

# ─────────────────────────────────────────────────────────
# * 2.6. Outputs: Sections - previous 3 months display ----
# ─────────────────────────────────────────────────────────
txt_period_start_prev_3m <- if(dt_current_month - months(2) < dt_year_start){
  format(dt_year_start, '%b-%y')
} else {
  format(dt_current_month - months(2), '%b-%y')
}
txt_period_end_prev_3m <- format(dt_current_month, '%b-%y')

plt <- ggplot(data = df_RPT_06_outputs_sections_prev_3m) %+%
  theme_bw(base_size = 10) %+%
  theme(plot.title = element_text(hjust = 0.5),
        axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) %+%
  labs(title = paste0('Glenday Sieve Plot for Output Sections for the Previous 3 Months\n',
                      txt_period_start_prev_3m, ' to ', txt_period_end_prev_3m),
       x = 'Output Section', y = 'Number of Outputs') %+%
  guides(fill = guide_legend(title = 'Glenday Sieve')) %+%
  geom_bar(aes(x = forcats::fct_reorder(str_wrap(section, 20), -volume), y = volume, fill = glenday), stat = 'identity') %+%
  scale_fill_manual(values = c('Green' = 'darkgreen',
                               'Yellow' = 'gold',
                               'Blue' = 'royalblue',
                                'Red' = 'red3'),
                     breaks = c('Green','Yellow','Blue','Red'))
plt
#block_section(prop_section(page_size = page_size(orient = 'landscape')))

# ────────────────────────────────────────────────────
# * 2.7. Outputs: Sections - year to date display ----
# ────────────────────────────────────────────────────
txt_period_start_ytd <- format(dt_year_start, '%b-%y')
txt_period_end_ytd <- format(dt_current_month, '%b-%y')

plt <- ggplot(data = df_RPT_07_outputs_sections_ytd) %+%
  theme_bw(base_size = 10) %+%
  theme(plot.title = element_text(hjust = 0.5),
        axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) %+%
  labs(title = paste0('Glenday Sieve Plot for Output Sections for the Year to Date\n',
                      txt_period_start_ytd, ' to ', txt_period_end_ytd),
       x = 'Output Section', y = 'Number of Outputs') %+%
  guides(fill = guide_legend(title = 'Glenday Sieve')) %+%
  geom_bar(aes(x = forcats::fct_reorder(str_wrap(section, 20), -volume), y = volume, fill = glenday), stat = 'identity') %+%
  scale_fill_manual(values = c('Green' = 'darkgreen',
                               'Yellow' = 'gold',
                               'Blue' = 'royalblue',
                                'Red' = 'red3'),
                     breaks = c('Green','Yellow','Blue','Red'))
plt
#block_section(prop_section(page_size = page_size(orient = 'landscape')))


# ───────────────────────────────────────────────────────
# * 2.8. Outputs: Detail - previous 3 months display ----
# ───────────────────────────────────────────────────────
plt <- ggplot(data = df_RPT_08_outputs_sections_and_outputs_prev_3m %>% dplyr::filter(glenday=='Green')) %+%
  theme_bw(base_size = 10) %+%
  theme(plot.title = element_text(hjust = 0.5),
        axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1, size = 8)) %+%
  labs(title = paste0('Glenday Sieve Plot (Green only) for Outputs for the Previous 3 Months\n',
                      txt_period_start_prev_3m, ' to ', txt_period_end_prev_3m),
       x = 'Output', y = 'Number of Outputs') %+%
  guides(fill = guide_legend(title = 'Glenday Sieve')) %+%
  geom_bar(aes(x = forcats::fct_reorder(str_wrap(output, 30), -volume), y = volume, fill = glenday), stat = 'identity') %+%
  scale_fill_manual(values = c('Green' = 'darkgreen',
                               'Yellow' = 'gold',
                               'Blue' = 'royalblue',
                                'Red' = 'red3'),
                     breaks = c('Green','Yellow','Blue','Red'))
plt
#block_section(prop_section(page_size = page_size(orient = 'landscape')))

# ──────────────────────────────────────────────────
# * 2.9. Outputs: Detail - year to date display ----
# ──────────────────────────────────────────────────
plt <- ggplot(data = df_RPT_09_outputs_sections_and_outputs_ytd %>% dplyr::filter(glenday=='Green')) %+%
  theme_bw(base_size = 10) %+%
  theme(plot.title = element_text(hjust = 0.5),
        axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1, size = 8)) %+%
  labs(title = paste0('Glenday Sieve Plot (Green only) for Outputs for the Year to Date\n',
                      txt_period_start_ytd, ' to ', txt_period_end_ytd),
       x = 'Output', y = 'Number of Outputs') %+%
  guides(fill = guide_legend(title = 'Glenday Sieve')) %+%
  geom_bar(aes(x = forcats::fct_reorder(str_wrap(output, 30), -volume), y = volume, fill = glenday), stat = 'identity') %+%
  scale_fill_manual(values = c('Green' = 'darkgreen',
                               'Yellow' = 'gold',
                               'Blue' = 'royalblue',
                                'Red' = 'red3'),
                     breaks = c('Green','Yellow','Blue','Red'))
plt
#block_section(prop_section(page_size = page_size(orient = 'landscape')))

# ───────────────────────────────────────────────────────────
# * 2.10. Outcomes: Sections - previous 3 months display ----
# ───────────────────────────────────────────────────────────
plt <- ggplot(data = df_RPT_10_outcomes_sections_prev_3m) %+%
  theme_bw(base_size = 10) %+%
  theme(plot.title = element_text(hjust = 0.5),
        axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) %+%
  labs(title = paste0('Glenday Sieve Plot for Outcome Sections for the Previous 3 Months\n',
                      txt_period_start_prev_3m, ' to ', txt_period_end_prev_3m),
       x = 'Outcome Section', y = 'Number of Outcomes') %+%
  guides(fill = guide_legend(title = 'Glenday Sieve')) %+%
  geom_bar(aes(x = forcats::fct_reorder(str_wrap(section, 20), -volume), y = volume, fill = glenday), stat = 'identity') %+%
  scale_fill_manual(values = c('Green' = 'darkgreen',
                               'Yellow' = 'gold',
                               'Blue' = 'royalblue',
                                'Red' = 'red3'),
                     breaks = c('Green','Yellow','Blue','Red'))
plt
#block_section(prop_section(page_size = page_size(orient = 'landscape')))

# ──────────────────────────────────────────────────────
# * 2.11. Outcomes: Sections - year to date display ----
# ──────────────────────────────────────────────────────
plt <- ggplot(data = df_RPT_11_outcomes_sections_ytd) %+%
  theme_bw(base_size = 10) %+%
  theme(plot.title = element_text(hjust = 0.5),
        axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) %+%
  labs(title = paste0('Glenday Sieve Plot for Outcome Sections for the Year to Date\n',
                      txt_period_start_prev_3m, ' to ', txt_period_end_prev_3m),
       x = 'Outcome Section', y = 'Number of Outcomes') %+%
  guides(fill = guide_legend(title = 'Glenday Sieve')) %+%
  geom_bar(aes(x = forcats::fct_reorder(str_wrap(section, 20), -volume), y = volume, fill = glenday), stat = 'identity') %+%
  scale_fill_manual(values = c('Green' = 'darkgreen',
                               'Yellow' = 'gold',
                               'Blue' = 'royalblue',
                                'Red' = 'red3'),
                     breaks = c('Green','Yellow','Blue','Red'))
plt
#block_section(prop_section(page_size = page_size(orient = 'landscape')))

# ─────────────────────────────────────────────────
# * 2.12. Outcomes: Detail - previous 3 months ----
# ─────────────────────────────────────────────────
plt <- ggplot(data = df_RPT_12_outcomes_sections_and_outcomes_prev_3m %>% dplyr::filter(glenday=='Green')) %+%
  theme_bw(base_size = 10) %+%
  theme(plot.title = element_text(hjust = 0.5),
        axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1, size = 8)) %+%
  labs(title = paste0('Glenday Sieve Plot (Green only) for Outcomes for the Previous 3 Months\n',
                      txt_period_start_prev_3m, ' to ', txt_period_end_prev_3m),
       x = 'Outcome', y = 'Number of Outcomes') %+%
  guides(fill = guide_legend(title = 'Glenday Sieve')) %+%
  geom_bar(aes(x = forcats::fct_reorder(str_wrap(outcome, 30), -volume), y = volume, fill = glenday), stat = 'identity') %+%
  scale_fill_manual(values = c('Green' = 'darkgreen',
                               'Yellow' = 'gold',
                               'Blue' = 'royalblue',
                                'Red' = 'red3'),
                     breaks = c('Green','Yellow','Blue','Red'))
plt
#block_section(prop_section(page_size = page_size(orient = 'landscape')))

# ────────────────────────────────────────────
# * 2.13. Outcomes: Detail - year to date ----
# ────────────────────────────────────────────
plt <- ggplot(data = df_RPT_13_outcomes_sections_and_outcomes_ytd %>% dplyr::filter(glenday=='Green')) %+%
  theme_bw(base_size = 10) %+%
  theme(plot.title = element_text(hjust = 0.5),
        axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1, size = 8)) %+%
  labs(title = paste0('Glenday Sieve Plot (Green only) for Outcomes for the Year to Date\n',
                      txt_period_start_prev_3m, ' to ', txt_period_end_prev_3m),
       x = 'Outcome', y = 'Number of Outcomes') %+%
  guides(fill = guide_legend(title = 'Glenday Sieve')) %+%
  geom_bar(aes(x = forcats::fct_reorder(str_wrap(outcome, 30), -volume), y = volume, fill = glenday), stat = 'identity') %+%
  scale_fill_manual(values = c('Green' = 'darkgreen',
                               'Yellow' = 'gold',
                               'Blue' = 'royalblue',
                                'Red' = 'red3'),
                     breaks = c('Green','Yellow','Blue','Red'))
plt
#block_section(prop_section(page_size = page_size(orient = 'landscape')))
