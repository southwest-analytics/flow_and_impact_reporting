# 0. Workaround for choose.file dialog issue ----
# ═══════════════════════════════════════════════

# Select files to import ----
# This has to be the first statement in the script as there is a bug with utils::choose.files

# Select ini file
ini_file <- utils::choose.files(caption = 'Please select the INI FILE', multi = FALSE)

# Select the project
project_list <- c('Community Flow' = 2472, 'High Flow' = 2473,
                  'Secondary Care' = 2474, 'Primary Care' = 2475,
                  'Community Flow' = 2476)
project_id <- project_list[utils::select.list(title = 'Select Flow Project', choices = names(project_list), multiple = FALSE, graphics = TRUE)]

# Select caseload tracker
caseload_tracker_file <- utils::choose.files(caption = 'Please select the CASELOAD TRACKER to import', multi = FALSE)

# Select activity file (for High Flow only)
if(project_id == 2473)
  activity_file <- utils::choose.files(caption = 'Please select the ED ACTIVITY file to import', multi = FALSE)

# Select reporting Workbooks
reporting_workbook_filelist <- utils::choose.files(caption = 'Please select the REPORTING WORKBOOKS to import', multi = TRUE)

library(lubridate)

# Get the start of the reporting year
year_start <- svDialogs::dlgInput(message = 'Enter Reporting Year Start in YYYY-MM-DD format (e.g. 2024-01-01)', default = '2024-01-01')$res
dt_year_start <- as.Date(year_start, '%Y-%m-%d')

# Get the current month
current_month <- format(Sys.Date() %m+% months(-1), '%Y-%m-01')
current_month <- svDialogs::dlgInput(message = paste0('Enter Current Month in YYYY-MM-DD format (e.g. ', current_month, ')'), default = current_month)$res
dt_current_month <- as.Date(current_month, '%Y-%m-%d')

# 1. Load libraries and define functions ----
# ═══════════════════════════════════════════

# * 1.1. Load libraries ----
# ──────────────────────────
library(tidyverse)
library(readxl)
library(ini)
library(flextable)
library(officer)
library(officedown)
library(svDialogs)
library(httr)
library(jsonlite)
library(uuid)
library(conflicted)

# * 1.2. Define functions ----
# ────────────────────────────

# * * 1.2.0. Glenday Functions ----
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

# * * 1.2.1. Import Functions ----

# * * * 1.2.1.1. Import Functions: Caseload Tracker ----
fnImportCaseloadTracker <- function(path, sheets){
  # Initialise the data frame to receive the caseload tracker data
  df <- data.frame()
  
  # Get the required field numbers from the ini file
  caseload_tracker_field_numbers <- c(
    unlist(as.integer(unname(ini_file_settings$caseload_tracker_demographics))),
    unlist(as.integer(unname(ini_file_settings$caseload_tracker_numbers_supported))),
    unlist(as.integer(unname(ini_file_settings$caseload_tracker_wemwbs_goals))))
  
  # Get the required field names from the ini file
  caseload_tracker_field_names <- c(
    unlist(names(ini_file_settings$caseload_tracker_demographics)),
    unlist(names(ini_file_settings$caseload_tracker_numbers_supported)),
    unlist(names(ini_file_settings$caseload_tracker_wemwbs_goals)))
  
  # Loop through each of the caseload tracker worksheets
  for(s in sheets){
    # Read in all bar the first two lines of the sheets (the headers)
    df_tmp <- read_excel(path = path, 
                         sheet = s,
                         col_type = 'text', 
                         col_names = (ini_file_settings$caseload_tracker_import$header==TRUE),
                         skip = as.integer(ini_file_settings$caseload_tracker_import$skip)
    ) %>% 
      # Select the required fields and rename them
      select(all_of(caseload_tracker_field_numbers)) %>%
      rename_with(.fn = ~caseload_tracker_field_names) %>%
      # Ignore any rows that don't have an ID
      dplyr::filter(!is.na(ct_id))
    
    # Bind to the combined data frame
    df <- df %>% bind_rows(df_tmp)
  }
  
  # Format the columns to required format
  df <- df %>% 
    mutate(ct_age = as.integer(ct_age),
           ct_status = as.factor(ct_status),
           ct_closure = as.factor(ct_closure),
           ct_start = as.Date(as.integer(ct_start), origin = '1899-12-30'),
           ct_end = as.Date(as.integer(ct_end), origin = '1899-12-30'),
           ct_wemwbs_score_in = as.integer(ct_wemwbs_score_in),
           ct_wemwbs_score_out = as.integer(ct_wemwbs_score_out),
           ct_goals_in = as.integer(ct_goals_in),
           ct_goals_out = as.integer(ct_goals_out))
  
  return(df)
}

# * * * 1.2.1.2. Import Functions: Caseload Tracker Data Quality ----
fnCaseloadTrackerDataQuality <- function(df){
  log <- file(description = 'caseload_tracker_data_quality.log',
              open = 'wt')
  cat('Field, Applicable Entries, Valid Entries\n', file = log)
  # ct_id: all should be valid as previously filter on !is.na
  log_entry <- sprintf('ct_id,%d,%d\n', 
                       NROW(df), 
                       df %>% dplyr::filter(!is.na(ct_id)) %>% NROW())
  cat(log_entry, file = log)
  # ct_gender: should be either 'Male' or 'Female'
  log_entry <- sprintf('ct_gender,%d,%d\n', 
                       NROW(df), 
                       df %>% dplyr::filter(ct_gender %in% c('Male', 'Female')) %>% NROW())
  cat(log_entry, file = log)
  # ct_age: should be an integer and not NA 
  log_entry <- sprintf('ct_age,%d,%d\n', 
                       NROW(df), 
                       df %>% dplyr::filter(is.integer(ct_age)) %>% NROW())
  cat(log_entry, file = log)
  # ct_loneliness_in: should be 'Yes' or 'No' for any Open or Closed cases (Paused or Pre-Engagement ignored)
  log_entry <- sprintf('ct_loneliness_in,%d,%d\n', 
                       df %>% dplyr::filter(ct_status %in% c('Open', 'Closed')) %>% NROW(),
                       df %>% dplyr::filter(ct_status %in% c('Open', 'Closed') & ct_loneliness_in %in% c('Yes','No')) %>% NROW())
  cat(log_entry, file = log)
  # ct_loneliness_out: should be 'Yes' or 'No' for any Open or Closed cases (Paused or Pre-Engagement ignored)
  log_entry <- sprintf('ct_loneliness_out,%d,%d\n', 
                       df %>% dplyr::filter(ct_status %in% c('Open', 'Closed')) %>% NROW(),
                       df %>% dplyr::filter(ct_status %in% c('Open', 'Closed') & ct_loneliness_out %in% c('Yes','No')) %>% NROW())
  cat(log_entry, file = log)
  # ct_postcode: should be anything but NA
  log_entry <- sprintf('ct_postcode,%d,%d\n', 
                       NROW(df),
                       df %>% dplyr::filter(!is.na(ct_postcode)) %>% NROW())
  cat(log_entry, file = log)
  # ct_status: should be 'Open', 'Closed', 'Paused', 'Pre-engagement'
  log_entry <- sprintf('ct_status,%d,%d\n', 
                       NROW(df),
                       df %>% dplyr::filter(ct_status %in% c('Open', 'Closed', 'Paused', 'Pre-engagement')) %>% NROW())
  cat(log_entry, file = log)
  # ct_closure: should be 'Closed successfully', 'Disengaged', 'Non-engagement', 'Declined', 'Exempt', 'Incorrect contact details' when ct_stauts is 'Closed'
  log_entry <- sprintf('ct_closure,%d,%d\n', 
                       df %>% dplyr::filter(ct_status=='Closed') %>% NROW(),
                       df %>% dplyr::filter(ct_status=='Closed' & ct_closure %in% c('Closed successfully', 'Disengaged', 'Non-engagement', 'Declined', 'Exempt', 'Incorrect contact details')) %>% NROW())
  cat(log_entry, file = log)
  # ct_start: should be a valid date unless ct_status is 'Pre-engagement' or ct_closure is 'Declined', 'Exempt', 'Incorrect contact details', 'Non-engagement'
  log_entry <- sprintf('ct_start,%d,%d\n', 
                       df %>% dplyr::filter(ct_status!='Pre-engagement' & !(ct_closure %in% c('Declined', 'Exempt', 'Incorrect contact details', 'Non-engagement'))) %>% NROW(),
                       df %>% dplyr::filter(ct_status!='Pre-engagement' & !(ct_closure %in% c('Declined', 'Exempt', 'Incorrect contact details', 'Non-engagement')) &!is.na(ct_start)) %>% NROW())
  cat(log_entry, file = log)
  # ct_end: should be a valid date if ct_status is 'Closed'
  log_entry <- sprintf('ct_end,%d,%d\n', 
                       df %>% dplyr::filter(ct_status == 'Closed') %>% NROW(),
                       df %>% dplyr::filter(ct_status == 'Closed' & !is.na(ct_end)) %>% NROW())
  cat(log_entry, file = log)
  # ct_wemwbs_score_in: should be an integer unless ct_closure is 'Declined', 'Exempt', 'Incorrect contact details' or ct_status is 'Pre-engagement'
  log_entry <- sprintf('ct_wemwbs_score_in,%d,%d\n', 
                       df %>% dplyr::filter(ct_status!='Pre-engagement' & !(ct_closure %in% c('Declined', 'Exempt', 'Incorrect contact details'))) %>% NROW(),
                       df %>% dplyr::filter(ct_status!='Pre-engagement' & !(ct_closure %in% c('Declined', 'Exempt', 'Incorrect contact details')) & !is.na(ct_wemwbs_score_in)) %>% NROW())
  cat(log_entry, file = log)
  # ct_wemwbs_score_out: should be an integer if ct_status is 'Closed' unless ct_closure is 'Declined', 'Exempt', 'Incorrect contact details'
  log_entry <- sprintf('ct_wemwbs_score_out,%d,%d\n', 
                       df %>% dplyr::filter(ct_status=='Closed' & !(ct_closure %in% c('Declined', 'Exempt', 'Incorrect contact details'))) %>% NROW(),
                       df %>% dplyr::filter(ct_status=='Closed' & !(ct_closure %in% c('Declined', 'Exempt', 'Incorrect contact details')) & !is.na(ct_wemwbs_score_out)) %>% NROW())
  cat(log_entry, file = log)  
  # ct_goals_in: should be an integer unless ct_closure is 'Declined', 'Exempt', 'Incorrect contact details'
  log_entry <- sprintf('ct_goals_in,%d,%d\n', 
                       df %>% dplyr::filter(ct_status!='Pre-engagement' & !(ct_closure %in% c('Declined', 'Exempt', 'Incorrect contact details'))) %>% NROW(),
                       df %>% dplyr::filter(ct_status!='Pre-engagement' & !(ct_closure %in% c('Declined', 'Exempt', 'Incorrect contact details')) & !is.na(ct_goals_in)) %>% NROW())
  cat(log_entry, file = log)
  # ct_goals_out: should be an integer if ct_status is 'Closed' unless ct_closure is 'Declined', 'Exempt', 'Incorrect contact details'
  log_entry <- sprintf('ct_goals_out,%d,%d\n', 
                       df %>% dplyr::filter(ct_status=='Closed' & !(ct_closure %in% c('Declined', 'Exempt', 'Incorrect contact details'))) %>% NROW(),
                       df %>% dplyr::filter(ct_status=='Closed' & !(ct_closure %in% c('Declined', 'Exempt', 'Incorrect contact details')) & !is.na(ct_goals_out)) %>% NROW())
  cat(log_entry, file = log)  
  close(log)
}

# * * * 1.2.1.3. Import Functions: Reporting Workbook - Data Points ----
fnImportReportingWorkbook_DataPoints <- function(path, sheet = 'Data Points'){
  # Only import the first 3 columns of the sheet and set the variable type
  df <- read_excel(path, sheet, range = cell_cols('A:C'), col_types = c('date','text','numeric')) %>%
    # Rename the column names as previously these have overwritten by caseworkers
    rename_with(.fn = ~c('month', 'metric', 'value')) %>% 
    # Remove any rows that have no month or no value
    dplyr::filter(!is.na(month) & !is.na(value)) %>%
    # Add the source filename to the data
    mutate(source = basename(file.path(f)))
  # Return the data frame
  return(df)
}

# * * * 1.2.1.4. Import Functions: Reporting Workbook - Support and Referrals ----
fnImportReportingWorkbook_SupportReferrals <- function(path, sheet = 'Support and Referrals'){
  # Only import the first 4 columns of the sheet and set the variable type
  df <- read_excel(path, sheet, range = cell_cols('A:D'), col_types = c('text','date','text','text')) %>%
    # Rename the column names as previously these have overwritten by caseworkers
    rename_with(.fn = ~c('client_id', 'month', 'section', 'support')) %>% 
    # Remove any rows that have an NA in column
    dplyr::filter( !(is.na(client_id) | is.na(month) | is.na(section) | is.na(support))) %>%
    # Add the source filename to the data
    mutate(source = basename(file.path(f)))
  # Return the data frame
  return(df)
}

# * * * 1.2.1.5. Import Functions: Reporting Workbook - Outputs ----
fnImportReportingWorkbook_Outputs <- function(path, sheet = 'Outputs'){
  # Only import the first 4 columns of the sheet and set the variable type
  df <- read_excel(path, sheet, range = cell_cols('A:D'), col_types = c('text','date','text','text')) %>%
    # Rename the column names as previously these have overwritten by caseworkers
    rename_with(.fn = ~c('client_id', 'month', 'section', 'output')) %>% 
    # Remove any rows that have an NA in column
    dplyr::filter( !(is.na(client_id) | is.na(month) | is.na(section) | is.na(output))) %>%
    # Add the source filename to the data
    mutate(source = basename(file.path(f)))
  # Return the data frame
  return(df)
}

# * * * 1.2.1.6. Import Functions: Reporting Workbook - Outcomes ----
fnImportReportingWorkbook_Outcomes <- function(path, sheet = 'Outcomes'){
  # Only import the first 4 columns of the sheet and set the variable type
  df <- read_excel(path = f, 
                       sheet = 'Outcomes', 
                       range = cell_cols('A:D'), 
                       col_types = c('text','date','text','text')) %>%
    # Rename the column names as previously these have overwritten by caseworkers
    rename_with(.fn = ~c('client_id', 'month', 'section', 'outcome')) %>% 
    # Remove any rows that have an NA in column
    dplyr::filter( !(is.na(client_id) | is.na(month) | is.na(section) | is.na(outcome))) %>%
    # Add the source filename to the data
    mutate(source = basename(file.path(f)))
  
  # Return the data frame
  return(df)
}

# * * * 1.2.1.7. Import Functions: Activity Workbook ----
fnImportActivityWorkbook <- function(path, sheet = 'All HIU Clients'){
  # DEVELOPMENT NOTE: If we want to only report on engaged clients this is where we will need 
  # to filter the data
  
  # Import the activity file
  df <- read_excel(path, 
                   sheet, 
                   skip = 2,
                   col_types = c('text', rep('numeric', 3), rep('date', 2), rep('numeric', 24)),
                   col_names = FALSE) %>%
    # Rename the column names as previously these have overwritten by caseworkers
    rename_with(.fn = ~c('client_id', 
                         'baseline_ed_3m', 'baseline_em_3m', 'baseline_amb_3m',
                         'start_date', 'end_date',
                         'em_3m_prior', 'em_3m_post_start', 'em_3m_during', 'em_3m_post_end',
                         'ed_3m_prior', 'ed_3m_post_start', 'ed_3m_during', 'ed_3m_post_end',
                         'amb_3m_prior', 'amb_3m_post_start', 'amb_3m_during', 'amb_3m_post_end',
                         'em_12m_prior', 'em_12m_post_start', 'em_12m_during', 'em_12m_post_end',
                         'ed_12m_prior', 'ed_12m_post_start', 'ed_12m_during', 'ed_12m_post_end',
                         'amb_12m_prior', 'amb_12m_post_start', 'amb_12m_during', 'amb_12m_post_end')) %>%
    dplyr::filter(!is.na(start_date))
    
  # Calculate the true ED values (i.e. add in the AMB comveyances to the ED attendances to get the overall ED attendances)
  df <- df %>%
    mutate(ed_3m_prior = ed_3m_prior + amb_3m_prior,
           ed_3m_post_start = ed_3m_post_start + amb_3m_post_start,
           ed_3m_during = ed_3m_during + amb_3m_during,
           ed_3m_post_end = ed_3m_post_end + amb_3m_post_end,
           ed_12m_prior = ed_12m_prior + amb_12m_prior,
           ed_12m_post_start = ed_12m_post_start + amb_12m_post_start,
           ed_12m_during = ed_12m_during + amb_12m_during,
           ed_12m_post_end = ed_12m_post_end + amb_12m_post_end)
  # Set any blank end_date to the start of the next month after the current month
  df <- df %>% mutate(end_date = if_else(is.na(end_date), dt_current_month %m+% months(1), end_date))
  # Calculate duration for each row
  df <- df %>% mutate(duration = as.numeric(difftime(end_date, start_date)), .after = 'end_date')
  # Standardise the 3m during periods to 91.25 days to be equivalent to 3 months
  df <- df %>% mutate(em_3m_during = em_3m_during/duration*91.25, 
                      ed_3m_during = ed_3m_during/duration*91.25, 
                      amb_3m_during = amb_3m_during/duration*91.25)
  # Standardise the 12m during periods to 365 days to be equivalent to 12 months
  df <- df %>% mutate(em_12m_during = em_12m_during/duration*365, 
                      ed_12m_during = ed_12m_during/duration*365, 
                      amb_12m_during = amb_12m_during/duration*365)
  
  # Group the data
  df_tmp <- df %>% dplyr::filter(start_date >= dt_current_month &
                                   start_date < dt_current_month %m+% months(1)) %>% 
    summarise(across(.cols = 8:31, .fns = function(x){sum(x, na.rm = TRUE)}),
              `Activity Reduction Cohort` = n()) %>%
    mutate(period = 'current_month', .before = 1)
  # Quarter 1
  if(dt_current_month >= dt_year_start){
    df_tmp <- df_tmp %>% bind_rows(
      df %>% dplyr::filter(start_date >= dt_year_start &
                             start_date < dt_year_start %m+% months(3)) %>% 
      summarise(across(.cols = 8:31, .fns = function(x){sum(x, na.rm = TRUE)}),
                `Activity Reduction Cohort` = n()) %>%
      mutate(period = 'q1', .before = 1))
  } else {
    df_tmp <- df_tmp %>% bind_rows(data.frame(period = 'q1'))
  }
  
  # Quarter 2
  if(dt_current_month >= dt_year_start %m+% months(3)){
    df_tmp <- df_tmp %>% bind_rows(
      df %>% dplyr::filter(start_date >= dt_year_start %m+% months(3) &
                             start_date < dt_year_start %m+% months(6)) %>% 
        summarise(across(.cols = 8:31, .fns = function(x){sum(x, na.rm = TRUE)}),
                  `Activity Reduction Cohort` = n()) %>%
        mutate(period = 'q2', .before = 1))
  } else {
    df_tmp <- df_tmp %>% bind_rows(data.frame(period = 'q2'))
  }

  if(dt_current_month >= dt_year_start %m+% months(6)){
    df_tmp <- df_tmp %>% bind_rows(
      df %>% dplyr::filter(start_date >= dt_year_start %m+% months(6) &
                             start_date < dt_year_start %m+% months(9)) %>% 
        summarise(across(.cols = 8:31, .fns = function(x){sum(x, na.rm = TRUE)}),
                  `Activity Reduction Cohort` = n()) %>%
        mutate(period = 'q3', .before = 1))
  } else {
    df_tmp <- df_tmp %>% bind_rows(data.frame(period = 'q3'))
  }
  
  if(dt_current_month >= dt_year_start %m+% months(9)){
    df_tmp <- df_tmp %>% bind_rows(
      df %>% dplyr::filter(start_date >= dt_year_start %m+% months(9) &
                             start_date < dt_year_start %m+% months(12)) %>% 
        summarise(across(.cols = 8:31, .fns = function(x){sum(x, na.rm = TRUE)}),
                  `Activity Reduction Cohort` = n()) %>%
        mutate(period = 'q4', .before = 1))
  } else {
    df_tmp <- df_tmp %>% bind_rows(data.frame(period = 'q4'))
  }
  
  df_tmp <- df_tmp %>% bind_rows(
    df %>% dplyr::filter(start_date >= dt_year_start &
                             start_date < dt_year_start %m+% months(12)) %>% 
        summarise(across(.cols = 8:31, .fns = function(x){sum(x, na.rm = TRUE)}),
                  `Activity Reduction Cohort` = n()) %>%
        mutate(period = 'ytd', .before = 1))

  # Calculate the reductions
  df_tmp <- df_tmp %>% 
    mutate(`Previous 3 Month Activity for Clients Supported in Period: ED Attendances` = ed_3m_prior,
           `Previous 3 Month Activity for Clients Supported in Period: Emergency Admissions` = em_3m_prior,
           `Previous 3 Month Activity for Clients Supported in Period: Ambulance Conveyances` = amb_3m_prior,
           `Previous 12 Month Activity for Clients Supported in Period: ED Attendances` = ed_12m_prior,
           `Previous 12 Month Activity for Clients Supported in Period: Emergency Admissions` = em_3m_prior,
           `Previous 12 Month Activity for Clients Supported in Period: Ambulance Conveyances` = amb_3m_prior,

           `EM Adms 3 Month Pre vs During` = sprintf('%.1f%% (from: %.0f to: %.0f)', (em_3m_during - em_3m_prior)/em_3m_prior*100, em_3m_prior, em_3m_during),
           `EM Adms 3 Month Pre vs Post End` = sprintf('%.1f%% (from: %.0f to: %.0f)', (em_3m_post_end - em_3m_prior)/em_3m_prior*100, em_3m_prior, em_3m_post_end),
           `ED Atts 3 Month Pre vs During` = sprintf('%.1f%% (from: %.0f to: %.0f)', (ed_3m_during - ed_3m_prior)/ed_3m_prior*100, ed_3m_prior, ed_3m_during),
           `ED Atts 3 Month Pre vs Post Start` = sprintf('%.1f%% (from: %.0f to: %.0f)', (ed_3m_post_start - ed_3m_prior)/ed_3m_prior*100, ed_3m_prior, ed_3m_post_start),
           `ED Atts 3 Month Pre vs Post End` = sprintf('%.1f%% (from: %.0f to: %.0f)', (em_3m_post_end - ed_3m_prior)/ed_3m_prior*100, ed_3m_prior, em_3m_post_end),
           `AMB Conveys 3 Month Pre vs Post Start` = sprintf('%.1f%% (from: %.0f to: %.0f)', (amb_3m_post_start - amb_3m_prior)/amb_3m_prior*100, amb_3m_prior, amb_3m_post_start),
           `AMB Conveys 3 Month Pre vs During` = sprintf('%.1f%% (from: %.0f to: %.0f)', (amb_3m_during - amb_3m_prior)/amb_3m_prior*100, amb_3m_prior, amb_3m_during),
           `AMB Conveys 3 Month Pre vs Post End` = sprintf('%.1f%% (from: %.0f to: %.0f)', (em_3m_post_end - amb_3m_prior)/amb_3m_prior*100, amb_3m_prior, em_3m_post_end),

           `EM Adms 12 Month Pre vs Post Start` = sprintf('%.1f%% (from: %.0f to: %.0f)', (em_12m_post_start - em_12m_prior)/em_12m_prior*100, em_12m_prior, em_12m_post_start),
           `EM Adms 12 Month Pre vs During` = sprintf('%.1f%% (from: %.0f to: %.0f)', (em_12m_during - em_12m_prior)/em_12m_prior*100, em_12m_prior, em_12m_during),
           `EM Adms 12 Month Pre vs Post End` = sprintf('%.1f%% (from: %.0f to: %.0f)', (em_12m_post_end - em_12m_prior)/em_12m_prior*100, em_12m_prior, em_12m_post_end),
           `ED Adms 12 Month Pre vs Post Start` = sprintf('%.1f%% (from: %.0f to: %.0f)', (ed_12m_post_start - ed_12m_prior)/ed_12m_prior*100, ed_12m_prior, ed_12m_post_start),
           `ED Adms 12 Month Pre vs During` = sprintf('%.1f%% (from: %.0f to: %.0f)', (ed_12m_during - ed_12m_prior)/ed_12m_prior*100, ed_12m_prior, ed_12m_during),
           `ED Adms 12 Month Pre vs Post End` = sprintf('%.1f%% (from: %.0f to: %.0f)', (ed_12m_post_end - ed_12m_prior)/ed_12m_prior*100, ed_12m_prior, ed_12m_post_end),
           `AMB Adms 12 Month Pre vs Post Start` = sprintf('%.1f%% (from: %.0f to: %.0f)', (amb_12m_post_start - amb_12m_prior)/amb_12m_prior*100, amb_12m_prior, amb_12m_post_start),
           `AMB Adms 12 Month Pre vs During` = sprintf('%.1f%% (from: %.0f to: %.0f)', (amb_12m_during - amb_12m_prior)/amb_12m_prior*100, amb_12m_prior, amb_12m_during),
           `AMB Adms 12 Month Pre vs Post End` = sprintf('%.1f%% (from: %.0f to: %.0f)', (amb_12m_post_end - amb_12m_prior)/amb_12m_prior*100, amb_12m_prior, amb_12m_post_end))
  
  df <- t(df_tmp %>% select(-period)) %>% as.data.frame()
  colnames(df) <- df_tmp$period
  df <- df %>% mutate(metric = rownames(.), .before = 1)
  rownames(df) <- NULL
  return(df)
}

# * * 1.2.2. Impact Reporting Section ----

# Set up main URL
url_text <- 'https://app.impactreporting.co.uk/api/v1/logs'

# * * * 1.2.2.1. Impact Reporting: Post Data ----
fnPostData <- function(x, uuid){
  body_text <- jsonlite::toJSON(
    list(
      'apikey' = 'd866628d-f9f3-4668-9849-cbba687774d9',
      'logs' = list(
        list(
          'ownerId' = 36517,
          'ownerType' = 'user',
          'activityId' = as.integer(x['id']),
          'projectId' = unname(project_id),
          'loggedByUserId' = 36517,
          'start' = format(as.Date(x['month']), '%d/%m/%Y'),
          'value' = as.numeric(x['value']),
          'text' = paste0('Bulk Upload [', x['uuid'], ']'),
          'classifications' = list(
            list(
              'name' = x['classification_name'],
              'value' = if_else(is.na(x['classification_value']), '', x['classification_value'])
            )
          )
        )
      )
    ), 
    auto_unbox = TRUE
  )
  res <- POST(url = url_text, body = body_text)
  return(res$status_code)
}


# * * 1.2.3. Headlines Section ----

# * * * 1.2.3.1. Headlines Section: New Clients Supported ----
fnGetNewClientsSupported <- function(df_ct){
  # Initialise the variables to receive the data
  curr_month <- as.integer(NA)
  q1 <- as.integer(NA)
  q2 <- as.integer(NA)
  q3 <- as.integer(NA)
  q4 <- as.integer(NA)
  ytd <- as.integer(NA)
  
  # Filter the input data frame to the selected cohort
  
  # Anyone with a ct_start within the period but without a ct_closure in the exception list in the ini file 
  # i.e. 'Declined', 'Non-engagement', 'Exempt', 'Incorrect contact details' see ini file [caseload_tracker_exclusions]
  # section settings for the list of ct_closure exclusions
  df_tmp <- df_ct %>% dplyr::filter(!(ct_closure %in% unlist(unname(ini_file_settings$caseload_tracker_exclusions))))
  
  # Trim the data to end at the end of the current month
  df_tmp <- df_tmp %>% dplyr::filter(ct_start <= dt_current_month %m+% months(1))
  
  # Calculate for each period
  curr_month <- df_tmp %>% dplyr::filter((ct_start >= dt_current_month) & 
                                           (ct_start < dt_current_month %m+% months(1))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    q1 <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start) & 
                                     (ct_start < dt_year_start %m+% months(3))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(3))
    q2 <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start %m+% months(3)) & 
                                     (ct_start < dt_year_start %m+% months(6))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(6))
    q3 <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start %m+% months(6)) & 
                                     (ct_start < dt_year_start %m+% months(9))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(9))
    q4 <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start %m+% months(9)) & 
                                     (ct_start < dt_year_start %m+% months(12))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    ytd <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start) & 
                                      (ct_start < dt_year_start %m+% months(12))) %>% NROW()
  df_tmp <- data.frame(metric = 'New Clients Supported in Period',
                       current_month = curr_month,
                       q1 = q1,
                       q2 = q2,
                       q3 = q3,
                       q4 = q4,
                       ytd = ytd)
  return(df_tmp)
}

# * * * 1.2.3.2. Headlines Section: Clients Supported in Period ----
fnGetClientsSupported <- function(df_ct){
  # Initialise the variables to receive the data
  curr_month <- as.integer(NA)
  q1 <- as.integer(NA)
  q2 <- as.integer(NA)
  q3 <- as.integer(NA)
  q4 <- as.integer(NA)
  ytd <- as.integer(NA)
  
  # Filter the input data frame to the selected cohort
  
  # Anyone with a ct_start before the end of the period and no end date or an end date that occurs after the 
  # start of the period, and a further check to ensure the ct_start >= ct_end if both are present and
  # without a ct_closure in the exception list in the ini file 
  # i.e. 'Declined', 'Non-engagement', 'Exempt', 'Incorrect contact details' see ini file [caseload_tracker_exclusions]
  # section settings for the list of ct_closure exclusions
  df_tmp <- df_ct %>% dplyr::filter(!(ct_closure %in% unlist(unname(ini_file_settings$caseload_tracker_exclusions))))
  
  # Trim the data to end at the end of the current month
  df_tmp <- df_tmp %>% dplyr::filter(ct_start <= dt_current_month %m+% months(1) & 
                                       (is.na(ct_end) | (ct_end < dt_current_month %m+% months(1))))
  
  # Calculate for each period
  curr_month <- df_tmp %>% dplyr::filter((ct_start < dt_current_month %m+% months(1)) & 
                                           (is.na(ct_end) | (ct_end >= dt_current_month & ct_start <= ct_end))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    q1 <- df_tmp %>% dplyr::filter((ct_start < dt_year_start %m+% months(3)) & 
                                     (is.na(ct_end) | (ct_end >= dt_year_start & ct_start <= ct_end))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    q2 <- df_tmp %>% dplyr::filter((ct_start < dt_year_start %m+% months(6)) & 
                                     (is.na(ct_end) | (ct_end >= dt_year_start %m+% months(3) & ct_start <= ct_end))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    q3 <- df_tmp %>% dplyr::filter((ct_start < dt_year_start %m+% months(9)) & 
                                     (is.na(ct_end) | (ct_end >= dt_year_start %m+% months(6) & ct_start <= ct_end))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    q4 <- df_tmp %>% dplyr::filter((ct_start < dt_year_start %m+% months(12)) & 
                                     (is.na(ct_end) | (ct_end >= dt_year_start %m+% months(9) & ct_start <= ct_end))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    ytd <- df_tmp %>% dplyr::filter((ct_start < dt_year_start %m+% months(12)) & 
                                      (is.na(ct_end) | (ct_end >= dt_year_start & ct_start <= ct_end))) %>% NROW()
  df_tmp <- data.frame(metric = 'Clients Supported in Period',
                       current_month = curr_month,
                       q1 = q1,
                       q2 = q2,
                       q3 = q3,
                       q4 = q4,
                       ytd = ytd)
  return(df_tmp)
}

# * * * 1.2.3.3. Headlines Section: Previous 3 Month Activity for Clients Supported in Period ----
#   ED Attendances*
#   Emergency Admissions*
#   Ambulance Conveyances*
# * * * 1.2.3.4. Headlines Section: Previous 12 Month Activity for Clients Supported in Period ----
#   ED Attendances*
#   Emergency Admissions*
#   Ambulance Conveyances*

# * * 1.2.4. Changes in Activity Section ----
# * * * 1.2.4.1. Changes in Activity Section: New Clients Supported in Period: As per 1.2.3.1. ----

# * * * 1.2.4.2. Changes in Activity Section: Reduction in Activity 3 months ----
#   ED Attendances*
#   Emergency Admissions*
#   Ambulance Conveyances*
# * * * 1.2.4.3. Changes in Activity Section: Reduction in Activity 12 months ----
#   ED Attendances*
#   Emergency Admissions*
#   Ambulance Conveyances*

# * * * 1.2.4.4. Changes in Activity Section: Reduction in Loneliness at End of Support ----
fnGetReductionInLoneliness <- function(df_ct){
  # Initialise the variables to receive the data
  curr_month <- as.integer(NA)
  q1 <- as.integer(NA)
  q2 <- as.integer(NA)
  q3 <- as.integer(NA)
  q4 <- as.integer(NA)
  ytd <- as.integer(NA)
  
  # Filter the input data frame to the selected cohort
  
  # Anyone with a ct_end within the period and without a ct_closure in the exception list in the ini file 
  # i.e. 'Declined', 'Non-engagement', 'Exempt', 'Incorrect contact details' see ini file [caseload_tracker_exclusions]
  # section settings for the list of ct_closure exclusions
  df_tmp <- df_ct %>% dplyr::filter(!(ct_closure %in% unlist(unname(ini_file_settings$caseload_tracker_exclusions))))
  
  # and who had a 'Yes' for the loneliness question on entry and a 'No' for the loneliness question on exit
  df_tmp <- df_tmp %>% dplyr::filter(ct_loneliness_in=='Yes' & ct_loneliness_out=='No')

  # Trim the data to end at the end of the current month
  df_tmp <- df_tmp %>% dplyr::filter(ct_end <= dt_current_month %m+% months(1))
  
  # Calculate for each period
  curr_month <- df_tmp %>% dplyr::filter((ct_end >= dt_current_month) & 
                                           (ct_end < dt_current_month %m+% months(1))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    q1 <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start) & 
                                     (ct_end < dt_year_start %m+% months(3))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(3))
    q2 <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start %m+% months(3)) & 
                                     (ct_end < dt_year_start %m+% months(6))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(6))
    q3 <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start %m+% months(6)) & 
                                     (ct_end < dt_year_start %m+% months(9))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(9))
    q4 <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start %m+% months(9)) & 
                                     (ct_end < dt_year_start %m+% months(12))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    ytd <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start) & 
                                      (ct_end < dt_year_start %m+% months(12))) %>% NROW()
  df_tmp <- data.frame(metric = 'Reduction in Loneliness at End of Support',
                       current_month = curr_month,
                       q1 = q1,
                       q2 = q2,
                       q3 = q3,
                       q4 = q4,
                       ytd = ytd)
  return(df_tmp)
}

  
# * * * 1.2.4.5. Changes in Activity Section: Improved Wellbeing at End of Support ----
fnGetIncreasedWEMWBSOnExit <- function(df_ct){
  # Initialise the variables to receive the data
  curr_month <- as.integer(NA)
  q1 <- as.integer(NA)
  q2 <- as.integer(NA)
  q3 <- as.integer(NA)
  q4 <- as.integer(NA)
  ytd <- as.integer(NA)
  
  # Filter the input data frame to the selected cohort
  
  # Anyone with a ct_end within the period and without a ct_closure in the exception list in the ini file 
  # i.e. 'Declined', 'Non-engagement', 'Exempt', 'Incorrect contact details' see ini file [caseload_tracker_exclusions]
  # section settings for the list of ct_closure exclusions
  df_tmp <- df_ct %>% dplyr::filter(!(ct_closure %in% unlist(unname(ini_file_settings$caseload_tracker_exclusions))))
  
  # and who had an  increase in WEMWBS score on exit
  df_tmp <- df_tmp %>% dplyr::filter(ct_wemwbs_score_out > ct_wemwbs_score_in)
  
  # Trim the data to end at the end of the current month
  df_tmp <- df_tmp %>% dplyr::filter(ct_end <= dt_current_month %m+% months(1))
  
  # Calculate for each period
  curr_month <- df_tmp %>% dplyr::filter((ct_end >= dt_current_month) & 
                                           (ct_end < dt_current_month %m+% months(1))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    q1 <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start) & 
                                     (ct_end < dt_year_start %m+% months(3))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(3))
    q2 <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start %m+% months(3)) & 
                                     (ct_end < dt_year_start %m+% months(6))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(6))
    q3 <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start %m+% months(6)) & 
                                     (ct_end < dt_year_start %m+% months(9))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(9))
    q4 <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start %m+% months(9)) & 
                                     (ct_end < dt_year_start %m+% months(12))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    ytd <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start) & 
                                      (ct_end < dt_year_start %m+% months(12))) %>% NROW()
  df_tmp <- data.frame(metric = 'Improved Wellbeing at End of Support',
                       current_month = curr_month,
                       q1 = q1,
                       q2 = q2,
                       q3 = q3,
                       q4 = q4,
                       ytd = ytd)
  return(df_tmp)
}


# * * * 1.2.4.6. Changes in Activity Section: Progressing at Least one Goal ----
fnGetProgressingAtLeastOneGoal <- function(df_ct){
  # Initialise the variables to receive the data
  curr_month <- as.integer(NA)
  q1 <- as.integer(NA)
  q2 <- as.integer(NA)
  q3 <- as.integer(NA)
  q4 <- as.integer(NA)
  ytd <- as.integer(NA)
  
  # Filter the input data frame to the selected cohort
  
  # Anyone with a ct_end within the period and without a ct_closure in the exception list in the ini file 
  # i.e. 'Declined', 'Non-engagement', 'Exempt', 'Incorrect contact details' see ini file [caseload_tracker_exclusions]
  # section settings for the list of ct_closure exclusions
  df_tmp <- df_ct %>% dplyr::filter(!(ct_closure %in% unlist(unname(ini_file_settings$caseload_tracker_exclusions))))
  
  # and who completed at least one goal on exit
  df_tmp <- df_tmp %>% dplyr::filter(ct_goals_out >= 1)
  
  # Trim the data to end at the end of the current month
  df_tmp <- df_tmp %>% dplyr::filter(ct_end <= dt_current_month %m+% months(1))
  
  # Calculate for each period
  curr_month <- df_tmp %>% dplyr::filter((ct_end >= dt_current_month) & 
                                           (ct_end < dt_current_month %m+% months(1))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    q1 <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start) & 
                                     (ct_end < dt_year_start %m+% months(3))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(3))
    q2 <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start %m+% months(3)) & 
                                     (ct_end < dt_year_start %m+% months(6))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(6))
    q3 <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start %m+% months(6)) & 
                                     (ct_end < dt_year_start %m+% months(9))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(9))
    q4 <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start %m+% months(9)) & 
                                     (ct_end < dt_year_start %m+% months(12))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    ytd <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start) & 
                                      (ct_end < dt_year_start %m+% months(12))) %>% NROW()
  df_tmp <- data.frame(metric = 'Completed At Least One Goal at End of Support',
                       current_month = curr_month,
                       q1 = q1,
                       q2 = q2,
                       q3 = q3,
                       q4 = q4,
                       ytd = ytd)
  return(df_tmp)
}

# * * * 1.2.4.7. Changes in Activity Section: Clients Ending Support ----
fnGetClientsEndingSupport <- function(df_ct){
  # Initialise the variables to receive the data
  curr_month <- as.integer(NA)
  q1 <- as.integer(NA)
  q2 <- as.integer(NA)
  q3 <- as.integer(NA)
  q4 <- as.integer(NA)
  ytd <- as.integer(NA)
  
  # Filter the input data frame to the selected cohort
  
  # Anyone with a ct_end within the period and without a ct_closure in the exception list in the ini file 
  # i.e. 'Declined', 'Non-engagement', 'Exempt', 'Incorrect contact details' see ini file [caseload_tracker_exclusions]
  # section settings for the list of ct_closure exclusions
  df_tmp <- df_ct %>% dplyr::filter(!(ct_closure %in% unlist(unname(ini_file_settings$caseload_tracker_exclusions))))
  
  # Trim the data to end at the end of the current month
  df_tmp <- df_tmp %>% dplyr::filter(ct_end <= dt_current_month %m+% months(1))
  
  # Calculate for each period
  curr_month <- df_tmp %>% dplyr::filter((ct_end >= dt_current_month) & 
                                           (ct_end < dt_current_month %m+% months(1))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    q1 <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start) & 
                                     (ct_end < dt_year_start %m+% months(3))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(3))
    q2 <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start %m+% months(3)) & 
                                     (ct_end < dt_year_start %m+% months(6))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(6))
    q3 <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start %m+% months(6)) & 
                                     (ct_end < dt_year_start %m+% months(9))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(9))
    q4 <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start %m+% months(9)) & 
                                     (ct_end < dt_year_start %m+% months(12))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    ytd <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start) & 
                                      (ct_end < dt_year_start %m+% months(12))) %>% NROW()
  df_tmp <- data.frame(metric = 'Clients Ending Support',
                       current_month = curr_month,
                       q1 = q1,
                       q2 = q2,
                       q3 = q3,
                       q4 = q4,
                       ytd = ytd)
  return(df_tmp)
}

# * * * 1.2.4.8. Changes in Activity Section: People Reporting a Positive Experience ----

# * * 1.2.5. KPI Section ----

# * * * 1.2.5.1. KPI Section: 80% of New Clients have an Entry WEMWBS Score ----
fnGetEntryWEMWBS <- function(df_ct){
  # Initialise the variables to receive the data
  curr_month <- as.character(NA)
  q1 <- as.character(NA)
  q2 <- as.character(NA)
  q3 <- as.character(NA)
  q4 <- as.character(NA)
  ytd <- as.character(NA)
  
  curr_month_denominator <- curr_month_numerator <- NA 
  q1_denominator <- q1_numerator <- NA
  q2_denominator <- q2_numerator <- NA
  q3_denominator <- q3_numerator <- NA
  q4_denominator <- q4_numerator <- NA
  ytd_denominator <- ytd_numerator <- NA
  
  # Filter the input data frame to the selected cohort
  
  # Anyone with a ct_start within the period but without a ct_closure in the exception list in the ini file 
  # i.e. 'Declined', 'Non-engagement', 'Exempt', 'Incorrect contact details' see ini file [caseload_tracker_exclusions]
  # section settings for the list of ct_closure exclusions
  df_tmp <- df_ct %>% dplyr::filter(!(ct_closure %in% unlist(unname(ini_file_settings$caseload_tracker_exclusions))))

  # Trim the data to end at the end of the current month
  df_tmp <- df_tmp %>% dplyr::filter(ct_start <= dt_current_month %m+% months(1))
  
  # Calculate denominator for each period
  curr_month_denominator <- df_tmp %>% dplyr::filter((ct_start >= dt_current_month) & 
                                           (ct_start < dt_current_month %m+% months(1))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    q1_denominator <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start) & 
                                     (ct_start < dt_year_start %m+% months(3))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(3))
    q2_denominator <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start %m+% months(3)) & 
                                     (ct_start < dt_year_start %m+% months(6))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(6))
    q3_denominator <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start %m+% months(6)) & 
                                     (ct_start < dt_year_start %m+% months(9))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(9))
    q4_denominator <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start %m+% months(9)) & 
                                     (ct_start < dt_year_start %m+% months(12))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    ytd_denominator <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start) & 
                                      (ct_start < dt_year_start %m+% months(12))) %>% NROW()
  # Trim to only valid WEMWBS entries
  df_tmp <- df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_in))
  
  # Calculate numerator for each period
  curr_month_numerator <- df_tmp %>% dplyr::filter((ct_start >= dt_current_month) & 
                                                       (ct_start < dt_current_month %m+% months(1))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    q1_numerator <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start) & 
                                                 (ct_start < dt_year_start %m+% months(3))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(3))
    q2_numerator <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start %m+% months(3)) & 
                                                 (ct_start < dt_year_start %m+% months(6))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(6))
    q3_numerator <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start %m+% months(6)) & 
                                                 (ct_start < dt_year_start %m+% months(9))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(9))
    q4_numerator <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start %m+% months(9)) & 
                                                 (ct_start < dt_year_start %m+% months(12))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    ytd_numerator <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start) & 
                                                  (ct_start < dt_year_start %m+% months(12))) %>% NROW()
  
  df_tmp <- data.frame(metric = 'Valid Entry WEMWBS',
                       current_month = sprintf('%.1f%% (%d/%d)', (curr_month_numerator / curr_month_denominator) * 100, curr_month_numerator, curr_month_denominator),
                       q1 = sprintf('%.1f%% (%d/%d)', (q1_numerator / q1_denominator) * 100, q1_numerator, q1_denominator),
                       q2 = sprintf('%.1f%% (%d/%d)', (q2_numerator / q2_denominator) * 100, q2_numerator, q2_denominator),
                       q3 = sprintf('%.1f%% (%d/%d)', (q3_numerator / q3_denominator) * 100, q3_numerator, q3_denominator),
                       q4 = sprintf('%.1f%% (%d/%d)', (q4_numerator / q4_denominator) * 100, q4_numerator, q4_denominator),
                       ytd = sprintf('%.1f%% (%d/%d)', (ytd_numerator / ytd_denominator) * 100, ytd_numerator, ytd_denominator))
  return(df_tmp)
}

# * * * 1.2.5.2. KPI Section: 80% of Closed Clients have an Exit WEMWBS Score ---- 
fnGetExitWEMWBS <- function(df_ct){
  # Initialise the variables to receive the data
  curr_month <- as.character(NA)
  q1 <- as.character(NA)
  q2 <- as.character(NA)
  q3 <- as.character(NA)
  q4 <- as.character(NA)
  ytd <- as.character(NA)
  
  curr_month_denominator <- curr_month_numerator <- NA 
  q1_denominator <- q1_numerator <- NA
  q2_denominator <- q2_numerator <- NA
  q3_denominator <- q3_numerator <- NA
  q4_denominator <- q4_numerator <- NA
  ytd_denominator <- ytd_numerator <- NA
  
  # Filter the input data frame to the selected cohort
  
  # Anyone with a ct_start within the period but without a ct_closure in the exception list in the ini file 
  # i.e. 'Declined', 'Non-engagement', 'Exempt', 'Incorrect contact details' see ini file [caseload_tracker_exclusions]
  # section settings for the list of ct_closure exclusions
  df_tmp <- df_ct %>% dplyr::filter(!(ct_closure %in% unlist(unname(ini_file_settings$caseload_tracker_exclusions))))
  
  # Trim the data to end at the end of the current month
  df_tmp <- df_tmp %>% dplyr::filter(ct_start <= dt_current_month %m+% months(1))
  
  # Calculate denominator for each period
  curr_month_denominator <- df_tmp %>% dplyr::filter((ct_end >= dt_current_month) & 
                                           (ct_end < dt_current_month %m+% months(1))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    q1_denominator <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start) & 
                                     (ct_end < dt_year_start %m+% months(3))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(3))
    q2_denominator <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start %m+% months(3)) & 
                                     (ct_end < dt_year_start %m+% months(6))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(6))
    q3_denominator <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start %m+% months(6)) & 
                                     (ct_end < dt_year_start %m+% months(9))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(9))
    q4_denominator <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start %m+% months(9)) & 
                                     (ct_end < dt_year_start %m+% months(12))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    ytd_denominator <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start) & 
                                      (ct_end < dt_year_start %m+% months(12))) %>% NROW()
  # Trim to only valid WEMWBS entries
  df_tmp <- df_tmp %>% dplyr::filter(!is.na(ct_wemwbs_score_out))
  
  # Calculate numerator for each period
  curr_month_numerator <- df_tmp %>% dplyr::filter((ct_end >= dt_current_month) & 
                                                       (ct_end < dt_current_month %m+% months(1))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    q1_numerator <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start) & 
                                                 (ct_end < dt_year_start %m+% months(3))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(3))
    q2_numerator <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start %m+% months(3)) & 
                                                 (ct_end < dt_year_start %m+% months(6))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(6))
    q3_numerator <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start %m+% months(6)) & 
                                                 (ct_end < dt_year_start %m+% months(9))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(9))
    q4_numerator <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start %m+% months(9)) & 
                                                 (ct_end < dt_year_start %m+% months(12))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    ytd_numerator <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start) & 
                                                  (ct_end < dt_year_start %m+% months(12))) %>% NROW()  
  df_tmp <- data.frame(metric = 'Valid Exit WEMWBS',
                       current_month = sprintf('%.1f%% (%d/%d)', (curr_month_numerator / curr_month_denominator) * 100, curr_month_numerator, curr_month_denominator),
                       q1 = sprintf('%.1f%% (%d/%d)', (q1_numerator / q1_denominator) * 100, q1_numerator, q1_denominator),
                       q2 = sprintf('%.1f%% (%d/%d)', (q2_numerator / q2_denominator) * 100, q2_numerator, q2_denominator),
                       q3 = sprintf('%.1f%% (%d/%d)', (q3_numerator / q3_denominator) * 100, q3_numerator, q3_denominator),
                       q4 = sprintf('%.1f%% (%d/%d)', (q4_numerator / q4_denominator) * 100, q4_numerator, q4_denominator),
                       ytd = sprintf('%.1f%% (%d/%d)', (ytd_numerator / ytd_denominator) * 100, ytd_numerator, ytd_denominator))
  return(df_tmp)
}

# * * * 1.2.5.3. KPI Section: 80% of New Clients have a Entry Loneliness Answer (Yes/No) ---- 
fnGetEntryLoneliness <- function(df_ct){
  # Initialise the variables to receive the data
  curr_month <- as.integer(NA)
  q1 <- as.integer(NA)
  q2 <- as.integer(NA)
  q3 <- as.integer(NA)
  q4 <- as.integer(NA)
  ytd <- as.integer(NA)

  curr_month_denominator <- curr_month_numerator <- NA 
  q1_denominator <- q1_numerator <- NA
  q2_denominator <- q2_numerator <- NA
  q3_denominator <- q3_numerator <- NA
  q4_denominator <- q4_numerator <- NA
  ytd_denominator <- ytd_numerator <- NA
  
  # Filter the input data frame to the selected cohort
  
  # Anyone with a ct_start within the period but without a ct_closure in the exception list in the ini file 
  # i.e. 'Declined', 'Non-engagement', 'Exempt', 'Incorrect contact details' see ini file [caseload_tracker_exclusions]
  # section settings for the list of ct_closure exclusions
  df_tmp <- df_ct %>% dplyr::filter(!(ct_closure %in% unlist(unname(ini_file_settings$caseload_tracker_exclusions))))
  
  # Trim the data to end at the end of the current month
  df_tmp <- df_tmp %>% dplyr::filter(ct_start <= dt_current_month %m+% months(1))
  
  # Calculate denominator for each period
  curr_month_denominator <- df_tmp %>% dplyr::filter((ct_start >= dt_current_month) & 
                                                       (ct_start < dt_current_month %m+% months(1))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    q1_denominator <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start) & 
                                                 (ct_start < dt_year_start %m+% months(3))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(3))
    q2_denominator <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start %m+% months(3)) & 
                                                 (ct_start < dt_year_start %m+% months(6))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(6))
    q3_denominator <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start %m+% months(6)) & 
                                                 (ct_start < dt_year_start %m+% months(9))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(9))
    q4_denominator <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start %m+% months(9)) & 
                                                 (ct_start < dt_year_start %m+% months(12))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    ytd_denominator <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start) & 
                                                  (ct_start < dt_year_start %m+% months(12))) %>% NROW()
  # Trim to only valid WEMWBS entries
  df_tmp <- df_tmp %>% dplyr::filter(ct_loneliness_in %in% c('Yes','No'))
  
  # Calculate numerator for each period
  curr_month_numerator <- df_tmp %>% dplyr::filter((ct_start >= dt_current_month) & 
                                                     (ct_start < dt_current_month %m+% months(1))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    q1_numerator <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start) & 
                                               (ct_start < dt_year_start %m+% months(3))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(3))
    q2_numerator <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start %m+% months(3)) & 
                                               (ct_start < dt_year_start %m+% months(6))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(6))
    q3_numerator <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start %m+% months(6)) & 
                                               (ct_start < dt_year_start %m+% months(9))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(9))
    q4_numerator <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start %m+% months(9)) & 
                                               (ct_start < dt_year_start %m+% months(12))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    ytd_numerator <- df_tmp %>% dplyr::filter((ct_start >= dt_year_start) & 
                                                (ct_start < dt_year_start %m+% months(12))) %>% NROW()
  
  df_tmp <- data.frame(metric = 'Valid Entry Loneliness',
                       current_month = sprintf('%.1f%% (%d/%d)', (curr_month_numerator / curr_month_denominator) * 100, curr_month_numerator, curr_month_denominator),
                       q1 = sprintf('%.1f%% (%d/%d)', (q1_numerator / q1_denominator) * 100, q1_numerator, q1_denominator),
                       q2 = sprintf('%.1f%% (%d/%d)', (q2_numerator / q2_denominator) * 100, q2_numerator, q2_denominator),
                       q3 = sprintf('%.1f%% (%d/%d)', (q3_numerator / q3_denominator) * 100, q3_numerator, q3_denominator),
                       q4 = sprintf('%.1f%% (%d/%d)', (q4_numerator / q4_denominator) * 100, q4_numerator, q4_denominator),
                       ytd = sprintf('%.1f%% (%d/%d)', (ytd_numerator / ytd_denominator) * 100, ytd_numerator, ytd_denominator))
  return(df_tmp)
}

# * * * 1.2.5.4. KPI Section: 80% of Closed Clients have a Exit Loneliness Answer (Yes/No) ----
fnGetExitLoneliness <- function(df_ct){
  # Initialise the variables to receive the data
  curr_month <- as.character(NA)
  q1 <- as.character(NA)
  q2 <- as.character(NA)
  q3 <- as.character(NA)
  q4 <- as.character(NA)
  ytd <- as.character(NA)
  
  curr_month_denominator <- curr_month_numerator <- NA 
  q1_denominator <- q1_numerator <- NA
  q2_denominator <- q2_numerator <- NA
  q3_denominator <- q3_numerator <- NA
  q4_denominator <- q4_numerator <- NA
  ytd_denominator <- ytd_numerator <- NA
  
  # Filter the input data frame to the selected cohort
  
  # Anyone with a ct_start within the period but without a ct_closure in the exception list in the ini file 
  # i.e. 'Declined', 'Non-engagement', 'Exempt', 'Incorrect contact details' see ini file [caseload_tracker_exclusions]
  # section settings for the list of ct_closure exclusions
  df_tmp <- df_ct %>% dplyr::filter(!(ct_closure %in% unlist(unname(ini_file_settings$caseload_tracker_exclusions))))
  
  # Trim the data to end at the end of the current month
  df_tmp <- df_tmp %>% dplyr::filter(ct_start <= dt_current_month %m+% months(1))
  
  # Calculate denominator for each period
  curr_month_denominator <- df_tmp %>% dplyr::filter((ct_end >= dt_current_month) & 
                                                       (ct_end < dt_current_month %m+% months(1))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    q1_denominator <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start) & 
                                                 (ct_end < dt_year_start %m+% months(3))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(3))
    q2_denominator <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start %m+% months(3)) & 
                                                 (ct_end < dt_year_start %m+% months(6))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(6))
    q3_denominator <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start %m+% months(6)) & 
                                                 (ct_end < dt_year_start %m+% months(9))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(9))
    q4_denominator <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start %m+% months(9)) & 
                                                 (ct_end < dt_year_start %m+% months(12))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    ytd_denominator <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start) & 
                                                  (ct_end < dt_year_start %m+% months(12))) %>% NROW()
  # Trim to only valid WEMWBS entries
  df_tmp <- df_tmp %>% dplyr::filter(!is.na(ct_loneliness_out))
  
  # Calculate numerator for each period
  curr_month_numerator <- df_tmp %>% dplyr::filter((ct_end >= dt_current_month) & 
                                                     (ct_end < dt_current_month %m+% months(1))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    q1_numerator <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start) & 
                                               (ct_end < dt_year_start %m+% months(3))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(3))
    q2_numerator <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start %m+% months(3)) & 
                                               (ct_end < dt_year_start %m+% months(6))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(6))
    q3_numerator <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start %m+% months(6)) & 
                                               (ct_end < dt_year_start %m+% months(9))) %>% NROW()
  if(dt_current_month >= dt_year_start %m+% months(9))
    q4_numerator <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start %m+% months(9)) & 
                                               (ct_end < dt_year_start %m+% months(12))) %>% NROW()
  if(dt_current_month >= dt_year_start)
    ytd_numerator <- df_tmp %>% dplyr::filter((ct_end >= dt_year_start) & 
                                                (ct_end < dt_year_start %m+% months(12))) %>% NROW()  
  df_tmp <- data.frame(metric = 'Valid Exit Loneliness',
                       current_month = sprintf('%.1f%% (%d/%d)', (curr_month_numerator / curr_month_denominator) * 100, curr_month_numerator, curr_month_denominator),
                       q1 = sprintf('%.1f%% (%d/%d)', (q1_numerator / q1_denominator) * 100, q1_numerator, q1_denominator),
                       q2 = sprintf('%.1f%% (%d/%d)', (q2_numerator / q2_denominator) * 100, q2_numerator, q2_denominator),
                       q3 = sprintf('%.1f%% (%d/%d)', (q3_numerator / q3_denominator) * 100, q3_numerator, q3_denominator),
                       q4 = sprintf('%.1f%% (%d/%d)', (q4_numerator / q4_denominator) * 100, q4_numerator, q4_denominator),
                       ytd = sprintf('%.1f%% (%d/%d)', (ytd_numerator / ytd_denominator) * 100, ytd_numerator, ytd_denominator))
  return(df_tmp)
}

# * * 1.2.6. Data Points Section ----

# * * * 1.2.6.1. Data Points Section: Create main table ----
fnGetDataPoints <- function(df_dp, metric_list){
  df_metrics <- data.frame(metric = metric_list)
  
  # Filter the data points to the selected metrics and up to and including the current month
  df_tmp <- df_dp %>% 
    dplyr::filter(metric %in% metric_list) %>%
    dplyr::filter(month <= dt_current_month %m+% months(1)) %>%
    group_by(month, metric) %>%
    summarise(value = sum(value, na.rm = TRUE),
              .groups = 'keep') %>%
    ungroup()
  
  # Create the data points table
  df <- df_tmp %>% 
    dplyr::filter(month == dt_current_month) %>%
    group_by(metric) %>%
    summarise(current_month = sum(value, na.rm = TRUE)) %>%
    ungroup() %>%
    full_join(
      df_tmp %>% 
        dplyr::filter(month >= dt_year_start & 
                        month < dt_year_start %m+% months(3)) %>%
        group_by(metric) %>%
        summarise(q1 = sum(value, na.rm = TRUE)) %>%
        ungroup(),
      by = 'metric'
    ) %>%
    full_join(
      df_tmp %>% 
        dplyr::filter(month >= dt_year_start %m+% months(3) & 
                        month < dt_year_start %m+% months(6)) %>%
        group_by(metric) %>%
        summarise(q2 = sum(value, na.rm = TRUE)) %>%
        ungroup(),
      by = 'metric'
    ) %>%
    full_join(
      df_tmp %>% 
        dplyr::filter(month >= dt_year_start %m+% months(6) & 
                        month < dt_year_start %m+% months(9)) %>%
        group_by(metric) %>%
        summarise(q3 = sum(value, na.rm = TRUE)) %>%
        ungroup(),
      by = 'metric'
    ) %>%
    full_join(
      df_tmp %>% 
        dplyr::filter(month >= dt_year_start %m+% months(9) & 
                        month < dt_year_start %m+% months(12)) %>%
        group_by(metric) %>%
        summarise(q4 = sum(value, na.rm = TRUE)) %>%
        ungroup(),
      by = 'metric'
    ) %>%
    full_join(
      df_tmp %>% 
        dplyr::filter(month >= dt_year_start & 
                        month < dt_year_start %m+% months(12)) %>%
        group_by(metric) %>%
        summarise(ytd = sum(value, na.rm = TRUE)) %>%
        ungroup(),
      by = 'metric'
    )

  # Ensure all metrics are represented and any NA are replaced with zero (for valid periods)  
  df <- df %>% full_join(df_metrics, by = 'metric')  
  replace_list <- list(current_month = 0, q1 = 0, ytd = 0)
  if(dt_current_month >= dt_year_start %m+% months(3))
    replace_list <- append(replace_list, list(q2 = 0))
  if(dt_current_month >= dt_year_start %m+% months(6))
    replace_list <- append(replace_list, list(q3 = 0))
  if(dt_current_month >= dt_year_start %m+% months(9))
    replace_list <- append(replace_list, list(q4 = 0))
  df <- df %>% replace_na(replace_list)  
  
  # Return the data frame
  return(df)
}

# * * * 1.2.6.2. Data Points Section: Get case load ----
fnGetCaseLoad <- function(df_ct){
  # Initialise the variables to receive the data
  curr_month <- as.integer(NA)
  q1 <- as.integer(NA)
  q2 <- as.integer(NA)
  q3 <- as.integer(NA)
  q4 <- as.integer(NA)
  ytd <- as.integer(NA)

  # Filter the data to exclude any records starting after end of current month
  df_tmp <- df_ct %>% dplyr::filter(ct_start < dt_current_month %m+% months(1))
  
  # Get any record where the ct_start occurs before end of period and end date occurs 
  # after the end of the period or is NA (NB: census point data) no exclusions applied to 
  # this data
  
  curr_month <- df_tmp %>% dplyr::filter(ct_start < dt_current_month %m+% months(1) &
                                           (is.na(ct_end) | ct_end >= dt_current_month %m+% months(1))) %>% NROW()
    
  if(dt_current_month >= dt_year_start)
    q1 <- df_tmp %>% dplyr::filter(ct_start < dt_year_start %m+% months(3) &
                                           (is.na(ct_end) | ct_end >= dt_year_start %m+% months(3))) %>% NROW()
  if(dt_current_month >= dt_year_start  %m+% months(3))
  q2 <- df_tmp %>% dplyr::filter(ct_start < dt_year_start %m+% months(6) &
                                   (is.na(ct_end) | ct_end >= dt_year_start %m+% months(6))) %>% NROW()

  if(dt_current_month >= dt_year_start %m+% months(6))
    q3 <- df_tmp %>% dplyr::filter(ct_start < dt_year_start %m+% months(9) &
                                   (is.na(ct_end) | ct_end >= dt_year_start %m+% months(9))) %>% NROW()
  
  if(dt_current_month >= dt_year_start %m+% months(9))
    q4 <- df_tmp %>% dplyr::filter(ct_start < dt_year_start %m+% months(9) &
                                   (is.na(ct_end) | ct_end >= dt_year_start %m+% months(9))) %>% NROW()
  
  if(dt_current_month >= dt_year_start)
    ytd <- df_tmp %>% dplyr::filter(ct_start < dt_year_start %m+% months(9) &
                                   (is.na(ct_end) | ct_end >= dt_year_start %m+% months(9))) %>% NROW()
  
  df <- data.frame(metric = 'Case Load at Census', current_month = curr_month,
                   q1, q2, q3, q4, ytd)
  return(df)
}

# * * 1.2.7. Support Provided Section ----

# * * * 1.2.7.1. Support Provided Section: Create main table ----
fnGetSupportProvided <- function(df_sr, metric_list){
  df_metrics <- data.frame(metric = metric_list)

  # Filter the data points to the selected metrics and up to and including the current month
  df_tmp <- df_sr %>% 
    mutate(metric = support) %>% 
    dplyr::filter(metric %in% metric_list) %>%
    dplyr::filter(month <= dt_current_month %m+% months(1)) %>%
    group_by(month, metric) %>%
    summarise(value = n(),
              .groups = 'keep') %>%
    ungroup()
  
  # Create the data points table
  df <- df_tmp %>% 
    dplyr::filter(month == dt_current_month) %>%
    group_by(metric) %>%
    summarise(current_month = sum(value, na.rm = TRUE)) %>%
    ungroup() %>%
    full_join(
      df_tmp %>% 
        dplyr::filter(month >= dt_year_start & 
                        month < dt_year_start %m+% months(3)) %>%
        group_by(metric) %>%
        summarise(q1 = sum(value, na.rm = TRUE)) %>%
        ungroup(),
      by = 'metric'
    ) %>%
    full_join(
      df_tmp %>% 
        dplyr::filter(month >= dt_year_start %m+% months(3) & 
                        month < dt_year_start %m+% months(6)) %>%
        group_by(metric) %>%
        summarise(q2 = sum(value, na.rm = TRUE)) %>%
        ungroup(),
      by = 'metric'
    ) %>%
    full_join(
      df_tmp %>% 
        dplyr::filter(month >= dt_year_start %m+% months(6) & 
                        month < dt_year_start %m+% months(9)) %>%
        group_by(metric) %>%
        summarise(q3 = sum(value, na.rm = TRUE)) %>%
        ungroup(),
      by = 'metric'
    ) %>%
    full_join(
      df_tmp %>% 
        dplyr::filter(month >= dt_year_start %m+% months(9) & 
                        month < dt_year_start %m+% months(12)) %>%
        group_by(metric) %>%
        summarise(q4 = sum(value, na.rm = TRUE)) %>%
        ungroup(),
      by = 'metric'
    ) %>%
    full_join(
      df_tmp %>% 
        dplyr::filter(month >= dt_year_start & 
                        month < dt_year_start %m+% months(12)) %>%
        group_by(metric) %>%
        summarise(ytd = sum(value, na.rm = TRUE)) %>%
        ungroup(),
      by = 'metric'
    )

  # Ensure all metrics are represented and any NA are replaced with zero (for valid periods)  
  df <- df %>% full_join(df_metrics, by = 'metric')  
  replace_list <- list(current_month = 0, q1 = 0, ytd = 0)
  if(dt_current_month >= dt_year_start %m+% months(3))
    replace_list <- append(replace_list, list(q2 = 0))
  if(dt_current_month >= dt_year_start %m+% months(6))
    replace_list <- append(replace_list, list(q3 = 0))
  if(dt_current_month >= dt_year_start %m+% months(9))
    replace_list <- append(replace_list, list(q4 = 0))
  df <- df %>% replace_na(replace_list)  
  
  # Return the data frame
  return(df)
}

# * * 1.2.8. Outputs Section ----

# * * * 1.2.8.1. Outputs 3m Section ----
fnOutputs3mSection <- function(df_op){
  df_tmp <- df_op %>% 
    dplyr::filter(
      as.Date(month) >= (dt_current_month %m+% months(-2)) & 
        as.Date(month) < (dt_current_month %m+% months(1)) &
        as.Date(month) >= dt_year_start) %>%
  group_by(section) %>%
  summarise(volume = n()) %>%
  ungroup()
  
  df <- fnGlenday(df_tmp, var.x = 'section', var.y = 'volume')

  return(df)
}

# * * * 1.2.8.2. Outputs 12m Section ----
fnOutputs12mSection <- function(df_op){
  df_tmp <- df_op %>% 
    dplyr::filter(
      as.Date(month) >= dt_year_start & 
        as.Date(month) < (dt_current_month %m+% months(1)) &
        as.Date(month) < (dt_year_start %m+% months(12))) %>%
    group_by(section) %>%
    summarise(volume = n()) %>%
    ungroup()
  
  df <- fnGlenday(df_tmp, var.x = 'section', var.y = 'volume')
  
  return(df)
}

# * * * 1.2.8.3. Outputs 3m Section and Metric ----
fnOutputs3mSectionMetric <- function(df_op){
  df_tmp <- df_op %>% 
    dplyr::filter(
      as.Date(month) >= (dt_current_month %m+% months(-2)) & 
        as.Date(month) < (dt_current_month %m+% months(1)) &
        as.Date(month) >= dt_year_start) %>%
    group_by(section, output) %>%
    summarise(volume = n(), .groups = 'keep') %>%
    ungroup()
  
  df <- fnGlenday(df_tmp, var.x = 'output', var.y = 'volume')
  
  return(df)
}

# * * * 1.2.8.4. Outputs 12m Section and Metric ----
fnOutputs12mSectionMetric <- function(df_op){
  df_tmp <- df_op %>% 
    dplyr::filter(
      as.Date(month) >= dt_year_start & 
        as.Date(month) < (dt_current_month %m+% months(1)) &
        as.Date(month) < (dt_year_start %m+% months(12))) %>%
    group_by(section, output) %>%
    summarise(volume = n(), .groups = 'keep') %>%
    ungroup()
  
  df <- fnGlenday(df_tmp, var.x = 'output', var.y = 'volume')
  
  return(df)
}

# * * 1.2.9. Outcomes Section ----

# * * * 1.2.9.1. Outcomes 3m Section ----
fnOutcomes3mSection <- function(df_oc){
  df_tmp <- df_oc %>% 
    dplyr::filter(
      as.Date(month) >= (dt_current_month %m+% months(-2)) & 
        as.Date(month) < (dt_current_month %m+% months(1)) &
        as.Date(month) >= dt_year_start) %>%
    group_by(section) %>%
    summarise(volume = n()) %>%
    ungroup()
  
  df <- fnGlenday(df_tmp, var.x = 'section', var.y = 'volume')
  
  return(df)
}

# * * * 1.2.9.2. Outcomes 12m Section ----
fnOutcomes12mSection <- function(df_oc){
  df_tmp <- df_oc %>% 
    dplyr::filter(
      as.Date(month) >= dt_year_start & 
        as.Date(month) < (dt_current_month %m+% months(1)) &
        as.Date(month) < (dt_year_start %m+% months(12))) %>%
    group_by(section) %>%
    summarise(volume = n()) %>%
    ungroup()
  
  df <- fnGlenday(df_tmp, var.x = 'section', var.y = 'volume')
  
  return(df)
}

# * * * 1.2.9.3. Outcomes 3m Section and Metric ----
fnOutcomes3mSectionMetric <- function(df_oc){
  df_tmp <- df_oc %>% 
    dplyr::filter(
      as.Date(month) >= (dt_current_month %m+% months(-2)) & 
        as.Date(month) < (dt_current_month %m+% months(1)) &
        as.Date(month) >= dt_year_start) %>%
    group_by(section, outcome) %>%
    summarise(volume = n(), .groups = 'keep') %>%
    ungroup()
  
  df <- fnGlenday(df_tmp, var.x = 'outcome', var.y = 'volume')
  
  return(df)
}

# * * * 1.2.9.4. Outcomes 12m Section and Metric ----
fnOutcomes12mSectionMetric <- function(df_oc){
  df_tmp <- df_oc %>% 
    dplyr::filter(
      as.Date(month) >= dt_year_start & 
        as.Date(month) < (dt_current_month %m+% months(1)) &
        as.Date(month) < (dt_year_start %m+% months(12))) %>%
    group_by(section, outcome) %>%
    summarise(volume = n(), .groups = 'keep') %>%
    ungroup()
  
  df <- fnGlenday(df_tmp, var.x = 'outcome', var.y = 'volume')
  
  return(df)
}


# 2. Import data ----
# ═══════════════════
ini_file_settings <- read.ini(ini_file)

# Select the caseload tracker sheet from the excel workbook
caseload_tracker_sheets <- readxl::excel_sheets(path = caseload_tracker_file)
caseload_tracker_sheets <- utils::select.list(title = 'CASELOAD TRACKER sheet(s)', choices = caseload_tracker_sheets, multiple = TRUE, graphics = TRUE)

# * 2.1. Caseload tracker ----
# ────────────────────────────

df_caseload_tracker <- fnImportCaseloadTracker(path = caseload_tracker_file,
                                               sheets = caseload_tracker_sheets)

fnCaseloadTrackerDataQuality(df_caseload_tracker)

# * 2.2. Reporting Workbooks ----
# ───────────────────────────────

# Initialise the data frames to receive the data
df_data_points <- data.frame()
df_support_and_referrals <- data.frame()
df_outputs <- data.frame()
df_outcomes <- data.frame()

# Loop through each reporting workbook
for(f in reporting_workbook_filelist){
  df_data_points <- df_data_points %>% bind_rows(fnImportReportingWorkbook_DataPoints(path = f))
  df_support_and_referrals <- df_support_and_referrals %>% bind_rows(fnImportReportingWorkbook_SupportReferrals(path = f))
  df_outputs <- df_outputs %>% bind_rows(fnImportReportingWorkbook_Outputs(path = f))
  df_outcomes <- df_outcomes %>% bind_rows(fnImportReportingWorkbook_Outcomes(path = f))
}

# * 2.3. Activity Workbook ----
# ─────────────────────────────
df_activity <- fnImportActivityWorkbook(path = activity_file)

# 3. Process Data ----
# ════════════════════

# * 3.1. Headline Section ----
# ────────────────────────────

# New Clients Supported in Period
# Clients Supported in Period
# Previous 3 Month Activity for Clients Supported in Period:
#   ED Attendances*
#   Emergency Admissions*
#   Ambulance Conveyances*
# Previous 12 Month Activity for Clients Supported in Period:
#   ED Attendances*
#   Emergency Admissions*
#   Ambulance Conveyances*
# ──────────────────────────────────────────────────────────────────────────────────
# NOTE: The sections marked * will need to be sourced from the RDUH BI team's report
# ──────────────────────────────────────────────────────────────────────────────────

df_headline_section <- data.frame(metric = as.character(),
                                  current_month = as.numeric(),
                                  q1 = as.numeric(),
                                  q2 = as.numeric(),
                                  q3 = as.numeric(),
                                  q4 = as.numeric(),
                                  ytd = as.numeric())

# New Clients Supported
df_headline_section <- df_headline_section %>% bind_rows(fnGetNewClientsSupported(df_caseload_tracker))

# Clients Supported
df_headline_section <- df_headline_section %>% bind_rows(fnGetClientsSupported(df_caseload_tracker))

# Previous Activity 3 and 12 Month: ED | EM | AMB 
# # Currently we are awaiting the RDUH BI Team's report to populate the following metrics
# df_tmp <- data.frame(metric = c('Previous 3 Month Activity for Clients Supported in Period: ED Attendances',
#                                 'Previous 3 Month Activity for Clients Supported in Period: Emergency Admissions',
#                                 'Previous 3 Month Activity for Clients Supported in Period: Ambulance Conveyances',
#                                 'Previous 12 Month Activity for Clients Supported in Period: ED Attendances',
#                                 'Previous 12 Month Activity for Clients Supported in Period: Emergency Admissions',
#                                 'Previous 12 Month Activity for Clients Supported in Period: Ambulance Conveyances'),
#                      current_month = rep(NA, 6),
#                      q1 = rep(NA, 6),
#                      q2 = rep(NA, 6),
#                      q3 = rep(NA, 6),
#                      q4 = rep(NA, 6),
#                      ytd = rep(NA, 6))
df_tmp <- df_activity[25:31,] %>%
  mutate(across(.cols = 2:7, .fns = as.integer))

df_headline_section <- df_headline_section %>% bind_rows(df_tmp)

# * 3.2. Changes in Activity Section ----
# ───────────────────────────────────────

# New Clients Supported in Period
# Reduction in Activity Starting 3 months from Intervention Start (NHSE Target 40%):
#   ED Attendances*
#   Emergency Admissions*
#   Ambulance Conveyances*
# Reduction in Activity 12 months Prior vs. 12 months Post Intervention (OND Target 40%)
#   ED Attendances*
#   Emergency Admissions*
#   Ambulance Conveyances*
# Reduction in People Experiencing Loneliness at End of Support (NHSE Target 66%)
# Clients Ending Support and Experiencing Improved Wellbeing
# People Progressing at Least one Goal (OND Target 90%)
# Clients Ending Support
# People Reporting a Positive Experience from our Support (NHSE Target 80%)**
# ─────────────────────────────────────────────────────────────────────────────────
# NOTE: The metrics marked * will need to be sourced from the RDUH BI team's report
#       The metrics marked ** are not available in the current data collections
# ─────────────────────────────────────────────────────────────────────────────────

df_activity_section <- data.frame(metric = as.character(),
                                  current_month = as.numeric(),
                                  q1 = as.numeric(),
                                  q2 = as.numeric(),
                                  q3 = as.numeric(),
                                  q4 = as.numeric(),
                                  ytd = as.numeric())

# New Clients Supported in Period
df_activity_section <- df_activity_section %>% bind_rows(fnGetNewClientsSupported(df_caseload_tracker))

# Reduction in People Experiencing Loneliness at End of Support (NHSE Target 66%)
# This makes the assumption that we will improve the experiencing loneliness support on entry and exit
df_activity_section <- df_activity_section %>% bind_rows(fnGetReductionInLoneliness(df_ct = df_caseload_tracker))

# Clients Ending Support and Experiencing Improved Wellbeing
# This makes the assumption that we will improve the WEMWBS on entry and exit
df_activity_section <- df_activity_section %>% bind_rows(fnGetIncreasedWEMWBSOnExit(df_ct = df_caseload_tracker))

# People Progressing at Least one Goal (OND Target 90%)
df_activity_section <- df_activity_section %>% bind_rows(fnGetProgressingAtLeastOneGoal(df_ct = df_caseload_tracker))

# Clients Ending Support
df_activity_section <- df_activity_section %>% bind_rows(fnGetClientsEndingSupport(df_ct = df_caseload_tracker))

# Activity Section
if(unname(project_id)==2473){
  df_activity_section_em_ed_amb <- df_activity[25:NROW(df_activity),]
} else {
  df_activity_section_em_ed_amb <- NULL
}

# People Reporting a Positive Experience from our Support (NHSE Target 80%)**
# This metric is currently not recorded

# * 3.3. Process KPIs Section ----
# ────────────────────────────────

df_kpi_section <- data.frame(metric = as.character(),
                             current_month = as.character(),
                             q1 = as.character(),
                             q2 = as.character(),
                             q3 = as.character(),
                             q4 = as.character(),
                             ytd = as.character())

# 80% of New Clients have an Entry WEMWBS Score
df_kpi_section <- df_kpi_section %>% bind_rows(fnGetEntryWEMWBS(df_ct = df_caseload_tracker))

# 80% of Closed Clients have an Exit WEMWBS Score 
df_kpi_section <- df_kpi_section %>% bind_rows(fnGetExitWEMWBS(df_ct = df_caseload_tracker))

# 80% of New Clients have a Entry Loneliness Answer (Yes/No) 
df_kpi_section <- df_kpi_section %>% bind_rows(fnGetEntryLoneliness(df_ct = df_caseload_tracker))

# 80% of Closed Clients have a Exit Loneliness Answer (Yes/No)
df_kpi_section <- df_kpi_section %>% bind_rows(fnGetExitLoneliness(df_ct = df_caseload_tracker))

# * 3.4. Data Points Section ----
# ───────────────────────────────

metric_list <- c('Number of wider beneficiaries', 'Clients who declined',
                 'Case concluded successfully', 'Closed cases due to disengagement',
                 'Closed cases due to death', 'Closed cases (other reasons, ie moving out of area)',
                 'Number of contacts/interventions with clients')

df_data_point_section <- fnGetDataPoints(df_dp = df_data_points, metric_list = metric_list) %>% 
  bind_rows(fnGetCaseLoad(df_ct = df_caseload_tracker))

# ────────────────────────────────────────────────────────────────────
# NOTE: The metrics marked * will be sourced from the caseload tracker
# ────────────────────────────────────────────────────────────────────

# * 3.5. Support Provided Section ----
# ────────────────────────────────────

metric_list <- c('Team Around the Person meeting conducted', 'Flow meeting with FC & Lead Professional',
             'One-to-one work with clients (per client) number of individual one to one interactions with client',
             'Continued ongoing contacts with professionals (total number of seperate contacts)',
             'Caseworker research undertaken to find solutions for clients', 'Caseworker support to access Personal Health Budget',
             'Caseworker support with Form filling', 'Caseworker support with IT incl. virtual meetings, emails etc',
             'Caseworker support to meet aspirations', 'Client involved in coproduction work (total number of seperate contacts)')

df_support_provided_section <- fnGetSupportProvided(df_sr = df_support_and_referrals, metric_list = metric_list)

# * 3.6. Outputs Section ----
# ───────────────────────────

df_outputs_3m_section <- fnOutputs3mSection(df_outputs)
df_outputs_12m_section <- fnOutputs12mSection(df_outputs)
df_outputs_3m_section_metric <- fnOutputs3mSectionMetric(df_outputs)
df_outputs_12m_section_metric <- fnOutputs12mSectionMetric(df_outputs)

# * 3.7. Outcomes Section ----
# ────────────────────────────

df_outcomes_3m_section <- fnOutcomes3mSection(df_outcomes)
df_outcomes_12m_section <- fnOutcomes12mSection(df_outcomes)
df_outcomes_3m_section_metric <- fnOutcomes3mSectionMetric(df_outcomes)
df_outcomes_12m_section_metric <- fnOutcomes12mSectionMetric(df_outcomes)

# * 3.8. Write data object ----
# ─────────────────────────────

save(list = c('dt_year_start', 'dt_current_month',
              'df_headline_section', 'df_activity_section', 'df_activity_section_em_ed_amb',
              'df_kpi_section', 'df_data_point_section', 'df_support_provided_section', 
              'df_outputs_3m_section', 'df_outputs_12m_section', 'df_outputs_3m_section_metric', 
              'df_outputs_12m_section_metric', 'df_outcomes_3m_section', 'df_outcomes_12m_section',
              'df_outcomes_3m_section_metric', 'df_outcomes_12m_section_metric'), file = 'data_objects.RObj')

# * 3.8. Impact Reporting Section ----
# ────────────────────────────────────

# * * 3.8.1. Impact Reporting: Load lookups and data ----
df_lookup <- read.csv(file = 'impact_reporting_api_entity_references.csv')
df_uploaded_data <- df_errors <- data.frame()

if(dlgMessage('Do you want to submit IMPACT data', 'yesno')$res=='yes'){
  # Generate UUID
  uuid <- uuid::UUIDgenerate()
  
  # * * 3.8.2. Impact Reporting: Data points ----
  df_report_data <- df_data_points %>% 
    dplyr::filter(month >= dt_current_month & month < dt_current_month %m+% months(1)) %>%
    # Group and summarise on metric
    group_by(month, metric) %>%
    summarise(value = sum(value, na.rm = TRUE),
              .groups = 'keep') %>%
    ungroup() %>%
    # Filter out any zero values
    dplyr::filter(value > 0) %>%
    # Add in the reporting_workbook_sheet and section fields
    mutate(reporting_workbook_sheet = 'Lookups_Data_Points',
           section = 'Data Points') %>%
    # Join to the impact reporting lookup data
    left_join(df_lookup, by = c('reporting_workbook_sheet', 'section', 'metric'))
  
  # Report any unmatched data to the error dataframe
  df_errors <- df_report_data %>% 
    dplyr::filter(is.na(id)) %>%
    select(reporting_workbook_sheet, section, metric, value) %>%
    mutate(uuid = uuid, .before = 1)
  
  # Select only valid data to upload
  df_report_data <- df_report_data %>% 
    dplyr::filter(!is.na(id)) %>%
    select(month, reporting_workbook_sheet, section, metric, value, id, name, classification_name, classification_value) %>%
    mutate(uuid = uuid, .before = 1)
  
  # Apply fnPostData to each row of the data points dataframe
  df_uploaded_data <- df_report_data %>% mutate(status = apply(df_report_data, 1, fnPostData, uuid))

  # * * 3.8.3. Impact Reporting: Support and referrals ----
  df_report_data <- df_support_and_referrals %>% 
    dplyr::filter(month >= dt_current_month & month < dt_current_month %m+% months(1)) %>%
    mutate(metric = support) %>% 
    # Group and summarise on metric
    group_by(month, section, metric) %>%
    summarise(value = n(),
              .groups = 'keep') %>%
    ungroup() %>%
    # Filter out any zero values
    dplyr::filter(value > 0) %>%
    # Add in the reporting_workbook_sheet and section fields
    mutate(reporting_workbook_sheet = 'Lookups_Support_Section') %>%
    # Join to the impact reporting lookup data
    left_join(df_lookup, by = c('reporting_workbook_sheet', 'section', 'metric'))
  
  # Report any unmatched data to the error dataframe
  df_errors <- df_errors %>% bind_rows(
    df_report_data %>% 
      dplyr::filter(is.na(id)) %>%
      select(reporting_workbook_sheet, section, metric, value) %>%
      mutate(uuid = uuid, .before = 1)
  )
  
  # Select only valid data to upload
  df_report_data <- df_report_data %>% 
    dplyr::filter(!is.na(id)) %>%
    select(month, reporting_workbook_sheet, section, metric, value, id, name, classification_name, classification_value) %>%
    mutate(uuid = uuid, .before = 1)
  
  # Apply fnPostData to each row of the data points dataframe
  df_uploaded_data <- df_uploaded_data %>% bind_rows(df_report_data %>% mutate(status = apply(df_report_data, 1, fnPostData, uuid)))

  # * * 3.8.4. Impact Reporting: Outputs ----
  df_report_data <- df_outputs %>% 
    dplyr::filter(month >= dt_current_month & month < dt_current_month %m+% months(1)) %>%
    mutate(metric = output) %>%
    # Group and summarise on metric
    group_by(month, section, metric) %>%
    summarise(value = n(),
              .groups = 'keep') %>%
    ungroup() %>%
    # Filter out any zero values
    dplyr::filter(value > 0) %>%
    # Add in the reporting_workbook_sheet and section fields
    mutate(reporting_workbook_sheet = 'Lookups_Outputs') %>%
    # Join to the impact reporting lookup data
    left_join(df_lookup, by = c('reporting_workbook_sheet', 'section', 'metric'))
  
  # Report any unmatched data to the error dataframe
  df_errors <- df_errors %>% bind_rows(
    df_report_data %>% 
      dplyr::filter(is.na(id)) %>%
      select(reporting_workbook_sheet, section, metric, value) %>%
      mutate(uuid = uuid, .before = 1)
  )
  
  # Select only valid data to upload
  df_report_data <- df_report_data %>% 
    dplyr::filter(!is.na(id)) %>%
    select(month, reporting_workbook_sheet, section, metric, value, id, name, classification_name, classification_value) %>%
    mutate(uuid = uuid, .before = 1)
  
  # Apply fnPostData to each row of the data points dataframe
  df_uploaded_data <- df_uploaded_data %>% bind_rows(df_report_data %>% mutate(status = apply(df_report_data, 1, fnPostData, uuid)))

  # * * 3.8.5. Impact Reporting: Outcomes ----
  df_report_data <- df_outcomes %>% 
    dplyr::filter(month >= dt_current_month & month < dt_current_month %m+% months(1)) %>%
    mutate(metric = outcome) %>%
    # Group and summarise on metric
    group_by(month, section, metric) %>%
    summarise(value = n(),
              .groups = 'keep') %>%
    ungroup() %>%
    # Filter out any zero values
    dplyr::filter(value > 0) %>%
    # Add in the reporting_workbook_sheet and section fields
    mutate(reporting_workbook_sheet = 'Lookups_Outcomes') %>%
    # Join to the impact reporting lookup data
    left_join(df_lookup, by = c('reporting_workbook_sheet', 'section', 'metric'))
  
  # Report any unmatched data to the error dataframe
  df_errors <- df_errors %>% bind_rows(
    df_report_data %>% 
      dplyr::filter(is.na(id)) %>%
      select(reporting_workbook_sheet, section, metric, value) %>%
      mutate(uuid = uuid, .before = 1)
  )
  
  # Select only valid data to upload
  df_report_data <- df_report_data %>% 
    dplyr::filter(!is.na(id)) %>%
    select(month, reporting_workbook_sheet, section, metric, value, id, name, classification_name, classification_value) %>%
    mutate(uuid = uuid, .before = 1)
  
  # Apply fnPostData to each row of the data points dataframe
  df_uploaded_data <- df_uploaded_data %>% bind_rows(df_report_data %>% mutate(status = apply(df_report_data, 1, fnPostData, uuid)))

  write.csv(df_errors, 'errors.csv')
  write.csv(df_uploaded_data, 'uploads.csv')
}


# 4. Create Report ----
# ═════════════════════

rmarkdown::render(input = 'Flow_Report.Rmd',
                  output_file = dlgSave(title = "Save file as", 
                                        default = "Flow_Report.docx")$res)
