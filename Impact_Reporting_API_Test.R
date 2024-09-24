# 0. Load libraries and define functions ----
# ═══════════════════════════════════════════

# * 0.1. Load libraries ----
# ──────────────────────────

# install.packages(c('httr', 'jsonlite'))
library(tidyverse)
library(httr)
library(jsonlite)
library(readxl)
library(lubridate)
library(uuid)
library(conflicted)

# Reporting month
# ═══════════════
dt_start <- as.Date('2024-05-01')
dt_end <- dt_start + months(1)

# Project ID
# ══════════
#project_id <- 2472  # One North Devon
#project_id <- 2473	# High Flow
#project_id <- 2474	# Secondary Care
#project_id <- 2475	# Primary Care
project_id <- 2476	# Community Flow


# * 0.2. Define functions ----
# ────────────────────────────

fnPostData <- function(x, uuid){
  body_text <- jsonlite::toJSON(
    list(
      'apikey' = 'd866628d-f9f3-4668-9849-cbba687774d9',
      'logs' = list(
        list(
          'ownerId' = 36517,
          'ownerType' = 'user',
          'activityId' = as.integer(x['id']),
          'projectId' = project_id,
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

# Set up main URL
url_text <- 'https://app.impactreporting.co.uk/api/v1/logs'

# 1. Load lookups and data ----
# ═════════════════════════════

df_lookup <- read_excel(path = 'impact-reporting-api-entity-references.xlsx', 
                        sheet = 'lookup')

filename <- file.choose()
load(filename)

# 2. Process and upload data ----
# ═══════════════════════════════

# Generate UUID
uuid <- uuid::UUIDgenerate()

# 2.1. Data points ----
# ─────────────────────

df_report_data <- df_data_points %>% 
  dplyr::filter(month >= dt_start & month < dt_end) %>%
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

# 2.2. Support and referrals ----
# ───────────────────────────────

df_report_data <- df_support_and_referrals %>% 
  dplyr::filter(month >= dt_start & month < dt_end) %>%
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

# 2.3. Outputs ----
# ─────────────────

df_report_data <- df_outputs %>% 
  dplyr::filter(month >= dt_start & month < dt_end) %>%
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

# 2.4. Outcomes ----
# ──────────────────

df_report_data <- df_outcomes %>% 
  dplyr::filter(month >= dt_start & month < dt_end) %>%
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

