##############################################################################
# ROBUST: Cosmic Quarry Mining Equipment Failure Claims Data 
# (Cleaned: Dropped <1% errors, Kept equipment_age)
##############################################################################
library(readxl)
library(dplyr)
library(stringr)
library(tidyr)
library(purrr)
library(rlang)
library(writexl)

# Define file path
file_path <- "Desktop/SOA_2026_Case_Study_Materials/srcsc-2026-claims-equipment-failure.xlsx"

# Import datasets
equipment_freq <- read_excel(file_path, sheet = "freq")
equipment_sev <- read_excel(file_path, sheet = "sev")

##############################################################################
# 1. Basic Cleaning & Formatting
##############################################################################

# Remove "_???" strings
equipment_freq <- equipment_freq %>% mutate(across(where(is.character), ~ sub("_\\?\\?\\?.*", "", .)))
equipment_sev <- equipment_sev %>% mutate(across(where(is.character), ~ sub("_\\?\\?\\?.*", "", .)))

# Fix known typos and force numerics/absolutes
clean_basics <- function(df) {
  df %>%
    mutate(
      equipment_type = case_when(
        equipment_type == "ReglAggregators" ~ "Mag-Lift Aggregator",
        equipment_type == "FexStram Carrier" ~ "FluxStream Carrier",
        equipment_type == "Flux Rider" ~ "Fusion Transport",
        TRUE ~ equipment_type
      ),
      equipment_type = trimws(equipment_type),
      solar_system = trimws(solar_system),
      
      # Force numeric and fix accidental negatives
      equipment_age = abs(as.numeric(equipment_age)),
      maintenance_int = abs(as.numeric(maintenance_int)),
      usage_int = abs(as.numeric(usage_int)),
      exposure = abs(as.numeric(exposure))
    )
}

equipment_freq <- clean_basics(equipment_freq)
equipment_sev <- clean_basics(equipment_sev)

# Fix claim_amount in Severity specifically
if("claim_amount" %in% names(equipment_sev)) {
  equipment_sev <- equipment_sev %>% mutate(claim_amount = abs(as.numeric(claim_amount)))
}

##############################################################################
# 2. Hierarchical Missing ID Recovery
##############################################################################
# If IDs are missing, try to group by exact machine specifications to fill them
recover_ids <- function(df) {
  df %>%
    group_by(equipment_type, equipment_age, solar_system, maintenance_int, usage_int) %>%
    fill(policy_id, equipment_id, .direction = "downup") %>%
    ungroup()
}

equipment_freq <- recover_ids(equipment_freq)
equipment_sev <- recover_ids(equipment_sev)

##############################################################################
# 3. Robust Claim Count Calculation (Hierarchical Fallbacks)
##############################################################################

# Precompute claim counts in Sev for possible lookup keys
sev_counts_both <- equipment_sev %>% count(policy_id, equipment_id, name = "sev_n_both")
sev_counts_policy <- equipment_sev %>% filter(!is.na(policy_id)) %>% count(policy_id, name = "sev_n_policy")
sev_counts_equip <- equipment_sev %>% filter(!is.na(equipment_id)) %>% count(equipment_id, name = "sev_n_equip")

equipment_freq <- equipment_freq %>%
  mutate(claim_count_original = as.numeric(claim_count)) %>%
  left_join(sev_counts_both, by = c("policy_id", "equipment_id")) %>%
  left_join(sev_counts_policy, by = "policy_id") %>%
  left_join(sev_counts_equip, by = "equipment_id") %>%
  mutate(
    # Hierarchical assignment logic
    claim_count_calculated = case_when(
      !is.na(policy_id) & !is.na(equipment_id) ~ sev_n_both,
      !is.na(policy_id) & is.na(equipment_id) ~ sev_n_policy,
      is.na(policy_id) & !is.na(equipment_id) ~ sev_n_equip,
      TRUE ~ 0L
    ),
    claim_count = coalesce(as.integer(claim_count_calculated), 0L),
    
    # Track how we found the count for auditing
    claim_count_lookup_rule = case_when(
      !is.na(policy_id) & !is.na(equipment_id) ~ "policy_id + equipment_id",
      !is.na(policy_id) & is.na(equipment_id) ~ "policy_id only",
      is.na(policy_id) & !is.na(equipment_id) ~ "equipment_id only",
      TRUE ~ "Default to 0"
    )
  ) %>%
  select(-sev_n_both, -sev_n_policy, -sev_n_equip, -claim_count_calculated)

##############################################################################
# 4. Advanced Cross-Healing Data Imputation
##############################################################################

allowed <- list(
  equipment_type = c("Quantum Bore", "Graviton Extractor", "FluxStream Carrier", 
                     "Mag-Lift Aggregator", "Fusion Transport", "Ion Pulverizer"),
  solar_system   = c("Helionis Cluster", "Epsilon", "Zeta")
)

bounds <- list(
  equipment_age   = c(0, 10),
  maintenance_int = c(100, 5000),
  usage_int       = c(0, 24),
  exposure        = c(0, 1),
  claim_amount    = c(11000, 790000)
)

is_valid_vec <- function(x, colname, allowed, bounds) {
  if (colname %in% names(allowed)) return(!is.na(x) & x %in% allowed[[colname]])
  if (colname %in% names(bounds)) {
    xn <- suppressWarnings(as.numeric(x))
    return(!is.na(xn) & xn >= bounds[[colname]][1] & xn <= bounds[[colname]][2])
  }
  !is.na(x)
}

collapse_to_risk_level <- function(df_b, cols, key_cols = c("policy_id", "equipment_id")) {
  df_b %>%
    mutate(.key = paste(.data[[key_cols[1]]], .data[[key_cols[2]]], sep = "||")) %>%
    group_by(.key) %>%
    summarise(
      across(all_of(cols), ~ { v <- unique(na.omit(.)); if (length(v) == 1) v else NA }),
      conflict_flag = any(map_lgl(across(all_of(cols)), ~ length(unique(na.omit(.x))) > 1)),
      .groups = "drop"
    )
}

fill_bad_from_other <- function(df_a, df_b, allowed, bounds, key_cols = c("policy_id", "equipment_id")) {
  common_cols <- setdiff(intersect(names(df_a), names(df_b)), c(key_cols, "claim_count", "data_quality_desc", "data_quality_flag"))
  
  a <- df_a %>% mutate(.key = paste(policy_id, equipment_id, sep = "||"))
  b <- collapse_to_risk_level(df_b, common_cols, key_cols)
  
  joined <- a %>% left_join(b, by = ".key", suffix = c("", ".b"))
  
  # Impute missing/invalid values
  for (col in common_cols) {
    a_valid <- is_valid_vec(joined[[col]], col, allowed, bounds)
    b_valid <- is_valid_vec(joined[[paste0(col, ".b")]], col, allowed, bounds)
    fill_idx <- (!a_valid) & b_valid
    joined[[col]][fill_idx] <- joined[[paste0(col, ".b")]][fill_idx]
  }
  
  joined %>% select(-ends_with(".b"), -.key)
}

# Execute bidirectional cross-healing
equipment_freq_filled <- fill_bad_from_other(equipment_freq, equipment_sev, allowed, bounds)
equipment_sev_filled <- fill_bad_from_other(equipment_sev, equipment_freq_filled, allowed, bounds)

# VERY IMPORTANT: Replace NAs generated by the join so we don't accidentally drop zero-claim policies!
equipment_freq_filled <- equipment_freq_filled %>% mutate(conflict_flag = replace_na(conflict_flag, FALSE))
equipment_sev_filled <- equipment_sev_filled %>% mutate(conflict_flag = replace_na(conflict_flag, FALSE))

##############################################################################
# 5. Impact Assessment & Column Breakdown
##############################################################################
print("=========================================================")
print("IMPACT ASSESSMENT: DATA QUALITY FLAGS")
print("=========================================================")

out_of_bounds_by_col <- function(df, allowed, bounds) {
  cols_to_check <- intersect(names(df), c(names(allowed), names(bounds)))
  
  map_dfr(cols_to_check, function(col) {
    invalid <- !is_valid_vec(df[[col]], col, allowed, bounds)
    tibble(
      column = col,
      n_out_of_bounds = sum(invalid, na.rm = TRUE),
      pct_out_of_bounds = round(100 * mean(invalid, na.rm = TRUE), 2)
    )
  }) %>% arrange(desc(n_out_of_bounds))
}

assess_impact <- function(df, dataset_name, allowed, bounds) {
  total_rows <- nrow(df)
  cat(sprintf("\n--- %s Dataset ---\n", toupper(dataset_name)))
  cat(sprintf("Total Rows: %d\n", total_rows))
  cat("Breakdown of out-of-bounds errors by column:\n")
  print(out_of_bounds_by_col(df, allowed, bounds))
}

assess_impact(equipment_freq_filled, "Frequency", allowed, bounds)
assess_impact(equipment_sev_filled, "Severity", allowed, bounds)

##############################################################################
# 6. Surgical Clean & Final Export 
##############################################################################

# Custom function to drop rows failing bounds EXCEPT for equipment_age
filter_except_age <- function(df, allowed, bounds) {
  # Get all dictionary columns present in the dataset, omitting equipment_age
  cols_to_check <- setdiff(intersect(names(df), c(names(allowed), names(bounds))), "equipment_age")
  
  # Create a matrix of TRUE/FALSE for invalid data in these specific columns
  invalid_matrix <- map_dfc(cols_to_check, function(col) {
    !is_valid_vec(df[[col]], col, allowed, bounds)
  })
  
  # Flag rows for removal if ANY of these other columns are invalid
  df$remove_flag <- rowSums(invalid_matrix) > 0
  
  # Filter out flagged rows and any unresolved conflicts, safely handling NAs
  df_clean <- df %>% 
    filter(remove_flag == FALSE & conflict_flag == FALSE) %>%
    select(
      -remove_flag,
      -conflict_flag,
      -any_of(c("data_quality_desc", "data_quality_flag"))
    )
  
  return(df_clean)
}

# Apply the surgical filter
equipment_freq_clean <- filter_except_age(equipment_freq_filled, allowed, bounds)
equipment_sev_clean <- filter_except_age(equipment_sev_filled, allowed, bounds)

# Calculate exactly how much we actually dropped
freq_drop_pct <- round((1 - nrow(equipment_freq_clean) / nrow(equipment_freq_filled)) * 100, 2)
sev_drop_pct <- round((1 - nrow(equipment_sev_clean) / nrow(equipment_sev_filled)) * 100, 2)

cat(sprintf("\n=========================================================\n"))
cat(sprintf("FINAL EXPORT SUMMARY\n"))
cat(sprintf("=========================================================\n"))
cat(sprintf("Frequency Data Dropped: %s%% (Targeted cleanup successful)\n", freq_drop_pct))
cat(sprintf("Severity Data Dropped:  %s%% (Targeted cleanup successful)\n", sev_drop_pct))
cat(sprintf("Note: equipment_age anomalies were deliberately preserved.\n"))

write_xlsx(
  list(
    Frequency_Clean = equipment_freq_clean,
    Severity_Clean  = equipment_sev_clean
  ),
  "Desktop/SOA_2026_Case_Study_Materials/equipment_failure_CLEANED_FINAL.xlsx"
)

cat(sprintf("\nData exported successfully to equipment_failure_CLEANED_FINAL.xlsx\n"))