##############################################################################
# Cosmic Quarry Mining Cargo Loss Claims Data
##############################################################################

# ---------------------------------------------------------------------------
# Load packages
# ---------------------------------------------------------------------------
library(readxl)
library(dplyr)
library(stringr)
library(tidyr)
library(purrr)
library(rlang)
library(writexl)

# ---------------------------------------------------------------------------
# Version control
# ---------------------------------------------------------------------------
# v1 | Sarina Truong | 24 Feb 2026 | Initial code
# v2 | Cleaned       | 11 Apr 2026 | Standardised structure, naming, comments

# ---------------------------------------------------------------------------
# File paths
# ---------------------------------------------------------------------------
claims_file_path <- "~/Downloads/university /year 4/semester 1/SOA Case Study/srcsc-2026-claims-cargo.xlsx"
output_file_path <- "~/Library/Mobile Documents/com~apple~CloudDocs/ACTL4001 Assignment/SOA_2026_Case_Study_Materials/Cargo Loss/cargo_loss_CLEAN.xlsx"

##############################################################################
# 1. Import raw data
##############################################################################

cargo_severity_raw <- read_excel(claims_file_path, sheet = "sev")
cargo_frequency_raw <- read_excel(claims_file_path, sheet = "freq")

# Create working copies
cargo_severity <- cargo_severity_raw
cargo_frequency <- cargo_frequency_raw

##############################################################################
# 2. Initial duplicate checks
##############################################################################

# Check duplicate risk records in the frequency data
cargo_frequency %>%
  group_by(policy_id, shipment_id) %>%
  filter(n() > 1) %>%
  ungroup()

# Check duplicate risk records in the severity data
cargo_severity %>%
  count(policy_id, shipment_id) %>%
  filter(n > 1) %>%
  nrow()

##############################################################################
# 3. Remove malformed text suffixes
##############################################################################

# Remove any trailing text beginning with "_???"
remove_malformed_suffix <- function(df) {
  df %>%
    mutate(
      across(
        where(is.character),
        ~ sub("_\\?\\?\\?.*", "", .)
      )
    )
}

cargo_frequency <- remove_malformed_suffix(cargo_frequency)
cargo_severity <- remove_malformed_suffix(cargo_severity)

##############################################################################
# 4. Clean cargo value and weight fields
##############################################################################

# Expected cargo value per kg by cargo type
cargo_value_multiples <- c(
  "gold"        = 135600,
  "cobalt"      = 52,
  "platinum"    = 54500,
  "lithium"     = 82,
  "supplies"    = 10,
  "rare earths" = 85,
  "titanium"    = 7
)

# Find the closest cargo type match based on an implied value multiple
get_closest_multiple_name <- function(implied_multiple, multiples_lookup) {
  if (is.na(implied_multiple) || !is.finite(implied_multiple)) {
    return(NA_character_)
  }
  
  names(multiples_lookup)[which.min(abs(implied_multiple - as.numeric(multiples_lookup)))]
}

# Clean cargo value, weight and cargo type fields
clean_value_weight_fields <- function(df, multiples_lookup, tolerance = 0.01) {
  df %>%
    mutate(
      cargo_value = ifelse(!is.na(cargo_value) & cargo_value < 0, abs(cargo_value), cargo_value),
      weight = ifelse(!is.na(weight) & weight < 0, abs(weight), weight),
      
      cargo_type_clean = tolower(trimws(cargo_type)),
      cargo_type_blank = is.na(cargo_type_clean) | cargo_type_clean == "",
      
      implied_multiple = ifelse(
        cargo_type_blank & !is.na(cargo_value) & !is.na(weight) & weight != 0,
        cargo_value / weight,
        NA_real_
      ),
      
      closest_match_name = vapply(
        implied_multiple,
        get_closest_multiple_name,
        FUN.VALUE = character(1),
        multiples_lookup = multiples_lookup
      ),
      
      closest_match_value = as.numeric(multiples_lookup[closest_match_name]),
      
      inferred_match_flag = ifelse(
        !is.na(implied_multiple) & !is.na(closest_match_value),
        abs(implied_multiple - closest_match_value) / closest_match_value <= tolerance,
        FALSE
      ),
      
      multiple_not_found_flag = cargo_type_blank & !is.na(implied_multiple) & !inferred_match_flag,
      
      cargo_type = ifelse(cargo_type_blank & inferred_match_flag, closest_match_name, cargo_type),
      multiple = as.numeric(multiples_lookup[tolower(trimws(cargo_type))]),
      
      weight = case_when(
        !is.na(cargo_value) & cargo_value %% 10 == 0 & !is.na(multiple) ~ cargo_value / multiple,
        TRUE ~ weight
      ),
      
      cargo_value = case_when(
        !is.na(weight) & weight %% 10 == 0 & !is.na(multiple) ~ weight * multiple,
        TRUE ~ cargo_value
      )
    ) %>%
    dplyr::select(1:17)
}

cargo_frequency <- clean_value_weight_fields(cargo_frequency, cargo_value_multiples)
cargo_severity <- clean_value_weight_fields(cargo_severity, cargo_value_multiples)

##############################################################################
# 5. Recalculate claim counts using severity records
##############################################################################

# Rebuild claim_count using severity-level records
clean_claim_count <- function(frequency_df, severity_df) {
  
  severity_count_both_keys <- severity_df %>%
    count(policy_id, shipment_id, name = "severity_count_both_keys")
  
  severity_count_policy_duration <- severity_df %>%
    filter(!is.na(policy_id)) %>%
    count(policy_id, transit_duration, name = "severity_count_policy_duration")
  
  severity_count_shipment_duration <- severity_df %>%
    filter(!is.na(shipment_id)) %>%
    count(shipment_id, transit_duration, name = "severity_count_shipment_duration")
  
  severity_count_duration <- severity_df %>%
    count(transit_duration, name = "severity_count_duration")
  
  frequency_df %>%
    mutate(claim_count_original = claim_count) %>%
    left_join(severity_count_both_keys, by = c("policy_id", "shipment_id")) %>%
    left_join(severity_count_policy_duration, by = c("policy_id", "transit_duration")) %>%
    left_join(severity_count_shipment_duration, by = c("shipment_id", "transit_duration")) %>%
    left_join(severity_count_duration, by = c("transit_duration")) %>%
    mutate(
      claim_count = case_when(
        !is.na(policy_id) & !is.na(shipment_id) ~ severity_count_both_keys,
        !is.na(policy_id) & is.na(shipment_id) ~ severity_count_policy_duration,
        is.na(policy_id) & !is.na(shipment_id) ~ severity_count_shipment_duration,
        TRUE ~ severity_count_duration
      ),
      claim_count = coalesce(as.integer(claim_count), 0L),
      claim_count_lookup_rule = case_when(
        !is.na(policy_id) & !is.na(shipment_id) ~ "policy_id + shipment_id",
        !is.na(policy_id) & is.na(shipment_id) ~ "policy_id + transit_duration",
        is.na(policy_id) & !is.na(shipment_id) ~ "shipment_id + transit_duration",
        TRUE ~ "transit_duration only"
      )
    ) %>%
    select(
      -severity_count_both_keys,
      -severity_count_policy_duration,
      -severity_count_shipment_duration,
      -severity_count_duration
    )
}

cargo_frequency_count <- clean_claim_count(cargo_frequency, cargo_severity)

# Check that all claim counts fall within the expected range
cargo_frequency_count %>%
  filter(!claim_count %in% 0:5)

##############################################################################
# 6. Define validation rules
##############################################################################

allowed_values <- list(
  cargo_type = c("cobalt", "gold", "platinum", "supplies", "titanium", "lithium", "rare earths"),
  container_type = c(
    "LongHaul Vault Canister",
    "QuantumCrate Module",
    "HardSeal Transit Crate",
    "DockArc Freight Case",
    "DeepSpace Haulbox"
  ),
  route_risk = c(0, 1, 2, 3, 4, 5)
)

value_bounds <- list(
  cargo_value = c(50000, 680000000),
  weight = c(1500, 250000),
  distance = c(1, 100),
  transit_duration = c(1, 60),
  pilot_experience = c(1, 30),
  vessel_age = c(1, 50),
  solar_radiation = c(0, 1),
  debris_density = c(0, 1),
  exposure = c(0, 1)
)

##############################################################################
# 7. Validation helpers
##############################################################################

# Check whether values satisfy allowed categories or numeric bounds
is_valid_vector <- function(x, column_name, allowed_lookup, bounds_lookup) {
  
  if (column_name %in% names(allowed_lookup)) {
    return(!is.na(x) & x %in% allowed_lookup[[column_name]])
  }
  
  if (column_name %in% names(bounds_lookup)) {
    lower_bound <- bounds_lookup[[column_name]][1]
    upper_bound <- bounds_lookup[[column_name]][2]
    x_numeric <- suppressWarnings(as.numeric(x))
    return(!is.na(x_numeric) & x_numeric >= lower_bound & x_numeric <= upper_bound)
  }
  
  !is.na(x)
}

# Collapse dataset B to one row per risk key to avoid join duplication
collapse_to_risk_level <- function(df, columns_to_check, key_columns = c("policy_id", "shipment_id")) {
  df %>%
    mutate(.key = paste(.data[[key_columns[1]]], .data[[key_columns[2]]], sep = "||")) %>%
    group_by(.key) %>%
    summarise(
      across(
        all_of(columns_to_check),
        ~ {
          unique_values <- unique(na.omit(.))
          if (length(unique_values) == 1) unique_values else NA
        }
      ),
      conflict_flag = {
        conflicts <- map_lgl(across(all_of(columns_to_check)), ~ length(unique(na.omit(.x))) > 1)
        any(conflicts)
      },
      .groups = "drop"
    )
}

# Fill invalid values in dataset A using valid values from dataset B
fill_invalid_from_other <- function(df_a,
                                    df_b,
                                    allowed_lookup,
                                    bounds_lookup,
                                    key_columns = c("policy_id", "shipment_id"),
                                    summary_label = "A_filled_from_B") {
  
  common_columns <- intersect(names(df_a), names(df_b))
  fill_columns <- setdiff(common_columns, key_columns)
  
  df_a_keyed <- df_a %>%
    mutate(.key = paste(policy_id, shipment_id, sep = "||"))
  
  df_b_collapsed <- collapse_to_risk_level(df_b, fill_columns, key_columns)
  
  if (any(df_b_collapsed$conflict_flag, na.rm = TRUE)) {
    warning("Conflicting values found within some (policy_id, shipment_id) groups in df_b.")
  }
  
  joined_df <- df_a_keyed %>%
    left_join(df_b_collapsed, by = ".key", suffix = c("", ".b"))
  
  invalid_before <- map_dfc(fill_columns, function(col) {
    !is_valid_vector(joined_df[[col]], col, allowed_lookup, bounds_lookup)
  })
  
  joined_df$out_of_bounds_before <- rowSums(invalid_before) > 0
  
  for (col in fill_columns) {
    a_valid <- is_valid_vector(joined_df[[col]], col, allowed_lookup, bounds_lookup)
    b_valid <- is_valid_vector(joined_df[[paste0(col, ".b")]], col, allowed_lookup, bounds_lookup)
    
    fill_index <- (!a_valid) & b_valid
    joined_df[[col]][fill_index] <- joined_df[[paste0(col, ".b")]][fill_index]
  }
  
  invalid_after <- map_dfc(fill_columns, function(col) {
    !is_valid_vector(joined_df[[col]], col, allowed_lookup, bounds_lookup)
  })
  
  joined_df$out_of_bounds_after <- rowSums(invalid_after) > 0
  
  cleaned_df <- joined_df %>%
    select(-ends_with(".b"), -.key)
  
  summary_table <- tibble(
    direction = summary_label,
    total_rows = nrow(cleaned_df),
    rows_out_of_bounds_before = sum(cleaned_df$out_of_bounds_before),
    rows_out_of_bounds_after = sum(cleaned_df$out_of_bounds_after),
    rows_fixed = sum(cleaned_df$out_of_bounds_before & !cleaned_df$out_of_bounds_after)
  )
  
  list(
    data = cleaned_df,
    summary = summary_table
  )
}

##############################################################################
# 8. Cross-fill invalid values between frequency and severity
##############################################################################

# Fill frequency data using severity data
cargo_frequency_fixed <- fill_invalid_from_other(
  df_a = cargo_frequency_count,
  df_b = cargo_severity,
  allowed_lookup = allowed_values,
  bounds_lookup = value_bounds,
  summary_label = "frequency_filled_from_severity"
)

cargo_frequency_filled <- cargo_frequency_fixed$data
summary_frequency_from_severity <- cargo_frequency_fixed$summary

# Fill severity data using frequency data
cargo_severity_fixed <- fill_invalid_from_other(
  df_a = cargo_severity,
  df_b = cargo_frequency_count,
  allowed_lookup = allowed_values,
  bounds_lookup = value_bounds,
  summary_label = "severity_filled_from_frequency"
)

cargo_severity_filled <- cargo_severity_fixed$data
summary_severity_from_frequency <- cargo_severity_fixed$summary

# Combined summary of cross-filling
bind_rows(summary_frequency_from_severity, summary_severity_from_frequency)

##############################################################################
# 9. Out-of-bounds summaries
##############################################################################

out_of_bounds_by_column <- function(df, allowed_lookup, bounds_lookup) {
  columns_to_check <- intersect(names(df), c(names(allowed_lookup), names(bounds_lookup)))
  
  map_dfr(columns_to_check, function(col) {
    invalid_flag <- !is_valid_vector(df[[col]], col, allowed_lookup, bounds_lookup)
    
    tibble(
      column = col,
      n_out_of_bounds = sum(invalid_flag, na.rm = TRUE),
      pct_out_of_bounds = 100 * mean(invalid_flag, na.rm = TRUE)
    )
  }) %>%
    arrange(desc(n_out_of_bounds))
}

oob_frequency <- out_of_bounds_by_column(cargo_frequency_filled, allowed_values, value_bounds)
oob_severity <- out_of_bounds_by_column(cargo_severity_filled, allowed_values, value_bounds)

oob_frequency
oob_severity

##############################################################################
# 10. Null, negative and logical checks
##############################################################################

# Count null values
colSums(is.na(cargo_frequency_filled))
colSums(is.na(cargo_severity_filled))

# Count total NA cells in a dataset
count_total_na <- function(df) {
  df %>%
    summarise(across(everything(), ~ sum(is.na(.)))) %>%
    summarise(total_na_cells = sum(across(everything()))) %>%
    pull(total_na_cells)
}

# Count negative numeric values
cargo_frequency_filled %>%
  summarise(across(where(is.numeric), ~ sum(. < 0, na.rm = TRUE)))

cargo_severity_filled %>%
  summarise(across(where(is.numeric), ~ sum(. < 0, na.rm = TRUE)))

# Force claim amount positive and flag impossible claims
cargo_severity_filled <- cargo_severity_filled %>%
  mutate(
    claim_amount = abs(claim_amount),
    claim_exceeds_cargo_flag = case_when(
      is.na(claim_amount) | is.na(cargo_value) ~ NA,
      claim_amount > cargo_value ~ TRUE,
      TRUE ~ FALSE
    )
  )

##############################################################################
# 11. Final clean datasets
##############################################################################

# Remove out-of-bounds records from frequency data
cargo_frequency_clean <- cargo_frequency_filled %>%
  filter(!out_of_bounds_after) %>%
  dplyr::select(1:17)

# Remove out-of-bounds and impossible claim records from severity data
cargo_severity_clean <- cargo_severity_filled %>%
  filter(!out_of_bounds_after) %>%
  filter(claim_exceeds_cargo_flag == FALSE) %>%
  dplyr::select(1:17)

##############################################################################
# 12. Export clean datasets
##############################################################################

write_xlsx(
  list(
    Frequency = cargo_frequency_clean,
    Severity = cargo_severity_clean
  ),
  output_file_path
)