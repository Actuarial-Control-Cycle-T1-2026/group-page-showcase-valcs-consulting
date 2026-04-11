# Set WD - change to your own iCloud path
setwd("~/Library/Mobile Documents/com~apple~CloudDocs/ACTL4001 Assignment/SOA_2026_Case_Study_Materials/0. Original Case Material")

###############################################################################
# Version Control
# v1  | Vihaan Jain    | 20 Feb 2026 | Initial build
# v2  | Vihaan Jain    | 23 Feb 2026 | Added OOB bounds fixing + neg fix
# v3  | Vihaan Jain    | 27 Feb 2026 | Cleaned up join logic, fixed claim_count
# v4  | Vihaan Jain    | 28 Feb 2026 | Minor tweaks to allowed list, added EDA filter
# v5  | Vihaan Jain    | 01 Mar 2026 | Tidied environment cleanup at end
###############################################################################

library(readxl)
library(stringr)
library(tidyr)
library(purrr)
library(rlang)
library(dplyr)

WC_folder_path <- "~/Library/Mobile Documents/com~apple~CloudDocs/ACTL4001 Assignment/SOA_2026_Case_Study_Materials/Workers Compensation"

wc_freq <- read_excel("srcsc-2026-claims-workers-comp.xlsx", sheet = "freq")
wc_sev  <- read_excel("srcsc-2026-claims-workers-comp.xlsx", sheet = "sev")


# ---- Strip trailing "???_" garbage strings ----

wc_freq <- wc_freq %>%
  mutate(across(where(is.character), ~ sub("_\\?\\?\\?.*", "", .)))

wc_sev <- wc_sev %>%
  mutate(across(where(is.character), ~ sub("_\\?\\?\\?.*", "", .)))


# ---- Fill in missing worker IDs ----
# Some rows in freq had blank IDs - generate placeholder ones

wc_freq <- wc_freq %>%
  mutate(
    worker_id = ifelse(
      is.na(worker_id) | worker_id == "",
      paste0("W-", str_pad(row_number(), width = 5, pad = "0")),
      worker_id
    )
  )

# use common cols (not worker_id) as a lookup key to fill missing IDs in sev
common_cols <- intersect(names(wc_freq), names(wc_sev)) %>% setdiff("worker_id")

freq_lookup <- wc_freq %>%
  filter(!is.na(worker_id)) %>%
  distinct(across(all_of(common_cols)), worker_id) %>%
  group_by(across(all_of(common_cols))) %>%
  filter(n() == 1) %>%
  ungroup() %>%
  dplyr::select(all_of(common_cols), worker_id)

wc_sev <- wc_sev %>%
  left_join(freq_lookup, by = common_cols, suffix = c("", "_from_freq")) %>%
  mutate(worker_id = dplyr::coalesce(worker_id, worker_id_from_freq)) %>%
  dplyr::select(-worker_id_from_freq)


# ---- Cross-fill missing cells using worker_id as key ----
# build one-row-per-worker lookups from each table

common_cols <- intersect(names(wc_freq), names(wc_sev)) %>% setdiff("worker_id")

freq_lookup <- wc_freq %>%
  group_by(worker_id) %>%
  slice_head(n = 1) %>%
  ungroup() %>%
  dplyr::select(worker_id, all_of(common_cols)) %>%
  rename_with(~ paste0(.x, "_from_freq"), all_of(common_cols))

sev_lookup <- wc_sev %>%
  group_by(worker_id) %>%
  slice_head(n = 1) %>%
  ungroup() %>%
  dplyr::select(worker_id, all_of(common_cols)) %>%
  rename_with(~ paste0(.x, "_from_sev"), all_of(common_cols))

# fill freq using sev
wc_freq_clean <- wc_freq %>% left_join(sev_lookup, by = "worker_id")
for (col in common_cols) {
  wc_freq_clean[[col]] <- dplyr::coalesce(wc_freq_clean[[col]],
                                          wc_freq_clean[[paste0(col, "_from_sev")]])
}
wc_freq_clean <- wc_freq_clean %>% dplyr::select(-any_of(paste0(common_cols, "_from_sev")))

# fill sev using freq
wc_sev_clean <- wc_sev %>% left_join(freq_lookup, by = "worker_id")
for (col in common_cols) {
  wc_sev_clean[[col]] <- dplyr::coalesce(wc_sev_clean[[col]],
                                         wc_sev_clean[[paste0(col, "_from_freq")]])
}
wc_sev_clean <- wc_sev_clean %>% dplyr::select(-any_of(paste0(common_cols, "_from_freq")))


# ---- Fix claim_count in freq: count actual rows in sev ----

sev_counts <- wc_sev_clean %>%
  filter(!is.na(worker_id)) %>%
  count(worker_id, name = "claim_count_from_sev")

wc_freq_clean <- wc_freq_clean %>%
  left_join(sev_counts, by = "worker_id") %>%
  mutate(
    claim_count = dplyr::coalesce(claim_count_from_sev, 0L),
    claim_count = as.integer(claim_count)
  ) %>%
  dplyr::select(-claim_count_from_sev)


# ---- OOB bounds definitions ----
# tweak these if the data spec changes

allowed <- list(
  accident_history_flag   = c(0, 1),
  psych_stress_index      = c(1, 2, 3, 4, 5),
  safety_training_index   = c(1, 2, 3, 4, 5),
  protective_gear_quality = c(1, 2, 3, 4, 5),
  hours_per_week          = c(20, 25, 30, 35, 40)
)

bounds <- list(
  experience_yrs    = c(0.2, 40),
  supervision_level = c(0, 1),
  gravity_level     = c(0.75, 1.50),
  base_salary       = c(20000, 130000),
  exposure          = c(0, 1),
  claim_count       = c(0, 2),       # freq only
  claim_length      = c(3, 1000),    # sev only
  claim_amount      = c(0, 2000000)  # sev only
)

# returns TRUE for valid (or NA) entries
is_valid_vec <- function(x, colname, allowed, bounds) {
  if (!colname %in% c(names(allowed), names(bounds))) return(rep(TRUE, length(x)))
  
  if (colname %in% names(allowed)) {
    return(is.na(x) | x %in% allowed[[colname]])
  }
  
  lo <- bounds[[colname]][1]
  hi <- bounds[[colname]][2]
  xn <- suppressWarnings(as.numeric(x))
  is.na(xn) | (xn >= lo & xn <= hi)
}

flag_oob_rows <- function(df, allowed, bounds) {
  cols <- intersect(names(df), c(names(allowed), names(bounds)))
  if (length(cols) == 0) { df$any_oob <- FALSE; return(df) }
  
  invalid_mat <- map_dfc(cols, ~ !is_valid_vec(df[[.x]], .x, allowed, bounds))
  df$any_oob <- rowSums(invalid_mat, na.rm = TRUE) > 0
  df
}

# collapse "other" table to 1 row per worker - avoids join explosion
collapse_one_row_per_worker <- function(df, cols, key = "worker_id") {
  df %>%
    group_by(.data[[key]]) %>%
    summarise(
      across(all_of(cols), ~ {
        v <- unique(na.omit(.x))
        if (length(v) == 1) v else NA
      }),
      .groups = "drop"
    )
}

# borrow in-bounds values from the other table where current values are OOB
fix_oob_from_other <- function(df_a, df_b, allowed, bounds, key = "worker_id") {
  common_cols <- intersect(names(df_a), names(df_b))
  cols_ruled  <- intersect(common_cols, c(names(allowed), names(bounds))) %>% setdiff(key)
  
  if (length(cols_ruled) == 0) return(df_a)
  
  b1 <- collapse_one_row_per_worker(df_b, cols_ruled, key = key)
  
  joined <- df_a %>% left_join(b1, by = key, suffix = c("", ".b"))
  
  for (col in cols_ruled) {
    a_valid <- is_valid_vec(joined[[col]], col, allowed, bounds)
    b_valid <- is_valid_vec(joined[[paste0(col, ".b")]], col, allowed, bounds)
    fill_idx <- (!a_valid) & b_valid
    joined[[col]][fill_idx] <- joined[[paste0(col, ".b")]][fill_idx]
  }
  
  joined %>% dplyr::select(-ends_with(".b"))
}

# before counts
freq_before <- flag_oob_rows(wc_freq_clean, allowed, bounds)
sev_before  <- flag_oob_rows(wc_sev_clean,  allowed, bounds)

tibble(
  dataset           = c("wc_freq_clean", "wc_sev_clean"),
  rows_with_any_oob = c(sum(freq_before$any_oob, na.rm = TRUE),
                        sum(sev_before$any_oob,  na.rm = TRUE))
)

# fix
wc_freq_clean <- fix_oob_from_other(wc_freq_clean, wc_sev_clean, allowed, bounds)
wc_sev_clean  <- fix_oob_from_other(wc_sev_clean,  wc_freq_clean, allowed, bounds)

# after counts
freq_after <- flag_oob_rows(wc_freq_clean, allowed, bounds)
sev_after  <- flag_oob_rows(wc_sev_clean,  allowed, bounds)

tibble(
  dataset           = c("wc_freq_clean", "wc_sev_clean"),
  rows_with_any_oob = c(sum(freq_after$any_oob, na.rm = TRUE),
                        sum(sev_after$any_oob,  na.rm = TRUE))
)

# set remaining OOB values to NA
set_oob_to_na <- function(df, allowed, bounds) {
  cols <- intersect(names(df), c(names(allowed), names(bounds)))
  for (col in cols) {
    valid <- is_valid_vec(df[[col]], col, allowed, bounds)
    df[[col]][!valid] <- NA
  }
  df
}

wc_freq_oobNA <- set_oob_to_na(wc_freq_clean, allowed, bounds)
wc_sev_oobNA  <- set_oob_to_na(wc_sev_clean,  allowed, bounds)


# ---- Flip negatives ----

make_negatives_positive <- function(df) {
  df %>% mutate(across(where(is.numeric), ~ ifelse(!is.na(.) & . < 0, abs(.), .)))
}

wc_freq_oobNA <- make_negatives_positive(wc_freq_oobNA)
wc_sev_oobNA  <- make_negatives_positive(wc_sev_oobNA)

# sanity check - should both be 0
count_neg_cells <- function(df) {
  df %>%
    summarise(across(where(is.numeric), ~ sum(. < 0, na.rm = TRUE))) %>%
    summarise(total = sum(across(everything()))) %>%
    pull(total)
}

count_neg_cells(wc_freq_oobNA)
count_neg_cells(wc_sev_oobNA)

# NA summary across pipeline stages
na_total <- function(df) {
  df %>%
    summarise(across(everything(), ~ sum(is.na(.)))) %>%
    summarise(total = sum(across(everything()))) %>%
    pull(total)
}

tibble(
  dataset        = c("wc_freq", "wc_sev", "wc_freq_clean", "wc_sev_clean"),
  total_na_cells = c(na_total(wc_freq), na_total(wc_sev),
                     na_total(wc_freq_clean), na_total(wc_sev_clean))
)


# ---- Rename and build EDA datasets ----

wc_freq_NA <- wc_freq_oobNA
wc_sev_NA  <- wc_sev_oobNA

# drop any rows with remaining NAs for modelling
wc_freq_eda <- wc_freq_NA %>% filter(complete.cases(.))
wc_sev_eda  <- wc_sev_NA  %>% filter(complete.cases(.))

# tidy up - keep only what we need going forward
rm(list = setdiff(ls(), c("wc_freq", "wc_sev",
                          "wc_freq_clean", "wc_sev_clean",
                          "wc_freq_NA", "wc_sev_NA",
                          "wc_freq_eda", "wc_sev_eda")))
