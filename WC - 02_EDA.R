###############################################################################
# Version Control
# v1  | Vihaan Jain  | 24 Feb 2026 | Initial EDA script - distributions + cat plots
# v2  | Vihaan Jain  | 27 Feb 2026 | Added log scale histogram, flipped coord on bar charts
# v3  | Vihaan Jain  | 01 Mar 2026 | Binned numeric plots for both freq + sev
# v4  | Vihaan Jain  | 02 Mar 2026 | Added injury_type and injury_cause for severity
###############################################################################

library(readxl)
library(dplyr)
library(stringr)
library(tidyr)
library(purrr)
library(ggplot2)
library(scales)

options(scipen = 999)


# Raw Distributions

# claim count is discrete so factor it
ggplot(wc_freq_eda, aes(x = factor(claim_count))) +
  geom_bar(fill = "grey30") +
  labs(title = "Distribution of Claim Count",
       x = "Number of Claims", y = "Frequency") +
  theme_minimal()

# claim amount - raw first, then log scale to see spread better
ggplot(wc_sev_eda, aes(x = claim_amount)) +
  geom_histogram(bins = 50, fill = "grey30", color = "white") +
  labs(title = "Distribution of Claim Amount",
       x = "Claim Amount", y = "Frequency") +
  theme_minimal()

ggplot(subset(wc_sev_eda, claim_amount > 0), aes(x = claim_amount)) +
  geom_histogram(bins = 50, fill = "grey30", color = "white") +
  scale_x_log10() +
  labs(title = "Distribution of Claim Amount (Log Scale)",
       x = "Claim Amount (log scale)", y = "Frequency") +
  theme_minimal()


# ---- Frequency: categorical variables vs claim rate ----

plot_count_by_cat <- function(df, var) {
  df %>%
    group_by(.data[[var]]) %>%
    summarise(
      n_rows          = n(),
      mean_claim_rate = mean(claim_count / exposure, na.rm = TRUE),
      prop_claims     = mean(claim_count > 0, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    ggplot(aes(x = reorder(.data[[var]], mean_claim_rate), y = mean_claim_rate)) +
    geom_col(fill = "grey30", width = 0.7) +
    geom_text(aes(label = sprintf("%.3f", mean_claim_rate)),
              hjust = -0.1, size = 3) +
    coord_flip() +
    labs(title = paste0("Mean claim rate by ", var),
         x = var, y = "Mean claim count per exposure-year") +
    theme_minimal()
}

# binned version for continuous predictors
plot_mean_count_by_num_bins <- function(df, var, bins = 10) {
  df %>%
    filter(!is.na(.data[[var]]), !is.na(claim_count)) %>%
    mutate(bin = cut(.data[[var]], breaks = bins, include.lowest = TRUE)) %>%
    group_by(bin) %>%
    summarise(n = n(), mean_claim_count = mean(claim_count, na.rm = TRUE), .groups = "drop") %>%
    ggplot(aes(x = bin, y = mean_claim_count)) +
    geom_col(fill = "grey30") +
    labs(title = paste0("Mean claim count by binned ", var),
         x = paste0(var, " (binned)"), y = "Mean claim count") +
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 45, hjust = 1))
}

plot_count_by_cat(wc_freq_eda, "solar_system")
plot_count_by_cat(wc_freq_eda, "station_id")
plot_count_by_cat(wc_freq_eda, "occupation")
plot_count_by_cat(wc_freq_eda, "employment_type")
plot_count_by_cat(wc_freq_eda, "accident_history_flag")
plot_count_by_cat(wc_freq_eda, "psych_stress_index")
plot_count_by_cat(wc_freq_eda, "hours_per_week")
plot_count_by_cat(wc_freq_eda, "safety_training_index")
plot_count_by_cat(wc_freq_eda, "protective_gear_quality")

plot_mean_count_by_num_bins(wc_freq_eda, "experience_yrs",   bins = 12)
plot_mean_count_by_num_bins(wc_freq_eda, "supervision_level", bins = 10)
plot_mean_count_by_num_bins(wc_freq_eda, "gravity_level",     bins = 10)
plot_mean_count_by_num_bins(wc_freq_eda, "base_salary",       bins = 12)


# ---- Severity: categorical variables vs claim amount ----

plot_amount_by_cat <- function(df, var) {
  ggplot(subset(df, !is.na(claim_amount) & claim_amount > 0),
         aes(x = .data[[var]], y = claim_amount)) +
    geom_boxplot(outlier.alpha = 0.2, fill = "grey80") +
    scale_y_log10(labels = label_number()) +
    labs(title = paste0("Claim amount by ", var, " (log scale)"),
         x = var, y = "Claim amount (log10 scale)") +
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 45, hjust = 1))
}

plot_mean_amount_by_num_bins <- function(df, var, bins = 10) {
  df %>%
    filter(!is.na(.data[[var]]), !is.na(claim_amount), claim_amount > 0) %>%
    mutate(bin = cut(.data[[var]], breaks = bins, include.lowest = TRUE)) %>%
    group_by(bin) %>%
    summarise(n = n(), mean_claim_amount = mean(claim_amount, na.rm = TRUE), .groups = "drop") %>%
    ggplot(aes(x = bin, y = mean_claim_amount)) +
    geom_col(fill = "grey30") +
    scale_y_log10(labels = label_number()) +
    labs(title = paste0("Mean claim amount by binned ", var, " (log scale)"),
         x = paste0(var, " (binned)"), y = "Mean claim amount (log10 scale)") +
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 45, hjust = 1))
}

plot_amount_by_cat(wc_sev_eda, "solar_system")
plot_amount_by_cat(wc_sev_eda, "station_id")
plot_amount_by_cat(wc_sev_eda, "occupation")
plot_amount_by_cat(wc_sev_eda, "employment_type")
plot_amount_by_cat(wc_sev_eda, "accident_history_flag")
plot_amount_by_cat(wc_sev_eda, "psych_stress_index")
plot_amount_by_cat(wc_sev_eda, "hours_per_week")
plot_amount_by_cat(wc_sev_eda, "safety_training_index")
plot_amount_by_cat(wc_sev_eda, "protective_gear_quality")
plot_amount_by_cat(wc_sev_eda, "injury_type")
plot_amount_by_cat(wc_sev_eda, "injury_cause")

plot_mean_amount_by_num_bins(wc_sev_eda, "experience_yrs",    bins = 12)
plot_mean_amount_by_num_bins(wc_sev_eda, "supervision_level", bins = 10)
plot_mean_amount_by_num_bins(wc_sev_eda, "gravity_level",     bins = 10)
plot_mean_amount_by_num_bins(wc_sev_eda, "base_salary",       bins = 12)
