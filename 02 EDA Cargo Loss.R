##############################################################################
# Cosmic Quarry Mining - Cargo Loss Exploratory Data Analysis
##############################################################################

# ---------------------------------------------------------------------------
# Load packages
# ---------------------------------------------------------------------------
library(readxl)
library(dplyr)
library(ggplot2)
library(tidyr)
library(scales)
library(GGally)
library(patchwork)
library(viridis)

# ---------------------------------------------------------------------------
# Version control
# ---------------------------------------------------------------------------
# v1 | Sarina Truong | Initial EDA
# v2 | Cleaned       | Standardised structure and naming

# ---------------------------------------------------------------------------
# File paths
# ---------------------------------------------------------------------------
original_file_path <- "~/Library/Mobile Documents/com~apple~CloudDocs/ACTL4001 Assignment/SOA_2026_Case_Study_Materials/0. Original Case Material/srcsc-2026-claims-cargo.xlsx"

clean_file_path <- "~/Library/Mobile Documents/com~apple~CloudDocs/ACTL4001 Assignment/SOA_2026_Case_Study_Materials/Cargo Loss/cargo_loss_CLEAN.xlsx"

##############################################################################
# 1. Import Data
##############################################################################

# Original raw datasets
cargo_frequency_original <- read_excel(original_file_path, sheet = "freq")
cargo_severity_original  <- read_excel(original_file_path, sheet = "sev")

# Cleaned datasets (from Script 1)
cargo_frequency <- read_excel(clean_file_path, sheet = "Frequency")
cargo_severity  <- read_excel(clean_file_path, sheet = "Severity")

##############################################################################
# 2. Remove duplicate rows
##############################################################################

sum(duplicated(cargo_frequency))
sum(duplicated(cargo_severity))

cargo_frequency <- cargo_frequency %>% distinct()
cargo_severity  <- cargo_severity  %>% distinct()

##############################################################################
# 3. Remove ID variables
##############################################################################

# Frequency dataset
frequency <- cargo_frequency[, -c(1,2)]

# Severity dataset
severity <- cargo_severity[, -c(1,2,3,4)]

##############################################################################
# 4. Frequency EDA
##############################################################################
# Claim count exploratory analysis

# ------------------------------------------------------------
# Cargo type composition by container
# ------------------------------------------------------------

cargo_mix <- frequency %>%
  group_by(container_type, cargo_type) %>%
  summarise(
    count = n(),
    .groups = "drop"
  ) %>%
  group_by(container_type) %>%
  mutate(
    percentage = count / sum(count) * 100
  )

ggplot(
  cargo_mix,
  aes(
    x = container_type,
    y = percentage,
    fill = cargo_type
  )
) +
  geom_col() +
  labs(
    title = "Cargo Type Composition by Container Type",
    x = "Container Type",
    y = "Percentage of Shipments",
    fill = "Cargo Type"
  ) +
  theme_minimal()

##############################################################################
# Weight distribution by cargo and container
##############################################################################

frequency %>%
  filter(!is.na(container_type), !is.na(cargo_type), !is.na(weight)) %>%
  ggplot(aes(x = cargo_type, y = weight)) +
  geom_boxplot() +
  facet_wrap(~ container_type) +
  labs(
    title = "Weight Distribution by Cargo and Container",
    x = "Cargo Type",
    y = "Weight"
  ) +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

##############################################################################
# Weight band distribution
##############################################################################

frequency %>%
  filter(!is.na(container_type), !is.na(cargo_type), !is.na(weight)) %>%
  mutate(
    weight_band = cut(
      weight,
      breaks = quantile(weight, probs = seq(0,1,0.2), na.rm = TRUE),
      include.lowest = TRUE
    )
  ) %>%
  group_by(container_type, cargo_type, weight_band) %>%
  summarise(count = n(), .groups = "drop") %>%
  group_by(container_type, cargo_type) %>%
  mutate(percentage = count / sum(count) * 100) %>%
  ggplot(aes(x = cargo_type, y = percentage, fill = weight_band)) +
  geom_col() +
  facet_wrap(~ container_type) +
  theme_minimal() +
  labs(
    title = "Weight Distribution by Container Type",
    x = "Cargo Type",
    y = "Percentage"
  ) +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

##############################################################################
# Average weight heatmap
##############################################################################

weight_summary <- frequency %>%
  group_by(container_type, cargo_type) %>%
  summarise(
    mean_weight = mean(weight),
    .groups = "drop"
  )

ggplot(
  weight_summary,
  aes(
    x = cargo_type,
    y = container_type,
    fill = mean_weight
  )
) +
  geom_tile() +
  theme_minimal() +
  labs(
    title = "Average Weight by Cargo and Container"
  )

##############################################################################
# Route risk vs claim frequency
##############################################################################

frequency %>%
  group_by(route_risk) %>%
  summarise(
    mean_claim = mean(claim_count, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  ggplot(aes(x = factor(route_risk), y = mean_claim)) +
  geom_col() +
  theme_minimal() +
  labs(
    title = "Mean Claim Count by Route Risk"
  )

##############################################################################
# Vessel age effect
##############################################################################

frequency %>%
  mutate(age_bin = ntile(vessel_age, 10)) %>%
  group_by(age_bin) %>%
  summarise(
    age_mid = median(vessel_age),
    mean_claim = mean(claim_count),
    .groups = "drop"
  ) %>%
  ggplot(aes(x = age_mid, y = mean_claim)) +
  geom_line() +
  geom_point() +
  theme_minimal() +
  labs(
    title = "Mean Claim Count vs Vessel Age"
  )

##############################################################################
# Pilot experience effect
##############################################################################

frequency %>%
  mutate(exp_bin = ntile(pilot_experience, 10)) %>%
  group_by(exp_bin) %>%
  summarise(
    exp_mid = median(pilot_experience),
    mean_claim = mean(claim_count),
    .groups = "drop"
  ) %>%
  ggplot(aes(x = exp_mid, y = mean_claim)) +
  geom_line() +
  geom_point() +
  theme_minimal()

##############################################################################
# 5. Severity EDA
##############################################################################

severity <- severity %>%
  mutate(
    cargo_type = as.factor(cargo_type),
    container_type = as.factor(container_type),
    route_risk = as.factor(route_risk),
    log_claim_amount = log1p(claim_amount),
    loss_ratio = claim_amount / cargo_value
  )

##############################################################################
# Claim distribution
##############################################################################

ggplot(severity, aes(x = log_claim_amount)) +
  geom_histogram(bins = 40, fill = "#9DC3E6", color = "black") +
  theme_minimal() +
  labs(
    title = "Log Claim Distribution"
  )

##############################################################################
# Claim vs cargo type
##############################################################################

ggplot(severity, aes(x = cargo_type, y = claim_amount)) +
  geom_boxplot() +
  scale_y_log10(labels = comma) +
  theme_minimal()

##############################################################################
# Claim vs cargo value
##############################################################################

ggplot(severity, aes(x = cargo_value, y = claim_amount)) +
  geom_point(alpha = 0.3) +
  geom_smooth(method = "loess", se = FALSE) +
  scale_x_log10(labels = comma) +
  scale_y_log10(labels = comma) +
  theme_minimal()

##############################################################################
# Binned helper function
##############################################################################

plot_binned <- function(df, x, y = "loss_ratio", bins = 12) {
  
  df %>%
    filter(is.finite(.data[[x]]), is.finite(.data[[y]])) %>%
    mutate(bin = ntile(.data[[x]], bins)) %>%
    group_by(bin) %>%
    summarise(
      x_mid = median(.data[[x]]),
      y_med = median(.data[[y]]),
      n = n(),
      .groups = "drop"
    ) %>%
    ggplot(aes(x = x_mid, y = y_med)) +
    geom_line() +
    geom_point() +
    theme_minimal()
}

##############################################################################
# Pilot experience vs severity
##############################################################################

plot_binned(severity, "pilot_experience", "loss_ratio")

##############################################################################
# Vessel age vs severity
##############################################################################

plot_binned(severity, "vessel_age", "loss_ratio")

##############################################################################
# Heatmap radiation vs debris
##############################################################################

severity_heat <- severity %>%
  mutate(
    rad_bin = cut_number(solar_radiation, n = 5),
    debris_bin = cut_number(debris_density, n = 5)
  ) %>%
  group_by(rad_bin, debris_bin) %>%
  summarise(
    median_claim = median(claim_amount),
    n = n(),
    .groups = "drop"
  )

ggplot(
  severity_heat,
  aes(
    x = rad_bin,
    y = debris_bin,
    fill = median_claim
  )
) +
  geom_tile() +
  scale_fill_viridis_c() +
  theme_minimal()

##############################################################################
# Bimodality check
##############################################################################

ggplot(severity, aes(x = log_claim_amount)) +
  geom_histogram(bins = 35, fill = "#9DC3E6") +
  facet_wrap(~ cargo_type) +
  theme_minimal()

##############################################################################
# Cargo value band split
##############################################################################

cut_point <- quantile(severity$cargo_value, 0.5, na.rm = TRUE)

severity_band <- severity %>%
  mutate(
    value_band = if_else(
      cargo_value >= cut_point,
      "High Value",
      "Low Value"
    )
  )

ggplot(severity_band, aes(x = log_claim_amount)) +
  geom_histogram(bins = 35) +
  facet_wrap(~ value_band) +
  theme_minimal()