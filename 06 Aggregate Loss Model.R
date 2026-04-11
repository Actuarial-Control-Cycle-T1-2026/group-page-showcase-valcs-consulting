##############################################################################
# Cosmic Quarry Mining - Aggregate Loss & Capital Modelling
##############################################################################

# ---------------------------------------------------------------------------
# Load packages
# ---------------------------------------------------------------------------
library(tidyverse)
library(purrr)
library(scales)
library(writexl)

# ---------------------------------------------------------------------------
# Version control
# ---------------------------------------------------------------------------
# v1 | Sarina Truong | Initial aggregate loss simulation
# v2 | Coverage limits added
# v3 | Solar system aggregation added
# v4 | Cleaned and standardised

##############################################################################
# 1. Prepare Severity Predictions
##############################################################################

predictive_data <- predictive_data %>%
  mutate(
    cargo_group = case_when(
      tolower(cargo_type) == "gold" ~ "gold",
      tolower(cargo_type) == "platinum" ~ "platinum",
      TRUE ~ "other"
    )
  )

##############################################################################
# 2. Predict Frequency
##############################################################################

predictive_data$pred_claim_count <- predict(
  zinb_reduced,
  newdata = predictive_data,
  type = "response"
)

predictive_data$mu_count <- predict(
  zinb_reduced,
  newdata = predictive_data,
  type = "count"
)

predictive_data$pi_zero <- predict(
  zinb_reduced,
  newdata = predictive_data,
  type = "zero"
)

theta_nb <- zinb_reduced$theta

##############################################################################
# 3. Severity Model Variances
##############################################################################

sigma2_gold  <- summary(lognormal_gold_reduced)$sigma^2
sigma2_plat  <- summary(lognormal_platinum_reduced)$sigma^2
sigma2_other <- summary(lognormal_other_reduced)$sigma^2

sigma_gold  <- sqrt(sigma2_gold)
sigma_plat  <- sqrt(sigma2_plat)
sigma_other <- sqrt(sigma2_other)

##############################################################################
# 4. Predict Severity
##############################################################################

predictive_data$pred_log_severity <- NA_real_

gold_idx  <- predictive_data$cargo_group == "gold"
plat_idx  <- predictive_data$cargo_group == "platinum"
other_idx <- predictive_data$cargo_group == "other"

predictive_data$pred_log_severity[gold_idx] <-
  predict(lognormal_gold_reduced, newdata = predictive_data[gold_idx, ])

predictive_data$pred_log_severity[plat_idx] <-
  predict(lognormal_platinum_reduced, newdata = predictive_data[plat_idx, ])

predictive_data$pred_log_severity[other_idx] <-
  predict(lognormal_other_reduced, newdata = predictive_data[other_idx, ])

##############################################################################
# 5. Expected Loss
##############################################################################

predictive_data$pred_severity <- case_when(
  predictive_data$cargo_group == "gold" ~
    exp(predictive_data$pred_log_severity + 0.5 * sigma2_gold),
  
  predictive_data$cargo_group == "platinum" ~
    exp(predictive_data$pred_log_severity + 0.5 * sigma2_plat),
  
  TRUE ~
    exp(predictive_data$pred_log_severity + 0.5 * sigma2_other)
)

predictive_data$expected_loss <-
  predictive_data$pred_claim_count *
  predictive_data$pred_severity

##############################################################################
# 6. Portfolio Summary
##############################################################################

total_expected_loss <-
  sum(predictive_data$expected_loss, na.rm = TRUE)

average_expected_loss <-
  mean(predictive_data$expected_loss, na.rm = TRUE)

##############################################################################
# 7. Frequency Simulation
##############################################################################

simulate_zinb_counts <- function(mu, pi_zero, theta) {
  
  structural_zero <-
    rbinom(length(mu), 1, pi_zero)
  
  ifelse(
    structural_zero == 1,
    0,
    rnbinom(length(mu), mu = mu, size = theta)
  )
}

##############################################################################
# 8. Severity Simulation
##############################################################################

simulate_row_severity <- function(
    n_claims,
    pred_log_severity,
    cargo_group
){
  
  if(n_claims == 0) return(0)
  
  sigma <- case_when(
    cargo_group == "gold" ~ sigma_gold,
    cargo_group == "platinum" ~ sigma_plat,
    TRUE ~ sigma_other
  )
  
  sum(
    rlnorm(
      n_claims,
      meanlog = pred_log_severity,
      sdlog = sigma
    )
  )
}

##############################################################################
# 9. Aggregate Loss Simulation
##############################################################################

simulate_aggregate_loss <- function(df){
  
  counts <- simulate_zinb_counts(
    df$mu_count,
    df$pi_zero,
    theta_nb
  )
  
  losses <- map2_dbl(
    counts,
    seq_along(counts),
    ~ simulate_row_severity(
      .x,
      df$pred_log_severity[.y],
      df$cargo_group[.y]
    )
  )
  
  sum(losses)
}

##############################################################################
# 10. Monte Carlo Simulation
##############################################################################

set.seed(123)

n_sim <- 10000

aggregate_losses <- replicate(
  n_sim,
  simulate_aggregate_loss(predictive_data)
)

##############################################################################
# 11. Distribution Summary
##############################################################################

loss_summary <- tibble(
  mean = mean(aggregate_losses),
  sd = sd(aggregate_losses),
  p95 = quantile(aggregate_losses,0.95),
  p99 = quantile(aggregate_losses,0.99),
  p995 = quantile(aggregate_losses,0.995),
  max = max(aggregate_losses)
)

loss_summary

##############################################################################
# 12. Plot Distribution
##############################################################################

ggplot(
  data.frame(loss = aggregate_losses),
  aes(x = loss)
) +
  geom_histogram(
    bins = 50,
    fill = "#9DC3E6",
    colour = "black"
  ) +
  geom_density() +
  theme_minimal() +
  scale_x_continuous(
    labels = label_number(scale = 1e-6)
  ) +
  labs(
    title = "Aggregate Loss Distribution",
    x = "Aggregate Loss ($m)"
  )

##############################################################################
# 13. Aggregate Loss by Solar System
##############################################################################

simulate_by_system <- function(df){
  
  counts <- simulate_zinb_counts(
    df$mu_count,
    df$pi_zero,
    theta_nb
  )
  
  df$sim_loss <-
    map2_dbl(
      counts,
      seq_along(counts),
      ~ simulate_row_severity(
        .x,
        df$pred_log_severity[.y],
        df$cargo_group[.y]
      )
    )
  
  df %>%
    group_by(solar_system) %>%
    summarise(
      loss = sum(sim_loss),
      .groups="drop"
    )
}

agg_by_system <- replicate(
  n_sim,
  simulate_by_system(predictive_data),
  simplify = FALSE
)

##############################################################################
# 14. Capital Metrics
##############################################################################

VaR_995 <- quantile(aggregate_losses,0.995)

required_capital <-
  VaR_995 -
  mean(aggregate_losses)

TVaR <- mean(
  aggregate_losses[
    aggregate_losses > VaR_995
  ]
)

capital_summary <- tibble(
  Mean = mean(aggregate_losses),
  VaR_995 = VaR_995,
  TVaR_995 = TVaR,
  Required_Capital = required_capital
)

capital_summary