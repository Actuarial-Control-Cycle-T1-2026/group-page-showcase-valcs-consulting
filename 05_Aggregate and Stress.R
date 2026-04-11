###############################################################################
# 05_aggregate_stress.R
###############################################################################
# Version Control
# v1  | Vihaan Jain   | 01 Mar 2026 | Initial aggregate simulation structure
# v2  | Sarina Troung  | 06 Mar 2026 | Added simulation by solar system breakdown
# v3  | Vihaan Jain   | 10 Mar 2026 | Reinsurance (stop-loss) + BI-aligned expense/profit assumptions
# v4  | Vihaan Jain   | 12 Mar 2026 | Fixed bias correction (ratio of means, not mean of ratios)
###############################################################################

library(dplyr)
library(tidyr)
library(readxl)
library(purrr)
library(ggplot2)
library(scales)
library(MASS)
library(writexl)

set.seed(42)
N_SIM         <- 10000
VALUATION_YR  <- 2175
N_PROJ        <- 10
EXPENSE_RATIO <- 0.20
TARGET_PROFIT <- 0.08

DARK_NAVY  <- "#1F4E79"
MID_BLUE   <- "#2E75B6"
LIGHT_BLUE <- "#9DC3E6"

##############################################################################
# 0) CLEAN ENVIRONMENT
##############################################################################

keep <- c(
  "wc_freq", "wc_sev", "wc_freq_clean", "wc_sev_clean",
  "wc_freq_eda", "wc_sev_eda",
  "freq_pricing", "sev_pricing", "sev_pricing_sigma2", "sev_pricing_is_lognormal",
  "freq_stress", "sev_stress", "sev_stress_sigma2",
  "get_sev_preds", "factor_vars_freq", "factor_vars_sev",
  "wc_freq_model", "wc_sev_model",
  "plot_amount_by_cat", "plot_count_by_cat",
  "plot_mean_amount_by_num_bins", "plot_mean_count_by_num_bins",
  "cq_portfolio",
  "N_SIM", "VALUATION_YR", "N_PROJ", "EXPENSE_RATIO", "TARGET_PROFIT",
  "DARK_NAVY", "MID_BLUE", "LIGHT_BLUE"
)
rm(list = setdiff(ls(), keep))

##############################################################################
# 1) INFLATION TERM STRUCTURE
##############################################################################

rates_raw <- read_excel(
  "~/Library/Mobile Documents/com~apple~CloudDocs/ACTL4001 Assignment/SOA_2026_Case_Study_Materials/0. Original Case Material/srcsc-2026-interest-and-inflation.xlsx",
  sheet = "Sheet1", skip = 2) %>%
  setNames(c("year", "inflation", "overnight", "spot_1y", "spot_10y")) %>%
  mutate(across(everything(), as.numeric)) %>%
  filter(!is.na(year), !is.na(inflation))

avg3 <- rates_raw %>% tail(3) %>% summarise(across(c(inflation, spot_1y, spot_10y), mean))
avg5 <- rates_raw %>% tail(5) %>% summarise(across(c(inflation, spot_1y, spot_10y), mean))
avgA <- rates_raw               %>% summarise(across(c(inflation, spot_1y, spot_10y), mean))

get_rates <- function(t) {
  if      (t <= 2) list(inf = avg3$inflation, disc = avg3$spot_1y)
  else if (t <= 5) list(inf = avg5$inflation, disc = avg5$spot_1y)
  else             list(inf = avgA$inflation, disc = avgA$spot_10y)
}

cum_inf  <- numeric(N_PROJ + 1)
cum_disc <- numeric(N_PROJ + 1)
cum_inf[1] <- cum_disc[1] <- 1.0
for (t in 1:N_PROJ) {
  r <- get_rates(t)
  cum_inf[t + 1]  <- cum_inf[t]  * (1 + r$inf)
  cum_disc[t + 1] <- cum_disc[t] * (1 / (1 + r$disc))
}

inflation_tbl <- tibble(
  year            = VALUATION_YR + 0:N_PROJ,
  t               = 0:N_PROJ,
  cum_inf_factor  = round(cum_inf,  6),
  cum_disc_factor = round(cum_disc, 6)
)
cat("--- Inflation term structure ---\n")
print(inflation_tbl)

##############################################################################
# 2) AUGMENT cq_portfolio
##############################################################################

set.seed(42)
n     <- nrow(cq_portfolio)
n_emp <- n

cq_portfolio <- cq_portfolio %>%
  mutate(
    supervision_level     = sample(wc_freq_clean$supervision_level, n, replace = TRUE),
    accident_history_flag = factor(
      sample(as.character(wc_freq_clean$accident_history_flag), n, replace = TRUE),
      levels = levels(wc_freq_model$accident_history_flag)
    ),
    psych_stress_index    = factor(
      sample(as.character(wc_freq_clean$psych_stress_index), n, replace = TRUE),
      levels = levels(wc_freq_model$psych_stress_index)
    ),
    safety_training_index = factor(
      sample(as.character(wc_freq_clean$safety_training_index), n, replace = TRUE),
      levels = levels(wc_freq_model$safety_training_index)
    ),
    injury_type  = factor(
      sample(as.character(wc_sev_clean$injury_type),  n, replace = TRUE),
      levels = levels(wc_sev_model$injury_type)
    ),
    injury_cause = factor(
      sample(as.character(wc_sev_clean$injury_cause), n, replace = TRUE),
      levels = levels(wc_sev_model$injury_cause)
    )
  )

##############################################################################
# 3) HELPER: SAFE PERCENTILE LEVEL FOR FACTOR VARIABLES
##############################################################################

pct_level <- function(x, p, model_levels) {
  vals    <- sort(unique(as.integer(as.character(model_levels))))
  target  <- quantile(as.integer(as.character(x)), p, na.rm = TRUE)
  closest <- vals[which.min(abs(vals - target))]
  as.character(closest)
}

##############################################################################
# 4) BUILD 5 STRESS PORTFOLIOS
##############################################################################

make_stress_portfolio <- function(base_df,
                                  psych_p       = NULL,
                                  safety_p      = NULL,
                                  acc_hist      = NULL,
                                  supervision_p = NULL) {
  out <- base_df
  
  if (!is.null(psych_p)) {
    lv  <- pct_level(wc_freq_clean$psych_stress_index, psych_p,
                     levels(wc_freq_model$psych_stress_index))
    out$psych_stress_index <- factor(lv, levels = levels(wc_freq_model$psych_stress_index))
  }
  if (!is.null(safety_p)) {
    lv  <- pct_level(wc_freq_clean$safety_training_index, 1 - safety_p,
                     levels(wc_freq_model$safety_training_index))
    out$safety_training_index <- factor(lv, levels = levels(wc_freq_model$safety_training_index))
  }
  if (!is.null(acc_hist)) {
    out$accident_history_flag <- factor(acc_hist,
                                        levels = levels(wc_freq_model$accident_history_flag))
  }
  if (!is.null(supervision_p)) {
    out$supervision_level <- quantile(wc_freq_clean$supervision_level,
                                      1 - supervision_p, na.rm = TRUE)
  }
  out
}

cq_s1 <- make_stress_portfolio(cq_portfolio, psych_p = 0.75, safety_p = 0.75, supervision_p = 0.75)
cq_s2 <- make_stress_portfolio(cq_portfolio, psych_p = 0.85, safety_p = 0.85, acc_hist = "1", supervision_p = 0.85)
cq_s3 <- cq_portfolio
cq_s4 <- cq_portfolio
cq_s5 <- make_stress_portfolio(cq_portfolio, psych_p = 0.90, safety_p = 0.90, acc_hist = "1", supervision_p = 0.90)

check_portfolio <- function(df, name) {
  vars <- c("psych_stress_index", "safety_training_index",
            "accident_history_flag", "supervision_level",
            "injury_type", "injury_cause")
  for (v in vars) {
    if (v %in% names(df)) {
      n_na <- sum(is.na(df[[v]]))
      if (n_na > 0) cat(sprintf("  WARNING [%s] %s: %d NAs\n", name, v, n_na))
    }
  }
  cat(sprintf("  [%s] portfolio check complete (%d rows)\n", name, nrow(df)))
}

cat("\nPortfolio NA checks:\n")
check_portfolio(cq_portfolio, "Base")
check_portfolio(cq_s1, "S1")
check_portfolio(cq_s2, "S2")
check_portfolio(cq_s5, "S5")

##############################################################################
# 5) BIAS CORRECTION — FIXED: ratio of means, not mean of ratios
##############################################################################

raw_sev_train <- get_sev_preds(sev_pricing, wc_sev_model,
                               lognormal = sev_pricing_is_lognormal,
                               sigma2    = sev_pricing_sigma2)

bias_correction <- mean(wc_sev_model$claim_amount, na.rm = TRUE) /
  mean(raw_sev_train, na.rm = TRUE)

cat(sprintf("\nMean actual severity (training)     : %.0f\n",
            mean(wc_sev_model$claim_amount, na.rm = TRUE)))
cat(sprintf("Mean raw predicted severity         : %.0f\n",
            mean(raw_sev_train, na.rm = TRUE)))
cat(sprintf("Bias correction (ratio of means)    : %.4f\n", bias_correction))

cq_portfolio$expected_freq <- predict(freq_pricing, newdata = cq_portfolio, type = "response")
cq_portfolio$expected_sev  <- get_sev_preds(sev_pricing, cq_portfolio,
                                            lognormal = sev_pricing_is_lognormal,
                                            sigma2    = sev_pricing_sigma2) * bias_correction
cq_portfolio$expected_pp   <- cq_portfolio$expected_freq * cq_portfolio$expected_sev

cat(sprintf("Mean expected freq                  : %.4f\n",
            mean(cq_portfolio$expected_freq, na.rm = TRUE)))
cat(sprintf("Mean expected sev (bias corrected)  : %.0f\n",
            mean(cq_portfolio$expected_sev,  na.rm = TRUE)))
cat(sprintf("Expected total claims/year          : %.0f\n",
            sum(cq_portfolio$expected_freq,  na.rm = TRUE)))
cat(sprintf("Total expected pure premium         : %.0f\n",
            sum(cq_portfolio$expected_pp,    na.rm = TRUE)))

##############################################################################
# 6) SIMULATION ENGINE
##############################################################################

run_simulation <- function(portfolio, freq_model, sev_model, sev_sigma2,
                           n_sim, bc, freq_mult = 1.0, label = "") {
  
  exp_freq <- predict(freq_model, newdata = portfolio, type = "response") * freq_mult
  log_mu   <- predict(sev_model,  newdata = portfolio)
  theta    <- freq_model$theta
  log_sig  <- sqrt(sev_sigma2)
  
  if (any(is.na(exp_freq))) {
    cat(sprintf("  [%s] %d NA exp_freq replaced with median\n", label, sum(is.na(exp_freq))))
    exp_freq[is.na(exp_freq)] <- median(exp_freq, na.rm = TRUE)
  }
  if (any(is.na(log_mu))) {
    cat(sprintf("  [%s] %d NA log_mu replaced with median\n", label, sum(is.na(log_mu))))
    log_mu[is.na(log_mu)] <- median(log_mu, na.rm = TRUE)
  }
  
  agg <- numeric(n_sim)
  for (i in seq_len(n_sim)) {
    counts       <- rnbinom(length(exp_freq), mu = exp_freq, size = theta)
    total_claims <- sum(counts)
    if (total_claims == 0) { agg[i] <- 0; next }
    emp_idx <- rep(seq_along(counts), times = counts)
    sevs    <- exp(rnorm(total_claims, mean = log_mu[emp_idx], sd = log_sig)) * bc
    agg[i]  <- sum(sevs)
  }
  agg
}

cat("\nRunning BASE simulation...\n")
agg_base <- run_simulation(cq_portfolio, freq_pricing, sev_pricing,
                           sev_pricing_sigma2, N_SIM, bias_correction,
                           freq_mult = 1.0, label = "Base")

cat("Running S1 — Mild Adverse...\n")
agg_s1 <- run_simulation(cq_s1, freq_pricing, sev_pricing,
                         sev_pricing_sigma2, N_SIM, bias_correction,
                         freq_mult = 1.10, label = "S1")

cat("Running S2 — Moderate Adverse...\n")
agg_s2 <- run_simulation(cq_s2, freq_pricing, sev_pricing,
                         sev_pricing_sigma2, N_SIM, bias_correction,
                         freq_mult = 1.20, label = "S2")

cat("S3 — High Inflation: Year 1 same as Base, shock applied in projection.\n")
agg_s3 <- agg_base

cat("S4 — Rapid Growth: Year 1 same as Base, shock applied in projection.\n")
agg_s4 <- agg_base

cat("Running S5 — Pandemic/Systemic...\n")
agg_s5 <- run_simulation(cq_s5, freq_stress, sev_stress,
                         sev_stress_sigma2, N_SIM, bias_correction,
                         freq_mult = 1.40, label = "S5")

cat("All simulations complete.\n")

cat("\n--- Simulation vector diagnostics ---\n")
for (nm in c("agg_base", "agg_s1", "agg_s2", "agg_s3", "agg_s4", "agg_s5")) {
  v <- get(nm)
  cat(sprintf("  %-10s  length: %5d  NA: %d  NaN: %d  zeros: %d  min: %.0f  max: %.0f\n",
              nm, length(v), sum(is.na(v)), sum(is.nan(v)),
              sum(v == 0, na.rm = TRUE), min(v, na.rm = TRUE), max(v, na.rm = TRUE)))
}

##############################################################################
# 7) REINSURANCE — PORTFOLIO STOP-LOSS TREATY
##############################################################################

REINS_ATTACH_PCT <- 0.99
REINS_LOAD       <- 1.20

apply_stoploss <- function(agg_vec, attach, limit = Inf) {
  ceded <- pmin(pmax(agg_vec - attach, 0), limit)
  net   <- agg_vec - ceded
  list(gross = agg_vec, net = net, ceded = ceded)
}

attach_wc <- quantile(agg_base, REINS_ATTACH_PCT, na.rm = TRUE)
limit_wc  <- Inf

cat(sprintf("\nReinsurance attachment (VaR %.0f%%): %.0f\n",
            REINS_ATTACH_PCT * 100, attach_wc))

sl_base <- apply_stoploss(agg_base, attach_wc, limit_wc)
sl_s1   <- apply_stoploss(agg_s1,   attach_wc, limit_wc)
sl_s2   <- apply_stoploss(agg_s2,   attach_wc, limit_wc)
sl_s3   <- apply_stoploss(agg_s3,   attach_wc, limit_wc)
sl_s4   <- apply_stoploss(agg_s4,   attach_wc, limit_wc)
sl_s5   <- apply_stoploss(agg_s5,   attach_wc, limit_wc)

agg_base_net <- sl_base$net
agg_s1_net   <- sl_s1$net
agg_s2_net   <- sl_s2$net
agg_s3_net   <- sl_s3$net
agg_s4_net   <- sl_s4$net
agg_s5_net   <- sl_s5$net

exp_ceded_base <- mean(sl_base$ceded, na.rm = TRUE)
reins_premium  <- REINS_LOAD * exp_ceded_base

cat(sprintf("Expected ceded loss (Base)    : %.0f\n", exp_ceded_base))
cat(sprintf("Reinsurance premium (%.0f%% load): %.0f\n", REINS_LOAD * 100, reins_premium))

##############################################################################
# 8) STATS FUNCTION
##############################################################################

calc_stats <- function(x, label, description) {
  n_na <- sum(is.na(x) | is.nan(x))
  if (n_na > 0)
    warning(sprintf("calc_stats [%s]: %d NA/NaN removed before stats.", label, n_na))
  x <- x[!is.na(x) & !is.nan(x)]
  
  tvar <- function(v, p) {
    thresh <- quantile(v, p, na.rm = TRUE)
    mean(v[v > thresh], na.rm = TRUE)
  }
  
  tibble(
    Scenario    = label,
    Description = description,
    N           = length(x),
    Mean        = mean(x,    na.rm = TRUE),
    Median      = median(x,  na.rm = TRUE),
    SD          = sd(x,      na.rm = TRUE),
    Variance    = var(x,     na.rm = TRUE),
    CV          = round(sd(x, na.rm = TRUE) / mean(x, na.rm = TRUE), 4),
    Skewness    = mean(((x - mean(x, na.rm = TRUE)) /
                          sd(x, na.rm = TRUE))^3, na.rm = TRUE),
    VaR_95      = quantile(x, 0.95,  na.rm = TRUE),
    VaR_99      = quantile(x, 0.99,  na.rm = TRUE),
    VaR_99_5    = quantile(x, 0.995, na.rm = TRUE),
    TVaR_95     = tvar(x, 0.95),
    TVaR_99     = tvar(x, 0.99),
    TVaR_99_5   = tvar(x, 0.995),
    Range_Low   = quantile(x, 0.025, na.rm = TRUE),
    Range_High  = quantile(x, 0.975, na.rm = TRUE)
  )
}

##############################################################################
# 9) SHORT-TERM STATISTICS
##############################################################################

short_stats_gross <- bind_rows(
  calc_stats(agg_base, "Base", "Central estimate — sampled risk profile"),
  calc_stats(agg_s1,   "S1",   "Mild Adverse: 75th pct risk factors, +10% freq"),
  calc_stats(agg_s2,   "S2",   "Moderate Adverse: 85th pct risk factors, +20% freq"),
  calc_stats(agg_s3,   "S3",   "High Inflation: +2pp inflation shock (Year 1 = Base)"),
  calc_stats(agg_s4,   "S4",   "Rapid Growth: +10pp headcount (Year 1 = Base)"),
  calc_stats(agg_s5,   "S5",   "Pandemic/Systemic: 90th pct risk factors, +40% freq")
)

short_stats_net <- bind_rows(
  calc_stats(agg_base_net, "Base", "Net of reinsurance"),
  calc_stats(agg_s1_net,   "S1",   "Net of reinsurance"),
  calc_stats(agg_s2_net,   "S2",   "Net of reinsurance"),
  calc_stats(agg_s3_net,   "S3",   "Net of reinsurance"),
  calc_stats(agg_s4_net,   "S4",   "Net of reinsurance"),
  calc_stats(agg_s5_net,   "S5",   "Net of reinsurance")
)

cat("\n╔══════════════════════════════════════════════════════════════════╗\n")
cat("║   SHORT-TERM AGGREGATE LOSS — GROSS (YEAR 1, 2175)              ║\n")
cat("╚══════════════════════════════════════════════════════════════════╝\n")
short_stats_gross %>%
  mutate(across(c(Mean, Median, SD, VaR_95, VaR_99, VaR_99_5,
                  TVaR_95, TVaR_99, TVaR_99_5, Range_Low, Range_High), ~ round(., 0)),
         Skewness = round(Skewness, 3)) %>%
  print(width = Inf)

cat("\n╔══════════════════════════════════════════════════════════════════╗\n")
cat("║   SHORT-TERM AGGREGATE LOSS — NET OF REINSURANCE (YEAR 1, 2175) ║\n")
cat("╚══════════════════════════════════════════════════════════════════╝\n")
short_stats_net %>%
  mutate(across(c(Mean, Median, SD, VaR_95, VaR_99, VaR_99_5,
                  TVaR_95, TVaR_99, TVaR_99_5, Range_Low, Range_High), ~ round(., 0)),
         Skewness = round(Skewness, 3)) %>%
  print(width = Inf)

##############################################################################
# 10) PREMIUM DERIVATION
##############################################################################

pure_premium_net   <- mean(agg_base_net, na.rm = TRUE) / n_emp
risk_margin_pct    <- (quantile(agg_base_net, 0.995, na.rm = TRUE) -
                         mean(agg_base_net, na.rm = TRUE)) /
  mean(agg_base_net, na.rm = TRUE)

net_loaded_pp      <- pure_premium_net * (1 + risk_margin_pct)
reins_cost_pp      <- reins_premium / n_emp
gross_premium_pp   <- (net_loaded_pp + reins_cost_pp) / (1 - EXPENSE_RATIO - TARGET_PROFIT)
gross_premium_total <- gross_premium_pp * n_emp

cat("\n╔══════════════════════════════════════════════╗\n")
cat("║   PREMIUM DERIVATION (NET OF REINSURANCE)    ║\n")
cat("╚══════════════════════════════════════════════╝\n")
cat(sprintf("  Pure premium — net (per employee)   : %12.2f\n", pure_premium_net))
cat(sprintf("  Risk margin — VaR99.5 on net        : %11.2f%%\n", risk_margin_pct * 100))
cat(sprintf("  Reinsurance cost (per employee)     : %12.2f\n", reins_cost_pp))
cat(sprintf("  Expense ratio                       : %11.2f%%\n", EXPENSE_RATIO * 100))
cat(sprintf("  Target profit margin                : %11.2f%%\n", TARGET_PROFIT * 100))
cat(sprintf("  Gross premium (per employee)        : %12.2f\n", gross_premium_pp))
cat(sprintf("  Gross premium (total portfolio)     : %12.0f\n", gross_premium_total))

##############################################################################
# 11) DENSITY PLOT — YEAR 1 GROSS vs NET
##############################################################################
ggplot() +
  geom_density(aes(x = agg_base,     fill = "Base (Gross)"),    alpha = 0.35, colour = NA) +
  geom_density(aes(x = agg_base_net, fill = "Base (Net)"),      alpha = 0.45, colour = NA) +
  geom_density(aes(x = agg_s2,       fill = "S2: Mod (Gross)"), alpha = 0.35, colour = NA) +
  geom_density(aes(x = agg_s2_net,   fill = "S2: Mod (Net)"),   alpha = 0.45, colour = NA) +
  geom_vline(xintercept = attach_wc,
             colour = DARK_NAVY, linetype = "dashed", linewidth = 0.8) +
  annotate("text", x = attach_wc, y = Inf, vjust = 1.5, hjust = -0.1,
           size = 3, colour = DARK_NAVY,
           label = sprintf("Stop-loss attachment\nĐ%.2fM", attach_wc / 1e6)) +
  scale_x_continuous(labels = comma) +
  scale_fill_manual(values = c(
    "Base (Gross)"    = LIGHT_BLUE,
    "Base (Net)"      = DARK_NAVY,
    "S2: Mod (Gross)" = "#90EE90",
    "S2: Mod (Net)"   = "forestgreen"
  )) +
  labs(
    title    = "WC Aggregate Loss — Gross vs Net of Reinsurance (Year 1, 2175)",
    subtitle = "Stop-loss treaty hard-caps retained exposure at attachment | S5 Pandemic excluded (see Appendix)",
    x = "Aggregate Loss", y = "Density", fill = NULL
  ) +
  theme_minimal() +
  theme(legend.position = "bottom")

##############################################################################
# 11b) SIMULATED vs EMPIRICAL OVERLAY
##############################################################################
set.seed(42)

B <- 10000

# empirical frequency distribution
emp_freq <- wc_freq_eda %>%
  group_by(policy_id) %>%
  summarise(n_claims = n(), .groups = "drop") %>%
  pull(n_claims)

# empirical severities
emp_sev <- wc_sev_eda$claim_amount

emp_agg <- numeric(B)

for(i in 1:B){
  
  # sample claim count
  N <- sample(emp_freq, 1, replace = TRUE)
  
  if(N == 0){
    emp_agg[i] <- 0
  } else {
    emp_agg[i] <- sum(sample(emp_sev, N, replace = TRUE))
  }
  
}
ggplot() +
  
  geom_density(aes(x = agg_base, fill = "Simulated (GLM)"),
               alpha = 0.45) +
  
  geom_density(aes(x = emp_agg, fill = "Empirical (Bootstrap)"),
               alpha = 0.45) +
  
  geom_vline(xintercept = mean(agg_base),
             colour = DARK_NAVY,
             linetype = "dashed") +
  
  geom_vline(xintercept = mean(emp_agg),
             colour = MID_BLUE,
             linetype = "dashed") +
  
  scale_fill_manual(values = c(
    "Simulated (GLM)" = MID_BLUE,
    "Empirical (Bootstrap)" = LIGHT_BLUE
  )) +
  
  labs(
    title = "Aggregate Loss Distribution: Model vs Empirical",
    subtitle = "Workers Compensation – Year 1 (2175)",
    x = "Aggregate Loss",
    y = "Density",
    fill = NULL
  ) +
  
  theme_minimal()

##############################################################################
# 12) LONG-TERM PROJECTION ENGINE
##############################################################################

cum_inf_s3 <- numeric(N_PROJ + 1)
cum_inf_s3[1] <- 1.0
for (t in 1:N_PROJ) {
  r <- get_rates(t)
  cum_inf_s3[t + 1] <- cum_inf_s3[t] * (1 + r$inf + 0.02)
}

solar_growth_base <- c("Helionis Cluster" = 0.25, "Epsilon" = 0.25, "Zeta" = 0.15)
solar_growth_s4   <- c("Helionis Cluster" = 0.35, "Epsilon" = 0.35, "Zeta" = 0.25)

project_year_full <- function(t, sim_vec,
                              growth_map = solar_growth_base,
                              inf_vec    = cum_inf) {
  x <- sim_vec[!is.na(sim_vec) & !is.nan(sim_vec)]
  
  growth_factors  <- growth_map[as.character(cq_portfolio$solar_system)]
  emp_scale       <- 1 + (growth_factors / N_PROJ) * t
  headcount_scale <- mean(emp_scale, na.rm = TRUE)
  scaled          <- x * headcount_scale * inf_vec[t + 1]
  
  tvar <- function(v, p) mean(v[v > quantile(v, p, na.rm = TRUE)], na.rm = TRUE)
  
  tibble(
    year            = VALUATION_YR + t,
    t               = t,
    cum_inf_factor  = round(inf_vec[t + 1], 4),
    headcount_scale = round(headcount_scale, 4),
    mean_loss       = mean(scaled,                na.rm = TRUE),
    sd_loss         = sd(scaled,                  na.rm = TRUE),
    VaR_95          = quantile(scaled, 0.95,  na.rm = TRUE),
    VaR_99          = quantile(scaled, 0.99,  na.rm = TRUE),
    VaR_99_5        = quantile(scaled, 0.995, na.rm = TRUE),
    TVaR_95         = tvar(scaled, 0.95),
    TVaR_99         = tvar(scaled, 0.99),
    TVaR_99_5       = tvar(scaled, 0.995),
    lower_95ci      = quantile(scaled, 0.025, na.rm = TRUE),
    upper_95ci      = quantile(scaled, 0.975, na.rm = TRUE),
    lower_1sd       = mean(scaled, na.rm = TRUE) - sd(scaled, na.rm = TRUE),
    upper_1sd       = mean(scaled, na.rm = TRUE) + sd(scaled, na.rm = TRUE)
  )
}

long_base     <- map_dfr(0:N_PROJ, project_year_full, sim_vec = agg_base)
long_s1       <- map_dfr(0:N_PROJ, project_year_full, sim_vec = agg_s1)
long_s2       <- map_dfr(0:N_PROJ, project_year_full, sim_vec = agg_s2)
long_s3       <- map_dfr(0:N_PROJ, project_year_full, sim_vec = agg_s3, inf_vec = cum_inf_s3)
long_s4       <- map_dfr(0:N_PROJ, project_year_full, sim_vec = agg_s4, growth_map = solar_growth_s4)
long_s5       <- map_dfr(0:N_PROJ, project_year_full, sim_vec = agg_s5)

long_base_net <- map_dfr(0:N_PROJ, project_year_full, sim_vec = agg_base_net)
long_s1_net   <- map_dfr(0:N_PROJ, project_year_full, sim_vec = agg_s1_net)
long_s2_net   <- map_dfr(0:N_PROJ, project_year_full, sim_vec = agg_s2_net)
long_s3_net   <- map_dfr(0:N_PROJ, project_year_full, sim_vec = agg_s3_net, inf_vec = cum_inf_s3)
long_s4_net   <- map_dfr(0:N_PROJ, project_year_full, sim_vec = agg_s4_net, growth_map = solar_growth_s4)
long_s5_net   <- map_dfr(0:N_PROJ, project_year_full, sim_vec = agg_s5_net)

##############################################################################
# 13) NET REVENUE FUNCTION
##############################################################################

calc_net_revenue <- function(long_tbl, label) {
  long_tbl %>%
    mutate(
      scenario       = label,
      gross_premium  = gross_premium_total * headcount_scale * cum_inf[t + 1],
      reins_cost_yr  = reins_premium * headcount_scale * cum_inf[t + 1],
      expenses       = gross_premium * EXPENSE_RATIO,
      net_revenue    = gross_premium - mean_loss - reins_cost_yr - expenses,
      net_rev_low    = gross_premium - upper_95ci - reins_cost_yr - expenses,
      net_rev_high   = gross_premium - lower_95ci - reins_cost_yr - expenses,
      return_on_prem = net_revenue / gross_premium,
      loss_ratio     = mean_loss    / gross_premium,
      reins_ratio    = reins_cost_yr / gross_premium,
      combined_ratio = (mean_loss + reins_cost_yr + expenses) / gross_premium
    ) %>%
    dplyr::select(scenario, year, gross_premium, mean_loss, reins_cost_yr, expenses,
                  net_revenue, net_rev_low, net_rev_high,
                  return_on_prem, loss_ratio, reins_ratio, combined_ratio)
}

rev_base <- calc_net_revenue(long_base_net, "Base")
rev_s1   <- calc_net_revenue(long_s1_net,   "S1: Mild")
rev_s2   <- calc_net_revenue(long_s2_net,   "S2: Moderate")
rev_s3   <- calc_net_revenue(long_s3_net,   "S3: High Inflation")
rev_s4   <- calc_net_revenue(long_s4_net,   "S4: Rapid Growth")
rev_s5   <- calc_net_revenue(long_s5_net,   "S5: Pandemic")

cat("\n╔══════════════════════════════════════════════════════════════════╗\n")
cat("║   NET REVENUE — BASE SCENARIO (2175–2185, NET OF REINSURANCE)   ║\n")
cat("╚══════════════════════════════════════════════════════════════════╝\n")
rev_base %>%
  mutate(across(c(gross_premium, mean_loss, reins_cost_yr, expenses,
                  net_revenue, net_rev_low, net_rev_high), ~ round(., 0)),
         across(c(return_on_prem, loss_ratio, reins_ratio, combined_ratio), ~ round(., 4))) %>%
  print(width = Inf)

##############################################################################
# 14) 10-YEAR CUMULATIVE SUMMARY
##############################################################################

summary_tbl <- bind_rows(rev_base, rev_s1, rev_s2, rev_s3, rev_s4, rev_s5) %>%
  group_by(scenario) %>%
  summarise(
    `Cum Gross Premium`   = sum(gross_premium),
    `Cum Mean Loss (Net)` = sum(mean_loss),
    `Cum Reins Cost`      = sum(reins_cost_yr),
    `Cum Expenses`        = sum(expenses),
    `Cum Net Revenue`     = sum(net_revenue),
    `Net Rev Low`         = sum(net_rev_low),
    `Net Rev High`        = sum(net_rev_high),
    `Avg Loss Ratio`      = mean(loss_ratio),
    `Avg Reins Ratio`     = mean(reins_ratio),
    `Avg Combined Ratio`  = mean(combined_ratio),
    `Avg Return on Prem`  = mean(return_on_prem)
  )

cat("\n╔══════════════════════════════════════════════════════════════════╗\n")
cat("║   10-YEAR CUMULATIVE SUMMARY — ALL SCENARIOS (NET OF REINS)     ║\n")
cat("╚══════════════════════════════════════════════════════════════════╝\n")
summary_tbl %>%
  mutate(across(c(`Cum Gross Premium`, `Cum Mean Loss (Net)`, `Cum Reins Cost`,
                  `Cum Expenses`, `Cum Net Revenue`, `Net Rev Low`, `Net Rev High`),
                ~ round(., 0)),
         across(c(`Avg Loss Ratio`, `Avg Reins Ratio`, `Avg Combined Ratio`,
                  `Avg Return on Prem`), ~ round(., 4))) %>%
  print(width = Inf)

##############################################################################
# 15) SHORT-TERM COST / RETURN SUMMARY
##############################################################################

short_summary <- short_stats_net %>%
  mutate(
    gross_premium_yr1 = gross_premium_total,
    reins_cost_yr1    = reins_premium,
    expenses_yr1      = gross_premium_total * EXPENSE_RATIO,
    net_rev_mean      = gross_premium_total - Mean - reins_premium - gross_premium_total * EXPENSE_RATIO,
    net_rev_low       = gross_premium_total - Range_High - reins_premium - gross_premium_total * EXPENSE_RATIO,
    net_rev_high      = gross_premium_total - Range_Low  - reins_premium - gross_premium_total * EXPENSE_RATIO,
    loss_ratio        = Mean / gross_premium_total,
    reins_ratio       = reins_premium / gross_premium_total,
    return_on_prem    = net_rev_mean / gross_premium_total
  )

cat("\n╔══════════════════════════════════════════════════════════════════╗\n")
cat("║   SHORT-TERM: COSTS / RETURNS / NET REVENUE — YEAR 1 (NET)     ║\n")
cat("╚══════════════════════════════════════════════════════════════════╝\n")
short_summary %>%
  dplyr::select(Scenario, Description,
                Mean, SD, VaR_99_5, TVaR_99_5,
                Range_Low, Range_High,
                net_rev_mean, net_rev_low, net_rev_high,
                loss_ratio, reins_ratio, return_on_prem) %>%
  mutate(across(c(Mean, SD, VaR_99_5, TVaR_99_5,
                  Range_Low, Range_High,
                  net_rev_mean, net_rev_low, net_rev_high), ~ round(., 0)),
         across(c(loss_ratio, reins_ratio, return_on_prem), ~ round(., 4))) %>%
  print(width = Inf)

##############################################################################
# 16) LONG-TERM PROJECTION PLOT (NET OF REINSURANCE)
##############################################################################

scenario_colors <- c(
  "Base"             = DARK_NAVY,
  "S1: Mild"         = MID_BLUE,
  "S2: Moderate"     = LIGHT_BLUE,
  "S3: High Infl"    = "darkorange",
  "S4: Rapid Growth" = "forestgreen",
  "S5: Pandemic"     = "tomato"
)

proj_plot_data <- bind_rows(
  long_base_net %>% mutate(scenario = "Base"),
  long_s1_net   %>% mutate(scenario = "S1: Mild"),
  long_s2_net   %>% mutate(scenario = "S2: Moderate"),
  long_s3_net   %>% mutate(scenario = "S3: High Infl"),
  long_s4_net   %>% mutate(scenario = "S4: Rapid Growth"),
  long_s5_net   %>% mutate(scenario = "S5: Pandemic")
)

ggplot(proj_plot_data, aes(x = year, colour = scenario)) +
  geom_ribbon(data = proj_plot_data %>% filter(scenario == "Base"),
              aes(ymin = lower_1sd, ymax = upper_1sd),
              fill = MID_BLUE, alpha = 0.10, colour = NA) +
  geom_line(aes(y = mean_loss), linewidth = 1.0) +
  geom_line(aes(y = VaR_99_5),  linewidth = 0.5, linetype = "dashed") +
  scale_colour_manual(values = scenario_colors) +
  scale_x_continuous(breaks = VALUATION_YR:(VALUATION_YR + N_PROJ)) +
  scale_y_continuous(labels = comma) +
  labs(
    title    = "WC Projected Aggregate Loss — 2175 to 2185 (Net of Reinsurance)",
    subtitle = "Solid = Mean | Dashed = VaR(99.5%) | Band = Base ± 1SD",
    x = "Year", y = "Aggregate Loss (Net)", colour = "Scenario"
  ) +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

##############################################################################
# 17) NET REVENUE PLOT
##############################################################################

rev_plot_data <- bind_rows(rev_base, rev_s1, rev_s2, rev_s3, rev_s4, rev_s5)

ggplot(rev_plot_data, aes(x = year, y = net_revenue, colour = scenario)) +
  geom_hline(yintercept = 0, linetype = "dotted", colour = "grey40") +
  geom_ribbon(data = rev_plot_data %>% filter(scenario == "Base"),
              aes(ymin = net_rev_low, ymax = net_rev_high),
              fill = MID_BLUE, alpha = 0.12, colour = NA) +
  geom_line(linewidth = 1.0) +
  scale_colour_manual(values = scenario_colors) +
  scale_x_continuous(breaks = VALUATION_YR:(VALUATION_YR + N_PROJ)) +
  scale_y_continuous(labels = comma) +
  labs(
    title    = "WC Net Revenue — 2175 to 2185 (Net of Reinsurance)",
    subtitle = "Band = Base 95% CI | Dotted = breakeven",
    x = "Year", y = "Net Revenue", colour = "Scenario"
  ) +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

##############################################################################
# 17b-d) EXCL. PANDEMIC PLOTS
##############################################################################

scenario_colors_no_s5 <- scenario_colors[names(scenario_colors) != "S5: Pandemic"]
proj_plot_no_s5 <- proj_plot_data %>% filter(scenario != "S5: Pandemic")
rev_plot_no_s5  <- rev_plot_data  %>% filter(scenario != "S5: Pandemic")

ggplot(proj_plot_no_s5, aes(x = year, colour = scenario)) +
  geom_ribbon(data = proj_plot_no_s5 %>% filter(scenario == "Base"),
              aes(ymin = lower_1sd, ymax = upper_1sd),
              fill = MID_BLUE, alpha = 0.10, colour = NA) +
  geom_line(aes(y = mean_loss), linewidth = 1.0) +
  geom_line(aes(y = VaR_99_5),  linewidth = 0.5, linetype = "dashed") +
  scale_colour_manual(values = scenario_colors_no_s5) +
  scale_x_continuous(breaks = VALUATION_YR:(VALUATION_YR + N_PROJ)) +
  scale_y_continuous(labels = comma) +
  labs(title    = "WC Projected Aggregate Loss — 2175 to 2185 (excl. Pandemic, Net of Reins)",
       subtitle = "Solid = Mean | Dashed = VaR(99.5%) | Band = Base ± 1SD",
       x = "Year", y = "Aggregate Loss (Net)", colour = "Scenario") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

ggplot(rev_plot_no_s5, aes(x = year, y = net_revenue, colour = scenario)) +
  geom_hline(yintercept = 0, linetype = "dotted", colour = "grey40") +
  geom_ribbon(data = rev_plot_no_s5 %>% filter(scenario == "Base"),
              aes(ymin = net_rev_low, ymax = net_rev_high),
              fill = MID_BLUE, alpha = 0.12, colour = NA) +
  geom_line(linewidth = 1.0) +
  scale_colour_manual(values = scenario_colors_no_s5) +
  scale_x_continuous(breaks = VALUATION_YR:(VALUATION_YR + N_PROJ)) +
  scale_y_continuous(labels = comma) +
  labs(title    = "WC Net Revenue — 2175 to 2185 (excl. Pandemic, Net of Reins)",
       subtitle = "Band = Base 95% CI | Dotted = breakeven",
       x = "Year", y = "Net Revenue", colour = "Scenario") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

ggplot() +
  geom_density(aes(x = agg_base_net, fill = "Base"),         alpha = 0.40) +
  geom_density(aes(x = agg_s1_net,   fill = "S1: Mild"),     alpha = 0.40) +
  geom_density(aes(x = agg_s2_net,   fill = "S2: Moderate"), alpha = 0.40) +
  geom_vline(xintercept = quantile(agg_base_net, 0.995, na.rm = TRUE),
             colour = DARK_NAVY, linetype = "dashed", linewidth = 0.7) +
  geom_vline(xintercept = quantile(agg_s2_net,   0.995, na.rm = TRUE),
             colour = MID_BLUE,  linetype = "dashed", linewidth = 0.7) +
  scale_x_continuous(labels = comma) +
  scale_fill_manual(values = c(
    "Base"         = DARK_NAVY,
    "S1: Mild"     = MID_BLUE,
    "S2: Moderate" = LIGHT_BLUE
  )) +
  labs(title    = "WC Aggregate Loss Distribution — Year 1 (excl. Pandemic, Net of Reins)",
       subtitle = "Dashed = VaR(99.5%) for Base and S2",
       x = "Aggregate Loss (Net)", y = "Density", fill = "Scenario") +
  theme_minimal()

##############################################################################
# 18) PRODUCT DESIGN — SIMULATED RETAINED vs CEDED
##############################################################################

var99_gross  <- quantile(agg_base, 0.99,  na.rm = TRUE)
var995_gross <- quantile(agg_base, 0.995, na.rm = TRUE)
mean_gross   <- mean(agg_base, na.rm = TRUE)

cat(sprintf("\n--- Product Design Anchors ---\n"))
cat(sprintf("Portfolio mean loss      : %.0f\n", mean_gross))
cat(sprintf("Portfolio VaR(99%%)       : %.0f\n", var99_gross))
cat(sprintf("Portfolio VaR(99.5%%)     : %.0f\n", var995_gross))
cat(sprintf("Per-employee mean        : %.2f\n",  mean_gross  / n_emp))
cat(sprintf("Per-employee VaR(99%%)   : %.2f\n",  var99_gross / n_emp))

# Product A: per-claim cap + hard aggregate limit
per_claim_limit_A <- 200000
agg_limit_A       <- var99_gross

run_simulation_capped <- function(portfolio, freq_model, sev_model, sev_sigma2,
                                  n_sim, bc, per_claim_cap = Inf) {
  exp_freq <- predict(freq_model, newdata = portfolio, type = "response")
  log_mu   <- predict(sev_model,  newdata = portfolio)
  theta    <- freq_model$theta
  log_sig  <- sqrt(sev_sigma2)
  
  if (any(is.na(exp_freq))) exp_freq[is.na(exp_freq)] <- median(exp_freq, na.rm = TRUE)
  if (any(is.na(log_mu)))   log_mu[is.na(log_mu)]     <- median(log_mu,   na.rm = TRUE)
  
  agg_gross <- numeric(n_sim)
  agg_net_A <- numeric(n_sim)
  
  for (i in seq_len(n_sim)) {
    counts       <- rnbinom(length(exp_freq), mu = exp_freq, size = theta)
    total_claims <- sum(counts)
    if (total_claims == 0) next
    emp_idx      <- rep(seq_along(counts), times = counts)
    sevs         <- exp(rnorm(total_claims, mean = log_mu[emp_idx], sd = log_sig)) * bc
    agg_gross[i] <- sum(sevs)
    agg_net_A[i] <- min(sum(pmin(sevs, per_claim_cap)), agg_limit_A)
  }
  
  # Product B: stop-loss at VaR(99%)
  agg_net_B  <- pmin(agg_gross, attach_wc)
  ceded_B    <- pmax(agg_gross - attach_wc, 0)
  
  list(
    gross   = agg_gross,
    net_A   = agg_net_A,
    ceded_A = agg_gross - agg_net_A,
    net_B   = agg_net_B,
    ceded_B = ceded_B
  )
}

set.seed(42)
prod_sim <- run_simulation_capped(
  portfolio     = cq_portfolio,
  freq_model    = freq_pricing,
  sev_model     = sev_pricing,
  sev_sigma2    = sev_pricing_sigma2,
  n_sim         = N_SIM,
  bc            = bias_correction,
  per_claim_cap = per_claim_limit_A
)

product_summary <- data.frame(
  Metric = c(
    "Mean gross loss",
    "Mean insurer pays — Product A (capped)",
    "Mean insurer pays — Product B (stop-loss)",
    "Mean ceded — Product A",
    "Mean ceded — Product B",
    "VaR(99%) gross",
    "VaR(99%) net — Product A",
    "VaR(99%) net — Product B",
    "Max gross",
    "Max net — Product A",
    "Max net — Product B"
  ),
  Value = c(
    mean(prod_sim$gross),
    mean(prod_sim$net_A),
    mean(prod_sim$net_B),
    mean(prod_sim$ceded_A),
    mean(prod_sim$ceded_B),
    quantile(prod_sim$gross, 0.99),
    quantile(prod_sim$net_A, 0.99),
    quantile(prod_sim$net_B, 0.99),
    max(prod_sim$gross),
    max(prod_sim$net_A),
    max(prod_sim$net_B)
  )
) %>% mutate(Value = round(Value, 0))

cat("\n╔══════════════════════════════════════════════════════════════════╗\n")
cat("║   PRODUCT DESIGN — RETAINED vs CEDED COMPARISON                 ║\n")
cat("╚══════════════════════════════════════════════════════════════════╝\n")
print(product_summary, row.names = FALSE)

ggplot() +
  geom_density(aes(x = prod_sim$gross,  fill = "Gross (no limit)"),     alpha = 0.30, colour = NA) +
  geom_density(aes(x = prod_sim$net_A,  fill = "Product A (capped)"),   alpha = 0.45, colour = NA) +
  geom_density(aes(x = prod_sim$net_B,  fill = "Product B (stop-loss)"),alpha = 0.45, colour = NA) +
  geom_vline(xintercept = attach_wc,   colour = DARK_NAVY, linetype = "dashed", linewidth = 0.7) +
  geom_vline(xintercept = agg_limit_A, colour = MID_BLUE,  linetype = "dashed", linewidth = 0.7) +
  annotate("text", x = attach_wc,   y = Inf, vjust = 2,   hjust = -0.1,
           size = 2.8, colour = DARK_NAVY,
           label = sprintf("Product B attachment (VaR 99%%): %.0f", attach_wc)) +
  annotate("text", x = agg_limit_A, y = Inf, vjust = 4,   hjust = -0.1,
           size = 2.8, colour = MID_BLUE,
           label = sprintf("Product A agg limit (VaR 99%%): %.0f", agg_limit_A)) +
  scale_x_continuous(labels = comma) +
  scale_fill_manual(values = c(
    "Gross (no limit)"      = LIGHT_BLUE,
    "Product A (capped)"    = MID_BLUE,
    "Product B (stop-loss)" = DARK_NAVY
  )) +
  labs(
    title    = "WC Insurer Retained Loss — Product A vs Product B (Year 1, 2175)",
    subtitle = "Product A: per-claim Đ200K + agg cap | Product B: unlimited per-claim + stop-loss",
    x = "Aggregate Loss Retained", y = "Density", fill = NULL
  ) +
  theme_minimal() +
  theme(legend.position = "bottom")

##############################################################################
# 19) EXPORT
##############################################################################

setwd("~/Library/Mobile Documents/com~apple~CloudDocs/ACTL4001 Assignment/SOA_2026_Case_Study_Materials/Workers Compensation")

wc_out <- tibble(
  sim_id     = 1:N_SIM,
  agg_wc     = agg_base,
  agg_wc_net = agg_base_net,
  prem_wc    = gross_premium_total
)
write.csv(wc_out, "wc_simulation_output.csv", row.names = FALSE)

##############################################################################
# 20) BY SOLAR SYSTEM SIMULATION
##############################################################################

run_simulation_by_system <- function(portfolio, freq_model, sev_model, sev_sigma2,
                                     n_sim, bc, freq_mult = 1.0, label = "") {
  exp_freq <- predict(freq_model, newdata = portfolio, type = "response") * freq_mult
  log_mu   <- predict(sev_model,  newdata = portfolio)
  theta    <- freq_model$theta
  log_sig  <- sqrt(sev_sigma2)
  
  if (any(is.na(exp_freq))) exp_freq[is.na(exp_freq)] <- median(exp_freq, na.rm = TRUE)
  if (any(is.na(log_mu)))   log_mu[is.na(log_mu)]     <- median(log_mu,   na.rm = TRUE)
  
  solar_systems <- unique(portfolio$solar_system)
  agg_mat       <- matrix(0, nrow = n_sim, ncol = length(solar_systems))
  colnames(agg_mat) <- solar_systems
  
  for (i in seq_len(n_sim)) {
    counts <- rnbinom(length(exp_freq), mu = exp_freq, size = theta)
    for (s in seq_along(solar_systems)) {
      ss         <- solar_systems[s]
      idx        <- which(portfolio$solar_system == ss)
      counts_ss  <- counts[idx]
      total_ss   <- sum(counts_ss)
      if (total_ss == 0) next
      emp_idx_ss <- rep(idx, times = counts_ss)
      sevs_ss    <- exp(rnorm(total_ss, mean = log_mu[emp_idx_ss], sd = log_sig)) * bc
      agg_mat[i, s] <- sum(sevs_ss)
    }
  }
  
  agg_df       <- as.data.frame(agg_mat)
  agg_df$Total <- rowSums(agg_df)
  agg_df
}

cat("\nRunning BASE simulation by solar system...\n")
agg_base_by_system <- run_simulation_by_system(
  portfolio  = cq_portfolio,
  freq_model = freq_pricing,
  sev_model  = sev_pricing,
  sev_sigma2 = sev_pricing_sigma2,
  n_sim      = N_SIM,
  bc         = bias_correction,
  freq_mult  = 1.0,
  label      = "Base"
)

agg_export <- tibble(
  sim_id     = 1:N_SIM,
  agg_wc_h   = agg_base_by_system$`Helionis Cluster`,
  agg_wc_e   = agg_base_by_system$`Epsilon`,
  agg_wc_z   = agg_base_by_system$`Zeta`,
  agg_wc     = agg_base,
  agg_wc_net = agg_base_net
)

write_xlsx(
  list(Aggregate_Losses = agg_export),
  "aggregate_losses_by_solar_system.xlsx"
)

##############################################################################
# 21) FINAL KEY OBJECTS SUMMARY
##############################################################################

cat("\n╔══════════════════════════════════════════════╗\n")
cat("║   KEY OBJECTS SAVED                          ║\n")
cat("╚══════════════════════════════════════════════╝\n")
cat("Gross sim vectors  : agg_base, agg_s1–s5\n")
cat("Net sim vectors    : agg_base_net, agg_s1_net–s5_net\n")
cat("Gross long tables  : long_base, long_s1–s5\n")
cat("Net long tables    : long_base_net, long_s1_net–s5_net\n")
cat("Revenue tables     : rev_base, rev_s1–s5\n")
cat("Stat tables        : short_stats_gross, short_stats_net, short_summary, summary_tbl\n")
cat("Product sim        : prod_sim (gross, net_A, net_B, ceded_A, ceded_B)\n")
cat(sprintf("n_emp               : %d\n",    n_emp))
cat(sprintf("bias_correction     : %.4f\n",  bias_correction))
cat(sprintf("attach_wc           : %.0f\n",  attach_wc))
cat(sprintf("gross_premium_pp    : %.2f\n",  gross_premium_pp))
cat(sprintf("gross_premium_total : %.0f\n",  gross_premium_total))
cat(sprintf("risk_margin_pct     : %.2f%%\n", risk_margin_pct * 100))
cat(sprintf("reins_premium       : %.0f\n",  reins_premium))
cat(sprintf("pure_premium_net    : %.2f\n",  pure_premium_net))
cat(sprintf("per_claim_limit_A   : %.0f\n",  per_claim_limit_A))
cat(sprintf("agg_limit_A (VaR99) : %.0f\n",  agg_limit_A))




DARK_NAVY  <- "#1F4E79"
MID_BLUE   <- "#2E75B6"
LIGHT_BLUE <- "#9DC3E6"
plot_empirical_agg <- function(agg_vec, title_name) {
  ev      <- mean(agg_vec, na.rm = TRUE) / 1e6
  sd_val  <- sd(agg_vec, na.rm = TRUE) / 1e6
  var_995 <- quantile(agg_vec, 0.995, na.rm = TRUE) / 1e6
  
  df <- data.frame(loss = agg_vec / 1e6)
  
  ggplot(df, aes(x = loss)) +
    geom_histogram(
      aes(y = after_stat(density)),
      bins   = 60,
      fill   = LIGHT_BLUE,
      colour = "black",
      alpha  = 0.85
    ) +
    geom_density(
      colour    = "black",
      linewidth = 1.2
    ) +
    annotate("text",
             x = Inf, y = Inf,
             label = paste0(
               "Expected Loss = $", round(ev, 2), "m\n",
               "Standard Deviation = $", round(sd_val, 2), "m\n",
               "VaR (99.5%) = $", round(var_995, 2), "m"
             ),
             hjust = 1.1, vjust = 1.5, size = 4.5,
             family = "mono") +
    coord_cartesian(xlim = c(3.8, 7.5)) +
    labs(
      title    = "Simulated Aggregate Loss Distribution",
      subtitle = title_name,
      x        = "Aggregate Loss (Đ millions)",
      y        = "Density"
    ) +
    theme_minimal()
}

plot_empirical_agg(agg_base, "Workers Compensation")
