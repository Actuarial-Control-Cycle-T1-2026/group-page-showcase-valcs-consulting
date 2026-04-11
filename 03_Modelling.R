###############################################################################
# 02_freq_sev_modelling.R
# Frequency x Severity modelling — SRCSC 2026 Workers Compensation
###############################################################################
# Version Control
# v1  | Vihaan Jain  | 01 Mar 2026 | Initial freq/sev model structure, NB + LN GLMs
# v2  | Vihaan Jain  | 01 Mar 2026 | Added GBM variable importance, train/test split
# v3  | Vihaan Jain  | 02 Mar 2026 | Severity tail diagnostics + distribution fitting
# v4  | Vihaan Jain  | 02 Mar 2026 | ZINB auto-select logic, CQ deployability flags
# v5  | Vihaan Jain  | 03 Mar 2026 | Residual diagnostics, QQ plots, final export block
###############################################################################

library(dplyr)
library(MASS)
library(gbm)
library(ggplot2)
library(scales)
library(pscl)
library(tidyr)
library(purrr)
library(actuar)
library(fitdistrplus)
library(moments)
library(stringr)
library(statmod)   # tweedie() GLM family — statmod only, do NOT load tweedie pkg

set.seed(123)


# ---- 1) Prep ----

factor_vars_freq <- c("occupation", "employment_type", "solar_system",
                      "station_id", "accident_history_flag", "psych_stress_index",
                      "hours_per_week", "safety_training_index",
                      "protective_gear_quality")

factor_vars_sev <- c(factor_vars_freq, "injury_type", "injury_cause")

wc_freq_model <- wc_freq_eda %>%
  mutate(across(all_of(factor_vars_freq), as.factor),
         high_supervision = as.factor(as.integer(supervision_level > 0.8)),
         log_exposure     = log(exposure)) %>%
  filter(exposure > 0)

wc_sev_model <- wc_sev_eda %>%
  mutate(across(all_of(factor_vars_sev), as.factor),
         high_supervision = as.factor(as.integer(supervision_level > 0.8)),
         log_claim        = log(claim_amount)) %>%
  filter(claim_amount > 0)


# ---- 2) Train / test split (80/20) ----

freq_idx   <- sample(nrow(wc_freq_model), 0.8 * nrow(wc_freq_model))
freq_train <- wc_freq_model[ freq_idx, ]
freq_test  <- wc_freq_model[-freq_idx, ]

sev_idx    <- sample(nrow(wc_sev_model), 0.8 * nrow(wc_sev_model))
sev_train  <- wc_sev_model[ sev_idx, ]
sev_test   <- wc_sev_model[-sev_idx, ]

cat("Freq train:", nrow(freq_train), "| test:", nrow(freq_test), "\n")
cat("Sev  train:", nrow(sev_train),  "| test:", nrow(sev_test),  "\n")


# ---- 3) Severity tail diagnostics ----

cat("\n========== SEVERITY TAIL DIAGNOSTICS ==========\n")
cat(sprintf("Mean:     %s\n",   round(mean(sev_train$claim_amount), 0)))
cat(sprintf("Median:   %s\n",   round(median(sev_train$claim_amount), 0)))
cat(sprintf("SD:       %s\n",   round(sd(sev_train$claim_amount), 0)))
cat(sprintf("Skewness: %.3f\n", skewness(sev_train$claim_amount)))
cat(sprintf("Kurtosis: %.3f\n", kurtosis(sev_train$claim_amount)))
cat(sprintf("P95/Mean: %.2f\n", quantile(sev_train$claim_amount, 0.95) /
              mean(sev_train$claim_amount)))
cat(sprintf("P99/Mean: %.2f\n", quantile(sev_train$claim_amount, 0.99) /
              mean(sev_train$claim_amount)))

# mean excess plot — upward slope = heavy tail (Pareto-like)
claim_sorted <- sort(sev_train$claim_amount)
n_cl         <- length(claim_sorted)

me_vals <- sapply(seq_len(n_cl - 1), function(i) {
  v <- claim_sorted[claim_sorted > claim_sorted[i]]
  if (length(v) == 0) NA_real_ else mean(v - claim_sorted[i])
})

tibble(threshold = claim_sorted[seq_len(n_cl - 1)], mean_excess = me_vals) %>%
  filter(threshold <= quantile(sev_train$claim_amount, 0.95),
         !is.na(mean_excess)) %>%
  ggplot(aes(x = threshold, y = mean_excess)) +
  geom_line(colour = "grey50", linewidth = 0.4) +
  geom_smooth(method = "loess", se = FALSE, colour = "black", linewidth = 1) +
  scale_x_continuous(labels = comma) +
  scale_y_continuous(labels = comma) +
  labs(title    = "Mean Excess Plot — WC Claim Severity",
       subtitle = "Upward = heavy tail | Flat = exponential | Downward = thin tail",
       x = "Threshold", y = "Mean Excess Loss E[X-u | X>u]") +
  theme_minimal()

# scale to /1000 before fitting to avoid MLE overflow
claims_scaled <- sev_train$claim_amount / 1000

fit_ln   <- fitdist(claims_scaled, "lnorm")
fit_gam  <- fitdist(claims_scaled, "gamma",
                    start = list(shape = 1, rate = 1 / mean(claims_scaled)))
fit_weib <- fitdist(claims_scaled, "weibull",
                    start = list(shape = 1, scale = mean(claims_scaled)))
fit_par  <- tryCatch(
  fitdist(claims_scaled, "pareto",
          start = list(shape = 2, scale = median(claims_scaled)),
          lower = c(0.1, 0.01)),
  error = function(e) { message("Pareto failed: ", e$message); NULL }
)

bind_rows(
  tibble(Distribution = "LogNormal", AIC = AIC(fit_ln),   BIC = BIC(fit_ln)),
  tibble(Distribution = "Gamma",     AIC = AIC(fit_gam),  BIC = BIC(fit_gam)),
  tibble(Distribution = "Weibull",   AIC = AIC(fit_weib), BIC = BIC(fit_weib)),
  if (!is.null(fit_par))
    tibble(Distribution = "Pareto",          AIC = AIC(fit_par), BIC = BIC(fit_par))
  else
    tibble(Distribution = "Pareto (failed)", AIC = NA_real_,     BIC = NA_real_)
) %>%
  arrange(AIC) %>%
  { cat("\n--- Marginal AIC (claims/1000 — family comparison only) ---\n"); print(.) }


# ---- 4) GBM variable importance ----

freq_gbm <- gbm(
  (claim_count / exposure) ~ occupation + employment_type + solar_system +
    station_id + accident_history_flag + psych_stress_index +
    hours_per_week + safety_training_index + protective_gear_quality +
    experience_yrs + supervision_level + gravity_level + base_salary,
  data = wc_freq_model, distribution = "gaussian",
  n.trees = 500, interaction.depth = 4, shrinkage = 0.01,
  bag.fraction = 0.75, cv.folds = 5, verbose = FALSE
)
best_freq <- gbm.perf(freq_gbm, method = "cv", plot.it = FALSE)

sev_gbm <- gbm(
  claim_amount ~ occupation + employment_type + solar_system +
    station_id + accident_history_flag + psych_stress_index +
    hours_per_week + safety_training_index + protective_gear_quality +
    injury_type + injury_cause +
    experience_yrs + supervision_level + gravity_level + base_salary,
  data = wc_sev_model, distribution = "gaussian",
  n.trees = 500, interaction.depth = 4, shrinkage = 0.01,
  bag.fraction = 0.75, cv.folds = 5, verbose = FALSE
)
best_sev <- gbm.perf(sev_gbm, method = "cv", plot.it = FALSE)

plot_varimp <- function(vimp, title) {
  ggplot(vimp, aes(x = reorder(var, rel.inf), y = rel.inf)) +
    geom_col(fill = "grey30") +
    geom_text(aes(label = sprintf("%.1f%%", rel.inf)), hjust = -0.1, size = 3) +
    coord_flip() +
    labs(title = title, x = "Variable", y = "Relative Influence (%)") +
    theme_minimal()
}

plot_varimp(summary(freq_gbm, n.trees = best_freq, plotit = FALSE),
            "GBM Variable Importance — Frequency")
plot_varimp(summary(sev_gbm,  n.trees = best_sev,  plotit = FALSE),
            "GBM Variable Importance — Severity")


# ---- 5) Frequency models ----
# Retained: M1 (full benchmark), M2 (stress test), M4 (CQ pricing NB),
#           M6 (CQ+safety), M7 (CQ ZINB)
# Dropped:  M3 (station_id not CQ-deployable), M5 (ZINB non-CQ predictors)

nb_ctrl <- glm.control(maxit = 200, epsilon = 1e-8)

freq_m1 <- suppressWarnings(glm.nb(
  claim_count ~ occupation + employment_type + solar_system +
    accident_history_flag + psych_stress_index + hours_per_week +
    safety_training_index + protective_gear_quality +
    experience_yrs + supervision_level + gravity_level + base_salary +
    offset(log_exposure),
  data = freq_train, control = nb_ctrl
))

freq_m2 <- suppressWarnings(glm.nb(
  claim_count ~ occupation + accident_history_flag + psych_stress_index +
    safety_training_index + experience_yrs + supervision_level +
    offset(log_exposure),
  data = freq_train, control = nb_ctrl
))

freq_m4 <- glm.nb(
  claim_count ~ occupation + employment_type + solar_system +
    experience_yrs + offset(log_exposure),
  data = freq_train, control = nb_ctrl
)

freq_m6 <- suppressWarnings(glm.nb(
  claim_count ~ occupation + solar_system +
    experience_yrs + safety_training_index + offset(log_exposure),
  data = freq_train, control = nb_ctrl
))

freq_m7_zinb <- zeroinfl(
  claim_count ~ occupation + employment_type + solar_system +
    experience_yrs + offset(log_exposure) | occupation,
  data = freq_train, dist = "negbin"
)

# theta in 1-10 range = genuine NB overdispersion, not near-Poisson
cat("\nTheta estimates (1-10 = genuine NB overdispersion):\n")
for (nm in c("m1", "m2", "m4", "m6")) {
  cat(sprintf("  %s theta: %.2f\n", toupper(nm),
              get(paste0("freq_", nm))$theta))
}

# auto-select NB vs ZINB on CQ predictors
zinb_delta <- AIC(freq_m4) - AIC(freq_m7_zinb)
cat(sprintf("\nM4 NB: %.1f | M7 ZINB: %.1f | Delta: %.1f %s\n",
            AIC(freq_m4), AIC(freq_m7_zinb), zinb_delta,
            ifelse(zinb_delta > 2, "(ZINB wins)", "(NB wins)")))

if (zinb_delta > 2) {
  freq_final      <- freq_m7_zinb
  freq_final_name <- "M7: ZINB CQ-Deployable"
  freq_final_tag  <- "M7"
} else {
  freq_final      <- freq_m4
  freq_final_name <- "M4: CQ-Deployable NB"
  freq_final_tag  <- "M4"
}
cat("★ freq_final:", freq_final_name, "\n")


# ---- 6) Frequency evaluation ----

eval_freq <- function(model, test_data, model_name) {
  preds  <- predict(model, newdata = test_data, type = "response")
  actual <- test_data$claim_count
  tibble(
    Model        = model_name,
    AIC          = round(AIC(model), 1),
    RMSE         = round(sqrt(mean((preds - actual)^2)), 6),
    MAE          = round(mean(abs(preds - actual)), 6),
    TestDeviance = round(2 * sum(ifelse(actual == 0, 0,
                                        actual * log(actual / preds)) - (actual - preds)), 2),
    Deployable   = ifelse(grepl("M4|M7", model_name), "YES", "NO"),
    FINAL        = ifelse(grepl(freq_final_tag, model_name), "★", "")
  )
}

freq_results <- bind_rows(
  eval_freq(freq_m1,      freq_test, "M1: Full NB"),
  eval_freq(freq_m2,      freq_test, "M2: EDA NB  [stress test]"),
  eval_freq(freq_m4,      freq_test, "M4: CQ NB"),
  eval_freq(freq_m6,      freq_test, "M6: CQ+safety NB"),
  eval_freq(freq_m7_zinb, freq_test, "M7: CQ ZINB")
)

cat("\n╔══════════════════════════════════════════════╗\n")
cat("║      FREQUENCY MODEL COMPARISON             ║\n")
cat("╚══════════════════════════════════════════════╝\n")
print(freq_results, n = Inf, width = Inf)

anova(freq_m4, freq_m6, freq_m2, freq_m1, test = "Chisq")

# actual vs fitted by occupation - useful sanity check across key segments
freq_test %>%
  mutate(actual_rate = claim_count / exposure,
         fitted_m2   = predict(freq_m2,    ., "response") / exposure,
         fitted_m4   = predict(freq_m4,    ., "response") / exposure,
         fitted_fin  = predict(freq_final, ., "response") / exposure) %>%
  group_by(occupation) %>%
  summarise(Actual    = mean(actual_rate), M2_Stress = mean(fitted_m2),
            M4_CQ     = mean(fitted_m4),   Final     = mean(fitted_fin),
            .groups = "drop") %>%
  pivot_longer(-occupation, names_to = "model", values_to = "rate") %>%
  ggplot(aes(x = occupation, y = rate, fill = model)) +
  geom_col(position = "dodge") +
  scale_fill_manual(values = c("Actual"    = "black",   "M2_Stress" = "#555555",
                               "M4_CQ"    = "#999999", "Final"     = "#cccccc")) +
  labs(title = "Actual vs Fitted Claim Rate by Occupation — Frequency",
       x = "Occupation", y = "Mean claim rate") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))


# ---- 7) Severity models ----
# Retained: S1 (Gamma full benchmark), S3 (Gamma CQ), S4 (LN EDA stress),
#           S5 (LN CQ pricing), S6 (InvGaussian CQ)
# Dropped:  S2 (Gamma EDA — dominated by S4/S5 on all metrics),
#           S7/S8 Tweedie (package conflict, unreliable on this data)

sev_s1 <- glm(
  claim_amount ~ occupation + employment_type + solar_system +
    accident_history_flag + psych_stress_index + safety_training_index +
    protective_gear_quality + injury_type + injury_cause +
    experience_yrs + supervision_level + base_salary,
  data = sev_train, family = Gamma(link = "log")
)

sev_s3 <- glm(
  claim_amount ~ occupation + solar_system + employment_type + experience_yrs,
  data = sev_train, family = Gamma(link = "log")
)

sev_s4        <- lm(
  log_claim ~ occupation + solar_system + psych_stress_index +
    safety_training_index + injury_type + injury_cause + experience_yrs,
  data = sev_train
)
sev_s4_sigma2 <- var(residuals(sev_s4))

sev_s5        <- lm(
  log_claim ~ occupation + solar_system + employment_type + experience_yrs,
  data = sev_train
)
sev_s5_sigma2 <- var(residuals(sev_s5))

sev_s6 <- glm(
  claim_amount ~ occupation + solar_system + employment_type + experience_yrs,
  data = sev_train, family = inverse.gaussian(link = "log")
)

cat(sprintf("\nS5 LN CQ sigma^2: %.4f | retransform mult: %.4f\n",
            sev_s5_sigma2, exp(0.5 * sev_s5_sigma2)))


# ---- 8) Severity helper + bias check ----

get_sev_preds <- function(model, newdata, lognormal = FALSE, sigma2 = NULL) {
  if (lognormal) exp(predict(model, newdata = newdata) + 0.5 * sigma2)
  else           predict(model, newdata = newdata, type = "response")
}

sev_models <- list(
  list(name = "S1: Gamma Full",             mod = sev_s1, ln = FALSE, s2 = NULL,         cq = FALSE),
  list(name = "S3: Gamma CQ",               mod = sev_s3, ln = FALSE, s2 = NULL,         cq = TRUE),
  list(name = "S4: LogNormal EDA [stress]", mod = sev_s4, ln = TRUE,  s2 = sev_s4_sigma2, cq = FALSE),
  list(name = "S5: LogNormal CQ",           mod = sev_s5, ln = TRUE,  s2 = sev_s5_sigma2, cq = TRUE),
  list(name = "S6: InvGaussian CQ",         mod = sev_s6, ln = FALSE, s2 = NULL,         cq = TRUE)
)

mean_actual  <- mean(sev_test$claim_amount)
bias_results <- map_dfr(sev_models, function(x) {
  p <- get_sev_preds(x$mod, sev_test, x$ln, x$s2)
  tibble(Model      = x$name,
         Mean_Pred  = round(mean(p), 0),
         Mean_Act   = round(mean_actual, 0),
         Bias_Ratio = round(mean(p) / mean_actual, 4),
         Bias_Pct   = round((mean(p) / mean_actual - 1) * 100, 2))
})

cat("\n╔══════════════════════════════════════════════╗\n")
cat("║      PORTFOLIO BIAS CHECK                   ║\n")
cat("║  Target: Bias_Ratio = 1.0                   ║\n")
cat("╚══════════════════════════════════════════════╝\n")
print(bias_results, n = Inf, width = Inf)


# ---- 9) Severity evaluation ----

eval_sev <- function(entry, test_data) {
  preds  <- get_sev_preds(entry$mod, test_data, entry$ln, entry$s2)
  actual <- test_data$claim_amount
  r2 <- if (entry$ln) round(summary(entry$mod)$adj.r.squared, 4)
  else          round(1 - entry$mod$deviance / entry$mod$null.deviance, 4)
  tibble(
    Model         = entry$name,
    AIC           = round(AIC(entry$mod), 1),
    RMSE          = round(sqrt(mean((preds - actual)^2)), 1),
    MAE           = round(mean(abs(preds - actual)), 1),
    MAPE_pct      = round(mean(abs((preds - actual) / actual)) * 100, 2),
    R2            = r2,
    R2_type       = if (entry$ln) "Adj-R2(log)" else "Pseudo-R2",
    Deployable_CQ = if (entry$cq) "YES" else "NO"
  )
}

sev_results <- map_dfr(sev_models, eval_sev, test_data = sev_test)

cat("\n╔══════════════════════════════════════════════╗\n")
cat("║      SEVERITY MODEL COMPARISON              ║\n")
cat("╚══════════════════════════════════════════════╝\n")
print(sev_results, n = Inf, width = Inf)


# ---- 10) Auto-select sev_final ----
# rule: lowest AIC among CQ-deployable, confirm bias ratio within 5% of 1.0

cq_ranked <- sev_results %>%
  filter(Deployable_CQ == "YES") %>%
  arrange(AIC)

cat("\n--- CQ-deployable models ranked by AIC ---\n")
print(cq_ranked, width = Inf)

best_name  <- cq_ranked$Model[1]
best_tag   <- str_extract(best_name, "^S[0-9]+")
best_entry <- sev_models[[which(map_chr(sev_models, "name") == best_name)]]

sev_final              <- best_entry$mod
sev_final_name         <- best_name
sev_final_is_lognormal <- best_entry$ln
sev_final_sigma2       <- if (best_entry$ln) best_entry$s2 else NA_real_

final_bias <- bias_results %>% filter(grepl(best_tag, Model)) %>% pull(Bias_Ratio)

cat("\n★ sev_final:", sev_final_name, "\n")
cat("  Bias ratio:", final_bias, "\n")

if (abs(final_bias - 1) > 0.05)
  cat("  WARNING: bias ratio outside ±5% — review before pricing.\n")

if (sev_final_is_lognormal)
  cat(sprintf("  sigma^2: %.4f | multiplier: %.4f\n",
              sev_final_sigma2, exp(0.5 * sev_final_sigma2)))


# ---- 11) Severity diagnostics — actual vs fitted by occupation ----

sev_test_diag <- sev_test %>%
  mutate(
    S4_LN_Stress  = get_sev_preds(sev_s4, cur_data(), TRUE, sev_s4_sigma2),
    S5_LN_CQ      = get_sev_preds(sev_s5, cur_data(), TRUE, sev_s5_sigma2),
    S6_InvGaus_CQ = get_sev_preds(sev_s6, cur_data())
  )

sev_test_diag %>%
  group_by(occupation) %>%
  summarise(Actual        = mean(claim_amount),
            S4_LN_Stress  = mean(S4_LN_Stress),
            S5_LN_CQ      = mean(S5_LN_CQ),
            S6_InvGaus_CQ = mean(S6_InvGaus_CQ),
            .groups = "drop") %>%
  pivot_longer(-occupation, names_to = "model", values_to = "amount") %>%
  ggplot(aes(x = occupation, y = amount, fill = model)) +
  geom_col(position = "dodge") +
  scale_fill_manual(values = c("Actual"        = "black",
                               "S4_LN_Stress"  = "#444444",
                               "S5_LN_CQ"      = "#888888",
                               "S6_InvGaus_CQ" = "#bbbbbb")) +
  scale_y_continuous(labels = comma) +
  labs(title = "Actual vs Fitted Claim Amount by Occupation — Severity",
       x = "Occupation", y = "Mean claim amount") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))


# ---- 11b) QQ plots + residuals — final models ----

# frequency: M4 NB pearson residuals
freq_resid_df <- tibble(
  fitted   = predict(freq_final, freq_test, type = "response"),
  actual   = freq_test$claim_count,
  residual = actual - fitted,
  pearson  = residual / sqrt(fitted)
)

ggplot(freq_resid_df, aes(x = fitted, y = pearson)) +
  geom_point(alpha = 0.2, size = 0.8, colour = "grey30") +
  geom_hline(yintercept = 0, linetype = "dashed") +
  geom_smooth(method = "loess", se = FALSE, colour = "black", linewidth = 0.8) +
  labs(title = "Frequency M4 NB — Pearson Residuals vs Fitted",
       x = "Fitted (expected claims)", y = "Pearson Residual") +
  theme_minimal()

ggplot(freq_resid_df, aes(sample = pearson)) +
  stat_qq(alpha = 0.3, size = 0.8, colour = "grey30") +
  stat_qq_line(colour = "black", linewidth = 0.8) +
  labs(title = "Frequency M4 NB — QQ Plot (Pearson Residuals)",
       x = "Theoretical Quantiles", y = "Sample Quantiles") +
  theme_minimal()

# severity: S5 LogNormal residuals in log space
sev_resid_df <- tibble(
  fitted_log = predict(sev_final, sev_test),
  actual_log = sev_test$log_claim,
  residual   = actual_log - fitted_log
)

ggplot(sev_resid_df, aes(x = fitted_log, y = residual)) +
  geom_point(alpha = 0.4, size = 0.8, colour = "grey30") +
  geom_hline(yintercept = 0, linetype = "dashed") +
  geom_smooth(method = "loess", se = FALSE, colour = "black", linewidth = 0.8) +
  labs(title = "Severity S5 LogNormal — Residuals vs Fitted (log space)",
       x = "Fitted log(claim amount)", y = "Residual") +
  theme_minimal()

ggplot(sev_resid_df, aes(sample = residual)) +
  stat_qq(alpha = 0.4, size = 0.8, colour = "grey30") +
  stat_qq_line(colour = "black", linewidth = 0.8) +
  labs(title = "Severity S5 LogNormal — QQ Plot (Log-space Residuals)",
       x = "Theoretical Quantiles", y = "Sample Quantiles") +
  theme_minimal()

# scale-location to check heteroskedasticity
ggplot(sev_resid_df, aes(x = fitted_log, y = sqrt(abs(residual)))) +
  geom_point(alpha = 0.4, size = 0.8, colour = "grey30") +
  geom_smooth(method = "loess", se = FALSE, colour = "black", linewidth = 0.8) +
  labs(title = "Severity S5 LogNormal — Scale-Location",
       x = "Fitted log(claim amount)", y = "sqrt(|Residual|)") +
  theme_minimal()


# ---- 12) Final summary ----

cat("\n╔══════════════════════════════════════════════╗\n")
cat("║         FINAL MODEL SELECTION               ║\n")
cat("╚══════════════════════════════════════════════╝\n")
cat("PRICING  Frequency:", freq_final_name, "\n")
cat("PRICING  Severity :", sev_final_name, "\n")
cat("STRESS   Frequency: freq_m2  (EDA NB)\n")
cat("STRESS   Severity : sev_s4   (LogNormal EDA)\n")
cat("\n--- Pass to 03_aggregate_simulation.R ---\n")
cat("freq_final, freq_m2\n")
cat("sev_final, sev_final_is_lognormal, sev_final_sigma2\n")
cat("sev_s4, sev_s4_sigma2\n")
cat("get_sev_preds()\n")

if (!is.na(sev_final_sigma2))
  cat(sprintf("sev_final_sigma2: %.4f | mult: %.4f\n",
              sev_final_sigma2, exp(0.5 * sev_final_sigma2)))
cat(sprintf("sev_s4_sigma2:    %.4f\n", sev_s4_sigma2))


# ---- 13) Export — rename finals for 03_aggregate_simulation.R ----

freq_pricing <- freq_final      # M4: CQ NB
freq_stress  <- freq_m2         # M2: EDA NB

sev_pricing              <- sev_final
sev_pricing_is_lognormal <- sev_final_is_lognormal
sev_pricing_sigma2       <- sev_final_sigma2

sev_stress        <- sev_s4
sev_stress_sigma2 <- sev_s4_sigma2

cat("\n--- Objects ready for 03_aggregate_simulation.R ---\n")
cat("freq_pricing  :", freq_final_name, "\n")
cat("freq_stress   : 