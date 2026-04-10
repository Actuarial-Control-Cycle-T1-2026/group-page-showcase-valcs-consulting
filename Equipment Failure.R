
# 0. Setup and Data Import ----
library(readxl)
library(dplyr)
library(stringr)
library(tidyr)
library(purrr)
library(rlang)
library(writexl)
library(MASS)    # For standard Negative Binomial
library(pscl)    # For Zero-Inflated models
library(lmtest)  # For Likelihood Ratio Tests

# Define file path
file_path <- "/Users/liya/Library/Mobile Documents/com~apple~CloudDocs/ACTL4001 Assignment/SOA_2026_Case_Study_Materials/Equipment Failure/equipment_failure_CLEANED_FINALv2.xlsx"

# Import the cleaned frequency dataset
frequency_v1 <- read_excel(file_path, sheet = "Frequency_final")
severity_v1 <- read_excel(file_path, sheet = "Severity_final")

# Modify equipment age to be categorical
frequency_v1$equipment_age_band <- cut(
  frequency_v1$equipment_age,
  breaks = c(-Inf, 4, 9, 14, 19, Inf),
  labels = c("<5", "5-9", "10-14", "15-19", "20+"),
  right = TRUE
)

severity_v1$equipment_age_band <- cut(
  severity_v1$equipment_age,
  breaks = c(-Inf, 4, 9, 14, 19, Inf),
  labels = c("<5", "5-9", "10-14", "15-19", "20+"),
  right = TRUE
)

equipment_freq_clean <- frequency_v1[,-4]

severity <- severity_v1[,-4]

##############################################################################
# MODELING FREQUENCY: Standard NB vs. Full ZINB vs. Reduced ZINB
# 1. Train/Test Split & Preparation ----
# Hold out 20% of the data to test out-of-sample accuracy
set.seed(2026) 
train_indices <- sample(seq_len(nrow(equipment_freq_clean)), size = 0.8 * nrow(equipment_freq_clean))

train_data <- equipment_freq_clean[train_indices, ]
test_data  <- equipment_freq_clean[-train_indices, ]

# Apply safe exposure offsets to prevent log(0) errors
train_data$exposure_safe <- ifelse(train_data$exposure == 0, 1e-6, train_data$exposure)
test_data$exposure_safe  <- ifelse(test_data$exposure == 0, 1e-6, test_data$exposure)


# 2. Model Fitting ----
cat("Fitting Model 1: Standard Negative Binomial...\n")
model_nb <- glm.nb(
  claim_count ~ equipment_age_band + maintenance_int + usage_int + equipment_type + solar_system + offset(log(exposure_safe)), 
  data = train_data
)

cat("Fitting Model 2: Full Zero-Inflated Negative Binomial (This may take a minute)...\n")
model_zinb_full <- zeroinfl(
  claim_count ~ equipment_age_band + maintenance_int + usage_int + equipment_type + solar_system + offset(log(exposure_safe)) | 
    equipment_age_band + maintenance_int + usage_int + equipment_type + solar_system + offset(log(exposure_safe)), 
  data = train_data, 
  dist = "negbin"
)

# (Usage, Solar System, Offset) to ensure model convergence and minimize AIC.
cat("Fitting Model 3: Reduced Zero-Inflated Negative Binomial...\n")
model_zinb_reduced <- zeroinfl(
  claim_count ~ equipment_age_band + maintenance_int + usage_int + equipment_type + solar_system + offset(log(exposure_safe)) | 
    equipment_age_band + equipment_type, 
  data = train_data, 
  dist = "negbin"
)

cat("Fitting Model 4: Reduced Negative Binomial Model removing exposure")
model_nb_red <- glm.nb(
  claim_count ~ equipment_age_band + maintenance_int + usage_int + equipment_type + solar_system , 
  data = train_data
)

# 4. Out-of-Sample Testing & Metric Calculation ----
cat("\nScoring Models on 20% Holdout Testing Data...\n")

# Predict expected continuous counts on unseen test data
test_data$pred_nb_cont        <- predict(model_nb, newdata = test_data, type = "response")
test_data$pred_zinb_full_cont <- predict(model_zinb_full, newdata = test_data, type = "response")
test_data$pred_zinb_red_cont  <- predict(model_zinb_reduced, newdata = test_data, type = "response")
test_data$pred_nb_red_cont  <- predict(model_nb_red, newdata = test_data, type = "response")


# Calculate Exact Match Accuracy (ignoring NAs)
acc_nb        <- mean(round(test_data$pred_nb_cont) == test_data$claim_count, na.rm = TRUE)
acc_zinb_full <- mean(round(test_data$pred_zinb_full_cont) == test_data$claim_count, na.rm = TRUE)
acc_zinb_red  <- mean(round(test_data$pred_zinb_red_cont) == test_data$claim_count, na.rm = TRUE)

# Calculate Root Mean Squared Error (RMSE) (ignoring NAs)
rmse_nb        <- sqrt(mean((test_data$pred_nb_cont - test_data$claim_count)^2, na.rm = TRUE))
rmse_zinb_full <- sqrt(mean((test_data$pred_zinb_full_cont - test_data$claim_count)^2, na.rm = TRUE))
rmse_zinb_red  <- sqrt(mean((test_data$pred_zinb_red_cont - test_data$claim_count)^2, na.rm = TRUE))
rmse_nb_red  <- sqrt(mean((test_data$pred_nb_red_cont - test_data$claim_count)^2, na.rm = TRUE))


# Calculate AIC & BIC
aic_nb        <- AIC(model_nb)
aic_zinb_full <- AIC(model_zinb_full)
aic_zinb_red  <- AIC(model_zinb_reduced)
aic_nb_red  <- AIC(model_nb_red)


bic_nb        <- BIC(model_nb)
bic_zinb_full <- BIC(model_zinb_full)
bic_zinb_red  <- BIC(model_zinb_reduced)
bic_nb_red  <- BIC(model_nb_red)

##############################################################################
# MODELING SEVERITY: Full Log Normal vs. Reduced Log Normal
# 1. Train/Test Split & Preparation ----
# Hold out 20% of the data to test out-of-sample accuracy
set.seed(2026) 
train_indices_sev <- sample(seq_len(nrow(severity)), size = 0.8 * nrow(severity))

train_severity <- severity[train_indices_sev, ]
test_severity  <- severity[-train_indices_sev, ]

# 2. Model Fitting ----
# Confirm EDA finding - the behaviour follows log-normal
qqnorm(log(train_severity$claim_amount),
       main = "QQ Plot of Log Claim Amounts")

qqline(log(train_severity$claim_amount), col = "red", lwd = 2)


qqnorm(
  log(train_severity$claim_amount),
  col = "steelblue",
  pch = 20
)

qqline(
  log(train_severity$claim_amount),
  col = "red",
  lwd = 2
)

# Full lognormal model
full_log <- glm(
  claim_amount ~ equipment_type + equipment_age_band + maintenance_int + usage_int + exposure + solar_system,
  data = train_severity,
  family = gaussian(link = "log")
)

summary(full_log)

# Selection using step model
step_model <- step(full_log, direction = "both")


# Reduced lognormal model - removing exposure, equipment age, maintenance_int
reduced_log <- glm(
  claim_amount ~equipment_type + usage_int + solar_system,
  data = train_severity,
  family = gaussian(link = "log")
)

summary(reduced_log)
# 3. Out-of-Sample Testing & Metric Calculation ----
# Predict expected continuous counts on unseen test data
test_severity$pred_full <- predict(full_log, newdata = test_severity, type = "response")
test_severity$pred_red <- predict(reduced_log, newdata = test_severity, type = "response")

# Calculate Root Mean Squared Error (RMSE) (ignoring NAs)
rmse_full        <- sqrt(mean((test_severity$pred_full - test_severity$claim_amount)^2, na.rm = TRUE))
rmse_red <- sqrt(mean((test_severity$pred_red - test_severity$claim_amount)^2, na.rm = TRUE))

# Calculate AIC & BIC
aic_full        <- AIC(full_log)
aic_red <- AIC(reduced_log)

bic_full       <- BIC(full_log)
bic_red <- BIC(reduced_log)

# Table
model_comparison <- data.frame(
  Model = c("Full Model", "Reduced Model"),
  RMSE = c(rmse_full, rmse_red),
  AIC = c(aic_full, aic_red),
  BIC = c(bic_full, bic_red)
)

model_comparison

##############################################################################
# Predicting: Data Import ----
# Define file path
file_path2 <- "/Users/liya/Library/Mobile Documents/com~apple~CloudDocs/ACTL4001 Assignment/SOA_2026_Case_Study_Materials/Equipment Failure/Equipment dataset.xlsx"
# Import the new dataset
Helionis_original <- read_excel(file_path2, sheet = "Helionis Cluster")
Bayesia_original <- read_excel(file_path2, sheet = "Bayesia")
oryn_original <- read_excel(file_path2, sheet = "Oryn Delta")

# Expanding the rows & simulate exposure as ~U(0,1) 
Helionis <- Helionis_original %>%
  uncount(number) %>%
  mutate(exposure = runif(n(), 0, 1))

Bayesia <- Bayesia_original %>%
  uncount(number) %>%
  mutate(exposure = runif(n(), 0, 1))

oryn <- oryn_original %>%
  uncount(number) %>%
  mutate(exposure = runif(n(), 0, 1))


# Apply safe exposure offsets to prevent log(0) errors
Helionis$exposure_safe <- ifelse(Helionis$exposure == 0, 1e-6, Helionis$exposure)
Bayesia$exposure_safe <- ifelse(Bayesia$exposure == 0, 1e-6, Bayesia$exposure)
oryn$exposure_safe <- ifelse(oryn$exposure == 0, 1e-6, oryn$exposure)

# Import risk factor scaling
rate <- read_excel(file_path2, sheet = "Rate")
Bayesia <- Bayesia %>%
  left_join(rate, by = "equipment_type")
oryn <- oryn %>%
  left_join(rate, by = "equipment_type")

## Simulation for 1 year ----
set.seed(123)

n_sim <- 10000
portfolio_loss <- numeric(n_sim)

# precompute sigma on log scale (severity)
resid_log <- residuals(reduced_log, type = "response") / predict(reduced_log, type = "response")
log_resid <- log(1 + resid_log)
sigma_sev <- sd(log_resid, na.rm = TRUE)


for(i in 1:n_sim){
  
  total_loss <- 0
  
  # ======================================================
  # HELIONIS
  # ======================================================
  N_h <- rnbinom(
    n = nrow(Helionis),
    size = model_nb$theta,
    mu = Helionis$claim_count
  )
  
  if(sum(N_h) > 0){
    mean_link_h <- predict(reduced_log, newdata = Helionis, type = "link")
    mean_rep_h <- rep(mean_link_h, times = N_h)
    
    sev_h <- rlnorm(
      n = length(mean_rep_h),
      meanlog = mean_rep_h,
      sdlog = sigma_sev
    )
    
    total_loss <- total_loss + sum(sev_h)
  }
  
  
  # ======================================================
  # BAYESIA
  # ======================================================
  N_b <- rnbinom(
    n = nrow(Bayesia),
    size = model_nb$theta,
    mu = Bayesia$freq_scaled
  )
  
  if(sum(N_b) > 0){
    mean_link_b <- predict(reduced_log, newdata = Bayesia, type = "link")
    mean_rep_b <- rep(mean_link_b, times = N_b)
    
    sev_b <- rlnorm(
      n = length(mean_rep_b),
      meanlog = mean_rep_b,
      sdlog = sigma_sev
    )
    
    # apply risk factor per claim (expanded to match sev_b)
    risk_rep_b <- rep(Bayesia$`Bayesia System`, times = N_b)
    sev_b <- sev_b * risk_rep_b
    
    total_loss <- total_loss + sum(sev_b)
  }
  
  
  # ======================================================
  # ORYN DELTA
  # ======================================================
  N_o <- rnbinom(
    n = nrow(oryn),
    size = model_nb$theta,
    mu = oryn$freq_scaled
  )
  
  if(sum(N_o) > 0){
    mean_link_o <- predict(reduced_log, newdata = oryn, type = "link")
    mean_rep_o <- rep(mean_link_o, times = N_o)
    
    sev_o <- rlnorm(
      n = length(mean_rep_o),
      meanlog = mean_rep_o,
      sdlog = sigma_sev
    )
    
    # apply risk factor per claim (expanded to match sev_o)
    risk_rep_o <- rep(oryn$`Oryn Delta`, times = N_o)
    sev_o <- sev_o * risk_rep_o
    
    total_loss <- total_loss + sum(sev_o)
  }
  
  
  # store result
  portfolio_loss[i] <- total_loss
}

summary(portfolio_loss)

portfolio_loss_m <- portfolio_loss / 1e6


results_df <- data.frame(portfolio_loss = portfolio_loss_m)
write_xlsx(results_df, "portfolio_loss_EF.xlsx")

## Loss distribution plot ----
EV <- mean(results_df$portfolio_loss)
VAR <- var(results_df$portfolio_loss)
sd <- sqrt(VAR)
VaR_995 <- quantile(results_df$portfolio_loss, 0.995)
ggplot(results_df, aes(x = portfolio_loss)) +
  geom_histogram(
    aes(y = ..density..),
    bins = 50,
    fill = "#9DC3E6",
    color = "black",
    linewidth = 0.5
  ) +
  geom_density(
    color = "black",
    linewidth = 0.8
  ) +
  labs(
    title = "Simulated Aggregate Loss Distribution – Equipment Failure",
    x = "Aggregate Loss ($m)",
    y = "Density"
  ) +
  annotate(
    "text",
    x = Inf,
    y = Inf,
    label = paste0(
      "Expected Loss = ", round(EV, 2), "m\n",
      "VaR (99.5%) = ", round(VaR_995, 2), "m\n",
      "Standard Deviation = ", round(sd, 2), "m\n"
    ),
    hjust = 1.1,
    vjust = 1.5,
    size = 4
  ) +
  theme_minimal(base_size = 12)