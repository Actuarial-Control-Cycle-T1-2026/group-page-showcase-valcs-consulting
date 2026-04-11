##############################################################################
# Cosmic Quarry Mining - Cargo Loss Severity Modelling
##############################################################################

# ---------------------------------------------------------------------------
# Load packages
# ---------------------------------------------------------------------------
library(tidyverse)
library(scales)
library(mgcv)
library(Metrics)
library(xgboost)
library(DHARMa)
library(lspline)
library(tibble)

# ---------------------------------------------------------------------------
# Version control
# ---------------------------------------------------------------------------
# v1 | Sarina Truong | Initial severity modelling
# v2 | Cleaned       | Standardised naming, structure and flow

##############################################################################
# 1. Prepare Data
##############################################################################

set.seed(123)

severity_data <- cargo_severity %>%
  mutate(
    cargo_type = as.factor(cargo_type),
    container_type = as.factor(container_type),
    route_risk = as.factor(route_risk),
    claim_amount = as.numeric(claim_amount),
    cargo_value = as.numeric(cargo_value),
    weight = as.numeric(weight),
    distance = as.numeric(distance),
    transit_duration = as.numeric(transit_duration),
    pilot_experience = as.numeric(pilot_experience),
    vessel_age = as.numeric(vessel_age),
    solar_radiation = as.numeric(solar_radiation),
    debris_density = as.numeric(debris_density)
  ) %>%
  filter(
    is.finite(claim_amount),
    claim_amount > 0,
    is.finite(cargo_value),
    cargo_value > 0
  ) %>%
  mutate(
    loss_ratio = claim_amount / cargo_value,
    loss_ratio = pmin(pmax(loss_ratio, 0), 1),
    log_claim = log1p(claim_amount),
    log_cargo_value = log1p(cargo_value),
    log_weight = log1p(weight),
    logit_lr = qlogis(pmin(pmax(loss_ratio, 1e-6), 1 - 1e-6))
  )

##############################################################################
# 2. Train Test Split
##############################################################################

n <- nrow(severity_data)

test_index <- sample.int(n, floor(0.2 * n))

severity_train <- severity_data[-test_index, ]
severity_test  <- severity_data[test_index, ]

##############################################################################
# 3. Gamma GLM
##############################################################################

gamma_model <- glm(
  claim_amount ~ 
    log(cargo_value) +
    weight +
    cargo_type +
    route_risk +
    container_type +
    solar_radiation +
    debris_density +
    transit_duration,
  
  data = severity_train,
  family = Gamma(link = "log")
)

summary(gamma_model)

##############################################################################
# Residual Diagnostics
##############################################################################

qqnorm(residuals(gamma_model, type = "deviance"))
qqline(residuals(gamma_model, type = "deviance"))

plot(
  fitted(gamma_model),
  residuals(gamma_model, type = "pearson")
)

abline(h = 0, col = "red")

##############################################################################
# 4. Lognormal Model
##############################################################################

lognormal_model <- glm(
  log(claim_amount) ~
    log(cargo_value) +
    weight +
    cargo_type +
    route_risk +
    container_type +
    solar_radiation +
    debris_density +
    transit_duration,
  
  data = severity_train,
  family = gaussian()
)

summary(lognormal_model)

##############################################################################
# Predictions
##############################################################################

severity_test$log_pred <- predict(
  lognormal_model,
  newdata = severity_test
)

severity_test$predicted_claim <- exp(
  severity_test$log_pred
)

##############################################################################
# Performance Metrics
##############################################################################

rmse <- sqrt(
  mean(
    (severity_test$claim_amount - severity_test$predicted_claim)^2
  )
)

mae <- mean(
  abs(severity_test$claim_amount - severity_test$predicted_claim)
)

mape <- mean(
  abs(
    (severity_test$claim_amount - severity_test$predicted_claim) /
      severity_test$claim_amount
  )
) * 100

log_rmse <- sqrt(
  mean(
    (log(severity_test$claim_amount) -
       log(severity_test$predicted_claim))^2
  )
)

##############################################################################
# 5. Reduced Lognormal Model
##############################################################################

lognormal_reduced <- glm(
  log(claim_amount) ~
    log(cargo_value) +
    route_risk +
    solar_radiation +
    debris_density,
  
  data = severity_train,
  family = gaussian()
)

summary(lognormal_reduced)

##############################################################################
# 6. Split by Cargo Type
##############################################################################

severity_gold <- severity_data %>%
  filter(cargo_type == "gold")

severity_platinum <- severity_data %>%
  filter(cargo_type == "platinum")

severity_other <- severity_data %>%
  filter(!cargo_type %in% c("gold", "platinum"))

##############################################################################
# Helper Train Test Split
##############################################################################

split_80_20 <- function(data, seed = 123) {
  
  set.seed(seed)
  
  n <- nrow(data)
  
  test_index <- sample.int(n, floor(0.2 * n))
  
  list(
    train = data[-test_index, ],
    test = data[test_index, ]
  )
}

gold_split <- split_80_20(severity_gold)
plat_split <- split_80_20(severity_platinum)
other_split <- split_80_20(severity_other)

gold_train <- gold_split$train
gold_test  <- gold_split$test

plat_train <- plat_split$train
plat_test  <- plat_split$test

other_train <- other_split$train
other_test  <- other_split$test

##############################################################################
# 7. Lognormal Models by Cargo Type
##############################################################################

lognormal_formula <- log(claim_amount) ~
  cargo_value +
  weight +
  route_risk +
  container_type +
  solar_radiation +
  debris_density +
  pilot_experience +
  vessel_age +
  distance +
  transit_duration

lognormal_gold <- lm(lognormal_formula, data = gold_train)
lognormal_platinum <- lm(lognormal_formula, data = plat_train)
lognormal_other <- lm(lognormal_formula, data = other_train)

##############################################################################
# 8. Reduced Models
##############################################################################

lognormal_formula_reduced <- log(claim_amount) ~
  log(cargo_value) +
  route_risk +
  solar_radiation +
  debris_density

lognormal_gold_reduced <- lm(
  lognormal_formula_reduced,
  data = gold_train
)

lognormal_platinum_reduced <- lm(
  lognormal_formula_reduced,
  data = plat_train
)

lognormal_other_reduced <- lm(
  lognormal_formula_reduced,
  data = other_train
)

##############################################################################
# 9. Prediction Function
##############################################################################

predict_lognormal <- function(model, newdata) {
  
  mu_hat <- predict(model, newdata)
  
  sigma2 <- summary(model)$sigma^2
  
  exp(mu_hat + 0.5 * sigma2)
}

##############################################################################
# 10. Predictions
##############################################################################

gold_pred <- predict_lognormal(
  lognormal_gold,
  gold_test
)

plat_pred <- predict_lognormal(
  lognormal_platinum,
  plat_test
)

other_pred <- predict_lognormal(
  lognormal_other,
  other_test
)

##############################################################################
# 11. Model Comparison Table
##############################################################################

model_summary <- tibble(
  
  dataset = c("Gold", "Platinum", "Other"),
  
  AIC = c(
    AIC(lognormal_gold),
    AIC(lognormal_platinum),
    AIC(lognormal_other)
  ),
  
  RMSE = c(
    sqrt(mean((gold_test$claim_amount - gold_pred)^2)),
    sqrt(mean((plat_test$claim_amount - plat_pred)^2)),
    sqrt(mean((other_test$claim_amount - other_pred)^2))
  ),
  
  MAE = c(
    mean(abs(gold_test$claim_amount - gold_pred)),
    mean(abs(plat_test$claim_amount - plat_pred)),
    mean(abs(other_test$claim_amount - other_pred))
  )
)

model_summary
