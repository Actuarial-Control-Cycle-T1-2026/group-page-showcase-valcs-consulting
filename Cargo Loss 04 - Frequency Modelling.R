##############################################################################
# Cosmic Quarry Mining - Cargo Loss Frequency Modelling
##############################################################################

# ---------------------------------------------------------------------------
# Load packages
# ---------------------------------------------------------------------------
library(tidyverse)
library(MASS)
library(pscl)

# ---------------------------------------------------------------------------
# Version control
# ---------------------------------------------------------------------------
# v1 | Sarina Truong | Initial frequency modelling
# v2 | Cleaned       | Standardised naming, structure and flow

##############################################################################
# 1. Prepare Data
##############################################################################

frequency_data <- cargo_frequency %>%
  mutate(
    cargo_type = as.factor(cargo_type),
    container_type = as.factor(container_type),
    route_risk = as.factor(route_risk),
    claim_count = as.numeric(claim_count),
    cargo_value = as.numeric(cargo_value),
    weight = as.numeric(weight),
    distance = as.numeric(distance),
    transit_duration = as.numeric(transit_duration),
    pilot_experience = as.numeric(pilot_experience),
    vessel_age = as.numeric(vessel_age),
    solar_radiation = as.numeric(solar_radiation),
    debris_density = as.numeric(debris_density)
  )

##############################################################################
# 2. Train Test Split
##############################################################################

set.seed(123)

n <- nrow(frequency_data)

test_index <- sample.int(n, floor(0.2 * n))

frequency_train <- frequency_data[-test_index, ]
frequency_test  <- frequency_data[test_index, ]

##############################################################################
# 3. Poisson Model
##############################################################################

# Check mean vs variance
mean(frequency_train$claim_count, na.rm = TRUE)
var(frequency_train$claim_count, na.rm = TRUE)

poisson_model <- glm(
  claim_count ~ 
    cargo_type +
    container_type +
    route_risk +
    cargo_value +
    weight +
    distance +
    transit_duration +
    pilot_experience +
    vessel_age +
    solar_radiation +
    debris_density +
    offset(log(exposure)),
  
  family = poisson(link = "log"),
  data = frequency_train
)

# Dispersion check
sum(residuals(poisson_model, type = "pearson")^2) /
  df.residual(poisson_model)

##############################################################################
# 4. Negative Binomial Model
##############################################################################

nb_model <- glm.nb(
  claim_count ~ 
    log(cargo_value) +
    cargo_type +
    weight +
    route_risk +
    distance +
    transit_duration +
    solar_radiation +
    container_type +
    debris_density +
    pilot_experience +
    vessel_age +
    offset(log(exposure)),
  
  data = frequency_train
)

summary(nb_model)

##############################################################################
# 5. Reduced Negative Binomial Model
##############################################################################

nb_reduced_model <- glm.nb(
  claim_count ~
    route_risk +
    container_type +
    debris_density +
    pilot_experience +
    offset(log(exposure)),
  
  data = frequency_train
)

summary(nb_reduced_model)

##############################################################################
# 6. Model Comparison
##############################################################################

model_comparison <- tibble(
  
  Model = c(
    "Poisson",
    "Negative Binomial",
    "Negative Binomial Reduced"
  ),
  
  AIC = c(
    AIC(poisson_model),
    AIC(nb_model),
    AIC(nb_reduced_model)
  ),
  
  BIC = c(
    BIC(poisson_model),
    BIC(nb_model),
    BIC(nb_reduced_model)
  )
)

model_comparison

##############################################################################
# 7. Zero Inflated Negative Binomial
##############################################################################

zinb_full <- zeroinfl(
  claim_count ~
    log(cargo_value) +
    cargo_type +
    weight +
    route_risk +
    distance +
    transit_duration +
    solar_radiation +
    container_type +
    debris_density +
    pilot_experience +
    vessel_age +
    offset(log(exposure))
  |
    log(cargo_value) +
    cargo_type +
    weight +
    route_risk +
    distance +
    transit_duration +
    solar_radiation +
    container_type +
    debris_density +
    pilot_experience +
    vessel_age,
  
  data = frequency_train,
  dist = "negbin"
)

##############################################################################
# Reduced ZINB
##############################################################################

zinb_reduced <- zeroinfl(
  claim_count ~
    weight +
    route_risk +
    distance +
    pilot_experience +
    solar_radiation +
    offset(log(exposure))
  |
    vessel_age +
    route_risk +
    solar_radiation +
    debris_density +
    pilot_experience,
  
  data = frequency_train,
  dist = "negbin"
)

##############################################################################
# Compare ZINB models
##############################################################################

AIC(zinb_full, zinb_reduced)

##############################################################################
# Residual diagnostics
##############################################################################

plot(
  fitted(zinb_full),
  residuals(zinb_full, type = "pearson"),
  xlab = "Fitted Values",
  ylab = "Pearson Residuals"
)

abline(h = 0, col = "red")

##############################################################################
# 8. Variable Selection Loop
##############################################################################

predictors <- c(
  "log(cargo_value)",
  "cargo_type",
  "weight",
  "route_risk",
  "distance",
  "transit_duration",
  "solar_radiation",
  "container_type",
  "debris_density",
  "pilot_experience",
  "vessel_age"
)

full_formula <- as.formula(
  paste(
    "claim_count ~",
    paste(predictors, collapse = " + "),
    "+ offset(log(exposure)) |",
    paste(predictors, collapse = " + ")
  )
)

fit_full <- zeroinfl(
  full_formula,
  data = frequency_train,
  dist = "negbin"
)

results <- tibble(
  Model = "Full Model",
  LogLik = logLik(fit_full),
  AIC = AIC(fit_full)
)

##############################################################################
# Variable selection loop
##############################################################################

for (p in predictors) {
  
  # Remove from count
  form_count <- as.formula(
    paste(
      "claim_count ~",
      paste(predictors[predictors != p], collapse = " + "),
      "+ offset(log(exposure)) |",
      paste(predictors, collapse = " + ")
    )
  )
  
  fit_count <- zeroinfl(
    form_count,
    data = frequency_train,
    dist = "negbin"
  )
  
  results <- bind_rows(
    results,
    tibble(
      Model = paste("Drop", p, "Count"),
      LogLik = logLik(fit_count),
      AIC = AIC(fit_count)
    )
  )
  
}

##############################################################################
# Rank models
##############################################################################

frequency_model_summary <- results %>%
  arrange(AIC) %>%
  mutate(
    Likelihood_Rank = rank(-LogLik)
  )

frequency_model_summary

##############################################################################
# 9. Test Set Performance
##############################################################################

test_predictions <- predict(
  zinb_reduced,
  newdata = frequency_test,
  type = "response"
)

frequency_mse <- mean(
  (frequency_test$claim_count - test_predictions)^2
)

frequency_mse

##############################################################################
# 10. Decile Testing
##############################################################################

frequency_results <- tibble(
  actual = frequency_test$claim_count,
  predicted = test_predictions,
  cargo_value = frequency_test$cargo_value
)

frequency_results %>%
  mutate(
    decile = ntile(predicted, 10)
  ) %>%
  group_by(decile) %>%
  summarise(
    mean_actual = mean(actual),
    mean_predicted = mean(predicted),
    .groups = "drop"
  )
