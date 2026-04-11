##############################################################################
# Cosmic Quarry Mining - Predictive Dataset Construction
##############################################################################

# ---------------------------------------------------------------------------
# Load packages
# ---------------------------------------------------------------------------
library(tidyverse)
library(fitdistrplus)

# ---------------------------------------------------------------------------
# Version control
# ---------------------------------------------------------------------------
# v1 | Sarina Truong | Initial predictive dataset
# v2 | Cleaned       | Standardised naming and structure

##############################################################################
# 1. Container Fleet Table
##############################################################################

container_fleet <- tibble(
  solar_system = c(
    rep("Helionis Cluster", 5),
    rep("Bayesia System", 5),
    rep("Oryn Delta", 5)
  ),
  
  container_type = rep(
    c(
      "DeepSpace Haulbox",
      "DockArc Freight Case",
      "HardSeal Transit Crate",
      "LongHaul Vault Canister",
      "QuantumCrate Module"
    ),
    3
  ),
  
  number_of_containers = c(
    58, 116, 580, 232, 174,
    56, 113, 564, 226, 169,
    39, 77, 387, 155, 116
  ),
  
  max_weight = rep(
    c(
      25000,
      50000,
      100000,
      150000,
      250000
    ),
    3
  )
)

##############################################################################
# 2. Cargo Weight Ranges
##############################################################################

cargo_weight_ranges <- tibble(
  cargo_type = c(
    "cobalt",
    "gold",
    "lithium",
    "platinum",
    "rare earths",
    "supplies",
    "titanium"
  ),
  
  min_weight = c(
    100000,
    2500,
    75000,
    1500,
    25000,
    50000,
    125000
  ),
  
  cargo_max_weight = c(
    200000,
    5000,
    150000,
    3000,
    50000,
    100000,
    250000
  )
)

##############################################################################
# 3. Cargo Value Per KG
##############################################################################

cargo_value_lookup <- tibble(
  cargo_type = c(
    "gold",
    "cobalt",
    "platinum",
    "lithium",
    "supplies",
    "rare earths",
    "titanium"
  ),
  
  value_per_kg = c(
    135600,
    52,
    54500,
    82,
    10,
    85,
    7
  )
)

##############################################################################
# 4. Expand to One Row Per Container
##############################################################################

predictive_data <- container_fleet %>%
  uncount(number_of_containers) %>%
  mutate(
    container_id = row_number()
  )

##############################################################################
# 5. Assign Cargo Types
##############################################################################

cargo_types <- cargo_weight_ranges$cargo_type

predictive_data <- predictive_data %>%
  group_by(solar_system) %>%
  mutate(
    cargo_type =
      cargo_types[
        (row_number() - 1) %% length(cargo_types) + 1
      ]
  ) %>%
  ungroup()

##############################################################################
# 6. Join Weight and Value Tables
##############################################################################

predictive_data <- predictive_data %>%
  left_join(
    cargo_weight_ranges,
    by = "cargo_type"
  ) %>%
  left_join(
    cargo_value_lookup,
    by = "cargo_type"
  )

##############################################################################
# 7. Simulate Cargo Weights
##############################################################################

set.seed(123)

predictive_data <- predictive_data %>%
  mutate(
    simulated_weight =
      runif(
        n(),
        min_weight,
        cargo_max_weight
      ),
    
    weight =
      pmin(
        simulated_weight,
        cargo_max_weight
      ),
    
    cargo_value =
      weight * value_per_kg
  )

##############################################################################
# 8. Round Values
##############################################################################

predictive_data <- predictive_data %>%
  mutate(
    simulated_weight = round(simulated_weight, 0),
    weight = round(weight, 0),
    cargo_value = round(cargo_value, 0)
  )

##############################################################################
# 9. Solar Radiation, Debris Density and Route Risk
##############################################################################

predictive_data <- predictive_data %>%
  mutate(
    
    solar_radiation =
      case_when(
        solar_system == "Helionis Cluster" ~ 0.4,
        solar_system == "Bayesia System" ~ 0.9,
        solar_system == "Oryn Delta" ~
          ifelse(
            runif(n()) < 0.05,
            0.9,
            0.1
          )
      ),
    
    debris_density =
      case_when(
        solar_system == "Helionis Cluster" ~ runif(n(), 0.7, 0.8),
        solar_system == "Bayesia System" ~ runif(n(), 0.2, 0.3),
        solar_system == "Oryn Delta" ~ runif(n(), 0.8, 0.9)
      ),
    
    route_risk =
      factor(
        case_when(
          solar_system == "Helionis Cluster" ~ sample(c(3,4), n(), replace = TRUE),
          solar_system == "Bayesia System" ~ 2,
          solar_system == "Oryn Delta" ~ 5
        ),
        levels = c(1,2,3,4,5)
      )
  )

##############################################################################
# 10. Exposure Simulation
##############################################################################

historical_exposure <- cargo_frequency$exposure

epsilon <- 1e-6

exposure_adjusted <- pmin(
  pmax(
    historical_exposure,
    epsilon
  ),
  1 - epsilon
)

fit_beta <- fitdist(
  exposure_adjusted,
  "beta",
  method = "mle"
)

alpha_hat <- fit_beta$estimate["shape1"]
beta_hat  <- fit_beta$estimate["shape2"]

set.seed(123)

predictive_data <- predictive_data %>%
  mutate(
    exposure =
      rbeta(
        n(),
        shape1 = alpha_hat,
        shape2 = beta_hat
      )
  )

##############################################################################
# 11. Simulate Pilot Experience
##############################################################################

predictive_data <- predictive_data %>%
  mutate(
    pilot_experience =
      rnorm(
        n(),
        mean = 15.00368,
        sd = 4.987284
      )
  )

##############################################################################
# 12. Simulate Distance
##############################################################################

predictive_data <- predictive_data %>%
  mutate(
    distance =
      rgamma(
        n(),
        shape = 1.3535152,
        rate = 0.0516307
      )
  )

##############################################################################
# 13. Simulate Vessel Age
##############################################################################

vessel_data <- cargo_frequency$vessel_age

fit_gamma_vessel <- fitdist(
  vessel_data,
  "gamma"
)

shape_hat <- fit_gamma_vessel$estimate["shape"]
rate_hat  <- fit_gamma_vessel$estimate["rate"]

set.seed(123)

predictive_data <- predictive_data %>%
  mutate(
    vessel_age =
      rgamma(
        n(),
        shape = shape_hat,
        rate = rate_hat
      )
  )

##############################################################################
# 14. Final Predictive Dataset
##############################################################################

predictive_data