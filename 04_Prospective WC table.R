###############################################################################
# 03_prospective_portfolio.R
# Build the CQ prospective portfolio for pricing
###############################################################################
# Version Control
# v1  | Vihaan Jain  | 01 Mar 2026 | Initial personnel import + occupation mapping
# v2  | Vihaan Jain  | 02 Mar 2026 | Row expansion with salary/age/exp simulation
# v3  | Vihaan Jain  | 03 Mar 2026 | Solar system mapping assumption documented
# v4  | Vihaan Jain  | 04 Mar 2026 | Factor validation, NA checks, summary printout
###############################################################################

library(dplyr)
library(tidyr)
library(readxl)
library(purrr)

set.seed(42)


# ---- 1) Import personnel file ----

personnel_raw <- read_excel(
  "~/Library/Mobile Documents/com~apple~CloudDocs/ACTL4001 Assignment/SOA_2026_Case_Study_Materials/0. Original Case Material/srcsc-2026-cosmic-quarry-personnel.xlsx",
  sheet = "Personnel", skip = 2
) %>%
  filter(!is.na(...1), !(...1 %in% c(
    "Management", "Administration", "Environmental & Safety",
    "Exploration Operations", "Extraction Operations", "Spacecraft Operations"
  ))) %>%
  rename(
    occupation = ...1,
    n_total    = `Number of Employees`,
    n_fulltime = `Full-Time Employees`,
    n_contract = `Contract Employees`,
    avg_salary = `Average Annualized Salary (Đ)`,
    avg_age    = `Average Age`
  ) %>%
  mutate(across(c(n_total, n_fulltime, n_contract, avg_salary, avg_age), as.numeric)) %>%
  filter(!is.na(n_total)) %>%
  mutate(occupation_clean = case_when(
    occupation == "Maintenance" & row_number() == 15 ~ "Maintenance",
    occupation == "Maintenance" & row_number() == 20 ~ "Spacecraft Maintenance",
    TRUE ~ occupation
  ))

cat("Personnel loaded:", nrow(personnel_raw), "roles |", sum(personnel_raw$n_total), "employees\n")


# ---- 2) Mappings + factor levels ----

# maps CQ job titles to the occupation levels the model was trained on
occupation_map <- c(
  "Executive"                = "Executive",
  "Vice President"           = "Manager",
  "Director"                 = "Manager",
  "HR"                       = "Administrator",
  "IT"                       = "Technology Officer",
  "Legal"                    = "Administrator",
  "Finance & Accounting"     = "Administrator",
  "Environmental Scientists" = "Scientist",
  "Safety Officer"           = "Safety Officer",
  "Medical Personel"         = "Administrator",
  "Geoligist"                = "Scientist",
  "Scientist"                = "Scientist",
  "Field technician"         = "Planetary Operations",
  "Drilling operators"       = "Drill Operator",
  "Maintenance"              = "Maintenance Staff",
  "Engineers"                = "Engineer",
  "Freight operators"        = "Planetary Operations",
  "Robotics technician"      = "Technology Officer",
  "Navigation officers"      = "Spacecraft Operator",
  "Spacecraft Maintenance"   = "Maintenance Staff",
  "Security personel"        = "Spacecraft Operator",
  "Steward"                  = "Administrator",
  "Galleyhand"               = "Planetary Operations"
)

valid_occupations <- levels(wc_freq_model$occupation)
valid_emp_type    <- levels(wc_freq_model$employment_type)  # "Contract", "Full-time"
valid_solar       <- levels(wc_freq_model$solar_system)     # "Epsilon", "Helionis Cluster", "Zeta"

# solar system mapping assumptions:
# CQ operates in: Helionis Cluster, Bayesia System, Oryn Delta
# Model was trained on: Helionis Cluster, Epsilon, Zeta
#   Bayesia  -> Epsilon   (binary star, elevated radiation, harsh conditions)
#   Oryn Delta -> Zeta    (dim dwarf star, remote, cold outer system)
# weighted by mine count: 30 / 15 / 10
solar_systems <- c("Helionis Cluster", "Epsilon", "Zeta")
solar_weights <- c(30, 15, 10) / 55


# ---- 3) Expand to individual rows ----

cq_portfolio <- personnel_raw %>%
  mutate(occupation_model = occupation_map[occupation_clean]) %>%
  filter(!is.na(occupation_model)) %>%
  pmap_dfr(function(occupation, occupation_clean, occupation_model,
                    n_total, n_fulltime, n_contract, avg_salary, avg_age) {
    
    n         <- as.integer(n_total)
    emp_types <- sample(c(rep("Full-time", as.integer(n_fulltime)),
                          rep("Contract",  as.integer(n_contract))))
    
    ages     <- pmax(18L, pmin(70L, round(rnorm(n, avg_age, 5))))
    exp_yrs  <- pmax(0, ages - 22 + round(rnorm(n, 0, 1)))
    
    # log-normal salary draw with 15% CV, recentred to preserve avg_salary
    sigma_sal <- sqrt(log(1 + 0.15^2))
    salaries  <- round(rlnorm(n, log(avg_salary) - 0.5 * sigma_sal^2, sigma_sal), -2)
    
    solar    <- sample(solar_systems, n, replace = TRUE, prob = solar_weights)
    exposure <- round(runif(n, 0.5, 1.0), 4)
    
    data.frame(
      cq_role         = occupation,
      occupation      = occupation_model,
      employment_type = emp_types,
      solar_system    = solar,
      age             = ages,
      experience_yrs  = exp_yrs,
      base_salary     = salaries,
      exposure        = exposure,
      stringsAsFactors = FALSE
    )
  })


# ---- 4) Factorise + validate ----

cq_portfolio <- cq_portfolio %>%
  mutate(
    occupation      = factor(occupation,      levels = valid_occupations),
    employment_type = factor(employment_type, levels = valid_emp_type),
    solar_system    = factor(solar_system,    levels = valid_solar),
    log_exposure    = log(exposure)
  )

# any NAs here means an occupation or solar_system fell outside model levels
cat("\nNA check:\n")
cat("  occupation NA:     ", sum(is.na(cq_portfolio$occupation)),      "\n")
cat("  employment_type NA:", sum(is.na(cq_portfolio$employment_type)), "\n")
cat("  solar_system NA:   ", sum(is.na(cq_portfolio$solar_system)),    "\n")

# sanity check simulated values against source data
cq_portfolio %>%
  group_by(cq_role) %>%
  summarise(n           = n(),
            mean_salary = round(mean(base_salary), 0),
            mean_age    = round(mean(age), 1),
            mean_exp    = round(mean(experience_yrs), 1),
            .groups = "drop") %>%
  left_join(dplyr::select(personnel_raw, cq_role = occupation, avg_salary, avg_age),
            by = "cq_role") %>%
  mutate(salary_err_pct = round((mean_salary / avg_salary - 1) * 100, 1),
         age_err_pct    = round((mean_age    / avg_age    - 1) * 100, 1)) %>%
  print(n = Inf, width = Inf)

cat(sprintf("\nPortfolio rows: %d | Personnel total: %d\n",
            nrow(cq_portfolio), sum(personnel_raw$n_total)))
cat("Columns:", paste(names(cq_portfolio), collapse = ", "), "\n")