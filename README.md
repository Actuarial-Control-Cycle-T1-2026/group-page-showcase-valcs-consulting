[![Review Assignment Due Date](https://classroom.github.com/assets/deadline-readme-button-22041afd0340ce965d47ae6ef1cefeee28c7c493a6346c4f15d667ab976d596c.svg)](https://classroom.github.com/a/FxAEmrI0)
# Actuarial Theory and Practice A: 2026 SOA Research Challenge

By Alan Steny, Carol Zhang, Liya Ruan, Sarina Truong & Vihaan Jain

---
# Objective Overview

We are members of the actuarial team at Galaxy General Insurance Company, one of the leading insurance companies across the interstellar expanse in 2075. The objective of this is to develop a successful response to a request for proposals (RFP) from Cosmic Quarry Mining Corporation who is seeking space mining insurance coverage for their expanding operation. This report analyses the associated risk and provide recommendations regarding the following products: equipment failure, cargo loss, workers’ compensation, and business interruption.

---
# Data


---
# Loss frequency and severity modelling

**Frequency Modelling**

A Negative Binomial GLM was selected as the frequency model across all product lines. Claim counts are discrete, non-negative, and preliminary exploratory data analysis indicated variability that violates the Poisson assumption. Given the concentration of policies with zero claims, a Zero-Inflated Negative Binomial model was initially considered to explicitly account for the excess zeros. However, the EDA revealed that the zero counts do not arise from a separate process, such as deductible thresholds. Consequently, the standard Negative Binomial was preferred, as it accommodates overdispersion while maintaining interpretability.

**Severity Modelling**

**Code**


The R Code for each product line can be accessed below:

[Equipment Failure](https://github.com/Actuarial-Control-Cycle-T1-2026/group-page-showcase-valcs-consulting/blob/main/Equipment%20Failure.R)

[Cargo Loss - Severity Model](https://github.com/Actuarial-Control-Cycle-T1-2026/group-page-showcase-valcs-consulting/blob/main/Cargo%20Loss%2003%20-%20Severity%20Model.R)

[Cargo Loss - Frequency Model](https://github.com/Actuarial-Control-Cycle-T1-2026/group-page-showcase-valcs-consulting/blob/main/Cargo%20Loss%2004%20-%20Frequency%20Model.R)

---
# Capital Modelling

---
# Risk assessment

---
# Assumptions
---
# Product Design
**Equipment Failure**

Coverage trigger: Equipment failure insurance is designed to cover the costs associated with repairing or replacing specific equipment following an unforeseen and sudden physical failure.

Benefit structure: Galaxy General will pay for the cost to restore the machine to its original condition prior to breakdown, including the costs of dismantling, freight, replacement parts and re-installation. If the cost of repair exceeds the actual cash value of the equipment at the time of breakdown, Galaxy General will pay the actual cash value of the equipment.

Exclusions: Galaxy General will not cover any events outside the defined benefit structure, namely any equipment breakdowns resulting from causes that are not sudden and unforeseen.

**Cargo Loss**
While Cargo Loss represents a significant operational risk for Cosmic Quarry Mining Corporation, our analysis indicates that offering cargo loss insurance is not currently commercially viable for Galaxy General. The transport of extremely high-value metals such as gold and platinum creates severe loss potential, requiring substantial capital and resulting in premiums that would likely be unsustainable for the client. Reducing premiums to make coverage affordable is not viable as this would expose the Galaxy General to losses that are not adequately compensated by premium. 

Accordingly, we do not recommend introducing a cargo loss product. However, if the value concentration of transported metals decrease in the future, or if market conditions allow risks to become more diversifiable and predictable, Galaxy General may reconsider offering cargo loss coverage.

___

### Congrats on completing the [2026 SOA Research Challenge](https://www.soa.org/research/opportunities/2026-student-research-case-study-challenge/)!


> Now it's time to build your own website to showcase your work.  
> Creating a website using GitHub Pages is simple and a great way to present your project.

This page is written in Markdown.
- Click the [assignment link](https://classroom.github.com/a/FxAEmrI0) to accept your assignment.

---

> Be creative! You can embed or link your [data](player_data_salaries_2020.csv), [code](sample-data-clean.ipynb), and [images](ACC.png) here.

More information on GitHub Pages can be found [here](https://pages.github.com/).

![](Actuarial.gif)
