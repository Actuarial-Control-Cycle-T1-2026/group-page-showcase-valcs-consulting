[![Review Assignment Due Date](https://classroom.github.com/assets/deadline-readme-button-22041afd0340ce965d47ae6ef1cefeee28c7c493a6346c4f15d667ab976d596c.svg)](https://classroom.github.com/a/FxAEmrI0)
# Actuarial Theory and Practice A: 2026 SOA Research Challenge

By Alan Steny, Carol Zhang, Liya Ruan, Sarina Truong & Vihaan Jain

---
# Objective Overview

We are members of the actuarial team at Galaxy General Insurance Company, one of the leading insurance companies across the interstellar expanse in 2075. The objective of this is to develop a successful response to a request for proposals (RFP) from Cosmic Quarry Mining Corporation who is seeking space mining insurance coverage for their expanding operation. This report analyses the associated risk and provide recommendations regarding the following products: equipment failure, cargo loss, workers’ compensation, and business interruption.

---
# Data
The reports analysis relied solely on project provided data, key limitations, assumptions, and their impacts are summarised below.
<img width="801" height="370" alt="image" src="https://github.com/user-attachments/assets/d7c2f2b8-82f7-4435-a7a2-6a1eb491a584" />

The R Code for data cleaning can be accessed below:

[Equipment Failure](https://github.com/Actuarial-Control-Cycle-T1-2026/group-page-showcase-valcs-consulting/blob/main/Equipement%20Failure%20Data%20Cleaning.R)

[Cargo Loss](https://github.com/Actuarial-Control-Cycle-T1-2026/group-page-showcase-valcs-consulting/blob/main/Cargo%20Loss%2001%20-%20Data%20Cleaning.r)

[Business Interruption](https://github.com/Actuarial-Control-Cycle-T1-2026/group-page-showcase-valcs-consulting/blob/main/Business%20Interruption%2001_data_cleaning.py)

[Workers' Compensation](https://github.com/Actuarial-Control-Cycle-T1-2026/group-page-showcase-valcs-consulting/blob/main/WC%20-%2001_DataCleaning.R)


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

The capital model for Cosmic Quarry’s portfolio is projected across a 2175 - 2185 window using 10,000 bootstrapped simulations for each product line. The simulations were then stress tested under 11 stress testing scenarios and are mapped across 3 capital requirements (100%, 150%, 200% SCR). In addition, the model incorporates a 40% quota share reinsurance structure with a 25% ceding commission returned to the insurer.

**Code**

The code for the capital model can be accessed below:

[Capital Modelling](https://github.com/Actuarial-Control-Cycle-T1-2026/group-page-showcase-valcs-consulting/blob/main/Capital%20Model%20Code.py)

**Outputs**

The outputs for the above capital model can be accessed below:
[Capital Model Outputs](https://github.com/Actuarial-Control-Cycle-T1-2026/group-page-showcase-valcs-consulting/blob/main/Capital%20Model%20Summary%20Excel%20.xlsx)

---
# Risk assessment

Solar system risk profiles are defined by unique stellar and orbital characteristics. The Helionis Cluster features irregular gravitational resonances and shifting debris that require periodic satellite repositioning. In contrast, the Bayesia System’s binary stars generate electromagnetic spikes and radiation extremes. Finally, the Oryn Delta presents a low-visibility environment with unpredictable solar flares and an asymmetric asteroid ring, which creates hazardous gravitational gradients for extraction infrastructure.

<img width="813" height="196" alt="image" src="https://github.com/user-attachments/assets/617ac83c-53ad-4d5c-ab15-34973e5ecf92" />
<img width="811" height="214" alt="image" src="https://github.com/user-attachments/assets/a9a59d2c-4583-4eb1-bc6d-bf28b35d71c4" />



---
# Assumptions

<img width="767" height="116" alt="image" src="https://github.com/user-attachments/assets/29326fb7-e592-47cd-8c1d-dea11ac40b75" />
<img width="765" height="340" alt="image" src="https://github.com/user-attachments/assets/c27191c1-2e4d-485f-94cd-ba99eb93539f" />



---
# Product Design

**Cargo Loss**

While Cargo Loss represents a significant operational risk for Cosmic Quarry Mining Corporation, our analysis indicates that offering cargo loss insurance is not currently commercially viable for Galaxy General. The transport of extremely high-value metals such as gold and platinum creates severe loss potential, requiring substantial capital and resulting in premiums that would likely be unsustainable for the client. Reducing premiums to make coverage affordable is not viable as this would expose the Galaxy General to losses that are not adequately compensated by premium. 

Accordingly, we do not recommend introducing a cargo loss product. However, if the value concentration of transported metals decrease in the future, or if market conditions allow risks to become more diversifiable and predictable, Galaxy General may reconsider offering cargo loss coverage.

**Equipment Failure**

Coverage trigger: Equipment failure insurance is designed to cover the costs associated with repairing or replacing specific equipment following an unforeseen and sudden physical failure.

Benefit structure: Galaxy General will pay for the cost to restore the machine to its original condition prior to breakdown, including the costs of dismantling, freight, replacement parts and re-installation. If the cost of repair exceeds the actual cash value of the equipment at the time of breakdown, Galaxy General will pay the actual cash value of the equipment.

Exclusions: Galaxy General will not cover any events outside the defined benefit structure, namely any equipment breakdowns resulting from causes that are not sudden and unforeseen.

**Workers Compensation**

Coverage Trigger: Workers’ Compensation coverage is triggered when Cosmic Quarry employees sustain work-related injuries or illnesses across any operating solar system. Coverage applies to all occupations, employment types, and locations, ensuring broad protection aligned with the social insurance mandate.

Benefit Structure: The policy provides unlimited per-claim coverage with no deductibles or sub-limits, covering medical costs, rehabilitation, and income replacement, with premiums risk-rated per employee and portfolio stop-loss reinsurance applied at extreme loss levels.

Exclusions: Coverage excludes pre-existing conditions, off-duty incidents, intoxication-related injuries, self-inflicted injuries, acts of war, and claims arising within the initial 90-day seasoning period.

**Business Interreuption**

Coverage Trigger: Business Interruption coverage is triggered by operational shutdowns or significant production disruptions, including equipment failure, supply chain interruptions, evacuations, debris field closures, or workforce shortages.

Benefit Structure: The policy indemnifies lost revenue based on operational downtime, with a $1.0M deductible, 14-day waiting period, $15M per-occurrence limit, and $45M annual aggregate, with premiums risk-rated based on operational and system characteristics.

Exclusions: Coverage excludes scheduled maintenance, gradual deterioration, market price changes, pre-existing faults, acts of war, double recovery, claims within the first 90 days, and regulatory or compliance-related penalties.









