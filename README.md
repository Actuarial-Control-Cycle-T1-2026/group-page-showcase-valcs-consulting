[![Review Assignment Due Date](https://classroom.github.com/assets/deadline-readme-button-22041afd0340ce965d47ae6ef1cefeee28c7c493a6346c4f15d667ab976d596c.svg)](https://classroom.github.com/a/FxAEmrI0)
# Actuarial Theory and Practice A: 2026 SOA Research Challenge

By Alan Steny, Carol Zhang, Liya Ruan, Sarina Truong & Vihaan Jain

---
# Objective Overview

We are members of the actuarial team at Galazy General Insurance Company, one of the leading insurance companies across the interstellar expanse in 2075. The objective of this is to develop a successful response to a request for proposals (RFP) from Cosmic Quarry Mining Corporation who is seeking space mining insurance coverage for their expanding operation. This report analyses the associated risk and provide recommendations regarding the following products: equipment failure, cargo loss, workers’ compensation, and business interruption.

---
# Data


---
# Loss frequency and severity modelling

**Frequency Modelling**

A Negative Binomial GLM was selected as the frequency model across all product lines. Claim counts are discrete, non-negative, and preliminary exploratory data analysis indicated variability that violates the Poisson assumption. Given the concentration of policies with zero claims, a Zero-Inflated Negative Binomial model was initially considered to explicitly account for the excess zeros. However, the EDA revealed that the zero counts do not arise from a separate process, such as deductible thresholds. Consequently, the standard Negative Binomial was preferred, as it accommodates overdispersion while maintaining interpretability.

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
