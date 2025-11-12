# Advanced Analytical Model: GBM Simulation Tool

## 100,000 Monte Carlo Paths • Fat Tails • 100% Excel

### Advanced Analytical Model: Geometric Brownian Motion (GBM) Simulation Tool

This project is a powerful, general-purpose quantitative risk model developed entirely in Microsoft Excel. It performs a Monte Carlo Simulation to forecast the future risk range of any continuous metric (e.g., commodity prices, stock prices, currency rates).

---

## Project Context and Purpose

I developed this model as an undergraduate junior double majoring in Applied and Computational Mathematics and Finance.

* Model Type: Geometric Brownian Motion (GBM) Monte Carlo Simulation.
* Wide Applicability: Chosen over simpler models (like Black-Scholes) because GBM has a broader range of uses for forecasting metrics beyond options pricing.
* Academic Rigor: Addresses the main critique of traditional modeling—the flawed assumption of a Normal distribution—by integrating custom statistical methods.

---

## Superlative Technical Features

### 1. Extreme Risk Modeling (Custom Distribution)

* Custom Risk Input: The model allows the user to define the likelihood of an extreme, rare event (Probability of a 3sigma event).
* High Precision Math: This input switches the core math to a specialized formula (Student's t-distribution), which provides more realistic forecasts by fully accounting for potential large price swings.
* Problem Solved (Accuracy): Since Excel's built-in functions can't perfectly calculate this specialized formula for all numbers, I designed a linear interpolation formula. This custom solution ensures the model maintains high statistical accuracy.

### 2. Automation and Scale

| Achievement | Technical Implementation | Impact |
| :--- | :--- | :--- |
| Capacity | Runs up to 100,000 simulations in Excel. | Provides a large, statistically reliable sample for precise risk analysis. |
| Automated Setup | Replaced the old, manual Goal Seek process with an automatic VLOOKUP solution to set the required math parameters. | Ensures the model is fast, error-free, and simple to set up instantly. |
| Dynamic Data | Used the OFFSET function to create self-adjusting data ranges that resize automatically when the simulation count changes. | Guarantees all charts and final outputs are always complete. |
| Structure | Built with over 20 Named Ranges and uses a formula to correctly convert the user's Annualized Growth Rate into the specific Drift term required for the GBM math. | Ensures mathematical accuracy across different time horizons. |

---

## Outputs and Reporting

The model provides comprehensive reporting across three formats: Metric Value, Total Percent Growth, and Average % Growth Per Year (CAGR).

* Risk Metrics: Provides Mean, Median, Min, Max, Standard Deviation, Percentiles, and Excess Kurtosis (a direct measure of risk concentration).
* Downside Analysis: Calculates the maximum probable loss (VaR) at 1% and 5% confidence levels, dynamically labeled as "Value at Risk (Loss)" or "Downside Profit Floor."
* Professional Charts: Features a dynamic histogram with automated data grouping and professionally formatted axis labels (created using nested text functions) for superior visualization.
* Data Sourcing: Historical Log Returns are calculated from time-series data.
