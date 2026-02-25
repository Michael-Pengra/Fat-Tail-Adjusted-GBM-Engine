# Multi-Asset Stochastic Jump Diffusion Monte Carlo Model (Excel)

This is a fully Excel-native Monte Carlo risk engine designed for scenario and distribution analysis of financial and business metrics.
The model began as a single-asset Geometric Brownian Motion (GBM) framework and was later expanded into a three-asset correlated jump diffusion engine. It is built entirely in Excel - no VBA, no external code.

## What the Model Does
The engine simulates thousands of possible future paths for a given metric (stock price, revenue, EBITDA, portfolio value, etc.) over a chosen time horizon.
It does not attempt to predict a single outcome. Instead, it measures:

* Distribution of possible outcomes
* Downside risk (VaR)
* Tail probabilities
* Confidence intervals
* Probability of breaching a target value

The mean and median will center around the user's expected growth input if the model is calibrated correctly. The purpose is risk measurement and range analysis, not point forecasting.

## Core Architecture

### 1. Fat-Tailed Distribution
Uses Student's t-distribution instead of normal returns
User selects a desired 3-sigma probability
Degrees of freedom are automatically calibrated via a lookup table (no manual Goal Seek)
Linear interpolation allows fractional degrees of freedom for smoother tail control
Realistic modeling of extreme market events

### 2. Jump Diffusion
Accounts for crashes, earnings shocks, or macro events via a jump component
Randomized number of shocks per year
Randomized shock magnitude
Independent shock volatility
Jump frequency can be suggested based on historical tail-event counts
Prevents unrealistic "smooth compounding only" behavior typical of basic GBM models

### 3. Multi-Asset Correlation (Latest Update)
Expanded into a 3-asset multivariate framework
Uses Cholesky decomposition to correlate continuous diffusion shocks
Systematic risk is correlated; jump components remain independent per asset
Each asset runs 50,000 simulations (150,000 total rows calculated per refresh)
Architecture expanded to 49 named ranges to keep formulas readable and prevent reference drift
Workbook size grew from 11MB -> 51MB after multivariate update

## Outputs
For each asset (and portfolio if weighted):

* Terminal value distribution
* Total return distribution
* CAGR distribution
* Mean, median, min, max
* Standard deviation
* Percentiles from 0.001% -> 99.999%
* 1% and 5% Value at Risk (VaR)
* Confidence intervals up to 99.999%
* Probability of falling below a target
* Excess kurtosis
* Dynamic histograms update automatically with the number of simulations
* Bin system redesigned to properly visualize long-horizon lognormal distributions without distortion by extreme compounding outliers

## Historical Suggestion Engine
Hidden support sheets include:

* Reconstruction of up to 100 years of annual returns
* Suggested expected growth inputs:
* 3-year CAGR
* Full historical CAGR
* Median return
* Best/worst year
* Estimate historical 3-sigma event frequencies
* Estimate average number of shock events per year
All suggestions update dynamically when new raw data is entered.

## Use Cases
Designed for any positive, proportional-growth metric that:

* Cannot go negative
* Scales with size
* Has drift + volatility characteristics
* Is not strongly mean-reverting

Examples:

* Equity prices
* Revenue growth paths
* EBITDA projections
* Portfolio NAV evolution
* LBO equity value distributions
* Exit multiple uncertainty
* FX exposure
* Free cash flow forecasts
This is a scenario and risk tool, not a point forecast engine.

## Versions
* Single-Asset Version - leaner structure
* Multi-Asset Correlated Version - full Cholesky framework

Multi-Asset File Link: https://fsu-my.sharepoint.com/:x:/r/personal/mhp22b_fsu_edu/Documents/Multi_Asset_GBM_MultiMetricEngine.xlsx?d=w7667b90788444c8e965a22dffced3e8d&csf=1&web=1&e=EB50vL
