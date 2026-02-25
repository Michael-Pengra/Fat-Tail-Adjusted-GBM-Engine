Multi-Asset Stochastic Jump Diffusion Monte Carlo Model (Excel)

This is a fully Excel-native Monte Carlo risk engine designed for scenario and distribution analysis of financial and business metrics.

The model began as a single-asset Geometric Brownian Motion (GBM) framework and was later expanded into a three-asset correlated jump diffusion engine.

It is built entirely in Excel — no VBA, no external code.

What the Model Does

The engine simulates thousands of possible future paths for a given metric (stock price, revenue, EBITDA, portfolio value, etc.) over a chosen time horizon.

It does not attempt to predict a single outcome.
Instead, it measures:

Distribution of possible outcomes

Downside risk (VaR)

Tail probabilities

Confidence intervals

Probability of breaching a target value

The mean and median will center around the user’s expected growth input if the model is calibrated correctly. The purpose is risk measurement and range analysis.

Core Architecture
1. Fat-Tailed Distribution

Instead of assuming normally distributed returns, the model uses a Student’s t-distribution.

User selects a desired 3-sigma probability.

Degrees of freedom are automatically calibrated through a lookup table (no manual Goal Seek).

Linear interpolation allows fractional degrees of freedom for smoother tail control.

This allows realistic modeling of extreme market events.

2. Jump Diffusion

To account for crashes, earnings shocks, or major macro events, the model includes a jump component:

Randomized number of shocks per year

Randomized shock magnitude

Independent shock volatility

Jump frequency can be suggested based on historical tail-event counts.

This prevents the unrealistic “smooth compounding only” behavior typical of basic GBM models.

3. Multi-Asset Correlation (Latest Update)

The model was expanded into a 3-asset multivariate framework.

A Cholesky decomposition matrix is used to correlate the continuous diffusion shocks.

Systematic risk is correlated.

Jump components remain independent per asset (to preserve idiosyncratic event risk).

Each asset runs 50,000 simulations (150,000 total rows calculated per refresh).

The architecture was expanded to 49 named ranges to keep formulas readable and prevent reference drift.

Workbook size increased from 11MB to 51MB after the multivariate update.

Outputs

For each asset (and portfolio, if weighted):

Terminal value distribution

Total return distribution

CAGR distribution

Mean, median, min, max

Standard deviation

Percentiles from 0.001% to 99.999%

1% and 5% Value at Risk

Confidence intervals up to 99.999%

Probability of falling below a target

Excess kurtosis

Dynamic histograms update automatically when the number of simulations changes.

The binning system was redesigned to properly visualize long-horizon lognormal distributions without being distorted by extreme compounding outliers.

Historical Suggestion Engine

The model includes hidden support sheets that:

Reconstruct up to 100 years of annual returns

Suggest expected growth inputs (3-year CAGR, full historical CAGR, median return, best/worst year)

Estimate historical 3-sigma event frequencies

Estimate average number of shock events per year

All suggestion values update dynamically when new raw data is entered.

Use Cases

The model is designed for any positive, proportional-growth metric that:

Cannot go negative

Scales with size

Has drift + volatility characteristics

Is not strongly mean-reverting

Examples:

Equity prices

Revenue growth paths

EBITDA projections

Portfolio NAV evolution

LBO equity value distributions

Exit multiple uncertainty

FX exposure

Free cash flow forecasts

Again, this is a scenario and risk tool — not a point forecast engine.

Versions

Two versions are maintained:

Single-Asset Version (leaner structure)

Multi-Asset Correlated Version (full Cholesky framework)
