Stochastic Jump Diffusion Model
This is a high performance quantitative risk model I built in Excel that uses a Stochastic Jump Diffusion Monte Carlo Simulation to forecast asset price paths. Unlike basic models, it uses fat tail distributions and discrete market shocks to capture the messy reality of financial markets.

Core Features
1. Fat Tail Risk
It solves the normal distribution flaw by using a Student’s t-distribution with linear interpolation for high statistical accuracy.

It uses a lookup engine to instantly set degrees of freedom based on user defined risk parameters, which replaces the old manual goal seek process.

2. Modeling Crashes
The model randomizes the frequency of jumps or shocks over the investment horizon to reflect how risk actually clusters in the real world.

It uses three different random variables per path to keep market sentiment, crash count, and crash severity uncorrelated.

3. Scale and Performance
It runs 100,000 simulation paths in under two seconds.

The model has a suggestion engine that scans 10 years of data to suggest CAGR, volatility, and 8-sigma event frequencies.

Everything is dynamic using named ranges and offset functions so the charts and histograms update in real time.

Risk Analytics and Dashboards
It outputs the mean, median, min, max, and standard deviation for the metric value, total return, and CAGR.

It calculates Value at Risk at 1% and 5% levels and provides automated confidence intervals up to 99.999%.

There is a dynamic histogram with automated data grouping and professionally formatted axis labels.

Strategic Use Cases
This is a tool for scenario and risk analysis across different areas of finance:

Asset Management: Looking at portfolio value evolution and FX rate risk.

Investment Banking and PE: Modeling LBO equity IRR risk and exit multiple uncertainty.

Corporate Strategy: Forecasting revenue and EBITDA growth paths subject to discrete shocks.
