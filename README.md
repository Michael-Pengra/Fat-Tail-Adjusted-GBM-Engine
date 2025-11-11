Geometric Brownian Motion (GBM) Monte Carlo Simulation Model
1. Model Purpose and Core Innovation
This project is a detailed financial model, built entirely in Microsoft Excel, that uses a Monte Carlo Simulation based on Geometric Brownian Motion (GBM) to forecast the future distribution of an asset's value.

The core of this model is its ability to create a realistic risk profile. Standard models assume markets follow a perfect bell curve (Normal Distribution), which is often wrong. My model solves this by letting the user define the risk, specifically:

Addressing Extreme Events: The model incorporates a t-distribution for the random market factor, which naturally allows for a much higher frequency of large price movements (like 3-standard deviation events) than the traditional normal curve.

User-Defined Risk: The user inputs the Desired Probability for a large price movement. The model then uses Excel's Goal Seek function to calculate the precise parameters needed to match that specific risk level.

Handling Excel Limitations: A specific linear interpolation formula was developed to allow the model to work with non-integer parameters in Excel's statistical functions, a necessary step for accurate calibration.

2. Data Processing and Formulas
Data Sourcing: Historical price data is pulled and manually pasted into the RawData sheet. The model then calculates daily Log Returns for volatility analysis.

Volatility: Daily Uncertainty is calculated as the standard deviation of the most recent three years of Log Returns.

Growth Rate Conversion: The model accepts a straightforward Simple Annual Growth Rate (CAGR) as input, which is then converted into the continuous-time Drift required for the GBM formula. This conversion ensures the simple input is correctly translated into the simulation's exponential growth logic:

Drift Formula (Copy-Pasteable): Drift = ( (LN(1+Simple Return) * Investment Horizon) + (0.5 * Annualized Uncertainty ^ 2) ) / Investment Horizon

Simulation Formula (Core Logic): The final price in each simulation is calculated using the GBM equation which incorporates the annualized terms and the custom t-distribution factor.

3. Model Architecture and Technical Features
Five-Sheet Structure: RawData, Inputs, Simulation 1 (Daily Steps), Simulation 2 (Final Outcomes), and Output.

Dynamic Automation: Over 20 Named Ranges are used for formula clarity and model reliability. The simulation output range is dynamically controlled using the OFFSET function to automatically resize based on the user-defined number of simulations.

Visualization Automation: The output sheet features a Dynamic Histogram that automatically calculates and adjusts its bins and labels when the number of simulations is changed. Custom formulas are used to ensure the axis labels are clear and accurate.

4. Output and Risk Reporting
The final output is a full report on the thousands of simulated price outcomes, providing both performance metrics and risk measures.

Risk Metrics (Value at Risk - VaR): The model calculates the 1% and 5% metric values. This is labelled Value at Risk (Loss) if the value is below the initial price, or Downside Profit Floor if it is still a gain.

Distribution Measurement (Non-Normality): The model calculates Excess Kurtosis. This is a critical measure of how non-normal the distribution is; the higher this number, the greater the probability of extreme price movements (either positive or negative).

Performance Metrics: Mean, Median, Min, Max, Standard Deviation, and Avg % Growth Per Year (CAGR) are reported, giving a complete statistical overview of the possible outcomes.

Percentiles: Comprehensive percentile data (1%, 5%, 25%, 50%, 75%, 95%, 99%) is provided to define the exact range of outcomes at various confidence levels.
