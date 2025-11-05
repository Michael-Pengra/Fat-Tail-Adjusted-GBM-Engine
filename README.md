Fat-Tail Adjusted Multi-Metric GBM Monte Carlo Engine
I personally designed and built this Generalized Geometric Brownian Motion (GBM) Monte Carlo simulation model entirely within Excel. It's engineered to be a powerful tool, providing robust long-term forecasts and risk analysis for any financial or operational metric that exhibits continuous, compounding growth or change.

The Quantitative Solution
I tackled the major statistical limitation of the Normal Distribution, which fails to account for the true frequency of extreme market events. This required implementing the Student's t-distribution to accurately model higher kurtosis. Software Limitation Bypass: A key challenge was that the Excel t-distribution function ignored the decimal Degrees of Freedom (DF) required for precise calibration. I engineered a linear interpolation formula to bypass this constraint, ensuring numerical precision and an exact match between the required P(3Sigma) and the model's output distribution.

Core Formulas 
Actual P (3Sigma) (Inputs)
=T.DIST.2T(3,INT(C18))+(C18-INT(C18))*(T.DIST.2T(3,INT(C18)+1)-T.DIST.2T(3,INT(C18)))
Final Price Simulation (Simulation 2)
=IF(NumberofSimulations>=A6,CurrentMetricValue*EXP((Drift*InvestmentHorizon-0.5*POWER(AnnualizedUncertainty,2))+(AnnualizedUncertainty*SQRT(TimeStepSize)*SQRT(InvestmentHorizon))*(T.INV(C4,INT(DF))+(DF-INT(DF))*(T.INV(C4,INT(DF)+1)-T.INV(C4,INT(DF))))),"")

Technical Architecture & Advanced Features
The model's robustness relies on dynamic named ranges and advanced visualization control:
Input Autocalculation (Array Manipulation)
The model automatically calculates key inputs (like uncertainty and observation window) using advanced dynamic formulas: 
Daily Uncertainty: =STDEV.S(OFFSET(RawData!D:D,MAX(0,COUNT(RawData!D:D)-756),0,MIN(756, COUNT(RawData!D:D)),1)) (Calculates standard deviation over the most recent 756 days or total available data).
Years of Raw Data: =ROUND((INDEX(RawData!B:B,COUNTA(RawData!B:B))-RawData!$B$3)/365,1) (Calculates the total time span of the raw data).

Dynamic Named Ranges:
SimRange=OFFSET('Simulation 2'!$B$4,0,0,NumberOfSimulations,1) Dynamically captures all simulation results (50,000+ cells).
BinWidth=ROUNDUP(RobustRange/NumBins, 2) The final calculated bin size for histogram visualization.
RobustRange=PERCENTILE.INC(SimRange, 0.999)-PERCENTILE.INC(SimRange, 0.001) Sets bin width based on 99.8% of data (fat-tail adjustment).

Dynamic Histogram Generation 
The histogram is fully dynamic, featuring a custom, visually clear x-axis: 
Bin Label Generation: =CONCAT("$",LEFT(I5,6),"-$",LEFT(I6,6)) (Used to dynamically build descriptive labels like "$350.00-$360.00" for the chart's x-axis).
Frequency Calculation: =FREQUENCY(SimRange, I5:I23) is used to calculate the counts across the dynamic SimRange. 

Demonstration & Results
Using DRI as a proof-of-concept (7.00% Expected Growth Rate), the model demonstrates its utility:
Forecast: Mean Price of $355.57 (7.04% Compounded Annual Growth Rate).
Risk Quantification: For a $350.00$ target, the model calculates a 37.6% probability of falling below target.
Kurtosis: The model confirms high volatility with Excess Kurtosis measured at 5.95.

How to Run the Model
1. Download the GBM_MultiMetric_Engine.xlsx file.
2. Open the file and navigate to the designated sheet to add the historical raw data. This data automatically calculates inputs like uncertainty and total steps.
3. Navigate to the Inputs sheet to set the core manual metrics (e.g., Investment Horizon, Drift, Number of Simulations).
4. Run Goal Seek (on the Inputs sheet) to find the appropriate Decimal Degrees of Freedom that achieves the desired P(3sigma) next to where the Degree of Freedom is calculated.
5. View the final forecast and risk metrics on the Output sheet.
