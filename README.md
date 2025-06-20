# raXL Stat, statistical Add-in for Data Science in Excel
raXL Stat, version v0.5 (June 19, 2025)
![raXL Stat v0 5](https://github.com/user-attachments/assets/841d7492-4868-40d0-a956-97fe03a8b2b4)

**raXL Stat is an add-in for Microsoft Excel that turns your favorite spreadsheet into a quantitative and predictive analysis software, offering a collection of functions to create statistical, econometric, financial and mathematical models.** You can call these functions directly from a spreadsheet and they will return the modeling results directly to it.

## Introduction
**raXL Stat**, is a statistical add-in for Microsoft Excel, developed in .NET with ExcelDna, that transforms spreadsheets into advanced tools for quantitative analysis, econometrics, finance, and time series. This manual details the installation, configuration, and use of all available user-defined functions (UDFs), organized by category, with full descriptions of each function, its purpose, and parameters.

raXL Stat is a statistical analysis software that will offer[^1] easy-to-use tools to perform and deliver quality work in a short time. It is developed[^2] to be used by both beginners and experts. The easiest and most intuitive way to run the functions is through the ribbon menu. If necessary, the user can directly write the functions in the spreadsheet cells or can invoke the functions from VBA (Visual Basic for Application) programming.

## System Requirements
- **Operating System**: Windows 7 or later.
- **Microsoft Excel**: Versions 2010 to 2024, or Microsoft 365 (32 or 64-bit).
- **Disk Space**: 10 MB free.
- **RAM**: Minimum 512 MB (4 GB recommended).
- **Internet Connection**: Requires .NET Framework 4.5.2 or later.

## Installation
1. **Download**:
   - Visit [https://ruben-apaza.blogspot.com/p/raxl-stat.html](https://ruben-apaza.blogspot.com/p/raxl-stat.html) or the GitHub repository.
   - Download the `raXL_Stat_v0.5.zip` file.
2. **Installation**:
   - The add-in does not require installation; simply open the `.xll` add-in and run it.
3. **Verification**:
   - A "raXL Stat" tab will appear in the toolbar, or functions will be available by typing `=ra.` in a cell.

## License and Activation
- **License Verification**: Use `=ra.raXLStat.License()` .
- **Supported Versions**: Check with `=ra.raXLStat.Version()`.

## Getting Started
1. **Accessing Functions**:
   - Open Excel, type `=ra.` in a cell to list UDFs, or use the "raXL Stat" tab.
2. **Data Preparation**:
   - Organize data in columns with clear headers (e.g., "Sales," "Date").
   - Check for missing or non-numeric data using `=ra.MissingData.Info(RangeX)`.
3. **Conventions**:
   - **RangeX**: Independent variables or data range (e.g., A1:B100).
   - **RangeY** or **RangeYt**: Dependent variable or time series (e.g., C1:C100).
   - **AscendentYt**: `TRUE` (most recent date at top), `FALSE` (most recent at bottom).
   - **Alpha**: Significance level (e.g., 0.05).
   - **Lag**: Lag (positive integer).
   - **ConstantC**: `TRUE` (include intercept), `FALSE` (exclude intercept), or fixed value.

raXL Stat functions can be run in three different ways: from the Excel tab or menu, from a user-defined function (UDF), or from VBA (Visual Basic for Applications) programming. For more details, you can check out the list of videos on how to use raXL Stat on our YouTube channel: https://www.youtube.com/watch?v=wYdGCkdN6cE&list=PLu4ltjreHhzO-cV1rHIis-K5_8numRqQV&pp=gAQBiAQB.
## Core Functionalities
The UDFs of raXL Stat are grouped by category. Each function is detailed with its **Function**, **Description**, and **Parameters**.

### 1. License and System
#### 1.1. ra.raXLStat.License
- **Function**: `ra.raXLStat.License(password)`
- **Description**: Verifies the add-in’s license status.
- **Parameters**:
  - `password`: String. Activation password.

#### 1.2. ra.raXLStat.Version
- **Function**: `ra.raXLStat.Version(password)`
- **Description**: Displays supported Excel versions.
- **Parameters**:
  - `password`: String. Activation password.

#### 1.3. ra.raXLStat.xIDCPU
- **Function**: `ra.raXLStat.xIDCPU(password)`
- **Description**: Returns the unique CPU identifier for activation.
- **Parameters**:
  - `password`: String. Activation password.

### 2. Descriptive Statistics
#### 2.1. ra.Descriptive.Stats
- **Function**: `ra.Descriptive.Stats(RangeX)`
- **Description**: Computes descriptive statistics (mean, median, standard deviation, min, max).
- **Parameters**:
  - `RangeX`: Data range (e.g., A1:A100).

#### 2.2. ra.Correlation.Matrix
- **Function**: `ra.Correlation.Matrix(RangeX, Population)`
- **Description**: Returns the correlation coefficient matrix of multiple data ranges, random variables, or time series.
- **Parameters**:
  - `RangeX`: Range of variables (e.g., A1:C100).
  - `Population`: Boolean. `TRUE` for population correlation, `FALSE` for sample.

#### 2.3. ra.Correlation.Matrix.Test
- **Function**: `ra.Correlation.Matrix.Test(RangeX, LowTriang, Alpha)`
- **Description**: Generates a Pearson correlation matrix with options for coefficients, R-squared, p-values, or significance tests.
- **Parameters**:
  - `RangeX`: Multi-column data range (e.g., A1:C100).
  - `LowTriang`: Integer. 0 (correlation coefficients R), 1 (R-squared), 2 (p-values for R based on t-Stat), 3 (test for R=0).
  - `Alpha`: Double. Significance level (default = 0.05).

#### 2.4. ra.Correlation.Vector
- **Function**: `ra.Correlation.Vector(RangeX, RangeY)`
- **Description**: Returns the correlation coefficients between a dependent variable and multiple independent variables.
- **Parameters**:
  - `RangeX`: Independent variables (e.g., A1:B100).
  - `RangeY`: Dependent variable (e.g., C1:C100).

#### 2.5. ra.Covariance.Matrix
- **Function**: `ra.Covariance.Matrix(RangeX, Population)`
- **Description**: Returns the covariance matrix for multiple data ranges, random variables, or time series.
- **Parameters**:
  - `RangeX`: Multi-column data range (e.g., A1:C100).
  - `Population`: Boolean. `TRUE` for population covariance, `FALSE` for sample.

#### 2.6. ra.Means.Vector
- **Function**: `ra.Means.Vector(RangeX)`
- **Description**: Returns a vector of means for multiple data sets.
- **Parameters**:
  - `RangeX`: Single or multi-column data range (e.g., A1:C100).

#### 2.7. ra.StdDev.Vector
- **Function**: `ra.StdDev.Vector(MatrixCovariance)`
- **Description**: Returns a vector of standard deviations from a covariance matrix.
- **Parameters**:
  - `MatrixCovariance`: Covariance matrix of returns.

### 3. Time Series Analysis [^3]
#### 3.1. ra.Autocorrelation.ACF
- **Function**: `ra.Autocorrelation.ACF(RangeYt, Lag, Interval, Alpha)`
- **Description**: Calculates the autocorrelation function (ACF) for a specified lag k.
- **Parameters**:
  - `RangeYt`: Time series (e.g., A1:A100).
  - `Lag`: Integer. Lag k (positive, default = 0).
  - `Interval`: Boolean. `TRUE` for confidence interval, `FALSE` (default).
  - `Alpha`: Double. Significance level (default = 0.05).

#### 3.2. ra.Autocovariance.Matrix
- **Function**: `ra.Autocovariance.Matrix(RangeYt, MaxLag)`
- **Description**: Returns the autocovariance function (ACVF) matrix for a specified maximum lag k.
- **Parameters**:
  - `RangeYt`: Time series (e.g., A1:A100).
  - `MaxLag`: Integer. Maximum lag k (positive).

#### 3.3. ra.Partial.Autocorr.PACF
- **Function**: `ra.Partial.Autocorr.PACF(RangeYt, Lag, Interval, Alpha)`
- **Description**: Calculates the partial autocorrelation function (PACF) for a specified lag k.
- **Parameters**:
  - `RangeYt`: Time series (e.g., A1:A100).
  - `Lag`: Integer. Lag k (positive, default = 0).
  - `Interval`: Boolean. `TRUE` for confidence interval, `FALSE` (default).
  - `Alpha`: Double. Significance level (default = 0.05).

#### 3.4. ra.LjungBox.Test
- **Function**: `ra.LjungBox.Test(RangeYt, Lag, AutocorrType, Alpha)`
- **Description**: Performs the Ljung-Box test for significant autocorrelation.
- **Parameters**:
  - `RangeYt`: Time series (e.g., A1:A100).
  - `Lag`: Integer. Number of lags (positive, default = 1).
  - `AutocorrType`: Boolean. `TRUE` for ACF (default), `FALSE` for PACF.
  - `Alpha`: Double. Significance level (default = 0.05).

#### 3.5. ra.BoxPierce.Test
- **Function**: `ra.BoxPierce.Test(RangeYt, Lag, AutocorrType, Alpha)`
- **Description**: Performs the Box-Pierce test for autocorrelation.
- **Parameters**:
  - `RangeYt`: Time series (e.g., A1:A100).
  - `Lag`: Integer. Number of lags (positive, default = 1).
  - `AutocorrType`: Boolean. `TRUE` for ACF (default), `FALSE` for PACF.
  - `Alpha`: Double. Significance level (default = 0.05).

#### 3.6. ra.DickeyFuller.ADF.Test
- **Function**: `ra.DickeyFuller.ADF.Test(RangeYt, AscendentYt, Lag, TermType, Alpha)`
- **Description**: Runs the Augmented Dickey-Fuller (ADF) test for stationarity.
- **Parameters**:
  - `RangeYt`: Time series (e.g., A1:A100).
  - `AscendentYt`: Boolean. `TRUE` for ascending order, `FALSE` (default).
  - `Lag`: Integer. Number of lags (positive, default = 1).
  - `TermType`: Integer. 0 (no constant), 1 (constant, default), 2 (constant and trend).
  - `Alpha`: Double. Significance level (default = 0.05).

#### 3.7. ra.DickeyFuller.ADF.Reg
- **Function**: `ra.DickeyFuller.ADF.Reg(RangeYt, AscendentYt, Lag, TermType)`
- **Description**: Runs the autoregression for the Augmented Dickey-Fuller (ADF) test for stationarity.
- **Parameters**:
  - `RangeYt`: Time series (e.g., A1:A100).
  - `AscendentYt`: Boolean. `TRUE` for ascending order, `FALSE` (default).
  - `Lag`: Integer. Number of lags (positive, default = 1).
  - `TermType`: Integer. 0 (no constant), 1 (constant, default), 2 (constant and trend).

#### 3.8. ra.KPSS.Test
- **Function**: `ra.KPSS.Test(RangeYt, AscendentYt, Lag, TermType, Alpha)`
- **Description**: Performs the KPSS test for stationarity.
- **Parameters**:
  - `RangeYt`: Time series (e.g., A1:A100).
  - `AscendentYt`: Boolean. `TRUE` for ascending order, `FALSE` (default).
  - `Lag`: Integer. Number of lags (positive, default = 1).
  - `TermType`: Integer. 1 (constant, default), 2 (constant and trend).
  - `Alpha`: Double. Significance level (default = 0.05).

#### 3.9. ra.KPSS.Reg
- **Function**: `ra.KPSS.Reg(RangeYt, AscendentYt, Lag, TermType)`
- **Description**: Performs the autoregression for the KPSS test for stationarity.
- **Parameters**:
  - `RangeYt`: Time series (e.g., A1:A100).
  - `AscendentYt`: Boolean. `TRUE` for ascending order, `FALSE` (default).
  - `Lag`: Integer. Number of lags (positive, default = 1).
  - `TermType`: Integer. 1 (constant, default), 2 (constant and trend).

#### 3.10. ra.ARIMA.Coeff
- **Function**: `ra.ARIMA.Coeff(RangeYt, AscendentYt, ARp, DiffD, MAq, OptMethod)`
- **Description**: Estimates coefficients for an ARIMA(p,d,q) model.
- **Parameters**:
  - `RangeYt`: Time series (e.g., A1:A100).
  - `AscendentYt`: Boolean. `TRUE` for ascending order, `FALSE` (default).
  - `ARp`: Integer. AR order (default = 1).
  - `DiffD`: Integer. Differencing order (default = 0).
  - `MAq`: Integer. MA order (default = 1).
  - `OptMethod`: String. Estimation method, default "NR" (Newton-Raphson).

#### 3.11. ra.ARMA.Fitted
- **Function**: `ra.ARMA.Fitted(RangeYt, AscendentYt, Constant, RangeAR, RangeMA)`
- **Description**: Returns fitted values of an ARMA model.
- **Parameters**:
  - `RangeYt`: Time series (e.g., A1:A100).
  - `AscendentYt`: Boolean. `TRUE` for ascending order, `FALSE` (default).
  - `Constant`: Double. Model constant.
  - `RangeAR`: Range of AR (phi) coefficients.
  - `RangeMA`: Range of MA (theta) coefficients.

#### 3.12. ra.ARMA.Forecast
- **Function**: `ra.ARMA.Forecast(RangeYt, AscendentYt, Constant, RangeAR, RangeMA, Forecast, Interval, Iterations)`
- **Description**: Forecasts future values using an ARMA model.
- **Parameters**:
  - `RangeYt`: Time series (e.g., A1:A100).
  - `AscendentYt`: Boolean. `TRUE` for ascending order, `FALSE` (default).
  - `Constant`: Double. Model constant.
  - `RangeAR`: Range of AR (phi) coefficients.
  - `RangeMA`: Range of MA (theta) coefficients.
  - `Forecast`: Integer. Number of periods to forecast (default = 10).
  - `Interval`: Integer. 0 (no interval, default), 1 (68%), 2 (95%), 3 (99.7%).
  - `Iterations`: Integer. Simulation iterations (default = 1000).

#### 3.13. ra.GARCH.Coeff
- **Function**: `ra.GARCH.Coeff(RangeYt, AscendentYt, AlphaP, BetaQ, CondMean, ErrorDist, OptMethod)`
- **Description**: Estimates coefficients for a GARCH(p,q) model.
- **Parameters**:
  - `RangeYt`: Time series (e.g., A1:A100).
  - `AscendentYt`: Boolean. `TRUE` for ascending order, `FALSE` (default).
  - `AlphaP`: Integer. ARCH order (default = 1).
  - `BetaQ`: Integer. GARCH order (default = 1).
  - `CondMean`: Boolean. `TRUE` to include conditional mean, `FALSE` (default).
  - `ErrorDist`: String. Error distribution, default "Normal".
  - `OptMethod`: String. Estimation method, default "NR" (Newton-Raphson).

#### 3.14. ra.GARCH.Fitted
- **Function**: `ra.GARCH.Fitted(RangeYt, AscendentYt, RangeAlpha, RangeBeta)`
- **Description**: Returns fitted values for a GARCH(p,q) model.
- **Parameters**:
  - `RangeYt`: Time series (e.g., A1:A100).
  - `AscendentYt`: Boolean. `TRUE` for ascending order, `FALSE` (default).
  - `RangeAlpha`: Range of ARCH (Alpha) coefficients.
  - `RangeBeta`: Range of GARCH (Beta) coefficients.

#### 3.15. ra.GARCH.Forecast
- **Function**: `ra.GARCH.Forecast(RangeYt, AscendentYt, RangeAlpha, RangeBeta, Forecast, Interval, Iterations)`
- **Description**: Forecasts volatility using a GARCH(p,q) model.
- **Parameters**:
  - `RangeYt`: Time series (e.g., A1:A100).
  - `AscendentYt`: Boolean. `TRUE` for ascending order, `FALSE` (default).
  - `RangeAlpha`: Range of ARCH (Alpha) coefficients.
  - `RangeBeta`: Range of GARCH (Beta) coefficients.
  - `Forecast`: Integer. Number of periods to forecast (default = 10).
  - `Interval`: Integer. 0 (no interval), 1 (68%), 2 (95%, default), 3 (99.7%).
  - `Iterations`: Integer. Simulation iterations (default = 1000).

#### 3.16. ra.GBM.Brownian.Coeff
- **Function**: `ra.GBM.Brownian.Coeff(RangeYt, AscendentYt)`
- **Description**: Calculates parameters (mean, standard deviation) for a Geometric Brownian Motion (GBM).
- **Parameters**:
  - `RangeYt`: Time series (e.g., A1:A100).
  - `AscendentYt`: Boolean. `TRUE` for ascending order, `FALSE` (default).

#### 3.17. ra.GBM.Brownian.Forecast
- **Function**: `ra.GBM.Brownian.Forecast(RangeMuSigma, Initial, Iterations, Forecast, Seed, Interval)`
- **Description**: Forecasts using Geometric Brownian Motion (GBM) with Monte Carlo simulation.
- **Parameters**:
  - `RangeMuSigma`: Range with mean and standard deviation.
  - `Initial`: Double. Initial value.
  - `Iterations`: Integer. Number of simulations (default = 100).
  - `Forecast`: Integer. Periods to forecast (default = 15).
  - `Seed`: Integer. Random seed (default = 1234).
  - `Interval`: Integer. 0 (simple random), 1 (68%), 2 (95%, default), 3 (99.7%).

#### 3.18. ra.FanChart.Table
- **Function**: `ra.FanChart.Table(RangeY, IntervalType)`
- **Description**: Returns columns for fan chart, trend band, or lines interval plot.
- **Parameters**:
  - `RangeY`: Multi-column data range (e.g., A1:C100).
  - `IntervalType`: Integer. 0 (mean, no interval), 1 (trend band 68%-95%), 2 (fan chart 68%-95%, default).

### 4. Normality Tests
#### 4.1. ra.JarqueBera.Test
- **Function**: `ra.JarqueBera.Test(RangeX, Population, Alpha)`
- **Description**: Performs the Jarque-Bera test for normality.
- **Parameters**:
  - `RangeX`: Data range (e.g., A1:A100).
  - `Population`: Boolean. `TRUE` for population (default), `FALSE` for sample.
  - `Alpha`: Double. Significance level (default = 0.05).

#### 4.2. ra.ShapiroWilk.Test
- **Function**: `ra.ShapiroWilk.Test(RangeX, Alpha)`
- **Description**: Performs the Shapiro-Wilk test for normality using the Royston algorithm (2 to 5000 observations).
- **Parameters**:
  - `RangeX`: Data range (e.g., A1:A100).
  - `Alpha`: Double. Significance level (default = 0.05).

#### 4.3. ra.AndersonDarling.Test
- **Function**: `ra.AndersonDarling.Test(RangeX, DistType, Alpha)`
- **Description**: Performs the Anderson-Darling test for normality.
- **Parameters**:
  - `RangeX`: Data range (e.g., A1:A100).
  - `DistType`: Integer. 0 (generic), 1 (normal, default), 2 (unmodified normal), 3 (lognormal).
  - `Alpha`: Double. Significance level (default = 0.05).

#### 4.4. ra.Skew.S
- **Function**: `ra.Skew.S(RangeX)`
- **Description**: Returns the skewness value of the distribution based on the sample.
- **Parameters**:
  - `RangeX`: Data range (e.g., A1:A100).

#### 4.5. ra.Kurt.S
- **Function**: `ra.Kurt.S(RangeX)`
- **Description**: Returns the kurtosis value of the distribution based on the sample.
- **Parameters**:
  - `RangeX`: Data range (e.g., A1:A100).

#### 4.6. ra.Skew.P
- **Function**: `ra.Skew.P(RangeX)`
- **Description**: Returns the skewness value of the distribution based on the population.
- **Parameters**:
  - `RangeX`: Data range (e.g., A1:A100).

#### 4.7. ra.Kurt.P
- **Function**: `ra.Kurt.P(RangeX, Subtract3)`
- **Description**: Returns the kurtosis value of the distribution based on the population.
- **Parameters**:
  - `RangeX`: Data range (e.g., A1:A100).
  - `Subtract3`: Boolean. `TRUE` to subtract 3 (normal distribution kurtosis = 0), `FALSE` (no subtraction).

### 5. Linear Regression (OLS)
#### 5.1. ra.LinearReg.Coeff
- **Function**: `ra.LinearReg.Coeff(RangeX, RangeY, ConstantC, RobustStdErrHC, Alpha)`
- **Description**: Estimates coefficients, standard errors, t-statistics, p-values, and confidence intervals.
- **Parameters**:
  - `RangeX`: Independent variables (e.g., A1:B100).
  - `RangeY`: Dependent variable (e.g., C1:C100).
  - `ConstantC`: Boolean or Double. `TRUE` (intercept, default), `FALSE` (no intercept), or fixed value.
  - `RobustStdErrHC`: Integer. 0=HC0, 1=HC1, 2=HC2, 3=HC3, 4=HC4, 5=HAC, 6=OLS (default).
  - `Alpha`: Double. Significance level (default = 0.05).

#### 5.2. ra.LinearReg.Fitted
- **Function**: `ra.LinearReg.Fitted(RangeX, RangeY, ConstantC, Interval, Alpha)`
- **Description**: Returns fitted values and optional confidence or prediction intervals.
- **Parameters**:
  - `RangeX`: Independent variables (e.g., A1:B100).
  - `RangeY`: Dependent variable (e.g., C1:C100).
  - `ConstantC`: Boolean or Double. `TRUE` (intercept, default), `FALSE` (no intercept), or fixed value.
  - `Interval`: Integer. 0 (no interval, default), 1 (confidence), 2 (prediction).
  - `Alpha`: Double. Significance level (default = 0.05).

#### 5.3. ra.LinearReg.Forecast
- **Function**: `ra.LinearReg.Forecast(RangeX, RangeY, RangeXo, ConstantC, Interval, Alpha)`
- **Description**: Generates forecasts for new data with optional intervals.
- **Parameters**:
  - `RangeX`: Independent variables (e.g., A1:B100).
  - `RangeY`: Dependent variable (e.g., C1:C100).
  - `RangeXo`: New independent variable data (e.g., E1:F10).
  - `ConstantC`: Boolean or Double. `TRUE` (intercept, default), `FALSE` (no intercept), or fixed value.
  - `Interval`: Boolean. `TRUE` (include interval), `FALSE` (forecast only, default).
  - `Alpha`: Double. Significance level (default = 0.05).

#### 5.4. ra.LinearReg.Residuals
- **Function**: `ra.LinearReg.Residuals(RangeX, RangeY, ConstantC, ResidualType, Alpha, Critical)`
- **Description**: Computes residuals, standardized, studentized, Pearson, DFFITS, leverage, or Cook’s distance.
- **Parameters**:
  - `RangeX`: Independent variables (e.g., A1:B100).
  - `RangeY`: Dependent variable (e.g., C1:C100).
  - `ConstantC`: Boolean or Double. `TRUE` (intercept, default), `FALSE` (no intercept), or fixed value.
  - `ResidualType`: Integer. 0=residuals, 1=σ, 2=standardized, 3=studentized, 4=Pearson, 5=DFFITS, 6=leverage, 7=Cook’s distance.
  - `Alpha`: Double. Significance level for z-Stat and t-Stat (default = 0.1587).
  - `Critical`: Double. Critical value for Cook’s (4/n), Leverage (2*mu), DFFITS (1).

#### 5.5. ra.LinearReg.Residuals.2
- **Function**: `ra.LinearReg.Residuals.2(RangeX, RangeY, ConstantC, MeasureType, Alpha, Critical)`
- **Description**: Generates a table for XY scatter plot and computes diagnostic measures for outlier detection.
- **Parameters**:
  - `RangeX`: Independent variables (e.g., A1:B100).
  - `RangeY`: Dependent variable (e.g., C1:C100).
  - `ConstantC`: Boolean or Double. `TRUE` (intercept, default), `FALSE` (no intercept), or fixed value.
  - `MeasureType`: Integer. 0=residuals-fitted, 1=standardized-fitted, 2=studentized-fitted, 3=standardized-leverage, 4=studentized-leverage, 5=studentized-Cook, 6=studentized-DFFITS.
  - `Alpha`: Double. Significance level for z-Stat and t-Stat (default = 0.05).
  - `Critical`: Double. Critical value for Cook’s (4/n), Leverage (2*mu), DFFITS (1).

#### 5.6. ra.LinearReg.Influence
- **Function**: `ra.LinearReg.Influence(RangeX, RangeY, readjust)`
- **Description**: Calculates Pratt’s measure for the relative importance of variables in R².
- **Parameters**:
  - `RangeX`: Independent variables (e.g., A1:B100).
  - `RangeY`: Dependent variable (e.g., C1:C100).
  - `readjust`: Boolean. `TRUE` to adjust indices to 100% (default), `FALSE` (no adjustment).

#### 5.7. ra.LinearReg.Anova
- **Function**: `ra.LinearReg.Anova(RangeX, RangeY, ConstantC)`
- **Description**: Generates an ANOVA table with sums of squares, mean squares, F-statistic, and p-value.
- **Parameters**:
  - `RangeX`: Independent variables (e.g., A1:B100).
  - `RangeY`: Dependent variable (e.g., C1:C100).
  - `ConstantC`: Boolean or Double. `TRUE` (intercept, default), `FALSE` (no intercept), or fixed value.

#### 5.8. ra.LinearReg.Stat
- **Function**: `ra.LinearReg.Stat(RangeX, RangeY, ConstantC)`
- **Description**: Computes fit metrics like R², adjusted R², standard error, Durbin-Watson, LLH, AIC, and BIC.
- **Parameters**:
  - `RangeX`: Independent variables (e.g., A1:B100).
  - `RangeY`: Dependent variable (e.g., C1:C100).
  - `ConstantC`: Boolean or Double. `TRUE` (intercept, default), `FALSE` (no intercept), or fixed value.

#### 5.9. ra.RamseyRESET.Test
- **Function**: `ra.RamseyRESET.Test(RangeX, RangeY, ConstantC, Power, Alpha)`
- **Description**: Performs the Ramsey RESET test to check for misspecification by testing if polynomial terms of fitted values improve the model.
- **Parameters**:
  - `RangeX`: Independent variables (e.g., A1:B100).
  - `RangeY`: Dependent variable (e.g., C1:C100).
  - `ConstantC`: Boolean or Double. `TRUE` (intercept, default), `FALSE` (no intercept), or fixed value.
  - `Power`: Integer. Maximum polynomial order of fitted values (default = 1 for quadratic).
  - `Alpha`: Double. Significance level (default = 0.05).

#### 5.10. ra.RamseyRESET.Reg
- **Function**: `ra.RamseyRESET.Reg(RangeX, RangeY, ConstantC, Power)`
- **Description**: Performs the regression for the Ramsey RESET test.
- **Parameters**:
  - `RangeX`: Independent variables (e.g., A1:B100).
  - `RangeY`: Dependent variable (e.g., C1:C100).
  - `ConstantC`: Boolean or Double. `TRUE` (intercept, default), `FALSE` (no intercept), or fixed value.
  - `Power`: Integer. Maximum polynomial order of fitted values (default = 1 for quadratic).

#### 5.11. ra.RecursiveCUSUM
- **Function**: `ra.RecursiveCUSUM(RangeX, RangeY, ConstantC, recursiveType)`
- **Description**: Performs the Recursive CUSUM test to detect structural breaks or parameter instability using recursive residuals.
- **Parameters**:
  - `RangeX`: Independent variables (e.g., A1:B100).
  - `RangeY`: Dependent variable (e.g., C1:C100).
  - `ConstantC`: Boolean or Double. `TRUE` (intercept, default), `FALSE` (no intercept), or fixed value.
  - `recursiveType`: Integer. 0 (CUSUM + 5% interval, default), 1 (recursive residuals), 2 (standardized recursive residuals).

#### 5.12. ra.Acorr.BreuschGodfrey.Test
- **Function**: `ra.Acorr.BreuschGodfrey.Test(RangeX, RangeY, ConstantC, Lag, Alpha, Chi2)`
- **Description**: Performs the Breusch-Godfrey test to detect autocorrelation in regression residuals up to a specified lag.
- **Parameters**:
  - `RangeX`: Independent variables (e.g., A1:B100).
  - `RangeY`: Dependent variable (e.g., C1:C100).
  - `ConstantC`: Boolean or Double. `TRUE` (intercept, default), `FALSE` (no intercept), or fixed value.
  - `Lag`: Integer. Number of lags to test (default = 2).
  - `Alpha`: Double. Significance level (default = 0.05).
  - `Chi2`: Boolean. `TRUE` for Chi-squared test (default), `FALSE` for F-Stat.

#### 5.13. ra.Acorr.BreuschGodfrey.Reg
- **Function**: `ra.Acorr.BreuschGodfrey.Reg(RangeX, RangeY, ConstantC, Lag)`
- **Description**: Performs the regression for the Breusch-Godfrey test.
- **Parameters**:
  - `RangeX`: Independent variables (e.g., A1:B100).
  - `RangeY`: Dependent variable (e.g., C1:C100).
  - `ConstantC`: Boolean or Double. `TRUE` (intercept, default), `FALSE` (no intercept), or fixed value.
  - `Lag`: Integer. Number of lags to test (default = 2).

#### 5.14. ra.Acorr.ACF.Test
- **Function**: `ra.Acorr.ACF.Test(RangeYt, Lag, Alpha)`
- **Description**: Tests for significant autocorrelation in residuals using the autocorrelation function (ACF) up to a specified lag.
- **Parameters**:
  - `RangeYt`: Time series or residuals (e.g., A1:A100).
  - `Lag`: Integer. Number of lags to test (default = 1).
  - `Alpha`: Double. Significance level (default = 0.05).

#### 5.15. ra.Acorr.DurbinWatson.Stat
- **Function**: `ra.Acorr.DurbinWatson.Stat(RangeYt)`
- **Description**: Computes the Durbin-Watson statistic to test for first-order autocorrelation in residuals.
- **Parameters**:
  - `RangeYt`: Time series or residuals (e.g., A1:A100).

#### 5.16. ra.MulticolLin.VIF
- **Function**: `ra.MulticolLin.VIF(RangeX, Column, measureType)`
- **Description**: Calculates the Variance Inflation Factor (VIF) for each independent variable to assess multicollinearity.
- **Parameters**:
  - `RangeX`: Multi-column data range (e.g., A1:B100).
  - `Column`: Integer. Column index of matrix X (default = 0 for all columns).
  - `measureType`: Integer. 0 (VIF, default), 1 (Tolerance), 2 (R-squared).

#### 5.17. ra.MulticolLin.Lambda
- **Function**: `ra.MulticolLin.Lambda(RangeX)`
- **Description**: Computes the eigenvalues (λ) of the correlation matrix of independent variables to evaluate multicollinearity.
- **Parameters**:
  - `RangeX`: Multi-column data range (e.g., A1:B100).

#### 5.18. ra.MulticolLin.Kappa
- **Function**: `ra.MulticolLin.Kappa(RangeX)`
- **Description**: Calculates the condition number (Kappa) of the independent variables’ matrix to detect multicollinearity.
- **Parameters**:
  - `RangeX`: Multi-column data range (e.g., A1:B100).

#### 5.19. ra.MulticolLin.Nu
- **Function**: `ra.MulticolLin.Nu(RangeX)`
- **Description**: Returns variances associated with principal axes (loadings) or eigenvectors (υ) of the correlation matrix.
- **Parameters**:
  - `RangeX`: Multi-column data range (e.g., A1:B100).

#### 5.20. ra.Leverage
- **Function**: `ra.Leverage(RangeX, Constant, MeasureType)`
- **Description**: Computes leverage values for each observation to identify influential points in independent variables.
- **Parameters**:
  - `RangeX`: Multi-column data range (e.g., A1:B100).
  - `Constant`: Boolean. `TRUE` (include intercept, default), `FALSE` (no intercept).
  - `MeasureType`: Integer. 0 (leverage, default), 1 (leverage + 2*mu), 2 (leverage + 2*mu outliers).

#### 5.21. ra.Normalized.Distances
- **Function**: `ra.Normalized.Distances(RangeX, Alpha, MeasureType)`
- **Description**: Calculates normalized distances for observations to detect outliers in independent variables.
- **Parameters**:
  - `RangeX`: Multi-column data range (e.g., A1:B100).
  - `Alpha`: Double. Significance level (default = 0.01).
  - `MeasureType`: Integer. 0 (normalized distances + unit, default), 1 (normalized distances + outliers).

### 6. Heteroskedasticity Tests
#### 6.1. ra.ARCH.Test
- **Function**: `ra.ARCH.Test(RangeYt, AscendentYt, Lag, Alpha, Chi2)`
- **Description**: Performs Engle’s LM Test for ARCH heteroskedasticity.
- **Parameters**:
  - `RangeYt`: Residuals (e.g., A1:A100).
  - `AscendentYt`: Boolean. `TRUE` for ascending order, `FALSE` (default).
  - `Lag`: Integer. Number of lags (default = 1).
  - `Alpha`: Double. Significance level (default = 0.05).
  - `Chi2`: Boolean. `TRUE` for Chi-squared test (default), `FALSE` for F-Stat.

#### 6.2. ra.ARCH.Test.Reg
- **Function**: `ra.ARCH.Test.Reg(RangeYt, AscendentYt, Lag)`
- **Description**: Performs the autoregression for Engle’s LM Test for ARCH.
- **Parameters**:
  - `RangeYt`: Residuals (e.g., A1:A100).
  - `AscendentYt`: Boolean. `TRUE` for ascending order, `FALSE` (default).
  - `Lag`: Integer. Number of lags (default = 1).

#### 6.3. ra.Het.BreuschPagan.Test
- **Function**: `ra.Het.BreuschPagan.Test(RangeX, RangeY, ConstantC, Alpha, Chi2)`
- **Description**: Performs the Breusch-Pagan-Godfrey test for homoskedasticity.
- **Parameters**:
  - `RangeX`: Independent variables (e.g., A1:B100).
  - `RangeY`: Dependent variable (e.g., C1:C100).
  - `ConstantC`: Boolean or Double. `TRUE` (intercept, default), `FALSE` (no intercept), or fixed value.
  - `Alpha`: Double. Significance level (default = 0.05).
  - `Chi2`: Boolean. `TRUE` for Chi-squared test (default), `FALSE` for F-Stat.

#### 6.4. ra.Het.BreuschPagan.Reg
- **Function**: `ra.Het.BreuschPagan.Reg(RangeX, RangeY, ConstantC)`
- **Description**: Performs regression for the Breusch-Pagan-Godfrey test.
- **Parameters**:
  - `RangeX`: Independent variables (e.g., A1:B100).
  - `RangeY`: Dependent variable (e.g., C1:C100).
  - `ConstantC`: Boolean or Double. `TRUE` (intercept, default), `FALSE` (no intercept), or fixed value.

#### 6.5. ra.Het.White.Test
- **Function**: `ra.Het.White.Test(RangeX, RangeY, ConstantC, CrossTerms, Alpha, Chi2)`
- **Description**: Performs White’s heteroskedasticity test.
- **Parameters**:
  - `RangeX`: Independent variables (e.g., A1:B100).
  - `RangeY`: Dependent variable (e.g., C1:C100).
  - `ConstantC`: Boolean or Double. `TRUE` (intercept, default), `FALSE` (no intercept), or fixed value.
  - `CrossTerms`: Boolean. `TRUE` to include White cross terms, `FALSE` (default).
  - `Alpha`: Double. Significance level (default = 0.05).
  - `Chi2`: Boolean. `TRUE` for Chi-squared test (default), `FALSE` for F-Stat.

#### 6.6. ra.Het.White.Reg
- **Function**: `ra.Het.White.Reg(RangeX, RangeY, ConstantC, CrossTerms)`
- **Description**: Performs regression for White’s heteroskedasticity test.
- **Parameters**:
  - `RangeX`: Independent variables (e.g., A1:B100).
  - `RangeY`: Dependent variable (e.g., C1:C100).
  - `ConstantC`: Boolean or Double. `TRUE` (intercept, default), `FALSE` (no intercept), or fixed value.
  - `CrossTerms`: Boolean. `TRUE` to include White cross terms, `FALSE` (default).

### 7. Risk and Portfolio Analysis
#### 7.1. ra.Portfolio.Risk
- **Function**: `ra.Portfolio.Risk(MatrixCovariance, VectorWeights)`
- **Description**: Returns the expected risk (volatility or standard deviation) of a portfolio.
- **Parameters**:
  - `MatrixCovariance`: Covariance matrix of returns.
  - `VectorWeights`: Vector of portfolio weights.

#### 7.2. ra.Portfolio.Return
- **Function**: `ra.Portfolio.Return(VectorMeanReturns, VectorWeights)`
- **Description**: Returns the expected return of a portfolio.
- **Parameters**:
  - `VectorMeanReturns`: Vector of mean asset returns.
  - `VectorWeights`: Vector of portfolio weights.

#### 7.3. ra.Portfolio.Weights.Optimal
- **Function**: `ra.Portfolio.Weights.Optimal(VectorMeanReturns, MatrixCovariance, ExpectedReturn)`
- **Description**: Returns the optimal investment weights in a portfolio on the efficient frontier.
- **Parameters**:
  - `VectorMeanReturns`: Vector of mean asset returns.
  - `MatrixCovariance`: Covariance matrix of returns.
  - `ExpectedReturn`: Double. Required or expected return.

#### 7.4. ra.Portfolio.Weights.Tangency
- **Function**: `ra.Portfolio.Weights.Tangency(VectorMeanReturns, MatrixCovariance, RiskFree)`
- **Description**: Returns the tangency portfolio weights on the efficient frontier.
- **Parameters**:
  - `VectorMeanReturns`: Vector of mean asset returns.
  - `MatrixCovariance`: Covariance matrix of returns.
  - `RiskFree`: Double. Risk-free rate.

#### 7.5. ra.Portfolio.Market.Line.Effic
- **Function**: `ra.Portfolio.Market.Line.Effic(VectorMeanReturns, MatrixCovariance, RiskFree, FrontierPoints)`
- **Description**: Returns expected return and risk for the Capital Market Line (CML) in a portfolio on the efficient frontier.
- **Parameters**:
  - `VectorMeanReturns`: Vector of mean asset returns.
  - `MatrixCovariance`: Covariance matrix of returns.
  - `RiskFree`: Double. Risk-free rate.
  - `FrontierPoints`: Integer. Number of points on the efficient frontier.

#### 7.6. ra.Portfolio.Market.Line.Opti
- **Function**: `ra.Portfolio.Market.Line.Opti(VectorMeanReturns, MatrixCovariance, RiskFree, ExpectedReturn)`
- **Description**: Returns the expected risk for the Capital Market Line (CML) optimal portfolio on the efficient frontier.
- **Parameters**:
  - `VectorMeanReturns`: Vector of mean asset returns.
  - `MatrixCovariance`: Covariance matrix of returns.
  - `RiskFree`: Double. Risk-free rate.
  - `ExpectedReturn`: Double. Target expected return.

#### 7.7. ra.Portfolio.Weights.MinVar
- **Function**: `ra.Portfolio.Weights.MinVar(VectorMeanReturns, MatrixCovariance)`
- **Description**: Returns the minimum variance portfolio weights on the efficient frontier.
- **Parameters**:
  - `VectorMeanReturns`: Vector of mean asset returns.
  - `MatrixCovariance`: Covariance matrix of returns.

#### 7.8. ra.Portfolio.Frontier.Efficient
- **Function**: `ra.Portfolio.Frontier.Efficient(VectorMeanReturns, MatrixCovariance, FrontierPoints)`
- **Description**: Returns expected return and risk for the efficient portfolio frontier.
- **Parameters**:
  - `VectorMeanReturns`: Vector of mean asset returns.
  - `MatrixCovariance`: Covariance matrix of returns.
  - `FrontierPoints`: Integer. Number of points on the efficient frontier.

#### 7.9. ra.Portfolio.Frontier.Optimal
- **Function**: `ra.Portfolio.Frontier.Optimal(VectorMeanReturns, MatrixCovariance, FrontierPoints)`
- **Description**: Returns the optimal expected return and risk for a portfolio on the efficient frontier.
- **Parameters**:
  - `VectorMeanReturns`: Vector of mean asset returns.
  - `MatrixCovariance`: Covariance matrix of returns.
  - `FrontierPoints`: Integer. Number of points on the efficient frontier.

#### 7.10. ra.Portfolio.Risk.Optimal
- **Function**: `ra.Portfolio.Risk.Optimal(VectorMeanReturns, MatrixCovariance, ExpectedReturn)`
- **Description**: Returns the optimal expected risk (standard deviation) for a portfolio on the efficient frontier.
- **Parameters**:
  - `VectorMeanReturns`: Vector of mean asset returns.
  - `MatrixCovariance`: Covariance matrix of returns.
  - `ExpectedReturn`: Double. Target expected return.

#### 7.11. ra.Portfolio.Simulation
- **Function**: `ra.Portfolio.Simulation(VectorMeanReturns, MatrixCovariance, IterSamples, Seed)`
- **Description**: Returns risk and return points for a portfolio simulation using Monte Carlo methods.
- **Parameters**:
  - `VectorMeanReturns`: Vector of mean asset returns.
  - `MatrixCovariance`: Covariance matrix of returns.
  - `IterSamples`: Integer. Number of simulation iterations.
  - `Seed`: Integer. Random seed (e.g., 1234).

#### 7.12. ra.Portfolio.Simul.Frontier
- **Function**: `ra.Portfolio.Simul.Frontier(VectorMeanReturns, MatrixCovariance, IterSamples, Seed, LevelZoom)`
- **Description**: Returns efficient risk and return points for a portfolio simulation (experimental).
- **Parameters**:
  - `VectorMeanReturns`: Vector of mean asset returns.
  - `MatrixCovariance`: Covariance matrix of returns.
  - `IterSamples`: Integer. Number of simulation iterations.
  - `Seed`: Integer. Random seed (e.g., 1234).
  - `LevelZoom`: Integer. Zoom level for frontier points (e.g., 1, 2).

#### 7.13. ra.VaR.Historical
- **Function**: `ra.VaR.Historical(RangeReturns, Alpha, Excludes)`
- **Description**: Returns the historical Value at Risk (H VaR) of a portfolio.
- **Parameters**:
  - `RangeReturns`: Single or multi-column data range of returns (e.g., A1:C100).
  - `Alpha`: Double. Significance level (e.g., 0.01, 0.05, 0.10).
  - `Excludes`: Boolean. `TRUE` to exclude first and last percentile values, `FALSE` to include.

#### 7.14. ra.VaR.Parametric
- **Function**: `ra.VaR.Parametric(RangeReturns, Alpha, Population)`
- **Description**: Returns the parametric Value at Risk (P VaR) of a portfolio, assuming normality.
- **Parameters**:
  - `RangeReturns`: Multi-column data range of returns (e.g., A1:C100).
  - `Alpha`: Double. Significance level (e.g., 0.01, 0.05, 0.10).
  - `Population`: Boolean. `TRUE` for population variance, `FALSE` for sample variance.

#### 7.15. ra.VaR.VarCovar
- **Function**: `ra.VaR.VarCovar(VectorMeanReturns, MatrixCovariance, VectorWeights, Alpha)`
- **Description**: Returns the variance-covariance Value at Risk (P VaR) of a portfolio, assuming normality.
- **Parameters**:
  - `VectorMeanReturns`: Vector of mean asset returns.
  - `MatrixCovariance`: Covariance matrix of returns.
  - `VectorWeights`: Vector of portfolio weights.
  - `Alpha`: Double. Significance level (e.g., 0.01, 0.05, 0.10).

#### 7.16. ra.CVaR.Historical
- **Function**: `ra.CVaR.Historical(RangeReturns, Alpha, Excludes)`
- **Description**: Returns the historical Conditional Value at Risk (H CVaR) or expected shortfall of a portfolio.
- **Parameters**:
  - `RangeReturns`: Single or multi-column data range of returns (e.g., A1:C100).
  - `Alpha`: Double. Significance level (e.g., 0.01, 0.05, 0.10).
  - `Excludes`: Boolean. `TRUE` to exclude first and last percentile values, `FALSE` to include.

#### 7.17. ra.CVaR.Parametric
- **Function**: `ra.CVaR.Parametric(RangeReturns, Alpha, Population)`
- **Description**: Returns the parametric Conditional Value at Risk (P CVaR) or expected shortfall, assuming normality.
- **Parameters**:
  - `RangeReturns`: Multi-column data range of returns (e.g., A1:C100).
  - `Alpha`: Double. Significance level (e.g., 0.01, 0.05, 0.10).
  - `Population`: Boolean. `TRUE` for population variance, `FALSE` for sample variance.

#### 7.18. ra.CVaR.VarCovar
- **Function**: `ra.CVaR.VarCovar(VectorMeanReturns, MatrixCovariance, VectorWeights, Alpha)`
- **Description**: Returns the variance-covariance Conditional Value at Risk (P CVaR) or expected shortfall, assuming normality.
- **Parameters**:
  - `VectorMeanReturns`: Vector of mean asset returns.
  - `MatrixCovariance`: Covariance matrix of returns.
  - `VectorWeights`: Vector of portfolio weights.
  - `Alpha`: Double. Significance level (e.g., 0.01, 0.05, 0.10).

#### 7.19. ra.VaR.Backtesting.Z
- **Function**: `ra.VaR.Backtesting.Z(RangeReturns, VaR, Alpha)`
- **Description**: Returns a backtesting vector of z-Stat, p-value, and calibration status for a VaR coefficient.
- **Parameters**:
  - `RangeReturns`: Column data range of returns (e.g., A1:A100).
  - `VaR`: Double. Value at Risk coefficient.
  - `Alpha`: Double. Significance level (e.g., 0.01, 0.05, 0.10).

#### 7.20. ra.VaR.Backtesting.Kupiec
- **Function**: `ra.VaR.Backtesting.Kupiec(RangeReturns, VaR, Alpha)`
- **Description**: Returns a backtesting vector of LR-Stat, p-value, and calibration status for a VaR coefficient.
- **Parameters**:
  - `RangeReturns`: Column data range of returns (e.g., A1:A100).
  - `VaR`: Double. Value at Risk coefficient.
  - `Alpha`: Double. Significance level (e.g., 0.01, 0.05, 0.10).

#### 7.21. ra.Means.Column
- **Function**: `ra.Means.Column(RangeReturns)`
- **Description**: Returns a column of means for a multi-column data set.
- **Parameters**:
  - `RangeReturns`: Multi-column data range (e.g., A1:C100).

### 8. Data Visualization
#### 8.1. ra.Histogram.Table
- **Function**: `ra.Histogram.Table(RangeX, CurveType, FrequencyType)`
- **Description**: Generates a table for histogram construction with automatic bins.
- **Parameters**:
  - `RangeX`: Single or multi-column data range (e.g., A1:C100).
  - `CurveType`: Integer. 0 (no curve, default), 1 (cumulative), 2 (normal).
  - `FrequencyType`: Boolean. `TRUE` for absolute frequency (default), `FALSE` for relative frequency.

#### 8.2. ra.Histogram.Table.Bin
- **Function**: `ra.Histogram.Table.Bin(RangeX, CurveType, FrequencyType, Start, Step, Stop)`
- **Description**: Generates a table for histogram construction with user-specified bins.
- **Parameters**:
  - `RangeX`: Single or multi-column data range (e.g., A1:C100).
  - `CurveType`: Integer. 0 (no curve, default), 1 (cumulative), 2 (normal).
  - `FrequencyType`: Boolean. `TRUE` for absolute frequency (default), `FALSE` for relative frequency.
  - `Start`: Double. Minimum bin value.
  - `Step`: Double. Step size between bins.
  - `Stop`: Double. Maximum bin value.

#### 8.3. ra.BoxPlot.Table
- **Function**: `ra.BoxPlot.Table(RangeX, Outliers, IQR)`
- **Description**: Generates a table for box-and-whisker plots.
- **Parameters**:
  - `RangeX`: Single or multi-column data range (e.g., A1:C100).
  - `Outliers`: Boolean. `TRUE` to detect outliers, `FALSE` (default).
  - `IQR`: Double. Interquartile range multiplier (default = 1.5).

#### 8.4. ra.Curve.Distribution
- **Function**: `ra.Curve.Distribution(RangeX, Population, Distribution)`
- **Description**: Generates a table of density or frequencies with automatic bins for plotting a curve.
- **Parameters**:
  - `RangeX`: Single or multi-column data range (e.g., A1:C100).
  - `Population`: Boolean. `TRUE` for population variance, `FALSE` for sample variance.
  - `Distribution`: String. Curve type, default "Normal" (Gaussian).

### 9. Utilities
#### 9.1. ra.MissingData.Info
- **Function**: `ra.MissingData.Info(RangeX)`
- **Description**: Identifies missing or non-numeric data in a range.
- **Parameters**:
  - `RangeX`: Single or multi-column data range (e.g., A1:C100).

#### 9.2. ra.Difference
- **Function**: `ra.Difference(RangeYt, AscendentYt, DifferenceD)`
- **Description**: Returns a column with the difference operation applied to a time series.
- **Parameters**:
  - `RangeYt`: Time series (e.g., A1:A100).
  - `AscendentYt`: Boolean. `TRUE` for ascending order, `FALSE` (default).
  - `DifferenceD`: Integer. Order of differencing (default = 1).

#### 9.3. ra.Interpolate.Point
- **Function**: `ra.Interpolate.Point(x, x1, x2, y1, y2, interpType)`
- **Description**: Performs linear, logarithmic, or harmonic interpolation.
- **Parameters**:
  - `x`: Double. Value to interpolate.
  - `x1`, `x2`: Double. Reference values.
  - `y1`, `y2`: Double. Corresponding values.
  - `interpType`: Integer. 0 (linear), 1 (logarithmic), 2 (harmonic).

#### 9.4. ra.Flip
- **Function**: `ra.Flip(RangeX, Flip)`
- **Description**: Flips a vertical range of cells if the condition is true.
- **Parameters**:
  - `RangeX`: Single or multi-column data range (e.g., A1:C100).
  - `Flip`: Boolean. `TRUE` to flip vertically (default), `FALSE` to duplicate.

#### 9.5. ra.Transpose
- **Function**: `ra.Transpose(RangeX, Transpose)`
- **Description**: Transposes a range from vertical to horizontal or vice versa.
- **Parameters**:
  - `RangeX`: Single or multi-column data range (e.g., A1:C100).
  - `Transpose`: Boolean. `TRUE` to transpose (default), `FALSE` to duplicate.

#### 9.6. ra.ShowLag
- **Function**: `ra.ShowLag(RangeYt, AscendentYt, ShowLag)`
- **Description**: Shows a column with lagged values of a time series.
- **Parameters**:
  - `RangeYt`: Time series (e.g., A1:A100).
  - `AscendentYt`: Boolean. `TRUE` for ascending order, `FALSE` (default).
  - `ShowLag`: Integer. Lag value (positive, default = all lags).

#### 9.7. ra.Rate.LN
- **Function**: `ra.Rate.LN(RangeYt, AscendentYt)`
- **Description**: Converts a positive time series into a time series of natural logarithmic (LN) rate of change.
- **Parameters**:
  - `RangeYt`: Time series with positive values (e.g., A1:A100).
  - `AscendentYt`: Boolean. `TRUE` for ascending order, `FALSE` (default).

#### 9.8. ra.Rate.Log
- **Function**: `ra.Rate.Log(RangeYt, AscendentYt, Base)`
- **Description**: Converts a positive time series into a time series of logarithmic (LOG) rate of change.
- **Parameters**:
  - `RangeYt`: Time series (e.g., A1:A100).
  - `AscendentYt`: Boolean. `TRUE` for ascending order, `FALSE` (default).
  - `Base`: Integer. Logarithmic base (default = 10).

#### 9.9. ra.Rate.Change
- **Function**: `ra.Rate.Change(RangeYt, AscendentYt)`
- **Description**: Converts a positive time series into a time series of rate of change.
- **Parameters**:
  - `RangeYt`: Time series (e.g., A1:A100).
  - `AscendentYt`: Boolean. `TRUE` for ascending order, `FALSE` (default).

#### 9.10. ra.Standardize
- **Function**: `ra.Standardize(RangeX, standardized)`
- **Description**: Returns standardized (mean, sigma) values from a distribution.
- **Parameters**:
  - `RangeX`: Single or multi-column data range (e.g., A1:C100).
  - `standardized`: Boolean. `TRUE` to standardize (default), `FALSE` to duplicate.

#### 9.11. ra.Normalize
- **Function**: `ra.Normalize(RangeX, normalized)`
- **Description**: Returns normalized N(0,1) values from a distribution.
- **Parameters**:
  - `RangeX`: Single or multi-column data range (e.g., A1:C100).
  - `normalized`: Boolean. `TRUE` to normalize (default), `FALSE` to duplicate.

## Practical Example: Integrated Analysis
1. **Data**:
   - Time series in A1:A100 (prices).
   - Independent variables in B1:C100 (advertising, price).
   - Dependent variable in D1:D100 (sales).
2. **Stationarity**:
   - ADF Test: `=ra.DickeyFuller.ADF.Test(A1:A100, TRUE, 1, 1, 0.05)`.
   - KPSS Test: `=ra.KPSS.Test(A1:A100, TRUE, 1, 1, 0.05)`.
3. **Regression**:
   - Coefficients: `=ra.LinearReg.Coeff(B1:C100, D1:D100, TRUE, 6, 0.05)`.
   - Forecasts: `=ra.LinearReg.Forecast(B1:C100, D1:D100, E1:F10, TRUE, TRUE, 0.05)`.
   - Specification: `=ra.RamseyRESET.Test(B1:C100, D1:D100, TRUE, 2, 0.05)`.
   - Stability: `=ra.RecursiveCUSUM(B1:C100, D1:D100, TRUE, 0)`.
   - Autocorrelation: `=ra.Acorr.BreuschGodfrey.Test(B1:C100, D1:D100, TRUE, 2, 0.05, TRUE)`.
    - Multicollinearity: `=ra.MulticolLin.VIF(B1:C100, 0, 0)`.
   - Leverage: `=ra.Leverage(B1:C100, TRUE, 0)`.
   - Outliers: `=ra.Normalized.Distances(B1:C100, 0.01, 0)`.
4. **Diagnostics**:
   - Residuals: `=ra.LinearReg.Residuals(B1:C100, D1:D100, TRUE, 3, 0.1587, 0.04)`.
   - Heteroskedasticity: `=ra.Het.BreuschPagan.Test(B1:C100, D1:D100, TRUE, 0.05, TRUE)`.
5. **Portfolio Analysis**:
   - Risk: `=ra.Portfolio.Risk(H1:J3, K1:K3)` (where H1:J3 is covariance matrix, K1:K3 is weights).
   - VaR: `=ra.VaR.Historical(L1:L100, 0.05, TRUE)` (where L1:L100 is returns).
6. **Visualization**:
   - Histogram: `=ra.Histogram.Table(D1:D100, 2, FALSE)`.
   - Box Plot: `=ra.BoxPlot.Table(D1:D100, TRUE, 1.5)`.

## Troubleshooting
- **License Error**: Verify with `=ra.raXLStat.License()`. Contact support at [https://ruben-apaza.blogspot.com](https://ruben-apaza.blogspot.com).
- **Invalid Data**: Use `=ra.MissingData.Info(RangeX)` to detect empty or non-numeric cells.
- **Functions Not Working**: Ensure the add-in is enabled and data meets requirements.

## Additional Resources
- **Documentation**: [https://ruben-apaza.blogspot.com/p/raxl-stat.html](https://ruben-apaza.blogspot.com/p/raxl-stat.html).
- **Support**: Contact the developer via the blog or email provided with the license.
- **Video tutorial**: Our YouTube channel: [https://www.youtube.com/c/rubenapaza](https://www.youtube.com/watch?v=wYdGCkdN6cE&list=PLu4ltjreHhzO-cV1rHIis-K5_8numRqQV&pp=gAQBiAQB). 

[^1]: raXL Stat version v.0[Beta] is a test version to which new public functions will be added.
[^2]: Acknowledgment: raXL Stat uses Excel-DNA: Copyright (c) 2024 Govert van Drimmelen.
[^3]: The ARIMA and GARCH functions use the Maximum Likelihood Estimation (MLE) method together with the Newton-Raphson (NR) optimization algorithm, however, other optimization methods such as Levenberg-Marquardt, BHHH, BFGS and others will be added in development.
