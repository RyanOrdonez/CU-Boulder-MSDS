# Stock Portfolio Diversification Tool

**Personalized stock portfolio recommendation engine using unsupervised machine learning to cluster equities by risk-return profiles and generate diversified allocations.**

---

## Project Overview

This project builds an end-to-end portfolio recommendation system that takes a user's risk tolerance and growth strategy as inputs and outputs a diversified stock portfolio. Rather than relying on traditional financial heuristics or modern portfolio theory alone, the system uses K-Means clustering on engineered financial features to group stocks by their behavioral characteristics, then selects across clusters to ensure true diversification.

The pipeline covers the full data science workflow: sourcing historical price data for the 100 highest market-cap stocks across 11 sectors, engineering meaningful financial features (quarterly returns, volatility metrics), performing outlier removal using IQR filtering, and applying K-Means clustering to discover natural groupings in the risk-return landscape. A moderate positive correlation (r = 0.42) between quarterly returns and volatility confirms the classic risk-return tradeoff and validates the feature engineering choices.

Portfolio recommendations are backtested against historical data to evaluate real-world performance, providing users with both forward-looking recommendations and evidence-based confidence in the methodology.

## Dataset

| Property | Value |
|---|---|
| Source | Yahoo Finance (via yfinance API) |
| Universe | 100 highest market-cap stocks per sector |
| Sectors Covered | 11 (Technology, Healthcare, Financials, etc.) |
| Data Type | Historical daily price data |
| Engineered Features | Quarterly returns, rolling volatility, sector labels |

## Methodology

### Pipeline

```
Data Collection (yfinance)
    -> Exploratory Data Analysis
    -> Feature Engineering (quarterly returns, volatility)
    -> Outlier Removal (IQR method)
    -> K-Means Clustering
    -> Portfolio Generation (user preferences)
    -> Backtesting
```

### Feature Engineering

- **Quarterly Returns** -- Percentage price change over rolling 3-month windows
- **Volatility** -- Standard deviation of daily returns over rolling windows
- **Correlation Finding** -- Quarterly return vs. volatility shows a moderate positive correlation (r = 0.42), confirming that higher-return stocks tend to carry higher risk

### Clustering

- **Algorithm:** K-Means
- **Input Features:** Normalized quarterly returns and volatility
- **Cluster Selection:** Elbow method and silhouette analysis
- **Purpose:** Group stocks into behavioral archetypes (e.g., high-growth/high-volatility, stable/low-return)

### Portfolio Generation

1. User specifies risk tolerance (conservative, moderate, aggressive) and growth strategy
2. System selects target clusters matching the user's profile
3. Stocks are sampled across clusters and sectors to maximize diversification
4. Portfolio is backtested against historical performance

## Results

### Backtest Performance (2021-2024, $10,000 starting investment)

| Metric | Portfolio | S&P 500 |
|---|---|---|
| Final Value | **$22,741** | $13,226 |
| Annualized Return | **28.95%** | 10.92% |
| Sharpe Ratio | **1.73** | 0.62 |
| Volatility | 16.77% | 17.58% |

The cluster-based portfolio outperformed the S&P 500 benchmark by a wide margin while maintaining comparable volatility, demonstrating that unsupervised grouping of stocks by risk-return features can identify above-market combinations.

### Additional Findings

| Analysis | Finding |
|---|---|
| Return-Volatility Correlation | r = 0.42 (moderate positive) |
| Cluster Count | 9 behavioral groupings via K-Means |
| Sector Diversity | Portfolios span multiple sectors by design |
| Portfolio Size | 20 equal-weighted stocks |

## Key Visualizations

- **Scatter Plots** -- Quarterly returns vs. volatility colored by cluster assignment
- **Elbow and Silhouette Plots** -- Optimal cluster count determination
- **Sector Distribution Charts** -- Stock counts and cluster membership by sector
- **Backtesting Performance Curves** -- Portfolio returns over historical periods
- **Correlation Heatmaps** -- Feature relationships across the stock universe

## Tech Stack

| Category | Tools |
|---|---|
| Language | Python |
| Machine Learning | Scikit-learn (K-Means, preprocessing) |
| Data Sources | yfinance |
| Data Processing | Pandas, NumPy |
| Visualization | Matplotlib, Seaborn |

## How to Run

```bash
# Clone the repository
git clone https://github.com/RyanOrdonez/Final_Project_DTSA5510.git
cd Final_Project_DTSA5510

# Install dependencies
pip install scikit-learn pandas numpy yfinance matplotlib seaborn

# Run the notebook
jupyter notebook
```

> **Note:** The yfinance API pulls live data from Yahoo Finance. Historical data availability may vary; ensure an active internet connection when running the data collection cells.

## Author

Ryan Ordonez -- MS in Data Science, University of Colorado Boulder
