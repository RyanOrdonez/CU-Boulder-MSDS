# SpaceX Falcon 9 Landing Prediction

A complete end-to-end data science capstone project that predicts the success of SpaceX Falcon 9 first-stage landings. By analyzing historical launch data, this project builds and evaluates multiple classification models to determine whether a booster will successfully land -- a key factor in estimating launch cost.

## Project Overview

SpaceX advertises Falcon 9 rocket launches at a cost of $62 million, significantly less than other providers whose costs exceed $165 million. Much of this savings comes from the ability to reuse the first stage. Predicting whether the first stage will land successfully enables better cost estimation and competitive analysis.

## Pipeline

1. **Data Collection** -- REST API calls to the SpaceX API and web scraping of Wikipedia launch records
2. **Data Wrangling** -- Cleaning, feature engineering, and label encoding of landing outcomes
3. **Exploratory Data Analysis** -- SQL-based analysis and data visualization with Matplotlib and Seaborn
4. **Interactive Visualizations** -- Folium maps for launch site proximity analysis and a Plotly Dash dashboard for interactive exploration
5. **Predictive Modeling** -- Logistic Regression, SVM, Decision Tree, and KNN classifiers tuned with GridSearchCV

## Key Results

- **Best Accuracy: 83.3%** across all four tuned models on the test set
- Launch success rates have increased over time
- KSC LC-39A has the highest launch success rate among all sites
- Payload mass and orbit type are strong predictors of landing success

## Technologies Used

- **Languages:** Python, SQL
- **Libraries:** Pandas, NumPy, Scikit-learn, Matplotlib, Seaborn, Plotly Dash, Folium, BeautifulSoup
- **Methods:** Logistic Regression, SVM, Decision Tree, KNN, GridSearchCV, Feature Engineering

## Repository Structure

```
spacey-capstone/
├── data/                          # Collected and processed datasets
├── notebooks/                     # Jupyter notebooks for each pipeline stage
├── spacex_dash_app.py             # Plotly Dash interactive dashboard
└── README.md
```

## Getting Started

### Prerequisites

```
pip install pandas numpy scikit-learn matplotlib seaborn plotly dash folium beautifulsoup4
```

### Running the Dashboard

```bash
python spacex_dash_app.py
```

Navigate to `http://127.0.0.1:8050/` to explore launch data interactively.

## Author

Ryan
