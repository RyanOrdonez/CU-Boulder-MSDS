# NYPD Shooting Incident Analysis

An R Markdown analysis of NYPD shooting incident data (2006 to present) investigating spatial, temporal, and demographic patterns. The goal of this project is to identify where community intervention programs -- such as after-school centers -- could be most effective in reducing youth-involved gun violence.

## Key Findings

- **Geographic concentration:** The Bronx and Brooklyn account for the highest number of shooting incidents across the study period.
- **Age demographics:** Perpetrators and victims are disproportionately concentrated in the 18-24 and 25-44 age groups.
- **Temporal patterns:** Incidents peak between 2:00 PM and 4:00 AM, with a sharp increase in evening hours.
- **Modeling:** A Generalized Linear Model (GLM) was fit to quantify relationships between demographic and geographic factors and incident frequency.

## Motivation

Understanding when and where shootings occur -- and who is most affected -- can inform targeted placement of community resources. The analysis specifically examines whether after-school community centers in high-incidence boroughs could help reduce youth crime during peak hours.

## Technologies Used

- **Language:** R
- **Libraries:** tidyverse, ggplot2, reshape2, pROC
- **Modeling:** Generalized Linear Model (GLM)
- **Format:** R Markdown (rendered to HTML/PDF)

## Data Source

[NYPD Shooting Incident Data (Historic)](https://data.cityofnewyork.us/Public-Safety/NYPD-Shooting-Incident-Data-Historic-/833y-fsy8) -- NYC Open Data, updated regularly.

## Author

Ryan
