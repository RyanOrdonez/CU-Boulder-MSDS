# Movie Ratings with Non-Negative Matrix Factorization

**Recommender System using NMF -- DTSA 5510: Unsupervised Algorithms in Machine Learning**

## Status

> This repository currently contains only the dataset files (train/test splits, movie metadata, user data). The analysis notebook needs to be added.

## Project Overview

This project applies Non-Negative Matrix Factorization (NMF) to a movie ratings dataset to build a collaborative filtering recommender system. NMF decomposes the user-item rating matrix into latent factor matrices, enabling prediction of missing ratings and generation of personalized movie recommendations.

## Dataset

| File | Description |
|------|-------------|
| `train.csv` | Training ratings (user-movie-rating triples) |
| `test.csv` | Held-out test ratings for evaluation |
| `movies.csv` | Movie metadata (titles, genres) |
| `users.csv` | User demographic information |

## TODO

- [ ] Add the analysis notebook with NMF implementation
- [ ] Include evaluation metrics (RMSE on test set)
- [ ] Add visualizations of latent topic distributions
- [ ] Compare NMF against baseline (mean rating, SVD)

## Technologies

Python, Scikit-learn, Pandas, NumPy, Matplotlib

## Author

Ryan Ordonez -- MS in Data Science, University of Colorado Boulder
