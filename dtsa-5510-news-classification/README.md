# BBC News Topic Classification with NMF

**Multi-class news article classification using Non-Negative Matrix Factorization, demonstrating that unsupervised topic modeling can outperform supervised baselines.**

---

## Project Overview

This project explores an unconventional approach to text classification: using Non-Negative Matrix Factorization (NMF), an unsupervised topic modeling technique, to classify BBC news articles into five categories. The central question is whether latent topic structure discovered without labels can rival or exceed the performance of traditional supervised classifiers trained on labeled data.

The answer is a decisive yes. NMF with KL-divergence on TF-IDF features achieves 96.19% classification accuracy -- outperforming both Logistic Regression (94.91%) and Random Forest (91.42%) trained on 50% of the labeled data. This result highlights the strength of NMF for datasets where topic boundaries are well-defined and the vocabulary is domain-specific, as is the case with news articles spanning distinct categories like sport, business, and technology.

The pipeline includes thorough text preprocessing with NLTK, TF-IDF and CountVectorizer feature extraction, and a systematic comparison across modeling approaches.

## Dataset

| Property | Value |
|---|---|
| Source | BBC News Dataset |
| Total Articles (raw) | 1,490 |
| Total Articles (after dedup) | 1,376 |
| Categories | 5 (Business, Entertainment, Politics, Sport, Tech) |
| Preprocessing | Duplicate removal, stopword filtering, lemmatization |

## Methodology

### Text Preprocessing

1. **Deduplication** -- Removed 114 duplicate articles
2. **Tokenization** -- NLTK word tokenizer
3. **Stopword Removal** -- Standard English stopwords via NLTK
4. **Feature Extraction** -- TF-IDF and CountVectorizer representations

### Models Compared

| Model | Type | Training Data | Approach |
|---|---|---|---|
| **NMF** | **Unsupervised** | **100% (no labels needed)** | **Topic modeling + cluster-to-label mapping** |
| Logistic Regression | Supervised | 50% labeled | Standard classification |
| Random Forest | Supervised | 50% labeled | Ensemble classification |

### NMF Configuration

- **n_components:** 5 (matching the known number of categories)
- **Divergence metric:** Kullback-Leibler
- **Input features:** TF-IDF matrix
- **Label assignment:** Topics mapped to categories via dominant topic per document

## Results

| Model | Accuracy | Supervision Required |
|---|---|---|
| **NMF (KL-divergence)** | **96.19%** | **None** |
| Logistic Regression | 94.91% | 50% labeled data |
| Random Forest | 91.42% | 50% labeled data |

### Key Insight

The unsupervised NMF model outperformed both supervised baselines by a meaningful margin. This is particularly notable because NMF requires zero labeled training examples -- it discovers topic structure purely from word co-occurrence patterns in the TF-IDF matrix. For well-separated topic domains like news categories, NMF provides an effective and label-efficient alternative to supervised classification.

## Key Visualizations

- **Topic-Term Heatmaps** -- Top terms per NMF component showing clean topic separation
- **Category Distribution** -- Article counts across the five BBC categories
- **Model Comparison Charts** -- Accuracy comparison across all three approaches
- **Confusion Matrices** -- Per-model error patterns and misclassification analysis

## Tech Stack

| Category | Tools |
|---|---|
| Language | Python |
| Machine Learning | Scikit-learn (NMF, LogisticRegression, RandomForest) |
| NLP | NLTK |
| Feature Extraction | TF-IDF, CountVectorizer |
| Data Processing | Pandas, NumPy |
| Visualization | Matplotlib |

## How to Run

```bash
# Clone the repository
git clone https://github.com/RyanOrdonez/DTSA5510_Week4_NewsClassification.git
cd DTSA5510_Week4_NewsClassification

# Install dependencies
pip install scikit-learn nltk pandas numpy matplotlib

# Download NLTK data
python -c "import nltk; nltk.download('stopwords'); nltk.download('punkt'); nltk.download('wordnet')"

# Run the notebook
jupyter notebook
```

## Author

Ryan Ordonez -- MS in Data Science, University of Colorado Boulder
