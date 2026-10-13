# Week 5 — Comprehensive Data Science Project Reporting and Strategic Recommendations

This final capstone integrates Weeks 1–4 using the official UCI Student Performance Mathematics dataset.

## Project question
What can the data reliably tell us about student outcomes, which patterns are statistically supported, and how useful are non-grade features for identifying students who reach the descriptive final-grade threshold of G3 >= 10?

## Synthesis of the previous weeks
- Week 1: acquisition, data-quality checks, cleaning, descriptive statistics and EDA.
- Week 2: data storytelling, choice of visual form, chart annotation and interpretation.
- Week 3: a Welch two-sample hypothesis test, confidence interval, effect size, Mann–Whitney sensitivity test, chi-square and Kruskal–Wallis.
- Week 4: model pipeline, leakage-aware predictors, cross-validation, test evaluation, confusion matrix, ROC curve, and feature importance.

## Dataset
UCI Machine Learning Repository — Student Performance: https://archive.ics.uci.edu/dataset/320/student%2Bperformance
DOI: https://doi.org/10.24432/C5TG7T

The script uses the official Mathematics `student-mat.csv` already held in the Week 1 repository data folder, or downloads the official UCI archive if needed. It never substitutes synthetic data.

## Run
```bash
python -m pip install -r requirements.txt
python week5_capstone.py
```

## Output
- `Week5_Comprehensive_Data_Science_Capstone_Report.docx`
- `Submission_Description.docx` and `submission_description.txt` (200+ words)
- `visualizations/` — integrated charts for EDA, hypothesis testing and ML performance
- `outputs/` — results JSON, audit, CV comparison, test metrics and analysis data

## Main caution
The dataset is observational and comes from two Portuguese secondary schools. The models are educational prototypes, not tools for automated decisions about individual students. G1 and G2 are excluded from the ML predictors to prevent grade-history leakage when those grades are not available at prediction time.
