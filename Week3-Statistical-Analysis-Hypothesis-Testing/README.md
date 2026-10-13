# Week 3 — Statistical Analysis and Hypothesis Testing in Python

This project applies statistical tests to the official UCI Student Performance Mathematics dataset. The primary hypothesis is that students with no previous class failures have a higher mean final grade than students with one or more previous failures.

## Dataset
UCI Machine Learning Repository — Student Performance: https://archive.ics.uci.edu/dataset/320/student%2Bperformance
DOI: https://doi.org/10.24432/C5TG7T

The script first uses the official source CSV included in the Week 1 project (`../data/student-mat.csv`); if it is missing, it attempts to download the UCI archive. It does not use synthetic data.

## Run
```bash
python -m pip install -r requirements.txt
python week3_hypothesis_testing.py
```

## Methods
- Primary: one-sided Welch independent-samples t-test, two-sided 95% confidence interval for mean difference, Hedges' g, plus a Mann–Whitney U sensitivity test.
- Secondary: chi-square test of independence between previous-failure group and final-grade threshold attainment, with Cramér's V.
- Exploratory: Kruskal–Wallis test for final grades across four study-time categories.
- Diagnostics: group distributions and Q–Q plots to contextualize the parametric test.

## Outputs
The script generates a polished Word report, a separate Word submission description, six visualizations, a data-quality audit, a JSON file containing all test statistics, and cleaned analysis data. All results are computed from the official UCI source records.
