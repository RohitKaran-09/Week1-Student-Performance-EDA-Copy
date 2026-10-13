# Week 1 — Student Performance Data Acquisition, Cleaning & EDA

This repository contains the project report, Python script, demo dataset, and EDA visualizations for the Week 1 data-science assignment.

## Official dataset
UCI Machine Learning Repository — Student Performance: https://archive.ics.uci.edu/dataset/320/student%2Bperformance
DOI: https://doi.org/10.24432/C5TG7T

The Python script attempts to acquire the official UCI data. If the download is unavailable, it uses the explicitly labelled synthetic demo CSV in `data/` only to exercise the workflow. Do not report demo-derived statistics as official dataset findings.

## Run
```bash
pip install pandas numpy matplotlib seaborn ucimlrepo
python student_performance_eda.py
```

The script writes cleaned data, an audit log, summary statistics, and regenerated visualizations to `outputs/`.

## Cleaning steps
- Remove exact duplicate rows.
- Trim whitespace in text fields.
- Convert numeric variables safely.
- Handle missing numeric and categorical values when necessary.
- Check plausible ranges for age, studytime, failures, absences, and grades.
- Export a before/after cleaning audit.

## Visualizations
The workflow produces a final-grade histogram, average grade by study-time category, absences-versus-grade scatter plot, and numeric correlation heatmap.

## Academic reference
Cortez, P. (2008). Student Performance [Dataset]. UCI Machine Learning Repository. https://doi.org/10.24432/C5TG7T
