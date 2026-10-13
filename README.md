# Week 1 — Student Performance Data Acquisition, Cleaning & EDA

A reproducible Python project for acquiring the public UCI Student Performance dataset, cleaning it, exploring trends and relationships, and building a Word report with data-based findings.

## Official dataset
UCI Machine Learning Repository — Student Performance: https://archive.ics.uci.edu/dataset/320/student%2Bperformance
DOI: https://doi.org/10.24432/C5TG7T

This project uses the official Mathematics course file, `student-mat.csv`, downloaded from the UCI ZIP archive. The dataset covers student grades, study habits, absences, and demographic/family context from Portuguese secondary schools. The dataset is licensed CC BY 4.0; source attribution is included in the report.

## Run locally
```bash
python -m pip install -r requirements.txt
python student_performance_eda.py
```

The script does not silently use synthetic data. If the UCI archive is unavailable and no official source CSV is already present, it exits with an error. It creates the final `Week1_Student_Performance_EDA_Report.docx`, four PNG charts in `visualizations/`, and cleaned data, a cleaning audit, descriptive statistics, run metadata, and computed insights in `outputs/`.

## Cleaning steps
- Strip leading/trailing whitespace from text values and remove exact duplicate rows.
- Convert known numeric columns safely.
- Validate plausible ranges for grades, age, absences, study-time and survey scales.
- Treat invalid values as missing; impute numeric missing cells with medians and categorical values with the mode (or `Unknown`).
- Export audit counts before/after cleaning.

## Visualizations
1. Distribution of final Mathematics grades (G3).
2. Average final grade by weekly study-time category, including group sample sizes.
3. Absences versus final grade scatter plot.
4. Correlation heatmap for study time, failures, absences and period grades.

## GitHub Actions
The `Build official UCI EDA report` workflow runs when project code/configuration changes or when manually dispatched. It retrieves the official UCI data, builds the report and output files, and commits generated deliverables back to the `main` branch.

## Files
- `student_performance_eda.py` — acquisition, cleaning, analysis, visualizations and report generation.
- `requirements.txt` — Python dependencies.
- `submission_description.txt` — 200+ word task-submission description.
- `data/student-mat.csv` — official source CSV added by the automated run.
- `visualizations/` — figures generated from the official source data.
- `outputs/` — cleaned dataset, audit, summary statistics and metrics.
- `Week1_Student_Performance_EDA_Report.docx` — final Word report generated from official-data results.

## References
Cortez, P. (2008). Student Performance [Dataset]. UCI Machine Learning Repository. https://doi.org/10.24432/C5TG7T

Cortez, P., & Silva, A. M. G. (2008). Using data mining to predict secondary school student performance. Proceedings of the 5th Annual Future Business Technology Conference.
