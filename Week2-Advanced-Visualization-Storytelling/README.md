# Week 2 — Advanced Data Visualization and Storytelling with Python

This project extends Week 1 by building a narrative from the official UCI Student Performance Mathematics dataset. It creates multiple analytical visuals with captions, reasons for chart choices, computed findings, implications, and a formatted Word report.

## Dataset
UCI Machine Learning Repository — Student Performance: https://archive.ics.uci.edu/dataset/320/student%2Bperformance
DOI: https://doi.org/10.24432/C5TG7T

The analysis uses the official `student-mat.csv` (395 records). It first checks the existing Week 1 `../data/student-mat.csv`; for a stand-alone copy it downloads the official UCI ZIP archive. It does not replace unavailable source data with synthetic rows.

## Run
```bash
python -m pip install -r requirements.txt
python week2_storytelling.py
```

## Deliverables
- `Week2_Advanced_Data_Visualization_Report.docx` — full report with narrative, visualizations, captions, design rationale, conclusions, and limitations.
- `Submission_Description.docx` and `submission_description.txt` — 200+ word description for the internship submission form.
- `visualizations/` — seven figures generated from the official dataset.
- `outputs/` — descriptive statistics, chart-level metrics, and source/run metadata.

## Story arc
1. Understand the spread and progression of grades.
2. Explore how study-time categories differ descriptively.
3. Inspect the strong association between second-period and final grades.
4. Compare absences and prior failures with final outcomes.
5. Convert patterns into cautious, evidence-based implications.

All results are descriptive. Associations in observational data do not establish causation.
