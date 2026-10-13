# Week 4 — Machine Learning Model Development and Evaluation

This project trains and evaluates classification models using the official UCI Student Performance Mathematics dataset. The task is to predict whether a student reaches the descriptive threshold G3 >= 10 from non-grade features such as study habits, previous failure count, attendance, and demographic/family context.

## Important design choice
G1 and G2 are deliberately excluded from the predictors. They are highly correlated with G3 and can create target leakage when the goal is to assess prediction without earlier grades. The 10/20 threshold is an analysis convention, not a universal school policy.

## Dataset
UCI Machine Learning Repository, Student Performance: https://archive.ics.uci.edu/dataset/320/student%2Bperformance
DOI: https://doi.org/10.24432/C5TG7T

## Run locally
```bash
python -m pip install -r requirements.txt
python week4_model_development.py
```

The script first checks for the official CSV in the parent Week 1 project's data folder and downloads the UCI archive if needed. No synthetic fallback is used.

## Models and evaluation
- Logistic Regression
- Decision Tree Classifier
- Random Forest Classifier

Preprocessing is performed with Scikit-learn pipelines: median imputation and scaling for numeric variables; most-frequent imputation and one-hot encoding for categorical variables. Model selection uses stratified 5-fold cross-validation on the training split, with ROC-AUC as the selection metric. Final evaluation uses a stratified held-out test split and reports accuracy, precision, recall, F1, ROC-AUC, confusion matrix, and ROC curve.

## Deliverables
The script writes the DOCX report, a separate Word submission description, the cleaned modeling dataset, evaluation metrics, model-selection results, and at least four charts under `visualizations/`.
