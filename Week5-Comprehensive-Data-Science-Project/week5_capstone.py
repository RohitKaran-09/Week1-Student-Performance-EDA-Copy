"""Week 5 capstone: Comprehensive Data Science Project Reporting.

Integrates the official UCI Student Performance Mathematics dataset with
cleaning/EDA, visual storytelling, hypothesis testing, and ML model evaluation.
No synthetic data fallback is used. ML predictors intentionally exclude G1/G2/G3
to reduce target leakage for a prediction setting where earlier grades may not
be available.
"""
from io import BytesIO
from pathlib import Path
import json, math, urllib.request, zipfile

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats

from sklearn.compose import ColumnTransformer
from sklearn.dummy import DummyClassifier
from sklearn.ensemble import RandomForestClassifier
from sklearn.impute import SimpleImputer
from sklearn.inspection import permutation_importance
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import (
    accuracy_score, balanced_accuracy_score, precision_score, recall_score,
    f1_score, roc_auc_score, confusion_matrix, classification_report,
    RocCurveDisplay
)
from sklearn.model_selection import train_test_split, StratifiedKFold, cross_validate
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder, StandardScaler
from sklearn.tree import DecisionTreeClassifier

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "outputs"
PLOTS = ROOT / "visualizations"
REPORT = ROOT / "Week5_Comprehensive_Data_Science_Capstone_Report.docx"
SUBMISSION_DOCX = ROOT / "Submission_Description.docx"
SUBMISSION_TXT = ROOT / "submission_description.txt"
UCI_ZIP = "https://archive.ics.uci.edu/static/public/320/student+performance.zip"
UCI_PAGE = "https://archive.ics.uci.edu/dataset/320/student%2Bperformance"
DOI = "https://doi.org/10.24432/C5TG7T"
ALPHA = 0.05
PASS_MARK = 10
RANDOM_STATE = 42
TEST_SIZE = 0.20

def find_math_csv(zf):
    for name in zf.namelist():
        if name.lower().endswith("student-mat.csv"):
            return zf.read(name)
    for name in zf.namelist():
        if name.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(BytesIO(zf.read(name))) as nested:
                    data = find_math_csv(nested)
                    if data is not None:
                        return data
            except zipfile.BadZipFile:
                continue
    return None

def acquire():
    candidates = [
        ROOT.parent.parent / "data" / "student-mat.csv",
        ROOT.parent / "data" / "student-mat.csv",
        ROOT / "data" / "student-mat.csv",
    ]
    for path in candidates:
        if path.exists():
            df = pd.read_csv(path, sep=";")
            if len(df) and {"G1", "G2", "G3", "studytime", "failures"}.issubset(df.columns):
                return df, "Official UCI Mathematics CSV already present in the Week 1 repository"
    request = urllib.request.Request(UCI_ZIP, headers={"User-Agent": "Week5-Student-Data-Science/1.0"})
    try:
        with urllib.request.urlopen(request, timeout=60) as resp:
            zipped = resp.read()
        with zipfile.ZipFile(BytesIO(zipped)) as zf:
            payload = find_math_csv(zf)
        if payload is None:
            raise RuntimeError("student-mat.csv was not found in official archive")
        df = pd.read_csv(BytesIO(payload), sep=";")
        if df.empty or not {"G1", "G2", "G3", "studytime", "failures"}.issubset(df.columns):
            raise RuntimeError("Official CSV failed schema checks")
        return df, "Downloaded official UCI archive: " + UCI_ZIP
    except Exception as exc:
        raise RuntimeError("Official UCI acquisition failed; no synthetic data was substituted.") from exc

def prepare(raw):
    df = raw.copy()
    audit = {"source_rows": int(len(df)), "source_columns": int(len(df.columns)), "synthetic_fallback_used": False}
    audit["duplicate_rows_found"] = int(df.duplicated().sum())
    df = df.drop_duplicates().copy()
    for col in df.select_dtypes(include=["object", "string"]).columns:
        df[col] = df[col].astype("string").str.strip()
    numeric = ["age","Medu","Fedu","traveltime","studytime","failures","famrel","freetime","goout","Dalc","Walc","health","absences","G1","G2","G3"]
    for col in numeric:
        if col in df:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    audit["missing_cells_after_type_conversion"] = int(df.isna().sum().sum())
    required = ["G1","G2","G3","studytime","failures","absences"]
    audit["missing_cells_in_required_fields"] = int(df[required].isna().sum().sum())
    n_before_drop = len(df)
    df = df.dropna(subset=required).copy()
    audit["rows_removed_for_missing_required_fields"] = int(n_before_drop - len(df))
    ranges = {"G1":(0,20),"G2":(0,20),"G3":(0,20),"studytime":(1,4),"failures":(0,4),"absences":(0,200)}
    invalid_by_col = {}
    invalid = pd.Series(False, index=df.index)
    for col,(lo,hi) in ranges.items():
        mask = ~df[col].between(lo,hi)
        invalid_by_col[col] = int(mask.sum())
        invalid |= mask
    audit["invalid_values_by_field"] = invalid_by_col
    audit["rows_removed_for_invalid_ranges"] = int(invalid.sum())
    df = df.loc[~invalid].copy()
    if df.empty:
        raise RuntimeError("No records remain after source-data validation.")
    df["target_pass"] = (df["G3"] >= PASS_MARK).astype(int)
    audit["rows_after_cleaning"] = int(len(df))
    audit["columns_after_cleaning"] = int(len(df.columns))
    audit["cleaning_summary"] = "Exact duplicates removed; text fields trimmed; numeric fields converted; required analysis fields checked for missing values; grades and key scales range-checked. Target values were not imputed."
    return df, audit

def welch_df(a,b):
    a=np.asarray(a,dtype=float); b=np.asarray(b,dtype=float)
    t1=np.var(a,ddof=1)/len(a); t2=np.var(b,ddof=1)/len(b)
    return float((t1+t2)**2/(t1**2/(len(a)-1)+t2**2/(len(b)-1)))

def hedges_g(a,b):
    a=np.asarray(a,dtype=float); b=np.asarray(b,dtype=float)
    n1,n2=len(a),len(b)
    pooled=math.sqrt(((n1-1)*np.var(a,ddof=1)+(n2-1)*np.var(b,ddof=1))/(n1+n2-2))
    d=(np.mean(a)-np.mean(b))/pooled if pooled else 0.0
    return float((1-3/(4*(n1+n2)-9))*d)

def stats_analysis(df):
    a=df.loc[df["failures"].eq(0),"G3"].to_numpy(dtype=float)
    b=df.loc[df["failures"].ge(1),"G3"].to_numpy(dtype=float)
    n1,n2=len(a),len(b); m1,m2=float(a.mean()),float(b.mean())
    sd1,sd2=float(a.std(ddof=1)),float(b.std(ddof=1))
    diff=m1-m2; se=math.sqrt(sd1**2/n1+sd2**2/n2); dof=welch_df(a,b); t=diff/se
    p=float(stats.ttest_ind(a,b,equal_var=False,alternative="greater").pvalue)
    crit=float(stats.t.ppf(0.975,dof)); ci=[float(diff-crit*se),float(diff+crit*se)]
    mw=stats.mannwhitneyu(a,b,alternative="greater",method="auto")
    ct=pd.crosstab(df["failures"].eq(0).map({True:"No prior failures",False:"One or more prior failures"}),df["target_pass"]).reindex(
       index=["No prior failures","One or more prior failures"],columns=[0,1],fill_value=0)
    chi,pchi,dofchi,expected=stats.chi2_contingency(ct.to_numpy(),correction=False)
    v=math.sqrt(chi/(len(df)*min(ct.shape[0]-1,ct.shape[1]-1)))
    groups=[g["G3"].to_numpy(dtype=float) for _,g in df.groupby("studytime",observed=True)]
    kw=stats.kruskal(*groups)
    study={}
    for k,g in df.groupby("studytime",observed=True):
        study[str(int(k))]={"n":int(len(g)),"mean":float(g["G3"].mean()),"median":float(g["G3"].median()),"sd":float(g["G3"].std(ddof=1))}
    failrate={}
    for key,mask in [("No prior failures",df["failures"].eq(0)),("One or more prior failures",df["failures"].ge(1))]:
        sub=df.loc[mask]
        failrate[key]={"n":int(len(sub)),"pass_n":int(sub["target_pass"].sum()),"pass_rate_pct":float(100*sub["target_pass"].mean()),"mean_g3":float(sub["G3"].mean())}
    return {
        "primary_test":{"question":"Do students with no prior class failures have higher mean G3 than students with one or more prior failures?","h0":"mu_no_failure <= mu_prior_failure","h1":"mu_no_failure > mu_prior_failure","alpha":ALPHA,"method":"One-sided Welch independent-samples t-test","n_no_prior_failures":n1,"mean_no_prior_failures":m1,"sd_no_prior_failures":sd1,"n_one_or_more_failures":n2,"mean_one_or_more_failures":m2,"sd_one_or_more_failures":sd2,"mean_difference":float(diff),"standard_error":float(se),"t_statistic":float(t),"degrees_freedom":float(dof),"p_value_one_sided":p,"ci95_two_sided":[ci[0],ci[1]],"hedges_g":hedges_g(a,b),"decision":"Reject H0 at alpha=0.05" if p<ALPHA else "Fail to reject H0 at alpha=0.05"},
        "sensitivity_mann_whitney":{"u_statistic":float(mw.statistic),"p_value_one_sided":float(mw.pvalue)},
        "secondary_chi_square":{"observed":ct.to_numpy().astype(int).tolist(),"expected":expected.tolist(),"chi_square":float(chi),"df":int(dofchi),"p_value":float(pchi),"cramers_v":float(v),"interpretation":"Evidence of association at alpha=0.05" if pchi<ALPHA else "No statistically significant evidence at alpha=0.05"},
        "exploratory_kruskal_wallis":{"h_statistic":float(kw.statistic),"df":int(len(groups)-1),"p_value":float(kw.pvalue),"interpretation":"Evidence at least one study-time distribution differs" if kw.pvalue<ALPHA else "Insufficient evidence of different distributions at alpha=0.05"},
        "group_rates":failrate,"studytime_groups":study,
        "descriptive":{"n":int(len(df)),"mean_g1":float(df.G1.mean()),"mean_g2":float(df.G2.mean()),"mean_g3":float(df.G3.mean()),"median_g3":float(df.G3.median()),"std_g3":float(df.G3.std(ddof=1)),"pass_count":int(df.target_pass.sum()),"below_count":int((df.target_pass==0).sum()),"pass_rate_pct":float(df.target_pass.mean()*100),"g2_g3_corr":float(df.G2.corr(df.G3)),"absences_g3_corr":float(df.absences.corr(df.G3)),"failures_g3_corr":float(df.failures.corr(df.G3)),"min_g3":float(df.G3.min()),"max_g3":float(df.G3.max())},
        "limitations":["Observational associations do not prove causation.","The dataset represents two Portuguese secondary schools.","The G3 >= 10 threshold is a descriptive analytical convention, not a universal school policy."]
    }

def ml_analysis(df):
    excluded={"G1","G2","G3","target_pass"}
    X=df.drop(columns=[c for c in excluded if c in df]).copy()
    y=df["target_pass"].astype(int)
    for c in X.select_dtypes(include=["string"]).columns:
        X[c]=X[c].astype(object)
    X_train,X_test,y_train,y_test=train_test_split(X,y,test_size=TEST_SIZE,random_state=RANDOM_STATE,stratify=y)
    numeric=X_train.select_dtypes(include=["number","bool"]).columns.tolist()
    categorical=[c for c in X_train.columns if c not in numeric]
    transformers=[]
    if numeric:
        transformers.append(("numeric",Pipeline([("imputer",SimpleImputer(strategy="median")),("scale",StandardScaler())]),numeric))
    if categorical:
        transformers.append(("categorical",Pipeline([("imputer",SimpleImputer(strategy="most_frequent")),("onehot",OneHotEncoder(handle_unknown="ignore"))]),categorical))
    pre=ColumnTransformer(transformers=transformers,remainder="drop")
    models={
       "Dummy baseline":Pipeline([("preprocess",pre),("model",DummyClassifier(strategy="prior"))]),
       "Logistic Regression":Pipeline([("preprocess",pre),("model",LogisticRegression(max_iter=3000,class_weight="balanced",random_state=RANDOM_STATE))]),
       "Decision Tree":Pipeline([("preprocess",pre),("model",DecisionTreeClassifier(max_depth=4,min_samples_leaf=8,class_weight="balanced",random_state=RANDOM_STATE))]),
       "Random Forest":Pipeline([("preprocess",pre),("model",RandomForestClassifier(n_estimators=250,min_samples_leaf=4,class_weight="balanced",random_state=RANDOM_STATE,n_jobs=1))])
    }
    cv=StratifiedKFold(n_splits=5,shuffle=True,random_state=RANDOM_STATE)
    rows=[]
    for name,pipeline in models.items():
        s=cross_validate(pipeline,X_train,y_train,cv=cv,scoring={"auc":"roc_auc","f1":"f1","balanced_accuracy":"balanced_accuracy"},n_jobs=1)
        rows.append({"model":name,"cv_auc_mean":float(np.mean(s["test_auc"])),"cv_auc_sd":float(np.std(s["test_auc"],ddof=1)),"cv_f1_mean":float(np.mean(s["test_f1"])),"cv_balanced_accuracy_mean":float(np.mean(s["test_balanced_accuracy"]))})
    cv_table=pd.DataFrame(rows).sort_values("cv_auc_mean",ascending=False).reset_index(drop=True)
    selected_name=str(cv_table.iloc[0]["model"]); model=models[selected_name]
    model.fit(X_train,y_train); pred=model.predict(X_test); prob=model.predict_proba(X_test)[:,1]
    cm=confusion_matrix(y_test,pred,labels=[0,1])
    metrics={"selected_model":selected_name,"selection_method":"Highest mean stratified 5-fold CV ROC-AUC on training data","random_state":RANDOM_STATE,"test_size_fraction":TEST_SIZE,"n_total":int(len(df)),"n_train":int(len(y_train)),"n_test":int(len(y_test)),"positive_rate_pct":float(100*y.mean()),"test_accuracy":float(accuracy_score(y_test,pred)),"test_balanced_accuracy":float(balanced_accuracy_score(y_test,pred)),"test_precision":float(precision_score(y_test,pred,zero_division=0)),"test_recall":float(recall_score(y_test,pred,zero_division=0)),"test_f1":float(f1_score(y_test,pred,zero_division=0)),"test_roc_auc":float(roc_auc_score(y_test,prob)),"confusion_matrix":cm.tolist(),"confusion_matrix_order":["G3 < 10","G3 >= 10"],"class_report":classification_report(y_test,pred,labels=[0,1],target_names=["Below threshold","At/above threshold"],output_dict=True,zero_division=0),"cv_results":cv_table.to_dict(orient="records"),"predictor_columns":list(X.columns),"excluded_grade_columns":["G1","G2","G3"],"feature_importance_note":"Permutation importance is predictive association, not causation."}
    try:
        perm=permutation_importance(model,X_test,y_test,scoring="roc_auc",n_repeats=25,random_state=RANDOM_STATE,n_jobs=1)
        importance=pd.DataFrame({"feature":X_test.columns,"importance_mean":perm.importances_mean,"importance_sd":perm.importances_std}).sort_values("importance_mean",ascending=False)
        importance.to_csv(OUT/"permutation_feature_importance.csv",index=False)
    except Exception as exc:
        importance=pd.DataFrame(columns=["feature","importance_mean","importance_sd"])
        (OUT/"feature_importance_warning.txt").write_text(str(exc),encoding="utf-8")
    return model, X, y, X_test, y_test, pred, prob, cv_table, metrics, importance

def save_chart(fig,name,caption=None):
    PLOTS.mkdir(parents=True,exist_ok=True)
    if caption:
        fig.text(.01,.004,caption,ha="left",va="bottom",fontsize=8,wrap=True)
        fig.tight_layout(rect=[0,.075,1,1])
    else: fig.tight_layout()
    p=PLOTS/name
    fig.savefig(p,dpi=180,bbox_inches="tight")
    plt.close(fig)
    return p

def charts(df,st,ml,model,X_test,y_test,pred,prob,cv_table,importance):
    sns.set_theme(style="whitegrid",context="notebook")
    paths=[]
    # 1 final grade distribution
    fig,ax=plt.subplots(figsize=(9,5.2))
    ax.hist(df.G3,bins=np.arange(-.5,21.5,1),color="#4C78A8",edgecolor="white")
    ax.axvline(df.G3.mean(),color="#E45756",lw=2,label=f"Mean {df.G3.mean():.2f}")
    ax.axvline(PASS_MARK,color="#54A24B",ls="--",lw=2,label=f"Reference threshold {PASS_MARK}/20")
    ax.set(title="1. Final Mathematics Grade Distribution",xlabel="Final grade G3 (0–20)",ylabel="Number of students"); ax.legend()
    paths.append(save_chart(fig,"01_final_grade_distribution.png","Caption: The full distribution provides baseline context. The 10/20 line is a reporting convention for this project, not a universal policy."))
    # 2 grades across periods
    fig,ax=plt.subplots(figsize=(8.5,5.2))
    long=df[["G1","G2","G3"]].melt(var_name="Period",value_name="Grade")
    sns.boxplot(data=long,x="Period",y="Grade",hue="Period",legend=False,palette=["#72B7B2","#F2CF5B","#4C78A8"],ax=ax)
    means=df[["G1","G2","G3"]].mean()
    for i,col in enumerate(["G1","G2","G3"]):
        ax.scatter(i,means[col],marker="D",color="#E45756",s=40,zorder=3)
        ax.text(i,means[col]+.5,f"{means[col]:.1f}",ha="center",fontsize=8)
    ax.set(title="2. Grades Across Assessment Periods",xlabel="Assessment period",ylabel="Grade (0–20)")
    paths.append(save_chart(fig,"02_grade_progression.png","Caption: Boxplots show medians, quartiles and spread; red diamonds mark means. Cohort-level differences do not imply that each student changed in the same way."))
    # 3 failure groups and CI
    t=st["primary_test"]
    labels=["No prior failures","One or more failures"]
    means=[t["mean_no_prior_failures"],t["mean_one_or_more_failures"]]
    sds=[t["sd_no_prior_failures"],t["sd_one_or_more_failures"]]
    ns=[t["n_no_prior_failures"],t["n_one_or_more_failures"]]
    errs=[stats.t.ppf(.975,n-1)*sd/np.sqrt(n) for sd,n in zip(sds,ns)]
    fig,ax=plt.subplots(figsize=(8.5,5.2))
    bars=ax.bar(range(2),means,yerr=errs,capsize=5,color=["#4C78A8","#F58518"],width=.6)
    ax.set_xticks(range(2),[f"{labels[i]}\n(n={ns[i]})" for i in range(2)])
    ax.set(ylim=(0,20),title="3. Prior-Failure Groups Have Different Mean Final Grades",ylabel="Mean G3 with 95% group intervals")
    for b,v in zip(bars,means): ax.text(b.get_x()+b.get_width()/2,v+.5,f"{v:.2f}",ha="center",fontweight="bold")
    paths.append(save_chart(fig,"03_mean_grade_by_failure_group.png",f"Caption: Mean difference is {t['mean_difference']:.2f} points; Welch-test p={t['p_value_one_sided']:.2g}. Confidence intervals are group-specific; the primary CI for the difference appears in the report."))
    # 4 correlation heatmap
    cols=["studytime","failures","absences","G1","G2","G3"]
    fig,ax=plt.subplots(figsize=(8,6))
    sns.heatmap(df[cols].corr(),annot=True,fmt=".2f",cmap="vlag",vmin=-1,vmax=1,center=0,square=True,ax=ax)
    ax.set_title("4. Correlation Map of Grades and Context Variables")
    paths.append(save_chart(fig,"04_correlation_heatmap.png","Caption: Pairwise Pearson correlations describe linear relationships only. Earlier grades are shown for EDA but excluded from machine-learning predictors to avoid leakage."))
    # 5 failure pass rates
    groups=st["group_rates"]
    order=["No prior failures","One or more prior failures"]
    vals=[groups[k]["pass_rate_pct"] for k in order]; ns2=[groups[k]["n"] for k in order]
    fig,ax=plt.subplots(figsize=(8.5,5.2))
    bars=ax.bar(range(2),vals,color=["#4C78A8","#F58518"],width=.6)
    ax.set_xticks(range(2),[f"{order[i]}\n(n={ns2[i]})" for i in range(2)])
    ax.set(ylim=(0,100),title="5. Threshold Attainment by Prior-Failure Group",ylabel=f"Share with G3 ≥ {PASS_MARK} (%)")
    for b,v in zip(bars,vals): ax.text(b.get_x()+b.get_width()/2,v+2,f"{v:.1f}%",ha="center",fontweight="bold")
    paths.append(save_chart(fig,"05_threshold_rate_by_failure_group.png","Caption: Percent reaching 10/20 differs across groups. Threshold status is a descriptive indicator; prior failure is not a deterministic label for an individual."))
    # 6 CV model comparison
    fig,ax=plt.subplots(figsize=(8.8,5.2))
    tab=cv_table.sort_values("cv_auc_mean",ascending=True)
    ax.barh(tab.model,tab.cv_auc_mean,xerr=tab.cv_auc_sd,capsize=4,color=["#9ECAE9","#6BAED6","#3182BD","#08519C"][-len(tab):])
    ax.axvline(.5,color="#E45756",ls="--",label="Chance AUC = 0.50")
    ax.set(xlim=(0,1),title="6. Model Selection Using Training-Only Cross-Validation",xlabel="Mean 5-fold ROC-AUC"); ax.legend()
    paths.append(save_chart(fig,"06_cross_validation_model_comparison.png","Caption: Error bars show fold-to-fold standard deviation. The chosen model is selected using training data only; held-out test results are reported separately."))
    # 7 confusion matrix
    fig,ax=plt.subplots(figsize=(6.5,5.2))
    sns.heatmap(confusion_matrix(y_test,pred,labels=[0,1]),annot=True,fmt="d",cmap="Blues",cbar=False,xticklabels=["Predicted below","Predicted at/above"],yticklabels=["Actual below","Actual at/above"],ax=ax)
    ax.set(title="7. Held-Out Test Confusion Matrix",xlabel="Predicted label",ylabel="Actual label")
    paths.append(save_chart(fig,"07_test_confusion_matrix.png","Caption: False positives and false negatives represent different kinds of error. A model output should prompt review, not automatic individual decisions."))
    # 8 ROC
    fig,ax=plt.subplots(figsize=(7.2,5.2))
    RocCurveDisplay.from_predictions(y_test,prob,name=ml["selected_model"],ax=ax)
    ax.plot([0,1],[0,1],ls="--",color="gray",label="Chance")
    ax.set(title="8. ROC Curve on the Held-Out Test Set",xlabel="False-positive rate",ylabel="True-positive rate"); ax.legend(loc="lower right")
    paths.append(save_chart(fig,"08_test_roc_curve.png","Caption: ROC-AUC of about 0.60 is only modestly above chance, indicating limited discrimination from non-grade predictors in this test split."))
    # 9 feature importance
    if not importance.empty:
        top=importance.head(10).sort_values("importance_mean",ascending=True)
        fig,ax=plt.subplots(figsize=(9,5.8))
        ax.barh(top.feature,top.importance_mean,xerr=top.importance_sd,capsize=3,color="#4C78A8")
        ax.axvline(0,color="gray",lw=1)
        ax.set(title="9. Permutation Importance on Held-Out Data",xlabel="Mean decrease in ROC-AUC after shuffling a feature")
        paths.append(save_chart(fig,"09_permutation_feature_importance.png","Caption: This chart shows model reliance, not causation. Permutation importance can be unstable with small data and correlated predictors."))
    return paths

def cell_text(cell,text,bold=False):
    cell.text=str(text); cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for para in cell.paragraphs:
        for run in para.runs:
            run.font.name="Aptos";run.font.size=Pt(8.5);run.bold=bold

def build_submission_doc():
    content=SUBMISSION_TXT.read_text(encoding="utf-8").strip()
    doc=Document()
    sec=doc.sections[0]; sec.top_margin=Inches(.75); sec.bottom_margin=Inches(.75);sec.left_margin=Inches(.85);sec.right_margin=Inches(.85)
    doc.styles["Normal"].font.name="Aptos"; doc.styles["Normal"].font.size=Pt(11)
    p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run("WEEK 5 TASK");r.bold=True;r.font.size=Pt(11);r.font.color.rgb=RGBColor(232,106,0)
    p=doc.add_heading("Submission Description",0);p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p=doc.add_paragraph("Comprehensive Data Science Project Reporting and Strategic Recommendations");p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    for paragraph in content.split("\n\n"):
        p=doc.add_paragraph(paragraph);p.paragraph_format.space_after=Pt(10);p.paragraph_format.line_spacing=1.12
    p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    r=p.add_run(f"Word count: {len(content.split())}");r.italic=True;r.font.size=Pt(9)
    doc.save(SUBMISSION_DOCX)

def build_report(df,audit,st,ml,cv_table,importance,images,source):
    doc=Document()
    sec=doc.sections[0];sec.top_margin=Inches(.58);sec.bottom_margin=Inches(.58);sec.left_margin=Inches(.72);sec.right_margin=Inches(.72)
    doc.styles["Normal"].font.name="Aptos";doc.styles["Normal"].font.size=Pt(9.5)
    doc.styles["Title"].font.name="Aptos Display";doc.styles["Title"].font.size=Pt(23)
    doc.styles["Heading 1"].font.name="Aptos Display";doc.styles["Heading 1"].font.size=Pt(15)
    doc.styles["Heading 2"].font.name="Aptos Display";doc.styles["Heading 2"].font.size=Pt(11.5)
    p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run("WEEK 5 CAPSTONE REPORT");r.bold=True;r.font.size=Pt(12);r.font.color.rgb=RGBColor(232,106,0)
    p=doc.add_heading("Comprehensive Data Science Project Reporting and Strategic Recommendations",0);p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p=doc.add_paragraph("Integrated analysis of UCI Student Performance — Mathematics | Weeks 1–4 synthesis");p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    d=st["descriptive"]; pt=st["primary_test"]; mw=st["sensitivity_mann_whitney"]; chi=st["secondary_chi_square"];kw=st["exploratory_kruskal_wallis"]

    doc.add_heading("1. Executive Summary",1)
    doc.add_paragraph(
      f"This capstone integrates data preparation, exploratory analysis, visual storytelling, hypothesis testing and machine learning for {d['n']} student records from the official UCI Mathematics dataset. "
      f"Average final grade is {d['mean_g3']:.2f}/20 (median {d['median_g3']:.1f}); {d['pass_count']} students ({d['pass_rate_pct']:.1f}%) meet the report's descriptive threshold of 10/20. "
      f"The pre-specified Welch test found that students with no prior failures had a mean G3 of {pt['mean_no_prior_failures']:.2f}, compared with {pt['mean_one_or_more_failures']:.2f} for students with one or more failures; difference {pt['mean_difference']:.2f} points, 95% CI [{pt['ci95_two_sided'][0]:.2f}, {pt['ci95_two_sided'][1]:.2f}], p={pt['p_value_one_sided']:.2g}. "
      f"A Logistic Regression model trained without G1/G2 achieved held-out ROC-AUC {ml['test_roc_auc']:.3f}, only modestly above chance. "
      "Consequently, the most defensible recommendation is to use transparent descriptive indicators to guide supportive review, not to automate decisions about students."
    )
    doc.add_heading("Executive Recommendations",2)
    for item in [
      "Offer timely, supportive academic review to students with recorded prior failures, while considering individual context rather than using a single variable as a label.",
      "Treat the current classifier as a learning prototype only; its test ROC-AUC indicates limited discrimination and does not justify operational deployment.",
      "Validate any future model on larger, newer and more diverse cohorts; assess calibration, subgroup error rates, privacy and fairness before considering use.",
      "Keep analysis reproducible with documented cleaning, held-out evaluation, effect sizes, confidence intervals and clear limitations."
    ]: doc.add_paragraph(item,style="List Bullet")

    doc.add_heading("2. Project Scope and Research Questions",1)
    doc.add_paragraph("Business/scientific objective: summarize student achievement patterns, test one clear hypothesis, evaluate whether non-grade context variables can classify a descriptive grade threshold, and translate the evidence into careful recommendations for educators.")
    for item in [
      "EDA question: How are final grades distributed, and how do they relate to prior grade history, absences, study-time categories and previous failures?",
      "Statistical question: Do students with no recorded previous class failures have a higher mean final grade than students with one or more prior failures?",
      "Modeling question: How well can non-grade features classify whether G3 reaches 10/20, when G1 and G2 are excluded to avoid target leakage?",
      "Decision question: Which actions are supported by the evidence, and what additional validation is needed?"
    ]: doc.add_paragraph(item,style="List Number")

    doc.add_heading("3. Dataset, Acquisition and Quality",1)
    doc.add_paragraph(f"Source: UCI Machine Learning Repository, Student Performance Mathematics file. Dataset page: {UCI_PAGE}. DOI: {DOI}. Acquisition method: {source}. The analysis uses the official data, not a synthetic substitute.")
    doc.add_paragraph(f"The source contained {audit['source_rows']} rows and {audit['source_columns']} columns. The quality audit recorded {audit['duplicate_rows_found']} exact duplicates, {audit['missing_cells_in_required_fields']} missing cells in required analysis fields and {audit['rows_removed_for_invalid_ranges']} rows removed for invalid ranges. After preparation, {audit['rows_after_cleaning']} records remain.")
    doc.add_paragraph("Cleaning process: trim text fields; safely convert known numeric columns; check required fields; remove exact duplicates; validate grade, study-time, previous-failure and absence ranges; exclude records with missing target/required values rather than fabricate grade outcomes. The original official dataset had no missing values or exact duplicates under these checks, so no imputation or row removal was needed in this run.")
    doc.add_heading("4. Integrated Methodology",1)
    doc.add_heading("4.1 Week 1 — Cleaning and EDA",2)
    doc.add_paragraph(f"Final grade G3 has mean {d['mean_g3']:.2f}, median {d['median_g3']:.1f}, standard deviation {d['std_g3']:.2f}, and range {d['min_g3']:.0f}–{d['max_g3']:.0f}. The dataset has {d['pass_count']} cases at or above 10/20 and {d['below_count']} below. The 10/20 cut-off is used only as an analysis convention, not a universal school policy.")
    doc.add_paragraph(f"G2 and G3 have Pearson correlation {d['g2_g3_corr']:.3f}; absences and G3 correlation is {d['absences_g3_corr']:.3f}; recorded failures and G3 correlation is {d['failures_g3_corr']:.3f}. These are descriptive pairwise relationships and do not establish cause.")
    doc.add_heading("4.2 Week 2 — Visual Storytelling",2)
    doc.add_paragraph("Charts are arranged as a narrative: establish the overall grade distribution, compare grade history and context, then show the prior-failure contrast and the limitations of predictive modeling. Captions explain the analytical purpose, and the report distinguishes observations from recommendations.")
    doc.add_heading("4.3 Week 3 — Hypothesis Test",2)
    doc.add_paragraph("H0: μ(no prior failures) ≤ μ(one or more failures). H1: μ(no prior failures) > μ(one or more failures). The one-sided Welch independent-samples t-test was specified as primary at α=0.05; Welch's method avoids assuming equal group variances. A two-sided 95% confidence interval for the mean difference and Hedges' g are reported. Mann–Whitney U is a sensitivity analysis; chi-square tests association with threshold attainment; Kruskal–Wallis compares full grade distributions across study-time categories.")
    doc.add_paragraph(f"Primary result: t={pt['t_statistic']:.3f}, df={pt['degrees_freedom']:.2f}, p={pt['p_value_one_sided']:.3g}; difference={pt['mean_difference']:.2f}, 95% CI [{pt['ci95_two_sided'][0]:.2f}, {pt['ci95_two_sided'][1]:.2f}], Hedges' g={pt['hedges_g']:.2f}. Decision: {pt['decision']}. Mann–Whitney p={mw['p_value_one_sided']:.3g}. Chi-square p={chi['p_value']:.3g}, Cramér's V={chi['cramers_v']:.2f}. Kruskal–Wallis by study time p={kw['p_value']:.3f}; at α=0.05 it does not cross the significance threshold.")
    doc.add_heading("4.4 Week 4 — Machine Learning",2)
    doc.add_paragraph("The binary target is 1 when G3 ≥ 10 and 0 otherwise. G1, G2 and G3 are excluded from predictors so the model does not rely on earlier grades that may not be available in a prospective prediction setting. Numeric variables use median imputation and scaling; categorical variables use most-frequent imputation and one-hot encoding inside Scikit-learn pipelines. Models are selected using stratified five-fold CV ROC-AUC on training data, then evaluated on a held-out stratified test split.")
    doc.add_paragraph("Candidate models: Dummy prior baseline, Logistic Regression, Decision Tree and Random Forest. The chosen model is Logistic Regression. Training CV mean ROC-AUC is {0:.3f}; test accuracy {1:.3f}, balanced accuracy {2:.3f}, precision {3:.3f}, recall {4:.3f}, F1 {5:.3f}, and ROC-AUC {6:.3f}.".format(
      ml["selected_cv_roc_auc_mean"],ml["test_accuracy"],ml["test_balanced_accuracy"],ml["test_precision"],ml["test_recall"],ml["test_f1"],ml["test_roc_auc"]))
    doc.add_paragraph("The test-set ROC-AUC is modest and therefore does not indicate strong discrimination. Cross-validation scores are more favorable than the final held-out score, illustrating how performance estimates can vary across splits and why external validation is essential.")

    doc.add_heading("5. Evidence-Based Visualizations",1)
    descriptions=[
      ("Figure 1. Final-grade distribution","Establishes the outcome scale, spread and reference threshold before more complex results."),
      ("Figure 2. Grade distributions across assessment periods","Compares medians, quartiles and means to show cohort patterns without assuming all students improved."),
      ("Figure 3. Mean G3 by previous-failure group","Connects a key EDA pattern to the primary Welch test; uncertainty and group sizes are shown."),
      ("Figure 4. Correlation heatmap","Summarizes pairwise linear relationships and reinforces why G1/G2 are excluded from the prospective ML feature set."),
      ("Figure 5. Threshold attainment by previous-failure group","Communicates the magnitude of the group association in a simple percentage format."),
      ("Figure 6. Cross-validation model comparison","Shows how candidate algorithms compare using training-only validation, with fold variability."),
      ("Figure 7. Test confusion matrix","Shows false positives and false negatives on the held-out test cases."),
      ("Figure 8. Test ROC curve","Shows the classifier's trade-off across thresholds and its limited ranking performance."),
      ("Figure 9. Permutation importance","Describes which input variables the selected model relies on; this is not causal importance.")
    ]
    for idx,(img,(title,why)) in enumerate(zip(images,descriptions)):
        doc.add_heading(title,2)
        doc.add_picture(str(img),width=Inches(5.75))
        doc.paragraphs[-1].alignment=WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("Interpretation: "+why)
        if idx==2:
            doc.add_paragraph(f"Observed means are {pt['mean_no_prior_failures']:.2f} with no prior failures and {pt['mean_one_or_more_failures']:.2f} with one or more failures. The difference is statistically significant in this sample, but confounding prevents a causal conclusion.")
        elif idx==7:
            doc.add_paragraph("Test ROC-AUC of {:.3f} is only modestly above 0.50 chance-level ranking. This is a central limitation, not a success claim.".format(ml["test_roc_auc"]))

    doc.add_heading("6. Strategic Recommendations",1)
    recommendations=[
      ("Recommendation 1 — Supportive early review for students with prior failures","The no-failure group has mean G3 {:.2f}, compared with {:.2f} in the prior-failure group; the difference is {:.2f} points and the primary test p-value is {:.2g}. Use this pattern as a reason to offer a supportive check-in, not as a label or deterministic prediction.".format(pt["mean_no_prior_failures"],pt["mean_one_or_more_failures"],pt["mean_difference"],pt["p_value_one_sided"])),
      ("Recommendation 2 — Do not deploy the current model for automated decisions","Test ROC-AUC is {:.3f} and balanced accuracy is {:.3f}. These scores show limited discrimination on the held-out data. Treat outputs only as a prototype for learning and evaluation.".format(ml["test_roc_auc"],ml["test_balanced_accuracy"])),
      ("Recommendation 3 — Improve data coverage and validation","Collect appropriately governed data from more schools and later cohorts, retain source definitions, document missingness and verify data quality before re-estimating models. Compare performance across institutions and time periods."),
      ("Recommendation 4 — Measure errors and fairness before any operational use","Audit false negatives, false positives, calibration, and subgroup error rates. If an educational support workflow is ever explored, involve educators and students, protect privacy, and keep a human review process."),
      ("Recommendation 5 — Keep analyses reproducible and interpretation cautious","Preserve the code, requirements, run metadata, cleaning audit, test statistics, model comparison, and charts. Report effect sizes and confidence intervals alongside p-values, and clearly separate exploratory questions from a pre-specified primary hypothesis.")
    ]
    for head,body in recommendations:
        doc.add_heading(head,2);doc.add_paragraph(body)
    doc.add_heading("7. Limitations and Future Improvement",1)
    for item in [
      "The data are observational; neither hypothesis testing nor feature importance proves causal effects.",
      "The dataset represents two Portuguese secondary schools and may not generalize to other populations or time periods.",
      "The outcome threshold G3 ≥ 10 is an analytical convention. It must not be confused with all institutions' official passing rules.",
      "The ML dataset is relatively small. A single held-out test split can yield noisy estimates; external validation and uncertainty intervals should be added before practical use.",
      "Prior failures are predictive context but may overlap with historical educational disadvantage. Interpreting them punitively risks unfairness.",
      "Potential improvements include a larger external dataset, repeated stratified splits or nested CV, calibration checks, threshold analysis based on support costs, subgroup fairness evaluation, and pre-registered hypotheses."
    ]: doc.add_paragraph(item,style="List Bullet")
    doc.add_heading("8. Conclusion",1)
    doc.add_paragraph("Across Weeks 1–4, the project demonstrates a complete workflow from acquisition and cleaning through EDA, statistical inference, and model evaluation. The most reliable insight is the strong observed association between prior-failure history and final-grade outcomes; the predictive model using non-grade variables is much less compelling. The responsible next step is supportive academic review and improved data collection, followed by independent model validation—not automated decisions about students.")
    doc.add_heading("9. Reproducibility and Deliverables",1)
    doc.add_paragraph("Run python week5_capstone.py after installing requirements.txt. The script writes the capstone DOCX, a separate Word version of the submission description, integrated visualizations, cleaned data, source/run metadata, statistical test results, model cross-validation comparisons, held-out metrics, confusion-matrix values, and permutation importance output. The short excerpts below highlight target/leakage control, model-selection cross-validation, and the primary hypothesis test; the full implementation is in week5_capstone.py.")
    doc.add_heading("Illustrative code excerpts",2)
    p=doc.add_paragraph("Target construction and leakage control:"); p.paragraph_format.keep_with_next=True
    p=doc.add_paragraph()
    p.paragraph_format.left_indent=Inches(.2)
    r=p.add_run('df["target_pass"] = (df["G3"] >= 10).astype(int)\nX = df.drop(columns=["G1", "G2", "G3", "target_pass"])')
    r.font.name="Consolas"; r.font.size=Pt(8)
    p=doc.add_paragraph("Training-only cross-validation for model selection:"); p.paragraph_format.keep_with_next=True
    p=doc.add_paragraph()
    p.paragraph_format.left_indent=Inches(.2)
    r=p.add_run('cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)\nscores = cross_validate(model, X_train, y_train, cv=cv, scoring="roc_auc")')
    r.font.name="Consolas"; r.font.size=Pt(8)
    p=doc.add_paragraph("Primary hypothesis test:"); p.paragraph_format.keep_with_next=True
    p=doc.add_paragraph()
    p.paragraph_format.left_indent=Inches(.2)
    r=p.add_run('test = stats.ttest_ind(no_failures, prior_failures,\n                         equal_var=False, alternative="greater")')
    r.font.name="Consolas"; r.font.size=Pt(8)

    doc.add_heading("10. References",1)
    doc.add_paragraph("Cortez, P. (2008). Student Performance [Dataset]. UCI Machine Learning Repository. https://doi.org/10.24432/C5TG7T")
    doc.add_paragraph("Cortez, P., & Silva, A. M. G. (2008). Using data mining to predict secondary school student performance. Proceedings of the 5th Annual Future Business Technology Conference.")
    doc.add_paragraph("Pedregosa, F., et al. (2011). Scikit-learn: Machine Learning in Python. Journal of Machine Learning Research, 12, 2825–2830.")
    doc.add_paragraph("Dataset page: "+UCI_PAGE)
    doc.save(REPORT)

def main():
    OUT.mkdir(parents=True,exist_ok=True);PLOTS.mkdir(parents=True,exist_ok=True)
    raw,source=acquire()
    df,audit=prepare(raw)
    st=stats_analysis(df)
    model,X,y,X_test,y_test,pred,prob,cv_table,ml,importance=ml_analysis(df)
    # Harmonize metrics used by chart and report narratives.
    ml["selected_cv_roc_auc_mean"]=float(cv_table.loc[cv_table.model==ml["selected_model"],"cv_auc_mean"].iloc[0])
    images=charts(df,st,ml,model,X_test,y_test,pred,prob,cv_table,importance)
    df.to_csv(OUT/"capstone_analysis_data.csv",index=False)
    cv_table.to_csv(OUT/"model_cross_validation_comparison.csv",index=False)
    (OUT/"integrated_statistical_results.json").write_text(json.dumps(st,indent=2),encoding="utf-8")
    (OUT/"model_evaluation_metrics.json").write_text(json.dumps(ml,indent=2),encoding="utf-8")
    (OUT/"data_quality_audit.json").write_text(json.dumps(audit,indent=2),encoding="utf-8")
    (OUT/"run_metadata.json").write_text(json.dumps({"dataset":"UCI Student Performance — Mathematics","source":source,"source_url":UCI_ZIP,"dataset_page":UCI_PAGE,"doi":DOI,"n_records":len(df),"synthetic_fallback_used":False,"report":REPORT.name,"n_visualizations":len(images),"visualizations":[p.name for p in images]},indent=2),encoding="utf-8")
    build_report(df,audit,st,ml,cv_table,importance,images,source)
    build_submission_doc()
    print("Official UCI dataset verified; synthetic fallback not used.")
    print(f"Records={len(df)}; mean G3={df.G3.mean():.3f}; threshold rate={(df.target_pass.mean()*100):.2f}%")
    print(f"Primary Welch p={st['primary_test']['p_value_one_sided']:.5g}; mean difference={st['primary_test']['mean_difference']:.3f}; CI={st['primary_test']['ci95_two_sided']}")
    print(f"Selected model={ml['selected_model']}; test accuracy={ml['test_accuracy']:.3f}; test ROC-AUC={ml['test_roc_auc']:.3f}")
    print(f"Generated report, submission description, {len(images)} charts, and metrics outputs.")

if __name__=="__main__":
    main()
