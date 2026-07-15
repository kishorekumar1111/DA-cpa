"""
Day 2 - Descriptive Statistics & Exploratory Data Analysis (EDA)
Project: Algorithmic Credit Scoring Using Behavioral and Transactional Data

Purpose
-------
Run a comprehensive EDA on the Day 1 cleaned dataset: descriptive statistics for every
numeric and categorical field, an outlier scan (IQR method), and a correlation study
between the behavioral/transactional variables and DefaultFlag - the variable this whole
project ultimately wants to predict.

Output: Master_Finance_Data_EDA.xlsx
"""

import pandas as pd
import numpy as np
from pathlib import Path

SRC = Path("/home/claude/day1/output/Master_Finance_Data_Cleaned.xlsx")
OUT_DIR = Path("/home/claude/day2/output")
OUT_DIR.mkdir(parents=True, exist_ok=True)

log = []
def note(step, detail):
    log.append({"Step": step, "Detail": detail})
    print(f"[{step}] {detail}")

# ---------------------------------------------------------------------------
# 1. LOAD
# ---------------------------------------------------------------------------
df = pd.read_excel(SRC, sheet_name="Cleaned_Flat_Table")
note("EXTRACT", f"Loaded {len(df):,} cleaned transactions from Day 1 output ('{SRC.name}').")

NUMERIC_COLS = ["Age", "Income", "TransactionAmount", "BehaviorScore", "LoanAmount",
                 "CreditScore", "DefaultFlag"]
CATEGORICAL_COLS = ["Gender", "Occupation", "Region", "MerchantCategory", "PaymentMethod"]

# ---------------------------------------------------------------------------
# 2. DESCRIPTIVE STATISTICS - NUMERIC
# ---------------------------------------------------------------------------
desc_rows = []
for col in NUMERIC_COLS:
    s = df[col].dropna()
    desc_rows.append({
        "Column": col,
        "Count": s.count(),
        "Missing": df[col].isna().sum(),
        "Mean": round(s.mean(), 2),
        "Median": round(s.median(), 2),
        "StdDev": round(s.std(), 2),
        "Min": round(s.min(), 2),
        "P25": round(s.quantile(0.25), 2),
        "P75": round(s.quantile(0.75), 2),
        "Max": round(s.max(), 2),
        "Skew": round(s.skew(), 3),
    })
desc_numeric = pd.DataFrame(desc_rows)
note("DESCRIBE-NUMERIC", f"Computed summary statistics for {len(NUMERIC_COLS)} numeric columns.")

# ---------------------------------------------------------------------------
# 3. DESCRIPTIVE STATISTICS - CATEGORICAL
# ---------------------------------------------------------------------------
cat_rows = []
for col in CATEGORICAL_COLS:
    vc = df[col].value_counts()
    for val, count in vc.items():
        cat_rows.append({"Column": col, "Value": val, "Count": count,
                          "Percent": round(count / len(df) * 100, 2)})
desc_categorical = pd.DataFrame(cat_rows)
note("DESCRIBE-CATEGORICAL", f"Computed value-count distributions for {len(CATEGORICAL_COLS)} "
                              f"categorical columns.")

# ---------------------------------------------------------------------------
# 4. OUTLIER SCAN (IQR method, 1.5x whisker rule)
# ---------------------------------------------------------------------------
outlier_rows = []
for col in ["Income", "TransactionAmount", "LoanAmount", "BehaviorScore", "CreditScore", "Age"]:
    s = df[col].dropna()
    q1, q3 = s.quantile(0.25), s.quantile(0.75)
    iqr = q3 - q1
    lo, hi = q1 - 1.5 * iqr, q3 + 1.5 * iqr
    n_outliers = ((s < lo) | (s > hi)).sum()
    outlier_rows.append({
        "Column": col, "Q1": round(q1, 2), "Q3": round(q3, 2), "IQR": round(iqr, 2),
        "LowerBound": round(lo, 2), "UpperBound": round(hi, 2),
        "OutlierCount": n_outliers, "OutlierPercent": round(n_outliers / len(s) * 100, 2),
    })
outliers = pd.DataFrame(outlier_rows)
note("OUTLIER-SCAN", f"Ran IQR (1.5x) outlier scan on 6 numeric columns. "
                      f"Highest outlier rate: "
                      f"{outliers.loc[outliers.OutlierPercent.idxmax(), 'Column']} "
                      f"({outliers.OutlierPercent.max()}%).")

# ---------------------------------------------------------------------------
# 5. CORRELATION MATRIX (numeric variables, incl. DefaultFlag as the target)
# ---------------------------------------------------------------------------
corr = df[NUMERIC_COLS].corr().round(3)
note("CORRELATION", "Computed Pearson correlation matrix across all numeric variables "
                     "(including DefaultFlag as the target variable).")

default_corr = (
    corr["DefaultFlag"].drop("DefaultFlag").sort_values(key=abs, ascending=False)
    .reset_index().rename(columns={"index": "Variable", "DefaultFlag": "CorrelationWithDefault"})
)
note("CORRELATION-TARGET", f"Strongest correlate of DefaultFlag: "
                            f"{default_corr.iloc[0]['Variable']} "
                            f"(r={default_corr.iloc[0]['CorrelationWithDefault']}).")

# ---------------------------------------------------------------------------
# 6. SEGMENT CUTS relevant to a credit-scoring EDA
# ---------------------------------------------------------------------------
default_by_occupation = df.groupby("Occupation")["DefaultFlag"].agg(["mean", "count"]).round(3)
default_by_occupation.columns = ["DefaultRate", "TransactionCount"]
default_by_occupation = default_by_occupation.sort_values("DefaultRate", ascending=False).reset_index()

default_by_region = df.groupby("Region")["DefaultFlag"].agg(["mean", "count"]).round(3)
default_by_region.columns = ["DefaultRate", "TransactionCount"]
default_by_region = default_by_region.sort_values("DefaultRate", ascending=False).reset_index()

income_bracket = pd.cut(df["Income_Imputed"], bins=[0, 25000, 50000, 100000, 200000, np.inf],
                         labels=["<25K", "25K-50K", "50K-100K", "100K-200K", "200K+"])
default_by_income = df.groupby(income_bracket, observed=True)["DefaultFlag"].agg(["mean", "count"]).round(3)
default_by_income.columns = ["DefaultRate", "TransactionCount"]
default_by_income = default_by_income.reset_index().rename(columns={"Income_Imputed": "IncomeBracket"})

note("SEGMENT-CUTS", "Built DefaultFlag rate cuts by Occupation, Region and Income bracket "
                      "for targeted risk profiling.")

# ---------------------------------------------------------------------------
# 7. SAVE
# ---------------------------------------------------------------------------
log_df = pd.DataFrame(log)
with pd.ExcelWriter(OUT_DIR / "Master_Finance_Data_EDA.xlsx", engine="openpyxl") as writer:
    desc_numeric.to_excel(writer, sheet_name="Descriptive_Numeric", index=False)
    desc_categorical.to_excel(writer, sheet_name="Descriptive_Categorical", index=False)
    outliers.to_excel(writer, sheet_name="Outlier_Scan", index=False)
    corr.to_excel(writer, sheet_name="Correlation_Matrix")
    default_corr.to_excel(writer, sheet_name="Correlation_vs_Default", index=False)
    default_by_occupation.to_excel(writer, sheet_name="Default_by_Occupation", index=False)
    default_by_region.to_excel(writer, sheet_name="Default_by_Region", index=False)
    default_by_income.to_excel(writer, sheet_name="Default_by_IncomeBracket", index=False)
    log_df.to_excel(writer, sheet_name="Methodology_Log", index=False)

note("SAVE", f"Workbook written to {OUT_DIR / 'Master_Finance_Data_EDA.xlsx'} with 9 sheets.")

print("\nDay 2 EDA complete.")
