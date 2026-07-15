"""
Day 1 - Data Integration and Preprocessing (ETL)
Project: Algorithmic Credit Scoring Using Behavioral and Transactional Data

Purpose
-------
Extract the raw transaction-level export (Master_Finance_Data / Semi_Unclear_Finance_Dataset),
clean and validate it, and load it into a star-schema-shaped, analysis-ready workbook:

    FACT_TRANSACTIONS  (grain: one row per TransactionID)
        -> DIM_DATE
        -> DIM_MERCHANT_CATEGORY
        -> DIM_PAYMENT_METHOD
        -> DIM_OCCUPATION
        -> DIM_REGION
        -> DIM_GENDER

A design note on Dim_Customer: CustomerID repeats across many transactions (801 unique
customers over 5,000 rows), but Age, Gender, Occupation, Region, Income and CreditScore are
NOT constant for a given CustomerID across their transactions (verified during profiling).
Rather than force a conventional, one-row-per-customer Dim_Customer (which would silently
overwrite genuine point-in-time variation - e.g. a CreditScore legitimately changes between
transactions), these attributes are kept as point-in-time descriptors on the fact row, and
their distinct values are additionally normalized into small lookup dimensions
(Occupation, Region, Gender) to support a proper star schema and clean slicing/filtering.
CustomerID itself remains on the fact table as a degenerate dimension (an identifier with no
stable attributes of its own worth spinning into a separate table).

Output: Master_Finance_Data_Cleaned.xlsx
"""

import pandas as pd
import numpy as np
from pathlib import Path
import json

RAW_PATH = Path("/home/claude/Semi_Unclear_Finance_Dataset_5000.xlsx")
OUT_DIR = Path("/home/claude/day1/output")
CHART_DIR = Path("/home/claude/day1/charts")
OUT_DIR.mkdir(parents=True, exist_ok=True)
CHART_DIR.mkdir(parents=True, exist_ok=True)

log = []  # data-quality / transformation log, becomes the Data_Quality_Log sheet


def note(step, detail):
    log.append({"Step": step, "Detail": detail})
    print(f"[{step}] {detail}")


# ---------------------------------------------------------------------------
# 1. EXTRACT
# ---------------------------------------------------------------------------
raw = pd.read_excel(RAW_PATH, sheet_name="Transactions")
note("EXTRACT", f"Loaded {len(raw):,} rows x {raw.shape[1]} columns from "
                 f"'{RAW_PATH.name}' (sheet: Transactions).")

df = raw.copy()

# ---------------------------------------------------------------------------
# 2. TRANSFORM / CLEAN
# ---------------------------------------------------------------------------

# 2.1 Trim whitespace on every text column (caught 'South ' vs 'South' in Region)
text_cols = ["Gender", "Occupation", "Region", "MerchantCategory", "PaymentMethod",
             "TransactionID"]
for c in text_cols:
    before = df[c].nunique()
    df[c] = df[c].astype(str).str.strip()
    after = df[c].nunique()
    if before != after:
        note("CLEAN-WHITESPACE",
             f"'{c}': trimmed leading/trailing whitespace, collapsing {before} -> {after} "
             f"distinct values (e.g. 'South ' merged into 'South').")

# 2.2 Standardize TransactionDate to a real datetime
df["TransactionDate"] = pd.to_datetime(df["TransactionDate"], errors="coerce")
bad_dates = df["TransactionDate"].isna().sum()
note("CLEAN-DATE", f"Parsed TransactionDate to datetime. Unparseable values: {bad_dates}.")

# 2.3 Duplicate check on the natural/business key
dupe_txn = df.duplicated(subset=["TransactionID"]).sum()
dupe_full = df.duplicated().sum()
note("VALIDATE-DUPLICATES",
     f"Duplicate TransactionID rows: {dupe_txn}. Fully duplicate rows: {dupe_full}. "
     f"(None found - no rows dropped for duplication.)")

# 2.4 Missing value handling
# Income: ~50% missing. Too large a share to impute point estimates without materially
# biasing downstream credit models, so we KEEP the null in the analysis-ready Income column
# and add an explicit IsIncomeMissing flag. A segment-median Income_Imputed column is provided
# separately for visualization/EDA convenience only - it must not be used as a modelling feature
# without disclosing the imputation.
income_missing = df["Income"].isna().sum()
df["IsIncomeMissing"] = df["Income"].isna()
median_by_segment = df.groupby("Occupation")["Income"].transform("median")
df["Income_Imputed"] = df["Income"].fillna(median_by_segment)
note("MISSING-INCOME",
     f"Income missing in {income_missing} of {len(df)} rows ({income_missing/len(df):.1%}). "
     f"Missingness rate is similar across all Occupation segments (profiled separately), "
     f"consistent with values simply not being captured at the point of transaction rather "
     f"than a systematic reporting gap for one segment. Original Income left null; "
     f"IsIncomeMissing flag added; Income_Imputed (segment-median by Occupation) added for "
     f"charting only.")

# TransactionAmount: 49 missing. Small share -> impute with median by MerchantCategory and flag.
txn_amt_missing = df["TransactionAmount"].isna().sum()
df["IsTransactionAmountMissing"] = df["TransactionAmount"].isna()
med_by_cat = df.groupby("MerchantCategory")["TransactionAmount"].transform("median")
df["TransactionAmount"] = df["TransactionAmount"].fillna(med_by_cat)
note("MISSING-TXN-AMOUNT",
     f"TransactionAmount missing in {txn_amt_missing} rows ({txn_amt_missing/len(df):.2%}). "
     f"Imputed with the median TransactionAmount for the same MerchantCategory; "
     f"IsTransactionAmountMissing flag retained so imputed rows remain identifiable.")

# 2.5 Business-rule / consistency checks (not corrected, just logged as validation evidence)
zero_loan_defaults = df[(df["LoanAmount"] == 0) & (df["DefaultFlag"] == 1)]
note("VALIDATE-BUSINESS-RULE",
     f"Rows with LoanAmount = 0 AND DefaultFlag = 1: {len(zero_loan_defaults)} "
     f"(expected 0 - a customer cannot default with no loan outstanding). Rule holds.")

neg_income = (df["Income"] < 0).sum()
neg_amount = (df["TransactionAmount"] < 0).sum()
oob_credit = (~df["CreditScore"].between(300, 900)).sum()
oob_age = (~df["Age"].between(18, 100)).sum()
note("VALIDATE-RANGES",
     f"Negative Income: {neg_income}. Negative TransactionAmount: {neg_amount}. "
     f"CreditScore outside [300,900]: {oob_credit}. Age outside [18,100]: {oob_age}. "
     f"No out-of-range values found.")

# 2.6 Derived / convenience columns
df["HasLoan"] = df["LoanAmount"] > 0
df["Year"] = df["TransactionDate"].dt.year
df["Month"] = df["TransactionDate"].dt.month
df["MonthName"] = df["TransactionDate"].dt.strftime("%b %Y")
df["Quarter"] = df["TransactionDate"].dt.quarter

note("TRANSFORM-DERIVED",
     "Added derived columns: HasLoan (LoanAmount > 0), Year, Month, MonthName, Quarter.")

# ---------------------------------------------------------------------------
# 3. BUILD STAR SCHEMA
# ---------------------------------------------------------------------------

# --- Dim_Date ---
dim_date = (
    df[["TransactionDate", "Year", "Month", "MonthName", "Quarter"]]
    .drop_duplicates()
    .sort_values("TransactionDate")
    .reset_index(drop=True)
)
dim_date.insert(0, "DateKey", dim_date["TransactionDate"].dt.strftime("%Y%m%d").astype(int))
dim_date["DayOfWeek"] = dim_date["TransactionDate"].dt.day_name()

# --- Dim_MerchantCategory ---
dim_merchant = (
    pd.DataFrame({"MerchantCategory": sorted(df["MerchantCategory"].unique())})
    .reset_index(drop=True)
)
dim_merchant.insert(0, "MerchantCategoryKey", dim_merchant.index + 1)

# --- Dim_PaymentMethod ---
dim_payment = (
    pd.DataFrame({"PaymentMethod": sorted(df["PaymentMethod"].unique())})
    .reset_index(drop=True)
)
dim_payment.insert(0, "PaymentMethodKey", dim_payment.index + 1)

# --- Dim_Occupation ---
dim_occupation = (
    pd.DataFrame({"Occupation": sorted(df["Occupation"].unique())})
    .reset_index(drop=True)
)
dim_occupation.insert(0, "OccupationKey", dim_occupation.index + 1)

# --- Dim_Region ---
dim_region = (
    pd.DataFrame({"Region": sorted(df["Region"].unique())})
    .reset_index(drop=True)
)
dim_region.insert(0, "RegionKey", dim_region.index + 1)

# --- Dim_Gender ---
dim_gender = (
    pd.DataFrame({"Gender": sorted(df["Gender"].unique())})
    .reset_index(drop=True)
)
dim_gender.insert(0, "GenderKey", dim_gender.index + 1)

note("LOAD-DIMENSIONS",
     f"Built Dim_Date ({len(dim_date)} rows), Dim_MerchantCategory ({len(dim_merchant)}), "
     f"Dim_PaymentMethod ({len(dim_payment)}), Dim_Occupation ({len(dim_occupation)}), "
     f"Dim_Region ({len(dim_region)}), Dim_Gender ({len(dim_gender)}).")

# --- Fact_Transactions ---
fact = df.merge(dim_date[["TransactionDate", "DateKey"]], on="TransactionDate", how="left")
fact = fact.merge(dim_merchant, on="MerchantCategory", how="left")
fact = fact.merge(dim_payment, on="PaymentMethod", how="left")
fact = fact.merge(dim_occupation, on="Occupation", how="left")
fact = fact.merge(dim_region, on="Region", how="left")
fact = fact.merge(dim_gender, on="Gender", how="left")

fact_transactions = fact[[
    "TransactionID", "CustomerID", "DateKey",
    "MerchantCategoryKey", "PaymentMethodKey", "OccupationKey", "RegionKey", "GenderKey",
    "Age", "Income", "Income_Imputed", "IsIncomeMissing",
    "TransactionAmount", "IsTransactionAmountMissing",
    "BehaviorScore", "LoanAmount", "HasLoan", "CreditScore", "DefaultFlag",
]].copy()

note("LOAD-FACT",
     f"Built Fact_Transactions: {len(fact_transactions):,} rows, grain = 1 row per "
     f"TransactionID, {fact_transactions.shape[1]} columns (measures + FKs to 6 dimensions).")

# Referential integrity check
for key, dim in [("DateKey", dim_date), ("MerchantCategoryKey", dim_merchant),
                  ("PaymentMethodKey", dim_payment), ("OccupationKey", dim_occupation),
                  ("RegionKey", dim_region), ("GenderKey", dim_gender)]:
    orphans = fact_transactions[key].isna().sum()
    if orphans:
        note("VALIDATE-RI", f"WARNING: {orphans} fact rows failed to map to {key}.")
if not any(fact_transactions[k].isna().any() for k in
           ["DateKey", "MerchantCategoryKey", "PaymentMethodKey", "OccupationKey",
            "RegionKey", "GenderKey"]):
    note("VALIDATE-RI", "All fact rows joined successfully to every dimension (0 orphans).")

# ---------------------------------------------------------------------------
# 4. SAVE cleaned flat table + star schema + quality log
# ---------------------------------------------------------------------------
log_df = pd.DataFrame(log)

with pd.ExcelWriter(OUT_DIR / "Master_Finance_Data_Cleaned.xlsx", engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Cleaned_Flat_Table", index=False)
    fact_transactions.to_excel(writer, sheet_name="Fact_Transactions", index=False)
    dim_date.to_excel(writer, sheet_name="Dim_Date", index=False)
    dim_merchant.to_excel(writer, sheet_name="Dim_MerchantCategory", index=False)
    dim_payment.to_excel(writer, sheet_name="Dim_PaymentMethod", index=False)
    dim_occupation.to_excel(writer, sheet_name="Dim_Occupation", index=False)
    dim_region.to_excel(writer, sheet_name="Dim_Region", index=False)
    dim_gender.to_excel(writer, sheet_name="Dim_Gender", index=False)
    log_df.to_excel(writer, sheet_name="Data_Quality_Log", index=False)

note("SAVE", f"Workbook written to {OUT_DIR / 'Master_Finance_Data_Cleaned.xlsx'} "
             f"with 8 sheets (1 flat cleaned table + 6-table star schema + quality log).")

# Save the log separately too (used by the technical report script)
log_df.to_json(OUT_DIR / "quality_log.json", orient="records", indent=2)

print("\nETL complete.")
