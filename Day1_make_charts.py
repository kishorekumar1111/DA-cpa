import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

plt.rcParams["font.family"] = "DejaVu Sans"
CHART_DIR = "/home/claude/day1/charts"

df = pd.read_excel("/home/claude/day1/output/Master_Finance_Data_Cleaned.xlsx", sheet_name="Cleaned_Flat_Table")

COLORS = ["#2E5266", "#6E8898", "#9FB8AD", "#D7B377", "#BF4E30", "#4C6444"]

# 1. Missing values before/after
fig, ax = plt.subplots(figsize=(7,4.5))
raw = pd.read_excel("/home/claude/Semi_Unclear_Finance_Dataset_5000.xlsx", sheet_name="Transactions")
missing_before = raw.isnull().sum()
missing_before = missing_before[missing_before > 0]
ax.bar(missing_before.index, missing_before.values, color="#BF4E30")
for i, v in enumerate(missing_before.values):
    ax.text(i, v + 30, f"{v}\n({v/len(raw):.1%})", ha="center", fontsize=9)
ax.set_title("Missing Values in Raw Dataset (5,000 rows)", fontsize=13, fontweight="bold")
ax.set_ylabel("Missing count")
ax.spines[["top","right"]].set_visible(False)
plt.tight_layout()
plt.savefig(f"{CHART_DIR}/01_missing_values.png", dpi=150)
plt.close()

# 2. Region distribution before/after cleaning
fig, axes = plt.subplots(1, 2, figsize=(10, 4.5))
raw["Region"].value_counts().plot(kind="bar", ax=axes[0], color="#BF4E30")
axes[0].set_title("Region - Before Cleaning\n(5 categories incl. 'South ' dup)", fontsize=11)
axes[0].tick_params(axis='x', rotation=30)
df["Region"].value_counts().plot(kind="bar", ax=axes[1], color="#2E5266")
axes[1].set_title("Region - After Trim & Standardize\n(4 categories)", fontsize=11)
axes[1].tick_params(axis='x', rotation=30)
for a in axes:
    a.spines[["top","right"]].set_visible(False)
plt.tight_layout()
plt.savefig(f"{CHART_DIR}/02_region_before_after.png", dpi=150)
plt.close()

# 3. Income distribution (imputed) by occupation
fig, ax = plt.subplots(figsize=(7.5,4.5))
occ_order = df.groupby("Occupation")["Income_Imputed"].median().sort_values().index
data = [df.loc[df.Occupation==o, "Income_Imputed"].dropna() for o in occ_order]
bp = ax.boxplot(data, labels=occ_order, patch_artist=True)
for patch, c in zip(bp['boxes'], COLORS):
    patch.set_facecolor(c)
ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x/1000:.0f}K"))
ax.set_title("Income Distribution by Occupation (median-imputed)", fontsize=13, fontweight="bold")
ax.set_ylabel("Income (₹)")
ax.spines[["top","right"]].set_visible(False)
plt.tight_layout()
plt.savefig(f"{CHART_DIR}/03_income_by_occupation.png", dpi=150)
plt.close()

# 4. Default rate by occupation and HasLoan
fig, ax = plt.subplots(figsize=(7.5,4.5))
default_rate = df[df.HasLoan].groupby("Occupation")["DefaultFlag"].mean().sort_values(ascending=False)
ax.bar(default_rate.index, default_rate.values*100, color=COLORS[:len(default_rate)])
for i, v in enumerate(default_rate.values):
    ax.text(i, v*100+0.5, f"{v:.1%}", ha="center", fontsize=9)
ax.set_title("Default Rate by Occupation (loan-holders only)", fontsize=13, fontweight="bold")
ax.set_ylabel("Default rate (%)")
ax.spines[["top","right"]].set_visible(False)
plt.tight_layout()
plt.savefig(f"{CHART_DIR}/04_default_rate_by_occupation.png", dpi=150)
plt.close()

# 5. Transactions over time (monthly)
fig, ax = plt.subplots(figsize=(9,4.5))
monthly = df.groupby(df["TransactionDate"].dt.to_period("M")).size()
monthly.index = monthly.index.astype(str)
ax.plot(monthly.index, monthly.values, marker="o", color="#2E5266", linewidth=2)
ax.set_title("Transaction Volume by Month", fontsize=13, fontweight="bold")
ax.set_ylabel("Number of transactions")
ax.tick_params(axis='x', rotation=60)
ax.spines[["top","right"]].set_visible(False)
plt.tight_layout()
plt.savefig(f"{CHART_DIR}/05_monthly_volume.png", dpi=150)
plt.close()

# 6. Credit score distribution
fig, ax = plt.subplots(figsize=(7.5,4.5))
ax.hist(df["CreditScore"], bins=30, color="#2E5266", edgecolor="white")
ax.axvline(df["CreditScore"].median(), color="#BF4E30", linestyle="--", linewidth=2,
           label=f"Median = {df['CreditScore'].median():.0f}")
ax.set_title("Credit Score Distribution (all transactions)", fontsize=13, fontweight="bold")
ax.set_xlabel("Credit Score")
ax.set_ylabel("Frequency")
ax.legend()
ax.spines[["top","right"]].set_visible(False)
plt.tight_layout()
plt.savefig(f"{CHART_DIR}/06_credit_score_dist.png", dpi=150)
plt.close()

print("Charts created:")
import os
for f in sorted(os.listdir(CHART_DIR)):
    print(" -", f)
