import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

plt.rcParams["font.family"] = "DejaVu Sans"
CHART_DIR = "/home/claude/day2/charts"

df = pd.read_excel("/home/claude/day1/output/Master_Finance_Data_Cleaned.xlsx", sheet_name="Cleaned_Flat_Table")
corr = pd.read_excel("/home/claude/day2/output/Master_Finance_Data_EDA.xlsx", sheet_name="Correlation_Matrix", index_col=0)
outliers = pd.read_excel("/home/claude/day2/output/Master_Finance_Data_EDA.xlsx", sheet_name="Outlier_Scan")

COLORS = ["#2E5266", "#6E8898", "#9FB8AD", "#D7B377", "#BF4E30", "#4C6444"]

# 1. Distribution grid (4 key numeric vars)
fig, axes = plt.subplots(2, 2, figsize=(10, 7.5))
specs = [("Age", "Age (years)"), ("Income", "Income (₹)"),
         ("TransactionAmount", "Transaction Amount (₹)"), ("CreditScore", "Credit Score")]
for ax, (col, label) in zip(axes.flat, specs):
    ax.hist(df[col].dropna(), bins=30, color="#2E5266", edgecolor="white")
    ax.set_title(label, fontsize=11, fontweight="bold")
    ax.spines[["top","right"]].set_visible(False)
fig.suptitle("Distribution of Key Numeric Variables", fontsize=14, fontweight="bold")
plt.tight_layout()
plt.savefig(f"{CHART_DIR}/01_distributions_grid.png", dpi=150)
plt.close()

# 2. Boxplots for outlier columns
fig, ax = plt.subplots(figsize=(9, 5))
cols = ["TransactionAmount", "LoanAmount", "Income"]
data = [df[c].dropna() / (1000 if c != "TransactionAmount" else 1) for c in cols]
bp = ax.boxplot([df[c].dropna() for c in cols], labels=cols, patch_artist=True, showfliers=True)
for patch, c in zip(bp['boxes'], COLORS):
    patch.set_facecolor(c)
ax.set_title("Boxplots — Outlier Scan (IQR method)", fontsize=13, fontweight="bold")
ax.spines[["top","right"]].set_visible(False)
plt.tight_layout()
plt.savefig(f"{CHART_DIR}/02_boxplots_outliers.png", dpi=150)
plt.close()

# 3. Correlation heatmap
fig, ax = plt.subplots(figsize=(7, 6))
im = ax.imshow(corr.values, cmap="RdBu_r", vmin=-1, vmax=1)
ax.set_xticks(range(len(corr.columns))); ax.set_xticklabels(corr.columns, rotation=45, ha="right")
ax.set_yticks(range(len(corr.index))); ax.set_yticklabels(corr.index)
for i in range(corr.shape[0]):
    for j in range(corr.shape[1]):
        ax.text(j, i, f"{corr.values[i,j]:.2f}", ha="center", va="center", fontsize=8,
                 color="white" if abs(corr.values[i,j])>0.5 else "black")
ax.set_title("Correlation Matrix (numeric variables)", fontsize=13, fontweight="bold")
plt.colorbar(im, ax=ax, shrink=0.8)
plt.tight_layout()
plt.savefig(f"{CHART_DIR}/03_correlation_heatmap.png", dpi=150)
plt.close()

# 4. Income vs CreditScore scatter, colored by DefaultFlag
fig, ax = plt.subplots(figsize=(8, 5.5))
for flag, label, color in [(0, "No Default", "#2E5266"), (1, "Default", "#BF4E30")]:
    sub = df[(df.DefaultFlag==flag) & df.Income.notna()]
    ax.scatter(sub["Income"], sub["CreditScore"], s=14, alpha=0.35, color=color, label=label)
ax.set_xlabel("Income (₹)")
ax.set_ylabel("Credit Score")
ax.set_title("Income vs Credit Score, by Default Status", fontsize=13, fontweight="bold")
ax.legend()
ax.spines[["top","right"]].set_visible(False)
plt.tight_layout()
plt.savefig(f"{CHART_DIR}/04_income_vs_creditscore.png", dpi=150)
plt.close()

# 5. Default rate by Occupation and Region (side by side)
fig, axes = plt.subplots(1, 2, figsize=(11, 4.8))
occ = df.groupby("Occupation")["DefaultFlag"].mean().sort_values(ascending=False)
axes[0].bar(occ.index, occ.values*100, color=COLORS[:len(occ)])
axes[0].set_title("Default Rate by Occupation", fontsize=11, fontweight="bold")
axes[0].set_ylabel("Default rate (%)")
axes[0].tick_params(axis='x', rotation=30)
reg = df.groupby("Region")["DefaultFlag"].mean().sort_values(ascending=False)
axes[1].bar(reg.index, reg.values*100, color=COLORS[:len(reg)])
axes[1].set_title("Default Rate by Region", fontsize=11, fontweight="bold")
axes[1].tick_params(axis='x', rotation=30)
for a in axes:
    a.spines[["top","right"]].set_visible(False)
plt.tight_layout()
plt.savefig(f"{CHART_DIR}/05_default_rate_occ_region.png", dpi=150)
plt.close()

# 6. Category frequency: MerchantCategory & PaymentMethod
fig, axes = plt.subplots(1, 2, figsize=(11, 4.8))
mc = df["MerchantCategory"].value_counts()
axes[0].bar(mc.index, mc.values, color=COLORS[:len(mc)])
axes[0].set_title("Transactions by Merchant Category", fontsize=11, fontweight="bold")
axes[0].tick_params(axis='x', rotation=30)
pm = df["PaymentMethod"].value_counts()
axes[1].bar(pm.index, pm.values, color=COLORS[:len(pm)])
axes[1].set_title("Transactions by Payment Method", fontsize=11, fontweight="bold")
axes[1].tick_params(axis='x', rotation=30)
for a in axes:
    a.spines[["top","right"]].set_visible(False)
plt.tight_layout()
plt.savefig(f"{CHART_DIR}/06_category_frequencies.png", dpi=150)
plt.close()

print("Charts created:")
import os
for f in sorted(os.listdir(CHART_DIR)):
    print(" -", f)
