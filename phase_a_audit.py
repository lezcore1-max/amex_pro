"""
Phase A — Data Audit
AmEx Campus Challenge 2026 Round 1
"""
import warnings
warnings.filterwarnings('ignore')

import pandas as pd
import numpy as np
import json
import os

DATA_DIR = r"c:\Users\vansh\Music\amex"
CSV_PATH = os.path.join(DATA_DIR, "campus_challenge_r1_data.csv")

print("=" * 70)
print("PHASE A — DATA AUDIT")
print("=" * 70)

# ── 1. Load & basic shape ────────────────────────────────────────────────────
df = pd.read_csv(CSV_PATH)
print(f"\n[1] Shape: {df.shape}  (expected 500000 rows × 24 cols)")
print(f"    Columns: {list(df.columns)}")
assert df.shape[0] == 500_000, "Row count mismatch!"
assert 'id' in df.columns, "id column missing!"
feat_cols = [c for c in df.columns if c != 'id']
print(f"    Feature columns: {len(feat_cols)}  ({feat_cols[0]} … {feat_cols[-1]})")
print(f"    Unique IDs: {df['id'].nunique()}  (should be 500000)")

# ── 2. Data types ────────────────────────────────────────────────────────────
print("\n[2] Data types:")
print(df.dtypes.to_string())

# ── 3. Missingness ──────────────────────────────────────────────────────────
print("\n[3] Missingness (%):")
miss_pct = df[feat_cols].isnull().mean().mul(100).round(2)
print(miss_pct.to_string())

# ── 4. Descriptive stats ─────────────────────────────────────────────────────
print("\n[4] Descriptive statistics:")
desc = df[feat_cols].describe(percentiles=[.01, .05, .25, .5, .75, .95, .99]).T
desc['missing_pct'] = miss_pct
print(desc[['count','mean','std','min','1%','5%','25%','50%','75%','95%','99%','max','missing_pct']].to_string())

# ── 5. Negative-value sanity check ──────────────────────────────────────────
print("\n[5] Negative value check:")
for col in feat_cols:
    neg_cnt = (df[col] < 0).sum()
    if neg_cnt > 0:
        neg_vals = df.loc[df[col] < 0, col]
        print(f"    {col}: {neg_cnt} negative values  [min={neg_vals.min():.4f}, p5={neg_vals.quantile(.05):.4f}]")
print("    (no output above = no negatives found)")

# ── 6. Structural-zero hypothesis: f6–f10 vs f5 ─────────────────────────────
print("\n[6] Structural-zero hypothesis — f6..f10 vs f5 (total spend)")
cat_cols = ['f6', 'f7', 'f8', 'f9', 'f10']

# a) Joint missingness among category columns
miss_mask = df[cat_cols].isnull()
joint_miss_all = miss_mask.all(axis=1).sum()
print(f"    Rows where ALL of f6-f10 are missing: {joint_miss_all} ({joint_miss_all/len(df)*100:.2f}%)")
any_miss = miss_mask.any(axis=1).sum()
print(f"    Rows where ANY of f6-f10 are missing: {any_miss} ({any_miss/len(df)*100:.2f}%)")

# b) Among rows where f5 is present, is the sum of f6-f10 close to f5?
# Use only rows where both f5 and all of f6-f10 are present
complete_mask = df[['f5'] + cat_cols].notnull().all(axis=1)
df_complete = df.loc[complete_mask].copy()
df_complete['cat_sum'] = df_complete[cat_cols].sum(axis=1)
df_complete['diff_abs'] = (df_complete['f5'] - df_complete['cat_sum']).abs()
df_complete['diff_pct'] = df_complete['diff_abs'] / (df_complete['f5'].abs() + 1e-9) * 100
print(f"\n    Rows with all 5 spend cols + f5 non-null: {len(df_complete)}")
print(f"    Median |f5 - sum(cats)|: {df_complete['diff_abs'].median():.2f}")
print(f"    Median % diff: {df_complete['diff_pct'].median():.2f}%")
print(f"    % rows where diff < 1%: {(df_complete['diff_pct'] < 1).mean()*100:.1f}%")
print(f"    % rows where diff < 5%: {(df_complete['diff_pct'] < 5).mean()*100:.1f}%")

# c) Among rows where f6-f10 are missing but f5 is present — are they zero-spenders?
cats_missing_f5_present = df[df[cat_cols].isnull().all(axis=1) & df['f5'].notnull()]
if len(cats_missing_f5_present) > 0:
    print(f"\n    Rows where cats all-missing but f5 present: {len(cats_missing_f5_present)}")
    print(f"    f5 median in those rows: {cats_missing_f5_present['f5'].median():.2f}")
    print(f"    f5=0 in those rows: {(cats_missing_f5_present['f5'] == 0).sum()}")
    pct_low = (cats_missing_f5_present['f5'] < 100).mean()*100
    print(f"    f5 < 100 in those rows: {pct_low:.1f}%")

# ── 7. f17 vs f18 relationship ───────────────────────────────────────────────
print("\n[7] f17 (Total Lend Line) vs f18 (Consumer Lend Line) — redundancy check:")
both_present = df[['f17','f18']].notnull().all(axis=1)
print(f"    Both non-null: {both_present.sum()}")
if both_present.sum() > 0:
    sub = df.loc[both_present, ['f17','f18']]
    corr_val = sub['f17'].corr(sub['f18'])
    print(f"    Pearson correlation: {corr_val:.4f}")
    ratio = sub['f18'] / (sub['f17'] + 1e-9)
    print(f"    f18/f17 median: {ratio.median():.4f}, mean: {ratio.mean():.4f}")
    print(f"    f18 > f17 cases: {(sub['f18'] > sub['f17']).sum()}")

# ── 8. Correlation matrix summary (top correlated pairs) ──────────────────
print("\n[8] Top 20 most correlated feature pairs (by |r|):")
corr_mat = df[feat_cols].corr()
# Extract upper triangle
pairs = []
for i, r in enumerate(feat_cols):
    for j, c in enumerate(feat_cols):
        if j > i:
            pairs.append({'feat1': r, 'feat2': c, 'corr': corr_mat.loc[r, c]})
pairs_df = pd.DataFrame(pairs).sort_values('corr', key=abs, ascending=False)
print(pairs_df.head(20).to_string(index=False))

# ── 9. Outlier / tail check ──────────────────────────────────────────────────
print("\n[9] Outlier / tail check — p99 vs max (winsorization candidates):")
heavy_tail_cols = ['f1','f4','f5','f6','f7','f8','f9','f10','f17','f18','f21']
for col in heavy_tail_cols:
    p99 = df[col].quantile(.99)
    p999 = df[col].quantile(.999)
    mx = df[col].max()
    cnt_above_p99 = (df[col] > p99).sum()
    if pd.notnull(p99):
        print(f"    {col}: p99={p99:,.2f}  p99.9={p999:,.2f}  max={mx:,.2f}  n>p99={cnt_above_p99}")

# ── 10. f4 and f21 co-missingness (rewards) ──────────────────────────────────
print("\n[10] f4 (Points Balance) vs f21 (Points Redeemed) missingness:")
miss_f4 = df['f4'].isnull()
miss_f21 = df['f21'].isnull()
both_miss = (miss_f4 & miss_f21).sum()
only_f4 = (miss_f4 & ~miss_f21).sum()
only_f21 = (~miss_f4 & miss_f21).sum()
neither = (~miss_f4 & ~miss_f21).sum()
print(f"    Both missing: {both_miss}  Only f4 missing: {only_f4}  Only f21 missing: {only_f21}  Both present: {neither}")

# ── 11. f22 vs f23 co-missingness ────────────────────────────────────────────
print("\n[11] f22 (Emails Opened) vs f23 (Emails Clicked) missingness:")
miss_f22 = df['f22'].isnull()
miss_f23 = df['f23'].isnull()
both_miss = (miss_f22 & miss_f23).sum()
only_f22 = (miss_f22 & ~miss_f23).sum()
only_f23 = (~miss_f22 & miss_f23).sum()
neither = (~miss_f22 & ~miss_f23).sum()
print(f"    Both missing: {both_miss}  Only f22 missing: {only_f22}  Only f23 missing: {only_f23}  Both present: {neither}")

# ── 12. f11 (Risk Score) distribution ────────────────────────────────────────
print("\n[12] f11 (Risk Score) distribution:")
print(df['f11'].describe(percentiles=[.01,.05,.25,.5,.75,.95,.99]))
print(f"    Range: [{df['f11'].min():.4f}, {df['f11'].max():.4f}]")
print(f"    Is scale 0-1? {df['f11'].max() <= 1.01}")
print(f"    Is scale 0-100? {df['f11'].max() > 1 and df['f11'].max() <= 100}")
print(f"    Is scale 0-999? {df['f11'].max() > 100}")

# ── Summary audit findings ────────────────────────────────────────────────────
print("\n" + "="*70)
print("AUDIT SUMMARY — key findings for Phase B/C design decisions")
print("="*70)
print(f"  Total rows: {len(df)} | Total cols: {len(df.columns)} (incl. id)")
print(f"  Columns with >10% missing: {(miss_pct > 10).sum()}")
heavy_miss = miss_pct[miss_pct > 10]
for f, pct in heavy_miss.items():
    print(f"    {f}: {pct:.1f}% missing")
print("  Done. Audit results above drive Phase B imputation strategy.")
