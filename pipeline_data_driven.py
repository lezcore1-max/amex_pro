import pandas as pd
import numpy as np
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
import os

print("AMEX CAMPUS CHALLENGE - DATA DRIVEN PIPELINE")
print("Loading data...")
df = pd.read_csv('campus_challenge_r1_data.csv')

# 1. Feature Selection based on data separation/bimodality
features = ['f1', 'f4', 'f5', 'f11', 'f17', 'f21']
print(f"Selected features: {features}")

# 2. Imputation (Data-driven median imputation to avoid assumptions about 'not enrolled')
df_imp = df.copy()
for f in features:
    df_imp[f] = df_imp[f].fillna(df_imp[f].median())

# 3. Standardization (Z-scores to remove arbitrary dollar weightings)
scaler = StandardScaler()
X_scaled = scaler.fit_transform(df_imp[features])

# 4. PCA for clean separation
print("Running PCA to find the axis of maximum variance...")
pca = PCA(n_components=1)
pc1 = pca.fit_transform(X_scaled).flatten()

# Align PC1 direction so higher spend/revolve means higher score
if np.corrcoef(pc1, df_imp['f5'])[0, 1] < 0:
    pc1 = -pc1

# 5. Calculate Score
print("Calculating percentiles...")
df['Prediction'] = pd.Series(pc1).rank(pct=True)

# 6. Save Submission
print("Saving submission...")
sub = df[['id', 'Prediction']].copy()
sub.rename(columns={'id': 'ID'}, inplace=True)
sub.to_excel('amex_submission_round1_v2.xlsx', index=False)
print("Saved amex_submission_round1_v2.xlsx successfully.")
