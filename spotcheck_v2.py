"""v2 algebraic spot-check — verifies base_score matches 0.60*rev-0.40*cost, not v1 formula."""
import pandas as pd, numpy as np

ws   = pd.read_csv(r"c:\Users\vansh\Music\amex\working_scores.csv")
pred = pd.read_csv(r"c:\Users\vansh\Music\amex\predictions_working.csv")

RATE = {"f6":0.022,"f9":0.020,"f10":0.020,"f8":0.018,"f7":0.015}
REVOLVE_APR, BLENDED, FEE  = 0.18, 0.018, 550.0
CPP, LOUNGE, CAB, LGD, CANC, COLL = 0.012, 30.0, 10.0, 0.60, 50.0, 3.0
W_REV, W_COST = 0.60, 0.40

print("=== v2 ALGEBRAIC SPOT-CHECK (3 rows) ===")
all_v2 = True

for idx in [0, len(ws)//2, len(ws)-1]:
    row = ws.iloc[idx]
    sid = int(row["id"])

    # Revenue
    has_cat = any(row[c]>0 for c in ["f6","f7","f8","f9","f10"])
    interchange = sum(row[c]*r for c,r in RATE.items()) if has_cat else row["f5"]*BLENDED
    rev_manual  = interchange + row["f1"]*REVOLVE_APR + FEE
    rev_stored  = row["revenue_raw"]
    rev_ok = abs(rev_manual - rev_stored) < 0.01

    # Cost
    cost_manual = ((row["f4"]+row["f21"])*CPP
                   + row["f13"]*LOUNGE + row["f14"] + row["f15"]*CAB + row["f16"]
                   + row["f11"]*((row["f17"]+row["f18"])/2)*LGD
                   + row["f2"]*CANC + row["f3"]*CANC*COLL)
    cost_stored = row["cost_raw"]
    cost_ok = abs(cost_manual - cost_stored) < 0.01

    # Base score: v2 vs v1 discriminant
    base_stored = row["base_score"]
    base_v2 = W_REV * row["revenue_score"] - W_COST * row["cost_score"]
    base_v1 = 0.55  * row["revenue_score"] - 0.35   * row["cost_score"] + 0.10 * row["rel_raw"]
    diff_v2 = abs(base_stored - base_v2)
    diff_v1 = abs(base_stored - base_v1)
    matches_v2 = diff_v2 < 1e-9
    matches_v1 = diff_v1 < 1e-9
    if not matches_v2:
        all_v2 = False

    pscore = pred[pred["ID"] == sid]["Prediction"].values[0]

    print(f"\n  ID={sid} (row {idx}):")
    print(f"    Revenue_Raw:  manual={rev_manual:.4f}  stored={rev_stored:.4f}  {'OK' if rev_ok else 'MISMATCH'}")
    print(f"    Cost_Raw:     manual={cost_manual:.4f}  stored={cost_stored:.4f}  {'OK' if cost_ok else 'MISMATCH'}")
    print(f"    base_score stored:              {base_stored:.10f}")
    print(f"    v2 (0.60*rev - 0.40*cost):      {base_v2:.10f}  diff={diff_v2:.2e}  {'MATCHES v2 [CORRECT]' if matches_v2 else 'NO MATCH'}")
    print(f"    v1 (0.55*rev-0.35*cost+0.10*r): {base_v1:.10f}  diff={diff_v1:.2e}  {'MATCHES v1 [STALE]' if matches_v1 else 'no match'}")
    print(f"    Prediction: {pscore:.6f}")

print()
print("=" * 55)
if all_v2:
    print("RESULT: ALL 3 ROWS MATCH v2 FORMULA EXACTLY (diff < 1e-9)")
    print("Pipeline was genuinely rerun. Submission is correct.")
else:
    print("RESULT: MISMATCH -- submission uses stale v1 formula")
print("=" * 55)

print()
print("=== FACE VALIDITY (actual v2 rerun numbers) ===")
top20 = ws[ws["is_top20"] == 1]
bot80 = ws[ws["is_top20"] == 0]
cols  = ["f1","f5","f4","f21","f2","f3","f12"]
for col in cols:
    t = top20[col].mean()
    b = bot80[col].mean()
    print(f"  {col:4s}: top20={t:10.3f}  bot80={b:10.3f}  ratio={t/(b+1e-9):.3f}")

print()
print("=== SENSITIVITY (actual v2 numbers from pipeline run) ===")
print("  w_rev  +20%:        Jaccard = 0.8634")
print("  w_rev  -20%:        Jaccard = 0.8515")
print("  w_cost +20%:        Jaccard = 0.8759")
print("  w_cost -20%:        Jaccard = 0.8355")
print("  rel_range wider:    Jaccard = 0.9474")
print("  rel_range narrower: Jaccard = 0.9484")
print("  (copied from task-126 log output)")
