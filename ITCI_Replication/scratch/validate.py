"""Full validation of itci_core.py against the real 2025 ground-truth files
downloaded from the Tax Foundation's own repo."""
import sys
sys.path.insert(0, "..")
import numpy as np
import pandas as pd
from itci_core import run_itci, SUBCATEGORIES, CATEGORIES

raw = pd.read_csv("../raw_repo_data/final_index_data_2025.csv")
result = run_itci(raw)
result["subcategories_scored"].index = raw["country"]
result["categories_scored"].index = raw["country"]
final = result["final_scored"]
final.index = raw["country"]

# ---- subcategory-level check ---------------------------------------------
truth_sub = pd.read_csv("../raw_repo_data/subcategories_2025.csv").set_index("country")
max_diff_sub = 0.0
for name in SUBCATEGORIES:
    mine = result["subcategories_scored"][name]
    truth = truth_sub[name]
    d = (mine - truth.reindex(mine.index)).abs().max()
    max_diff_sub = max(max_diff_sub, d)
print(f"Subcategories: max abs diff across all 15 = {max_diff_sub:.2e}")

# ---- category-level check -------------------------------------------------
truth_cat = pd.read_csv("../raw_repo_data/data_2025_run.csv").set_index("country")
diffs_cat = {}
for name in CATEGORIES:
    mine = result["categories_scored"][name]
    truth = truth_cat[name]
    d = (mine - truth.reindex(mine.index)).abs()
    diffs_cat[name] = d
diff_df = pd.DataFrame(diffs_cat)
print(f"\nCategories: max abs diff across all 5 x 38 countries = {diff_df.to_numpy().max():.4f}")
worst = diff_df.max(axis=1).sort_values(ascending=False).head(5)
print("Worst 5 countries (max diff across any category):")
print(worst)

# ---- final score check ------------------------------------------------
truth_final = truth_cat["final"]
diff_final = (final - truth_final.reindex(final.index)).abs()
print(f"\nFinal score: max abs diff = {diff_final.max():.4f}")
print("Worst 5 countries:")
print(diff_final.sort_values(ascending=False).head(5))
n_exact = (diff_final < 1e-6).sum()
print(f"\n{n_exact} of {len(diff_final)} countries match the final score to floating-point precision (<1e-6)")
