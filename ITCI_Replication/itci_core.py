"""Python port of the Tax Foundation's International Tax Competitiveness
Index (ITCI) core scoring methodology
(https://github.com/TaxFoundation/international-tax-competitiveness-index,
R_code_files/08_index_calculations.R).

Validated to floating-point precision against the repo's own real 2025
ground-truth outputs (subcategories_2025.csv, categories_score.csv,
data_2025_run.csv) for 36 of 38 countries; the remaining 2 (Colombia,
Costa Rica) differ by <0.01 on a 0-100 scale in exactly one category
(income) -- isolated to a likely small upstream data revision between
snapshots, not a methodology difference (see scratch/validate.py).

Pipeline (matches 08_index_calculations.R exactly):
  1. Z-score every raw variable across countries: z = (x - mean(x)) / std(x)
  2. Flip sign of the 30 "higher = less competitive" variables
  3. Average z-scores (equal weight) within each of 15 subcategories -> RAW
     subcategory score
  4. score2() rescale each RAW subcategory score to 0-100 (independently)
  5. Average the RAW (pre-score2) subcategory scores (equal weight) within
     each of 5 categories -> RAW category score (NOT the rescaled ones --
     validated empirically, see scratch/validate.py)
  6. score2() rescale each RAW category score to 0-100
  7. Average the RAW (pre-score2) category scores (equal weight) -> RAW
     final score
  8. score2() rescale the RAW final score to 0-100
"""
import numpy as np
import pandas as pd

# ---------------------------------------------------------------- structure
SUBCATEGORIES = {
    "corporate_rate": ["corporate_rate"],
    "cost_recovery": ["loss_carryback", "loss_carryforward", "machines_cost_recovery",
                       "buildings_cost_recovery", "intangibles_cost_recovery", "inventory",
                       "allowance_corporate_equity"],
    "incentives": ["patent_box", "r_and_d_credit", "digital_services_tax",
                   "corporate_alt_minimum", "corporate_surtax", "corporate_other_rev"],
    "income_tax": ["top_income_rate", "threshold_top_income_rate", "tax_wedge"],
    "income_tax_complexity": ["personal_surtax", "personal_other_rev"],
    "capital_gains_and_dividends": ["capital_gains_rate", "dividends_rate"],
    "consumption_tax_rate": ["vat_rate"],
    "consumption_tax_base": ["vat_threshold", "vat_base"],
    "real_property_tax": ["property_tax", "property_tax_collections"],
    "wealth_taxes": ["net_wealth", "estate_or_inheritance_tax"],
    "capital_taxes": ["transfer_tax", "asset_tax", "capital_duties", "financial_transaction_tax"],
    "territorial": ["dividends_exemption", "capital_gains_exemption", "country_limitations"],
    "withholding_taxes": ["dividends_withholding_tax", "interest_withholding_tax",
                           "royalties_withholding_tax"],
    "tax_treaties": ["tax_treaties"],
    "international_regulations": ["cfc_rules", "thin_capitalization_rules", "minimum_tax"],
}

CATEGORIES = {
    "corporate": ["corporate_rate", "cost_recovery", "incentives"],
    "consumption": ["consumption_tax_rate", "consumption_tax_base"],
    "property": ["real_property_tax", "wealth_taxes", "capital_taxes"],
    "income": ["capital_gains_and_dividends", "income_tax", "income_tax_complexity"],
    "cross_border": ["territorial", "withholding_taxes", "tax_treaties",
                      "international_regulations"],
}

# The 30 variables where a HIGHER raw value means LESS competitive (rates,
# surtaxes, withholding taxes, restrictive rules, ...) -- sign-flipped after
# z-scoring so a higher z-score always means more competitive everywhere.
FLIP = [
    "corporate_rate", "patent_box", "r_and_d_credit", "digital_services_tax",
    "corporate_alt_minimum", "corporate_surtax", "corporate_other_rev",
    "top_income_rate", "threshold_top_income_rate", "tax_wedge", "personal_surtax",
    "personal_other_rev", "capital_gains_rate", "dividends_rate",
    "vat_rate", "vat_threshold",
    "property_tax_collections", "net_wealth", "estate_or_inheritance_tax", "transfer_tax",
    "asset_tax", "capital_duties", "financial_transaction_tax",
    "country_limitations", "dividends_withholding_tax", "interest_withholding_tax",
    "royalties_withholding_tax", "cfc_rules", "thin_capitalization_rules", "minimum_tax",
]

ALL_VARIABLES = [v for cols in SUBCATEGORIES.values() for v in cols]

# corporate_other_rev / personal_other_rev can be recorded as a negative
# revenue-impact figure for some countries (e.g. Sweden's personal_other_rev
# = -0.039785) -- 08_index_calculations.R converts these to absolute value
# before z-scoring. Missing this caused a real, isolated bug (Sweden's
# income_tax_complexity subcategory off by 0.33/100) caught during
# validation against the repo's own ground truth -- see scratch/validate.py.
ABS_VALUE_VARIABLES = ["corporate_other_rev", "personal_other_rev"]


# -------------------------------------------------------------------- steps
def compute_zscores(raw: pd.DataFrame) -> pd.DataFrame:
    """Step 1-2: z-score each variable (sample std, ddof=1), NaN -> 0, flip sign
    of the 30 "higher=worse" variables."""
    X = raw[ALL_VARIABLES].apply(pd.to_numeric, errors="coerce")
    for v in ABS_VALUE_VARIABLES:
        X[v] = X[v].abs()
    Z = (X - X.mean()) / X.std(ddof=1)
    Z = Z.fillna(0.0)
    for v in FLIP:
        Z[v] = -Z[v]
    return Z


def score2(x: pd.Series) -> pd.Series:
    """The ITCI's own rescale-to-100 transform: shift so the minimum becomes
    1 (eliminates negatives), then scale so the maximum becomes 100."""
    shifted = x + (-x.min() + 1)
    return shifted / shifted.max() * 100.0


def compute_subcategories(Z: pd.DataFrame) -> pd.DataFrame:
    """Step 3: equal-weighted average of variable z-scores within each
    subcategory (RAW, pre-score2)."""
    return pd.DataFrame({name: Z[cols].mean(axis=1) for name, cols in SUBCATEGORIES.items()})


def compute_categories(subcategories_raw: pd.DataFrame) -> pd.DataFrame:
    """Step 5: equal-weighted average of RAW (pre-score2) subcategory scores
    within each category. Uses the RAW values, not the score2-rescaled
    ones -- validated empirically against categories_score.csv."""
    return pd.DataFrame({name: subcategories_raw[cols].mean(axis=1) for name, cols in CATEGORIES.items()})


def compute_final(categories_raw: pd.DataFrame) -> pd.Series:
    """Step 7: equal-weighted average of RAW (pre-score2) category scores."""
    return categories_raw.mean(axis=1)


def run_itci(raw: pd.DataFrame) -> dict:
    """Full pipeline. `raw` must have columns = ALL_VARIABLES (+ any ID
    columns, ignored). Returns a dict of intermediate and final DataFrames/
    Series, all indexed to match `raw`'s row order."""
    Z = compute_zscores(raw)
    sub_raw = compute_subcategories(Z)
    sub_scored = sub_raw.apply(score2, axis=0)
    cat_raw = compute_categories(sub_raw)
    cat_scored = cat_raw.apply(score2, axis=0)
    final_raw = compute_final(cat_raw)
    final_scored = score2(final_raw)
    return {
        "zscores": Z,
        "subcategories_raw": sub_raw,
        "subcategories_scored": sub_scored,
        "categories_raw": cat_raw,
        "categories_scored": cat_scored,
        "final_raw": final_raw,
        "final_scored": final_scored,
    }
