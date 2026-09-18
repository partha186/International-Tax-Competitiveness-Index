# International Tax Competitiveness Index 

A from-scratch Python (and [Excel](ITCI_Calculator.xlsx)) replica of the Tax
Foundation's [International Tax Competitiveness Index](https://github.com/TaxFoundation/international-tax-competitiveness-index)
(ITCI) — the R script `R_code_files/08_index_calculations.R`, ported and
validated to **floating-point precision against the repo's own real 2025
ground-truth output** for all 38 countries.

**Structure**: 42 raw tax-policy variables → 15 subcategories → 5 categories
→ 1 final score, each level an equal-weighted average with a rescale-to-100
step. No factor analysis, no eigen-decomposition — everything here is
closed-form arithmetic (z-scores, averages, min/max rescaling), 

Structure: 42 raw tax-policy variables (corporate rate, VAT, withholding taxes, etc.) roll up into 15 subcategories → 5 categories (corporate, consumption, property, income, cross-border) → 1 final competitiveness score, for 38 OECD countries.
Z-score everything so rates/thresholds/binary rules are all comparable, then sign-flip 30 of the 42 variables where a higher raw value means less competitive (e.g. tax rates), so "higher z-score = more competitive" holds universally.
score2 rescale: shift/scale each level's average to a clean 0–100 range (worst country = ~1, best = 100). Applied independently at the subcategory, category, and final-score levels.
Two genuine ambiguities in the original R code, resolved empirically (not by guesswork):
corporate_other_rev/personal_other_rev need abs() before z-scoring (Sweden has a negative value there) — missing it silently shifts Sweden's score by 0.33 points.
Categories are averaged from the raw (pre-rescale) subcategory z-scores, not the rescaled 0–100 ones — the R script reuses the variable name subcategories for both, making this genuinely undecidable from a code read alone. Confirmed by matching the Tax Foundation's own published categories_score.csv to 1e-15 precision.
Validation: all 38 countries match the Tax Foundation's real 2025 output to floating-point precision (~1e-14) at every stage — subcategories, categories, and the final score.
Result: Estonia ranks #1 (100.0), followed by Latvia, New Zealand, Switzerland; property-tax scores correlate most strongly with the final ranking, matching the Tax Foundation's own reported finding.
