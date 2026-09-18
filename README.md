# International Tax Competitiveness Index 

A from-scratch Python (and [Excel](ITCI_Calculator.xlsx)) replica of the Tax
Foundation's [International Tax Competitiveness Index](https://github.com/TaxFoundation/international-tax-competitiveness-index)
(ITCI) — the R script `R_code_files/08_index_calculations.R`, ported and
validated to **floating-point precision against the repo's own real 2025
ground-truth output** for all 38 countries.

**Structure**: 42 raw tax-policy variables → 15 subcategories → 5 categories
→ 1 final score, each level an equal-weighted average with a rescale-to-100
step. No factor analysis, no eigen-decomposition — everything here is
closed-form arithmetic (z-scores, averages, min/max rescaling), which is
also why the companion Excel calculator needs no iterative engine anywhere,
unlike the WGI/UCM/Prince projects.

**Worth knowing before you read the code**: the R script's own variable-name
reuse makes two real methodology choices genuinely ambiguous from reading
the code alone. Both were resolved here empirically — by testing hypotheses
directly against the repo's real intermediate output files — not by
guessing. Section 6 covers this in detail.
