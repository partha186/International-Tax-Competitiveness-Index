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
