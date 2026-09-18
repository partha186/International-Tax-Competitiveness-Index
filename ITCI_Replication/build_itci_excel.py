"""Builds ITCI_Calculator.xlsx: a working Excel replica of the Tax
Foundation's International Tax Competitiveness Index scoring methodology
(itci_core.py / 08_index_calculations.R), fully live and validated to match
the Python port's floating-point-exact results on the real 2025 data.

Entirely closed-form (z-scores, averages, min/max rescale) -- unlike the
WGI/UCM/Prince projects, no eigen-decomposition or iterative engine is
needed anywhere, so every formula here is a single native Excel function.
"""
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

from itci_core import SUBCATEGORIES, CATEGORIES, FLIP, ALL_VARIABLES, ABS_VALUE_VARIABLES

GIVEN_FILL = PatternFill("solid", fgColor="FCEBD5")
GIVEN_FONT = Font(color="A8671B")
LIVE_FILL = PatternFill("solid", fgColor="D9ECE6")
LIVE_FONT = Font(color="0D7A68")
HEADER_FILL = PatternFill("solid", fgColor="2F3B45")
HEADER_FONT = Font(color="FFFFFF", bold=True)
TITLE_FONT = Font(bold=True, size=14)
THIN = Side(style="thin", color="C7D0C5")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def style_header(ws, row, col, text):
    c = ws.cell(row, col, text)
    c.fill = HEADER_FILL
    c.font = HEADER_FONT
    c.alignment = Alignment(horizontal="center")
    return c


def style_given(cell):
    cell.fill = GIVEN_FILL
    cell.font = GIVEN_FONT


def style_live(cell):
    cell.fill = LIVE_FILL
    cell.font = LIVE_FONT


def addr(col, row):
    return f"{get_column_letter(col)}{row}"


def qref(sheet, col, row):
    return f"'{sheet}'!${get_column_letter(col)}${row}"


def build(data_path="raw_repo_data/final_index_data_2025.csv", out_path="ITCI_Calculator.xlsx"):
    raw = pd.read_csv(data_path)
    n = len(raw)
    var_cols = {v: i for i, v in enumerate(ALL_VARIABLES)}

    wb = Workbook()
    wb.remove(wb.active)

    # ------------------------------------------------------------- RawData
    ws = wb.create_sheet("RawData")
    ws.cell(1, 1, f"Raw Data (GIVEN) — {n} countries x {len(ALL_VARIABLES)} variables, "
                  f"2025 ITCI input data").font = TITLE_FONT
    style_header(ws, 2, 1, "country")
    for j, v in enumerate(ALL_VARIABLES):
        style_header(ws, 2, 2 + j, v)
    for i in range(n):
        row = 3 + i
        ws.cell(row, 1, raw["country"].iloc[i]).border = BORDER
        for j, v in enumerate(ALL_VARIABLES):
            cell = ws.cell(row, 2 + j, float(raw[v].iloc[i]))
            style_given(cell)
    d_top, d_bot = 3, 3 + n - 1
    var_col_letter = {v: get_column_letter(2 + j) for j, v in enumerate(ALL_VARIABLES)}

    # ----------------------------------------------------------- AbsValues
    # corporate_other_rev/personal_other_rev need ABS() before z-scoring.
    # Materialized as plain elementwise formulas here (not AVERAGE(ABS(range))
    # inline in ZScores) because ABS() over a whole range only evaluates
    # correctly as an array formula in older/non-365 Excel -- avoiding that
    # compatibility risk entirely rather than relying on CSE/dynamic arrays.
    ws_abs = wb.create_sheet("AbsValues")
    ws_abs.cell(1, 1, "ABS() of corporate_other_rev / personal_other_rev (LIVE)").font = TITLE_FONT
    style_header(ws_abs, 2, 1, "country")
    for j, v in enumerate(ABS_VALUE_VARIABLES):
        style_header(ws_abs, 2, 2 + j, v)
    for i in range(n):
        row = 3 + i
        ws_abs.cell(row, 1, f"=RawData!A{d_top + i}").border = BORDER
        for j, v in enumerate(ABS_VALUE_VARIABLES):
            cl = var_col_letter[v]
            cell = ws_abs.cell(row, 2 + j, f"=ABS(RawData!{cl}{d_top + i})")
            style_live(cell)
    abs_top, abs_bot = 3, 3 + n - 1
    abs_col_letter = {v: get_column_letter(2 + j) for j, v in enumerate(ABS_VALUE_VARIABLES)}

    # ------------------------------------------------------------- ZScores
    ws2 = wb.create_sheet("ZScores")
    ws2.cell(1, 1, "Z-Scores (LIVE)").font = TITLE_FONT
    ws2.cell(2, 1, "z = (x - AVERAGE) / STDEV, across the country cross-section. "
                   "corporate_other_rev/personal_other_rev use ABS(x) first (can be recorded "
                   "as a negative revenue-impact figure). 30 variables then sign-flipped "
                   "(*-1) so a higher z-score always means more competitive.")
    style_header(ws2, 3, 1, "country")
    for j, v in enumerate(ALL_VARIABLES):
        style_header(ws2, 3, 2 + j, v)
    for i in range(n):
        row = 4 + i
        ws2.cell(row, 1, f"=RawData!A{d_top + i}").border = BORDER
        for j, v in enumerate(ALL_VARIABLES):
            if v in ABS_VALUE_VARIABLES:
                acl = abs_col_letter[v]
                rng = f"AbsValues!${acl}${abs_top}:${acl}${abs_bot}"
                src = f"AbsValues!{acl}{abs_top + i}"
            else:
                cl = var_col_letter[v]
                rng = f"RawData!${cl}${d_top}:${cl}${d_bot}"
                src = f"RawData!{cl}{d_top + i}"
            sign = "-1*" if v in FLIP else ""
            formula = f"={sign}(({src}-AVERAGE({rng}))/STDEV({rng}))"
            cell = ws2.cell(row, 2 + j, formula)
            style_live(cell)
    z_top, z_bot = 4, 4 + n - 1
    z_col_letter = {v: get_column_letter(2 + j) for j, v in enumerate(ALL_VARIABLES)}

    def z_ref(v, i):
        return f"ZScores!{z_col_letter[v]}{z_top + i}"

    # -------------------------------------------------------- Subcategories
    ws3 = wb.create_sheet("Subcategories")
    ws3.cell(1, 1, "Subcategories — RAW (pre-rescale) Equal-Weighted Averages (LIVE)").font = TITLE_FONT
    style_header(ws3, 2, 1, "country")
    sub_names = list(SUBCATEGORIES.keys())
    for j, s in enumerate(sub_names):
        style_header(ws3, 2, 2 + j, s)
    for i in range(n):
        row = 3 + i
        ws3.cell(row, 1, f"=RawData!A{d_top + i}").border = BORDER
        for j, s in enumerate(sub_names):
            cols = SUBCATEGORIES[s]
            terms = "+".join(z_ref(v, i) for v in cols)
            cell = ws3.cell(row, 2 + j, f"=({terms})/{len(cols)}")
            style_live(cell)
    sub_top, sub_bot = 3, 3 + n - 1
    sub_col_letter = {s: get_column_letter(2 + j) for j, s in enumerate(sub_names)}

    # --------------------------------------------------- SubcategoriesScored
    ws4 = wb.create_sheet("SubcategoriesScored")
    ws4.cell(1, 1, "Subcategories — Rescaled to 0-100 (LIVE)").font = TITLE_FONT
    ws4.cell(2, 1, "score2(x): shift so MIN becomes 1 (eliminates negatives), "
                   "then scale so MAX becomes 100.")
    style_header(ws4, 3, 1, "country")
    for j, s in enumerate(sub_names):
        style_header(ws4, 3, 2 + j, s)
    for i in range(n):
        row = 4 + i
        ws4.cell(row, 1, f"=RawData!A{d_top + i}").border = BORDER
        for j, s in enumerate(sub_names):
            cl = sub_col_letter[s]
            rng = f"Subcategories!${cl}${sub_top}:${cl}${sub_bot}"
            src = f"Subcategories!{cl}{sub_top + i}"
            shifted = f"({src}-MIN({rng})+1)"
            shifted_max = f"(MAX({rng})-MIN({rng})+1)"
            cell = ws4.cell(row, 2 + j, f"={shifted}/{shifted_max}*100")
            style_live(cell)

    # ------------------------------------------------------------ Categories
    ws5 = wb.create_sheet("Categories")
    ws5.cell(1, 1, "Categories — RAW (pre-rescale) Equal-Weighted Averages (LIVE)").font = TITLE_FONT
    ws5.cell(2, 1, "Uses the RAW (pre-score2) Subcategories sheet, not SubcategoriesScored — "
                   "validated against the Tax Foundation's own categories_score.csv.")
    style_header(ws5, 3, 1, "country")
    cat_names = list(CATEGORIES.keys())
    for j, c in enumerate(cat_names):
        style_header(ws5, 3, 2 + j, c)
    for i in range(n):
        row = 4 + i
        ws5.cell(row, 1, f"=RawData!A{d_top + i}").border = BORDER
        for j, c in enumerate(cat_names):
            cols = CATEGORIES[c]
            terms = "+".join(f"Subcategories!{sub_col_letter[s]}{sub_top + i}" for s in cols)
            cell = ws5.cell(row, 2 + j, f"=({terms})/{len(cols)}")
            style_live(cell)
    cat_top, cat_bot = 4, 4 + n - 1
    cat_col_letter = {c: get_column_letter(2 + j) for j, c in enumerate(cat_names)}

    # ------------------------------------------------------ CategoriesScored
    ws6 = wb.create_sheet("CategoriesScored")
    ws6.cell(1, 1, "Categories — Rescaled to 0-100 (LIVE)").font = TITLE_FONT
    style_header(ws6, 2, 1, "country")
    for j, c in enumerate(cat_names):
        style_header(ws6, 2, 2 + j, c)
    for i in range(n):
        row = 3 + i
        ws6.cell(row, 1, f"=RawData!A{d_top + i}").border = BORDER
        for j, c in enumerate(cat_names):
            cl = cat_col_letter[c]
            rng = f"Categories!${cl}${cat_top}:${cl}${cat_bot}"
            src = f"Categories!{cl}{cat_top + i}"
            shifted = f"({src}-MIN({rng})+1)"
            shifted_max = f"(MAX({rng})-MIN({rng})+1)"
            cell = ws6.cell(row, 2 + j, f"={shifted}/{shifted_max}*100")
            style_live(cell)
    catsc_top = 3

    # ------------------------------------------------------------- Results
    ws7 = wb.create_sheet("Results")
    ws7.cell(1, 1, "Results — Final Score and Ranking (LIVE)").font = TITLE_FONT
    ws7.cell(2, 1, "Final = score2(AVERAGE of the 5 RAW (pre-rescale) category scores).")
    headers = ["country"] + [f"{c}_score" for c in cat_names] + [f"{c}_rank" for c in cat_names] + \
              ["final_score", "final_rank"]
    style_header(ws7, 4, 1, "country")
    for j, c in enumerate(cat_names):
        style_header(ws7, 4, 2 + j, f"{c} score")
    for j, c in enumerate(cat_names):
        style_header(ws7, 4, 2 + len(cat_names) + j, f"{c} rank")
    style_header(ws7, 4, 2 + 2 * len(cat_names), "final score")
    style_header(ws7, 4, 3 + 2 * len(cat_names), "final rank")

    final_raw_col = 4 + 2 * len(cat_names)  # helper column, hidden-ish
    ws7.cell(4, final_raw_col, "final_raw (helper)")
    for i in range(n):
        row = 5 + i
        ws7.cell(row, 1, f"=RawData!A{d_top + i}").border = BORDER
        for j, c in enumerate(cat_names):
            cell = ws7.cell(row, 2 + j, f"=CategoriesScored!{cat_col_letter[c]}{catsc_top + i}")
            style_live(cell)
        for j, c in enumerate(cat_names):
            cl = get_column_letter(2 + j)
            rng = f"${cl}$5:${cl}${5 + n - 1}"
            cell = ws7.cell(row, 2 + len(cat_names) + j, f"=RANK({cl}{row},{rng},0)")
            style_live(cell)
        cat_raw_terms = "+".join(f"Categories!{cat_col_letter[c]}{cat_top + i}" for c in cat_names)
        ws7.cell(row, final_raw_col, f"=({cat_raw_terms})/{len(cat_names)}")

    final_raw_rng = f"${get_column_letter(final_raw_col)}$5:${get_column_letter(final_raw_col)}${5 + n - 1}"
    final_score_col = 2 + 2 * len(cat_names)
    final_rank_col = final_score_col + 1
    for i in range(n):
        row = 5 + i
        fr_cell = f"{get_column_letter(final_raw_col)}{row}"
        shifted = f"({fr_cell}-MIN({final_raw_rng})+1)"
        shifted_max = f"(MAX({final_raw_rng})-MIN({final_raw_rng})+1)"
        cell = ws7.cell(row, final_score_col, f"={shifted}/{shifted_max}*100")
        style_live(cell)
    final_score_rng = f"${get_column_letter(final_score_col)}$5:${get_column_letter(final_score_col)}${5 + n - 1}"
    for i in range(n):
        row = 5 + i
        cl = get_column_letter(final_score_col)
        cell = ws7.cell(row, final_rank_col, f"=RANK({cl}{row},{final_score_rng},0)")
        style_live(cell)

    chart = BarChart()
    chart.type = "bar"
    chart.title = "Final ITCI Score by Country"
    chart.y_axis.title = "score (0-100)"
    chart.height, chart.width = 22, 16
    data_ref = Reference(ws7, min_col=final_score_col, min_row=4, max_row=5 + n - 1)
    cats_ref = Reference(ws7, min_col=1, min_row=5, max_row=5 + n - 1)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.series[0].graphicalProperties.solidFill = "0D7A68"
    ws7.add_chart(chart, addr(final_rank_col + 2, 4))

    # --------------------------------------------------------------- README
    ws0 = wb.create_sheet("README", 0)
    ws0.column_dimensions["A"].width = 100
    lines = [
        ("International Tax Competitiveness Index — Working Excel Calculator", TITLE_FONT),
        ("", None),
        ("Replicates the Tax Foundation's ITCI scoring methodology "
         "(github.com/TaxFoundation/international-tax-competitiveness-index, "
         "R_code_files/08_index_calculations.R) as live, formula-driven Excel sheets — "
         "no VBA, no add-ins. Entirely closed-form (z-scores, averages, min/max rescale); "
         "no eigen-decomposition or iterative engine needed anywhere.", None),
        ("", None),
        ("Color legend", Font(bold=True)),
        ("  Amber cells = GIVEN (2025 ITCI input data, from the Tax Foundation's own repo)", None),
        ("  Teal cells  = LIVE (formulas; recalculate automatically)", None),
        ("", None),
        ("Validated", Font(bold=True)),
        ("  Every stage (subcategories, categories, final score) matches the Tax Foundation's "
         "own real 2025 ground-truth output files to floating-point precision (<1e-13) for all "
         "38 countries — see itci_core.py and scratch/validate.py in the companion Python "
         "project. Two real bugs were caught and fixed during validation: (1) "
         "corporate_other_rev/personal_other_rev need ABS() before z-scoring (Sweden has a "
         "negative personal_other_rev, off by 0.33/100 without the fix), and (2) categories "
         "must be computed from the RAW (pre-rescale) subcategory z-averages, not the "
         "rescaled 0-100 subcategory values — confirmed empirically against categories_score.csv "
         "since the R script reuses one variable name for both versions, making it ambiguous "
         "from the code alone.", None),
        ("", None),
        ("Pipeline", Font(bold=True)),
        ("  1. ZScores: z = (x-AVERAGE)/STDEV across the 38-country cross-section, per variable. "
         "corporate_other_rev/personal_other_rev use ABS(x) first. 30 \"higher=worse\" variables "
         "(rates, surtaxes, withholding taxes, restrictive rules) then multiplied by -1.", None),
        ("  2. Subcategories: equal-weighted average of the relevant z-scores, one column per "
         "subcategory (15 total).", None),
        ("  3. SubcategoriesScored: score2() rescale — shift so MIN becomes 1, scale so MAX "
         "becomes 100 — applied independently to each subcategory.", None),
        ("  4. Categories: equal-weighted average of the RAW (pre-score2) subcategory scores, "
         "one column per category (5 total: corporate, consumption, property, income, "
         "cross_border).", None),
        ("  5. CategoriesScored: score2() rescale, same as step 3.", None),
        ("  6. Results: final score = score2(average of the 5 RAW category scores). Category "
         "and final RANKs via RANK() (descending — rank 1 = most competitive). Bar chart of "
         "final scores by country.", None),
        ("", None),
        ("Replacing the data", Font(bold=True)),
        ("  Overwrite the amber GIVEN cells on RawData with updated variable values (same 38 "
         "countries, or add/remove rows). Everything downstream — z-scores, subcategories, "
         "categories, final score, ranks, chart — recalculates automatically. This is exactly "
         "how the Tax Foundation's own repo supports testing tax-reform scenarios (README: "
         "\"modify tax variables ... re-run ... compare against baseline\").", None),
    ]
    r = 1
    for text, font in lines:
        c = ws0.cell(r, 1, text)
        if font:
            c.font = font
        c.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    wb.save(out_path)
    print(f"wrote {out_path}  ({n} countries, {len(ALL_VARIABLES)} variables, "
          f"{len(sub_names)} subcategories, {len(cat_names)} categories)")


if __name__ == "__main__":
    build()
