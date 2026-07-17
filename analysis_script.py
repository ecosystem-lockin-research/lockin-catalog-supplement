"""
Survey analysis script for the paper
"Digital Sovereignty under Pressure in Consumer Ecosystems"
(submitted to ICSOB 2026, double-blind review)

This script reproduces all numbers reported in Section 4.2 (User-Level
Validation), including the existence claim and composition claim analyses,
together with their 95% Wilson confidence intervals.

Inputs:
    Responses1.xlsx   (Google Forms export, sheet "Form Responses 2")

Outputs (written next to this script):
    sample_composition.csv          Ecosystem distribution 
    existence_claim_results.csv     Per-mechanism existence claim 
    composition_claim_results.csv   Per-dimension composition claim
    cleaning_diagnostics.csv        Counts of inconsistent responses

Cleaning conventions:
    1. Multi-select responses from disagreers (Likert 1, 2, or 3) are dropped
       (composition is conditional on perceiving the mechanism as a barrier)
    2. Multi-select responses from non-applicability respondents are dropped
       (switching costs cannot be attributed to a non-existent feature)
    3. Respondents who agreed but selected no cost dimension are treated as
       missing for the corresponding composition denominator

Run:
    pip install openpyxl statsmodels
    python analysis_script.py
"""

import csv
import os
from collections import Counter
from statistics import mean, stdev, median

import openpyxl
from statsmodels.stats.proportion import proportion_confint


# 1. Configuration: file location, mechanism mapping, response strings
# ----------------------------------------------------------------------------

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(HERE, "Responses1.xlsx")
SHEET_NAME = "Form Responses 2"

# Excluded ecosystem responses (those who exited at the second screen)
ECO_EXCLUDE = {"I do not consider myself part of any ecosystem"}

# Likert response strings as they appear in the Google Forms export,
# mapped to numeric values for descriptive statistics
LIKERT_NUMERIC = {
    "1 = Strongly disagree": 1,
    "2 = Disagree": 2,
    "3 = Neither agree nor disagree": 3,
    "4 = Agree": 4,
    "5 = Strongly agree": 5,
}
NOT_A_FEATURE = "This is not a feature of the ecosystem I use"
AGREE_RESPONSES = {"4 = Agree", "5 = Strongly agree"}
DISAGREE_OR_NEUTRAL = {
    "1 = Strongly disagree",
    "2 = Disagree",
    "3 = Neither agree nor disagree",
}

# Mechanism catalog with column indices (1-indexed) for the Likert and
# multi-select columns in the Google Forms export, plus the literature-
# attributed Burnham dimensions for each mechanism.
#
# The dimension texts are matched against multi-select cell content via
# substring search (Google Forms exports multi-select as comma-separated text).
MECHANISMS = [
    {
        "id": 1,
        "name": "Seamless coordination",
        "likert_col": 3,
        "multi_col": 4,
        "dimensions": {
            "Procedural": "The time and effort to set up similar connections between new products",
            "Financial": "Losing the benefits I now get from having my products work together",
            "Relational": "The loss of the feeling that my products belong together as one coherent set",
        },
    },
    {
        "id": 2,
        "name": "Cross-device continuity",
        "likert_col": 5,
        "multi_col": 6,
        "dimensions": {
            "Procedural": "The time and effort to relearn how a different ecosystem",
            "Financial": "The loss of the learning effort I have already invested in my current ecosystem",
            "Relational": "The loss of the familiar style of my current ecosystem that I have come to identify with",
        },
    },
    {
        "id": 3,
        "name": "Ecosystem-exclusive 1P",
        "likert_col": 7,
        "multi_col": 8,
        "dimensions": {
            "Procedural": "The time and effort to find and learn alternative apps or services on a different ecosystem",
            "Financial": "The loss of these specific apps and services that I currently use",
            "Relational": "The loss of the distinctive experience of these apps and services that I have come to identify with",
        },
    },
    {
        "id": 4,
        "name": "Interdependent 1P",
        "likert_col": 9,
        "multi_col": 10,
        "dimensions": {
            "Procedural": "The time and effort to find replacement products compatible with a different ecosystem",
            "Financial": "The loss of value in the dependent products that would no longer work properly",
        },
    },
    {
        "id": 5,
        "name": "Multi-product bundle",
        "likert_col": 11,
        "multi_col": 12,
        "dimensions": {
            "Procedural": "The time and effort to find and set up separate replacements for each bundled service",
            "Financial": "The loss of the bundled discount, since buying services separately would cost more",
            "Relational": "The coordination needed with family or household members who share my bundle",
        },
    },
    {
        "id": 6,
        "name": "Cross-promotional trials",
        "likert_col": 13,
        "multi_col": 14,
        "dimensions": {
            "Procedural": "The time and effort to find and set up replacements for the additional services I now use",
            "Relational": "The loss of the familiar experience of the additional services I tried and have come to identify with",
        },
    },
    {
        "id": 7,
        "name": "Limiting third-party access",
        "likert_col": 15,
        "multi_col": 16,
        "dimensions": {
            "Procedural": "The time and effort to find and learn third-party alternatives within a different ecosystem",
            "Financial": "The loss of credentials, settings, or routines I have built up in the ecosystem",
            "Relational": "The loss of the familiarity with the first-party services I have come to rely on",
        },
    },
    {
        "id": 8,
        "name": "Preferencing first-party",
        "likert_col": 17,
        "multi_col": 18,
        "dimensions": {
            "Procedural": "The time and effort to find and set up third-party alternatives for the brand",
            "Relational": "The loss of the familiar use of the brand",
        },
    },
    {
        "id": 9,
        "name": "Centralized ecosystem-exclusive account",
        "likert_col": 19,
        "multi_col": 20,
        "dimensions": {
            "Procedural": "The time and effort to set up a new account on a different ecosystem",
            "Financial": "The loss of the data, settings, and content I have accumulated against my current account",
            "Relational": "The loss of the digital identity I have built up around my account over time",
        },
    },
    {
        "id": 10,
        "name": "Non-portable content",
        "likert_col": 21,
        "multi_col": 22,
        "dimensions": {
            "Procedural": "The time and effort to rebuild an equivalent content library on a different ecosystem",
            "Financial": "The loss of the digital content I have purchased and accumulated over time",
        },
    },
    {
        "id": 11,
        "name": "Personified AI assistant",
        "likert_col": 23,
        "multi_col": 24,
        "dimensions": {
            "Procedural": "The time and effort to acclimate to a different assistant and recreate my personalized settings",
            "Relational": "The loss of the familiar way of interacting with the assistant that I have come to identify with",
        },
    },
]

# 2. Helpers
# ----------------------------------------------------------------------------

def has_option(cell, option_text):
    """Return True if the multi-select cell contains the given option."""
    if cell is None:
        return False
    return option_text in str(cell)


def has_any_tick(cell):
    """Return True if the multi-select cell is non-empty."""
    return cell is not None and str(cell).strip() != ""


def wilson_ci(n_success, n_total, alpha=0.05):
    """Return (lower, upper) Wilson score 95% CI as percentages."""
    lo, hi = proportion_confint(n_success, n_total, alpha=alpha, method="wilson")
    return lo * 100, hi * 100



# 3. Load and pre-filter data
# ----------------------------------------------------------------------------

print(f"Loading {DATA_FILE}")
wb = openpyxl.load_workbook(DATA_FILE, data_only=True)
ws = wb[SHEET_NAME]
rows = list(ws.iter_rows(min_row=2, values_only=True))
print(f"  Total response rows: {len(rows)}")

valid_rows = [r for r in rows if r[1] not in ECO_EXCLUDE]
print(f"  Valid respondents (in an ecosystem): {len(valid_rows)}")



# 4. Sample composition 
# ----------------------------------------------------------------------------

print("\n=== Sample composition ===")
eco_counts = Counter(r[1] for r in valid_rows)
total_valid = len(valid_rows)

with open(os.path.join(HERE, "sample_composition.csv"), "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["Ecosystem", "n", "percent"])
    for eco, n in eco_counts.most_common():
        pct = 100 * n / total_valid
        print(f"  {eco}: {n} ({pct:.1f}%)")
        writer.writerow([eco, n, f"{pct:.2f}"])
    writer.writerow(["Total", total_valid, "100.00"])



# 5. Existence claim 
# ----------------------------------------------------------------------------
# For each mechanism, compute:
#   n_app    = denominator (Likert 1-5, excludes "Not a feature")
#   n_NA     = count of "Not a feature" responses
#   M, SD    = mean and SD of Likert scores over n_app
#   Med      = median Likert score over n_app
#   n_agr    = count of Likert 4 or 5 (the numerator)
#   Rate     = n_agr / n_app
#   95% CI   = Wilson interval around Rate
# ----------------------------------------------------------------------------

print("\n=== Existence claim per mechanism ===")
existence_rows = []
print(
    f"{'M':>3}  {'Mechanism':<40} "
    f"{'n_app':>5} {'n_NA':>4} {'Mean':>4} {'SD':>4} {'Med':>3} "
    f"{'n_agr':>5} {'Rate':>6} {'95% Wilson CI':>17} {'Decision':>10}"
)
print("-" * 115)
for mech in MECHANISMS:
    likert_responses = [r[mech["likert_col"] - 1] for r in valid_rows]
    n_NA = sum(1 for v in likert_responses if v == NOT_A_FEATURE)
    numeric = [LIKERT_NUMERIC[v] for v in likert_responses if v in LIKERT_NUMERIC]
    n_app = len(numeric)
    n_agr = sum(1 for v in numeric if v >= 4)
    rate = n_agr / n_app
    ci_lo, ci_hi = wilson_ci(n_agr, n_app)
    m_val = mean(numeric)
    sd_val = stdev(numeric)
    med_val = median(numeric)
    if rate >= 0.5 and ci_lo > 50:
        decision = "Robust"
    elif rate >= 0.5:
        decision = "Marginal"
    else:
        decision = "Not validated"
    print(
        f"M{mech['id']:<2}  {mech['name']:<40} "
        f"{n_app:>5} {n_NA:>4} {m_val:>4.2f} {sd_val:>4.2f} {med_val:>3.0f} "
        f"{n_agr:>5} {rate*100:>5.1f}% [{ci_lo:>5.1f}, {ci_hi:>5.1f}] {decision:>10}"
    )
    existence_rows.append([
        mech["id"], mech["name"], n_app, n_NA,
        f"{m_val:.2f}", f"{sd_val:.2f}", med_val,
        n_agr, f"{rate*100:.1f}",
        f"{ci_lo:.1f}", f"{ci_hi:.1f}", decision,
    ])

with open(os.path.join(HERE, "existence_claim_results.csv"), "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow([
        "M", "Mechanism", "n_app", "n_NA",
        "Mean", "SD", "Median",
        "n_agreed", "Rate_percent",
        "Wilson_CI_lower_percent", "Wilson_CI_upper_percent", "Decision",
    ])
    writer.writerows(existence_rows)



# 6. Cleaning diagnostics 
# ----------------------------------------------------------------------------

print("\n=== Cleaning diagnostics (per Section 3.2.3 conventions) ===")
print(f"{'M':>3}  {'L1-3 + Burnham tick':>22}  {'NotFeature + Burnham':>22}  {'Agreed + no tick':>17}")
print("-" * 75)
diag_rows = []
for mech in MECHANISMS:
    lc, mc = mech["likert_col"], mech["multi_col"]
    lik13_tick = sum(
        1 for r in valid_rows
        if r[lc - 1] in DISAGREE_OR_NEUTRAL and has_any_tick(r[mc - 1])
    )
    notfeat_tick = sum(
        1 for r in valid_rows
        if r[lc - 1] == NOT_A_FEATURE and has_any_tick(r[mc - 1])
    )
    agreed_no_tick = sum(
        1 for r in valid_rows
        if r[lc - 1] in AGREE_RESPONSES and not has_any_tick(r[mc - 1])
    )
    print(f"M{mech['id']:<2}  {lik13_tick:>22}  {notfeat_tick:>22}  {agreed_no_tick:>17}")
    diag_rows.append([
        mech["id"], mech["name"],
        lik13_tick, notfeat_tick, agreed_no_tick,
    ])

with open(os.path.join(HERE, "cleaning_diagnostics.csv"), "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow([
        "M", "Mechanism",
        "L1to3_plus_Burnham_tick_dropped",
        "NotFeature_plus_Burnham_tick_dropped",
        "Agreed_plus_no_tick_treated_as_missing",
    ])
    writer.writerows(diag_rows)



# 7. Composition claim 
# ----------------------------------------------------------------------------
# For each mechanism EXCEPT M11 (whose existence claim fails and which is excluded), for each
# literature-attributed Burnham dimension:
#   n_fu     = follow-up denominator (Likert 4/5 AND at least one option ticked)
#   n_sel    = count of respondents in n_fu who selected this dimension
#   Rate     = n_sel / n_fu
#   95% CI   = Wilson interval around Rate
# ----------------------------------------------------------------------------

print("\n=== Composition claim per (mechanism, dimension) pair ===")
print(
    f"{'M':>3}  {'Dimension':<12} "
    f"{'n_fu':>4} {'n_sel':>5} {'Rate':>6} {'95% Wilson CI':>17} {'Decision':>15}"
)
print("-" * 80)

composition_rows = []
for mech in MECHANISMS:
    # Exclude M11: its existence claim was not validated, so composition is not assessed
    if mech["id"] == 11:
        continue

    # Build n_fu: agreed (Likert 4/5) AND at least one option ticked
    followup = [
        r for r in valid_rows
        if r[mech["likert_col"] - 1] in AGREE_RESPONSES
        and has_any_tick(r[mech["multi_col"] - 1])
    ]
    n_fu = len(followup)

    for dim_name, opt_text in mech["dimensions"].items():
        n_sel = sum(1 for r in followup if has_option(r[mech["multi_col"] - 1], opt_text))
        rate = n_sel / n_fu
        ci_lo, ci_hi = wilson_ci(n_sel, n_fu)
        if rate >= 0.5 and ci_lo > 50:
            decision = "Robust"
        elif rate >= 0.5:
            decision = "Marginal"
        else:
            decision = "Not corroborated"
        print(
            f"M{mech['id']:<2}  {dim_name:<12} "
            f"{n_fu:>4} {n_sel:>5} {rate*100:>5.1f}% "
            f"[{ci_lo:>5.1f}, {ci_hi:>5.1f}] {decision:>15}"
        )
        composition_rows.append([
            mech["id"], mech["name"], dim_name,
            n_fu, n_sel,
            f"{rate*100:.1f}",
            f"{ci_lo:.1f}", f"{ci_hi:.1f}", decision,
        ])

with open(os.path.join(HERE, "composition_claim_results.csv"), "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow([
        "M", "Mechanism", "Dimension",
        "n_followup", "n_selected", "Rate_percent",
        "Wilson_CI_lower_percent", "Wilson_CI_upper_percent", "Decision",
    ])
    writer.writerows(composition_rows)

# 8. Done
# ----------------------------------------------------------------------------

print("\nFour CSVs written to:", HERE)
print("  - sample_composition.csv")
print("  - existence_claim_results.csv")
print("  - cleaning_diagnostics.csv")
print("  - composition_claim_results.csv")
