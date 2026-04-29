# Methodology

How the cap-table benchmark was built, scored, and audited.

## Why a new benchmark

[SpreadsheetBench](https://spreadsheetbench.io) is the canonical public spreadsheet-agent benchmark. It's 400 expert-annotated atomic operations — write a SUMIF, fix a lookup, debug a formula — averaging about 2 minutes per task. The leaderboard's headline comparisons currently anchor on Anthropic Opus 4.6 and OpenAI GPT-5.4, model generations that are now an iteration or two behind production.

We needed something different for the cap-table modeling work that defines venture and family-office due diligence:

- **End-to-end model construction**, not atomic operations
- **Live-formula propagation** as a measured property, not just output values
- **Current model generations** (Opus 4.7, GPT-5.5)
- **Domain-specific structure**: secondary tender, hybrid convertibles, multi-class anti-dilution, ESOP fixed-points

That's the scenario in `scenarios/kelvin_v4/`.

## Scenario design

The scenario models a synthetic Series F transaction. The deal includes:

- $1.4B pre-money valuation (a slight down-round vs the prior post-money — to trigger anti-dilution)
- $30M new primary investment
- $15M secondary tender at 0.85× the new round PPS
- Two outstanding convertible notes converting at the round closing — one capped, one discounted
- Pre-money ESOP refresh to 13% of post-money fully diluted, coupled into the round PPS as a fixed-point
- Anti-dilution adjustments on Series D and Series E (both prior classes are above the new PPS)
- A standard waterfall at five exit values from $700M to $8B with conversion election (preference vs convert) as a Nash equilibrium across seven priced classes

We track **167 specific output values**:

| Category | Keys |
|---|---|
| Pro-forma ownership percentages | 15 |
| Anti-dilution ratios + new conversion prices | 12 |
| Convertible note conversion (PPS, branch, shares) | 10 |
| Tender mechanics | 3 |
| Round solution (F PPS, post-F FD, etc.) | 4 |
| Waterfall dollars + elections | 95 |
| Sensitivity grid (3×3) | 18 |
| Return solver (3×, 4× per exit) | 10 |
| **Total** | **167** |

Full enumeration is in `scenarios/kelvin_v4/CANONICAL_INPUTS.md`.

## Dual-build ground truth

The single most important methodological choice: **we build truth two ways and reconcile them before accepting any value as canonical.**

1. **Python builder** (`truth/truth.py`): pure-decimal arithmetic at 60-digit precision, iterative fixed-point solver, no Excel involved. Produces `truth.json`.
2. **Excel builder** (`build_excel.py`): formula-driven openpyxl workbook with the same modeling logic expressed as cell formulas. Produces `truth/ground_truth.xlsx`.

We then run `harness.reconcile` to verify every one of the 167 tracked values agrees between the two builders within strict tolerance. PASS is required before any engine evaluation begins.

This catches a class of bugs that single-build truth never finds: arithmetic conventions where Python and Excel disagree silently (rounding, floating-point edge cases, formula evaluation order, named-range scope). v4 reconcile passed at 167/167 matched.

## Scoring

Each engine output is scored on three dimensions.

### Strict tier (the headline number)

For each of 167 tracked keys, the extracted value must be within tight tolerance of canonical truth:

| Field type | Strict tolerance |
|---|---|
| Ratios (`*.new_ratio`) | abs ≤ 1e-9 |
| Percentages (`*.pct`, `*.required_pct`) | abs ≤ 1e-6 |
| Dollar amounts | abs ≤ $1 OR rel ≤ 1e-4 (greater) |
| Share counts | abs ≤ 1e-2 (effectively integer) |
| Per-share prices | abs ≤ $1e-4 OR rel ≤ 1e-6 |
| Election strings ("preference" / "convert") | exact match |
| Convertible-branch strings ("cap" / "discount") | exact match |

Strict pass% = (keys passing strict) / 167.

### Audit-probe (live-formula propagation)

We perturb the input — change `RoundTerms.primary_raise_usd` from $30M to $33M — write the new value into the engine's output workbook (via named-range lookup, with label-scan fallback), recalculate the workbook headlessly via LibreOffice, re-extract the 167 values, and compare to `truth_perturbed.json` (which we built by re-running the Python truth builder with the perturbed input).

Audit-probe pass% = (keys updating correctly under perturbation) / 167.

The interpretation: a workbook full of live formulas that wire correctly to the input cell will update most cells when the input changes. A workbook with hardcoded constants (or formulas that wire to internal scratch values, not the boundary inputs) will not. The audit-probe tells you whether the engine's output is a *model* or a *report*.

Two engines emitted no named cells at all for the input boundary; their audit-probe scores are dominated by passthrough cells (cells whose value is the same in canonical and perturbed truth) and don't reflect actual recompute capability.

### VC-usable + directional tiers

For sensitivity analysis, we also report VC-usable (≤1pp / 1% rel tolerance) and directional (≤5pp / 5% rel tolerance) pass rates. These widen the strict tolerance so a small numerical drift doesn't fail the cell.

## Pre-publish audits

Before publishing, we ran five skeptical audits against our own scoring. All five reports are in the analysis folder.

| Audit | Question | Verdict |
|---|---|---|
| **D1** | Does the extractor treat all 4 engines uniformly? | ≤1pp residual bias. Headline ranking is engine quality, not measurement bias. |
| **D2** | Are the gaps statistically significant given σ ~10pp? | Skip-tier (Tetra > GPT-5.5) significant after Holm-Bonferroni (p=0.019, d=1.45). Adjacent-tier directional but underpowered at n=10. |
| **D3** | Where does the agent edge concentrate? | Tetra leads concentrated in waterfall + sensitivity + solver categories; ties on simpler primitives. |
| **D4** | Do high-strict engines have live formulas or value-dumps? | Tetra most "live" — 6/10 named-cell coverage, 65.6% audit-probe. GPT-5.5 elaborate-but-not-wired (2/10 named-cell coverage). Opus 4.7 not auditable on this metric (0/10). |
| **D5** | Are extractor disambiguation paths correct? | 8 disambiguation bugs identified and fixed. After fix, no engine-specific extractor bias remains. |

The audits caught **11 measurement bugs** that would have biased our results before they were corrected:

1. VC-tier scoring against the wrong (legacy v1) key set — fixed in score.py
2. Ownership column-picker selecting the wrong column type — fixed in extract.py
3. Series E / 30 synthesis path firing on v4 (a v3 invariant misapplied) — fixed
4. Common-row deduplication in waterfall extraction — fixed
5. GPT-5.5 round-shares extracting from a debug column — fixed
6. Opus 4.7 USD/shares confusion in round-block extraction — fixed
7. Our_fund label rule too broad — fixed
8. F_LEAD primary+secondary commingling — fixed
9. Employees label rule matching Series E investors (E1_INV vs E1) — fixed
10. `total_distributed` tolerance kind misrouting — fixed
11. PPS heuristic too narrow — fixed

After all fixes, the v4 reconcile remained 167/167 PASS, confirming no regression on the dual-build invariant.

## What we deliberately don't claim

- We do **not** claim our strict score is "the right level of difficulty." It's an end-to-end Series F. It's harder than SpreadsheetBench's atomic tasks. It's easier than a custom $5B PE deal. It's representative of the kind of work venture analysts do, which is the regime we built for.
- We do **not** make engine-quality claims outside of cap-table modeling. This is a cap-table benchmark. A model that's worse here may be excellent at, say, code generation.
- We do **not** claim no extractor bugs remain. We caught and fixed 11; we audited five different ways. There may be a twelfth.
- We do **not** assign blame for engine reliability where the harness is ours. We exclude reliability claims that we can't cleanly attribute to the engine versus our integration.

## Reproducibility

Everything in this repo runs locally. No cloud services required to verify the truth oracle, score an engine output, or run the audit-probe — just Python 3.11, openpyxl, scipy, and LibreOffice headless.

The engine output workbooks themselves are not in this repo (they're large, and we want vendors to have a chance to verify against our scoring before we make them public). They're available upon request.
