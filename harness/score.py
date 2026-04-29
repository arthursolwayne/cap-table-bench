"""
Scoring engine for the tetra-v-shortcut benchmark.

Grades a single API run's xlsx output against the ground-truth JSON across
five rubric dimensions: correctness (tiered), auditability, structural
fidelity, timing, and (v4+) audit-probe. Robustness is computed upstream
(across variants) and is NOT handled here.

Public API:
    - score_run(output_xlsx_path, truth_json_path, run_result_path,
                output_report_path=None) -> ScoreResult
    - score_audit_probe(canonical_output_xlsx, perturbations,
                        truth_perturbed_path, tracked_keys) -> float | None
    - ScoreResult (dataclass)
    - TierResult (dataclass)

Run as a CLI:
    python score.py --output out.xlsx --truth truth.json --run result.json \
        [--report score.md]

Audit-probe (v4+)
-----------------
The audit-probe is a sub-test that runs INSIDE this scoring engine. It
opens the engine's output workbook, overwrites a small set of named input
cells with perturbed values (e.g. v4: `RoundTerms_primary_raise_usd` $30M
→ $33M), forces a LibreOffice headless recalc, re-extracts all 167 tracked
truth values, and compares them to `truth_perturbed.json`. Live-formula
workbooks pass ~100%; value-dump workbooks fail on every dependent cell.
The `audit_probe_pass_pct` dimension is folded into the composite at 0.20
weight (subtracted from structural and timing — see WEIGHTS below). For
scenarios without a `truth_perturbed.json` (v3 and older), the dimension
is None and its weight is redistributed back to correctness; see
`_composite()`.

Dependencies: openpyxl, decimal, stdlib.
"""

from __future__ import annotations

import argparse
import json
import math
import sys
import traceback
from dataclasses import dataclass, field, asdict
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

try:
    from openpyxl import load_workbook
    from openpyxl.worksheet.worksheet import Worksheet
except ImportError:  # pragma: no cover
    load_workbook = None  # type: ignore
    Worksheet = Any  # type: ignore

try:
    from harness.extract import extract as _semantic_extract
    from harness.extract import perturb_and_recalc as _perturb_and_recalc
except ImportError:  # pragma: no cover
    try:
        from extract import extract as _semantic_extract  # type: ignore
        from extract import perturb_and_recalc as _perturb_and_recalc  # type: ignore
    except ImportError:
        _semantic_extract = None  # type: ignore
        _perturb_and_recalc = None  # type: ignore


# ---------------------------------------------------------------------------
# Tunable targets — edit here, not in the body of the code.
# ---------------------------------------------------------------------------

# Rubric weights. Robustness handled in reporting layer (not here).
# Audit-probe (v4+) takes 0.20 from structural+timing. When the audit-probe
# is N/A (older scenarios with no truth_perturbed.json), its weight is
# redistributed back to correctness. See _composite().
WEIGHTS: dict[str, float] = {
    "correctness": 0.50,    # uses VC-usable tier as the "correctness subscore"
    "auditability": 0.20,
    "structural": 0.05,     # was 0.10 pre-v4
    "timing": 0.05,         # was 0.15 pre-v4
    "audit_probe": 0.20,    # new in v4
    # Robustness (0.05) is allocated in the reporting layer, not here.
}

# Timing targets (wall clock, submit-to-complete). Tune freely.
TIMING_TARGET_FULL_CREDIT_S: float = 600.0      # 10 min
TIMING_TARGET_ZERO_CREDIT_S: float = 1800.0     # 30 min

# Auditability formula-fraction thresholds
FORMULA_FRACTION_FULL: float = 0.95
FORMULA_FRACTION_MID: float = 0.75    # maps to 0.5
FORMULA_FRACTION_FLOOR: float = 0.50  # below here → 0

# Correctness tier thresholds
# Percentages: absolute percentage-point error (so 0.001 = 0.1pp)
# Dollars: relative error, fractional (so 0.001 = 0.1%)
TIER_STRICT_PP: Decimal = Decimal("0.001")      # 0.1 pp (pct stored as fraction)
TIER_STRICT_REL: Decimal = Decimal("0.001")     # 0.1%
TIER_VC_PP: Decimal = Decimal("0.01")           # 1 pp
TIER_VC_REL: Decimal = Decimal("0.01")          # 1%
TIER_DIR_PP: Decimal = Decimal("0.05")          # 5 pp
TIER_DIR_REL: Decimal = Decimal("0.05")         # 5%

# For the one absolute-dollar check (waterfall conservation): $1 abs tolerance.
WATERFALL_CONSERVATION_ABS_USD: Decimal = Decimal("1.00")


# ---------------------------------------------------------------------------
# Result types
# ---------------------------------------------------------------------------


@dataclass
class TierResult:
    tier_name: str
    n_keys: int
    n_passed: int
    n_missing: int
    details: list[dict[str, Any]] = field(default_factory=list)

    @property
    def pct_passed(self) -> float:
        if self.n_keys == 0:
            return 0.0
        return self.n_passed / self.n_keys


@dataclass
class ScoreResult:
    # Correctness — three tiers, independent
    correctness_strict: TierResult
    correctness_vc_usable: TierResult
    correctness_directional: TierResult

    # Auditability
    formula_fraction: float
    assumptions_tab_present: bool
    safe_conversion_detail_shown: bool
    auditability_score: float

    # Structural
    required_tabs_present: dict[str, bool]
    structural_score: float

    # Timing
    wall_clock_s: float
    submit_to_first_progress_s: float | None
    submit_to_complete_s: float | None
    complete_to_download_done_s: float | None
    credits_used: float | None
    timing_score: float

    # Composite
    composite_score: float

    # Full raw numbers for the report
    raw: dict = field(default_factory=dict)

    # Nash-election reasoning (isolated from dollar correctness).
    # Counts the tracked `.election` keys that match truth exactly.
    # For kelvin_v3 there are 25 (5 classes × 5 exits). Older scenarios
    # have fewer; scenarios without elections leave these at 0.
    elections_total: int = 0
    elections_correct: int = 0
    elections_correct_pct: float = 0.0

    # Audit-probe (v4+). None when the scenario ships no truth_perturbed.json,
    # in which case the audit-probe weight redistributes to correctness.
    audit_probe_pass_pct: float | None = None
    audit_probe_n_keys: int = 0
    audit_probe_n_passed: int = 0
    audit_probe_applied: list[str] = field(default_factory=list)
    audit_probe_unapplied: list[str] = field(default_factory=list)
    audit_probe_error: str | None = None


# ---------------------------------------------------------------------------
# Key list — KEEP IN SYNC WITH reconcile.py (see §11 of the spec).
# TODO: later refactor: lift key-iteration into a shared module (harness/keys.py)
# and import from it rather than duplicating here.
# ---------------------------------------------------------------------------


def _sensitivity_esop_token(esop: float) -> str:
    # 0.10 -> esop_10, 0.125 -> esop_12p5, 0.15 -> esop_15
    if esop == 0.125:
        return "esop_12p5"
    pct = int(round(esop * 100))
    return f"esop_{pct}"


def _infer_kind_from_key(key: str) -> str:
    """Map a flat key name to its tolerance kind. Rules match §11 conventions
    across v0/v1/v2/v3/v4: any last-segment ending in `_pct` (`.pct`,
    `.our_pct`, `.required_pct`, ...) → pct; `.new_ratio` → ratio (unit-less,
    bounded ≥1.0, MUST NOT be /100-normalized); `.shares*` → shares;
    `.election` → election; `.binding*` → binding; any `.dollars`/`.usd`/
    `.conversion_price*`/`usd_to_*`/`_exit` → usd; else string."""
    lower = key.lower()
    last_seg = lower.split(".")[-1]
    if lower.endswith(".election"):
        return "election"
    if "binding_term" in lower or lower.endswith(".binding"):
        return "binding"
    # Ratios are unit-less and bounded in [1.0, ∞) for triggered AD classes —
    # they are NOT percentages and must not be /100-normalized. Keep separate
    # from pct so the defined-name extractor branch can dispatch correctly.
    if lower.endswith(".new_ratio"):
        return "ratio"
    # Any last-segment ending in `_pct` (incl. `.pct`, `.our_pct`,
    # `.required_pct`, future `.X_pct`).
    if last_seg.endswith("_pct") or last_seg == "pct":
        return "pct"
    if "shares" in last_seg:
        return "shares"
    if lower.endswith(".total_distributed"):
        return "usd_abs_1"
    if (lower.endswith(".dollars") or lower.endswith("_usd")
            or "dollars_at" in lower
            or "conversion_price" in lower or "accreted_principal" in lower
            or "usd_to_" in lower):
        return "usd"
    return "string"


def iter_tracked_keys(truth_json_path: "Path | None" = None) -> list[dict[str, Any]]:
    """
    Return a list of {key, kind} dicts for every tracked output in §11.

    If `truth_json_path` is provided, derive the key set from truth.json —
    scenario-agnostic, works for v0/v1/v2/v3 and beyond. Falls back to the
    hardcoded v1 superset if no path is given (legacy behavior).

    kind ∈ {"pct", "usd", "shares", "election", "binding", "string"}.
    Kinds determine tolerance semantics.
    """
    if truth_json_path is not None:
        try:
            import json as _json
            with open(truth_json_path) as _f:
                data = _json.load(_f)
            # Flatten nested dicts to dot-notation if needed; truth.json is
            # already flat per all scenarios' Python builders.
            return [{"key": k, "kind": _infer_kind_from_key(k)}
                    for k in sorted(data.keys())]
        except Exception:
            pass  # fall through to legacy v1 superset
    keys: list[dict[str, Any]] = []

    # 11.1 Pro-forma ownership
    ownership_parts = [
        "founders", "employees", "options_issued", "options_unissued",
        "seed", "series_a", "safes", "venture_debt", "warrants",
        "series_b_new", "our_fund",
    ]
    for p in ownership_parts:
        keys.append({"key": f"ownership.{p}.pct", "kind": "pct"})

    # 11.2 Per-SAFE (15 × 3)
    for i in range(1, 16):
        keys.append({"key": f"safe.{i}.binding_term", "kind": "binding"})
        keys.append({"key": f"safe.{i}.conversion_price_usd", "kind": "usd"})
        keys.append({"key": f"safe.{i}.shares_issued", "kind": "shares"})

    # 11.3 Venture debt
    keys.append({"key": "venture_debt.accreted_principal_usd", "kind": "usd"})
    keys.append({"key": "venture_debt.binding_term", "kind": "binding"})
    keys.append({"key": "venture_debt.conversion_price_usd", "kind": "usd"})
    keys.append({"key": "venture_debt.shares_issued", "kind": "shares"})
    keys.append({"key": "venture_debt.vdw_shares", "kind": "shares"})

    # 11.4 Anti-dilution
    for series in ("seed", "series_a"):
        keys.append({"key": f"antidilution.{series}.new_conversion_price_usd", "kind": "usd"})
        keys.append({"key": f"antidilution.{series}.new_ratio", "kind": "ratio"})  # unit-less, ≥1.0; pp-abs ok

    # 11.5 Waterfall — 5 exits × (dollars lines + elections + total)
    waterfall_dollar_classes = [
        "venture_debt_as_b", "series_b", "series_a", "seed", "safes",
        "common", "options_issued", "warrants", "our_fund",
    ]
    waterfall_election_classes = [
        ("series_b", "election"),
        ("series_a", "election"),
        ("seed", "election"),
    ]
    for exit_m in (100, 250, 500, 1000, 2000):
        for cls in waterfall_dollar_classes:
            keys.append({"key": f"waterfall.{exit_m}M.{cls}.dollars", "kind": "usd"})
        for cls, sub in waterfall_election_classes:
            keys.append({"key": f"waterfall.{exit_m}M.{cls}.{sub}", "kind": "election"})
        keys.append({
            "key": f"waterfall.{exit_m}M.total_distributed",
            "kind": "usd_abs_1",  # special: ±$1 absolute, not relative
        })

    # 11.6 Sensitivity — 9 × 2
    for esop in (0.10, 0.125, 0.15):
        tok = _sensitivity_esop_token(esop)
        for pre in (60, 80, 120):
            keys.append({
                "key": f"sensitivity.{tok}.pre_{pre}M.our_pct",
                "kind": "pct",
            })
            keys.append({
                "key": f"sensitivity.{tok}.pre_{pre}M.dollars_at_500M_exit",
                "kind": "usd",
            })

    # 11.7 Return solver — 5 × 2
    for mult in ("3x", "4x"):
        for exit_m in (100, 250, 500, 1000, 2000):
            keys.append({
                "key": f"solver.{mult}.exit_{exit_m}M.required_pct",
                "kind": "pct",
            })

    return keys


# ---------------------------------------------------------------------------
# Output xlsx — defined-name lookup & region iteration
# ---------------------------------------------------------------------------


def _key_to_named_range(key: str) -> str:
    """Excel defined names can't use '.'. Map dots to underscores per §14."""
    return key.replace(".", "_")


def _resolve_defined_name(wb, named: str) -> Any:
    """
    Resolve a workbook-level defined name to its first cell's value.
    Returns None if not found or unresolvable.
    """
    try:
        dn = wb.defined_names.get(named)
    except Exception:
        dn = None
    if dn is None:
        # Case-insensitive fallback
        try:
            for name in wb.defined_names:
                if name.lower() == named.lower():
                    dn = wb.defined_names[name]
                    break
        except Exception:
            dn = None
    if dn is None:
        return None
    try:
        dests = list(dn.destinations)
    except Exception:
        return None
    if not dests:
        return None
    sheet_name, coord = dests[0]
    try:
        ws = wb[sheet_name]
    except Exception:
        return None
    try:
        cell_range = ws[coord]
    except Exception:
        return None
    # May be a single cell or a tuple-of-tuples
    if isinstance(cell_range, tuple):
        try:
            return cell_range[0][0].value
        except Exception:
            try:
                return cell_range[0].value
            except Exception:
                return None
    return getattr(cell_range, "value", None)


def _extract_output_value(wb, key: str) -> Any:
    """
    Try to resolve a tracked key's value from the output workbook.
    Primary path: workbook-level defined name `key_with_underscores`.
    Fallback: None (missing).
    """
    named = _key_to_named_range(key)
    val = _resolve_defined_name(wb, named)
    return val


# ---------------------------------------------------------------------------
# Correctness
# ---------------------------------------------------------------------------


def _to_decimal(x: Any) -> Decimal | None:
    """Coerce a value to Decimal, handling None/strings/ints/floats."""
    if x is None:
        return None
    if isinstance(x, bool):
        return None  # bools are not numbers here
    if isinstance(x, Decimal):
        return x
    if isinstance(x, (int,)):
        return Decimal(x)
    if isinstance(x, float):
        if math.isnan(x) or math.isinf(x):
            return None
        # Round-trip via str to avoid float cruft
        return Decimal(str(x))
    if isinstance(x, str):
        s = x.strip().replace(",", "").replace("$", "").replace("%", "")
        if not s:
            return None
        try:
            return Decimal(s)
        except (InvalidOperation, ValueError):
            return None
    return None


def _compare_one(
    kind: str,
    truth_val: Any,
    out_val: Any,
    pp_tol: Decimal,
    rel_tol: Decimal,
) -> tuple[bool, dict[str, Any]]:
    """
    Compare a single value against truth at a given tier. Returns (passed, detail).
    kind ∈ {"pct", "ratio", "usd", "shares", "election", "binding", "string", "usd_abs_1"}.
    """
    detail: dict[str, Any] = {"truth": truth_val, "output": out_val, "kind": kind}

    if out_val is None:
        detail["reason"] = "missing"
        return False, detail

    if kind in ("election", "binding", "string"):
        ok = str(truth_val).strip().lower() == str(out_val).strip().lower()
        if not ok:
            detail["reason"] = "string_mismatch"
        return ok, detail

    t = _to_decimal(truth_val)
    o = _to_decimal(out_val)
    if t is None or o is None:
        detail["reason"] = "non_numeric"
        return False, detail

    diff = (o - t).copy_abs()

    if kind in ("pct", "ratio"):
        # AD ratios share the same absolute tolerance class as percentages
        # (§13: pp-abs ≤ pp_tol). Distinct kind so extractors can skip
        # /100-normalization for ratios.
        ok = diff <= pp_tol
        detail["abs_pp_err"] = float(diff)
        detail["tol_pp"] = float(pp_tol)
        if not ok:
            detail["reason"] = "pp_out_of_tol"
        return ok, detail

    if kind == "shares":
        # Share counts: integers. Exact match at strict; allow small rel at lower tiers.
        if t == 0:
            ok = diff == 0
        else:
            # Treat like a relative check, reusing rel_tol
            rel = diff / t.copy_abs()
            ok = rel <= rel_tol
            detail["rel_err"] = float(rel)
        detail["abs_err"] = float(diff)
        if not ok:
            detail["reason"] = "shares_out_of_tol"
        return ok, detail

    if kind == "usd_abs_1":
        ok = diff <= WATERFALL_CONSERVATION_ABS_USD
        detail["abs_usd_err"] = float(diff)
        if not ok:
            detail["reason"] = "waterfall_conservation_violated"
        return ok, detail

    if kind == "usd":
        if t == 0:
            # Relative-error trap. Fall back to absolute-$0.01 for exact-zero truth.
            ok = diff <= Decimal("0.01")
            detail["abs_err"] = float(diff)
        else:
            rel = diff / t.copy_abs()
            ok = rel <= rel_tol
            detail["rel_err"] = float(rel)
            detail["tol_rel"] = float(rel_tol)
        if not ok:
            detail["reason"] = "usd_out_of_tol"
        return ok, detail

    # Unknown kind
    detail["reason"] = f"unknown_kind:{kind}"
    return False, detail


def _score_tier(
    tier_name: str,
    truth: dict,
    output_vals: dict[str, Any],
    pp_tol: Decimal,
    rel_tol: Decimal,
    tracked: list[dict[str, Any]] | None = None,
) -> TierResult:
    keys = tracked if tracked is not None else iter_tracked_keys()
    n_keys = len(keys)
    n_passed = 0
    n_missing = 0
    details: list[dict[str, Any]] = []
    for k in keys:
        key = k["key"]
        kind = k["kind"]
        truth_val = truth.get(key)
        out_val = output_vals.get(key)
        passed, detail = _compare_one(kind, truth_val, out_val, pp_tol, rel_tol)
        detail["key"] = key
        detail["passed"] = passed
        if out_val is None:
            n_missing += 1
        if passed:
            n_passed += 1
        details.append(detail)
    return TierResult(
        tier_name=tier_name,
        n_keys=n_keys,
        n_passed=n_passed,
        n_missing=n_missing,
        details=details,
    )


# ---------------------------------------------------------------------------
# Auditability
# ---------------------------------------------------------------------------


TRACKED_REGION_TAB_HINTS = (
    "pro-forma", "pro_forma", "proforma", "pro forma",
    "waterfall",
    "sensitivity",
    "return-solver", "return_solver", "return solver", "returns",
)


def _sheet_matches_any(name: str, hints: tuple[str, ...]) -> bool:
    n = name.strip().lower()
    for h in hints:
        if h in n:
            return True
    return False


def _iter_tracked_sheets(wb):
    for sn in wb.sheetnames:
        if _sheet_matches_any(sn, TRACKED_REGION_TAB_HINTS):
            yield wb[sn]


def _compute_formula_fraction(wb) -> float:
    formulas = 0
    hardcodes = 0
    for ws in _iter_tracked_sheets(wb):
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if v is None:
                    continue
                dt = getattr(cell, "data_type", None)
                if dt == "f":
                    formulas += 1
                elif isinstance(v, (int, float, Decimal)) and not isinstance(v, bool):
                    hardcodes += 1
                else:
                    # strings / labels — skip
                    pass
    total = formulas + hardcodes
    if total == 0:
        return 0.0
    return formulas / total


def _formula_fraction_to_score(frac: float) -> float:
    if frac >= FORMULA_FRACTION_FULL:
        return 1.0
    if frac <= FORMULA_FRACTION_FLOOR:
        return 0.0
    if frac >= FORMULA_FRACTION_MID:
        # Linear from MID→FULL mapped to 0.5→1.0
        span = FORMULA_FRACTION_FULL - FORMULA_FRACTION_MID
        return 0.5 + 0.5 * (frac - FORMULA_FRACTION_MID) / span
    # FLOOR → MID mapped to 0.0 → 0.5
    span = FORMULA_FRACTION_MID - FORMULA_FRACTION_FLOOR
    return 0.5 * (frac - FORMULA_FRACTION_FLOOR) / span


def _has_assumptions_tab(wb) -> bool:
    for sn in wb.sheetnames:
        if "assumption" in sn.strip().lower():
            return True
    return False


_SAFE_KEYWORDS = ("cap", "discount", "binding", "term", "mfn", "conversion")


def _safe_conversion_detail_shown(wb) -> bool:
    """
    Heuristic: look for a sheet named ~"Pro-Forma" or ~"SAFE*" that contains at
    least 15 rows AND whose header row shows ≥ 2 of the SAFE-related keywords.
    """
    candidate_sheets = []
    for sn in wb.sheetnames:
        low = sn.strip().lower()
        if "safe" in low or "pro-forma" in low or "proforma" in low or "pro forma" in low:
            candidate_sheets.append(wb[sn])

    for ws in candidate_sheets:
        # Scan the first ~30 rows for a header row with ≥ 2 SAFE keywords
        header_row_idx = None
        for r_idx, row in enumerate(ws.iter_rows(max_row=30, values_only=True), start=1):
            if not row:
                continue
            joined = " ".join(str(c).lower() for c in row if c is not None)
            hits = sum(1 for kw in _SAFE_KEYWORDS if kw in joined)
            if hits >= 2:
                header_row_idx = r_idx
                break
        if header_row_idx is None:
            continue
        # Count non-empty data rows after the header
        data_rows = 0
        for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
            if any(c is not None and str(c).strip() != "" for c in row):
                data_rows += 1
            if data_rows >= 15:
                return True
    return False


def _auditability_score(
    formula_fraction: float,
    assumptions_present: bool,
    safe_detail: bool,
) -> float:
    # Weight: formula-fraction is the big signal; the two booleans are smaller.
    ff = _formula_fraction_to_score(formula_fraction)
    # 0.6 formulas, 0.2 assumptions, 0.2 safe-detail
    return 0.6 * ff + 0.2 * (1.0 if assumptions_present else 0.0) + 0.2 * (1.0 if safe_detail else 0.0)


# ---------------------------------------------------------------------------
# Structural fidelity
# ---------------------------------------------------------------------------


REQUIRED_TABS: list[tuple[str, tuple[str, ...]]] = [
    ("Assumptions", ("assumption",)),
    ("Pro-Forma", ("pro-forma", "proforma", "pro forma", "pro_forma")),
    ("Waterfall", ("waterfall",)),
    ("Sensitivity", ("sensitivity",)),
    ("Return-Solver", ("return-solver", "return_solver", "return solver", "returns")),
]


def _sheet_is_nonempty(ws) -> bool:
    try:
        for row in ws.iter_rows(values_only=True):
            if any(c is not None and str(c).strip() != "" for c in row):
                return True
    except Exception:
        return False
    return False


def _structural_fidelity(wb) -> tuple[dict[str, bool], float]:
    present: dict[str, bool] = {}
    sheet_names = [sn.strip().lower() for sn in wb.sheetnames]
    for label, hints in REQUIRED_TABS:
        found = False
        for sn, low in zip(wb.sheetnames, sheet_names):
            if any(h in low for h in hints):
                ws = wb[sn]
                if _sheet_is_nonempty(ws):
                    found = True
                    break
        present[label] = found
    score = sum(1 for v in present.values() if v) / len(REQUIRED_TABS)
    return present, score


# ---------------------------------------------------------------------------
# Timing
# ---------------------------------------------------------------------------


def _timing_score_from_wall(wall_s: float) -> float:
    if wall_s <= TIMING_TARGET_FULL_CREDIT_S:
        return 1.0
    if wall_s >= TIMING_TARGET_ZERO_CREDIT_S:
        return 0.0
    span = TIMING_TARGET_ZERO_CREDIT_S - TIMING_TARGET_FULL_CREDIT_S
    return 1.0 - (wall_s - TIMING_TARGET_FULL_CREDIT_S) / span


def _extract_timings(run_dict: dict) -> dict[str, Any]:
    """
    Pull the key timing intervals from a serialized RunResult dict.
    Uses monotonic ns where possible, falling back to None.
    """
    events = run_dict.get("events") or []

    def _first_ns(phase: str) -> int | None:
        for ev in events:
            if ev.get("phase") == phase:
                return ev.get("t_mono_ns")
        return None

    def _last_ns(phase: str) -> int | None:
        last = None
        for ev in events:
            if ev.get("phase") == phase:
                last = ev.get("t_mono_ns")
        return last

    start_ns = events[0].get("t_mono_ns") if events else None
    end_ns = events[-1].get("t_mono_ns") if events else None

    wall_s = None
    if start_ns is not None and end_ns is not None:
        wall_s = (end_ns - start_ns) / 1e9

    submit_ns = _first_ns("submit_start")
    first_progress_ns = _first_ns("progress")
    complete_ns = _first_ns("complete")
    download_done_ns = _last_ns("download_done")

    def _delta(a, b):
        if a is None or b is None:
            return None
        return (b - a) / 1e9

    return {
        "wall_clock_s": wall_s if wall_s is not None else 0.0,
        "submit_to_first_progress_s": _delta(submit_ns, first_progress_ns),
        "submit_to_complete_s": _delta(submit_ns, complete_ns),
        "complete_to_download_done_s": _delta(complete_ns, download_done_ns),
        "credits_used": run_dict.get("credits_used"),
    }


# ---------------------------------------------------------------------------
# Composite
# ---------------------------------------------------------------------------


def _composite(
    correctness_subscore: float,
    auditability: float,
    structural: float,
    timing: float,
    audit_probe: float | None = None,
) -> float:
    """Weighted composite over rubric dimensions.

    When `audit_probe` is None (scenario has no truth_perturbed.json), its
    0.20 weight is redistributed to correctness so older scenarios are not
    penalized for a sub-test that doesn't exist for them. v3/v0/paxos_v2
    all hit this branch.
    """
    w = WEIGHTS
    if audit_probe is None:
        # Redistribute audit_probe weight to correctness.
        eff_correctness_w = w["correctness"] + w.get("audit_probe", 0.0)
        total_w = (eff_correctness_w + w["auditability"]
                   + w["structural"] + w["timing"])
        s = (
            eff_correctness_w * correctness_subscore
            + w["auditability"] * auditability
            + w["structural"] * structural
            + w["timing"] * timing
        ) / total_w
    else:
        total_w = (w["correctness"] + w["auditability"]
                   + w["structural"] + w["timing"]
                   + w.get("audit_probe", 0.0))
        s = (
            w["correctness"] * correctness_subscore
            + w["auditability"] * auditability
            + w["structural"] * structural
            + w["timing"] * timing
            + w.get("audit_probe", 0.0) * audit_probe
        ) / total_w
    return max(0.0, min(1.0, s))


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------


def _fmt_seconds(s: float | None) -> str:
    if s is None:
        return "n/a"
    return f"{s:.1f}s"


def _fmt_pct(x: float) -> str:
    return f"{x * 100:.1f}%"


def render_report(sr: ScoreResult, meta: dict[str, Any]) -> str:
    api = meta.get("api", "?")
    scenario = meta.get("scenario", "?")
    replicate = meta.get("replicate", "?")

    lines: list[str] = []
    lines.append(f"# Score: {api} / {scenario} / replicate {replicate}")
    lines.append("")
    lines.append(f"**Composite: {sr.composite_score:.3f} / 1.000**")
    lines.append("")

    lines.append("## Correctness")
    lines.append("| Tier | Keys passed | % |")
    lines.append("|---|---|---|")
    for t in (sr.correctness_strict, sr.correctness_vc_usable, sr.correctness_directional):
        lines.append(
            f"| {t.tier_name} | {t.n_passed}/{t.n_keys} | {_fmt_pct(t.pct_passed)} |"
        )
    lines.append("")

    lines.append("## Auditability")
    lines.append(f"- formula_fraction: {sr.formula_fraction:.2f}")
    lines.append(f"- assumptions_tab_present: {'yes' if sr.assumptions_tab_present else 'no'}")
    lines.append(f"- safe_conversion_detail_shown: {'yes' if sr.safe_conversion_detail_shown else 'no'}")
    lines.append(f"- **subscore: {sr.auditability_score:.2f}**")
    lines.append("")

    lines.append("## Structural")
    for label, present in sr.required_tabs_present.items():
        lines.append(f"- {label}: {'yes' if present else 'missing'}")
    lines.append(f"- **subscore: {sr.structural_score:.2f}**")
    lines.append("")

    lines.append("## Timing")
    lines.append(f"- wall_clock: {_fmt_seconds(sr.wall_clock_s)}")
    lines.append(f"- submit -> first progress: {_fmt_seconds(sr.submit_to_first_progress_s)}")
    lines.append(f"- submit -> complete: {_fmt_seconds(sr.submit_to_complete_s)}")
    lines.append(f"- complete -> download done: {_fmt_seconds(sr.complete_to_download_done_s)}")
    lines.append(f"- credits_used: {sr.credits_used if sr.credits_used is not None else 'n/a'}")
    lines.append(f"- **subscore: {sr.timing_score:.2f}**")
    lines.append("")

    lines.append("## Audit-probe")
    if sr.audit_probe_pass_pct is None:
        skip = sr.raw.get("audit_probe_skip_reason", "n/a (no truth_perturbed.json)")
        lines.append(f"- skipped: {skip}")
    else:
        lines.append(f"- pass_pct: {_fmt_pct(sr.audit_probe_pass_pct)} "
                     f"({sr.audit_probe_n_passed}/{sr.audit_probe_n_keys})")
        if sr.audit_probe_applied:
            lines.append(f"- applied: {', '.join(sr.audit_probe_applied)}")
        if sr.audit_probe_unapplied:
            lines.append(f"- unapplied (no target cell found): "
                         f"{', '.join(sr.audit_probe_unapplied)}")
        if sr.audit_probe_error:
            lines.append(f"- error: {sr.audit_probe_error}")
    lines.append("")

    errs = sr.raw.get("errors") or []
    if errs:
        lines.append("## Errors (non-fatal)")
        for e in errs:
            lines.append(f"- {e}")
        lines.append("")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Audit-probe (v4+)
# ---------------------------------------------------------------------------


# Per-scenario perturbation specs. Maps scenario name → list of
# (defined_name_key, value, description) tuples. The harness consults this
# at score time; a missing entry means "no audit-probe configured for this
# scenario" → audit_probe_pass_pct = None.
#
# kelvin_v4 perturbation per CANONICAL_INPUTS.md §11:
#   RoundTerms.primary_raise_usd: $30M → $33M
#
# Adding a scenario: list its perturbation cell(s) here. Truth must also
# ship a sibling `truth_perturbed.json` recomputed under the same
# perturbations; otherwise score_audit_probe degrades gracefully to None.
SCENARIO_PERTURBATIONS: dict[str, list[tuple[str, float, str]]] = {
    "kelvin_v4": [
        ("RoundTerms.primary_raise_usd", 33_000_000.0,
         "Series F primary raise $30M → $33M"),
    ],
}


def _scenario_perturbations(scenario: str | None) -> dict[str, float] | None:
    """Return the perturbation dict for a scenario, or None."""
    if not scenario:
        return None
    spec = SCENARIO_PERTURBATIONS.get(scenario)
    if not spec:
        return None
    return {k: v for (k, v, _desc) in spec}


def score_audit_probe(
    *,
    canonical_output_xlsx: Path,
    perturbations: dict[str, float],
    truth_perturbed_path: Path,
    tracked_keys: list[dict[str, Any]] | None = None,
    pp_tol: Decimal = TIER_VC_PP,
    rel_tol: Decimal = TIER_VC_REL,
    work_dir: Path | None = None,
) -> dict[str, Any]:
    """Run the audit-probe sub-test against an engine's output workbook.

    Steps:
      1) `perturb_and_recalc(canonical_output_xlsx, perturbations, work)`
         writes perturbed values into the named input cells and runs
         LibreOffice headless to recalc. Returns the recalculated workbook.
      2) Re-extract all tracked keys from the recalculated workbook using
         the same `extract.extract()` we use for canonical scoring.
      3) Compare extracted values to `truth_perturbed.json` using
         VC-usable tolerance (1pp / 1%) — same tolerance class used for the
         primary VC-usable tier.

    Returns a dict with keys: pass_pct, n_keys, n_passed, n_missing,
    applied, unapplied, recalc_path, details, error. Errors at any stage
    are surfaced via the `error` field; pass_pct is 0.0 in that case.

    A live-formula workbook should hit ~1.0; a pure value-dump workbook
    should hit ~0.0 except for the cells we directly overwrote.
    """
    out: dict[str, Any] = {
        "pass_pct": 0.0,
        "n_keys": 0,
        "n_passed": 0,
        "n_missing": 0,
        "applied": [],
        "unapplied": [],
        "recalc_path": None,
        "details": [],
        "error": None,
    }

    if _perturb_and_recalc is None:
        out["error"] = "harness.extract.perturb_and_recalc unavailable"
        return out
    if _semantic_extract is None:
        out["error"] = "harness.extract.extract unavailable"
        return out
    if not Path(canonical_output_xlsx).exists():
        out["error"] = f"output xlsx missing: {canonical_output_xlsx}"
        return out
    if not Path(truth_perturbed_path).exists():
        out["error"] = f"truth_perturbed.json missing: {truth_perturbed_path}"
        return out

    # Working directory for the perturbed copy.
    if work_dir is None:
        work_dir = Path(canonical_output_xlsx).parent / "audit_probe"
    work_dir.mkdir(parents=True, exist_ok=True)
    perturbed_path = work_dir / "output_perturbed.xlsx"

    # 1) Perturb + recalc
    try:
        recalc_path = _perturb_and_recalc(
            Path(canonical_output_xlsx),
            perturbations,
            perturbed_path,
        )
    except Exception as e:
        out["error"] = f"perturb_and_recalc failed: {e}"
        return out
    out["recalc_path"] = str(recalc_path)

    # Pull the apply/unapply sidecar (best-effort).
    try:
        sidecar = Path(str(recalc_path) + ".perturb.json")
        if sidecar.exists():
            sc = json.loads(sidecar.read_text())
            out["applied"] = sc.get("applied", [])
            out["unapplied"] = sc.get("unapplied", [])
    except Exception:  # noqa: BLE001
        pass

    # 2) Re-extract from the recalculated workbook.
    # Pass truth_perturbed_path so the extractor uses the scenario's actual
    # key set (v3/v4 schemas have keys like notes.*, tender.*, round.* that
    # the legacy v1 superset doesn't know about).
    try:
        er = _semantic_extract(
            xlsx_path=Path(recalc_path),
            truth_json_path=Path(truth_perturbed_path),
        )
        extracted = dict(er.values)
    except Exception as e:
        out["error"] = f"extract on recalculated workbook failed: {e}"
        return out

    # 3) Load truth_perturbed.json and compare.
    try:
        truth_p = json.loads(Path(truth_perturbed_path).read_text())
        if not isinstance(truth_p, dict):
            out["error"] = "truth_perturbed.json is not a dict"
            return out
    except Exception as e:
        out["error"] = f"failed to load truth_perturbed.json: {e}"
        return out

    if tracked_keys is None:
        tracked_keys = iter_tracked_keys(truth_perturbed_path)

    n_keys = 0
    n_passed = 0
    n_missing = 0
    details: list[dict[str, Any]] = []
    for k_spec in tracked_keys:
        key = k_spec["key"]
        if key not in truth_p:
            continue  # only score keys present in perturbed truth
        n_keys += 1
        kind = k_spec.get("kind") or _infer_kind_from_key(key)
        truth_val = truth_p.get(key)
        out_val = extracted.get(key)
        if out_val is None:
            n_missing += 1
        passed, det = _compare_one(kind, truth_val, out_val, pp_tol, rel_tol)
        det["key"] = key
        det["passed"] = passed
        details.append(det)
        if passed:
            n_passed += 1

    pct = (n_passed / n_keys) if n_keys > 0 else 0.0
    out.update({
        "pass_pct": pct,
        "n_keys": n_keys,
        "n_passed": n_passed,
        "n_missing": n_missing,
        "details": details,
    })
    return out


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------


def score_run(
    *,
    output_xlsx_path: Path,
    truth_json_path: Path,
    run_result_path: Path,
    output_report_path: Path | None = None,
) -> ScoreResult:
    """
    Produce a ScoreResult for a single API run.

    This function never raises; errors are collected into `raw['errors']` and
    missing inputs result in floor scores.
    """
    raw: dict[str, Any] = {"errors": []}
    errors: list[str] = raw["errors"]

    # --- Load truth ---
    truth: dict = {}
    try:
        with Path(truth_json_path).open("r") as f:
            truth = json.load(f)
        if not isinstance(truth, dict):
            errors.append(f"truth.json is not a dict, got {type(truth).__name__}")
            truth = {}
    except Exception as e:
        errors.append(f"failed to load truth.json: {e}")

    # --- Load run result ---
    run_dict: dict = {}
    try:
        with Path(run_result_path).open("r") as f:
            run_dict = json.load(f)
        if not isinstance(run_dict, dict):
            errors.append(f"run_result is not a dict")
            run_dict = {}
    except Exception as e:
        errors.append(f"failed to load run result: {e}")

    # --- Load output workbook ---
    wb = None
    wb_data = None
    if load_workbook is None:
        errors.append("openpyxl not available; structural & auditability scores will be 0")
    else:
        try:
            # data_only=False: preserves formula data_type info for auditability
            wb = load_workbook(str(output_xlsx_path), data_only=False)
        except Exception as e:
            errors.append(f"failed to open output xlsx (formula mode): {e}")
        try:
            # data_only=True: gives cached values for correctness comparison
            wb_data = load_workbook(str(output_xlsx_path), data_only=True)
        except Exception as e:
            errors.append(f"failed to open output xlsx (data mode): {e}")

    # --- Extract output values ---
    # First try the layout-agnostic semantic extractor (handles both
    # defined-name cases and arbitrary layouts). Fall back to bare
    # defined-name lookup if that module isn't importable for some reason.
    # Derive the scenario-specific key set from truth.json. Enables v2/v3/v*
    # without per-scenario scorer edits. Falls back to v1 superset in
    # iter_tracked_keys() if truth.json isn't readable.
    tracked = iter_tracked_keys(truth_json_path)

    output_vals: dict[str, Any] = {}
    if _semantic_extract is not None:
        try:
            # Forward truth_json_path so the extractor's key set matches the
            # scenario's schema (works for v0/v1/v2/v3/v4+).
            er = _semantic_extract(
                xlsx_path=Path(output_xlsx_path),
                truth_json_path=Path(truth_json_path),
            )
            output_vals.update(er.values)
            # Record extractor stats for downstream reporting
            raw["extractor"] = {
                "extracted": sum(1 for k in tracked if er.values.get(k["key"]) is not None),
                "missing": list(er.missing),
                "ambiguous": [a[0] for a in er.ambiguous],
            }
        except Exception as e:
            errors.append(f"semantic extractor failed: {e}")
    # Ensure every tracked key has an entry (None for missing) so downstream
    # tier scoring can distinguish "missing" from "wrong".
    if wb_data is not None:
        for k in tracked:
            if output_vals.get(k["key"]) is None:
                try:
                    # Fallback path: direct defined-name lookup
                    fallback = _extract_output_value(wb_data, k["key"])
                    if fallback is not None:
                        output_vals[k["key"]] = fallback
                    else:
                        output_vals.setdefault(k["key"], None)
                except Exception as e:
                    errors.append(f"value extract failed for {k['key']}: {e}")
                    output_vals.setdefault(k["key"], None)

    # --- Correctness tiers ---
    try:
        strict = _score_tier("Strict (<=0.1pp/0.1%)", truth, output_vals, TIER_STRICT_PP, TIER_STRICT_REL, tracked=tracked)
    except Exception as e:
        errors.append(f"strict tier failed: {e}\n{traceback.format_exc()}")
        strict = TierResult("Strict (<=0.1pp/0.1%)", 0, 0, 0, [])
    try:
        vc = _score_tier("VC-usable (<=1pp/1%)", truth, output_vals, TIER_VC_PP, TIER_VC_REL, tracked=tracked)
    except Exception as e:
        errors.append(f"vc tier failed: {e}")
        vc = TierResult("VC-usable (<=1pp/1%)", 0, 0, 0, [])
    try:
        directional = _score_tier("Directional (<=5pp/5%)", truth, output_vals, TIER_DIR_PP, TIER_DIR_REL, tracked=tracked)
    except Exception as e:
        errors.append(f"directional tier failed: {e}")
        directional = TierResult("Directional (<=5pp/5%)", 0, 0, 0, [])

    # --- Auditability ---
    formula_fraction = 0.0
    assumptions_present = False
    safe_detail = False
    if wb is not None:
        try:
            formula_fraction = _compute_formula_fraction(wb)
        except Exception as e:
            errors.append(f"formula_fraction failed: {e}")
        try:
            assumptions_present = _has_assumptions_tab(wb)
        except Exception as e:
            errors.append(f"assumptions_tab detect failed: {e}")
        try:
            safe_detail = _safe_conversion_detail_shown(wb)
        except Exception as e:
            errors.append(f"safe_conversion detect failed: {e}")
    auditability = _auditability_score(formula_fraction, assumptions_present, safe_detail)

    # --- Structural ---
    required_tabs_present: dict[str, bool] = {label: False for label, _ in REQUIRED_TABS}
    structural = 0.0
    target_wb = wb or wb_data
    if target_wb is not None:
        try:
            required_tabs_present, structural = _structural_fidelity(target_wb)
        except Exception as e:
            errors.append(f"structural fidelity failed: {e}")

    # --- Timing ---
    timings = {
        "wall_clock_s": 0.0,
        "submit_to_first_progress_s": None,
        "submit_to_complete_s": None,
        "complete_to_download_done_s": None,
        "credits_used": None,
    }
    try:
        timings.update(_extract_timings(run_dict))
    except Exception as e:
        errors.append(f"timing extraction failed: {e}")
    wall_s = timings["wall_clock_s"] or 0.0
    timing_score = _timing_score_from_wall(wall_s) if wall_s > 0 else 0.0

    # --- Nash elections: isolate string-match accuracy on .election keys ---
    # Separable from dollar correctness because a model can pick the right
    # election and then round the dollars slightly wrong, or vice versa.
    # For kelvin_v3 there are 25 election keys (5 classes × 5 exits).
    elections_total = 0
    elections_correct = 0
    try:
        for k_spec in tracked:
            k = k_spec["key"]
            if k_spec.get("kind") == "election" or k.endswith(".election"):
                if k not in truth:
                    continue
                elections_total += 1
                tval = truth.get(k)
                oval = output_vals.get(k)
                if (tval is not None and oval is not None
                        and str(tval).strip().lower() == str(oval).strip().lower()):
                    elections_correct += 1
    except Exception as e:
        errors.append(f"elections_correct_pct failed: {e}")
    elections_correct_pct = (
        elections_correct / elections_total if elections_total > 0 else 0.0
    )

    # --- Audit-probe (v4+) ---
    # Run the perturb-and-recalc sub-test if the scenario has both a
    # registered perturbation in SCENARIO_PERTURBATIONS and a
    # truth_perturbed.json sibling next to truth_json_path. v3/v0/paxos_v2
    # have no perturbed truth; the dimension stays None and its weight
    # redistributes back to correctness in _composite().
    audit_probe_pass_pct: float | None = None
    audit_probe_n_keys = 0
    audit_probe_n_passed = 0
    audit_probe_applied: list[str] = []
    audit_probe_unapplied: list[str] = []
    audit_probe_error: str | None = None

    scenario_name = run_dict.get("scenario") if isinstance(run_dict, dict) else None
    perturbations = _scenario_perturbations(scenario_name)
    truth_perturbed_path = Path(truth_json_path).parent / "truth_perturbed.json"

    if (perturbations is not None
            and truth_perturbed_path.exists()
            and Path(output_xlsx_path).exists()):
        try:
            ap = score_audit_probe(
                canonical_output_xlsx=Path(output_xlsx_path),
                perturbations=perturbations,
                truth_perturbed_path=truth_perturbed_path,
                tracked_keys=tracked,
                pp_tol=TIER_VC_PP,
                rel_tol=TIER_VC_REL,
            )
            audit_probe_pass_pct = float(ap["pass_pct"])
            audit_probe_n_keys = int(ap["n_keys"])
            audit_probe_n_passed = int(ap["n_passed"])
            audit_probe_applied = list(ap.get("applied") or [])
            audit_probe_unapplied = list(ap.get("unapplied") or [])
            audit_probe_error = ap.get("error")
            if audit_probe_error:
                errors.append(f"audit_probe: {audit_probe_error}")
        except Exception as e:
            errors.append(f"audit_probe failed: {e}\n{traceback.format_exc()}")
            audit_probe_error = str(e)
            audit_probe_pass_pct = 0.0  # collapse to zero rather than None on hard error
    else:
        # Reasons for skipping (recorded but non-fatal):
        if perturbations is None and scenario_name:
            raw.setdefault("audit_probe_skip_reason",
                           f"no perturbations registered for scenario {scenario_name!r}")
        elif not truth_perturbed_path.exists():
            raw.setdefault("audit_probe_skip_reason",
                           f"truth_perturbed.json missing: {truth_perturbed_path}")
        elif not Path(output_xlsx_path).exists():
            raw.setdefault("audit_probe_skip_reason",
                           f"output xlsx missing: {output_xlsx_path}")

    # --- Composite: use VC-usable as the "correctness subscore" ---
    correctness_subscore = vc.pct_passed
    composite = _composite(
        correctness_subscore, auditability, structural, timing_score,
        audit_probe=audit_probe_pass_pct,
    )

    raw.update({
        "key_count": len(tracked),
        "formula_fraction": formula_fraction,
        "timings": timings,
        "output_vals_sample": {
            k: output_vals.get(k)
            for k in list(output_vals)[:5]
        },
        "weights": WEIGHTS,
        "timing_targets": {
            "full_credit_s": TIMING_TARGET_FULL_CREDIT_S,
            "zero_credit_s": TIMING_TARGET_ZERO_CREDIT_S,
        },
    })

    sr = ScoreResult(
        correctness_strict=strict,
        correctness_vc_usable=vc,
        correctness_directional=directional,
        formula_fraction=formula_fraction,
        assumptions_tab_present=assumptions_present,
        safe_conversion_detail_shown=safe_detail,
        auditability_score=auditability,
        required_tabs_present=required_tabs_present,
        structural_score=structural,
        elections_total=elections_total,
        elections_correct=elections_correct,
        elections_correct_pct=elections_correct_pct,
        wall_clock_s=wall_s,
        submit_to_first_progress_s=timings.get("submit_to_first_progress_s"),
        submit_to_complete_s=timings.get("submit_to_complete_s"),
        complete_to_download_done_s=timings.get("complete_to_download_done_s"),
        credits_used=timings.get("credits_used"),
        timing_score=timing_score,
        composite_score=composite,
        audit_probe_pass_pct=audit_probe_pass_pct,
        audit_probe_n_keys=audit_probe_n_keys,
        audit_probe_n_passed=audit_probe_n_passed,
        audit_probe_applied=audit_probe_applied,
        audit_probe_unapplied=audit_probe_unapplied,
        audit_probe_error=audit_probe_error,
        raw=raw,
    )

    # --- Optional report emission ---
    if output_report_path is not None:
        try:
            meta = {
                "api": run_dict.get("api", "?"),
                "scenario": run_dict.get("scenario", "?"),
                "replicate": run_dict.get("replicate", "?"),
            }
            md = render_report(sr, meta)
            Path(output_report_path).parent.mkdir(parents=True, exist_ok=True)
            Path(output_report_path).write_text(md)
        except Exception as e:
            errors.append(f"report emission failed: {e}")

    return sr


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def _cli(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Score a single API run against ground truth.")
    p.add_argument("--output", required=True, help="Path to API's output xlsx")
    p.add_argument("--truth", required=True, help="Path to truth.json")
    p.add_argument("--run", required=True, help="Path to serialized RunResult JSON")
    p.add_argument("--report", default=None, help="Optional: path to write a markdown report")
    args = p.parse_args(argv)

    sr = score_run(
        output_xlsx_path=Path(args.output),
        truth_json_path=Path(args.truth),
        run_result_path=Path(args.run),
        output_report_path=Path(args.report) if args.report else None,
    )

    # Emit a compact JSON summary on stdout. Don't dump full per-key details.
    summary = {
        "composite_score": sr.composite_score,
        "correctness_strict_pct": sr.correctness_strict.pct_passed,
        "correctness_vc_usable_pct": sr.correctness_vc_usable.pct_passed,
        "correctness_directional_pct": sr.correctness_directional.pct_passed,
        "auditability_score": sr.auditability_score,
        "structural_score": sr.structural_score,
        "timing_score": sr.timing_score,
        "audit_probe_pass_pct": sr.audit_probe_pass_pct,
        "audit_probe_n_keys": sr.audit_probe_n_keys,
        "audit_probe_n_passed": sr.audit_probe_n_passed,
        "wall_clock_s": sr.wall_clock_s,
        "elections_total": sr.elections_total,
        "elections_correct": sr.elections_correct,
        "elections_correct_pct": sr.elections_correct_pct,
        "n_errors": len(sr.raw.get("errors") or []),
    }
    print(json.dumps(summary, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(_cli())
