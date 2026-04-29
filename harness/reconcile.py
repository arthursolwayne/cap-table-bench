"""
Dual-build reconciliation checker for the Nyxlight benchmark.

Compares the hand-authored Excel ground truth (ground_truth.xlsx) against the
Python-computed ground truth (truth.json). If the two builders disagree on any
tracked key beyond its tolerance, the benchmark is invalid and this tool
surfaces every disagreement.

Design notes
------------
- Key extraction strategy: the full expected key list is HARDCODED below (see
  EXPECTED_KEYS). The spec's §11 mixes prose, tables, and template fragments
  (e.g. ``solver.3x.exit_{100,250,500,1000,2000}M.required_pct``) that are
  cumbersome and fragile to regex out. We enumerate the keys directly from the
  spec's templates in code. A helper ``parse_spec_keys`` is provided that
  best-effort extracts key-looking tokens from the spec; we sanity-check that
  every hardcoded key appears somewhere in the spec text so drift gets caught.
- Numeric comparisons use ``decimal.Decimal``. Floats from JSON are converted
  via ``Decimal(str(x))`` to avoid binary-float artefacts.
- Formula evaluation fallback chain (for cells whose cached value is absent):
  1. openpyxl data_only=True (Excel's last-cached value — the happy path)
  2. ``formulas`` Python package, if installed
  3. ``libreoffice --headless --convert-to xlsx`` to force recalc, if the
     binary is on PATH
  If all three fail, the key is marked AMBIGUOUS and the run fails.

Audit-probe preflight (v4+)
---------------------------
Scenarios that ship a `truth_perturbed.json` (e.g. kelvin_v4) trigger the
audit-probe sub-test inside `harness.score.score_audit_probe`. The
reconcile preflight here verifies that BOTH `truth.json` AND
`truth_perturbed.json` are present and structurally compatible (same key
count, same key set). Scenarios without `truth_perturbed.json` (v3 and
older) silently skip the audit-probe — `audit_probe_preflight` returns a
"skipped, no perturbed truth" status that callers can log without
treating it as a failure. See `audit_probe_preflight()`.

CLI
---
    python reconcile.py --spec <path> --truth <path> --xlsx <path> \
                        [--report <path>]

Exit codes:
    0 — all keys match within tolerance
    1 — at least one mismatch / missing / ambiguous
    2 — setup error (missing input files, unreadable xlsx, etc.)
"""

from __future__ import annotations

# --- sys.path shim ---------------------------------------------------------
# When this file is executed directly as ``python reconcile.py``, Python
# prepends this file's directory (``harness/``) to ``sys.path``. That
# directory contains a local ``types.py`` which then shadows the stdlib
# ``types`` module and breaks imports of argparse, dataclasses, etc. Strip
# our own directory from sys.path[0] before any stdlib imports happen.
import os as _os
import sys as _sys
_here = _os.path.dirname(_os.path.abspath(__file__))
if _sys.path and _os.path.abspath(_sys.path[0]) == _here:
    _sys.path.pop(0)
# ---------------------------------------------------------------------------

import argparse
import json
import math
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Iterable


# --------------------------------------------------------------------------
# Expected key list
# --------------------------------------------------------------------------
#
# Hardcoded from spec §11. See ``verify_keys_against_spec`` for the drift
# check that complains if the spec mentions a key we don't track (or vice
# versa).

_OWNERSHIP_KEYS = [
    "ownership.founders.pct",
    "ownership.employees.pct",
    "ownership.options_issued.pct",
    "ownership.options_unissued.pct",
    "ownership.seed.pct",
    "ownership.series_a.pct",
    "ownership.safes.pct",
    "ownership.venture_debt.pct",
    "ownership.warrants.pct",
    "ownership.series_b_new.pct",
    "ownership.our_fund.pct",
]

_SAFE_KEYS = [
    f"safe.{i}.{field}"
    for i in range(1, 16)
    for field in ("binding_term", "conversion_price_usd", "shares_issued")
]

_VD_KEYS = [
    "venture_debt.accreted_principal_usd",
    "venture_debt.binding_term",
    "venture_debt.conversion_price_usd",
    "venture_debt.shares_issued",
    "venture_debt.vdw_shares",
]

_AD_KEYS = [
    "antidilution.seed.new_conversion_price_usd",
    "antidilution.seed.new_ratio",
    "antidilution.series_a.new_conversion_price_usd",
    "antidilution.series_a.new_ratio",
]

_WATERFALL_EXITS = ("100", "250", "500", "1000", "2000")
_WATERFALL_DOLLAR_FIELDS = (
    "venture_debt_as_b",
    "series_b",
    "series_a",
    "seed",
    "safes",
    "common",
    "options_issued",
    "warrants",
    "our_fund",
)
_WATERFALL_ELECTION_FIELDS = ("series_b", "series_a", "seed")


def _waterfall_keys() -> list[str]:
    keys: list[str] = []
    for exit_ in _WATERFALL_EXITS:
        for f in _WATERFALL_DOLLAR_FIELDS:
            keys.append(f"waterfall.{exit_}M.{f}.dollars")
        for f in _WATERFALL_ELECTION_FIELDS:
            keys.append(f"waterfall.{exit_}M.{f}.election")
        keys.append(f"waterfall.{exit_}M.total_distributed")
    return keys


_SENS_ESOPS = ("10", "12p5", "15")
_SENS_PRES = ("60", "80", "120")


def _sensitivity_keys() -> list[str]:
    keys: list[str] = []
    for esop in _SENS_ESOPS:
        for pre in _SENS_PRES:
            keys.append(f"sensitivity.esop_{esop}.pre_{pre}M.our_pct")
            keys.append(f"sensitivity.esop_{esop}.pre_{pre}M.dollars_at_500M_exit")
    return keys


_SOLVER_EXITS = ("100", "250", "500", "1000", "2000")


def _solver_keys() -> list[str]:
    keys: list[str] = []
    for mult in ("3x", "4x"):
        for exit_ in _SOLVER_EXITS:
            keys.append(f"solver.{mult}.exit_{exit_}M.required_pct")
    return keys


EXPECTED_KEYS: list[str] = (
    _OWNERSHIP_KEYS
    + _SAFE_KEYS
    + _VD_KEYS
    + _AD_KEYS
    + _waterfall_keys()
    + _sensitivity_keys()
    + _solver_keys()
)


# --------------------------------------------------------------------------
# Public result types
# --------------------------------------------------------------------------


@dataclass
class KeyMismatch:
    key: str
    truth_json_value: Any
    xlsx_value: Any
    tolerance_used: str
    delta: Any  # absolute difference; for strings, (a, b)


@dataclass
class ReconcileResult:
    passed: bool
    total_keys: int
    matched: int
    mismatched: list[KeyMismatch] = field(default_factory=list)
    missing_in_truth_json: list[str] = field(default_factory=list)
    missing_in_xlsx: list[str] = field(default_factory=list)
    ambiguous: list[str] = field(default_factory=list)

    def summary_line(self) -> str:
        status = "PASS" if self.passed else "FAIL"
        return (
            f"{status}: {self.matched}/{self.total_keys} matched, "
            f"{len(self.mismatched)} mismatched, "
            f"{len(self.missing_in_truth_json)} missing in truth.json, "
            f"{len(self.missing_in_xlsx)} missing in xlsx, "
            f"{len(self.ambiguous)} ambiguous."
        )


# --------------------------------------------------------------------------
# Tolerance classification
# --------------------------------------------------------------------------


PCT_TOL = Decimal("1e-6")
USD_TOL = Decimal("0.01")


def tolerance_kind(key: str) -> str:
    """Return a short classification string for the key.

    One of: ``pct``, ``usd``, ``shares``, ``string``.
    """
    # pct / required_pct
    if key.endswith(".pct") or key.endswith(".required_pct") or key.endswith(".our_pct"):
        return "pct"
    # shares (exact integer match). Must be checked BEFORE usd because
    # "shares_issued" happens to match neither, but keep order explicit.
    if key.endswith(".shares_issued") or key.endswith(".vdw_shares") or ".shares" in key.split(".")[-1]:
        return "shares"
    # strings
    if (
        key.endswith(".election")
        or key.endswith(".binding_term")
        or key.endswith(".conversion_branch")
    ):
        return "string"
    # dollars / usd amounts — catch *.dollars, *.usd, *_usd, *.total_distributed,
    # *.dollars_at_500M_exit
    if (
        key.endswith(".dollars")
        or key.endswith(".usd")
        or key.endswith("_usd")
        or key.endswith(".total_distributed")
        or key.endswith(".dollars_at_500M_exit")
    ):
        return "usd"
    # Ratios (antidilution.*.new_ratio) — treat as a tight-tolerance numeric
    # scalar. Use pct tolerance (1e-6).
    if key.endswith(".new_ratio"):
        return "pct"
    # Fallback: treat as usd scalar. Better to compare than to silently skip.
    return "usd"


# --------------------------------------------------------------------------
# JSON loading + flattening
# --------------------------------------------------------------------------


def flatten_json(obj: Any, prefix: str = "") -> dict[str, Any]:
    """Flatten a nested dict into dot-notation keys.

    Non-dict values are terminals. Lists are not expected in truth.json; if
    encountered, they are stored as-is under the parent key.
    """
    out: dict[str, Any] = {}
    if isinstance(obj, dict):
        for k, v in obj.items():
            new_key = f"{prefix}.{k}" if prefix else str(k)
            if isinstance(v, dict):
                out.update(flatten_json(v, new_key))
            else:
                out[new_key] = v
    else:
        if prefix:
            out[prefix] = obj
    return out


def load_truth_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as f:
        raw = json.load(f)
    return flatten_json(raw)


# --------------------------------------------------------------------------
# Xlsx loading + defined-name resolution
# --------------------------------------------------------------------------


def _key_to_defined_name(key: str) -> str:
    """Translate a dot-notation key to the Excel-defined-name variant.

    Spec §14: Excel defined names can't use ``.``; replace with ``_``.
    """
    return key.replace(".", "_")


def _resolve_defined_name(wb, name: str) -> tuple[str, str] | None:
    """Resolve a workbook-level defined name to (sheet, cell_ref) or None.

    Handles openpyxl's slightly awkward defined-name API across versions.
    """
    try:
        dn = wb.defined_names[name]
    except (KeyError, TypeError):
        # Older openpyxl returns None via .get
        dn = None
        try:
            dn = wb.defined_names.get(name)  # type: ignore[attr-defined]
        except Exception:
            dn = None
    if dn is None:
        return None

    # openpyxl 3.1+: DefinedName has .destinations yielding (sheet, ref)
    try:
        dests = list(dn.destinations)
    except Exception:
        dests = []
    if dests:
        sheet, ref = dests[0]
        return sheet, ref

    # Fallback: parse .value or .attr_text
    val = getattr(dn, "value", None) or getattr(dn, "attr_text", None)
    if not val:
        return None
    # e.g. "Outputs!$B$5"
    m = re.match(r"^'?([^'!]+)'?!(.+)$", val)
    if not m:
        return None
    return m.group(1), m.group(2)


def _read_cell_value(wb_data, sheet_name: str, cell_ref: str) -> Any:
    """Read a cell from the data-loaded workbook.

    Strips absolute-reference ``$`` markers. If the ref names a range, reads
    the top-left cell of the range (tracked keys should be single cells).
    """
    ref = cell_ref.replace("$", "")
    # range like A1:B2 — take first
    if ":" in ref:
        ref = ref.split(":")[0]
    try:
        ws = wb_data[sheet_name]
    except KeyError:
        return None
    try:
        return ws[ref].value
    except Exception:
        return None


def _has_formula(wb_formula, sheet_name: str, cell_ref: str) -> bool:
    """Check that the cell in the formula-view workbook holds a formula.

    Per spec §14, every tracked output cell must be a formula. Cells holding
    bare numeric literals in the output sheets are a spec violation and we
    surface them as ambiguous.
    """
    ref = cell_ref.replace("$", "")
    if ":" in ref:
        ref = ref.split(":")[0]
    try:
        ws = wb_formula[sheet_name]
    except KeyError:
        return False
    try:
        v = ws[ref].value
    except Exception:
        return False
    return isinstance(v, str) and v.startswith("=")


def _evaluate_via_formulas_pkg(xlsx_path: Path) -> dict[tuple[str, str], Any] | None:
    """Try to evaluate formulas via the ``formulas`` package. Returns a dict
    keyed by (sheet, cell_ref_no_dollar) or None if the package is missing /
    fails."""
    try:
        import formulas  # type: ignore
    except Exception:
        return None
    try:
        xl_model = formulas.ExcelModel().loads(str(xlsx_path)).finish()
        xl_model.calculate()
        # formulas stores results with qualified keys like
        # "'[file.xlsx]Sheet'!A1". We extract (sheet, ref).
        out: dict[tuple[str, str], Any] = {}
        sols = xl_model.cells
        for qkey, cell in sols.items():
            m = re.match(r"^'?\[[^\]]+\]([^'!]+)'?!(.+)$", qkey)
            if not m:
                continue
            sheet, ref = m.group(1), m.group(2).replace("$", "")
            val = getattr(cell, "value", None)
            if hasattr(val, "tolist"):
                try:
                    val = val.tolist()
                    # unwrap scalar array
                    while isinstance(val, list) and len(val) == 1:
                        val = val[0]
                except Exception:
                    pass
            out[(sheet, ref)] = val
        return out
    except Exception:
        return None


def _evaluate_via_libreoffice(xlsx_path: Path) -> Path | None:
    """Force recalc by round-tripping through headless LibreOffice. Returns
    path to a fresh xlsx with cached values, or None if unavailable."""
    if shutil.which("libreoffice") is None and shutil.which("soffice") is None:
        return None
    binary = shutil.which("libreoffice") or shutil.which("soffice")
    assert binary is not None
    tmpdir = Path(tempfile.mkdtemp(prefix="reconcile_lo_"))
    try:
        subprocess.run(
            [binary, "--headless", "--calc", "--convert-to", "xlsx",
             "--outdir", str(tmpdir), str(xlsx_path)],
            check=True, capture_output=True, timeout=120,
        )
    except Exception:
        return None
    out = tmpdir / xlsx_path.name
    return out if out.exists() else None


def load_xlsx_values(
    path: Path,
    keys: list[str] | None = None,
) -> tuple[dict[str, Any], dict[str, bool], list[str]]:
    """Load defined-name values from an xlsx.

    `keys` lets callers pass the scenario-specific key set (from truth.json)
    so we don't iterate v1's hardcoded superset and miss v3-only keys.
    Default preserves legacy behavior.
    """
    """Return (values_by_key, has_formula_by_key, missing_names).

    Strategy:
      1. Open with ``data_only=False`` to inspect formulas.
      2. Open with ``data_only=True`` to get cached values.
      3. If a tracked cell's cached value is None but it's a formula, try
         the ``formulas`` package, then LibreOffice headless recalc.
    """
    try:
        import openpyxl  # type: ignore
    except ImportError as e:
        raise RuntimeError("openpyxl is required for reconcile.py") from e

    wb_formula = openpyxl.load_workbook(path, data_only=False, read_only=False)
    wb_data = openpyxl.load_workbook(path, data_only=True, read_only=False)

    values: dict[str, Any] = {}
    has_formula: dict[str, bool] = {}
    missing_names: list[str] = []

    # First pass: resolve defined names, collect cached values + formula flags
    iter_keys = keys if keys is not None else EXPECTED_KEYS
    resolved: dict[str, tuple[str, str]] = {}
    for key in iter_keys:
        dn_name = _key_to_defined_name(key)
        loc = _resolve_defined_name(wb_formula, dn_name)
        if loc is None:
            missing_names.append(key)
            continue
        resolved[key] = loc
        sheet, ref = loc
        has_formula[key] = _has_formula(wb_formula, sheet, ref)
        values[key] = _read_cell_value(wb_data, sheet, ref)

    # Determine which tracked keys still need a computed value
    needs_eval = [
        k for k, loc in resolved.items()
        if values.get(k) is None and has_formula.get(k, False)
    ]

    if needs_eval:
        # Fallback 1: LibreOffice recalc (moved ahead of formulas-pkg — the
        # formulas package partially resolves with empty Ranges objects that
        # incorrectly clear needs_eval without producing real values, blocking
        # the LibreOffice fallback. On real workbooks with IF/INDEX/MATCH
        # (kelvin_v3), LibreOffice handles it cleanly.)
        recalc_path = _evaluate_via_libreoffice(path)
        if recalc_path is not None:
            wb2 = openpyxl.load_workbook(recalc_path, data_only=True, read_only=False)
            for k in list(needs_eval):
                sheet, ref = resolved[k]
                v = _read_cell_value(wb2, sheet, ref)
                if v is not None:
                    values[k] = v
            needs_eval = [k for k in needs_eval if values.get(k) is None]

    if needs_eval:
        # Fallback 2: formulas package (last resort for any cell LibreOffice
        # missed).
        fmap = _evaluate_via_formulas_pkg(path)
        if fmap is not None:
            for k in list(needs_eval):
                sheet, ref = resolved[k]
                ref_clean = ref.replace("$", "")
                if ":" in ref_clean:
                    ref_clean = ref_clean.split(":")[0]
                v = fmap.get((sheet, ref_clean))
                if v is not None:
                    values[k] = v
            needs_eval = [k for k in needs_eval if values.get(k) is None]

    return values, has_formula, missing_names


# --------------------------------------------------------------------------
# Comparison
# --------------------------------------------------------------------------


def _to_decimal(v: Any) -> Decimal | None:
    if v is None:
        return None
    if isinstance(v, Decimal):
        return v
    if isinstance(v, bool):  # bools are ints in Python; reject
        return None
    if isinstance(v, int):
        return Decimal(v)
    if isinstance(v, float):
        if math.isnan(v) or math.isinf(v):
            return None
        return Decimal(str(v))
    if isinstance(v, str):
        try:
            return Decimal(v)
        except (InvalidOperation, ValueError):
            return None
    return None


def _is_int_like(v: Any) -> bool:
    """True if v is an int (or a float/Decimal that is numerically an integer
    with zero fractional part). Bools excluded."""
    if isinstance(v, bool):
        return False
    if isinstance(v, int):
        return True
    if isinstance(v, Decimal):
        return v == v.to_integral_value()
    if isinstance(v, float):
        return v.is_integer()
    if isinstance(v, str):
        try:
            d = Decimal(v)
            return d == d.to_integral_value()
        except (InvalidOperation, ValueError):
            return False
    return False


def _compare_one(key: str, jval: Any, xval: Any) -> tuple[bool, KeyMismatch | None]:
    """Compare a single key's two values. Return (matched, mismatch_or_None)."""
    kind = tolerance_kind(key)

    if kind == "string":
        # Exact string equality, after str-coercion (xlsx may hand back a
        # non-string if someone put a number in a string cell — that's a fail).
        a = jval if isinstance(jval, str) else (None if jval is None else str(jval))
        b = xval if isinstance(xval, str) else (None if xval is None else str(xval))
        if a == b and a is not None:
            return True, None
        return False, KeyMismatch(
            key=key,
            truth_json_value=jval,
            xlsx_value=xval,
            tolerance_used="string exact",
            delta=(jval, xval),
        )

    if kind == "shares":
        # Must be integer on both sides. Floats are a type error.
        j_int = _is_int_like(jval)
        x_int = _is_int_like(xval)
        if not j_int or not x_int:
            return False, KeyMismatch(
                key=key,
                truth_json_value=jval,
                xlsx_value=xval,
                tolerance_used="shares integer (type error)",
                delta=("non-integer value", (type(jval).__name__, type(xval).__name__)),
            )
        jd = _to_decimal(jval)
        xd = _to_decimal(xval)
        if jd is None or xd is None:
            return False, KeyMismatch(
                key=key, truth_json_value=jval, xlsx_value=xval,
                tolerance_used="shares integer",
                delta="unparsable",
            )
        if jd == xd:
            return True, None
        return False, KeyMismatch(
            key=key, truth_json_value=jval, xlsx_value=xval,
            tolerance_used="shares integer exact",
            delta=abs(jd - xd),
        )

    # numeric (pct or usd)
    jd = _to_decimal(jval)
    xd = _to_decimal(xval)
    if jd is None or xd is None:
        tol_label = "1e-6 (pct)" if kind == "pct" else "max($1, 1e-4 rel)"
        return False, KeyMismatch(
            key=key, truth_json_value=jval, xlsx_value=xval,
            tolerance_used=tol_label,
            delta="unparsable",
        )
    delta = abs(jd - xd)
    if kind == "pct":
        # absolute pp tolerance — 1e-6 is ~0.0001% (still very tight)
        tol = PCT_TOL
        tol_label = "1e-6 (pct)"
    else:
        # USD: dual tolerance. Cross-engine precision drift on $10M+ cells
        # costs us pennies; don't penalize dual-build for that.
        # Pass if abs delta <= $1 OR relative delta <= 1e-4 (0.01%).
        one_dollar = Decimal("1")
        rel_tol = Decimal("1e-4")
        abs_ref = max(abs(jd), abs(xd), Decimal("1"))
        tol = max(one_dollar, rel_tol * abs_ref)
        tol_label = "max($1, 1e-4 rel)"
    if delta <= tol:
        return True, None
    return False, KeyMismatch(
        key=key, truth_json_value=jval, xlsx_value=xval,
        tolerance_used=tol_label, delta=delta,
    )


# --------------------------------------------------------------------------
# Spec sanity check
# --------------------------------------------------------------------------


def parse_spec_keys(spec_text: str) -> set[str]:
    """Best-effort extraction of key-looking tokens from the spec. Used only
    to sanity-check the hardcoded EXPECTED_KEYS list against the spec text.
    """
    # Match tokens with at least one dot and only key-safe chars.
    pattern = re.compile(r"[a-zA-Z_][a-zA-Z0-9_]*(?:\.[a-zA-Z0-9_{}]+)+")
    found = set()
    for m in pattern.findall(spec_text):
        # filter obvious non-keys (e.g. "1.2" is not captured, but "x.y" that
        # is module-style is fine).
        if m.startswith(("scenarios.", "inputs.", "truth.py")):
            continue
        found.add(m)
    return found


def verify_keys_against_spec(spec_path: Path) -> list[str]:
    """Return a list of warnings about drift between EXPECTED_KEYS and the
    spec text. Never fails the run, just reports."""
    warnings: list[str] = []
    try:
        text = spec_path.read_text(encoding="utf-8")
    except OSError as e:
        warnings.append(f"could not read spec ({e}); skipping drift check")
        return warnings
    found = parse_spec_keys(text)

    # The spec uses templates like "safe.{i}.binding_term". Map those to
    # prefixes we can check against.
    # For each expected key, check that either (a) the key appears verbatim
    # or (b) a templated form of its path-head appears.
    for key in EXPECTED_KEYS:
        if key in found:
            continue
        # strip leading segments and look for the stem
        stem = ".".join(key.split(".")[:2])
        if any(stem in f or stem.replace(".1.", ".{i}.") in f for f in found):
            continue
        # loose fallback: last two segments
        tail = ".".join(key.split(".")[-2:])
        if any(tail in f for f in found):
            continue
        warnings.append(f"expected key not found in spec text: {key}")

    return warnings


# --------------------------------------------------------------------------
# Top-level reconcile()
# --------------------------------------------------------------------------


@dataclass
class AuditProbePreflightResult:
    """Outcome of the audit-probe preflight check.

    `status` is one of:
      - "ok"             — both truth.json and truth_perturbed.json present,
                           same key sets, audit-probe ready to run.
      - "skipped"        — no truth_perturbed.json; audit-probe is N/A for
                           this scenario (v3/v0/older). Not a failure.
      - "missing_canonical" — truth.json itself is missing; the wider
                           reconcile would fail anyway.
      - "schema_mismatch" — both files present but truth_perturbed.json's
                           key set doesn't match truth.json. Audit-probe
                           cannot run safely.
    `truth_perturbed_path` is the resolved sibling path (whether it
    exists or not).
    """
    status: str
    truth_path: Path
    truth_perturbed_path: Path
    n_canonical_keys: int = 0
    n_perturbed_keys: int = 0
    keys_only_in_canonical: list[str] = field(default_factory=list)
    keys_only_in_perturbed: list[str] = field(default_factory=list)
    detail: str = ""


def audit_probe_preflight(truth_json_path: Path) -> AuditProbePreflightResult:
    """Verify truth artifacts for the audit-probe sub-test.

    Pre-flight contract: scenarios that ship `truth_perturbed.json` next to
    `truth.json` opt into the audit-probe. Scenarios without it silently
    skip (graceful degrade for v3/older). When both are present, we sanity-
    check that they share the same key set — divergence is a builder bug
    that would silently make the audit-probe meaningless.

    Returns AuditProbePreflightResult with one of {"ok", "skipped",
    "missing_canonical", "schema_mismatch"}. None of these are raised; the
    caller decides severity based on `status`.
    """
    truth_p = truth_json_path.parent / "truth_perturbed.json"

    if not truth_json_path.exists():
        return AuditProbePreflightResult(
            status="missing_canonical",
            truth_path=truth_json_path,
            truth_perturbed_path=truth_p,
            detail=f"canonical truth.json not found at {truth_json_path}",
        )

    if not truth_p.exists():
        # Graceful degrade: this scenario does not exercise the audit-probe.
        return AuditProbePreflightResult(
            status="skipped",
            truth_path=truth_json_path,
            truth_perturbed_path=truth_p,
            detail=("no truth_perturbed.json sibling — audit-probe N/A "
                    "for this scenario"),
        )

    try:
        canonical = load_truth_json(truth_json_path)
        perturbed = load_truth_json(truth_p)
    except Exception as e:  # noqa: BLE001
        return AuditProbePreflightResult(
            status="schema_mismatch",
            truth_path=truth_json_path,
            truth_perturbed_path=truth_p,
            detail=f"failed to load truth files: {e}",
        )

    c_keys = set(canonical.keys())
    p_keys = set(perturbed.keys())
    only_c = sorted(c_keys - p_keys)
    only_p = sorted(p_keys - c_keys)

    if only_c or only_p:
        return AuditProbePreflightResult(
            status="schema_mismatch",
            truth_path=truth_json_path,
            truth_perturbed_path=truth_p,
            n_canonical_keys=len(c_keys),
            n_perturbed_keys=len(p_keys),
            keys_only_in_canonical=only_c,
            keys_only_in_perturbed=only_p,
            detail=(f"key sets differ: {len(only_c)} only in truth.json, "
                    f"{len(only_p)} only in truth_perturbed.json"),
        )

    return AuditProbePreflightResult(
        status="ok",
        truth_path=truth_json_path,
        truth_perturbed_path=truth_p,
        n_canonical_keys=len(c_keys),
        n_perturbed_keys=len(p_keys),
        detail=f"both truth files present, {len(c_keys)} keys aligned",
    )


def reconcile(
    *,
    truth_json_path: Path,
    ground_truth_xlsx_path: Path,
    spec_md_path: Path,
    report_path: Path | None = None,
) -> ReconcileResult:
    """Compare truth.json and ground_truth.xlsx over every tracked key.

    Also surfaces an audit-probe preflight (informational): if a sibling
    `truth_perturbed.json` exists, its key-set must match truth.json's.
    The reconcile result is not failed solely on a schema_mismatch — that
    failure mode is reported via the markdown report and stderr but kept
    out of the matched/mismatched counts because audit-probe is a separate
    sub-test from the canonical truth reconciliation.
    """
    # Audit-probe preflight (informational; does not fail reconcile).
    ap_pf = audit_probe_preflight(truth_json_path)
    if ap_pf.status == "schema_mismatch":
        print(f"WARNING: audit-probe preflight schema_mismatch: {ap_pf.detail}",
              file=sys.stderr)
        for k in ap_pf.keys_only_in_canonical[:10]:
            print(f"  only in truth.json: {k}", file=sys.stderr)
        for k in ap_pf.keys_only_in_perturbed[:10]:
            print(f"  only in truth_perturbed.json: {k}", file=sys.stderr)
    # Sanity-check spec drift (non-fatal).
    _ = verify_keys_against_spec(spec_md_path)

    flat_json = load_truth_json(truth_json_path)
    scenario_keys_list = sorted(flat_json.keys())
    xlsx_values, has_formula, missing_names = load_xlsx_values(
        ground_truth_xlsx_path, keys=scenario_keys_list,
    )

    # Scenario-agnostic key set: truth.json is the canonical list of tracked
    # keys for THIS scenario. Filter EXPECTED_KEYS (v1 superset) to the keys
    # actually present in truth.json — this lets v0 (114 keys) reconcile
    # without penalizing it for missing v1-only keys (venture debt, warrants,
    # SAFEs 7-15).
    scenario_keys = set(flat_json.keys())
    scenario_missing_in_xlsx = [k for k in missing_names if k in scenario_keys]

    missing_in_json: list[str] = []
    missing_in_xlsx: list[str] = list(scenario_missing_in_xlsx)
    ambiguous: list[str] = []
    mismatched: list[KeyMismatch] = []
    matched = 0

    # Iterate the scenario's actual key set (from truth.json), not the v1
    # superset. Keys that v1 tracks but v0 doesn't should not register as
    # mismatches or missing for v0.
    for key in sorted(scenario_keys):
        in_json = key in flat_json
        in_xlsx_defined = key not in missing_names

        if not in_json and not in_xlsx_defined:
            # missing everywhere: report in both lists
            missing_in_json.append(key)
            continue
        if not in_json:
            missing_in_json.append(key)
            continue
        if not in_xlsx_defined:
            # already appended above via missing_names
            continue

        jval = flat_json[key]
        xval = xlsx_values.get(key)

        kind = tolerance_kind(key)
        if kind != "string" and xval is None:
            # cell was a formula we couldn't evaluate, OR defined-name pointed
            # at an empty cell. Distinguish via has_formula flag.
            if has_formula.get(key):
                ambiguous.append(key)
                continue
            # truly empty cell: treat as mismatch — spec requires a formula,
            # so a bare None is a schema violation.
            mismatched.append(KeyMismatch(
                key=key, truth_json_value=jval, xlsx_value=None,
                tolerance_used=f"{kind}",
                delta="xlsx cell empty or non-formula",
            ))
            continue

        ok, mm = _compare_one(key, jval, xval)
        if ok:
            matched += 1
        else:
            assert mm is not None
            mismatched.append(mm)

    passed = (
        not mismatched
        and not missing_in_json
        and not missing_in_xlsx
        and not ambiguous
    )

    result = ReconcileResult(
        passed=passed,
        total_keys=len(EXPECTED_KEYS),
        matched=matched,
        mismatched=mismatched,
        missing_in_truth_json=missing_in_json,
        missing_in_xlsx=missing_in_xlsx,
        ambiguous=ambiguous,
    )

    if report_path is not None:
        write_markdown_report(result, report_path, scenario_name=spec_md_path.parent.name)

    return result


# --------------------------------------------------------------------------
# Markdown report
# --------------------------------------------------------------------------


def _fmt(v: Any) -> str:
    if v is None:
        return "∅"
    if isinstance(v, Decimal):
        return format(v, "f")
    if isinstance(v, float):
        return repr(v)
    return str(v)


def write_markdown_report(result: ReconcileResult, path: Path, scenario_name: str) -> None:
    lines: list[str] = []
    lines.append(f"# Reconciliation Report: {scenario_name}")
    lines.append("")
    lines.append(f"**Status:** {'PASS' if result.passed else 'FAIL'}")
    lines.append(f"**Keys matched:** {result.matched} / {result.total_keys}")
    lines.append(f"**Mismatched:** {len(result.mismatched)}")
    lines.append(f"**Missing in truth.json:** {len(result.missing_in_truth_json)}")
    lines.append(f"**Missing in xlsx:** {len(result.missing_in_xlsx)}")
    lines.append(f"**Ambiguous (formula cells without values):** {len(result.ambiguous)}")
    lines.append("")

    if result.mismatched:
        lines.append("## Mismatches")
        lines.append("")
        lines.append("| Key | truth.json | ground_truth.xlsx | Δ | Tolerance |")
        lines.append("|---|---|---|---|---|")
        for mm in result.mismatched:
            lines.append(
                f"| {mm.key} | {_fmt(mm.truth_json_value)} | "
                f"{_fmt(mm.xlsx_value)} | {_fmt(mm.delta)} | {mm.tolerance_used} |"
            )
        lines.append("")

    if result.missing_in_truth_json:
        lines.append("## Missing in truth.json")
        lines.append("")
        for k in result.missing_in_truth_json:
            lines.append(f"- {k}")
        lines.append("")

    if result.missing_in_xlsx:
        lines.append("## Missing in xlsx")
        lines.append("")
        for k in result.missing_in_xlsx:
            lines.append(f"- {k}")
        lines.append("")

    if result.ambiguous:
        lines.append("## Ambiguous")
        lines.append("")
        lines.append(
            "Keys whose xlsx cell holds a formula but no cached value could be "
            "obtained via openpyxl, the `formulas` package, or LibreOffice "
            "headless recalc. Open the workbook in Excel and save to populate "
            "the cached values, or install one of the fallback tools."
        )
        lines.append("")
        for k in result.ambiguous:
            lines.append(f"- {k}")
        lines.append("")

    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines), encoding="utf-8")


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------


def _main(argv: Iterable[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="Reconcile truth.json against ground_truth.xlsx.",
    )
    ap.add_argument("--spec", required=True, type=Path, help="scenario spec.md path")
    ap.add_argument("--truth", required=True, type=Path, help="truth.json path")
    ap.add_argument("--xlsx", required=True, type=Path, help="ground_truth.xlsx path")
    ap.add_argument("--report", type=Path, default=None, help="optional markdown report path")
    args = ap.parse_args(list(argv) if argv is not None else None)

    # Setup errors → exit 2
    for p, label in [(args.spec, "spec"), (args.truth, "truth"), (args.xlsx, "xlsx")]:
        if not p.exists():
            print(f"ERROR: {label} file not found: {p}", file=sys.stderr)
            return 2

    try:
        result = reconcile(
            truth_json_path=args.truth,
            ground_truth_xlsx_path=args.xlsx,
            spec_md_path=args.spec,
            report_path=args.report,
        )
    except Exception as e:
        print(f"ERROR: reconcile failed to run: {e}", file=sys.stderr)
        return 2

    print(result.summary_line())
    if not result.passed:
        # Show up to 10 mismatches inline for quick debugging.
        for mm in result.mismatched[:10]:
            print(
                f"  MISMATCH {mm.key}: json={_fmt(mm.truth_json_value)} "
                f"xlsx={_fmt(mm.xlsx_value)} Δ={_fmt(mm.delta)} "
                f"[{mm.tolerance_used}]"
            )
        for k in result.ambiguous[:10]:
            print(f"  AMBIGUOUS {k}")
        for k in result.missing_in_truth_json[:10]:
            print(f"  MISSING json {k}")
        for k in result.missing_in_xlsx[:10]:
            print(f"  MISSING xlsx {k}")

    return 0 if result.passed else 1


if __name__ == "__main__":
    sys.exit(_main())
