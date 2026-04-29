"""
Layout-agnostic value extractor for the tetra-v-shortcut benchmark.

Given an arbitrary .xlsx produced by a spreadsheet-generation API (Tetra,
Shortcut, or otherwise), extract the values for every §11 tracked key of the
Nyxlight scenario. Falls back through a cascade of strategies — fastest and
most-reliable first:

  1) Workbook-level defined names that match §11→underscore translation.
  2) Fuzzy tab + keyword row/column match across common layouts (Pro-Forma,
     Waterfall, Sensitivity, Return-Solver, Assumptions).
  3) Compute-derived values (ownership from shares, etc.) when primary cells
     aren't directly present.

The extractor DOES NOT read truth.json, use an LLM, or hit the network. It is
a pure openpyxl + stdlib re/fuzzy module.

Public API:
    extract(xlsx_path=..., spec_md_path=None) -> ExtractResult
    perturb_and_recalc(workbook_path, perturbations, output_path) -> Path

Audit-probe support (v4+)
-------------------------
`perturb_and_recalc` is the input-knob mutator used by
`harness.score.score_audit_probe`. It writes new values into the engine's
workbook for a small set of named input cells (e.g.
`RoundTerms_primary_raise_usd`), invokes LibreOffice headless to recalc all
formulas, and returns the path to the recalculated workbook. The audit-
probe sub-test then re-runs `extract()` over the recalculated workbook and
compares the 167 keys to `truth_perturbed.json`. Workbooks with live
formulas score ~1.0 on the audit-probe; value-dump workbooks score ~0 on
all dependent cells. See contracts.PerturbationSpec / AuditProbeResult.
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
from dataclasses import dataclass, field, asdict
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Callable, Iterable

try:
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    load_workbook = None  # type: ignore


def _maybe_libreoffice_recalc(xlsx_path: Path) -> Path:
    """Best-effort: round-trip the xlsx through LibreOffice headless to force
    formula evaluation. Returns the path to use for reading. On any failure,
    returns the original path unchanged.

    Needed because some API outputs (codex notably) emit pure-formula sheets
    with no cached values — openpyxl can't evaluate, we'd see None everywhere.
    """
    import shutil
    import subprocess
    import tempfile

    binary = shutil.which("libreoffice") or shutil.which("soffice")
    if binary is None:
        return xlsx_path

    try:
        out_dir = Path(tempfile.mkdtemp(prefix="extract_recalc_"))
        subprocess.run(
            [binary, "--headless", "--calc", "--convert-to", "xlsx",
             "--outdir", str(out_dir), str(xlsx_path)],
            check=True, capture_output=True, timeout=120,
        )
        recalc = out_dir / xlsx_path.name
        if recalc.exists() and recalc.stat().st_size > 0:
            return recalc
    except Exception:  # noqa: BLE001
        pass
    return xlsx_path


def _libreoffice_recalc_to(xlsx_path: Path, out_dir: Path) -> Path | None:
    """Run LibreOffice headless on `xlsx_path`, writing the recalculated
    workbook into `out_dir`. Returns the recalculated file path or None if
    the binary isn't available / the conversion failed.

    Distinct from `_maybe_libreoffice_recalc` because the audit-probe needs
    the recalc artifact at a known location (so we can re-extract from it),
    not a transient temp dir we silently swap behind the caller.
    """
    import shutil
    import subprocess

    binary = shutil.which("libreoffice") or shutil.which("soffice")
    if binary is None:
        return None
    out_dir.mkdir(parents=True, exist_ok=True)
    try:
        subprocess.run(
            [binary, "--headless", "--calc", "--convert-to", "xlsx",
             "--outdir", str(out_dir), str(xlsx_path)],
            check=True, capture_output=True, timeout=180,
        )
    except Exception:  # noqa: BLE001
        return None
    cand = out_dir / xlsx_path.name
    if cand.exists() and cand.stat().st_size > 0:
        return cand
    return None


def _find_perturbation_target(
    wb,
    perturbation_key: str,
) -> tuple[str, str] | None:
    """Locate the cell to overwrite for a given perturbation key.

    Strategy (matches the extractor's general philosophy: defined-name fast
    path, fall back to fuzzy label-region search):

      1) Direct workbook-level defined name on `key.replace(".", "_")`.
         This is the contract per spec §14 — engines that follow the
         convention land here.
      2) Case-variant defined-name lookup (handled inside
         `_resolve_defined_name`).
      3) Label-region scan via `_scalar_label_search`. The key's last
         dot-segment is tokenized (e.g. `RoundTerms.primary_raise_usd` →
         tokens {"primary", "raise", "usd"}) and matched against label cells
         on Assumptions / Inputs / RoundTerms tabs. Returns the (sheet, A1)
         of the first numeric cell in the matching row.

    Returns (sheet, A1-coord) or None if no target is found.
    """
    named = perturbation_key.replace(".", "_")
    val, sheet, coord = _resolve_defined_name(wb, named)
    if sheet is not None and coord is not None:
        # Strip any range syntax — defined names typically point at a single
        # cell, but could be a range; we want the top-left.
        first = str(coord).split(":")[0].replace("$", "")
        return sheet, first

    # Label-region fallback. Tokenize the last dot-segment, drop semantic
    # noise tokens, and search.
    last = perturbation_key.split(".")[-1]
    raw_tokens = [t for t in re.split(r"[^a-z0-9]+", last.lower()) if t]
    # Treat each non-trivial token as a single-element OR group → AND
    # across all groups (i.e. all tokens must appear in the label).
    drop = {"usd", "value", "amount", "input"}  # too generic
    tok_groups: list[tuple[str, ...]] = [
        (t,) for t in raw_tokens if t not in drop and len(t) > 1
    ]
    if not tok_groups:
        return None

    # Bias toward input-style sheets first.
    found = _scalar_label_search(
        wb,
        token_groups=tok_groups,
        prefer_sheets=(
            "assumption", "assumptions", "inputs", "input",
            "roundterms", "round terms", "round_terms",
            "pro-forma", "proforma", "pro forma", "pro_forma",
        ),
    )
    if found is None:
        return None
    sheet, row, col, _label, _val = found
    return sheet, _cell_a1(row, col)


def perturb_and_recalc(
    workbook_path: "Path | str",
    perturbations: dict[str, float],
    output_path: "Path | str",
) -> Path:
    """Apply input perturbations to `workbook_path` and recalculate.

    Audit-probe primitive. Used by `harness.score.score_audit_probe`.

    Steps:
      1) Copy the source workbook to `output_path` (we never mutate the
         engine's original output xlsx).
      2) For each (key, value) in `perturbations`, locate the named cell
         (defined-name fast path; label-region fallback) and overwrite its
         value. Defined-name targets are the contract; the label fallback
         exists because some engines forget to define a name on input cells
         even when they do for output cells.
      3) Run LibreOffice headless to force a full formula recalculation.
         The recalculated copy lands at `output_path` (overwritten in
         place, alongside any LibreOffice intermediate file).
      4) Return the path to the recalculated workbook.

    Raises:
      RuntimeError if openpyxl or LibreOffice is unavailable (these are
      the same dependencies extract() needs; failure means the audit-probe
      cannot run on this host).

    Quirks:
      - LibreOffice's --convert-to xlsx round-trip can change cell number
        formats and may strip volatile features. For audit-probe purposes
        we only care about the cached values of tracked-key cells, which
        survive cleanly.
      - Defined names that point at multi-cell ranges write only the
        top-left cell. This is the right behavior for input-knob cells
        (always single cells per spec §14).
      - openpyxl's `wb.save()` over a file that originally contained a
        cached calc chain leaves stale `<calcChain>` entries; LibreOffice
        rebuilds them during recalc, so this is not a concern downstream.
    """
    import shutil

    if load_workbook is None:
        raise RuntimeError("openpyxl not installed; perturb_and_recalc unavailable")

    src = Path(workbook_path)
    dst = Path(output_path)
    if not src.exists():
        raise FileNotFoundError(str(src))
    dst.parent.mkdir(parents=True, exist_ok=True)

    # Copy original → dst. We mutate dst in place.
    shutil.copyfile(src, dst)

    # Open in formula mode (data_only=False) so we don't strip formulas
    # while writing input cells.
    wb = load_workbook(str(dst), data_only=False)

    applied: list[str] = []
    unapplied: list[str] = []

    for key, value in perturbations.items():
        target = _find_perturbation_target(wb, key)
        if target is None:
            unapplied.append(key)
            continue
        sheet_name, a1 = target
        try:
            ws = wb[sheet_name]
            ws[a1].value = value
            applied.append(key)
        except Exception:  # noqa: BLE001
            unapplied.append(key)

    # Persist the input-mutated workbook.
    wb.save(str(dst))

    # Force a full recalc via LibreOffice. We write the recalc into a
    # sibling temp dir to avoid LibreOffice clobbering the file we just
    # saved (it locks the source during conversion).
    import tempfile

    with tempfile.TemporaryDirectory(prefix="audit_probe_lo_") as td:
        recalc = _libreoffice_recalc_to(dst, Path(td))
        if recalc is None:
            # No LibreOffice available — return the value-mutated workbook
            # unchanged. Downstream extract() will then fall back to its
            # own _maybe_libreoffice_recalc (also a no-op), and any cells
            # that needed formula re-evaluation will read as None / stale.
            # The audit-probe pass_pct will collapse, which is the correct
            # signal to the user: "we couldn't probe, treat as inconclusive."
            return dst
        # Move the recalculated file over dst.
        shutil.copyfile(recalc, dst)

    # Stash the apply/unapply lists on the returned path's parent for the
    # caller to retrieve via a sidecar JSON. Simpler than threading another
    # return value through; the audit-probe scorer reads this file.
    sidecar = dst.with_suffix(dst.suffix + ".perturb.json")
    try:
        sidecar.write_text(json.dumps({
            "applied": applied,
            "unapplied": unapplied,
            "perturbations": {k: float(v) for k, v in perturbations.items()},
            "source": str(src),
        }, indent=2))
    except Exception:  # noqa: BLE001
        pass

    return dst


# ---------------------------------------------------------------------------
# Result types
# ---------------------------------------------------------------------------


@dataclass
class CellRef:
    sheet: str
    cell: str           # A1 notation; may be "A1+A2+..." for computed keys
    raw_value: Any
    how_matched: str    # human-readable rule that fired


@dataclass
class ExtractResult:
    values: dict[str, Any]
    provenance: dict[str, CellRef]
    missing: list[str]
    ambiguous: list[tuple[str, list[CellRef]]]

    def to_dict(self) -> dict[str, Any]:
        return {
            "values": {k: _json_safe(v) for k, v in self.values.items()},
            "provenance": {
                k: {
                    "sheet": cr.sheet,
                    "cell": cr.cell,
                    "raw_value": _json_safe(cr.raw_value),
                    "how_matched": cr.how_matched,
                }
                for k, cr in self.provenance.items()
            },
            "missing": list(self.missing),
            "ambiguous": [
                (
                    k,
                    [
                        {
                            "sheet": cr.sheet,
                            "cell": cr.cell,
                            "raw_value": _json_safe(cr.raw_value),
                            "how_matched": cr.how_matched,
                        }
                        for cr in cands
                    ],
                )
                for k, cands in self.ambiguous
            ],
        }


def _json_safe(v: Any) -> Any:
    if isinstance(v, Decimal):
        return float(v)
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return None
    return v


# ---------------------------------------------------------------------------
# Canonical key list (mirrors score.iter_tracked_keys)
# ---------------------------------------------------------------------------


def _sensitivity_esop_token(esop: float) -> str:
    if esop == 0.125:
        return "esop_12p5"
    pct = int(round(esop * 100))
    return f"esop_{pct}"


def _infer_kind_from_key(key: str) -> str:
    """Map a flat key name to its tolerance kind. Mirrors score._infer_kind_from_key
    so extract() can derive a scenario-specific tracked-key list from truth.json
    without having to import score.py (which imports extract.py).

    Rules: any last-segment ending in `_pct` → pct; `.new_ratio` → ratio
    (unit-less, MUST NOT be /100-normalized); `.shares*` → shares;
    `.election` → election; `.binding*` → binding; usd heuristics
    (`.dollars`, `_usd`, `usd_to_*`, conversion_price, accreted_principal) → usd."""
    lower = key.lower()
    last_seg = lower.split(".")[-1]
    if lower.endswith(".election"):
        return "election"
    if "binding_term" in lower or lower.endswith(".binding"):
        return "binding"
    # Ratios are unit-less and bounded ≥1.0 for triggered AD classes — they
    # are NOT percentages. Distinct kind so the defined-name dispatcher can
    # skip /100-normalization.
    if lower.endswith(".new_ratio"):
        return "ratio"
    # Any last-segment ending in `_pct` (incl. `.pct`, `.our_pct`,
    # `.required_pct`).
    if last_seg.endswith("_pct") or last_seg == "pct":
        return "pct"
    if "shares" in last_seg:
        return "shares"
    if (lower.endswith(".dollars") or lower.endswith("_usd")
            or "dollars_at" in lower or lower.endswith(".total_distributed")
            or "conversion_price" in lower or "accreted_principal" in lower
            or "usd_to_" in lower):
        return "usd"
    return "string"


def _tracked_keys_from_truth(truth_json_path: Path) -> list[dict[str, Any]] | None:
    """Read a truth.json file and return a list of {key, kind} dicts. Returns
    None if the file can't be read/parsed; caller should fall back to the
    hardcoded v1 superset (`iter_tracked_keys()`)."""
    try:
        with open(truth_json_path) as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return None
        return [
            {"key": k, "kind": _infer_kind_from_key(k)}
            for k in sorted(data.keys())
        ]
    except Exception:
        return None


def iter_tracked_keys() -> list[dict[str, Any]]:
    keys: list[dict[str, Any]] = []

    ownership_parts = [
        "founders", "employees", "options_issued", "options_unissued",
        "seed", "series_a", "safes", "venture_debt", "warrants",
        "series_b_new", "our_fund",
    ]
    for p in ownership_parts:
        keys.append({"key": f"ownership.{p}.pct", "kind": "pct"})

    for i in range(1, 16):
        keys.append({"key": f"safe.{i}.binding_term", "kind": "binding"})
        keys.append({"key": f"safe.{i}.conversion_price_usd", "kind": "usd"})
        keys.append({"key": f"safe.{i}.shares_issued", "kind": "shares"})

    keys.append({"key": "venture_debt.accreted_principal_usd", "kind": "usd"})
    keys.append({"key": "venture_debt.binding_term", "kind": "binding"})
    keys.append({"key": "venture_debt.conversion_price_usd", "kind": "usd"})
    keys.append({"key": "venture_debt.shares_issued", "kind": "shares"})
    keys.append({"key": "venture_debt.vdw_shares", "kind": "shares"})

    for series in ("seed", "series_a"):
        keys.append({"key": f"antidilution.{series}.new_conversion_price_usd", "kind": "usd"})
        keys.append({"key": f"antidilution.{series}.new_ratio", "kind": "ratio"})

    waterfall_dollar_classes = [
        "venture_debt_as_b", "series_b", "series_a", "seed", "safes",
        "common", "options_issued", "warrants", "our_fund",
    ]
    waterfall_elections = [("series_b",), ("series_a",), ("seed",)]
    for exit_m in (100, 250, 500, 1000, 2000):
        for cls in waterfall_dollar_classes:
            keys.append({"key": f"waterfall.{exit_m}M.{cls}.dollars", "kind": "usd"})
        for (cls,) in waterfall_elections:
            keys.append({"key": f"waterfall.{exit_m}M.{cls}.election", "kind": "election"})
        keys.append({"key": f"waterfall.{exit_m}M.total_distributed", "kind": "usd_abs_1"})

    for esop in (0.10, 0.125, 0.15):
        tok = _sensitivity_esop_token(esop)
        for pre in (60, 80, 120):
            keys.append({"key": f"sensitivity.{tok}.pre_{pre}M.our_pct", "kind": "pct"})
            keys.append({"key": f"sensitivity.{tok}.pre_{pre}M.dollars_at_500M_exit", "kind": "usd"})

    for mult in ("3x", "4x"):
        for exit_m in (100, 250, 500, 1000, 2000):
            keys.append({"key": f"solver.{mult}.exit_{exit_m}M.required_pct", "kind": "pct"})

    return keys


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _norm(s: Any) -> str:
    """Lowercase, strip, squash non-alnum to single underscores."""
    if s is None:
        return ""
    t = str(s).lower()
    t = re.sub(r"[^a-z0-9%$]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def _all_tokens(s: Any) -> set[str]:
    return set(_norm(s).split()) if s is not None else set()


def _to_decimal(x: Any) -> Decimal | None:
    if x is None:
        return None
    if isinstance(x, bool):
        return None
    if isinstance(x, Decimal):
        return x
    if isinstance(x, int):
        return Decimal(x)
    if isinstance(x, float):
        if math.isnan(x) or math.isinf(x):
            return None
        return Decimal(str(x))
    if isinstance(x, str):
        s = x.strip().replace(",", "").replace("$", "")
        if s.endswith("%"):
            try:
                return Decimal(s[:-1].strip()) / Decimal("100")
            except (InvalidOperation, ValueError):
                return None
        # Strip trailing M/B/K suffix
        m = re.match(r"^([-+]?[0-9]*\.?[0-9]+)([kmb])$", s.lower())
        if m:
            num = Decimal(m.group(1))
            suf = m.group(2)
            mul = {"k": Decimal("1000"), "m": Decimal("1000000"), "b": Decimal("1000000000")}[suf]
            return num * mul
        if not s:
            return None
        try:
            return Decimal(s)
        except (InvalidOperation, ValueError):
            return None
    return None


def _cell_a1(row: int, col: int) -> str:
    return f"{get_column_letter(col)}{row}"


def _fuzzy_sheet_match(wb, needles: tuple[str, ...]) -> list[str]:
    """Return sheet names whose normalized name contains any needle (normalized)."""
    out: list[str] = []
    for sn in wb.sheetnames:
        n = _norm(sn)
        for nd in needles:
            if _norm(nd) in n:
                out.append(sn)
                break
    return out


def _iter_nonempty_cells(ws, max_rows: int = 2000, max_cols: int = 100):
    rmax = min(ws.max_row or 0, max_rows)
    cmax = min(ws.max_column or 0, max_cols)
    for r in range(1, rmax + 1):
        for c in range(1, cmax + 1):
            v = ws.cell(r, c).value
            if v is not None and (not isinstance(v, str) or v.strip()):
                yield r, c, v


# ---------------------------------------------------------------------------
# Defined-name fast path
# ---------------------------------------------------------------------------


def _resolve_defined_name(wb, named: str) -> tuple[Any, str | None, str | None]:
    """
    Return (value, sheet, coord). All three None on failure.
    """
    dn = None
    try:
        dn = wb.defined_names.get(named)
    except Exception:
        dn = None
    if dn is None:
        try:
            for name in wb.defined_names:
                if name.lower() == named.lower():
                    dn = wb.defined_names[name]
                    break
        except Exception:
            dn = None
    if dn is None:
        return None, None, None
    try:
        dests = list(dn.destinations)
    except Exception:
        return None, None, None
    if not dests:
        return None, None, None
    sheet_name, coord = dests[0]
    try:
        ws = wb[sheet_name]
    except Exception:
        return None, None, None
    try:
        cell_range = ws[coord]
    except Exception:
        return None, None, None
    if isinstance(cell_range, tuple):
        try:
            val = cell_range[0][0].value
            c = cell_range[0][0].coordinate
            return val, sheet_name, c
        except Exception:
            try:
                val = cell_range[0].value
                c = cell_range[0].coordinate
                return val, sheet_name, c
            except Exception:
                return None, None, None
    return getattr(cell_range, "value", None), sheet_name, getattr(cell_range, "coordinate", coord)


# ---------------------------------------------------------------------------
# Row-scanning helpers used by most strategies
# ---------------------------------------------------------------------------


LABEL_MAX_COL = 8        # look for labels in first 8 columns
MAX_SCAN_ROWS = 500


def _find_label_rows(
    ws,
    match_fn: Callable[[str], bool],
    max_rows: int = MAX_SCAN_ROWS,
    label_cols: int = LABEL_MAX_COL,
) -> list[tuple[int, int, str]]:
    """
    Scan label columns for any cell whose stringified value passes match_fn.
    Return list of (row, col, label_text).
    """
    out: list[tuple[int, int, str]] = []
    rmax = min(ws.max_row or 0, max_rows)
    cmax = min(ws.max_column or 0, label_cols)
    for r in range(1, rmax + 1):
        for c in range(1, cmax + 1):
            v = ws.cell(r, c).value
            if v is None:
                continue
            if not isinstance(v, str):
                continue
            if match_fn(v):
                out.append((r, c, v))
    return out


def _first_numeric_in_row(
    ws,
    row: int,
    start_col: int,
    max_col: int = 50,
) -> tuple[int, Any] | None:
    cmax = min(ws.max_column or 0, max_col)
    for c in range(start_col, cmax + 1):
        v = ws.cell(row, c).value
        if isinstance(v, (int, float, Decimal)) and not isinstance(v, bool):
            return c, v
    return None


def _numeric_cells_in_row(
    ws,
    row: int,
    start_col: int = 1,
    max_col: int = 100,
) -> list[tuple[int, Any]]:
    cmax = min(ws.max_column or 0, max_col)
    out = []
    for c in range(start_col, cmax + 1):
        v = ws.cell(row, c).value
        if isinstance(v, (int, float, Decimal)) and not isinstance(v, bool):
            out.append((c, v))
    return out


def _numeric_or_coerced_in_row(
    ws,
    row: int,
    start_col: int = 1,
    max_col: int = 100,
) -> list[tuple[int, Decimal]]:
    cmax = min(ws.max_column or 0, max_col)
    out = []
    for c in range(start_col, cmax + 1):
        v = ws.cell(row, c).value
        d = _to_decimal(v)
        if d is not None:
            out.append((c, d))
    return out


# ---------------------------------------------------------------------------
# Ownership extraction (Pro-Forma)
# ---------------------------------------------------------------------------


# Pattern: (category_key, set of required tokens, set of disallowed tokens)
# "Required" means row's label must contain at least one of each group.
OWNERSHIP_LABEL_RULES: list[tuple[str, list[set[str]], set[str]]] = [
    # Match founders/employees carefully. A row like "Common (Founders + RSAs)"
    # is a COMBINED founders+RSAs row — it's legitimate to classify only as
    # `founders` (first rule wins under single-match logic) AND NOT also as
    # employees. Disallow `founders` (the plural, which is what "Founders +
    # RSAs" contains) on the employees rule.
    # Disallow `secondary`/`lead` so a row like "F_LEAD secondary (common
     # from founders)" — present in v4 cap tables — does NOT get rolled into
     # `ownership.founders.pct`. Those shares belong to `f_lead_secondary`.
    ("founders",          [{"founder", "founders"}],                            {"rsa", "employee", "secondary", "lead", "f_lead"}),
    # Employees rule: require an explicit employee/RSA token. The bare
    # "e1"/"e8" tokens used to live here but they collide with Series E
    # investor IDs like "E1_INV"/"E8_INV" — those rows MUST NOT classify as
    # employees. We also exclude "inv" via the disallow set (after _norm,
    # "E1_INV" tokenizes to {e1, inv}) so any future "E1_INV" style label
    # is firmly rejected.
    ("employees",         [{"employee", "employees", "rsa", "rsas"}],
        {"founder", "founders", "option", "options", "grants", "issued",
         "grant", "inv", "investor", "investors"}),
    ("options_issued",    [{"option", "options"}, {"issued"}],                 {"unissued", "reserve", "refresh", "esop"}),
    ("options_unissued",  [{"option", "options", "esop", "pool", "reserve"}, {"unissued", "reserve", "refresh", "esop"}], {"issued"}),
    ("seed",              [{"seed"}],                                          set()),
    ("series_a",          [{"series"}, {"a"}],                                 {"b", "c", "d", "e"}),
    ("series_b",          [{"series"}, {"b"}],                                 {"a", "c", "d", "e", "new"}),
    ("series_c",          [{"series"}, {"c"}],                                 {"a", "b", "d", "e"}),
    ("series_d",          [{"series"}, {"d"}],                                 {"a", "b", "c", "e"}),
    ("safes",             [{"safe", "safes"}],                                 {"conversion", "table", "total"}),
    ("venture_debt",      [{"venture", "debt", "vd"}],                         {"warrant", "vdw", "residual"}),
    ("warrants",          [{"warrant", "warrants", "w1", "w2", "vdw"}],        set()),
    # "new"-round ownership: the current round's new investors. For v0/v1/v2
    # this is B (excludes our fund since truth tracks our_fund separately and
    # series_b_new excludes it). For v3 this is E (INCLUDES our fund since v3's
    # ownership.series_e_new.pct is the roll-up of all Series E, and our_fund
    # is tracked separately as an additional key).
    ("series_b_new",      [{"series"}, {"b"}],                                 {"a", "c", "d", "e", "f", "our", "fund"}),
    ("series_e_new",      [{"series"}, {"e"}],                                 {"a", "b", "c", "d", "f"}),  # include our fund row
    # v4: series_f_new is the new Series F preferred shares (F_LEAD primary +
    # our_fund primary). Excludes the F_LEAD secondary (common from tender)
    # and excludes prior series so a row like "Series F Preferred" lands here.
    ("series_f_new",      [{"series"}, {"f"}],                                 {"a", "b", "c", "d", "e", "secondary"}),
    # our_fund: must mention "our fund" but NOT be a row that explicitly
    # excludes our fund ("ex our fund", "new investors", "external"). These
    # phrases signal "everyone else in this round besides us".
    ("our_fund",          [{"our"}, {"fund"}],                                 {"ex", "external", "new", "excluding"}),
    # Looser fallback for layouts that label our investor without the literal
    # word "fund" (e.g., shortcut a8 emits "Our Shares (pro-rata of Series C)"
    # / "Our % of Series C"). These rows still uniquely identify our_fund's
    # ownership pct on those scenarios. Disallow any token suggesting a
    # different fund or generic-fund-but-not-ours so we don't false-positive
    # on "Founders Fund" / "Hedge Fund" / "Crossover Fund".
    ("our_fund",          [{"our"}, {"shares", "check", "investment", "proceeds", "ownership", "pct", "%"}],
        {"ex", "external", "excluding", "new", "founders", "founder",
         "hedge", "crossover", "investors"}),
]


def _label_ownership_category(label: str) -> str | None:
    """Legacy single-match lookup. Kept for callers that only need one cat."""
    cats = _label_ownership_categories(label)
    return cats[0] if cats else None


def _label_ownership_categories(label: str) -> list[str]:
    """Return ALL categories that match this label. Used so that a row like
    'Our Fund (Series E)' can populate both `our_fund` and `series_e_new` — the
    latter being a roll-up of all Series E investors INCLUDING our fund."""
    tokens = _all_tokens(label)
    if not tokens:
        return []
    out: list[str] = []
    for cat, required_groups, disallowed in OWNERSHIP_LABEL_RULES:
        if disallowed & tokens:
            continue
        if all(group & tokens for group in required_groups):
            if cat not in out:
                out.append(cat)
    return out


def _extract_ownership(wb, provenance: dict, values: dict) -> None:
    """
    Find a Pro-Forma / Cap Table tab. Locate a header row that contains both a
    share column and an FD% column. Then classify each row below into a §11
    ownership category and sum values. Prefer %-column; else compute from
    shares/total.
    """
    sheet_names = _fuzzy_sheet_match(wb, (
        "pro-forma", "proforma", "pro forma", "pro_forma",
        "cap table", "cap_table", "captable",
        "post-series", "post series", "pro-forma cap",
    ))
    if not sheet_names:
        return

    for sn in sheet_names:
        ws = wb[sn]
        if not _try_ownership_on_sheet(ws, sn, provenance, values):
            continue
        # First successful sheet wins
        return


def _find_header_row(
    ws,
    scan_rows: int = 80,
    keywords_required: tuple[tuple[str, ...], ...] = (("stakeholder", "class", "name"),),
    keywords_any: tuple[str, ...] = (),
) -> int | None:
    """
    Returns the 1-based row index of the first row whose cells collectively
    satisfy ALL required keyword tuples. Each tuple is an OR-group.
    """
    rmax = min(ws.max_row or 0, scan_rows)
    for r in range(1, rmax + 1):
        joined = " ".join(
            _norm(ws.cell(r, c).value)
            for c in range(1, min(ws.max_column or 0, 20) + 1)
            if ws.cell(r, c).value is not None
        )
        if not joined:
            continue
        ok = True
        for group in keywords_required:
            if not any(k in joined for k in group):
                ok = False
                break
        if not ok:
            continue
        if keywords_any and not any(k in joined for k in keywords_any):
            continue
        return r
    return None


def _header_columns(ws, header_row: int, max_col: int = 30) -> dict[int, str]:
    cmax = min(ws.max_column or 0, max_col)
    out: dict[int, str] = {}
    for c in range(1, cmax + 1):
        v = ws.cell(header_row, c).value
        if v is not None:
            out[c] = _norm(v)
    return out


def _find_col_by_keywords(
    headers: dict[int, str],
    any_of: tuple[str, ...],
    none_of: tuple[str, ...] = (),
) -> int | None:
    for c, txt in headers.items():
        if any(k in txt for k in any_of) and not any(k in txt for k in none_of):
            return c
    return None


def _try_ownership_on_sheet(ws, sn: str, provenance: dict, values: dict) -> bool:
    # Find a "cap table" style header: has Shares and (Fully Diluted / FD / %) cols.
    # Some sheets have two tables (e.g., SAFE detail then cap table) — we need
    # the one with ownership rows. Try each plausible header row.
    candidates: list[int] = []
    rmax = min(ws.max_row or 0, 200)
    for r in range(1, rmax + 1):
        joined = " ".join(
            _norm(ws.cell(r, c).value)
            for c in range(1, min(ws.max_column or 0, 20) + 1)
            if ws.cell(r, c).value is not None
        )
        if not joined:
            continue
        has_shares = ("shares" in joined or "share count" in joined)
        has_pct = ("fd" in joined or "fully diluted" in joined or "%" in joined
                   or "ownership" in joined or "pct" in joined)
        if has_shares and has_pct:
            candidates.append(r)

    for header_row in candidates:
        headers = _header_columns(ws, header_row, max_col=30)
        shares_col = _find_col_by_keywords(headers, ("shares",), ("strike",))
        pct_col = _find_col_by_keywords(
            headers,
            ("fully diluted", "fd %", "fd%", "ownership", "ownership %", "pct", "%"),
            (),
        )
        if shares_col is None and pct_col is None:
            continue

        # Label column = leftmost non-empty of first 3 cols in header row
        label_col = None
        for c in sorted(headers.keys()):
            if c != shares_col and c != pct_col:
                label_col = c
                break
        if label_col is None:
            label_col = 1

        # Scan rows below the header until blank or "total"
        per_cat_pct: dict[str, list[tuple[Decimal, int, int]]] = {}
        per_cat_shares: dict[str, list[tuple[Decimal, int, int]]] = {}
        total_fd_shares: Decimal | None = None
        first_data_row = header_row + 1
        for r in range(first_data_row, min(ws.max_row or 0, first_data_row + 150) + 1):
            label_val = ws.cell(r, label_col).value
            label = str(label_val) if label_val is not None else ""
            if not label.strip():
                # Allow a few blanks; but if label column empty across 3 cols,
                # consider it a section break — break if we've seen data.
                all_blank = all(
                    ws.cell(r, c).value is None
                    for c in range(1, min(ws.max_column or 0, 10) + 1)
                )
                if all_blank and per_cat_pct:
                    break
                continue
            low = _norm(label)
            if "total" in low and ("fully diluted" in low or "fd" in low or low.startswith("total")):
                # Capture total for fallback ownership computation
                if shares_col is not None:
                    v = _to_decimal(ws.cell(r, shares_col).value)
                    if v is not None:
                        total_fd_shares = v
                # Also stop — we've reached the end
                break
            cats = _label_ownership_categories(label)
            if not cats:
                continue
            for cat in cats:
                if pct_col is not None:
                    v = _to_decimal(ws.cell(r, pct_col).value)
                    if v is not None:
                        per_cat_pct.setdefault(cat, []).append((v, r, pct_col))
                if shares_col is not None:
                    v = _to_decimal(ws.cell(r, shares_col).value)
                    if v is not None:
                        per_cat_shares.setdefault(cat, []).append((v, r, shares_col))

        if not per_cat_pct and not per_cat_shares:
            continue

        # Success on this sheet. Record values.
        for cat, _rules, _disallow in OWNERSHIP_LABEL_RULES:
            key = f"ownership.{cat}.pct"
            if cat in per_cat_pct:
                rows = per_cat_pct[cat]
                total_pct = sum((p[0] for p in rows), Decimal(0))
                # If pct stored as "4.76" (not 0.0476), heuristically rescale.
                total_pct = _normalize_pct(total_pct)
                cells = "+".join(_cell_a1(r, c) for _v, r, c in rows)
                values[key] = float(total_pct)
                provenance[key] = CellRef(
                    sheet=sn,
                    cell=cells,
                    raw_value=float(total_pct),
                    how_matched=f"ownership:{cat} fd%% rows",
                )
            elif cat in per_cat_shares and total_fd_shares and total_fd_shares > 0:
                rows = per_cat_shares[cat]
                total_shares = sum((p[0] for p in rows), Decimal(0))
                pct = total_shares / total_fd_shares
                cells = "+".join(_cell_a1(r, c) for _v, r, c in rows)
                values[key] = float(pct)
                provenance[key] = CellRef(
                    sheet=sn,
                    cell=cells + " / Total FD",
                    raw_value=float(pct),
                    how_matched=f"ownership:{cat} computed shares/total_fd",
                )
        return True
    return False


def _normalize_pct(v: Decimal) -> Decimal:
    """
    Heuristic: ownership percentages are <= 1 (fractions). If a value appears
    to be >1 and <=100, assume it's already in 'percent units' and divide by 100.
    """
    if v is None:
        return v
    if v > Decimal(1) and v <= Decimal(100):
        return v / Decimal(100)
    return v


# ---------------------------------------------------------------------------
# Unit detection (millions vs raw)
# ---------------------------------------------------------------------------
#
# Some API outputs declare share counts and dollars in MILLIONS rather than raw
# numbers (e.g., a founder row showing "4" meaning 4,000,000 shares; a waterfall
# cell showing "300" meaning $300,000,000). When this happens, every downstream
# value emitted from that sheet must be multiplied by 1e6 before scoring or it
# will compare ~6 orders of magnitude too small.
#
# Detection runs once per workbook (per sheet, actually) at the start of
# extract(). Results are stashed on `wb._unit_scales: dict[sheet, dict]`.
# Each sheet's dict has:
#   "dollars_mult": 1 or 1_000_000
#   "shares_mult":  1 or 1_000_000
#   "dollars_reason": str  (which heuristic fired; "" if no scaling)
#   "shares_reason":  str
#
# Scaling is applied at each emit point that produces a `kind=usd|shares` value
# from a specific source sheet — see _scale_dollars_for / _scale_shares_for.

# Substrings that, when found in a header cell or banner/title text, EXPLICITLY
# declare millions units. Match is case-insensitive, against the *normalized*
# (lowercased, punctuation→space) text.
_MILLIONS_DOLLARS_MARKERS = (
    "$ in millions", "in millions $", "dollars in millions", "$m",
    "($m)", "$ m ", "in $m", "values in millions",
    "($ in millions",  # banner like "Post-Series E ($ in millions"
)
_MILLIONS_SHARES_MARKERS = (
    "shares in millions", "shares m", "shares (m)", "as converted m",
    "as-converted m", "as converted (m)", "as-converted (m)",
    "orig shares m", "orig shares (m)", "original shares m",
    "original shares (m)", "ac (m)", "as conv m", "as-conv m",
    "as conv (m)", "as-conv (m)",
)


def _scan_text_for_units(ws, max_rows: int = 30, max_cols: int = 20) -> str:
    """Return concatenated normalized text of the first ~max_rows of a sheet,
    suitable for substring-matching against unit markers."""
    rmax = min(ws.max_row or 0, max_rows)
    cmax = min(ws.max_column or 0, max_cols)
    parts: list[str] = []
    for r in range(1, rmax + 1):
        for c in range(1, cmax + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip():
                parts.append(_norm(v))
    return " | ".join(parts)


def _detect_sheet_units(ws) -> dict[str, Any]:
    """
    Inspect a single worksheet for explicit millions declarations in headers
    and banners. Falls back to a magnitude sanity check on cap-table share
    columns when no explicit text is found.

    Returns: {"dollars_mult", "shares_mult", "dollars_reason", "shares_reason"}.
    """
    info = {
        "dollars_mult": 1,
        "shares_mult": 1,
        "dollars_reason": "",
        "shares_reason": "",
    }

    # 1) Explicit text scan: banners + header rows.
    text = _scan_text_for_units(ws, max_rows=40)
    for marker in _MILLIONS_DOLLARS_MARKERS:
        if marker in text:
            info["dollars_mult"] = 1_000_000
            info["dollars_reason"] = f"explicit:'{marker}'"
            break
    for marker in _MILLIONS_SHARES_MARKERS:
        if marker in text:
            info["shares_mult"] = 1_000_000
            info["shares_reason"] = f"explicit:'{marker}'"
            break

    # 2) Magnitude sanity check (only if shares not already declared).
    # Look at any column whose header EXACTLY normalizes to a generic "shares"-
    # like header. Sample up to 6 numeric values from rows just below. If the
    # median is < 50 (and > 0), shares are almost certainly in millions —
    # founders / employees / preferred-share blocks at raw scale are ALWAYS
    # in the millions of shares.
    #
    # We deliberately require an exact-token match (not substring) — phrases
    # like "Seed preferred (as-converted)" are ROW labels in a Class column,
    # not COLUMN headers, and reading numbers below them crosses table
    # boundaries (false positive on codex-style narrow cap tables).
    SHARES_HEADER_TOKENS = {
        "shares", "as converted", "as-converted", "as conv", "as-conv",
        "orig shares", "original shares", "share count", "shares as conv",
        "shares as converted", "shares as-converted",
    }
    if info["shares_mult"] == 1:
        rmax = min(ws.max_row or 0, 30)
        cmax = min(ws.max_column or 0, 30)
        for hr in range(1, rmax + 1):
            for c in range(1, cmax + 1):
                hv = ws.cell(hr, c).value
                if not isinstance(hv, str):
                    continue
                hnorm = _norm(hv)
                if hnorm not in SHARES_HEADER_TOKENS:
                    continue
                if "price" in hnorm or "pps" in hnorm or "$" in hnorm:
                    continue
                # Stop sampling at the next non-numeric break (a string row
                # below indicates we've crossed into another table).
                samples: list[float] = []
                for rr in range(hr + 1, min(rmax + 30, (ws.max_row or 0)) + 1):
                    vv = ws.cell(rr, c).value
                    if vv is None:
                        continue
                    if isinstance(vv, str):
                        # New section — stop, don't bleed into next table.
                        if vv.strip():
                            break
                        continue
                    if isinstance(vv, (int, float)) and not isinstance(vv, bool):
                        if vv > 0:
                            samples.append(float(vv))
                            if len(samples) >= 6:
                                break
                if len(samples) >= 3:
                    samples.sort()
                    median = samples[len(samples) // 2]
                    # Founders/employees at raw scale will be 100k-30M.
                    # In millions, they're 0.5..30.
                    if 0 < median < 50:
                        info["shares_mult"] = 1_000_000
                        info["shares_reason"] = (
                            f"magnitude:median_shares={median:g} in"
                            f" col {get_column_letter(c)} below '{hv}'"
                        )
                        break
            if info["shares_mult"] != 1:
                break

    return info


def _detect_workbook_units(wb) -> dict[str, dict[str, Any]]:
    """Run unit detection per sheet and return mapping. Also emits an
    EXTRACT_DEBUG log line per sheet that triggers a non-trivial conversion."""
    import os as _os
    debug = _os.environ.get("EXTRACT_DEBUG") == "1"
    out: dict[str, dict[str, Any]] = {}
    for sn in wb.sheetnames:
        try:
            info = _detect_sheet_units(wb[sn])
        except Exception:  # noqa: BLE001
            info = {"dollars_mult": 1, "shares_mult": 1,
                    "dollars_reason": "", "shares_reason": ""}
        out[sn] = info
        if debug and (info["dollars_mult"] != 1 or info["shares_mult"] != 1):
            print(
                f"[extract:units] sheet='{sn}' dollars_mult={info['dollars_mult']}"
                f" ({info['dollars_reason']}) shares_mult={info['shares_mult']}"
                f" ({info['shares_reason']})",
                file=sys.stderr,
            )
    return out


def _unit_scales(wb) -> dict[str, dict[str, Any]]:
    """Lookup-or-empty accessor."""
    return getattr(wb, "_unit_scales", {}) or {}


def _dollars_mult(wb, sheet_name: str | None) -> int:
    if not sheet_name:
        return 1
    return _unit_scales(wb).get(sheet_name, {}).get("dollars_mult", 1)


def _shares_mult(wb, sheet_name: str | None) -> int:
    if not sheet_name:
        return 1
    return _unit_scales(wb).get(sheet_name, {}).get("shares_mult", 1)


# ---------------------------------------------------------------------------
# Per-SAFE extraction (Pro-Forma)
# ---------------------------------------------------------------------------


BINDING_NORMALIZE = [
    # (priority, predicate_on_normalized_text, canonical)
    # Order matters: check MFN-specific first.
    (12, lambda t: ("mfn" in t) and ("15" in t or "safe 15" in t) and "discount" in t,
         "mfn_from_safe_15_discount"),
    (11, lambda t: ("mfn" in t) and ("14" in t or "safe 14" in t),
         "mfn_from_safe_14_cap"),
    (10, lambda t: "mfn_from_safe_14" in t, "mfn_from_safe_14_cap"),
    (10, lambda t: "mfn_from_safe_15" in t, "mfn_from_safe_15_discount"),
    (9,  lambda t: "dormant" in t, "discount"),
    (9,  lambda t: "cap" in t and "discount" in t and "mfn" not in t, "cap_discount"),
    (8,  lambda t: "cap" in t and "mfn" not in t, "cap"),
    (8,  lambda t: "discount" in t and "mfn" not in t, "discount"),
    (5,  lambda t: "cap_discount" in t, "cap_discount"),
    (4,  lambda t: t == "cap", "cap"),
    (4,  lambda t: t == "discount", "discount"),
]


def _normalize_binding(
    raw: Any,
    context_row_text: str = "",
    self_safe_id: int | None = None,
) -> str | None:
    """
    Map messy binding-term strings/cells to the canonical tokens.
    context_row_text lets us pick up MFN targets from an "MFN Elected" cell
    like "$45M cap (SAFE 14)" or "SAFE 15 discount".
    self_safe_id: when processing a SAFE detail row, the ID of that SAFE
    (so a self-reference like "SAFE 14" on SAFE 14's own row isn't treated
    as an MFN election).
    """
    if raw is None and not context_row_text:
        return None
    t = _norm(raw) if raw is not None else ""
    ct = _norm(context_row_text)
    combined = (t + " " + ct).strip()

    # Special MFN heuristic: if the context text references a DIFFERENT SAFE
    # (14 or 15) and this isn't the self-row, treat as MFN-election.
    def _other_safe_refs(s: str) -> list[int]:
        out = []
        for m in re.finditer(r"safe\s*(\d{1,2})", s):
            n = int(m.group(1))
            if n in (14, 15):
                out.append(n)
        return out

    refs = _other_safe_refs(ct)
    # Exclude self-id from refs
    if self_safe_id is not None:
        refs = [n for n in refs if n != self_safe_id]
    if refs and "mfn" not in combined:
        # Prefer the most-specific ref in the context (latest, typically)
        combined = "mfn " + combined + f" safe {refs[-1]}"

    for _prio, pred, canon in BINDING_NORMALIZE:
        try:
            if pred(combined):
                return canon
        except Exception:
            continue
    return None


def _extract_safes(wb, provenance: dict, values: dict) -> None:
    """
    Locate a SAFE detail table; extract binding_term, conversion_price_usd,
    shares_issued for SAFE 1..15.
    """
    sheet_candidates = _fuzzy_sheet_match(wb, (
        "pro-forma", "proforma", "pro forma", "pro_forma",
        "safe", "safes", "cap table", "captable",
    ))
    # Dedup but preserve order
    seen = set()
    ordered = []
    for s in sheet_candidates:
        if s not in seen:
            ordered.append(s)
            seen.add(s)

    for sn in ordered:
        ws = wb[sn]
        if _try_safe_detail_on_sheet(ws, sn, provenance, values):
            return


def _try_safe_detail_on_sheet(ws, sn: str, provenance: dict, values: dict) -> bool:
    # Find header row that has SAFE + (binding/term/conversion/shares/pps)
    rmax = min(ws.max_row or 0, 200)
    header_row = None
    for r in range(1, rmax + 1):
        joined = " ".join(
            _norm(ws.cell(r, c).value)
            for c in range(1, min(ws.max_column or 0, 20) + 1)
            if ws.cell(r, c).value is not None
        )
        if not joined:
            continue
        has_safe = "safe" in joined
        has_binding = ("binding" in joined or "term" in joined or "conversion" in joined
                       or "pps" in joined or "price" in joined)
        has_shares = "shares" in joined
        if has_safe and has_binding and has_shares:
            header_row = r
            break
    if header_row is None:
        return False

    headers = _header_columns(ws, header_row, max_col=30)
    id_col = _find_col_by_keywords(headers, ("safe id", "safe #", "safe", "id"), ("safes",))
    binding_col = _find_col_by_keywords(headers, ("binding", "binding term", "term"), ("terms",))
    pps_col = _find_col_by_keywords(
        headers,
        ("conversion pps", "conversion price", "binding pps", "price per share", "pps"),
        ("series b pps", "cap-implied", "discount-implied", "cap_implied", "discount_implied",
         "cap implied", "discount implied"),
    )
    shares_col = _find_col_by_keywords(headers, ("shares issued", "shares"), ("pre-b", "total"))
    mfn_col = _find_col_by_keywords(headers, ("mfn elected", "mfn"), ("only",))

    if id_col is None:
        # Try left-most column as id if every row below has "SAFE <n>"
        id_col = min(headers.keys())

    # Walk rows below until we've seen SAFE 1..15 or run out / hit "Total"
    found = False
    for r in range(header_row + 1, min(ws.max_row or 0, header_row + 60) + 1):
        id_val = ws.cell(r, id_col).value
        n = _parse_safe_id(id_val)
        if n is None:
            # Maybe the ID is on another column
            for c in headers.keys():
                n = _parse_safe_id(ws.cell(r, c).value)
                if n is not None:
                    break
            if n is None:
                # "Total" row or empty → stop scanning
                label_low = _norm(id_val)
                if "total" in label_low:
                    break
                continue
        if not (1 <= n <= 15):
            continue
        found = True

        # Collect context string for binding detection.
        row_strings = []
        for c in range(1, min(ws.max_column or 0, 20) + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str):
                row_strings.append(v)
        context_text = " | ".join(row_strings)

        # Binding term
        raw_binding = ws.cell(r, binding_col).value if binding_col else None
        # Include MFN column in context for "$45M cap (SAFE 14)" / "Dormant"
        raw_mfn = ws.cell(r, mfn_col).value if mfn_col else None
        extra_ctx = f"{raw_mfn or ''} {context_text}"
        canon = _normalize_binding(raw_binding, extra_ctx, self_safe_id=n)
        # Fallback: look at the broader context only.
        if canon is None:
            canon = _normalize_binding(None, context_text, self_safe_id=n)
        if canon is not None:
            key = f"safe.{n}.binding_term"
            values[key] = canon
            provenance[key] = CellRef(
                sheet=sn,
                cell=_cell_a1(r, binding_col or id_col),
                raw_value=raw_binding if raw_binding is not None else raw_mfn,
                how_matched="safe_detail: binding-term normalize",
            )

        # Conversion PPS
        if pps_col is not None:
            pps_val = _to_decimal(ws.cell(r, pps_col).value)
            if pps_val is not None:
                key = f"safe.{n}.conversion_price_usd"
                values[key] = float(pps_val)
                provenance[key] = CellRef(
                    sheet=sn,
                    cell=_cell_a1(r, pps_col),
                    raw_value=float(pps_val),
                    how_matched="safe_detail: conversion PPS column",
                )

        # Shares issued
        if shares_col is not None:
            sh_val = _to_decimal(ws.cell(r, shares_col).value)
            if sh_val is not None:
                key = f"safe.{n}.shares_issued"
                mult = _shares_mult(ws.parent, sn)
                values[key] = int(round(float(sh_val) * mult))
                how = "safe_detail: shares column"
                if mult != 1:
                    how += f" ×{mult} (units:{_unit_scales(ws.parent)[sn]['shares_reason']})"
                provenance[key] = CellRef(
                    sheet=sn,
                    cell=_cell_a1(r, shares_col),
                    raw_value=float(sh_val),
                    how_matched=how,
                )

    return found


def _parse_safe_id(v: Any) -> int | None:
    if v is None:
        return None
    if isinstance(v, int) and not isinstance(v, bool):
        if 1 <= v <= 15:
            return v
        return None
    if isinstance(v, float):
        if v.is_integer() and 1 <= v <= 15:
            return int(v)
        return None
    if isinstance(v, str):
        m = re.search(r"safe[\s_#:-]*0*(\d{1,2})", v.lower())
        if m:
            n = int(m.group(1))
            if 1 <= n <= 15:
                return n
        # Sometimes just the number
        m2 = re.fullmatch(r"\s*0*(\d{1,2})\s*", v)
        if m2:
            n = int(m2.group(1))
            if 1 <= n <= 15:
                return n
    return None


# ---------------------------------------------------------------------------
# Venture debt extraction
# ---------------------------------------------------------------------------


def _extract_venture_debt(wb, provenance: dict, values: dict) -> None:
    """
    Find single-cell values via label rows across any sheet. We search
    Assumptions, Pro-Forma, and any sheet containing VD-related text.
    """
    # For accreted principal, binding term, conversion price, shares issued,
    # VDW shares — each is a scalar we search for.
    targets: list[tuple[str, list[tuple[str, ...]], str, str]] = [
        # (key, match_groups[list of OR-token sets], kind, description)
        ("venture_debt.accreted_principal_usd",
            [("venture debt", "vd", "debt", "accreted"), ("accreted", "principal")],
            "usd", "accreted principal"),
        ("venture_debt.conversion_price_usd",
            [("venture debt", "vd", "debt"), ("conversion", "pps", "price")],
            "usd", "debt conversion price"),
        ("venture_debt.shares_issued",
            [("venture debt", "vd", "debt"), ("shares",)],
            "shares", "debt shares"),
        ("venture_debt.vdw_shares",
            [("vdw", "warrant"), ("shares",)],
            "shares", "VDW shares"),
    ]

    for key, groups, kind, desc in targets:
        found = _scalar_label_search(wb, groups, disallow=_vd_disallow_for(key))
        if found is None:
            continue
        sheet, row, col, label, val = found
        v = _to_decimal(val)
        if v is None:
            continue
        # Apply per-sheet unit scaling. PPS (kind=usd for "conversion price")
        # is per-share so scale-invariant — it's only "accreted_principal_usd"
        # under kind=usd that genuinely scales with dollars units.
        is_pps = (
            ("conversion" in key and ("price" in key or "pps" in key))
            or key.endswith("_pps_usd")
            or key.endswith(".f_pps_usd")
        )
        scale_how = ""
        if kind == "shares":
            mult = _shares_mult(wb, sheet)
            v = v * Decimal(mult)
            values[key] = int(round(float(v)))
            if mult != 1:
                scale_how = f" ×{mult} (units:{_unit_scales(wb)[sheet]['shares_reason']})"
        elif kind == "usd" and not is_pps:
            mult = _dollars_mult(wb, sheet)
            v = v * Decimal(mult)
            values[key] = float(v)
            if mult != 1:
                scale_how = f" ×{mult} (units:{_unit_scales(wb)[sheet]['dollars_reason']})"
        else:
            values[key] = float(v)
        provenance[key] = CellRef(
            sheet=sheet,
            cell=_cell_a1(row, col),
            raw_value=val,
            how_matched=f"label_scan:{desc} label='{label}'{scale_how}",
        )

    # Binding term: look at Pro-Forma/Assumptions for a row indicating which of
    # 20% discount vs cap was picked. Two useful signals: "Binding Term" column
    # in a VD-specific row, or a narrative like "cap binds" / "discount binds".
    _extract_vd_binding_term(wb, provenance, values)


def _vd_disallow_for(key: str) -> set[str]:
    if key == "venture_debt.conversion_price_usd":
        # Avoid matching "VDW" / "warrant" rows
        return {"vdw", "warrant", "coverage", "accreted"}
    if key == "venture_debt.shares_issued":
        return {"vdw", "warrant", "coverage"}
    if key == "venture_debt.vdw_shares":
        return {"conversion price", "pps", "strike", "coverage %"}
    if key == "venture_debt.accreted_principal_usd":
        return {"warrant", "vdw", "conversion", "pps"}
    return set()


def _scalar_label_search(
    wb,
    token_groups: list[tuple[str, ...]],
    disallow: set[str] | None = None,
    prefer_sheets: tuple[str, ...] = (
        "assumption", "pro-forma", "proforma", "pro forma", "pro_forma",
    ),
) -> tuple[str, int, int, str, Any] | None:
    """
    Search every sheet's label columns for a cell whose label tokens satisfy
    every token_group (OR within group, AND across groups) and isn't in
    `disallow`. Return (sheet, row, col, label_text, nearest_numeric_value).
    Prefers preferred sheets first; within a sheet, prefers rows where the label
    is in the first couple of cols.
    """
    disallow = disallow or set()
    # Sort sheets: preferred first
    sheets: list[str] = []
    for pref in prefer_sheets:
        for sn in wb.sheetnames:
            if pref in _norm(sn) and sn not in sheets:
                sheets.append(sn)
    for sn in wb.sheetnames:
        if sn not in sheets:
            sheets.append(sn)

    for sn in sheets:
        ws = wb[sn]
        rmax = min(ws.max_row or 0, MAX_SCAN_ROWS)
        cmax = min(ws.max_column or 0, LABEL_MAX_COL)
        for r in range(1, rmax + 1):
            # Concatenate first few label cols for this row
            label_text = " ".join(
                str(ws.cell(r, c).value)
                for c in range(1, cmax + 1)
                if ws.cell(r, c).value is not None and isinstance(ws.cell(r, c).value, str)
            )
            if not label_text.strip():
                continue
            n = _norm(label_text)
            if not n:
                continue
            if any(d in n for d in disallow):
                continue
            ok = True
            for group in token_groups:
                if not any(tok in n for tok in group):
                    ok = False
                    break
            if not ok:
                continue
            # Grab the first numeric in this row, scanning from col 2
            nums = _numeric_cells_in_row(ws, r, start_col=1, max_col=30)
            # Exclude numerics that are inside the label columns (quite narrow — if
            # label col has a number we skip that). We want the numeric AFTER the label.
            # Find last label column with text (not number):
            last_label_col = 0
            for c in range(1, cmax + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str) and v.strip():
                    last_label_col = c
            nums = [(c, v) for c, v in nums if c > last_label_col]
            if not nums:
                continue
            col, val = nums[0]
            return sn, r, col, label_text, val
    return None


def _extract_vd_binding_term(wb, provenance: dict, values: dict) -> None:
    """
    Find VD binding term. Primary: a row with venture debt + binding/type
    whose value cell is a cap/discount string. Fallback: infer from the
    conversion PPS vs Series B PPS (if PPS ≤ SB_PPS*(1-discount) by a small
    margin, discount bound; else cap bound).
    """
    sheets_order = _fuzzy_sheet_match(wb, (
        "pro-forma", "proforma", "pro forma", "pro_forma",
        "assumption", "assumptions", "debt", "venture",
    ))
    for sn in sheets_order + [s for s in wb.sheetnames if s not in sheets_order]:
        ws = wb[sn]
        rmax = min(ws.max_row or 0, MAX_SCAN_ROWS)
        for r in range(1, rmax + 1):
            label_parts = []
            last_label_col = 0
            for c in range(1, 6):
                v = ws.cell(r, c).value
                if isinstance(v, str) and v.strip():
                    label_parts.append(v)
                    last_label_col = c
            label = " ".join(label_parts)
            n = _norm(label)
            if not n:
                continue
            has_debt = ("venture debt" in n or " vd " in f" {n} " or " debt " in f" {n} "
                        or n.startswith("debt") or n.endswith("debt"))
            has_topic = ("binding" in n or "term" in n or "type" in n)
            if not (has_debt and has_topic):
                continue
            for c in range(last_label_col + 1, min(ws.max_column or 0, 30) + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str) and v.strip():
                    canon = _normalize_binding(v)
                    if canon in ("cap", "discount"):
                        values["venture_debt.binding_term"] = canon
                        provenance["venture_debt.binding_term"] = CellRef(
                            sheet=sn,
                            cell=_cell_a1(r, c),
                            raw_value=v,
                            how_matched="vd_binding: label/value pair",
                        )
                        return

    # Fallback: infer from debt_conversion_pps vs SB_PPS and discount.
    if "venture_debt.binding_term" in values:
        return
    debt_pps = _to_decimal(values.get("venture_debt.conversion_price_usd"))
    if debt_pps is None or debt_pps <= 0:
        return
    sb_pps = _find_series_b_pps(wb)
    discount = _find_vd_discount(wb)
    if sb_pps is None or discount is None:
        return
    discount_pps = sb_pps * (Decimal(1) - discount)
    # If actual conversion PPS is below discount_pps (within 1% tolerance),
    # cap is binding (since cap gives MORE shares, i.e. lower PPS). Otherwise
    # discount is binding.
    # Note: cap-implied PPS < discount-implied PPS means cap binds.
    if debt_pps < discount_pps * Decimal("0.995"):
        canon = "cap"
    elif debt_pps > discount_pps * Decimal("1.005"):
        # Above discount-implied: something odd; default to discount.
        canon = "discount"
    else:
        canon = "discount"
    values["venture_debt.binding_term"] = canon
    provenance["venture_debt.binding_term"] = CellRef(
        sheet="(inferred)",
        cell="-",
        raw_value=f"debt_pps={debt_pps} sb_pps={sb_pps} disc={discount}",
        how_matched="vd_binding: inferred from PPS vs SB*discount",
    )


def _find_series_b_pps(wb) -> Decimal | None:
    # Look for "Series B PPS" / "SB PPS"
    found = _scalar_label_search(
        wb,
        [{"series b", "sb", "series_b"}, {"pps", "price"}],
        disallow={"cap", "discount", "implied"},
    )
    if found is None:
        return None
    return _to_decimal(found[4])


def _find_vd_discount(wb) -> Decimal | None:
    found = _scalar_label_search(
        wb,
        [{"venture debt", "vd", "debt"}, {"discount"}],
        disallow={"pps", "implied", "warrant", "vdw"},
    )
    if found is None:
        # Try generic "Conversion Discount" row on Debt or Assumptions
        found = _scalar_label_search(
            wb,
            [{"conversion"}, {"discount"}],
            disallow={"safe", "pps", "vdw", "warrant"},
        )
    if found is None:
        return None
    return _to_decimal(found[4])


# ---------------------------------------------------------------------------
# Anti-dilution extraction
# ---------------------------------------------------------------------------


def _extract_antidilution(wb, provenance: dict, values: dict) -> None:
    # Search for label rows like "Seed Adjusted CP", "SA Adjusted CP",
    # "Seed new ratio", etc. v3 added anti-dilution keys for series B/C/D
    # (even though they're dormant for the Nash case).
    # FIX: tightened token groups to exclude "cp" (which matches inside "OCP",
    # the *original* conversion price input cell). Now requires "new" or
    # "adjusted" or "ncp" explicitly. Also added "ocp" / "original" to
    # disallow lists so input-side OCP cells never win the label-scan.
    _NCP_TOK_GROUP = ("conversion", "adjusted", "new", "ncp")  # dropped bare "cp"
    _NCP_VALUE_TOK = ("price", "pps", "ncp")  # dropped bare "cp" here too
    _OCP_DISALLOW = {"ocp", "original conversion", "orig cp"}

    targets: list[tuple[str, list[tuple[str, ...]], str, set[str]]] = [
        # (key, token-groups, description, disallow)
        ("antidilution.seed.new_conversion_price_usd",
            [("seed",), _NCP_TOK_GROUP, _NCP_VALUE_TOK],
            "seed new CP", {"series a", "series_a", "sa ",
                            "series b", "series_b", "sb ",
                            "series c", "series_c", "sc ",
                            "series d", "series_d", "sd "} | _OCP_DISALLOW),
        ("antidilution.series_a.new_conversion_price_usd",
            [("series a", "sa", "series_a"), _NCP_TOK_GROUP, _NCP_VALUE_TOK],
            "series A new CP", {"seed", "series b", "sb ", "series_b",
                                "series c", "sc ", "series_c",
                                "series d", "sd ", "series_d"} | _OCP_DISALLOW),
        ("antidilution.series_b.new_conversion_price_usd",
            [("series b", "sb", "series_b"), _NCP_TOK_GROUP, _NCP_VALUE_TOK],
            "series B new CP", {"seed", "series a", "sa ", "series_a",
                                "series c", "sc ", "series_c",
                                "series d", "sd ", "series_d"} | _OCP_DISALLOW),
        ("antidilution.series_c.new_conversion_price_usd",
            [("series c", "sc", "series_c"), _NCP_TOK_GROUP, _NCP_VALUE_TOK],
            "series C new CP", {"seed", "series a", "sa ", "series_a",
                                "series b", "sb ", "series_b",
                                "series d", "sd ", "series_d"} | _OCP_DISALLOW),
        ("antidilution.series_d.new_conversion_price_usd",
            [("series d", "sd", "series_d"), _NCP_TOK_GROUP, _NCP_VALUE_TOK],
            "series D new CP", {"seed", "series a", "sa ", "series_a",
                                "series b", "sb ", "series_b",
                                "series c", "sc ", "series_c"} | _OCP_DISALLOW),
        ("antidilution.seed.new_ratio",
            [("seed",), ("ratio",)],
            "seed new ratio", {"series a", "series_a", "series b", "series_b",
                               "series c", "series_c", "series d", "series_d"}),
        ("antidilution.series_a.new_ratio",
            [("series a", "sa", "series_a"), ("ratio",)],
            "series A new ratio", {"seed", "series b", "series_b",
                                   "series c", "series_c", "series d", "series_d"}),
        ("antidilution.series_b.new_ratio",
            [("series b", "sb", "series_b"), ("ratio",)],
            "series B new ratio", {"seed", "series a", "series_a",
                                   "series c", "series_c", "series d", "series_d"}),
        ("antidilution.series_c.new_ratio",
            [("series c", "sc", "series_c"), ("ratio",)],
            "series C new ratio", {"seed", "series a", "series_a",
                                   "series b", "series_b", "series d", "series_d"}),
        ("antidilution.series_d.new_ratio",
            [("series d", "sd", "series_d"), ("ratio",)],
            "series D new ratio", {"seed", "series a", "series_a",
                                   "series b", "series_b", "series c", "series_c"}),
    ]

    # FIX: run the table-scan FIRST. The dedicated AD table (header row with
    # NCP / Ratio columns and rows keyed by class) is the canonical layout
    # and gives an internally-consistent (NCP, ratio) pair from the same row.
    # Label-scan is fallback for keys the table-scan didn't find.
    _extract_antidilution_table(wb, provenance, values)

    for key, groups, desc, disallow in targets:
        if values.get(key) is not None:
            continue  # table-scan already filled this key
        found = _scalar_label_search(wb, groups, disallow=disallow)
        if found is None:
            continue
        sn, r, c, label, val = found
        v = _to_decimal(val)
        if v is None:
            continue
        # For ratio: it should be ~1.0; don't rescale.
        values[key] = float(v)
        provenance[key] = CellRef(
            sheet=sn,
            cell=_cell_a1(r, c),
            raw_value=val,
            how_matched=f"antidilution: {desc}",
        )

    # Compute-derived fallback for *.new_ratio when absent.
    # Ratio = OriginalPPS / NewCP  (common per pref share).
    for series in ("seed", "series_a"):
        ratio_key = f"antidilution.{series}.new_ratio"
        if ratio_key in values and values[ratio_key] is not None:
            continue
        ncp_key = f"antidilution.{series}.new_conversion_price_usd"
        if ncp_key not in values:
            continue
        ncp = _to_decimal(values[ncp_key])
        if ncp is None or ncp == 0:
            continue
        # Find original PPS via section-aware scan
        orig = _find_original_pps(wb, series)
        if orig is None or orig == 0:
            continue
        if orig < Decimal("0.01") or orig > Decimal("1000"):
            continue
        ratio = orig / ncp
        values[ratio_key] = float(ratio)
        provenance[ratio_key] = CellRef(
            sheet="(computed)",
            cell="original_pps / new_cp",
            raw_value=float(ratio),
            how_matched=f"antidilution:{series} ratio computed = original_pps / new_cp",
        )


def _extract_antidilution_table(wb, provenance: dict, values: dict) -> None:
    """
    Handle Tetra-style anti-dilution tables:
      Row (header):  Round | OCP | Trigger | NCP | Ratio | As-Conv Shares
      Row (data):    Seed    | 1.00 | No trigger | 1.00 | 1.0 | 5,000,000
      Row (data):    Series A| 3.00 | No trigger | 3.00 | 1.0 | 5,000,000
      ...
    For each series class (seed, series_a, series_b, series_c, series_d),
    look up the row labeled with that class and pull the NCP and ratio values.
    Does NOT overwrite keys already extracted by the label-scan pass.
    """
    class_patterns: list[tuple[str, list[str], list[str]]] = [
        # (series_key, normalized-label tokens that indicate this row,
        #  disallowed tokens)
        ("seed", ["seed"], ["series a", "series_a", "series b", "series_b",
                            "series c", "series_c", "series d", "series_d"]),
        ("series_a", ["series a", "series_a"], ["seed", "series b", "series_b",
                                                 "series c", "series_c",
                                                 "series d", "series_d"]),
        ("series_b", ["series b", "series_b"], ["seed", "series a", "series_a",
                                                 "series c", "series_c",
                                                 "series d", "series_d"]),
        ("series_c", ["series c", "series_c"], ["seed", "series a", "series_a",
                                                 "series b", "series_b",
                                                 "series d", "series_d"]),
        ("series_d", ["series d", "series_d"], ["seed", "series a", "series_a",
                                                 "series b", "series_b",
                                                 "series c", "series_c"]),
    ]

    for sn in wb.sheetnames:
        ws = wb[sn]
        rmax = min(ws.max_row or 0, 200)
        cmax = min(ws.max_column or 0, 20)

        # Find a header row that contains both an NCP-ish column and a Ratio col.
        for hr in range(1, rmax + 1):
            headers_norm: dict[int, str] = {}
            for c in range(1, cmax + 1):
                v = ws.cell(hr, c).value
                if isinstance(v, str):
                    headers_norm[c] = _norm(v)
            joined = " ".join(headers_norm.values())
            if not joined:
                continue
            has_ncp = any(
                k in t for t in headers_norm.values()
                for k in ("ncp", "new cp", "new conversion", "adjusted cp",
                          "adjusted conversion", "new pps", "new conv")
            )
            has_ratio = any("ratio" in t for t in headers_norm.values())
            if not (has_ncp and has_ratio):
                continue
            # Find the NCP and Ratio columns
            # NOTE (Shortcut layout): Shortcut emits "New Conv Price ($)" as
            # the NCP header, which doesn't contain any of the substrings
            # "ncp" / "new cp" / "new conversion" / "adjusted ..." / "new pps".
            # Add "new conv" to the recognizer so Pro-Forma!F is picked up.
            ncp_col = None
            ratio_col = None
            for c, t in headers_norm.items():
                if ncp_col is None and ("ncp" in t or "new cp" in t
                                        or "new conversion" in t or "adjusted cp" in t
                                        or "adjusted conversion" in t or "new pps" in t
                                        or "new conv" in t):
                    ncp_col = c
                if ratio_col is None and "ratio" in t:
                    ratio_col = c
            if ncp_col is None or ratio_col is None:
                continue
            # Label column: prefer an explicit "class"/"round" header near (but
            # ≤) the NCP column, else fall back to leftmost non-header col
            # that's ≤ ncp_col (so we don't grab a label from a side-by-side
            # unrelated table on the left).
            label_col = None
            for c, t in headers_norm.items():
                if c <= ncp_col and (
                    "class" in t or "round" in t or "series" in t
                    or "security" in t or "name" in t
                ):
                    label_col = c
                    break
            if label_col is None:
                # Prefer the rightmost header col that is still ≤ ncp_col and
                # NOT ncp/ratio itself — that is the label column of the
                # antidilution sub-table.
                candidates = [c for c in headers_norm.keys()
                              if c <= ncp_col and c not in (ncp_col, ratio_col)]
                label_col = max(candidates) if candidates else 1

            # Walk data rows below
            for r in range(hr + 1, min(rmax, hr + 30) + 1):
                label = ws.cell(r, label_col).value
                if not isinstance(label, str) or not label.strip():
                    # Allow a blank row then continue? No — stop on blank.
                    all_empty = all(
                        ws.cell(r, c).value is None
                        for c in range(1, min(cmax, 10) + 1)
                    )
                    if all_empty:
                        break
                    continue
                nlabel = _norm(label)
                # Stop if we hit a "Total" row
                if nlabel.startswith("total") or nlabel == "check":
                    break
                # Match to a series class
                for series_key, required, disallowed in class_patterns:
                    if any(d in nlabel for d in disallowed):
                        continue
                    if not any(req in nlabel for req in required):
                        continue
                    # Emit NCP (only if not already set)
                    ncp_key = f"antidilution.{series_key}.new_conversion_price_usd"
                    if ncp_key not in values:
                        v = _to_decimal(ws.cell(r, ncp_col).value)
                        if v is not None:
                            values[ncp_key] = float(v)
                            provenance[ncp_key] = CellRef(
                                sheet=sn,
                                cell=_cell_a1(r, ncp_col),
                                raw_value=ws.cell(r, ncp_col).value,
                                how_matched=f"antidilution_table:{series_key} NCP",
                            )
                    ratio_key = f"antidilution.{series_key}.new_ratio"
                    if ratio_key not in values:
                        v = _to_decimal(ws.cell(r, ratio_col).value)
                        if v is not None:
                            values[ratio_key] = float(v)
                            provenance[ratio_key] = CellRef(
                                sheet=sn,
                                cell=_cell_a1(r, ratio_col),
                                raw_value=ws.cell(r, ratio_col).value,
                                how_matched=f"antidilution_table:{series_key} ratio",
                            )
                    break
            # First successful header row per sheet
            break


def _find_original_pps(wb, series: str) -> Decimal | None:
    """
    Find the original issue price of Seed or Series A preferred.
    Looks for rows labeled "Original PPS" / "PPS" / "Price per share" / "Issue
    Price" that fall under a "Seed" or "Series A" section title in the same
    sheet. Also handles Carta-like "Securities" tabs and a "RoundTerms" tab.
    """
    wanted = "seed" if series == "seed" else "series a"
    other = "series a" if series == "seed" else "seed"

    # Strategy 1: find any row where the label directly contains both
    # series name and (pps/original/price).
    if series == "seed":
        disallow = {"series a", "series_a", "sa "}
        groups = [{"seed"}, {"original pps", "pps", "price per share", "issue price", "original"}]
    else:
        disallow = {"seed"}
        groups = [{"series a", "sa", "series_a"}, {"original pps", "pps", "price per share", "issue price", "original"}]

    # Turn groups into substring match (token-level OR would miss "original pps")
    for sn in wb.sheetnames:
        ws = wb[sn]
        rmax = min(ws.max_row or 0, MAX_SCAN_ROWS)
        cmax = min(ws.max_column or 0, LABEL_MAX_COL)
        for r in range(1, rmax + 1):
            label = " ".join(
                str(ws.cell(r, c).value)
                for c in range(1, cmax + 1)
                if isinstance(ws.cell(r, c).value, str)
            )
            n = _norm(label)
            if not n:
                continue
            if any(d in n for d in disallow):
                continue
            ok = True
            for g in groups:
                if not any(tok in n for tok in g):
                    ok = False
                    break
            if not ok:
                continue
            # Grab numeric after label
            for c in range(1, min(ws.max_column or 0, 30) + 1):
                v = ws.cell(r, c).value
                d = _to_decimal(v)
                if d is not None:
                    return d

    # Strategy 2: section-aware. Find rows whose label starts with a section
    # title matching "wanted" (e.g., "Seed Terms", "Seed"), then scan the next
    # ~20 rows for "Original PPS" / "PPS" / "Issue Price" within that section.
    for sn in wb.sheetnames:
        ws = wb[sn]
        rmax = min(ws.max_row or 0, MAX_SCAN_ROWS)
        cmax = min(ws.max_column or 0, LABEL_MAX_COL)
        section_starts: list[tuple[int, str]] = []
        for r in range(1, rmax + 1):
            for c in range(1, cmax + 1):
                v = ws.cell(r, c).value
                if not isinstance(v, str):
                    continue
                nv = _norm(v)
                if not nv:
                    continue
                if wanted in nv and ("term" in nv or "round" in nv or nv == wanted):
                    section_starts.append((r, v))
                    break
        for sr, _title in section_starts:
            end = sr + 25
            for r in range(sr + 1, min(rmax, end) + 1):
                label_parts = []
                for c in range(1, cmax + 1):
                    v = ws.cell(r, c).value
                    if isinstance(v, str) and v.strip():
                        label_parts.append(v)
                lbl = " ".join(label_parts)
                nl = _norm(lbl)
                if not nl:
                    continue
                # If we hit another section header, stop.
                if other in nl and ("term" in nl or "round" in nl or nl == other):
                    break
                if ("original pps" in nl or "pps" in nl or "issue price" in nl
                        or "original price" in nl or "price per share" in nl):
                    for c in range(1, min(ws.max_column or 0, 30) + 1):
                        v = ws.cell(r, c).value
                        d = _to_decimal(v)
                        if d is not None:
                            return d

    # Strategy 3: a RoundTerms-like tab, with round names in column 1.
    for sn in wb.sheetnames:
        if "round" not in _norm(sn) and "term" not in _norm(sn) and "security" not in _norm(sn):
            continue
        ws = wb[sn]
        # Find column with label match & column with "pps"
        header_row = _find_header_row(
            ws,
            scan_rows=10,
            keywords_required=(("round", "class", "series", "name"), ("pps", "price", "price/share")),
        )
        if header_row is None:
            continue
        headers = _header_columns(ws, header_row, max_col=30)
        label_col = _find_col_by_keywords(headers, ("round", "class", "series", "name"))
        pps_col = _find_col_by_keywords(headers, ("pps", "price", "price/share"))
        if label_col is None or pps_col is None:
            continue
        for r in range(header_row + 1, min(ws.max_row or 0, header_row + 30) + 1):
            lab = ws.cell(r, label_col).value
            if not isinstance(lab, str):
                continue
            nl = _norm(lab)
            if wanted in nl and other not in nl:
                d = _to_decimal(ws.cell(r, pps_col).value)
                if d is not None:
                    return d
    return None


# ---------------------------------------------------------------------------
# Waterfall extraction
# ---------------------------------------------------------------------------


WATERFALL_ROW_RULES: list[tuple[str, list[set[str]], set[str]]] = [
    # (class_key, required groups, disallowed)
    # Order matters: our_fund FIRST (it's a distinct row in v3 Tetra's
    # "Our Fund proceeds (Series E x 1/30)"), then venture_debt_as_b for v0/
    # v1/v2, then series E→A in decreasing specificity. The disallow sets
    # exclude "our"/"fund" from the plain-series rules so that our_fund row
    # isn't accidentally classified as a series row.
    ("our_fund", [{"our"}, {"fund"}], {"moic", "election"}),
    ("venture_debt_as_b",
        [{"venture", "debt", "vd"}, {"residual", "remaining", "unconverted", "as_b", "as b", "pre-conversion"}],
        set()),
    # v3 adds classes E, D, C in addition to A/B.
    ("series_e", [{"series", "se"}, {"e", "se"}],
        {"our", "fund", "election", "reference", "participation"}),
    ("series_d", [{"series", "sd"}, {"d", "sd"}],
        {"our", "fund", "election", "participation"}),
    ("series_c", [{"series", "sc"}, {"c", "sc"}],
        {"our", "fund", "election", "participation"}),
    ("series_b", [{"series", "sb"}, {"b", "sb"}],
        {"our", "fund", "election", "participation"}),
    ("series_a", [{"series", "sa"}, {"a", "sa"}],
        {"our", "fund", "election"}),
    ("seed", [{"seed"}, {"payout", "proceeds", "total", "distribution", "investors"}],
        {"our", "fund", "election"}),
    ("seed", [{"seed"}], {"our", "fund", "election"}),
    ("safes", [{"safe", "safes"}], set()),
    ("common", [{"common", "rsa", "rsas", "founders", "founder"}], {"option", "warrant"}),
    ("options_issued", [{"option", "options"}, {"issued", "issued)", "(issued"}],
        {"unissued", "remaining", "pool", "contribution", "strike receipts"}),
    ("options_issued", [{"options", "option"}, {"net"}, {"strike"}],
        {"unissued", "unissued)", "remaining", "pool", "contribution"}),
    ("warrants", [{"warrant", "warrants", "vdw", "w1", "w2"}], set()),
]

# Election row rules — v3 adds Series C and D elections (5 classes total).
# Labels seen in practice:
#   "Series E election"       → {series, e, election}
#   "E election"              → {e, election}
#   "Series D (elects)"       → {series, d, elects}
# So we accept any label containing the class letter plus election-phrase.
# Rules are first-match-wins, ordered E→A so most-specific letter checks come
# first.
WATERFALL_ELECTION_ROW_RULES: list[tuple[str, list[set[str]], set[str]]] = [
    ("series_e", [{"e", "se"}, {"election", "elects", "picks"}],
        {"reference", "seed"}),
    ("series_d", [{"d", "sd"}, {"election", "elects", "picks"}],
        {"seed"}),
    ("series_c", [{"c", "sc"}, {"election", "elects", "picks"}],
        {"seed"}),
    ("series_b", [{"b", "sb"}, {"election", "elects", "picks"}],
        {"seed"}),
    ("series_a", [{"a", "sa"}, {"election", "elects", "picks"}],
        {"seed"}),
    ("seed", [{"seed"}, {"election", "elects", "picks"}], set()),
]

# Payout row rules (prefer "payout" label over "preference path"/"convert path")
# NOTE: "our" and "fund" are in each series_X disallow set so the dedicated
# `Our Fund proceeds (Series E x 1/30)` row is NOT classified as series_e.
# our_fund is emitted separately (WATERFALL_ROW_RULES's our_fund rule). Also
# note: series ordering DOES matter because rules are iterated first-match-wins
# when called via _label_waterfall_class(), so we list the most-specific series
# first (E, D, C, B, A) to avoid "Series E" accidentally matching series_a via
# the generic {a} token check etc. (which here are already excluded via the
# explicit disallow sets).
WATERFALL_PAYOUT_ROW_RULES: list[tuple[str, list[set[str]], set[str]]] = [
    ("our_fund", [{"our"}, {"fund"},
        {"payout", "proceeds", "total", "distribution", "investors", "$"}],
        {"election", "moic"}),
    ("series_e", [{"series", "se"}, {"e", "se"},
        {"payout", "proceeds", "total", "distribution", "investors", "$"}],
        {"our", "fund", "election", "reference"}),
    ("series_d", [{"series", "sd"}, {"d", "sd"},
        {"payout", "proceeds", "total", "distribution", "investors", "$"}],
        {"our", "fund", "election"}),
    ("series_c", [{"series", "sc"}, {"c", "sc"},
        {"payout", "proceeds", "total", "distribution", "investors", "$"}],
        {"our", "fund", "election"}),
    ("series_b", [{"series", "sb"}, {"b", "sb"},
        {"payout", "proceeds", "total", "distribution", "investors", "$"}],
        {"our", "fund", "election"}),
    ("series_a", [{"series", "sa"}, {"a", "sa"},
        {"payout", "proceeds", "total", "distribution", "investors", "$"}],
        {"our", "fund", "election"}),
    ("seed", [{"seed"},
        {"payout", "proceeds", "total", "distribution", "investors"}],
        {"our", "fund", "election"}),
]


def _label_waterfall_class(label: str, rules) -> str | None:
    tokens = _all_tokens(label)
    for cls, required_groups, disallowed in rules:
        if disallowed & tokens:
            continue
        if all(g & tokens for g in required_groups):
            return cls
    return None


ELECTION_NORMALIZE = [
    (lambda t: "participate" in t or "part-cap" in t or "partcap" in t or "part cap" in t or "participating" in t,
     "participate_capped"),
    (lambda t: "preference" in t or "pref" in t and "path" not in t, "preference"),
    (lambda t: "convert" in t, "convert"),
]


def _normalize_election(v: Any) -> str | None:
    t = _norm(v)
    if not t:
        return None
    for pred, canon in ELECTION_NORMALIZE:
        if pred(t):
            return canon
    return None


def _extract_waterfall(wb, provenance: dict, values: dict) -> None:
    sheet_candidates = _fuzzy_sheet_match(wb, (
        "waterfall", "liquidation", "liquidation analysis",
    ))
    if not sheet_candidates:
        return

    for sn in sheet_candidates:
        ws = wb[sn]
        if _try_waterfall_on_sheet(ws, sn, provenance, values):
            return


def _try_waterfall_on_sheet(ws, sn: str, provenance: dict, values: dict) -> bool:
    # Find the exit-header row: a row containing several numerics that look
    # like exit valuations (i.e., $50M..$20B expressed either as dollars or in
    # $M units), plus optional string forms like "$100M"/"$1B".
    # The exit values themselves become the key suffix (e.g. 700 → "700M")
    # rather than being matched to a hardcoded list — this lets v3's 700/1200/
    # 2000/4000/8000 coexist with v0/v1/v2's 100/250/500/1000/2000.
    def _parse_exit_value(v: Any) -> int | None:
        """Return exit value in $M (integer) if `v` looks like an exit, else None.
        Accepts $dollars (>= $50M), $M-scale (50..20000), and strings like
        '$100M', '1B', '$1.2B'."""
        if v is None or isinstance(v, bool):
            return None
        d = _to_decimal(v) if not isinstance(v, str) else None
        if d is not None:
            dnum = float(d)
            if 50 <= dnum <= 20000 and abs(dnum - round(dnum)) < 1e-6:
                return int(round(dnum))
            # In dollars: $50M..$20B
            if 50_000_000 <= dnum <= 20_000_000_000:
                m = dnum / 1_000_000
                if abs(m - round(m)) < 1e-3:
                    return int(round(m))
            return None
        if isinstance(v, str):
            s = v.strip().replace("$", "").replace(",", "")
            s_low = s.lower()
            m = re.match(r"^\s*\$?\s*([0-9]*\.?[0-9]+)\s*([mbk])?\s*$", s_low)
            if m:
                num = float(m.group(1))
                suf = m.group(2)
                if suf == "b":
                    num *= 1000
                elif suf == "k":
                    num /= 1000.0
                if 50 <= num <= 20000 and abs(num - round(num)) < 1e-3:
                    return int(round(num))
        return None

    exit_header_row = None
    exit_col_map: dict[int, int] = {}   # col → exit_m
    rmax = min(ws.max_row or 0, 80)
    cmax = min(ws.max_column or 0, 40)
    for r in range(1, rmax + 1):
        col_to_exit: dict[int, int] = {}
        for c in range(1, cmax + 1):
            e = _parse_exit_value(ws.cell(r, c).value)
            if e is not None:
                col_to_exit[c] = e
        # Only accept rows where ≥4 columns parse as distinct exit values and
        # they're monotone increasing (exit ladder).
        if len(col_to_exit) >= 4:
            sorted_by_col = [e for _, e in sorted(col_to_exit.items())]
            uniq = len(set(sorted_by_col)) == len(sorted_by_col)
            monotone = all(sorted_by_col[i] < sorted_by_col[i+1]
                           for i in range(len(sorted_by_col) - 1))
            if uniq and monotone:
                exit_header_row = r
                exit_col_map = col_to_exit
                break

    if exit_header_row is None or not exit_col_map:
        return False

    # Preserve original variable semantics used below.
    exits_wanted = sorted(set(exit_col_map.values()))

    # For each data row below header, classify and extract numeric per exit col.
    # We use two passes:
    #  - collect dollar rows per class
    #  - collect election rows per class
    #  - collect payout rows (prefer these over multi-path rows)
    dollars_by_class: dict[str, list[tuple[int, dict[int, Decimal], str, int]]] = {}
    # key: class, value: list of (priority, {exit_m: val}, label, row_idx)
    elections_by_class: dict[str, list[tuple[int, dict[int, str], str, int]]] = {}
    totals_row_idx: int | None = None
    totals_vals: dict[int, Decimal] = {}

    def _looks_helper(label: str) -> bool:
        n = _norm(label)
        # Only treat as helper if it's clearly section-header text, not a
        # dollar row that happens to mention "strike" (e.g. "Issued Options
        # (net of strike)" should still be captured).
        helper_tokens = (
            "helpers", "uncapped pps", "remainder", "total strike receipts",
            "equilibrium", "rough exit", "itm ", "reconciliation", "moic",
            "pps after", "uncapped", "capped excess", "redistrib",
            " path ", "preference path", "convert path", "participation path",
            "conversion path", "part uncapped", "capped?", "strike receipts",
            "effective pps",
            # v3 Tetra cascade helpers:
            "per-share", "per share", "if convert", "convert proceeds",
            "preference proceeds", "pool entering", "pool after",
            "denom entering", "denom after",
            "cascade", "final per-share", "final per share",
            "participation proceeds", "strike receipt",
            # v3 Tetra "Best-Response" helper rows (convert value etc.)
            "convert value", "preference value", "participating value",
            "best response", "best-response",
            "senior prefs", "senior_prefs", "conv_div", "conv div",
            "conv proc", "pref proc",
        )
        for k in helper_tokens:
            if k in n:
                return True
        # Also treat "--- foo ---" / "===" sections as helpers
        raw = label.strip()
        if raw.startswith("---") or raw.startswith("==="):
            return True
        return False

    for r in range(exit_header_row + 1, min(ws.max_row or 0, exit_header_row + 200) + 1):
        # Gather label from first few cols
        label_parts = []
        last_label_col = 0
        for c in range(1, min(ws.max_column or 0, 6) + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip():
                label_parts.append(v)
                last_label_col = c
        label = " ".join(label_parts)
        if not label.strip():
            continue
        if _looks_helper(label):
            continue

        # Get values at each exit col
        row_vals: dict[int, Any] = {}
        for col, e in exit_col_map.items():
            v = ws.cell(r, col).value
            if v is not None:
                row_vals[e] = v
        if not row_vals:
            continue

        n = _norm(label)
        # Totals row. "Check (Total − Exit)" rows often appear right after the
        # actual "Total Distributed" row with all-zero values; don't let them
        # overwrite a real totals capture.
        is_check_row = (
            "check" in n
            and ("total" in n)
            and ("exit" in n or "diff" in n or "delta" in n)
        )
        if is_check_row:
            continue
        if "total distributed" in n or "total dist" in n or n == "total" or "grand total" in n or "total exit" in n:
            # First capture wins.
            if not totals_vals:
                totals_row_idx = r
                for e, v in row_vals.items():
                    d = _to_decimal(v)
                    if d is not None:
                        totals_vals[e] = d
            continue

        # Is this an election row? (string values, not numbers)
        numeric_count = sum(1 for v in row_vals.values() if isinstance(v, (int, float, Decimal)) and not isinstance(v, bool))
        string_count = sum(1 for v in row_vals.values() if isinstance(v, str))
        if string_count >= 2 and string_count > numeric_count:
            cls = _label_waterfall_class(label, WATERFALL_ELECTION_ROW_RULES)
            if cls is None:
                # Fallback: in an ELECTIONS section, the row label may just
                # be a class name ("Series E") without the word "election".
                # If the row's string values all parse as election tokens
                # (preference/convert/participate), classify by series letter
                # alone.
                election_strs = [v for v in row_vals.values()
                                 if isinstance(v, str)
                                 and _normalize_election(v) is not None]
                if len(election_strs) >= 2:
                    cls = _label_waterfall_class(label, WATERFALL_ROW_RULES)
                    # Only accept series-class fallbacks (seed/series_{a,b,c,d,e})
                    if cls not in ("seed", "series_a", "series_b",
                                   "series_c", "series_d", "series_e"):
                        cls = None
            if cls:
                elec = {}
                for e, v in row_vals.items():
                    canon = _normalize_election(v)
                    if canon is not None:
                        elec[e] = canon
                if elec:
                    # priority: "election" rows explicitly labeled
                    prio = 10 if "election" in n or "elects" in n else 5
                    elections_by_class.setdefault(cls, []).append((prio, elec, label, r))
            continue

        # Dollar row classification: prefer "Payout" rows over "preference/convert path"
        cls = _label_waterfall_class(label, WATERFALL_PAYOUT_ROW_RULES)
        prio = 0
        if cls is not None:
            prio = 20  # payout/proceeds preferred
        else:
            cls = _label_waterfall_class(label, WATERFALL_ROW_RULES)
            if cls is None:
                continue
            # Penalize "path" rows and "residual" for non-VD
            if "convert path" in n or "preference path" in n or "pref path" in n:
                continue  # skip preliminary paths
            if cls == "venture_debt_as_b":
                prio = 10
            else:
                prio = 5

        d_map: dict[int, Decimal] = {}
        for e, v in row_vals.items():
            d = _to_decimal(v)
            if d is not None:
                d_map[e] = d
        # Reject single-cell hits when this looks like an exit-by-exit table
        # (≥4 exits in header). A single populated cell almost always means a
        # static-input row (e.g. "Liquidation preferences ($M)" sub-block where
        # each class lists its pref in column C only) that happens to share its
        # row label with an actual distribution row further down. Prefer to
        # let the real distribution row win.
        if d_map and len(exit_col_map) >= 4 and len(d_map) < 2:
            continue
        if d_map:
            dollars_by_class.setdefault(cls, []).append((prio, d_map, label, r))

    # Emit
    # For "common" specifically, the source layout might split founders and
    # RSAs into two rows (e.g. "Common (F1-F3)" and "RSAs (E1-E10)"); sum them.
    # For other classes, prefer a single best hit (highest prio, earliest row).
    dmult = _dollars_mult(ws.parent, sn)
    dscale_tag = ""
    if dmult != 1:
        dscale_tag = f" ×{dmult} (units:{_unit_scales(ws.parent)[sn]['dollars_reason']})"
    for cls, hits in dollars_by_class.items():
        if cls == "common" and len(hits) > 1:
            # Sum all same-class hits (both founders & RSAs contribute to common).
            # DEDUP first: if two rows produce identical dollar maps (some
            # engines emit a recap "Common + RSA" row that mirrors the
            # primary "Common + RSAs" row), summing them double-counts.
            # Keep the first occurrence of each unique dollar-map signature.
            seen_sigs: set[tuple] = set()
            deduped_hits: list = []
            for h in hits:
                _prio, d_map, _label, _row_idx = h
                sig = tuple(sorted((e, float(v)) for e, v in d_map.items()))
                if sig in seen_sigs:
                    continue
                seen_sigs.add(sig)
                deduped_hits.append(h)
            hits = deduped_hits
            combined: dict[int, Decimal] = {}
            labels = []
            rows_used = []
            for _prio, d_map, label, row_idx in hits:
                labels.append(label)
                rows_used.append(row_idx)
                for e, v in d_map.items():
                    combined[e] = combined.get(e, Decimal(0)) + v
            for e, v in combined.items():
                key = f"waterfall.{e}M.{cls}.dollars"
                values[key] = float(v) * dmult
                col = None
                for cc, ee in exit_col_map.items():
                    if ee == e:
                        col = cc
                        break
                provenance[key] = CellRef(
                    sheet=sn,
                    cell="+".join(_cell_a1(r, col or 1) for r in rows_used),
                    raw_value=float(v),
                    how_matched=f"waterfall:common dollars (sum of {len(hits)} rows: {', '.join(labels)[:80]}){dscale_tag}",
                )
            continue

        # Pick highest-priority hit
        hits_sorted = sorted(hits, key=lambda h: (-h[0], h[3]))
        _prio, d_map, label, row_idx = hits_sorted[0]
        for e, v in d_map.items():
            key = f"waterfall.{e}M.{cls}.dollars"
            values[key] = float(v) * dmult
            col = None
            for cc, ee in exit_col_map.items():
                if ee == e:
                    col = cc
                    break
            provenance[key] = CellRef(
                sheet=sn,
                cell=_cell_a1(row_idx, col or 1),
                raw_value=float(v),
                how_matched=f"waterfall:{cls} dollars (row='{label}'){dscale_tag}",
            )

    # Elections
    for cls, hits in elections_by_class.items():
        hits_sorted = sorted(hits, key=lambda h: (-h[0], h[3]))
        _prio, elec_map, label, row_idx = hits_sorted[0]
        for e, canon in elec_map.items():
            key = f"waterfall.{e}M.{cls}.election"
            values[key] = canon
            col = None
            for cc, ee in exit_col_map.items():
                if ee == e:
                    col = cc
                    break
            provenance[key] = CellRef(
                sheet=sn,
                cell=_cell_a1(row_idx, col or 1),
                raw_value=None,
                how_matched=f"waterfall:{cls} election (row='{label}')",
            )

    # Totals
    if totals_vals:
        for e, v in totals_vals.items():
            key = f"waterfall.{e}M.total_distributed"
            values[key] = float(v) * dmult
            col = None
            for cc, ee in exit_col_map.items():
                if ee == e:
                    col = cc
                    break
            provenance[key] = CellRef(
                sheet=sn,
                cell=_cell_a1(totals_row_idx or 0, col or 1),
                raw_value=float(v),
                how_matched=f"waterfall:total_distributed{dscale_tag}",
            )
    else:
        # Fallback: totals = sum of exit values (spec says total must == exit).
        # Actually better: use the exit header values as the totals.
        for col, e in exit_col_map.items():
            v = ws.cell(exit_header_row, col).value
            d = _to_decimal(v)
            if d is None:
                continue
            # Normalize to dollars
            dval = float(d)
            if dval < 10000:  # given as $M
                dval *= 1_000_000
            key = f"waterfall.{e}M.total_distributed"
            values.setdefault(key, dval)
            provenance.setdefault(key, CellRef(
                sheet=sn,
                cell=_cell_a1(exit_header_row, col),
                raw_value=v,
                how_matched="waterfall:total (from exit header)",
            ))

    # Default: venture_debt_as_b dollars = 0 if we couldn't find the row
    for e in exits_wanted:
        key = f"waterfall.{e}M.venture_debt_as_b.dollars"
        if key not in values:
            values[key] = 0.0
            provenance[key] = CellRef(
                sheet=sn,
                cell="-",
                how_matched="waterfall:default venture_debt_as_b=0 (converted at B)",
                raw_value=0.0,
            )

    # v3 compute-derived fallback for our_fund: if our_fund.dollars wasn't
    # found as its own row but series_e.dollars IS available, our_fund is
    # $series_e / 30 per v3 spec (our $10M check = 1/30 of the $300M Series
    # E raise). This is a structural invariant in v3 truth.json.
    #
    # Gate this synthesis to v3 only — detect v4 by the presence of any
    # series_f.dollars waterfall row (a v4-only class). On v4, our_fund's
    # waterfall payout is roughly series_f * (our_fund_shares / total_F_shares),
    # NOT series_e/30 — so the v3 synthesis underestimates by ~24% at high
    # exits and must NOT fire.
    has_series_f = any(
        values.get(f"waterfall.{e}M.series_f.dollars") is not None
        for e in exits_wanted
    )
    if not has_series_f:
        for e in exits_wanted:
            of_key = f"waterfall.{e}M.our_fund.dollars"
            if of_key in values and values[of_key] is not None:
                continue
            se_key = f"waterfall.{e}M.series_e.dollars"
            sev = values.get(se_key)
            if sev is None:
                continue
            try:
                of_val = float(sev) / 30.0
            except (TypeError, ValueError):
                continue
            values[of_key] = of_val
            provenance[of_key] = CellRef(
                sheet="(computed)",
                cell="series_e.dollars / 30",
                raw_value=of_val,
                how_matched="waterfall:our_fund computed = series_e/30 (v3 invariant)",
            )

    return True


# ---------------------------------------------------------------------------
# Sensitivity grid extraction
# ---------------------------------------------------------------------------


def _detect_sensitivity_exit(tnorm: str) -> int | None:
    """Extract the exit value ($M) from a normalized sensitivity grid title.
    Handles phrasings like 'at $500M exit', '$2B', 'at 2000m', 'proceeds @ $2B'."""
    if not tnorm:
        return None
    # Patterns: "at $2b", "$2b exit", "at 2b", "at 2000m", "2000m"
    # Normalized strings already lowercase no-punct — dollar signs can appear.
    m = re.search(r"\$?\s*([0-9]+(?:\.[0-9]+)?)\s*b\b", tnorm)
    if m:
        n = float(m.group(1))
        if 0.05 <= n <= 20:
            return int(round(n * 1000))
    m = re.search(r"\$?\s*([0-9]{2,5})\s*m\b", tnorm)
    if m:
        n = int(m.group(1))
        if 50 <= n <= 20000:
            return n
    # "at 2000" (no suffix, but clearly an exit-$M integer)
    m = re.search(r"\bat\s+\$?\s*([0-9]{2,5})\b", tnorm)
    if m:
        n = int(m.group(1))
        if 50 <= n <= 20000:
            return n
    return None


def _extract_sensitivity(wb, provenance: dict, values: dict) -> None:
    sheet_candidates = _fuzzy_sheet_match(wb, (
        "sensitivity", "scenario", "grid",
    ))
    if not sheet_candidates:
        return
    for sn in sheet_candidates:
        ws = wb[sn]
        _try_sensitivity_on_sheet(ws, sn, provenance, values)


def _try_sensitivity_on_sheet(ws, sn: str, provenance: dict, values: dict) -> None:
    """
    Find 2D grids with an ESOP axis and a pre-money axis. The axis VALUES are
    read from the sheet itself — we don't match to a hardcoded {10%, 12.5%, 15%}
    × {60, 80, 120} set, because v3 uses {10%, 12%, 15%} × {1200, 1500, 2000}
    $M and future scenarios may pick yet different axes. Output keys encode
    the detected axis values (e.g., `esop_12`, `pre_1200M`).

    The dollars grid's exit value is also detected from context (v0/v1/v2 use
    500M; v3 uses 2000M).
    """
    rmax = min(ws.max_row or 0, 200)
    cmax = min(ws.max_column or 0, 60)

    def _val_matches_pre(v: Any) -> int | None:
        """Return pre-money in $M if v parses as a plausible pre-money value
        (range $30M..$5B) that is ALSO a 'round' multiple of $10M. This
        tighter rule avoids false positives from intermediate-calc numbers
        like 40, 51, 99, 277 (Series E PPS, per-share values, etc.) that
        happen to land in the pre-money numeric range."""
        if v is None or isinstance(v, bool):
            return None
        if isinstance(v, (int, float)):
            vv = float(v)
            # $M scale (30..5000) and multiple of 10
            if 30 <= vv <= 5000 and abs(vv - round(vv)) < 0.5 and int(round(vv)) % 10 == 0:
                return int(round(vv))
            # $ scale (multiple of $10M)
            if 30_000_000 <= vv <= 5_000_000_000:
                m = vv / 1_000_000
                if abs(m - round(m)) < 0.5 and int(round(m)) % 10 == 0:
                    return int(round(m))
            return None
        if isinstance(v, str):
            s = v.strip().lower().replace(",", "").replace("$", "").strip()
            # Strip "m" suffix
            suf_mul = 1.0
            if s.endswith("b"):
                suf_mul = 1000.0
                s = s[:-1].strip()
            elif s.endswith("m"):
                s = s[:-1].strip()
            try:
                n = float(s) * suf_mul
                if 30 <= n <= 5000 and abs(n - round(n)) < 0.5 and int(round(n)) % 10 == 0:
                    return int(round(n))
            except ValueError:
                pass
        return None

    def _val_matches_esop(v: Any) -> float | None:
        """Return ESOP fraction (e.g., 0.10, 0.12, 0.125, 0.15) for any
        plausible ESOP axis value in range 5%..30%."""
        if v is None or isinstance(v, bool):
            return None
        if isinstance(v, (int, float)):
            vv = float(v)
            if 0.05 <= vv <= 0.30:
                return round(vv, 4)
            if 5 <= vv <= 30:  # expressed as percent
                return round(vv / 100, 4)
            return None
        if isinstance(v, str):
            s = v.strip().lower().replace("%", "").replace(" ", "")
            try:
                n = float(s)
                if 0.05 <= n <= 0.30:
                    return round(n, 4)
                if 5 <= n <= 30:
                    return round(n / 100, 4)
            except ValueError:
                pass
        return None

    # Find every pre-money-header row. Require at least 2 DISTINCT pre-money
    # values (otherwise a row with "1200 | 1200 | 1200" scratch column headers
    # yields a spurious grid detection).
    pre_header_rows: list[tuple[int, dict[int, int]]] = []  # (row, {col: pre_m})
    for r in range(1, rmax + 1):
        col_to_pre: dict[int, int] = {}
        for c in range(1, cmax + 1):
            pre = _val_matches_pre(ws.cell(r, c).value)
            if pre is not None:
                col_to_pre[c] = pre
        if len(col_to_pre) >= 2 and len(set(col_to_pre.values())) >= 2:
            pre_header_rows.append((r, col_to_pre))

    # For each pre_header_row, find ESOP label rows below (within 10 rows)
    # and collect the 3x3 grid of values. Also try to identify whether this
    # grid is dollars or percentages by the surrounding context (scan 2-5 rows
    # above the pre_header for a title cell).
    for header_row, pre_cols in pre_header_rows:
        # Locate the nearest section title above: the LAST non-empty text row
        # between header_row-6 and header_row-1 that is clearly a standalone
        # title (contains no cells that matched pre-money values).
        title = ""
        # Collect ALL text in the lookback window (up to 12 rows back). Prefer
        # the nearest text row as the primary title, but capture any earlier
        # "Scratch / Per-Scenario / Intermediate" banner so we can skip the
        # whole block as not-a-real-grid.
        lookback_context = ""
        for rr in range(header_row - 1, max(0, header_row - 13), -1):
            has_pre_val = any(_val_matches_pre(ws.cell(rr, cc).value) is not None
                              for cc in range(1, cmax + 1))
            if has_pre_val:
                continue
            row_text_parts = []
            for cc in range(1, cmax + 1):
                v = ws.cell(rr, cc).value
                if isinstance(v, str) and v.strip():
                    row_text_parts.append(v.strip())
            if row_text_parts:
                joined = " ".join(row_text_parts)
                if not title:
                    title = joined
                lookback_context += " " + joined
        context_norm = _norm(lookback_context)
        # Only skip if this looks like a scratch / derivation block — i.e.,
        # explicit "Scratch:" banner or per-cell-calc. Don't use generic
        # phrases like "full recomputation" which can also appear in the
        # legitimate top-of-sheet banner.
        if ("scratch" in context_norm
                or "per scenario computation" in context_norm
                or "per-scenario computation" in context_norm
                or "scenario computation" in context_norm
                or "intermediate calc" in context_norm
                or "intermediate work" in context_norm
                or "first principles" in context_norm
                or "scenario matrix" in context_norm
                or "scenarios (" in context_norm):
            continue
        # ESOP label column: leftmost label col
        esop_col = None
        # Scan rows below header
        grid: dict[float, dict[int, Decimal]] = {}
        data_rows_seen = 0
        for r in range(header_row + 1, min(header_row + 25, rmax) + 1):
            # Find ESOP label in first 3 cols
            esop = None
            label_col = None
            for c in range(1, 5):
                e = _val_matches_esop(ws.cell(r, c).value)
                if e is not None:
                    esop = e
                    label_col = c
                    break
            if esop is None:
                # If we've already collected data, a blank row means grid end
                if data_rows_seen >= 1:
                    any_pre_val = any(
                        _to_decimal(ws.cell(r, c).value) is not None
                        for c in pre_cols.keys()
                    )
                    if not any_pre_val:
                        break
                continue
            row_vals: dict[int, Decimal] = {}
            for c, pre_m in pre_cols.items():
                d = _to_decimal(ws.cell(r, c).value)
                if d is not None:
                    row_vals[pre_m] = d
            if row_vals:
                grid[esop] = row_vals
                data_rows_seen += 1
                esop_col = label_col
            if len(grid) >= 3:
                break

        if not grid:
            continue

        # Classify this grid: dollars vs percentages
        tnorm = _norm(title)
        any_val = next(iter(next(iter(grid.values())).values()))
        avnum = float(any_val)
        # Skip scratch / intermediate calc tables (not a tracked output).
        if ("scratch" in tnorm or "per-scenario computation" in tnorm
                or "per scenario computation" in tnorm
                or "intermediate" in tnorm or "helper" in tnorm):
            continue
        # MOIC always skipped — not a tracked output. Be careful: " x " can
        # legitimately appear in titles like "ESOP x Pre-Money" (axis label,
        # not MOIC). Require stronger MOIC markers: a literal "moic" or
        # "multiple" token, a bare "Nx"/"N x" suffix, or "return multiple".
        is_moic = (
            "moic" in tnorm
            or "multiple" in tnorm
            or bool(re.search(r"\b\d+\s*x\b", tnorm))
            or bool(re.search(r"\bmoic\b", tnorm))
        )
        if is_moic:
            continue
        is_dollars = False
        is_pct = False
        if ("pct" in tnorm or "%" in tnorm or "ownership" in tnorm
                or "our_pct" in tnorm or "our pct" in tnorm
                or "ownership %" in tnorm or "post-e" in tnorm):
            is_pct = True
        elif ("dollar" in tnorm or "proceed" in tnorm or "return" in tnorm
              or "at 500m" in tnorm or "500m" in tnorm or "500 m" in tnorm
              or "at 2000m" in tnorm or "at 2b" in tnorm or "$2b" in tnorm
              or "exit" in tnorm or "$" in tnorm):
            is_dollars = True

        if not is_dollars and not is_pct:
            # Fallback to magnitude
            if avnum > 100:
                is_dollars = True
            elif 0 <= avnum <= 1:
                is_pct = True
            else:
                # Could be MOIC (1-30) — skip
                continue

        # Detect exit value for the dollars grid. Truth keys encode which
        # exit the grid represents (e.g., v0/v1/v2 use 500M, v3 uses 2000M).
        # Look in the title for an exit reference; fall back to the most
        # common exit in the sheet's waterfall.
        exit_m_for_dollars = _detect_sensitivity_exit(tnorm)
        if exit_m_for_dollars is None:
            # Scan the whole sheet for "at $XM exit" or "at $XB exit" text
            for rr in range(max(1, header_row - 12), min(header_row + 2, rmax) + 1):
                for cc in range(1, cmax + 1):
                    val = ws.cell(rr, cc).value
                    if isinstance(val, str):
                        e = _detect_sensitivity_exit(_norm(val))
                        if e is not None:
                            exit_m_for_dollars = e
                            break
                if exit_m_for_dollars is not None:
                    break
        if exit_m_for_dollars is None:
            exit_m_for_dollars = 500  # legacy default for v0/v1/v2

        # Emit — need row/col of each cell, so re-walk the grid storage
        # (we stored esop→pre_m→val; we need the source cell too.)
        # Easiest: re-scan header_row+1..25 again, mirroring the collection.
        grid_cells: dict[float, dict[int, tuple[int, int, Decimal]]] = {}
        for r in range(header_row + 1, min(header_row + 25, rmax) + 1):
            esop_here = None
            for c in range(1, 5):
                e = _val_matches_esop(ws.cell(r, c).value)
                if e is not None:
                    esop_here = e
                    break
            if esop_here is None:
                continue
            for c, pre_m in pre_cols.items():
                d = _to_decimal(ws.cell(r, c).value)
                if d is not None:
                    grid_cells.setdefault(esop_here, {})[pre_m] = (r, c, d)
            if len(grid_cells) >= 3 and all(len(v) >= 3 for v in grid_cells.values()):
                break

        for esop, row_cells in grid_cells.items():
            tok = _sensitivity_esop_token(esop)
            for pre_m, (rr, cc, val) in row_cells.items():
                if is_pct:
                    key = f"sensitivity.{tok}.pre_{pre_m}M.our_pct"
                    pct = _normalize_pct(val)
                    values[key] = float(pct)
                    provenance[key] = CellRef(
                        sheet=sn,
                        cell=_cell_a1(rr, cc),
                        raw_value=float(val),
                        how_matched=f"sensitivity:our_pct title='{title.strip()[:60]}'",
                    )
                else:  # dollars
                    key = f"sensitivity.{tok}.pre_{pre_m}M.dollars_at_{exit_m_for_dollars}M_exit"
                    dv = float(val)
                    # Apply per-sheet explicit unit detection first; fall back to
                    # the magnitude heuristic (sub-$1k cells almost certainly
                    # mean $M) for sheets without an explicit declaration.
                    sens_dmult = _dollars_mult(ws.parent, sn)
                    sens_tag = ""
                    if sens_dmult != 1:
                        dv = dv * sens_dmult
                        sens_tag = (
                            f" ×{sens_dmult}"
                            f" (units:{_unit_scales(ws.parent)[sn]['dollars_reason']})"
                        )
                    elif abs(dv) < 1000:
                        dv = dv * 1_000_000
                        sens_tag = " ×1e6 (magnitude<1000)"
                    values[key] = dv
                    provenance[key] = CellRef(
                        sheet=sn,
                        cell=_cell_a1(rr, cc),
                        raw_value=float(val),
                        how_matched=f"sensitivity:dollars@{exit_m_for_dollars}M title='{title.strip()[:60]}'{sens_tag}",
                    )


# ---------------------------------------------------------------------------
# Return Solver extraction
# ---------------------------------------------------------------------------


def _extract_solver(wb, provenance: dict, values: dict) -> None:
    sheet_candidates = _fuzzy_sheet_match(wb, (
        "return-solver", "return_solver", "return solver", "returns", "solver",
    ))
    if not sheet_candidates:
        return
    for sn in sheet_candidates:
        ws = wb[sn]
        _try_solver_on_sheet(ws, sn, provenance, values)


def _try_solver_on_sheet(ws, sn: str, provenance: dict, values: dict) -> None:
    rmax = min(ws.max_row or 0, 80)
    cmax = min(ws.max_column or 0, 20)

    # Find header row with "3x"/"4x" cells or "3.0×" text
    def _mult_for(v: Any) -> str | None:
        if v is None:
            return None
        if isinstance(v, (int, float)):
            if abs(float(v) - 3.0) < 1e-6:
                return "3x"
            if abs(float(v) - 4.0) < 1e-6:
                return "4x"
            return None
        s = _norm(v)
        if "3x" in s or "3 x" in s or "3.0x" in s or "3x net" in s or s.startswith("3 "):
            return "3x"
        if "4x" in s or "4 x" in s or "4.0x" in s or "4x net" in s or s.startswith("4 "):
            return "4x"
        # Also look for "$15m" / "$20m"
        if "$15" in s:
            return "3x"
        if "$20" in s:
            return "4x"
        return None

    header_row = None
    mult_cols: dict[int, str] = {}
    for r in range(1, rmax + 1):
        col_map = {}
        for c in range(1, cmax + 1):
            m = _mult_for(ws.cell(r, c).value)
            if m is not None:
                col_map[c] = m
        if len(col_map) >= 2:
            header_row = r
            mult_cols = col_map
            break
    if header_row is None:
        # Try transposed layout: mult in rows, exits in columns
        _try_solver_transposed(ws, sn, provenance, values)
        return

    # Data rows: find exit value in first column(s). Accept any plausible exit
    # value ($50M..$20B) rather than a fixed whitelist — v3 uses 700/1200/
    # 2000/4000/8000 vs v0/v1/v2's 100/250/500/1000/2000.
    def _match_exit(v: Any) -> int | None:
        if v is None or isinstance(v, bool):
            return None
        if isinstance(v, (int, float)):
            vv = float(v)
            # $M scale
            if 50 <= vv <= 20000 and abs(vv - round(vv)) < 0.5:
                return int(round(vv))
            # $ scale
            if 50_000_000 <= vv <= 20_000_000_000:
                m = vv / 1_000_000
                if abs(m - round(m)) < 1e-3:
                    return int(round(m))
            return None
        if isinstance(v, str):
            s = v.strip().lower().replace("$", "").replace(",", "").replace(" ", "")
            suf_mul = 1.0
            if s.endswith("b"):
                suf_mul = 1000.0
                s = s[:-1]
            elif s.endswith("m"):
                s = s[:-1]
            try:
                n = float(s) * suf_mul
                if 50 <= n <= 20000 and abs(n - round(n)) < 1e-3:
                    return int(round(n))
            except ValueError:
                return None
        return None

    for r in range(header_row + 1, min(rmax, header_row + 30) + 1):
        exit_m = None
        for c in range(1, 5):
            e = _match_exit(ws.cell(r, c).value)
            if e is not None:
                exit_m = e
                break
        if exit_m is None:
            continue
        for col, mult in mult_cols.items():
            v = ws.cell(r, col).value
            d = _to_decimal(v)
            if d is None:
                continue
            pct = _normalize_pct(d)
            key = f"solver.{mult}.exit_{exit_m}M.required_pct"
            values[key] = float(pct)
            provenance[key] = CellRef(
                sheet=sn,
                cell=_cell_a1(r, col),
                raw_value=float(d),
                how_matched=f"solver:{mult}@{exit_m}M",
            )


def _try_solver_transposed(ws, sn: str, provenance: dict, values: dict) -> None:
    """
    Alternate layout: exits in a header row, multiples in first column.
    """
    rmax = min(ws.max_row or 0, 80)
    cmax = min(ws.max_column or 0, 20)

    def _match_exit(v: Any) -> int | None:
        if v is None or isinstance(v, bool):
            return None
        if isinstance(v, (int, float)):
            vv = float(v)
            if 50 <= vv <= 20000 and abs(vv - round(vv)) < 0.5:
                return int(round(vv))
            if 50_000_000 <= vv <= 20_000_000_000:
                m = vv / 1_000_000
                if abs(m - round(m)) < 1e-3:
                    return int(round(m))
            return None
        if isinstance(v, str):
            s = v.strip().lower().replace("$", "").replace(",", "").replace(" ", "")
            suf_mul = 1.0
            if s.endswith("b"):
                suf_mul = 1000.0
                s = s[:-1]
            elif s.endswith("m"):
                s = s[:-1]
            try:
                n = float(s) * suf_mul
                if 50 <= n <= 20000 and abs(n - round(n)) < 1e-3:
                    return int(round(n))
            except ValueError:
                return None
        return None

    def _mult_for(v: Any) -> str | None:
        if v is None:
            return None
        if isinstance(v, (int, float)):
            if abs(float(v) - 3.0) < 1e-6:
                return "3x"
            if abs(float(v) - 4.0) < 1e-6:
                return "4x"
        s = _norm(v) if isinstance(v, str) else ""
        if "3x" in s or "3 x" in s or "3.0x" in s:
            return "3x"
        if "4x" in s or "4 x" in s or "4.0x" in s:
            return "4x"
        return None

    # Find a row with >= 4 exits
    header_row = None
    exit_cols: dict[int, int] = {}
    for r in range(1, rmax + 1):
        colmap = {}
        for c in range(1, cmax + 1):
            e = _match_exit(ws.cell(r, c).value)
            if e is not None:
                colmap[c] = e
        if len(colmap) >= 4:
            header_row = r
            exit_cols = colmap
            break
    if header_row is None:
        return

    for r in range(header_row + 1, min(rmax, header_row + 30) + 1):
        mult = None
        for c in range(1, 5):
            m = _mult_for(ws.cell(r, c).value)
            if m is not None:
                mult = m
                break
        if mult is None:
            continue
        for col, exit_m in exit_cols.items():
            v = ws.cell(r, col).value
            d = _to_decimal(v)
            if d is None:
                continue
            pct = _normalize_pct(d)
            key = f"solver.{mult}.exit_{exit_m}M.required_pct"
            values[key] = float(pct)
            provenance[key] = CellRef(
                sheet=sn,
                cell=_cell_a1(r, col),
                raw_value=float(d),
                how_matched=f"solver:{mult}@{exit_m}M (transposed)",
            )


# ---------------------------------------------------------------------------
# v4 extension extractors
# ---------------------------------------------------------------------------
#
# kelvin_v4 introduced five new tracked-key families that the v0..v3
# extractor never modeled:
#
#   notes.{note_a,note_b}.* (10 keys)    — convertible-note conversion math
#   tender.*                (3 keys)     — secondary tender mechanics
#   round.*                 (4 keys)     — fixed-point round outputs
#   antidilution.series_e.* (2 keys)     — Series E AD (E's first AD adj.)
#   waterfall.{exit}M.series_f.*         — F preferred line in waterfall
#   waterfall.{exit}M.notes.dollars      — combined NA+NB waterfall line
#   solver.{3,4}x.exit_{700,1200,...}M.* — v4 exit ladder (also v3 ladder)
#   ownership.{f_lead_secondary,note_a,note_b,series_f_new,series_e}.pct
#
# These functions follow the existing `_extract_*` style: defined-name path
# is handled centrally in `_apply_defined_names`, so each function is the
# label-scan / table-scan fallback for keys that defined-names didn't fill.
# Each function NEVER overwrites a value already in `values`.
#
# Diagnostic logging (enabled by `EXTRACT_DEBUG=1`) prints a one-liner per
# key resolved by fuzzy match so the operator can sanity-check pickings.
# ---------------------------------------------------------------------------


def _row_label_text(ws, row: int, max_col: int = 6) -> str:
    """Concatenate the leftmost text cells in a row into a single label."""
    parts: list[str] = []
    cmax = min(ws.max_column or 0, max_col)
    for c in range(1, cmax + 1):
        v = ws.cell(row, c).value
        if isinstance(v, str) and v.strip():
            parts.append(v)
    return " ".join(parts)


def _last_numeric_in_row(
    ws,
    row: int,
    start_col: int = 2,
    max_col: int = 60,
    accept: "callable | None" = None,
) -> tuple[int, Decimal] | None:
    """Return the *last* (rightmost) numeric cell in a row.

    Many engines emit iteration tables ("Iter 0..6" with the converged value
    in the rightmost column). For solved-output keys we prefer the last
    numeric over the first.

    If `accept` is given, it is called with the float value and must return
    True for the cell to be considered. This lets callers reject obviously
    out-of-range values (e.g. ownership-% trailing cells when scanning for
    a share count) without abandoning the row entirely.
    """
    cmax = min(ws.max_column or 0, max_col)
    last: tuple[int, Decimal] | None = None
    for c in range(start_col, cmax + 1):
        d = _to_decimal(ws.cell(row, c).value)
        if d is None:
            continue
        if accept is not None:
            try:
                if not accept(float(d)):
                    continue
            except (TypeError, ValueError):
                continue
        last = (c, d)
    return last


def _scan_label_rows(
    ws,
    max_rows: int = 250,
    label_cols: int = 6,
):
    """Yield (row, normalized_label, raw_label) for every row whose first
    few cols contain non-empty string text. Helper for label-anchored
    region searches."""
    rmax = min(ws.max_row or 0, max_rows)
    for r in range(1, rmax + 1):
        raw = _row_label_text(ws, r, max_col=label_cols)
        if not raw.strip():
            continue
        yield r, _norm(raw), raw


def _debug_log(msg: str) -> None:
    import os as _os
    if _os.environ.get("EXTRACT_DEBUG") == "1":
        print(f"[extract:v4] {msg}", file=sys.stderr)


# ----- Notes block (notes.note_{a,b}.*) ------------------------------------


def _extract_notes(wb, provenance: dict, values: dict) -> None:
    """Extract `notes.note_a.*` and `notes.note_b.*` (10 keys total).

    Strategy:
      1) Look for a notes table whose header row contains BOTH a "cap"
         (or "cap-implied" / "cap pps") column AND a "discount" /
         "disc-implied" column, plus a "branch" or "shares issued"
         column. This is the canonical layout in Tetra/Shortcut/Claude
         Pro-Forma 'Convertible Notes Conversion' / 'CONVERTIBLE NOTE
         CONVERSION' / 'NOTES BLOCK' sections.
      2) For each row whose label contains 'note a' / 'note b' / 'NA1' /
         'NB1', read cap/discount/conv/branch/shares from the header
         columns.
      3) Fallback: scalar label-scan per key on rows like
         "Note A cap PPS" / "Note A Conv PPS" / "Note A Shares" — used
         for engines (e.g. Claude-bare) that emit one-key-per-row helper
         layouts. Always uses the last numeric in the row to skip past
         iter-0 columns.
    """
    note_keys = (
        "cap_implied_pps_usd",
        "discount_implied_pps_usd",
        "conversion_pps_usd",
        "conversion_branch",
        "shares_issued",
    )
    targets = ["note_a", "note_b"]

    # Pass 1: locate a notes table by header row.
    for sn in wb.sheetnames:
        ws = wb[sn]
        snorm = _norm(sn)
        # Tables typically live on Pro-Forma; some put a dedicated "Notes"
        # tab but that's the inputs file in v4 (no computed fields). Skip
        # input-only tabs.
        if snorm in {"stakeholders", "securities", "roundterms", "round terms"}:
            continue
        rmax = min(ws.max_row or 0, 250)
        cmax = min(ws.max_column or 0, 25)

        for hr in range(1, rmax + 1):
            headers: dict[int, str] = {}
            for c in range(1, cmax + 1):
                v = ws.cell(hr, c).value
                if isinstance(v, str) and v.strip():
                    headers[c] = _norm(v)
            joined = " ".join(headers.values())
            if not joined:
                continue
            has_cap = any(("cap" in t and ("pps" in t or "implied" in t or "cap pps" in t))
                          for t in headers.values())
            has_disc = any(("disc" in t and ("pps" in t or "implied" in t))
                           for t in headers.values())
            has_conv = any(("conv" in t and ("pps" in t or "price" in t)) or t == "conversion pps"
                           or "branch" in t or "shares issued" in t or t == "shares"
                           for t in headers.values())
            if not (has_cap and has_disc and has_conv):
                continue

            # Find cols
            def _find(predicates) -> int | None:
                for c, t in headers.items():
                    for p in predicates:
                        if p(t):
                            return c
                return None

            cap_col = _find([
                lambda t: ("cap" in t and ("pps" in t or "implied" in t or "price" in t)
                           and "discount" not in t and "disc" not in t),
            ])
            disc_col = _find([
                lambda t: ("disc" in t and ("pps" in t or "implied" in t or "price" in t)
                           and "cap" not in t),
            ])
            conv_col = _find([
                lambda t: (("conv" in t and ("pps" in t or "price" in t))
                           and "cap" not in t and "disc" not in t),
            ])
            branch_col = _find([lambda t: t == "branch" or t.endswith(" branch") or t.startswith("branch ") or "branch" == t.split()[-1] if t else False])
            if branch_col is None:
                branch_col = _find([lambda t: "branch" in t and "implied" not in t])
            shares_col = _find([lambda t: ("shares" in t) and ("issued" in t or t == "shares")])
            if shares_col is None:
                shares_col = _find([lambda t: "shares" in t and "ratio" not in t])

            if cap_col is None or disc_col is None:
                continue

            # Walk data rows
            label_col = min(headers.keys()) if headers else 1
            for r in range(hr + 1, min(rmax, hr + 30) + 1):
                label = _row_label_text(ws, r, max_col=max(label_col, 3))
                if not label:
                    # Allow blank between header and data, but stop after data
                    continue
                nlabel = _norm(label)
                if "total" in nlabel and "note" not in nlabel:
                    break
                # Match note A / note B (or NA1/NB1)
                target_key: str | None = None
                if ("note a" in nlabel or " na1 " in f" {nlabel} "
                        or nlabel.endswith("na1") or nlabel == "note a"):
                    target_key = "note_a"
                elif ("note b" in nlabel or " nb1 " in f" {nlabel} "
                          or nlabel.endswith("nb1") or nlabel == "note b"):
                    target_key = "note_b"
                if target_key is None:
                    continue

                emit = {
                    "cap_implied_pps_usd": (cap_col, "usd"),
                    "discount_implied_pps_usd": (disc_col, "usd"),
                    "conversion_pps_usd": (conv_col, "usd"),
                    "conversion_branch": (branch_col, "string"),
                    "shares_issued": (shares_col, "shares"),
                }
                for sub, (col, kind) in emit.items():
                    if col is None:
                        continue
                    fkey = f"notes.{target_key}.{sub}"
                    if fkey in values and values[fkey] is not None:
                        continue
                    raw = ws.cell(r, col).value
                    if raw is None:
                        continue
                    if kind == "string":
                        s = str(raw).strip().lower()
                        # Canonicalize: "cap" / "discount" tokens
                        if "cap" in s and "disc" not in s:
                            canon = "cap"
                        elif "disc" in s and "cap" not in s:
                            canon = "discount"
                        else:
                            canon = s
                        values[fkey] = canon
                        provenance[fkey] = CellRef(
                            sheet=sn, cell=_cell_a1(r, col),
                            raw_value=raw,
                            how_matched=f"notes_table:{target_key}.{sub}",
                        )
                    elif kind == "shares":
                        d = _to_decimal(raw)
                        if d is None:
                            continue
                        mult = _shares_mult(wb, sn)
                        values[fkey] = int(round(float(d) * mult))
                        provenance[fkey] = CellRef(
                            sheet=sn, cell=_cell_a1(r, col),
                            raw_value=raw,
                            how_matched=f"notes_table:{target_key}.{sub}"
                                       + (f" ×{mult}" if mult != 1 else ""),
                        )
                    else:  # usd (PPS — scale-invariant)
                        d = _to_decimal(raw)
                        if d is None:
                            continue
                        values[fkey] = float(d)
                        provenance[fkey] = CellRef(
                            sheet=sn, cell=_cell_a1(r, col),
                            raw_value=raw,
                            how_matched=f"notes_table:{target_key}.{sub}",
                        )
            # Don't break — continue scanning so a second table on the
            # same sheet (rare, but possible if headers repeat in a
            # solved-section) gets a chance to fill remaining keys.

    # Pass 2: fallback per-key label scan for one-key-per-row layouts.
    # Many engines emit "Note A Cap PPS", "Note A Conv PPS", "Note A Shares"
    # etc. as separate rows with the value in the rightmost numeric cell
    # (iter table) or the second column (scalar).
    fallback: list[tuple[str, list[tuple[str, ...]], str]] = [
        # (key, token_groups, kind)
        ("notes.note_a.cap_implied_pps_usd",
            [("note", "n_a", "na"), ("a",), ("cap",), ("pps", "implied", "price")],
            "usd_pps"),
        ("notes.note_a.discount_implied_pps_usd",
            [("note", "n_a", "na"), ("a",),
             ("disc", "discount"), ("pps", "implied", "price")],
            "usd_pps"),
        ("notes.note_a.conversion_pps_usd",
            [("note", "n_a", "na"), ("a",),
             ("conv", "conversion"), ("pps", "price")],
            "usd_pps"),
        ("notes.note_a.shares_issued",
            [("note", "n_a", "na"), ("a",), ("shares",)],
            "shares"),
        ("notes.note_a.conversion_branch",
            [("note", "n_a", "na"), ("a",), ("branch",)],
            "branch"),
        ("notes.note_b.cap_implied_pps_usd",
            [("note", "n_b", "nb"), ("b",), ("cap",), ("pps", "implied", "price")],
            "usd_pps"),
        ("notes.note_b.discount_implied_pps_usd",
            [("note", "n_b", "nb"), ("b",),
             ("disc", "discount"), ("pps", "implied", "price")],
            "usd_pps"),
        ("notes.note_b.conversion_pps_usd",
            [("note", "n_b", "nb"), ("b",),
             ("conv", "conversion"), ("pps", "price")],
            "usd_pps"),
        ("notes.note_b.shares_issued",
            [("note", "n_b", "nb"), ("b",), ("shares",)],
            "shares"),
        ("notes.note_b.conversion_branch",
            [("note", "n_b", "nb"), ("b",), ("branch",)],
            "branch"),
    ]

    for key, groups, kind in fallback:
        if values.get(key) is not None:
            continue
        # We need to bias toward "solved/converged" sections; locate by
        # walking sheets manually so we can prefer the last numeric in row.
        # Use the same disallow set across all notes keys to prevent
        # cross-sheet bleed (e.g. helper "Senior pref" rows).
        disallow = {"senior", "remainder"}
        for sn in wb.sheetnames:
            if values.get(key) is not None:
                break
            if _norm(sn) in {"stakeholders", "securities", "roundterms",
                             "round terms", "waterfall"}:
                continue
            ws = wb[sn]
            for r, nlabel, raw_label in _scan_label_rows(ws, max_rows=300):
                if any(d in nlabel for d in disallow):
                    continue
                # AND across groups, OR within group
                ok = True
                for g in groups:
                    if not any(tok in nlabel for tok in g):
                        ok = False
                        break
                if not ok:
                    continue
                # Find last numeric (for iter tables) OR a non-numeric
                # branch string for branch keys.
                if kind == "branch":
                    # Find rightmost string cell that is "cap" or "discount".
                    cmax = min(ws.max_column or 0, 60)
                    found_branch = None
                    for cc in range(cmax, 1, -1):
                        v = ws.cell(r, cc).value
                        if isinstance(v, str) and v.strip():
                            s = v.strip().lower()
                            if "cap" in s and "disc" not in s:
                                found_branch = ("cap", cc, v)
                                break
                            if "disc" in s and "cap" not in s:
                                found_branch = ("discount", cc, v)
                                break
                    if found_branch is None:
                        continue
                    canon, cc, raw = found_branch
                    values[key] = canon
                    provenance[key] = CellRef(
                        sheet=sn, cell=_cell_a1(r, cc),
                        raw_value=raw,
                        how_matched=f"notes_label_scan:{key} label='{raw_label[:60]}'",
                    )
                    _debug_log(f"{key} <- '{raw_label[:60]}' = {canon}")
                    break
                # Numeric kinds: take last numeric in row.
                lr = _last_numeric_in_row(ws, r, start_col=2, max_col=60)
                if lr is None:
                    continue
                col, d = lr
                if kind == "shares":
                    if abs(float(d)) < 1:
                        continue  # skip zero/initial-iter cells
                    mult = _shares_mult(wb, sn)
                    values[key] = int(round(float(d) * mult))
                else:  # usd_pps (per-share, scale-invariant)
                    values[key] = float(d)
                provenance[key] = CellRef(
                    sheet=sn, cell=_cell_a1(r, col),
                    raw_value=float(d),
                    how_matched=f"notes_label_scan:{key} label='{raw_label[:60]}'",
                )
                _debug_log(f"{key} <- '{raw_label[:60]}' = {values[key]}")
                break


# ----- Tender block (tender.*) ---------------------------------------------


def _extract_tender(wb, provenance: dict, values: dict) -> None:
    """Extract `tender.shares_transferred`, `tender.usd_to_founders`,
    `tender.f_lead_holding_secondary_shares`.

    Labels seen in v4 outputs (Tetra/Shortcut/Claude):
      "Tender shares (transferred from founders)"  → shares_transferred
      "Total tender shares (transferred)"          → shares_transferred
      "Tender Shares"                              → shares_transferred
      "Tender shares total"                        → shares_transferred
      "Tender size ($)"                            → usd_to_founders
      "Secondary Raise"                            → usd_to_founders
      "F_LEAD secondary common (from tender)"      → f_lead_holding_secondary_shares
      "F_LEAD Secondary Holding"                   → f_lead_holding_secondary_shares
    """
    targets: list[tuple[str, list[tuple[str, ...]], str, set[str]]] = [
        ("tender.shares_transferred",
            [("tender",), ("shares", "transferred")],
            "shares",
            {"pps", "size", "lead", "f_lead"}),
        ("tender.usd_to_founders",
            [("tender", "secondary"),
             ("size", "raise", "usd", "founders", "$")],
            "usd",
            {"shares", "pps", "multiplier", "ratio"}),
        ("tender.f_lead_holding_secondary_shares",
            [("f_lead", "f lead", "lead"),
             ("secondary",)],
            "shares",
            {"pps", "common", "fund"}),
    ]

    for key, groups, kind, disallow in targets:
        if values.get(key) is not None:
            continue
        for sn in wb.sheetnames:
            if values.get(key) is not None:
                break
            if _norm(sn) in {"stakeholders", "securities", "roundterms",
                             "round terms"}:
                continue
            ws = wb[sn]
            for r, nlabel, raw_label in _scan_label_rows(ws, max_rows=300):
                if any(d in nlabel for d in disallow):
                    continue
                ok = True
                for g in groups:
                    if not any(tok in nlabel for tok in g):
                        ok = False
                        break
                if not ok:
                    continue
                # Use last numeric (handles iter tables and the rare case
                # where the value is in a multi-iter solver row).
                lr = _last_numeric_in_row(ws, r, start_col=2, max_col=60)
                if lr is None:
                    continue
                col, d = lr
                fnum = float(d)
                if kind == "shares":
                    if abs(fnum) < 1:
                        continue
                    mult = _shares_mult(wb, sn)
                    values[key] = int(round(fnum * mult))
                elif kind == "usd":
                    # Tender size truth = $15,000,000. If row reports in
                    # millions ($15) scale up.
                    if abs(fnum) < 1_000_000 and abs(fnum) >= 1:
                        # Assume millions
                        fnum *= 1_000_000
                    values[key] = fnum
                provenance[key] = CellRef(
                    sheet=sn, cell=_cell_a1(r, col),
                    raw_value=float(d),
                    how_matched=f"tender_label_scan:{key} label='{raw_label[:60]}'",
                )
                _debug_log(f"{key} <- '{raw_label[:60]}' = {values[key]}")
                break

    # Last-resort: f_lead_holding_secondary_shares == tender.shares_transferred
    # per spec §12.4 (cross-check). If the dedicated label didn't fire but
    # we have shares_transferred, use that.
    flead_key = "tender.f_lead_holding_secondary_shares"
    if values.get(flead_key) is None:
        st = values.get("tender.shares_transferred")
        if st is not None:
            values[flead_key] = st
            provenance[flead_key] = CellRef(
                sheet="(computed)", cell="tender.shares_transferred",
                raw_value=st,
                how_matched="tender:f_lead_holding = shares_transferred (spec invariant)",
            )


# ----- Round solution (round.*) --------------------------------------------


def _extract_round(wb, provenance: dict, values: dict) -> None:
    """Extract `round.f_pps_usd`, `round.post_f_fd_shares`,
    `round.our_fund_shares`, `round.f_lead_primary_shares`."""

    targets: list[tuple[str, list[tuple[str, ...]], str, set[str]]] = [
        ("round.f_pps_usd",
            [("f", "series f"),
             ("pps", "price")],
            "pps",
            {"cap", "discount", "implied", "tender", "ncp", "ocp",
             "uncapped", "post"}),
        ("round.post_f_fd_shares",
            [("post-f", "post_f", "post f", "t_f", "tf",
              "post-series f", "post series f"),
             ("fd", "shares", "total", "fully-diluted", "fully diluted",
              "t_f")],
            "shares",
            # Normalized labels collapse '-' to space, so disallow on
            # bare 'pre' tokens to skip pre-F-FD inputs (which match
            # the same shape: "Pre-F fully-diluted total").
            {"pre f", "pre-f", "pre_f", "pre fd", "esop", "broad-based",
             "broad based"}),
        ("round.our_fund_shares",
            [("our",), ("fund", "f primary", "f primary shares",
                         "shares", "f-class", "f class")],
            "shares",
            {"pps", "check", "ownership", "moic"}),
        ("round.f_lead_primary_shares",
            [("f_lead", "f lead", "lead"),
             ("primary",)],
            "shares",
            # Disallow USD/raise tokens: rows like "series_f_primary_raise_usd"
            # contain f_lead+primary in their concatenated label but hold a
            # raise dollar amount, not a share count.
            {"check", "secondary", "pps", "raise", "usd", "$",
             "capital", "dollar", "dollars"}),
    ]

    for key, groups, kind, disallow in targets:
        if values.get(key) is not None:
            continue
        # Track all candidates with their (sheet, row, col, value) so we
        # can pick the best by some rule (prefer "solved"/"converged" rows
        # over "iter 0"). Strategy: pick the last numeric in the matching
        # row, but among multiple matching rows, prefer those whose label
        # contains "solved" / "converged" / "final" / "(solved)" markers.
        candidates: list[tuple[int, str, int, int, float]] = []
        # priority, sheet, row, col, value
        for sn in wb.sheetnames:
            if _norm(sn) in {"stakeholders", "securities", "roundterms",
                             "round terms"}:
                continue
            ws = wb[sn]
            for r, nlabel, raw_label in _scan_label_rows(ws, max_rows=300):
                if any(d in nlabel for d in disallow):
                    continue
                ok = True
                for g in groups:
                    if not any(tok in nlabel for tok in g):
                        ok = False
                        break
                if not ok:
                    continue
                # Per-kind sanity filter applied to EACH cell as we walk the
                # row right-to-left, so trailing debug cells (PPS/ownership
                # %s past the share count, USD raise figures past the share
                # count) are skipped rather than disqualifying the whole row.
                if kind == "pps":
                    accept = lambda v: 0.5 < v < 200  # noqa: E731
                elif kind == "shares":
                    # Plausible share counts: >=100 (excludes PPS/ratios/debug)
                    # and <1e8 (excludes USD raise amounts e.g. 30,000,000).
                    accept = lambda v: 100 <= abs(v) < 1e8  # noqa: E731
                else:
                    accept = None
                lr = _last_numeric_in_row(ws, r, start_col=2, max_col=60,
                                           accept=accept)
                if lr is None:
                    continue
                col, d = lr
                fnum = float(d)
                # Priority: solved/converged labels rank highest. Iter
                # tables also bubble up because last-col is the converged
                # value, but a row labeled "solved" / "converged" is more
                # canonical.
                prio = 0
                if "solved" in nlabel or "converged" in nlabel or "final" in nlabel:
                    prio = 10
                # Demote rows that look like inputs (often near top of
                # Assumptions and equal to e.g. raw input). For our_fund
                # shares we want output, not input "Our fund check".
                if "check" in nlabel and "shares" not in nlabel:
                    prio -= 5
                candidates.append((prio, sn, r, col, fnum))
        if not candidates:
            continue
        # Pick highest priority; tie-break by latest sheet (rendered last).
        candidates.sort(key=lambda c: (-c[0],))
        prio, sn, r, col, fnum = candidates[0]
        if kind == "pps":
            values[key] = fnum
        else:  # shares
            mult = _shares_mult(wb, sn)
            values[key] = int(round(fnum * mult))
        provenance[key] = CellRef(
            sheet=sn, cell=_cell_a1(r, col),
            raw_value=fnum,
            how_matched=f"round_label_scan:{key} (prio={prio})",
        )
        _debug_log(f"{key} <- {sn}!{_cell_a1(r, col)} = {values[key]}")


# ----- Series E AD --------------------------------------------------------


def _extract_series_e_ad(wb, provenance: dict, values: dict) -> None:
    """Extend AD extraction to cover Series E.

    The existing `_extract_antidilution_table` scans for an NCP/Ratio
    header row and matches series_{seed,a,b,c,d}. v4 added series_e
    rows. We re-walk the same tables here and emit series_e keys.
    Also runs the label-scan fallback for engines that put each AD
    output on its own row.
    """
    keys_to_fill = [
        ("antidilution.series_e.new_conversion_price_usd",
            [("series e", "se", "series_e", "e ", " e ", "(e)"),
             ("ncp", "new cp", "new conversion", "adjusted",
              "post-f", "post_f", "ncp e", "ncp_e")],
            "usd",
            {"seed", "series a", "series_a", "series b", "series_b",
             "series c", "series_c", "series d", "series_d", "ocp",
             "original"}),
        ("antidilution.series_e.new_ratio",
            [("series e", "se", "series_e", "e ", " e ", "(e)"),
             ("ratio",)],
            "ratio",
            {"seed", "series a", "series_a", "series b", "series_b",
             "series c", "series_c", "series d", "series_d"}),
    ]

    # 1) Re-walk AD tables (same logic as _extract_antidilution_table)
    for sn in wb.sheetnames:
        ws = wb[sn]
        rmax = min(ws.max_row or 0, 250)
        cmax = min(ws.max_column or 0, 25)
        for hr in range(1, rmax + 1):
            headers_norm: dict[int, str] = {}
            for c in range(1, cmax + 1):
                v = ws.cell(hr, c).value
                if isinstance(v, str):
                    headers_norm[c] = _norm(v)
            joined = " ".join(headers_norm.values())
            if not joined:
                continue
            has_ncp = any(
                k in t for t in headers_norm.values()
                for k in ("ncp", "new cp", "new conversion", "adjusted cp",
                          "adjusted conversion", "new pps", "new conv")
            )
            has_ratio = any("ratio" in t for t in headers_norm.values())
            if not (has_ncp and has_ratio):
                continue
            ncp_col = None
            ratio_col = None
            for c, t in headers_norm.items():
                if ncp_col is None and ("ncp" in t or "new cp" in t
                                        or "new conversion" in t
                                        or "adjusted cp" in t
                                        or "adjusted conversion" in t
                                        or "new pps" in t
                                        or "new conv" in t):
                    ncp_col = c
                if ratio_col is None and "ratio" in t:
                    ratio_col = c
            if ncp_col is None or ratio_col is None:
                continue
            label_col = None
            for c, t in headers_norm.items():
                if c <= ncp_col and (
                    "class" in t or "round" in t or "series" in t
                    or "security" in t or "name" in t
                ):
                    label_col = c
                    break
            if label_col is None:
                cands = [c for c in headers_norm.keys()
                         if c <= ncp_col and c not in (ncp_col, ratio_col)]
                label_col = max(cands) if cands else 1

            for r in range(hr + 1, min(rmax, hr + 30) + 1):
                label = ws.cell(r, label_col).value
                if not isinstance(label, str) or not label.strip():
                    continue
                nlabel = _norm(label)
                if nlabel.startswith("total") or nlabel == "check":
                    break
                # Match Series E
                if not (nlabel == "e" or "series e" in nlabel
                        or "series_e" in nlabel
                        or nlabel.startswith("e ") or nlabel.endswith(" e")
                        or nlabel == "se"):
                    continue
                if any(x in nlabel for x in ("seed", "series a", "series_a",
                                              "series b", "series_b",
                                              "series c", "series_c",
                                              "series d", "series_d",
                                              "series f", "series_f")):
                    continue
                ncp_key = "antidilution.series_e.new_conversion_price_usd"
                if values.get(ncp_key) is None:
                    v = _to_decimal(ws.cell(r, ncp_col).value)
                    if v is not None:
                        values[ncp_key] = float(v)
                        provenance[ncp_key] = CellRef(
                            sheet=sn, cell=_cell_a1(r, ncp_col),
                            raw_value=ws.cell(r, ncp_col).value,
                            how_matched="antidilution_table:series_e NCP",
                        )
                ratio_key = "antidilution.series_e.new_ratio"
                if values.get(ratio_key) is None:
                    v = _to_decimal(ws.cell(r, ratio_col).value)
                    if v is not None:
                        values[ratio_key] = float(v)
                        provenance[ratio_key] = CellRef(
                            sheet=sn, cell=_cell_a1(r, ratio_col),
                            raw_value=ws.cell(r, ratio_col).value,
                            how_matched="antidilution_table:series_e ratio",
                        )
                break  # one E row per table
            break  # one header per sheet

    # 2) Label-scan fallback for one-key-per-row layouts (e.g. Tetra
    # Assumptions: "NCP_E (post-F)" / "Post-F NCP for E (effective)").
    for key, groups, kind, disallow in keys_to_fill:
        if values.get(key) is not None:
            continue
        candidates: list[tuple[int, str, int, int, float]] = []
        for sn in wb.sheetnames:
            ws = wb[sn]
            for r, nlabel, raw_label in _scan_label_rows(ws, max_rows=300):
                if any(d in nlabel for d in disallow):
                    continue
                ok = True
                for g in groups:
                    if not any(tok in nlabel for tok in g):
                        ok = False
                        break
                if not ok:
                    continue
                lr = _last_numeric_in_row(ws, r, start_col=2, max_col=60)
                if lr is None:
                    continue
                col, d = lr
                fnum = float(d)
                if kind == "usd":
                    if not (0.5 < fnum < 200):
                        continue
                elif kind == "ratio":
                    if not (0.95 < fnum < 5):
                        continue
                prio = 0
                if "solved" in nlabel or "post-f" in nlabel or "post_f" in nlabel or "effective" in nlabel:
                    prio = 10
                candidates.append((prio, sn, r, col, fnum))
        if not candidates:
            continue
        candidates.sort(key=lambda c: (-c[0],))
        prio, sn, r, col, fnum = candidates[0]
        values[key] = fnum
        provenance[key] = CellRef(
            sheet=sn, cell=_cell_a1(r, col),
            raw_value=fnum,
            how_matched=f"series_e_ad_label_scan:{key} (prio={prio})",
        )
        _debug_log(f"{key} <- {sn}!{_cell_a1(r, col)} = {fnum}")


# ----- Series F & notes waterfall lines -----------------------------------


def _extract_v4_waterfall_extras(wb, provenance: dict, values: dict) -> None:
    """Extend waterfall extraction to cover series_f.* and notes.dollars
    keys that the v0..v3 extractor doesn't model.

    Layout assumption: a waterfall sheet with an exit-header row (already
    detected by `_extract_waterfall`). We re-walk that sheet's data rows
    and pick out:
      - rows whose label contains 'series f' / 'F preferred' / 'F-class' →
        emit series_f.dollars per exit. Election cells are emitted by
        rows like 'Series F election' / 'Election F'.
      - rows whose label contains 'note' (combined notes) → emit
        notes.dollars per exit.
    Re-uses the exit-header detection from `_try_waterfall_on_sheet` by
    duplicating the small parser inline (≤30 LoC) — cheaper than
    refactoring a shared exit-detector across 8 callers.
    """
    sheets = _fuzzy_sheet_match(wb, ("waterfall", "liquidation"))
    if not sheets:
        return

    def _parse_exit(v: Any) -> int | None:
        if v is None or isinstance(v, bool):
            return None
        if isinstance(v, (int, float)):
            vv = float(v)
            if 50 <= vv <= 20000 and abs(vv - round(vv)) < 1e-3:
                return int(round(vv))
            if 50_000_000 <= vv <= 20_000_000_000:
                m = vv / 1_000_000
                if abs(m - round(m)) < 1e-3:
                    return int(round(m))
            return None
        if isinstance(v, str):
            s = v.strip().replace("$", "").replace(",", "").lower()
            m = re.match(r"^\s*([0-9]*\.?[0-9]+)\s*([mbk])?\s*$", s)
            if m:
                num = float(m.group(1))
                suf = m.group(2)
                if suf == "b":
                    num *= 1000
                elif suf == "k":
                    num /= 1000.0
                if 50 <= num <= 20000 and abs(num - round(num)) < 1e-3:
                    return int(round(num))
            # "Exit 1: $700M" / "$1.2B" / "$700M"
            m = re.search(r"\$?\s*([0-9]+(?:\.[0-9]+)?)\s*([mb])\b", s)
            if m:
                num = float(m.group(1))
                suf = m.group(2)
                if suf == "b":
                    num *= 1000
                if 50 <= num <= 20000 and abs(num - round(num)) < 1e-3:
                    return int(round(num))
        return None

    for sn in sheets:
        ws = wb[sn]
        rmax = min(ws.max_row or 0, 250)
        cmax = min(ws.max_column or 0, 30)

        # Find exit-header row: same heuristic as _try_waterfall_on_sheet
        exit_col_map: dict[int, int] = {}
        exit_header_row: int | None = None
        for r in range(1, rmax + 1):
            cm: dict[int, int] = {}
            for c in range(1, cmax + 1):
                e = _parse_exit(ws.cell(r, c).value)
                if e is not None:
                    cm[c] = e
            if len(cm) >= 4:
                vs = [e for _, e in sorted(cm.items())]
                if len(set(vs)) == len(vs) and all(vs[i] < vs[i+1] for i in range(len(vs)-1)):
                    exit_header_row = r
                    exit_col_map = cm
                    break
        if exit_header_row is None:
            continue

        dmult = _dollars_mult(wb, sn)
        scale_tag = ""
        if dmult != 1:
            scale_tag = f" ×{dmult}"

        # Helper rows to skip
        helper_tokens = (
            "iteration block", "individual", "pref ($)", "ind iter",
            "convert@", "iter1", "iter2", "iter3", "iter4", "iter5",
            "aggregates", "check", "total proceeds", "total exit",
            "convert pool", "convert path", "preference path",
            "election ", "elects", "convert@final",
            "iter →", "indicator", "itm ", "indicator (",
            "shares", "preference",  # raw inputs in summary refs
            "convert shares", "ind iter",
        )

        def _is_helper(label_norm: str, row_vals_str: bool) -> bool:
            if any(h in label_norm for h in helper_tokens):
                return True
            return False

        # Pass A: dollar rows — series_f / notes
        # Pass B: election rows — series_f
        # We also want to be conservative: only accept rows where >=4 of 5
        # exit-cols have numeric (for dollars) or string (for elections).

        # Track candidates: (priority, label, row_idx, {exit_m: val})
        candidates_dollars: dict[str, list[tuple[int, str, int, dict[int, Decimal]]]] = {
            "series_f": [], "notes": [],
        }
        candidates_elections: dict[str, list[tuple[int, str, int, dict[int, str]]]] = {
            "series_f": [],
        }

        for r in range(exit_header_row + 1, rmax + 1):
            label = _row_label_text(ws, r, max_col=4)
            if not label.strip():
                continue
            n = _norm(label)

            # Get values at exit cols
            row_vals: dict[int, Any] = {}
            for col, e in exit_col_map.items():
                v = ws.cell(r, col).value
                if v is not None:
                    row_vals[e] = v
            if not row_vals:
                continue

            numeric_count = sum(1 for v in row_vals.values()
                                if isinstance(v, (int, float, Decimal))
                                and not isinstance(v, bool))
            string_count = sum(1 for v in row_vals.values()
                                if isinstance(v, str))

            # Classify Series F dollar/election rows
            is_series_f = (
                ("series f" in n or "series_f" in n
                 or n == "f" or n.startswith("f ")
                 or "f preferred" in n or "f-class" in n or "f class" in n)
                and not any(x in n for x in (
                    "series f_lead", "f_lead", "f lead",
                    "f_fund", "f primary check", "f primary new",
                    "f primary shares", "f primary $",  # input refs
                    "f primary new shares",  # solver helper
                ))
            )

            # Notes (combined) dollar row
            is_notes = (
                ("notes" in n or "note a + note b" in n or "notes (combined)" in n
                 or "notes a+b" in n or "notes proceeds" in n
                 or "na+nb" in n or "n_a + n_b" in n)
                and not ("note a" in n and "note b" not in n)
                and not ("note b" in n and "note a" not in n)
                and "convert" not in n
                and "branch" not in n
                and "shares" not in n
                and "pps" not in n
                and "principal" not in n
                and "cap" not in n
                and "discount" not in n
                and "disc" not in n
            )

            # Election row?
            if string_count >= 2 and string_count > numeric_count:
                if is_series_f and any(_normalize_election(v) is not None
                                        for v in row_vals.values() if isinstance(v, str)):
                    elec_map: dict[int, str] = {}
                    for e, v in row_vals.items():
                        canon = _normalize_election(v)
                        if canon is not None:
                            elec_map[e] = canon
                    if elec_map:
                        prio = 10 if ("election" in n or "elects" in n) else 5
                        candidates_elections["series_f"].append(
                            (prio, label, r, elec_map))
                continue

            # Dollar row
            if numeric_count >= 2:
                if _is_helper(n, False):
                    # Skip iter helper rows but allow proceeds rows
                    if "proceeds" not in n and "final" not in n and "$" not in n and "dollars" not in n:
                        continue
                d_map: dict[int, Decimal] = {}
                for e, v in row_vals.items():
                    d = _to_decimal(v)
                    if d is not None:
                        d_map[e] = d
                if not d_map or len(d_map) < 2:
                    continue
                if is_series_f:
                    prio = 0
                    if "proceeds" in n or "payout" in n or "total" in n or "final" in n or "$" in n:
                        prio = 20
                    elif "iter" in n or "indicator" in n or "convert@" in n:
                        prio = -10
                    else:
                        prio = 5
                    candidates_dollars["series_f"].append((prio, label, r, d_map))
                if is_notes:
                    prio = 10
                    if "proceeds" in n or "$" in n:
                        prio = 20
                    candidates_dollars["notes"].append((prio, label, r, d_map))

        # Emit
        for cls, hits in candidates_dollars.items():
            if not hits:
                continue
            hits.sort(key=lambda h: (-h[0], h[2]))
            prio, label, r_idx, d_map = hits[0]
            for e, v in d_map.items():
                key = f"waterfall.{e}M.{cls}.dollars"
                if key in values and values[key] is not None:
                    continue
                col = None
                for cc, ee in exit_col_map.items():
                    if ee == e:
                        col = cc
                        break
                values[key] = float(v) * dmult
                provenance[key] = CellRef(
                    sheet=sn, cell=_cell_a1(r_idx, col or 1),
                    raw_value=float(v),
                    how_matched=f"waterfall_v4:{cls} dollars row='{label[:60]}'{scale_tag}",
                )

        for cls, hits in candidates_elections.items():
            if not hits:
                continue
            hits.sort(key=lambda h: (-h[0], h[2]))
            prio, label, r_idx, e_map = hits[0]
            for e, canon in e_map.items():
                key = f"waterfall.{e}M.{cls}.election"
                if key in values and values[key] is not None:
                    continue
                col = None
                for cc, ee in exit_col_map.items():
                    if ee == e:
                        col = cc
                        break
                values[key] = canon
                provenance[key] = CellRef(
                    sheet=sn, cell=_cell_a1(r_idx, col or 1),
                    raw_value=None,
                    how_matched=f"waterfall_v4:{cls} election row='{label[:60]}'",
                )


# ----- v4 solver (3x/4x × 700/1200/2000/4000/8000 with company %) ---------


def _extract_v4_solver(wb, provenance: dict, values: dict) -> None:
    """Extract `solver.{3x,4x}.exit_{700,1200,2000,4000,8000}M.required_pct`
    when the existing `_extract_solver` falls short.

    The v3/v4 ladder uses 700/1200/2000/4000/8000M (vs v0/v1/v2's
    100/250/500/1000/2000M). The base extractor's `_match_exit` already
    accepts that range, but `_mult_for` requires ASCII '3x'/'4x' — Tetra
    uses unicode '3×' / '4×', and labels like "Req. F-Class % for 3×"
    versus "Req. Company % for 3×" need disambiguation: the truth is
    the COMPANY percentage, not the F-class percentage.
    """
    sheets = _fuzzy_sheet_match(wb, (
        "return-solver", "return_solver", "return solver", "returns", "solver",
    ))
    if not sheets:
        return

    def _has_mult(text: str, mult: str) -> bool:
        """Check if a text-cell denotes the given multiple (3x or 4x).
        Accepts ASCII 'x', unicode '×' (×), and 'X'."""
        if not text:
            return False
        t = text.lower().replace("×", "x")
        t = re.sub(r"\s+", " ", t)
        if mult == "3x":
            return bool(re.search(r"\b3\s*\.?\s*0?\s*x\b", t)) or " 3x" in t or t.startswith("3x")
        if mult == "4x":
            return bool(re.search(r"\b4\s*\.?\s*0?\s*x\b", t)) or " 4x" in t or t.startswith("4x")
        return False

    def _is_company_pct_label(text: str) -> bool:
        if not text:
            return False
        t = text.lower().replace("×", "x")
        if "company" in t or "company %" in t or "company%" in t:
            return True
        # Fallback: "required ownership" / "required %" without "f-class"
        if "required" in t and "f-class" not in t and "f class" not in t:
            return "%" in t or "pct" in t or "ownership" in t
        return False

    def _is_fclass_pct_label(text: str) -> bool:
        if not text:
            return False
        t = text.lower().replace("×", "x")
        return ("f-class" in t or "f class" in t or "f_class" in t)

    def _match_exit(v: Any) -> int | None:
        if v is None or isinstance(v, bool):
            return None
        if isinstance(v, (int, float)):
            vv = float(v)
            if 50 <= vv <= 20000 and abs(vv - round(vv)) < 1e-3:
                return int(round(vv))
            if 50_000_000 <= vv <= 20_000_000_000:
                m = vv / 1_000_000
                if abs(m - round(m)) < 1e-3:
                    return int(round(m))
            return None
        if isinstance(v, str):
            s = v.strip().lower().replace(",", "")
            # "Exit 1: $700M" / "$1.2B" / "1.2B"
            m = re.search(r"\$?\s*([0-9]+(?:\.[0-9]+)?)\s*([mb])\b", s)
            if m:
                num = float(m.group(1))
                if m.group(2) == "b":
                    num *= 1000
                if 50 <= num <= 20000 and abs(num - round(num)) < 1e-3:
                    return int(round(num))
        return None

    for sn in sheets:
        ws = wb[sn]
        rmax = min(ws.max_row or 0, 80)
        cmax = min(ws.max_column or 0, 30)

        # Find header row that has 3× and 4× cells, plus optionally
        # disambiguation columns for company-vs-fclass.
        # Layout A: ("Exit Value", "...", "Req. F-Class % for 3×",
        #           "Req. F-Class % for 4×", "Req. Company % for 3×",
        #           "Req. Company % for 4×", ...)  (Tetra)
        # Layout B: a single 3×/4× column pair right after "Exit" — old
        # style v0..v3.
        for hr in range(1, rmax + 1):
            hcells: dict[int, str] = {}
            for c in range(1, cmax + 1):
                v = ws.cell(hr, c).value
                if isinstance(v, str) and v.strip():
                    hcells[c] = v.strip()
            if not hcells:
                continue

            # Build maps: 3x-company, 3x-fclass, 4x-company, 4x-fclass
            cols_3x_company: list[int] = []
            cols_3x_fclass: list[int] = []
            cols_4x_company: list[int] = []
            cols_4x_fclass: list[int] = []
            cols_3x_unknown: list[int] = []
            cols_4x_unknown: list[int] = []
            for c, txt in hcells.items():
                t = txt.lower().replace("×", "x")
                has_3 = _has_mult(t, "3x")
                has_4 = _has_mult(t, "4x")
                if not (has_3 or has_4):
                    continue
                is_co = _is_company_pct_label(txt)
                is_fc = _is_fclass_pct_label(txt)
                if has_3:
                    if is_co:
                        cols_3x_company.append(c)
                    elif is_fc:
                        cols_3x_fclass.append(c)
                    else:
                        cols_3x_unknown.append(c)
                if has_4:
                    if is_co:
                        cols_4x_company.append(c)
                    elif is_fc:
                        cols_4x_fclass.append(c)
                    else:
                        cols_4x_unknown.append(c)

            # Pick the right column per multiple. Prefer company; then
            # fall back to unknown (the v0..v3 one-pair-of-columns
            # layout). Skip f-class columns — those are about Series F
            # share of class, not company-level required ownership.
            col_3x = (cols_3x_company[0] if cols_3x_company
                      else (cols_3x_unknown[0] if cols_3x_unknown else None))
            col_4x = (cols_4x_company[0] if cols_4x_company
                      else (cols_4x_unknown[0] if cols_4x_unknown else None))
            if col_3x is None and col_4x is None:
                # No usable columns on this row. Could be a title row
                # ("Required ... to return 3×/4× ...") that mentions both
                # multiples but isn't the actual column-header row.
                # Continue scanning subsequent rows.
                continue
            # We need an exit-bearing data row to trust this header. If
            # nothing matches in the next ~30 rows, fall through to keep
            # scanning later header rows.
            row_count_emitted = 0
            for r in range(hr + 1, min(rmax, hr + 30) + 1):
                # Find exit value in any of first 4 cols (text or num).
                exit_m: int | None = None
                for cc in range(1, 5):
                    e = _match_exit(ws.cell(r, cc).value)
                    if e is not None:
                        exit_m = e
                        break
                if exit_m is None:
                    continue
                for mult, col in (("3x", col_3x), ("4x", col_4x)):
                    if col is None:
                        continue
                    d = _to_decimal(ws.cell(r, col).value)
                    if d is None:
                        continue
                    pct = _normalize_pct(d)
                    key = f"solver.{mult}.exit_{exit_m}M.required_pct"
                    if values.get(key) is not None:
                        continue
                    values[key] = float(pct)
                    provenance[key] = CellRef(
                        sheet=sn, cell=_cell_a1(r, col),
                        raw_value=float(d),
                        how_matched=f"solver_v4:{mult}@{exit_m}M (col={get_column_letter(col)})",
                    )
                    row_count_emitted += 1
            # If this header produced any data rows, stop scanning.
            if row_count_emitted > 0:
                break

    # ------- Layout C: Shortcut/Codex flat-grid Return-Solver --------------
    # Shortcut's Return-Solver:
    #   B5  "Exit Value ($)"  | C5..G5  (700M..8000M)
    #   B16 "Required Ownership (% of T_F)"  (3x section, header at B13)
    #   B22 "Required Ownership (% of T_F)"  (4x section, header at B19)
    # The exits run across columns; the multiple gating is a section title
    # in column B a few rows above. v4_solver Layout-A/B above looks for an
    # exit value in cols 1-4 of the SAME row as the answer cell, which never
    # matches Shortcut/Codex's flat grid.
    #
    # Codex's Return-Solver:
    #   A3.."Exit value", B3="Target multiple", D3="Required Series F class
    #     ownership %", E3="Implied our ownership %"
    #   Rows 4..13: 5 exits × 2 multiples (long form)
    # Both engines write `required_pct` in COMPANY (T_F) terms which is what
    # truth tracks.
    _extract_v4_solver_flat_grid(wb, provenance, values)


def _extract_v4_solver_flat_grid(wb, provenance: dict, values: dict) -> None:
    """Layout C extractor for `solver.{3x,4x}.exit_*.required_pct`.

    Two flat-grid sub-layouts handled:

    * **Shortcut transposed grid.** A "Required Ownership (% of T_F)" row
      with exit values laid out across columns to its right. The 3x and 4x
      grids are stacked vertically; the multiple is a section title (e.g.
      "3× TARGET ($30M)") a few rows above the answer row. We pair each
      "Required Ownership" row with the most-recent multiple-section title
      above it.

    * **Codex long-form grid.** A header row with columns "Exit value",
      "Target multiple", "Implied our ownership %". Each data row is a
      single (exit, multiple) tuple. We look up the our-ownership column
      and emit one key per row.
    """
    sheets = _fuzzy_sheet_match(wb, (
        "return-solver", "return_solver", "return solver", "returns", "solver",
    ))
    if not sheets:
        return

    def _has_mult(text: str, mult: str) -> bool:
        if not text:
            return False
        t = text.lower().replace("×", "x")
        t = re.sub(r"\s+", " ", t)
        if mult == "3x":
            return (bool(re.search(r"(?:^|[^0-9])3\s*\.?\s*0?\s*x\b", t))
                    or " 3x" in t or t.startswith("3x")
                    or "3 x target" in t or "3x target" in t)
        if mult == "4x":
            return (bool(re.search(r"(?:^|[^0-9])4\s*\.?\s*0?\s*x\b", t))
                    or " 4x" in t or t.startswith("4x")
                    or "4 x target" in t or "4x target" in t)
        return False

    def _parse_exit(v: Any) -> int | None:
        if v is None or isinstance(v, bool):
            return None
        if isinstance(v, (int, float)):
            vv = float(v)
            if 50 <= vv <= 20000 and abs(vv - round(vv)) < 1e-3:
                return int(round(vv))
            if 50_000_000 <= vv <= 20_000_000_000:
                m = vv / 1_000_000
                if abs(m - round(m)) < 1e-3:
                    return int(round(m))
            return None
        if isinstance(v, str):
            s = v.strip().lower().replace(",", "")
            m = re.search(r"\$?\s*([0-9]+(?:\.[0-9]+)?)\s*([mb])\b", s)
            if m:
                num = float(m.group(1))
                if m.group(2) == "b":
                    num *= 1000
                if 50 <= num <= 20000 and abs(num - round(num)) < 1e-3:
                    return int(round(num))
        return None

    for sn in sheets:
        ws = wb[sn]
        rmax = min(ws.max_row or 0, 200)
        cmax = min(ws.max_column or 0, 30)

        # ----- Sub-layout 1: Shortcut transposed grid ----------------------
        # Step 1: find a row whose values look like a sequence of exit
        # magnitudes (700M, 1.2B, 2B, 4B, 8B). This is the column header
        # for the flat grid.
        exit_header_row: int | None = None
        exit_col_map: dict[int, int] = {}
        for r in range(1, rmax + 1):
            cm: dict[int, int] = {}
            for c in range(1, cmax + 1):
                e = _parse_exit(ws.cell(r, c).value)
                if e is not None:
                    cm[c] = e
            if len(cm) >= 4:
                vs = [e for _, e in sorted(cm.items())]
                if len(set(vs)) == len(vs) and all(vs[i] < vs[i+1]
                                                    for i in range(len(vs)-1)):
                    exit_header_row = r
                    exit_col_map = cm
                    break

        if exit_header_row is not None:
            # Step 2: walk rows below and identify
            #   (a) "section title" rows that mention 3x or 4x, and
            #   (b) "Required Ownership" rows that hold the percentages.
            # Pair each (b) with the most-recent (a) above it.
            current_mult: str | None = None
            for r in range(exit_header_row + 1, rmax + 1):
                # Section-title detection: look at columns 1-3 for 3×/4×
                # text. Allow "3× TARGET ($30M)" / "3x section" / "Multiple 3x".
                section_text = ""
                for c in range(1, min(cmax, 4) + 1):
                    v = ws.cell(r, c).value
                    if isinstance(v, str) and v.strip():
                        section_text += " " + v.strip()
                if section_text:
                    if _has_mult(section_text, "3x"):
                        current_mult = "3x"
                    elif _has_mult(section_text, "4x"):
                        current_mult = "4x"
                # Answer-row detection
                label = _row_label_text(ws, r, max_col=4)
                if not label.strip() or current_mult is None:
                    continue
                nlabel = _norm(label)
                # Match "Required Ownership" / "Required %" / "Required (% of)"
                # but NOT "Required Fraction of F" (that's the F-class share,
                # not the company %).
                is_company_pct = (
                    ("required" in nlabel and (
                        "ownership" in nlabel or "% of t_f" in nlabel
                        or "% of tf" in nlabel or "company" in nlabel
                        or ("%" in nlabel and "fraction" not in nlabel
                            and "f-class" not in nlabel and "f class" not in nlabel)
                    ))
                )
                if not is_company_pct:
                    continue
                # Walk the exit columns and emit
                for col, exit_m in exit_col_map.items():
                    d = _to_decimal(ws.cell(r, col).value)
                    if d is None:
                        continue
                    pct = _normalize_pct(d)
                    key = f"solver.{current_mult}.exit_{exit_m}M.required_pct"
                    if values.get(key) is not None:
                        continue
                    values[key] = float(pct)
                    provenance[key] = CellRef(
                        sheet=sn, cell=_cell_a1(r, col),
                        raw_value=float(d),
                        how_matched=(f"solver_v4_flat:{current_mult}@{exit_m}M "
                                      f"(label='{label[:40]}')"),
                    )

        # ----- Sub-layout 2: Codex long-form grid --------------------------
        # Find a header row containing "exit", "multiple", and an "ownership"
        # column (preferably "implied our" — that's the company % of T_F).
        for hr in range(1, rmax + 1):
            hcells: dict[int, str] = {}
            for c in range(1, cmax + 1):
                v = ws.cell(hr, c).value
                if isinstance(v, str) and v.strip():
                    hcells[c] = _norm(v)
            if not hcells:
                continue
            exit_col_l: int | None = None
            mult_col: int | None = None
            our_col: int | None = None
            f_class_col: int | None = None
            for c, t in hcells.items():
                if exit_col_l is None and "exit" in t and "value" in t:
                    exit_col_l = c
                elif exit_col_l is None and t.strip() == "exit":
                    exit_col_l = c
                if mult_col is None and ("multiple" in t or "target multiple" in t
                                          or t.strip() == "multiple"):
                    mult_col = c
                # Prefer "implied our" / "company" column for required_pct
                if our_col is None and (
                    "implied our" in t
                    or "our ownership" in t
                    or "company ownership" in t
                    or "company %" in t
                ):
                    our_col = c
                if f_class_col is None and (
                    "f class" in t or "f-class" in t or "series f class" in t
                ):
                    f_class_col = c
            if exit_col_l is None or mult_col is None or our_col is None:
                continue
            # Walk data rows
            for r in range(hr + 1, min(rmax, hr + 30) + 1):
                exit_v = _parse_exit(ws.cell(r, exit_col_l).value)
                if exit_v is None:
                    continue
                mult_raw = ws.cell(r, mult_col).value
                if mult_raw is None or isinstance(mult_raw, bool):
                    continue
                if isinstance(mult_raw, (int, float)):
                    mult_int = int(round(float(mult_raw)))
                    if mult_int == 3:
                        mult_tag = "3x"
                    elif mult_int == 4:
                        mult_tag = "4x"
                    else:
                        continue
                elif isinstance(mult_raw, str):
                    s = mult_raw.lower().replace("×", "x")
                    if "3" in s and "x" in s:
                        mult_tag = "3x"
                    elif "4" in s and "x" in s:
                        mult_tag = "4x"
                    else:
                        continue
                else:
                    continue
                d = _to_decimal(ws.cell(r, our_col).value)
                if d is None:
                    continue
                pct = _normalize_pct(d)
                key = f"solver.{mult_tag}.exit_{exit_v}M.required_pct"
                # Long-form layout has an explicit, unambiguous
                # ("Implied our ownership %") header — trust it over
                # whatever the v0..v3 base solver may have written from
                # an accidental column hit (e.g. a state-counter column
                # whose values happen to look like percentages).
                values[key] = float(pct)
                provenance[key] = CellRef(
                    sheet=sn, cell=_cell_a1(r, our_col),
                    raw_value=float(d),
                    how_matched=(f"solver_v4_long:{mult_tag}@{exit_v}M "
                                  f"col={get_column_letter(our_col)}"),
                )
            break  # one header per sheet


# ----- Codex-specific layout extractors (v4) ------------------------------
#
# Codex (gpt-5-codex via Managed Agents) emits a 5-sheet workbook with
# labels and column orderings that the v0..v3 + early-v4 extractors don't
# match. The forensic post-mortem (`runs/kelvin_v4/analysis/codex_forensic
# .md`) documents 121/167 keys silently returning None on Codex outputs
# even though the model wrote correct values. The functions below add
# Codex-aware label scans for the highest-value 30+ keys: the Pro-Forma
# scalar summary at rows 4-13, the labeled Notes/AD blocks at rows 16-29,
# the stacked-grid Sensitivity (3 sub-tables vertically), and the compact
# Waterfall at rows 4-15 with class labels in column A and exits in B-F.
#
# These functions run AFTER all generic extractors; each only fills keys
# the prior passes left as None. Codex math itself is uniformly off by
# ~0.76% on F-PPS-derived quantities (a known engine-side issue), so
# many extracted values will VC-pass but strict-fail. That's expected
# and is not an extractor concern.


def _extract_codex_layout(wb, provenance: dict, values: dict) -> None:
    """Codex-specific Pro-Forma / Waterfall / Sensitivity / Solver readers.

    Idempotent: only fills keys that are still None after every prior
    extractor strategy has run. Safe to call on non-Codex workbooks (each
    sub-extractor early-returns if its expected layout markers aren't
    present).
    """
    _extract_codex_pro_forma_scalars(wb, provenance, values)
    _extract_codex_pro_forma_iter_ladder(wb, provenance, values)
    _extract_codex_notes_block(wb, provenance, values)
    _extract_codex_ad_block(wb, provenance, values)
    _extract_codex_waterfall(wb, provenance, values)
    _extract_codex_sensitivity(wb, provenance, values)


def _set_if_missing(values: dict, provenance: dict, key: str, val: Any,
                     *, sheet: str, cell: str, raw: Any, how: str) -> None:
    """Helper: only write to values if the slot is currently None.

    This preserves the v4 contract that Codex-specific readers are
    fallbacks, never overrides for keys an earlier extractor confidently
    resolved.
    """
    if values.get(key) is not None:
        return
    if val is None:
        return
    values[key] = val
    provenance[key] = CellRef(
        sheet=sheet, cell=cell, raw_value=raw, how_matched=how,
    )


def _extract_codex_pro_forma_scalars(wb, provenance: dict, values: dict) -> None:
    """Read the Codex Pro-Forma!A4:C13 scalar summary block.

    Layout (from `runs/kelvin_v4/analysis/codex_forensic.md`):
       A4:  "Series F PPS"                       B4 = pps
       A5:  "Series F new shares"                B5 = shares
       A6:  "Our fund Series F shares"           B6 = shares
       A7:  "F_LEAD Series F primary shares"     B7 = shares
       A8:  "Note A converted shares"            B8 = shares
       A9:  "Note B converted shares"            B9 = shares
       A10: "ESOP top-up shares"                 B10 = shares
       A11: "Post-F FD shares"                   B11 = shares
       A12: "Our fund post-F ownership"          B12 = pct
       A13: "F_LEAD secondary common shares"     B13 = shares (= tender shares)
    """
    sheets = _fuzzy_sheet_match(wb, ("pro-forma", "pro forma", "proforma"))
    if not sheets:
        return

    label_to_key: list[tuple[tuple[str, ...], str]] = [
        # tokens (all must be in normalized label), key
        (("series f pps",), "round.f_pps_usd"),
        (("post-f fd shares",), "round.post_f_fd_shares"),
        (("our fund series f shares",), "round.our_fund_shares"),
        (("f_lead series f primary shares",), "round.f_lead_primary_shares"),
        (("f lead series f primary shares",), "round.f_lead_primary_shares"),
        (("note a converted shares",), "notes.note_a.shares_issued"),
        (("note b converted shares",), "notes.note_b.shares_issued"),
        (("f_lead secondary common shares",),
            "tender.f_lead_holding_secondary_shares"),
        (("f lead secondary common shares",),
            "tender.f_lead_holding_secondary_shares"),
        (("our fund post-f ownership",), "ownership.our_fund.pct"),
        (("our fund post f ownership",), "ownership.our_fund.pct"),
    ]

    for sn in sheets:
        ws = wb[sn]
        rmax = min(ws.max_row or 0, 30)  # summary is at top of sheet
        for r in range(1, rmax + 1):
            label = ws.cell(r, 1).value
            if not isinstance(label, str):
                continue
            nlabel = _norm(label)
            if not nlabel:
                continue
            v = ws.cell(r, 2).value
            d = _to_decimal(v)
            if d is None:
                continue
            for tokens, key in label_to_key:
                if all(t in nlabel for t in tokens):
                    if key == "ownership.our_fund.pct":
                        # Already a fraction; normalize defensively
                        val = float(_normalize_pct(d))
                    else:
                        val = float(d)
                    _set_if_missing(
                        values, provenance, key, val,
                        sheet=sn, cell=_cell_a1(r, 2), raw=v,
                        how=f"codex_proforma_summary:{key}",
                    )

        # tender.shares_transferred is the same value as
        # f_lead_holding_secondary_shares (Codex models it that way).
        sb = values.get("tender.f_lead_holding_secondary_shares")
        if (sb is not None
                and values.get("tender.shares_transferred") is None):
            # Try to find the same provenance row
            for r in range(1, rmax + 1):
                label = ws.cell(r, 1).value
                if isinstance(label, str) and "secondary common shares" in _norm(label):
                    _set_if_missing(
                        values, provenance, "tender.shares_transferred",
                        sb,
                        sheet=sn, cell=_cell_a1(r, 2), raw=ws.cell(r, 2).value,
                        how="codex_proforma_summary:tender.shares_transferred (=secondary common)",
                    )
                    break


def _extract_codex_pro_forma_iter_ladder(wb, provenance: dict, values: dict) -> None:
    """Read the converged (last) row of the Codex iteration ladder at
    Pro-Forma!AA3:BE81.

    The ladder column layout (per the forensic report):
        AA  iteration index (0..78)
        AB  F PPS
        AC  F new shares
        AD..AG  AD-trigger gates for Seed/A/B/C
        AH  NCP_D
        AI  NCP_E
        AJ..AM  as-converted shares for Seed..C (post-AD)
        AN  D as-converted (post-AD)
        AO  E as-converted (post-AD)
        AR  Note A cap PPS
        AS  Note A discount PPS
        AT  Note A conversion PPS
        AU  Note A shares
        AV  Note B cap PPS
        AW  Note B discount PPS
        AX  Note B conversion PPS
        AY  Note B shares
        AZ  ESOP top-up
        BA  Pre-money FD
        BB  Post-F FD
        BC  Next-iter PPS
        BD  Note A branch label  ("cap" or "discount")
        BE  Note B branch label

    Strategy: detect the ladder by finding a row block where AA is a
    sequential integer 0..N (N>=20). Take the LAST row's converged
    values. Skip if the workbook isn't Codex-shaped.
    """
    sheets = _fuzzy_sheet_match(wb, ("pro-forma", "pro forma", "proforma"))
    if not sheets:
        return

    # Column letters → 1-based indices
    AA, AB, AC = 27, 28, 29
    AH, AI = 34, 35
    AR, AS_, AT = 44, 45, 46
    AU = 47
    AV, AW, AX = 48, 49, 50
    AY = 51
    AZ, BA, BB = 52, 53, 54
    BD, BE = 56, 57

    for sn in sheets:
        ws = wb[sn]
        rmax = min(ws.max_row or 0, 200)
        # Detect ladder: find rows where AA is integer >=0 and increases
        # by 1 between consecutive rows. Keep last row.
        ladder_rows: list[int] = []
        last_val: float | None = None
        for r in range(1, rmax + 1):
            v = ws.cell(r, AA).value
            if not isinstance(v, (int, float)) or isinstance(v, bool):
                continue
            vv = float(v)
            if vv != int(vv) or vv < 0 or vv > 200:
                continue
            if last_val is None or int(vv) == int(last_val) + 1:
                ladder_rows.append(r)
                last_val = vv
            elif int(vv) == 0:
                # Reset — start a new ladder
                ladder_rows = [r]
                last_val = vv
        if len(ladder_rows) < 20:
            continue
        last_row = ladder_rows[-1]

        # Verify this looks like the Codex ladder: AB should be a
        # plausible PPS (10..200), BB should be a post-F FD (1e7..1e9).
        ab = _to_decimal(ws.cell(last_row, AB).value)
        bb = _to_decimal(ws.cell(last_row, BB).value)
        if ab is None or not (Decimal("5") < ab < Decimal("500")):
            continue
        if bb is None or not (Decimal("1000000") < bb < Decimal("1000000000")):
            continue

        # F PPS
        _set_if_missing(values, provenance, "round.f_pps_usd",
                         float(ab), sheet=sn,
                         cell=_cell_a1(last_row, AB),
                         raw=ws.cell(last_row, AB).value,
                         how=f"codex_iter_ladder:row={last_row} (F PPS)")
        # Post-F FD
        _set_if_missing(values, provenance, "round.post_f_fd_shares",
                         float(bb), sheet=sn,
                         cell=_cell_a1(last_row, BB),
                         raw=ws.cell(last_row, BB).value,
                         how=f"codex_iter_ladder:row={last_row} (post-F FD)")
        # F new shares = primary raise / PPS = F_LEAD primary + our_fund
        ac = _to_decimal(ws.cell(last_row, AC).value)
        if ac is not None:
            # round.f_lead_primary_shares = AC * (20/30) (F_LEAD = 20M, our = 10M)
            # round.our_fund_shares       = AC * (10/30)
            # We can't know the split without input data; skip — these
            # come from the scalar summary block instead.
            pass
        # NCP_D, NCP_E
        ah = _to_decimal(ws.cell(last_row, AH).value)
        if ah is not None:
            _set_if_missing(values, provenance,
                             "antidilution.series_d.new_conversion_price_usd",
                             float(ah), sheet=sn,
                             cell=_cell_a1(last_row, AH),
                             raw=ws.cell(last_row, AH).value,
                             how=f"codex_iter_ladder:row={last_row} NCP_D")
        ai = _to_decimal(ws.cell(last_row, AI).value)
        if ai is not None:
            _set_if_missing(values, provenance,
                             "antidilution.series_e.new_conversion_price_usd",
                             float(ai), sheet=sn,
                             cell=_cell_a1(last_row, AI),
                             raw=ws.cell(last_row, AI).value,
                             how=f"codex_iter_ladder:row={last_row} NCP_E")
        # Notes Note A
        ar = _to_decimal(ws.cell(last_row, AR).value)
        if ar is not None:
            _set_if_missing(values, provenance,
                             "notes.note_a.cap_implied_pps_usd",
                             float(ar), sheet=sn,
                             cell=_cell_a1(last_row, AR),
                             raw=ws.cell(last_row, AR).value,
                             how=f"codex_iter_ladder:row={last_row} note_a cap PPS")
        as_ = _to_decimal(ws.cell(last_row, AS_).value)
        if as_ is not None:
            _set_if_missing(values, provenance,
                             "notes.note_a.discount_implied_pps_usd",
                             float(as_), sheet=sn,
                             cell=_cell_a1(last_row, AS_),
                             raw=ws.cell(last_row, AS_).value,
                             how=f"codex_iter_ladder:row={last_row} note_a disc PPS")
        at = _to_decimal(ws.cell(last_row, AT).value)
        if at is not None:
            _set_if_missing(values, provenance,
                             "notes.note_a.conversion_pps_usd",
                             float(at), sheet=sn,
                             cell=_cell_a1(last_row, AT),
                             raw=ws.cell(last_row, AT).value,
                             how=f"codex_iter_ladder:row={last_row} note_a conv PPS")
        au = _to_decimal(ws.cell(last_row, AU).value)
        if au is not None:
            _set_if_missing(values, provenance,
                             "notes.note_a.shares_issued",
                             float(au), sheet=sn,
                             cell=_cell_a1(last_row, AU),
                             raw=ws.cell(last_row, AU).value,
                             how=f"codex_iter_ladder:row={last_row} note_a shares")
        # Note B
        av = _to_decimal(ws.cell(last_row, AV).value)
        if av is not None:
            _set_if_missing(values, provenance,
                             "notes.note_b.cap_implied_pps_usd",
                             float(av), sheet=sn,
                             cell=_cell_a1(last_row, AV),
                             raw=ws.cell(last_row, AV).value,
                             how=f"codex_iter_ladder:row={last_row} note_b cap PPS")
        aw = _to_decimal(ws.cell(last_row, AW).value)
        if aw is not None:
            _set_if_missing(values, provenance,
                             "notes.note_b.discount_implied_pps_usd",
                             float(aw), sheet=sn,
                             cell=_cell_a1(last_row, AW),
                             raw=ws.cell(last_row, AW).value,
                             how=f"codex_iter_ladder:row={last_row} note_b disc PPS")
        ax = _to_decimal(ws.cell(last_row, AX).value)
        if ax is not None:
            _set_if_missing(values, provenance,
                             "notes.note_b.conversion_pps_usd",
                             float(ax), sheet=sn,
                             cell=_cell_a1(last_row, AX),
                             raw=ws.cell(last_row, AX).value,
                             how=f"codex_iter_ladder:row={last_row} note_b conv PPS")
        ay = _to_decimal(ws.cell(last_row, AY).value)
        if ay is not None:
            _set_if_missing(values, provenance,
                             "notes.note_b.shares_issued",
                             float(ay), sheet=sn,
                             cell=_cell_a1(last_row, AY),
                             raw=ws.cell(last_row, AY).value,
                             how=f"codex_iter_ladder:row={last_row} note_b shares")
        # Branch labels — straight string read
        bd = ws.cell(last_row, BD).value
        if isinstance(bd, str) and bd.lower() in ("cap", "discount"):
            _set_if_missing(values, provenance,
                             "notes.note_a.conversion_branch", bd.lower(),
                             sheet=sn, cell=_cell_a1(last_row, BD),
                             raw=bd,
                             how=f"codex_iter_ladder:row={last_row} note_a branch")
        be = ws.cell(last_row, BE).value
        if isinstance(be, str) and be.lower() in ("cap", "discount"):
            _set_if_missing(values, provenance,
                             "notes.note_b.conversion_branch", be.lower(),
                             sheet=sn, cell=_cell_a1(last_row, BE),
                             raw=be,
                             how=f"codex_iter_ladder:row={last_row} note_b branch")
        return  # one ladder per workbook


def _extract_codex_notes_block(wb, provenance: dict, values: dict) -> None:
    """Read the Codex Pro-Forma "Notes Block" at A26-I29.

    Layout:
       Row 27: Note | Principal | Cap | Discount | Cap-implied PPS |
               Discount-implied PPS | Conversion PPS | Branch | Shares
       Row 28: Note A ...
       Row 29: Note B ...

    More canonical than the iter ladder because the column headers carry
    semantic names directly (avoids the AB-AY column-letter aliasing).
    Runs after the ladder reader so the ladder fills any gaps the
    Notes Block can't.
    """
    sheets = _fuzzy_sheet_match(wb, ("pro-forma", "pro forma", "proforma",
                                     "notes"))
    if not sheets:
        return

    note_label_to_key = {
        "note_a": {
            "cap-implied pps": "notes.note_a.cap_implied_pps_usd",
            "cap implied pps": "notes.note_a.cap_implied_pps_usd",
            "discount-implied pps": "notes.note_a.discount_implied_pps_usd",
            "discount implied pps": "notes.note_a.discount_implied_pps_usd",
            "conversion pps": "notes.note_a.conversion_pps_usd",
            "conv pps": "notes.note_a.conversion_pps_usd",
            "branch": "notes.note_a.conversion_branch",
            "shares issued": "notes.note_a.shares_issued",
            "shares": "notes.note_a.shares_issued",
        },
        "note_b": {
            "cap-implied pps": "notes.note_b.cap_implied_pps_usd",
            "cap implied pps": "notes.note_b.cap_implied_pps_usd",
            "discount-implied pps": "notes.note_b.discount_implied_pps_usd",
            "discount implied pps": "notes.note_b.discount_implied_pps_usd",
            "conversion pps": "notes.note_b.conversion_pps_usd",
            "conv pps": "notes.note_b.conversion_pps_usd",
            "branch": "notes.note_b.conversion_branch",
            "shares issued": "notes.note_b.shares_issued",
            "shares": "notes.note_b.shares_issued",
        },
    }

    for sn in sheets:
        ws = wb[sn]
        rmax = min(ws.max_row or 0, 100)
        cmax = min(ws.max_column or 0, 25)
        # Find a "Notes Block" / "Note" header row
        header_row: int | None = None
        for r in range(1, rmax + 1):
            row_vals: list[str] = []
            for c in range(1, cmax + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str):
                    row_vals.append(_norm(v))
            joined = " ".join(row_vals)
            # A header row should mention "principal" or "cap" + "discount"
            # AND have multiple textual cells suggesting a column header.
            if (("principal" in joined and ("cap" in joined or "shares" in joined))
                or ("cap" in joined and "discount" in joined and "pps" in joined)):
                header_row = r
                break
        if header_row is None:
            continue
        # Header column lookup
        col_label: dict[int, str] = {}
        for c in range(1, cmax + 1):
            v = ws.cell(header_row, c).value
            if isinstance(v, str):
                col_label[c] = _norm(v)
        # Walk rows below
        for r in range(header_row + 1, min(rmax, header_row + 8) + 1):
            label = ws.cell(r, 1).value
            if not isinstance(label, str):
                continue
            nl = _norm(label)
            target_note: str | None = None
            if nl in ("note a", "note_a", "notea") or nl.startswith("note a"):
                target_note = "note_a"
            elif nl in ("note b", "note_b", "noteb") or nl.startswith("note b"):
                target_note = "note_b"
            if target_note is None:
                continue
            for c, cl in col_label.items():
                if c == 1:
                    continue
                key = None
                for kw, k in note_label_to_key[target_note].items():
                    if kw in cl:
                        key = k
                        break
                if key is None:
                    continue
                cv = ws.cell(r, c).value
                if key.endswith(".conversion_branch"):
                    if isinstance(cv, str) and cv.lower() in ("cap", "discount"):
                        _set_if_missing(values, provenance, key, cv.lower(),
                                         sheet=sn, cell=_cell_a1(r, c),
                                         raw=cv,
                                         how=f"codex_notes_block:{key}")
                else:
                    d = _to_decimal(cv)
                    if d is None:
                        continue
                    val = float(d)
                    if key.endswith(".pct"):
                        val = float(_normalize_pct(d))
                    _set_if_missing(values, provenance, key, val,
                                     sheet=sn, cell=_cell_a1(r, c),
                                     raw=cv,
                                     how=f"codex_notes_block:{key}")
        return  # one Notes Block per sheet pair


def _extract_codex_ad_block(wb, provenance: dict, values: dict) -> None:
    """Read the Codex Pro-Forma "AD Block" at rows 16-23.

    Codex's AD Block is a labeled per-class table: the trigger? column is
    True/False, but the NCPs themselves only appear on the iter ladder
    rows AH/AI. The block here captures Seed/A/B/C/D/E rows with their
    OCP and trigger flag. NCP/ratio columns may or may not be present —
    when absent (Codex), we fall back to the iter ladder reader.

    For dormant series (Seed/A/B/C — `Trigger? = False`), we synthesize
    NCP=OCP and ratio=1.0 since by definition no anti-dilution adjustment
    fires when not triggered.
    """
    sheets = _fuzzy_sheet_match(wb, ("pro-forma", "pro forma", "proforma"))
    if not sheets:
        return
    series_map = {
        "seed": "seed",
        "series a": "series_a", "series_a": "series_a",
        "series b": "series_b", "series_b": "series_b",
        "series c": "series_c", "series_c": "series_c",
        "series d": "series_d", "series_d": "series_d",
        "series e": "series_e", "series_e": "series_e",
    }
    for sn in sheets:
        ws = wb[sn]
        rmax = min(ws.max_row or 0, 100)
        cmax = min(ws.max_column or 0, 15)
        # Find an "AD Block" title row, then header row immediately below
        for r in range(1, rmax + 1):
            v = ws.cell(r, 1).value
            if isinstance(v, str) and "ad block" in _norm(v):
                header_row = r + 1
                # Read header
                col_label: dict[int, str] = {}
                for c in range(1, cmax + 1):
                    hv = ws.cell(header_row, c).value
                    if isinstance(hv, str):
                        col_label[c] = _norm(hv)
                # Data rows
                ocp_col = next((c for c, l in col_label.items()
                                 if "ocp" in l), None)
                trig_col = next((c for c, l in col_label.items()
                                  if "trigger" in l), None)
                if ocp_col is None or trig_col is None:
                    break
                for rr in range(header_row + 1, min(rmax, header_row + 12) + 1):
                    lab = ws.cell(rr, 1).value
                    if not isinstance(lab, str):
                        break
                    nlab = _norm(lab)
                    series = None
                    for kw, sk in series_map.items():
                        if nlab == kw or nlab.startswith(kw):
                            series = sk
                            break
                    if series is None:
                        continue
                    trig = ws.cell(rr, trig_col).value
                    ocp = _to_decimal(ws.cell(rr, ocp_col).value)
                    if ocp is None:
                        continue
                    is_triggered = bool(trig) if isinstance(trig, bool) else (
                        isinstance(trig, str) and "true" in trig.lower()
                    )
                    if not is_triggered:
                        # Dormant: NCP = OCP, ratio = 1.0
                        ncp_key = f"antidilution.{series}.new_conversion_price_usd"
                        ratio_key = f"antidilution.{series}.new_ratio"
                        _set_if_missing(values, provenance, ncp_key, float(ocp),
                                         sheet=sn, cell=_cell_a1(rr, ocp_col),
                                         raw=ws.cell(rr, ocp_col).value,
                                         how=f"codex_ad_block:{series} dormant (NCP=OCP)")
                        _set_if_missing(values, provenance, ratio_key, 1.0,
                                         sheet=sn, cell=_cell_a1(rr, trig_col),
                                         raw=trig,
                                         how=f"codex_ad_block:{series} dormant ratio=1.0")
                break
        # else: no AD Block on this sheet, try next


def _extract_codex_waterfall(wb, provenance: dict, values: dict) -> None:
    """Read Codex's compact Waterfall at rows 4-15, cols A-F (proceeds)
    and H-M (election by priced class).

    Layout (per `codex_forensic.md`):
        Row 4:  Class | 700M | 1200M | 2000M | 4000M | 8000M
        Row 5:  F    | $...
        Row 6:  E    | $...
        Row 7:  D    | $...
        Row 8:  C    | $...
        Row 9:  B    | $...
        Row 10: A    | $...
        Row 11: Seed | $...
        Row 12: Notes| $...
        Row 13: Common + RSAs | $...
        Row 14: Issued options| $...
        Row 15: Total| $...
    Election table at H4:M14 with same row layout but text values.

    Codex doesn't model `our_fund` as a separate waterfall line — it's
    folded into series_f (primary) + common (secondary tender). We don't
    fabricate `our_fund.dollars` from secondary inference; the harness
    will return None for those keys and they will score 0 (engine-side
    layout decision).
    """
    sheets = _fuzzy_sheet_match(wb, ("waterfall", "liquidation"))
    if not sheets:
        return

    class_map = {
        # normalized label → ownership key suffix
        "f": "series_f", "series f": "series_f",
        "e": "series_e", "series e": "series_e",
        "d": "series_d", "series d": "series_d",
        "c": "series_c", "series c": "series_c",
        "b": "series_b", "series b": "series_b",
        "a": "series_a", "series a": "series_a",
        "seed": "seed",
        "notes": "notes",
        "common": "common", "common rsas": "common", "common + rsas": "common",
        "common rsa": "common",
        "issued options": "options_issued",
        "options issued": "options_issued",
    }

    for sn in sheets:
        ws = wb[sn]
        rmax = min(ws.max_row or 0, 200)
        cmax = min(ws.max_column or 0, 30)

        # Detect a "Class | 700M | 1200M | ..." header row in cols A-F
        # AND optionally a parallel "Class | 700M | ..." in cols H-M
        # (election table).
        header_row: int | None = None
        proc_cols: dict[int, int] = {}
        for r in range(1, rmax + 1):
            cm: dict[int, int] = {}
            for c in range(2, min(cmax, 12) + 1):
                v = ws.cell(r, c).value
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    fv = float(v)
                    if 50_000_000 <= fv <= 20_000_000_000:
                        m = fv / 1_000_000
                        if abs(m - round(m)) < 1e-3:
                            cm[c] = int(round(m))
            if len(cm) >= 4:
                # Verify col A on this row is "Class" or similar (header)
                a_val = ws.cell(r, 1).value
                if isinstance(a_val, str) and "class" in _norm(a_val):
                    header_row = r
                    proc_cols = cm
                    break
        if header_row is None:
            continue

        # Find optional election header row (could be the same row, with
        # a second "Class" header further right, OR a separate row).
        elec_header_row: int | None = None
        elec_cols: dict[int, int] = {}
        # Same-row search: look for column-of-7+ that has another Class
        # label and exit values to its right.
        for c_class in range(8, min(cmax, 20) + 1):
            v = ws.cell(header_row, c_class).value
            if not isinstance(v, str) or "class" not in _norm(v):
                continue
            cm: dict[int, int] = {}
            for c in range(c_class + 1, min(cmax, c_class + 10) + 1):
                vv = ws.cell(header_row, c).value
                if isinstance(vv, (int, float)) and not isinstance(vv, bool):
                    fv = float(vv)
                    if 50_000_000 <= fv <= 20_000_000_000:
                        m = fv / 1_000_000
                        if abs(m - round(m)) < 1e-3:
                            cm[c] = int(round(m))
            if len(cm) >= 4:
                elec_header_row = header_row
                elec_cols = cm
                break

        # Walk the data rows
        for r in range(header_row + 1, min(rmax, header_row + 20) + 1):
            label_cell = ws.cell(r, 1).value
            if not isinstance(label_cell, str):
                continue
            nl = _norm(label_cell)
            # Total row gets handled separately (it's not a class).
            is_total = nl == "total" or nl.startswith("total ")
            cls_key = None
            if not is_total:
                # Match longest label first (so "common rsas" beats "common")
                for kw in sorted(class_map.keys(), key=lambda s: -len(s)):
                    if nl == kw or nl.startswith(kw + " ") or nl == kw:
                        cls_key = class_map[kw]
                        break
                if cls_key is None:
                    # Fallback: single-letter class labels (F/E/D/C/B/A)
                    if nl in ("f", "e", "d", "c", "b", "a", "seed"):
                        cls_key = class_map.get(nl)
                if cls_key is None:
                    continue
            if cls_key == "common":
                # Truth tracks both `common.dollars` and (for some
                # exits) implicit overlap with our_fund secondary.
                # Conservative: only emit `common`.
                pass
            # Emit dollars (only for class rows)
            if cls_key is not None:
                for col, exit_m in proc_cols.items():
                    d = _to_decimal(ws.cell(r, col).value)
                    if d is None:
                        continue
                    key = f"waterfall.{exit_m}M.{cls_key}.dollars"
                    _set_if_missing(values, provenance, key, float(d),
                                     sheet=sn, cell=_cell_a1(r, col),
                                     raw=ws.cell(r, col).value,
                                     how=f"codex_waterfall:{cls_key}@{exit_m}M")
                # Emit elections from a parallel column block
                for col, exit_m in elec_cols.items():
                    ev = ws.cell(r, col).value
                    canon = _normalize_election(ev)
                    if canon is None:
                        continue
                    key = f"waterfall.{exit_m}M.{cls_key}.election"
                    _set_if_missing(values, provenance, key, canon,
                                     sheet=sn, cell=_cell_a1(r, col),
                                     raw=ev,
                                     how=f"codex_waterfall_elec:{cls_key}@{exit_m}M")
            # If row label is "Total", emit waterfall.{exit}M.total_distributed
            if is_total:
                for col, exit_m in proc_cols.items():
                    d = _to_decimal(ws.cell(r, col).value)
                    if d is None:
                        continue
                    key = f"waterfall.{exit_m}M.total_distributed"
                    _set_if_missing(values, provenance, key, float(d),
                                     sheet=sn, cell=_cell_a1(r, col),
                                     raw=ws.cell(r, col).value,
                                     how=f"codex_waterfall:total@{exit_m}M")
        return  # one waterfall per workbook


def _extract_codex_sensitivity(wb, provenance: dict, values: dict) -> None:
    """Read the Codex Sensitivity stacked-grid layout.

    Codex stacks three sub-tables vertically on the Sensitivity sheet:
       Sub-table 1 (rows 4-7):   "Ownership % / proceeds at $2,000M exit"
                                 — formatted strings like "0.806595% / $17.0M"
       Sub-table 2 (rows 10-14): "Ownership %"          (numeric our_pct)
       Sub-table 3 (rows 16-20): "Proceeds"             (numeric dollars at 2000M)

    Each sub-table has the same axes:
       Col A: ESOP fraction (0.11, 0.13, 0.15)
       Cols B-D (header row above): pre-money 1200M / 1400M / 1800M

    We use sub-tables 2 and 3 (numeric) and skip sub-table 1 (formatted
    strings would require parsing).
    """
    sheets = _fuzzy_sheet_match(wb, ("sensitivity", "sensitivities"))
    if not sheets:
        return

    def _esop_token(esop_decimal: float) -> str | None:
        # Match 0.11/0.13/0.15 → "esop_11" / "esop_13" / "esop_15"
        for pct in (11, 13, 15):
            if abs(esop_decimal * 100 - pct) < 0.5:
                return f"esop_{pct}"
        return None

    def _pre_token(pre_usd: float) -> str | None:
        for pre in (1200, 1400, 1800):
            if abs(pre_usd / 1_000_000 - pre) < 1:
                return f"pre_{pre}M"
        return None

    for sn in sheets:
        ws = wb[sn]
        rmax = min(ws.max_row or 0, 60)
        cmax = min(ws.max_column or 0, 10)

        # Find each sub-table's section header (text in col A like
        # "Ownership %" or "Proceeds"), then the immediately-following
        # header row (with pre-money values in B/C/D) and the 3 data rows.
        section_label_to_kind = {
            "ownership %": "our_pct",
            "ownership%": "our_pct",
            "proceeds": "dollars_at_2000M_exit",
        }
        for r in range(1, rmax + 1):
            v = ws.cell(r, 1).value
            if not isinstance(v, str):
                continue
            nv = _norm(v)
            kind: str | None = None
            for sl, k in section_label_to_kind.items():
                if nv == sl or nv.startswith(sl):
                    kind = k
                    break
            if kind is None:
                continue
            # The header row for this sub-table: scan rows below for a
            # row whose B/C/D are numeric pre-money values.
            pre_header_row: int | None = None
            pre_cols: dict[int, int] = {}
            for hr in range(r + 1, min(rmax, r + 4) + 1):
                cm: dict[int, int] = {}
                for c in range(2, min(cmax, 6) + 1):
                    cv = ws.cell(hr, c).value
                    if isinstance(cv, (int, float)) and not isinstance(cv, bool):
                        fv = float(cv)
                        if 1_000_000_000 <= fv <= 5_000_000_000:
                            cm[c] = int(round(fv))
                if len(cm) >= 2:
                    pre_header_row = hr
                    pre_cols = cm
                    break
            if pre_header_row is None:
                continue
            # Walk 3 data rows below; col A has ESOP fraction
            for dr in range(pre_header_row + 1, min(rmax, pre_header_row + 6) + 1):
                a_val = ws.cell(dr, 1).value
                if not isinstance(a_val, (int, float)) or isinstance(a_val, bool):
                    continue
                esop_tok = _esop_token(float(a_val))
                if esop_tok is None:
                    continue
                for col, pre_usd in pre_cols.items():
                    pre_tok = _pre_token(float(pre_usd))
                    if pre_tok is None:
                        continue
                    cv = ws.cell(dr, col).value
                    d = _to_decimal(cv)
                    if d is None:
                        continue
                    if kind == "our_pct":
                        val = float(_normalize_pct(d))
                    else:
                        val = float(d)
                    key = f"sensitivity.{esop_tok}.{pre_tok}.{kind}"
                    _set_if_missing(values, provenance, key, val,
                                     sheet=sn, cell=_cell_a1(dr, col),
                                     raw=cv,
                                     how=f"codex_sensitivity:{esop_tok}/{pre_tok}/{kind}")


# ----- Stakeholder-table ownership fallback (v4) --------------------------


# Map from stakeholder ID/name patterns to ownership.* categories. Used
# when the per-class ownership extractor finds a per-stakeholder cap
# table (one row per F1/E1/SEED1/A1/...) instead of class roll-up rows.
# Each entry: (category, list of regex patterns matched against the
# normalized stakeholder label). First matching pattern wins per row,
# but ORDER matters — most-specific patterns (e.g. options_unissued
# before options_issued; series_e *_inv before bare employees) come
# first.
_STAKEHOLDER_PATTERNS: list[tuple[str, list[str]]] = [
    # Most specific first.
    ("options_unissued", [r"unissued option", r"unallocated",
                           r"option pool", r"esop pool", r"^pool\b",
                           r"^esop$"]),
    ("options_issued", [r"issued option", r"vested.*esop",
                         r"vested option"]),
    ("series_e", [r"^e\d+_inv\b", r"\bseries e\b", r"e pref",
                   r"e preferred", r"\bseries_e\b"]),
    ("employees", [
        r"^e\d+\b", r"employee rsa", r"employee", r"\brsa\b",
    ]),
    ("founders", [
        r"^f1\b", r"^f2\b", r"^f3\b",
        r"\bceo\b", r"\bcto\b", r"\bchief sci\b", r"founder",
    ]),
    ("seed", [r"^seed1\b", r"^seed\b", r"seed pref"]),
    ("series_a", [r"^a1\b", r"\bseries a\b", r"^a$", r"a pref"]),
    ("series_b", [r"^b1\b", r"\bseries b\b", r"^b$", r"b pref"]),
    ("series_c", [r"^c1\b", r"\bseries c\b", r"^c$", r"c pref"]),
    ("series_d", [r"^d1\b", r"\bseries d\b", r"^d$", r"d pref"]),
    ("note_a", [r"^na1\b", r"note a holder", r"\bnote a\b"]),
    ("note_b", [r"^nb1\b", r"note b holder", r"\bnote b\b"]),
    # F_LEAD secondary (common from tender) — must come before plain
    # F_LEAD / series_f_new because F_LEAD also appears in series_f_new
    # rows (primary preferred).
    ("f_lead_secondary", [r"secondary", r"f lead.*tender",
                           r"tender.*f lead", r"f_lead.*common",
                           r"common.*from tender",
                           r"f lead.*common.*tender"]),
    ("our_fund", [r"our fund", r"^f.fund\b", r"\bf fund\b",
                   r"our \$10m"]),
    ("series_f_new", [r"f primary", r"f preferred", r"\bseries f\b",
                       r"^f.lead\b", r"\bf lead\b",
                       r"\bseries_f\b", r"^series f preferred"]),
]


def _classify_stakeholder(label: str) -> str | None:
    """Return the ownership category for a per-stakeholder cap-table
    row, or None if no pattern matches.

    `label` may be a single column or a concatenated set of columns
    (e.g. ID + 'Role / Class' descriptor) — concatenation is the
    caller's responsibility."""
    n = _norm(label)
    if not n:
        return None
    for cat, patterns in _STAKEHOLDER_PATTERNS:
        for p in patterns:
            if re.search(p, n):
                return cat
    return None


def _extract_ownership_v4(wb, provenance: dict, values: dict) -> None:
    """Per-stakeholder ownership fallback: when a Pro-Forma sheet lists
    one row per stakeholder (F1, F2, ..., E10_INV, NA1, NB1, F_LEAD,
    F_FUND, ...) instead of class roll-up rows, sum each stakeholder's
    ownership % into the right category. Fills only keys not already
    populated by `_extract_ownership` (which prefers class-level rows).

    Special-handling: F_LEAD typically shows up in TWO rows in v4
    workbooks — primary (preferred) and secondary (common from tender).
    The classifier discriminates by 'secondary' / 'tender' substring. A
    bare F_LEAD row (no secondary marker) is treated as
    series_f_new.
    """
    sheet_names = _fuzzy_sheet_match(wb, (
        "pro-forma", "proforma", "pro forma", "pro_forma",
        "cap table", "captable",
    ))
    if not sheet_names:
        return

    # Find the categories the truth requires that are STILL missing.
    # Skip the whole strategy if all already filled — saves a redundant
    # scan on engines that produced clean class-level rollups.
    desired = {
        "founders", "employees", "options_issued", "options_unissued",
        "seed", "series_a", "series_b", "series_c", "series_d",
        "series_e", "series_f_new", "note_a", "note_b",
        "f_lead_secondary", "our_fund",
    }
    missing_cats = {c for c in desired
                    if values.get(f"ownership.{c}.pct") is None}
    if not missing_cats:
        return

    for sn in sheet_names:
        ws = wb[sn]
        rmax = min(ws.max_row or 0, 250)
        cmax = min(ws.max_column or 0, 25)

        # Locate a stakeholder-level cap table: header row containing
        # "Stakeholder" / "Holder" / "Name" + "Ownership" / "%" / "FD".
        header_row = None
        pct_col = None
        shares_col = None
        label_col = None
        for hr in range(1, rmax + 1):
            headers: dict[int, str] = {}
            for c in range(1, cmax + 1):
                v = ws.cell(hr, c).value
                if isinstance(v, str) and v.strip():
                    headers[c] = _norm(v)
            joined = " ".join(headers.values())
            if not joined:
                continue
            has_stakeholder = ("stakeholder" in joined or "holder" in joined
                                or "name" in joined)
            has_pct = ("ownership" in joined or "%" in joined or "pct" in joined
                       or "fully diluted" in joined or "fd %" in joined or "fd%" in joined)
            if not (has_stakeholder and has_pct):
                continue

            # Pick columns. PREFER explicit pct-like headers (% / pct /
            # ownership / percent) over bare "fd" — many layouts have BOTH
            # a "Total FD" share-count column AND an "Ownership %" column
            # side-by-side; the share column must NOT be picked as pct.
            for c, t in headers.items():
                if label_col is None and ("stakeholder" in t or "holder" in t
                                           or t == "name"):
                    label_col = c
                if pct_col is None and ("%" in t or "pct" in t
                                         or "ownership" in t
                                         or "percent" in t):
                    pct_col = c
                if shares_col is None and t in ("shares", "total shares",
                                                  "total fd shares"):
                    shares_col = c
            # Fallback: only use bare "fd" / "fully diluted" columns when
            # no explicit pct-like header was found anywhere.
            if pct_col is None:
                for c, t in headers.items():
                    if ("fully diluted" in t or "fd" in t) and "shares" not in t:
                        pct_col = c
                        break
            if label_col is None:
                label_col = min(headers.keys())
            if pct_col is None:
                continue
            header_row = hr
            break

        if header_row is None:
            continue

        # Walk data rows; sum %-by-category.
        per_cat_pct: dict[str, list[tuple[Decimal, int, int]]] = {}
        per_cat_label: dict[str, list[str]] = {}
        for r in range(header_row + 1, min(rmax, header_row + 80) + 1):
            label_val = ws.cell(r, label_col).value
            if not isinstance(label_val, str) or not label_val.strip():
                # Allow stragglers but break on a cluster of blanks.
                all_blank = all(
                    ws.cell(r, c).value is None
                    for c in range(1, min(ws.max_column or 0, 10) + 1)
                )
                if all_blank and per_cat_pct:
                    break
                continue
            # Concatenate label with adjacent text columns ('Role /
            # Class' column on Tetra; 'Security' column on Shortcut)
            # before pattern-matching. Tetra encodes "F1 | Founder
            # (CEO)" / "E1 | Employee RSA" / "F_LEAD | F lead investor"
            # / "F_FUND | Our fund (F primary)" — the role column is
            # what discriminates a bare F-letter from F_LEAD's role.
            concat_parts = [label_val]
            for cc in range(label_col + 1, min(ws.max_column or 0, label_col + 6) + 1):
                if cc == pct_col or cc == shares_col:
                    continue
                v = ws.cell(r, cc).value
                if isinstance(v, str) and v.strip():
                    concat_parts.append(v)
            label_concat = " | ".join(concat_parts)
            n = _norm(label_concat)
            # Stop at totals / check rows
            if (n.startswith("total") or n == "check"
                    or ("total" in n and any(t in n for t in ("fd", "fully",
                                                               "post-f", "post f",
                                                               "post_f")))):
                break
            cat = _classify_stakeholder(label_concat)
            if cat is None:
                continue
            v = _to_decimal(ws.cell(r, pct_col).value)
            if v is None:
                continue
            per_cat_pct.setdefault(cat, []).append((v, r, pct_col))
            per_cat_label.setdefault(cat, []).append(label_concat)

        if not per_cat_pct:
            continue

        # Emit only for categories not already in values.
        for cat, rows in per_cat_pct.items():
            key = f"ownership.{cat}.pct"
            if values.get(key) is not None:
                continue
            total_pct = sum((v[0] for v in rows), Decimal(0))
            total_pct = _normalize_pct(total_pct)
            cells = "+".join(_cell_a1(r, c) for _v, r, c in rows)
            labels = ", ".join(per_cat_label.get(cat, [])[:4])
            values[key] = float(total_pct)
            provenance[key] = CellRef(
                sheet=sn, cell=cells,
                raw_value=float(total_pct),
                how_matched=f"ownership_v4:{cat} stakeholder-table sum (rows={len(rows)}: {labels[:80]})",
            )
        # First viable sheet wins.
        return


def _extract_ownership_v4_derived(wb, provenance: dict, values: dict) -> None:
    """Compute-derived ownership keys (last-resort fallback).

    When the stakeholder-level table extractor can't separate F_LEAD's
    primary-preferred shares from its secondary-common (tender) shares
    — e.g. a workbook with one F_LEAD row and a multi-column split
    "F Primary | F Secondary" we don't currently parse — fall back to
    deriving the missing keys from values we DID extract:

      ownership.f_lead_secondary.pct
        = tender.f_lead_holding_secondary_shares / round.post_f_fd_shares

      ownership.series_f_new.pct
        = (round.our_fund_shares + round.f_lead_primary_shares)
          / round.post_f_fd_shares

    These identities follow directly from spec §12.5 (round solution).
    Skipped for any key already present.
    """
    T_F = _to_decimal(values.get("round.post_f_fd_shares"))
    if T_F is None or T_F <= 0:
        return

    # f_lead_secondary
    fls_key = "ownership.f_lead_secondary.pct"
    if values.get(fls_key) is None:
        fls_shares = _to_decimal(values.get("tender.f_lead_holding_secondary_shares"))
        if fls_shares is None:
            fls_shares = _to_decimal(values.get("tender.shares_transferred"))
        if fls_shares is not None and fls_shares > 0:
            pct = float(fls_shares / T_F)
            values[fls_key] = pct
            provenance[fls_key] = CellRef(
                sheet="(computed)",
                cell="tender.f_lead_holding_secondary_shares / round.post_f_fd_shares",
                raw_value=pct,
                how_matched="ownership_v4_derived:f_lead_secondary = tender_secondary / T_F",
            )

    # series_f_new — if the table-level extractor produced a value but the
    # row is merged (Tetra's "F_LEAD" row commingles primary preferred + the
    # secondary common in a single Ownership % cell), the recorded value
    # will be inflated by the f_lead_secondary contribution. Detect this
    # by checking whether the recorded value is materially larger than the
    # round-derived (our_fund + f_lead_primary) / T_F identity. If so,
    # overwrite — the round identity is exact by construction.
    sfn_key = "ownership.series_f_new.pct"
    our = _to_decimal(values.get("round.our_fund_shares")) or Decimal(0)
    flp = _to_decimal(values.get("round.f_lead_primary_shares")) or Decimal(0)
    # Both inputs must be present for the round identity to be exact.
    # If only one is present, the sum would be artificially low and would
    # incorrectly override a perfectly-good extracted value.
    have_inputs = (our > 0 and flp > 0)
    if values.get(sfn_key) is None and have_inputs:
        pct = float((our + flp) / T_F)
        values[sfn_key] = pct
        provenance[sfn_key] = CellRef(
            sheet="(computed)",
            cell="(round.our_fund_shares + round.f_lead_primary_shares) / round.post_f_fd_shares",
            raw_value=pct,
            how_matched="ownership_v4_derived:series_f_new = (our + f_lead_primary) / T_F",
        )
    elif values.get(sfn_key) is not None and have_inputs:
        # Merged-row override: if extracted is >15% larger than the
        # round-derived value, the row commingled primary+secondary.
        derived = float((our + flp) / T_F)
        extracted = float(values[sfn_key] or 0)
        if extracted > 0 and derived > 0 and (extracted - derived) / derived > 0.15:
            values[sfn_key] = derived
            provenance[sfn_key] = CellRef(
                sheet="(computed)",
                cell="(round.our_fund_shares + round.f_lead_primary_shares) / round.post_f_fd_shares",
                raw_value=derived,
                how_matched=(
                    "ownership_v4_derived:series_f_new override "
                    "(table value commingled primary+secondary; "
                    f"replaced {extracted:.4f} with {derived:.4f})"
                ),
            )


# ---------------------------------------------------------------------------
# Orchestrator
# ---------------------------------------------------------------------------


def extract(
    *,
    xlsx_path: Path,
    spec_md_path: Path | None = None,
    truth_json_path: Path | None = None,
) -> ExtractResult:
    """
    Extract tracked values from an arbitrary cap-table-shaped xlsx.

    spec_md_path is unused at runtime; the canonical key list is normally
    baked into iter_tracked_keys() (a v1 superset covering Nyxlight/Paxos).

    truth_json_path, when provided, overrides the hardcoded key list with
    the keys present in that truth.json file (kind inferred via
    _infer_kind_from_key). This is required for v3/v4+ scenarios whose
    schemas (notes.*, tender.*, round.*, antidilution.series_e.*, ...)
    aren't in the v1 superset. Callers that don't pass truth_json_path
    keep the legacy v1 behavior unchanged.
    """
    values: dict[str, Any] = {}
    provenance: dict[str, CellRef] = {}

    if load_workbook is None:
        raise RuntimeError("openpyxl not installed")

    xlsx_path = Path(xlsx_path)
    if not xlsx_path.exists():
        raise FileNotFoundError(str(xlsx_path))

    # Resolve the tracked-key set up-front. Falls back to the hardcoded v1
    # superset if truth_json_path is None or unreadable.
    tracked_override: list[dict[str, Any]] | None = None
    if truth_json_path is not None:
        tracked_override = _tracked_keys_from_truth(Path(truth_json_path))
    tracked = tracked_override if tracked_override is not None else iter_tracked_keys()

    # Many API outputs (notably codex) emit pure-formula workbooks with no
    # cached values — openpyxl's data_only=True returns None everywhere. Force
    # a LibreOffice headless recalc up-front so the data actually surfaces.
    # If recalc fails (LibreOffice not installed, workbook malformed) we fall
    # back to the original path — extraction may still recover some cells.
    xlsx_for_read = _maybe_libreoffice_recalc(xlsx_path)

    # Use data_only=True: we want the cached numeric output from any formulas.
    wb = load_workbook(str(xlsx_for_read), data_only=True)

    # --- Detect per-sheet unit declarations (millions vs raw) ---
    # Some workbooks present share counts and dollar amounts in millions
    # (e.g. founder row showing "4" meaning 4M shares). Scaling is applied
    # per-sheet at each emit point in the pipeline below.
    try:
        wb._unit_scales = _detect_workbook_units(wb)
    except Exception:  # noqa: BLE001
        wb._unit_scales = {}

    # --- Strategy 1: defined names (fast path) ---
    _apply_defined_names(wb, values, provenance, tracked_keys=tracked)

    # --- Strategy 2: semantic/keyword extraction per region ---
    # Each strategy is wrapped to contain failures to its own region —
    # we would rather mark a subset missing than abort the whole extract.
    # Set EXTRACT_DEBUG=1 to re-raise and see tracebacks.
    import os as _os
    debug = _os.environ.get("EXTRACT_DEBUG") == "1"

    for name, fn in (
        ("ownership", _extract_ownership),
        ("safes", _extract_safes),
        ("venture_debt", _extract_venture_debt),
        ("antidilution", _extract_antidilution),
        ("waterfall", _extract_waterfall),
        ("sensitivity", _extract_sensitivity),
        ("solver", _extract_solver),
        # v4 extension extractors. Run AFTER the v0..v3 strategies so
        # they only fill in keys those couldn't resolve. Each extractor
        # is a no-op on v0..v3 truth schemas (the keys it emits don't
        # appear in earlier scenarios' tracked-key sets), so this is
        # safe across scenarios.
        ("notes", _extract_notes),
        ("tender", _extract_tender),
        ("round", _extract_round),
        ("series_e_ad", _extract_series_e_ad),
        ("v4_waterfall_extras", _extract_v4_waterfall_extras),
        ("v4_solver", _extract_v4_solver),
        ("ownership_v4", _extract_ownership_v4),
        ("ownership_v4_derived", _extract_ownership_v4_derived),
        # Codex-specific layout reader. Runs LAST so it only fills keys
        # that earlier passes left as None. Most cells it targets are at
        # canonical Codex addresses (Pro-Forma scalar summary, AD Block,
        # Notes Block, iter ladder row 81, compact Waterfall, stacked
        # Sensitivity grid, long-form Return-Solver). No-op on workbooks
        # without those layout markers.
        ("codex_layout", _extract_codex_layout),
    ):
        try:
            fn(wb, provenance, values)
        except Exception as e:
            if debug:
                import traceback
                traceback.print_exc()
                print(f"[extract] {name} strategy raised: {e}", file=sys.stderr)

    # Compute missing list — over the tracked-key set we resolved above
    # (override or v1 superset).
    missing: list[str] = []
    for k in tracked:
        if k["key"] not in values or values[k["key"]] is None:
            missing.append(k["key"])
    ambiguous: list[tuple[str, list[CellRef]]] = []

    return ExtractResult(
        values=values,
        provenance=provenance,
        missing=missing,
        ambiguous=ambiguous,
    )


def _apply_defined_names(
    wb,
    values: dict,
    provenance: dict,
    tracked_keys: list[dict[str, Any]] | None = None,
) -> None:
    keys_iter = tracked_keys if tracked_keys is not None else iter_tracked_keys()
    for k in keys_iter:
        key = k["key"]
        named = key.replace(".", "_")
        val, sheet, coord = _resolve_defined_name(wb, named)
        if val is None:
            continue
        # For pct, try to normalize
        kind = k["kind"]
        # PPS keys (per-share dollar prices) are scale-invariant; bare USD keys
        # ARE scaled per the source sheet's unit declaration.
        is_pps = (
            ("conversion" in key and ("price" in key or "pps" in key))
            or key.endswith("_pps_usd")
            or key.endswith(".f_pps_usd")
        )
        scale_how = "defined_name"
        if kind == "pct":
            d = _to_decimal(val)
            if d is not None:
                nd = _normalize_pct(d)
                values[key] = float(nd)
            else:
                values[key] = val
        elif kind == "ratio":
            # AD ratios are unit-less and bounded in [1.0, ∞) for triggered
            # classes. NEVER /100-normalize — that breaks values like 1.04
            # (Series D with material AD adjustment).
            d = _to_decimal(val)
            if d is not None:
                values[key] = float(d)
            else:
                values[key] = val
        elif kind == "shares":
            d = _to_decimal(val)
            if d is not None:
                mult = _shares_mult(wb, sheet)
                values[key] = int(round(float(d) * mult))
                if mult != 1:
                    scale_how = (
                        f"defined_name ×{mult} "
                        f"(units:{_unit_scales(wb)[sheet]['shares_reason']})"
                    )
            else:
                values[key] = val
        elif kind in ("usd", "usd_abs_1"):
            d = _to_decimal(val)
            if d is not None:
                if is_pps:
                    values[key] = float(d)
                else:
                    mult = _dollars_mult(wb, sheet)
                    values[key] = float(d) * mult
                    if mult != 1:
                        scale_how = (
                            f"defined_name ×{mult} "
                            f"(units:{_unit_scales(wb)[sheet]['dollars_reason']})"
                        )
            else:
                values[key] = val
        elif kind == "election":
            canon = _normalize_election(val)
            values[key] = canon if canon is not None else val
        elif kind == "binding":
            canon = _normalize_binding(val)
            values[key] = canon if canon is not None else val
        else:
            values[key] = val
        provenance[key] = CellRef(
            sheet=sheet or "?",
            cell=coord or "?",
            raw_value=val,
            how_matched=scale_how,
        )


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def _cli(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Layout-agnostic extractor for §11 tracked keys.")
    p.add_argument("--xlsx", required=True, help="Path to input xlsx")
    p.add_argument("--spec", default=None, help="Optional path to spec.md (informational only)")
    p.add_argument("--truth", default=None, help="Optional path to truth.json — derives the scenario-specific key set instead of the v1 superset")
    p.add_argument("--json-out", default=None, help="Optional: write full extraction JSON here")
    args = p.parse_args(argv)

    result = extract(
        xlsx_path=Path(args.xlsx),
        spec_md_path=Path(args.spec) if args.spec else None,
        truth_json_path=Path(args.truth) if args.truth else None,
    )

    # Total reflects the same key set extract() used (override or v1 superset).
    if args.truth:
        truth_keys = _tracked_keys_from_truth(Path(args.truth))
        total = len(truth_keys) if truth_keys is not None else len(iter_tracked_keys())
    else:
        total = len(iter_tracked_keys())
    got = total - len(result.missing)
    print(f"extracted {got}/{total} keys  |  missing={len(result.missing)}  |  ambiguous={len(result.ambiguous)}")

    # List missing keys (grouped)
    if result.missing:
        print("\nMissing keys:")
        for k in result.missing:
            print(f"  - {k}")

    if args.json_out:
        out = result.to_dict()
        Path(args.json_out).parent.mkdir(parents=True, exist_ok=True)
        Path(args.json_out).write_text(json.dumps(out, indent=2, default=str))
        print(f"\nWrote {args.json_out}")
    return 0


if __name__ == "__main__":
    sys.exit(_cli())
