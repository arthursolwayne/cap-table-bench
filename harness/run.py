"""
Benchmark orchestrator — one entry point to run a scenario against
Tetra, Shortcut, or both; capture all outputs; and invoke scoring.

Usage:
    python -m harness.run --scenario nyxlight_v1 --api both --replicates 3
    python harness/run.py --scenario nyxlight_v1 --api tetra --tetra-sas-url ...

Design notes:
- This file is the ONLY orchestrator. It owns pre-flight, run planning,
  exception handling, scoring invocation, and summary rendering.
- Sibling modules (shortcut_client, tetra_client, score, reconcile) are
  imported via try/except so this file is still syntactically valid and
  CLI-usable for --help even if siblings aren't written yet. At actual
  run time the missing import surfaces as a clear error.
- "Bad ground truth is worse than no benchmark" — we refuse to run unless
  reconcile passes OR the user explicitly passes --skip-reconcile.
"""

from __future__ import annotations

# When this file is invoked as a script (``python harness/run.py``),
# Python prepends ``harness/`` to sys.path, which causes the local
# ``harness/types.py`` to shadow the stdlib ``types`` module and breaks
# every subsequent import. Fix that before importing anything else:
# drop the script-dir entry and insert the repo root so ``harness`` is
# importable as a package. ``python -m harness.run`` already does the
# right thing, but this keeps the naive invocation working too.
import os as _os
import sys as _sys
if __name__ == "__main__" and __package__ in (None, ""):
    _here = _os.path.dirname(_os.path.abspath(__file__))
    _repo_root = _os.path.dirname(_here)
    # Remove harness/ from sys.path so stdlib `types` wins.
    _sys.path[:] = [p for p in _sys.path if _os.path.abspath(p) != _here]
    if _repo_root not in _sys.path:
        _sys.path.insert(0, _repo_root)

import argparse
import dataclasses
import json
import logging
import os
import signal
import sys
import threading
import time
import traceback
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable

# --- Shared contract (must exist; this file is useless without it) ---
from harness.contracts import (  # noqa: E402
    RunEvent,
    RunResult,
    now_iso,
    mono_ns,
    write_events_jsonl,
    write_run_result,
)

# --- Sibling modules — imported lazily/defensively. ---
# These may not exist yet during parallel authoring. We resolve them at
# call time so --help works regardless.
try:  # pragma: no cover - import guard
    from harness import shortcut_client as _shortcut_client  # type: ignore
except Exception:  # noqa: BLE001
    _shortcut_client = None  # type: ignore[assignment]

try:  # pragma: no cover
    from harness import tetra_client as _tetra_client  # type: ignore
except Exception:  # noqa: BLE001
    _tetra_client = None  # type: ignore[assignment]

try:  # pragma: no cover
    from harness import score as _score_mod  # type: ignore
except Exception:  # noqa: BLE001
    _score_mod = None  # type: ignore[assignment]

try:  # pragma: no cover
    from harness import reconcile as _reconcile_mod  # type: ignore
except Exception:  # noqa: BLE001
    _reconcile_mod = None  # type: ignore[assignment]


# --- Repo layout ---
REPO_ROOT = Path(__file__).resolve().parent.parent
SCENARIOS_ROOT = REPO_ROOT / "scenarios"
HARNESS_ROOT = REPO_ROOT / "harness"
DEFAULT_OUTPUT_ROOT = REPO_ROOT / "runs"
# Prompt now lives under each scenario (scenarios/{scenario}/prompt.md) so
# parameters match the scenario's truth. Old global harness/prompt.md would
# silently inject v1 parameters into v0 runs and invalidate them.


# --- Logging ---
logging.basicConfig(
    level=logging.INFO,
    stream=sys.stderr,
    format="%(asctime)sZ %(levelname)s %(name)s: %(message)s",
    datefmt="%Y-%m-%dT%H:%M:%S",
)
logging.Formatter.converter = time.gmtime  # force UTC
log = logging.getLogger("harness.run")


# --- SIGINT handling ---
# We want graceful shutdown: finalize current run artifacts, write summary.
_shutdown = threading.Event()


def _install_sigint_handler() -> None:
    def _handler(signum: int, frame: Any) -> None:  # noqa: ARG001
        if _shutdown.is_set():
            # Second Ctrl-C → hard exit
            log.error("Second SIGINT — forcing exit.")
            sys.exit(130)
        log.warning("SIGINT received — finishing current run, then shutting down.")
        _shutdown.set()

    signal.signal(signal.SIGINT, _handler)


# --- .env loading ---
def _load_dotenv() -> None:
    """Load .env from repo root if python-dotenv is available. No-op otherwise."""
    env_path = REPO_ROOT / ".env"
    try:
        from dotenv import load_dotenv  # type: ignore

        if env_path.exists():
            load_dotenv(env_path)
            log.info("Loaded env from %s (python-dotenv).", env_path)
        else:
            log.info("No .env at %s; relying on process environment.", env_path)
    except Exception:  # noqa: BLE001
        # Minimal KEY=VALUE fallback so we don't hard-require python-dotenv.
        if env_path.exists():
            log.info("python-dotenv not installed; doing naive .env parse of %s.", env_path)
            for line in env_path.read_text().splitlines():
                s = line.strip()
                if not s or s.startswith("#") or "=" not in s:
                    continue
                k, _, v = s.partition("=")
                k = k.strip()
                v = v.strip().strip('"').strip("'")
                os.environ.setdefault(k, v)
        else:
            log.info("No python-dotenv, no .env file — using process environment only.")


# --- Pre-flight ---
@dataclasses.dataclass
class PreflightResult:
    ok: bool
    checks: list[tuple[str, bool, str]]   # (label, passed, detail)
    prompt_text: str | None
    paths: dict[str, Path]

    def render(self) -> str:
        lines = ["Pre-flight check:"]
        for label, passed, detail in self.checks:
            mark = "OK  " if passed else "FAIL"
            lines.append(f"  [{mark}] {label}" + (f"  — {detail}" if detail else ""))
        return "\n".join(lines)


def preflight(scenario: str, apis: list[str], tetra_sas_url: str | None) -> PreflightResult:
    checks: list[tuple[str, bool, str]] = []

    scenario_dir = SCENARIOS_ROOT / scenario
    paths = {
        "scenario_dir": scenario_dir,
        "input_xlsx": scenario_dir / "inputs" / "cap_table.xlsx",
        "truth_json": scenario_dir / "truth" / "truth.json",
        "ground_truth_xlsx": scenario_dir / "truth" / "ground_truth.xlsx",
        "spec_md": scenario_dir / "spec.md",
        "prompt_md": scenario_dir / "prompt.md",
    }

    # File existence checks
    for label, p in [
        ("scenario directory", paths["scenario_dir"]),
        ("cap_table.xlsx", paths["input_xlsx"]),
        ("truth.json", paths["truth_json"]),
        ("ground_truth.xlsx", paths["ground_truth_xlsx"]),
        ("spec.md", paths["spec_md"]),
        ("scenario prompt.md", paths["prompt_md"]),
    ]:
        exists = p.exists()
        checks.append((f"{label} exists", exists, str(p)))

    # Env keys
    shortcut_key = os.environ.get("SHORTCUT_API_KEY")
    tetra_key = os.environ.get("TETRA_API_KEY")
    if "shortcut" in apis:
        checks.append(("SHORTCUT_API_KEY set", bool(shortcut_key), "env"))
    # All DealGlass engines route through tetra_client and share the same
    # key + path requirement.
    dg_engines = {"tetra", "opus-4.7", "gpt-5.5"} & set(apis)
    if dg_engines:
        checks.append(("TETRA_API_KEY set", bool(tetra_key),
                       f"shared across engines: {sorted(dg_engines)}"))
        checks.append(
            ("--tetra-sas-url provided", bool(tetra_sas_url),
             f"required for engines: {sorted(dg_engines)}"),
        )

    # Prompt load
    prompt_text: str | None = None
    if paths["prompt_md"].exists():
        try:
            prompt_text = paths["prompt_md"].read_text(encoding="utf-8")
            checks.append(("scenario prompt.md readable",
                           bool(prompt_text.strip()),
                           f"{len(prompt_text)} chars"))
        except Exception as e:  # noqa: BLE001
            checks.append(("scenario prompt.md readable", False, str(e)))

    # Sibling module availability (only warn; caller decides severity)
    if "shortcut" in apis:
        checks.append(("harness.shortcut_client import",
                       _shortcut_client is not None,
                       "" if _shortcut_client else "module missing"))
    if dg_engines:
        checks.append(("harness.tetra_client import",
                       _tetra_client is not None,
                       "" if _tetra_client else "module missing"))
    checks.append(("harness.score import",
                   _score_mod is not None,
                   "" if _score_mod else "module missing"))
    checks.append(("harness.reconcile import",
                   _reconcile_mod is not None,
                   "" if _reconcile_mod else "module missing"))

    ok = all(passed for _, passed, _ in checks)
    return PreflightResult(ok=ok, checks=checks, prompt_text=prompt_text, paths=paths)


# --- Reconcile step ---
def run_reconcile(paths: dict[str, Path], report_path: Path) -> bool:
    """Invoke harness.reconcile. Returns True iff reconcile passes."""
    if _reconcile_mod is None:
        log.error("Cannot reconcile: harness.reconcile module is missing. "
                  "Re-run with --skip-reconcile to override.")
        return False
    try:
        result = _reconcile_mod.reconcile(
            truth_json_path=paths["truth_json"],
            ground_truth_xlsx_path=paths["ground_truth_xlsx"],
            spec_md_path=paths["spec_md"],
            report_path=report_path,
        )
    except Exception as e:  # noqa: BLE001
        log.error("Reconcile raised an exception: %s", e)
        log.debug("%s", traceback.format_exc())
        return False

    # Duck-type the result: prefer .ok attribute; fall back to truthiness.
    ok = bool(getattr(result, "ok", result))
    if ok:
        log.info("Reconcile PASSED. Report: %s", report_path)
    else:
        diffs = getattr(result, "differences", None)
        log.error("Reconcile FAILED. Report: %s. Differences: %s", report_path, diffs)
    return ok


# --- Run ID ---
def make_run_id(api: str, scenario: str, replicate: int, when: datetime) -> str:
    """
    {ISO_timestamp}_{api}_{scenario}_r{n}

    We use compact ISO (no colons) for filesystem safety, with explicit Z.
    Example: 2026-04-15T12-34-56Z_tetra_nyxlight_v1_r1
    """
    ts = when.astimezone(timezone.utc).strftime("%Y-%m-%dT%H-%M-%SZ")
    return f"{ts}_{api}_{scenario}_r{replicate}"


def make_batch_timestamp(when: datetime) -> str:
    return when.astimezone(timezone.utc).strftime("%Y-%m-%dT%H-%M-%SZ")


# --- Per-run execution ---
@dataclasses.dataclass
class RunOutcome:
    api: str
    replicate: int
    run_id: str
    run_dir: Path
    result: RunResult | None
    score: Any | None
    score_md_path: Path | None
    error: str | None
    wall_clock_s: float | None


def _synth_failure_result(
    *, api: str, scenario: str, replicate: int, run_id: str, error: str,
) -> RunResult:
    """Build a minimal RunResult for a failure outside the client call."""
    start_ev = RunEvent.make("run_start", api=api, scenario=scenario, replicate=replicate)
    err_ev = RunEvent.make("error", message=error)
    end_ev = RunEvent.make("run_end")
    return RunResult(
        run_id=run_id,
        api=api,  # type: ignore[arg-type]
        scenario=scenario,
        replicate=replicate,
        submitted_at_iso=start_ev.t_wall_iso,
        completed_at_iso=end_ev.t_wall_iso,
        api_run_id=None,
        events=[start_ev, err_ev, end_ev],
        output_xlsx_path=None,
        credits_used=None,
        error=error,
    )


def _persist_result(result: RunResult, run_dir: Path) -> tuple[Path, Path]:
    """Write events.jsonl and result.json to run_dir. Returns (events_path, result_path)."""
    run_dir.mkdir(parents=True, exist_ok=True)
    events_path = run_dir / "events.jsonl"
    result_path = run_dir / "result.json"
    # events.jsonl: append all events (client may have already written; we rewrite
    # canonically here from the returned RunResult to guarantee completeness).
    if events_path.exists():
        events_path.unlink()
    write_events_jsonl(result.events, events_path)
    write_run_result(result, result_path)
    return events_path, result_path


def _invoke_client(
    *,
    api: str,
    scenario: str,
    replicate: int,
    run_dir: Path,
    prompt_text: str,
    input_xlsx_path: Path,
    tetra_sas_url: str | None,
    shortcut_api_key: str | None,
    tetra_api_key: str | None,
) -> RunResult:
    """Dispatch to the appropriate client. Raises on missing module / config."""
    if api == "shortcut":
        if _shortcut_client is None:
            raise RuntimeError("harness.shortcut_client not importable")
        if not shortcut_api_key:
            raise RuntimeError("SHORTCUT_API_KEY not set")
        return _shortcut_client.run(
            prompt_text=prompt_text,
            input_xlsx_path=input_xlsx_path,
            output_dir=run_dir,
            scenario=scenario,
            replicate=replicate,
            api_key=shortcut_api_key,
        )
    if api in ("tetra", "opus-4.7", "gpt-5.5"):
        # All three DealGlass engines share the same /xlsx transport.
        # tetra_client.run() accepts engine=... and routes accordingly.
        if _tetra_client is None:
            raise RuntimeError("harness.tetra_client not importable")
        if not tetra_api_key:
            raise RuntimeError("TETRA_API_KEY not set")
        if not tetra_sas_url:
            raise RuntimeError(f"--tetra-sas-url required for {api} runs")
        return _tetra_client.run(
            prompt_text=prompt_text,
            azure_sas_url=tetra_sas_url,
            output_dir=run_dir,
            scenario=scenario,
            replicate=replicate,
            api_key=tetra_api_key,
            engine=api,
        )
    raise ValueError(f"unknown api: {api!r}")


def _invoke_scorer(
    *, output_xlsx_path: Path | None, truth_json_path: Path,
    run_result_path: Path, score_md_path: Path,
) -> Any | None:
    """Invoke score_run. Tolerant of a missing output xlsx (failed run) —
    the scorer is expected to return a zero-score result in that case."""
    if _score_mod is None:
        log.warning("harness.score module missing — skipping scoring for this run.")
        return None
    try:
        return _score_mod.score_run(
            output_xlsx_path=output_xlsx_path if output_xlsx_path else Path("/nonexistent"),
            truth_json_path=truth_json_path,
            run_result_path=run_result_path,
            output_report_path=score_md_path,
        )
    except Exception as e:  # noqa: BLE001
        log.error("Scorer raised: %s", e)
        log.debug("%s", traceback.format_exc())
        return None


def execute_run(
    *,
    api: str,
    scenario: str,
    replicate: int,
    batch_dir: Path,
    prompt_text: str,
    paths: dict[str, Path],
    tetra_sas_url: str | None,
) -> RunOutcome:
    when = datetime.now(timezone.utc)
    run_id = make_run_id(api, scenario, replicate, when)
    run_dir = batch_dir / api / f"replicate_{replicate}"
    run_dir.mkdir(parents=True, exist_ok=True)

    log.info("[%s r%d] START run_id=%s dir=%s", api, replicate, run_id, run_dir)

    shortcut_key = os.environ.get("SHORTCUT_API_KEY")
    tetra_key = os.environ.get("TETRA_API_KEY")

    result: RunResult | None = None
    error: str | None = None
    t0 = time.monotonic()
    try:
        result = _invoke_client(
            api=api,
            scenario=scenario,
            replicate=replicate,
            run_dir=run_dir,
            prompt_text=prompt_text,
            input_xlsx_path=paths["input_xlsx"],
            tetra_sas_url=tetra_sas_url,
            shortcut_api_key=shortcut_key,
            tetra_api_key=tetra_key,
        )
        # Client can also report an in-band error via result.error
        if result.error:
            error = result.error
            log.warning("[%s r%d] client returned error in-band: %s", api, replicate, error)
    except Exception as e:  # noqa: BLE001
        error = f"{type(e).__name__}: {e}"
        log.error("[%s r%d] client raised: %s", api, replicate, error)
        log.debug("%s", traceback.format_exc())
        result = _synth_failure_result(
            api=api, scenario=scenario, replicate=replicate,
            run_id=run_id, error=error,
        )
    finally:
        wall = time.monotonic() - t0

    # Ensure the result has the harness-assigned run_id (clients may or may not).
    if result is not None and result.run_id != run_id:
        log.info("[%s r%d] client returned run_id=%s; overriding with harness id=%s",
                 api, replicate, result.run_id, run_id)
        result.run_id = run_id

    # Persist artifacts even on failure.
    assert result is not None  # always set above
    events_path, result_path = _persist_result(result, run_dir)
    log.info("[%s r%d] artifacts: %s, %s", api, replicate, events_path.name, result_path.name)

    # Score regardless of success — a failed run scores as zero.
    score_md_path = run_dir / "score.md"
    output_xlsx = Path(result.output_xlsx_path) if result.output_xlsx_path else None
    score = _invoke_scorer(
        output_xlsx_path=output_xlsx,
        truth_json_path=paths["truth_json"],
        run_result_path=result_path,
        score_md_path=score_md_path,
    )

    outcome = RunOutcome(
        api=api,
        replicate=replicate,
        run_id=run_id,
        run_dir=run_dir,
        result=result,
        score=score,
        score_md_path=score_md_path if score_md_path.exists() else None,
        error=error,
        wall_clock_s=wall,
    )
    log.info("[%s r%d] END wall=%.2fs error=%s", api, replicate, wall, error)
    return outcome


# --- Summary ---
def _score_field(score: Any, name: str) -> float | None:
    """Tolerant accessor: dataclass attr or dict key."""
    if score is None:
        return None
    v = getattr(score, name, None)
    if v is None and isinstance(score, dict):
        v = score.get(name)
    try:
        return float(v) if v is not None else None
    except (TypeError, ValueError):
        return None


def _fmt(x: float | None, spec: str = ".4f") -> str:
    if x is None:
        return "—"
    return format(x, spec)


SUMMARY_HEADER = (
    "| api | replicate | composite | correctness_strict | correctness_vc_usable"
    " | correctness_directional | wall_clock_s | error |"
)
SUMMARY_DIVIDER = "|---|---|---|---|---|---|---|---|"


def render_summary(outcomes: list[RunOutcome], batch_dir: Path, scenario: str) -> str:
    lines: list[str] = []
    lines.append(f"# Benchmark summary — {scenario}")
    lines.append("")
    lines.append(f"- batch dir: `{batch_dir}`")
    lines.append(f"- generated: {now_iso()}")
    lines.append(f"- runs: {len(outcomes)}")
    lines.append("")
    lines.append(SUMMARY_HEADER)
    lines.append(SUMMARY_DIVIDER)
    for o in outcomes:
        composite = _score_field(o.score, "composite")
        c_strict = _score_field(o.score, "correctness_strict")
        c_vc = _score_field(o.score, "correctness_vc_usable")
        c_dir = _score_field(o.score, "correctness_directional")
        err = (o.error or "").replace("|", "\\|").replace("\n", " ")[:120]
        lines.append(
            f"| {o.api} | {o.replicate} | {_fmt(composite)} | {_fmt(c_strict)} "
            f"| {_fmt(c_vc)} | {_fmt(c_dir)} | {_fmt(o.wall_clock_s, '.2f')} | {err} |"
        )
    lines.append("")
    return "\n".join(lines)


def write_summary(outcomes: list[RunOutcome], batch_dir: Path, scenario: str) -> Path:
    md = render_summary(outcomes, batch_dir, scenario)
    path = batch_dir / "summary.md"
    path.write_text(md, encoding="utf-8")
    # Also drop a machine-readable version.
    json_path = batch_dir / "summary.json"
    with json_path.open("w", encoding="utf-8") as f:
        json.dump(
            [
                {
                    "api": o.api,
                    "replicate": o.replicate,
                    "run_id": o.run_id,
                    "run_dir": str(o.run_dir),
                    "wall_clock_s": o.wall_clock_s,
                    "error": o.error,
                    "composite": _score_field(o.score, "composite"),
                    "correctness_strict": _score_field(o.score, "correctness_strict"),
                    "correctness_vc_usable": _score_field(o.score, "correctness_vc_usable"),
                    "correctness_directional": _score_field(o.score, "correctness_directional"),
                }
                for o in outcomes
            ],
            f,
            indent=2,
        )
    return path


# --- Run plan expansion ---
def build_plan(apis: list[str], replicates: int) -> list[tuple[str, int]]:
    """Return list of (api, replicate) pairs. Sequential order, api-major."""
    plan: list[tuple[str, int]] = []
    for api in apis:
        for n in range(1, replicates + 1):
            plan.append((api, n))
    return plan


# --- CLI ---
def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="run.py",
        description=(
            "Benchmark orchestrator. Runs a scenario against Tetra and/or "
            "Shortcut, persists per-run artifacts, and writes a summary."
        ),
    )
    p.add_argument("--scenario", required=True,
                   help="Scenario name under scenarios/ (e.g. nyxlight_v1).")
    p.add_argument("--api", required=True,
                   choices=["tetra", "shortcut", "opus-4.7", "gpt-5.5", "both", "all-dg"],
                   help="Which API(s) to exercise. "
                        "'both' = tetra + shortcut (legacy). "
                        "'all-dg' = all 3 DealGlass engines (tetra + opus-4.7 + gpt-5.5).")
    p.add_argument("--replicates", type=int, default=1,
                   help="Replicates per API (default: 1).")
    p.add_argument("--tetra-sas-url", default=None,
                   help="Azure SAS URL for the cap_table.xlsx (required for tetra).")
    p.add_argument("--skip-reconcile", action="store_true",
                   help="Skip the ground-truth reconcile pre-check. Use consciously.")
    p.add_argument("--output-root", default=str(DEFAULT_OUTPUT_ROOT),
                   help="Root directory for run outputs (default: ./runs).")
    p.add_argument("--parallel", action="store_true",
                   help="Run tetra and shortcut concurrently (default: sequential). "
                        "Replicates within an API still run sequentially.")
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    _install_sigint_handler()
    args = parse_args(argv)
    _load_dotenv()

    if args.api == "both":
        apis = ["tetra", "shortcut"]
    elif args.api == "all-dg":
        apis = ["tetra", "opus-4.7", "gpt-5.5"]
    else:
        apis = [args.api]
    log.info("Plan: scenario=%s apis=%s replicates=%d parallel=%s",
             args.scenario, apis, args.replicates, args.parallel)

    # Pre-flight
    pf = preflight(args.scenario, apis, args.tetra_sas_url)
    print(pf.render(), file=sys.stderr)
    if not pf.ok:
        log.error("Pre-flight failed. Fix the items above and re-run.")
        return 2
    assert pf.prompt_text is not None

    # Reconcile
    batch_when = datetime.now(timezone.utc)
    batch_ts = make_batch_timestamp(batch_when)
    batch_dir = Path(args.output_root) / args.scenario / batch_ts
    batch_dir.mkdir(parents=True, exist_ok=True)
    log.info("Batch dir: %s", batch_dir)

    if args.skip_reconcile:
        log.warning("Reconcile SKIPPED by user flag. Scores may be meaningless "
                    "if ground truth is inconsistent.")
    else:
        reconcile_report = batch_dir / "reconcile.md"
        ok = run_reconcile(pf.paths, reconcile_report)
        if not ok:
            log.error("Reconcile failed — refusing to run benchmark. "
                      "Fix ground truth OR pass --skip-reconcile to override.")
            return 3

    # Build plan and execute
    plan = build_plan(apis, args.replicates)
    outcomes: list[RunOutcome] = []

    def _do(api: str, rep: int) -> RunOutcome:
        return execute_run(
            api=api,
            scenario=args.scenario,
            replicate=rep,
            batch_dir=batch_dir,
            prompt_text=pf.prompt_text or "",
            paths=pf.paths,
            tetra_sas_url=args.tetra_sas_url,
        )

    try:
        if args.parallel and len(apis) > 1:
            # One worker per API; replicates within each API stay sequential.
            per_api: dict[str, list[int]] = {a: [] for a in apis}
            for api, rep in plan:
                per_api[api].append(rep)

            def _run_api_sequence(api: str) -> list[RunOutcome]:
                out: list[RunOutcome] = []
                for rep in per_api[api]:
                    if _shutdown.is_set():
                        log.warning("Shutdown requested — skipping %s r%d.", api, rep)
                        break
                    try:
                        out.append(_do(api, rep))
                    except Exception as e:  # noqa: BLE001
                        log.error("[%s r%d] orchestrator-level failure: %s", api, rep, e)
                        log.debug("%s", traceback.format_exc())
                return out

            with ThreadPoolExecutor(max_workers=len(apis)) as pool:
                futs = {pool.submit(_run_api_sequence, api): api for api in apis}
                for fut in as_completed(futs):
                    outcomes.extend(fut.result())
        else:
            for api, rep in plan:
                if _shutdown.is_set():
                    log.warning("Shutdown requested — skipping remaining runs.")
                    break
                try:
                    outcomes.append(_do(api, rep))
                except Exception as e:  # noqa: BLE001
                    log.error("[%s r%d] orchestrator-level failure: %s", api, rep, e)
                    log.debug("%s", traceback.format_exc())
    finally:
        # Always write a summary of whatever we completed.
        # Stable sort: api asc, replicate asc.
        outcomes.sort(key=lambda o: (o.api, o.replicate))
        summary_path = write_summary(outcomes, batch_dir, args.scenario)
        log.info("Summary written to %s", summary_path)
        print("", file=sys.stderr)
        print(render_summary(outcomes, batch_dir, args.scenario))

    if _shutdown.is_set():
        return 130
    # Non-zero exit if every run errored.
    if outcomes and all(o.error is not None for o in outcomes):
        return 1
    return 0


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
