"""
Shared interface contract for the benchmark harness.

All clients, scorers, and the orchestrator import from here.
One source of truth for the data shapes — do not duplicate these.

Audit-probe (added v4)
----------------------
Some scenarios (e.g. kelvin_v4) ship a `cap_table_perturbed.xlsx` that
differs from the canonical input in a single named cell (for v4:
`RoundTerms.primary_raise_usd` $30M → $33M). Each scenario with a perturbed
input also ships a `truth_perturbed.json` containing the 167-key truth set
recomputed under the perturbed input. The harness audit-probe sub-test:
opens the engine's output workbook, overwrites the named-input cells with
the perturbed values, runs LibreOffice headless to recalc, re-extracts the
167 truth keys, and reconciles them against `truth_perturbed.json`. A
workbook with live formulas passes ~100%; a value-dump workbook passes
only on the cells we directly overwrote (~0% on dependents).

The contract types added here (`PerturbationSpec`, `AuditProbeResult`) are
the data shapes used by `harness.extract.perturb_and_recalc` and
`harness.score.score_audit_probe`. They are scenario-agnostic — any
scenario with a `truth_perturbed.json` can wire in.
"""

from __future__ import annotations

import time
from dataclasses import dataclass, field, asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Literal


# -------- Canonical event phases --------
# Both clients emit events using these phase strings. The scorer's timing
# extractor keys off them — do not invent new phase strings without updating
# score.py at the same time.
PhaseStr = Literal[
    "run_start",         # harness begins this run (first event)
    "upload_start",      # shortcut: multipart upload begins
    "upload_done",       # shortcut: multipart upload returns fileId
    "submit_start",      # POST to submit endpoint begins
    "submit_ack",        # POST returns (runId / job_id / SSE "accepted")
    "sse_event",         # tetra: any SSE event seen on the wire (detail: {event,data})
    "poll",              # shortcut: GET /{runId} issued
    "status_change",     # observed status transition (detail: {from, to})
    "progress",          # tetra: "progress" SSE or shortcut progress hint
    "complete",          # server reports job complete (before we download)
    "download_start",    # GET for result bytes begins
    "download_done",     # result bytes written to disk
    "error",             # any error — detail carries info
    "run_end",           # harness ends this run (last event)
]


def now_iso() -> str:
    """UTC wall-clock, ISO8601 with Z suffix."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.%fZ")


def mono_ns() -> int:
    """Monotonic nanoseconds. Do not use time.time() for intervals."""
    return time.monotonic_ns()


@dataclass
class RunEvent:
    t_mono_ns: int                 # monotonic ns, relative to process start
    t_wall_iso: str                # ISO8601 UTC
    phase: str                     # one of PhaseStr
    detail: dict[str, Any] = field(default_factory=dict)

    @classmethod
    def make(cls, phase: str, **detail: Any) -> "RunEvent":
        return cls(
            t_mono_ns=mono_ns(),
            t_wall_iso=now_iso(),
            phase=phase,
            detail=detail,
        )


@dataclass
class RunResult:
    run_id: str                    # harness-assigned, e.g. "2026-04-15T12-34-56Z_tetra_nyxlight_v1_r1"
    api: Literal["tetra", "shortcut"]
    scenario: str                  # e.g. "nyxlight_v1"
    replicate: int                 # 1..N for this (api, scenario)
    submitted_at_iso: str          # first submit_start timestamp
    completed_at_iso: str | None   # last complete or error timestamp
    api_run_id: str | None         # server-side runId / job_id (whatever the API returns)
    events: list[RunEvent] = field(default_factory=list)
    output_xlsx_path: str | None = None
    credits_used: float | None = None   # shortcut reports this; tetra does not
    error: str | None = None

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        return d

    def duration_wall_s(self) -> float | None:
        """End-to-end wall clock in seconds. None if incomplete."""
        if not self.events:
            return None
        start = self.events[0].t_mono_ns
        end = self.events[-1].t_mono_ns
        return (end - start) / 1e9

    def duration_to_phase_s(self, phase: str) -> float | None:
        """Seconds from run_start to first event with given phase. None if absent."""
        if not self.events:
            return None
        start = self.events[0].t_mono_ns
        for ev in self.events:
            if ev.phase == phase:
                return (ev.t_mono_ns - start) / 1e9
        return None


def write_events_jsonl(events: list[RunEvent], path: Path) -> None:
    """Append events to a JSONL log file. One event per line."""
    import json
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a") as f:
        for ev in events:
            f.write(json.dumps(asdict(ev), default=str) + "\n")


def write_run_result(result: RunResult, path: Path) -> None:
    """Write the full RunResult as pretty JSON. Overwrites."""
    import json
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w") as f:
        json.dump(result.to_dict(), f, indent=2, default=str)


# ---------------------------------------------------------------------------
# Audit-probe types (v4+)
# ---------------------------------------------------------------------------


@dataclass
class PerturbationSpec:
    """A single input-cell perturbation to apply to an engine's output xlsx.

    `key` is the dot-notation truth-key (e.g. `RoundTerms.primary_raise_usd`)
    that maps to a defined name in the engine's workbook (with `.` → `_`).
    `value` is the new numeric value to write into that cell. `description`
    is human-readable provenance (e.g. "$33M vs $30M canonical") used in
    score reports.
    """
    key: str
    value: float
    description: str = ""

    @property
    def defined_name(self) -> str:
        """Excel defined-name form (`.` → `_`)."""
        return self.key.replace(".", "_")


@dataclass
class AuditProbeResult:
    """Outcome of one audit-probe sub-test.

    `n_keys`/`n_passed` count truth_perturbed.json keys; `pass_pct` is the
    fraction in tolerance after recalculation. `applied`/`unapplied` track
    which `PerturbationSpec`s found a target cell in the workbook.
    `recalc_path` is the LibreOffice-recalculated workbook for inspection.
    `error` is non-None iff something fatal happened (LibreOffice missing,
    workbook corrupt, etc.) — in which case `pass_pct` is 0.0 and the score
    dimension is best treated as zero rather than None.
    """
    n_keys: int
    n_passed: int
    n_missing: int
    pass_pct: float
    applied: list[str] = field(default_factory=list)        # defined names found + written
    unapplied: list[str] = field(default_factory=list)      # defined names not found
    recalc_path: str | None = None
    details: list[dict[str, Any]] = field(default_factory=list)
    error: str | None = None
