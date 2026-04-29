"""
Generate the 4 blog charts for the kelvin_v4 cap-table benchmark.

Outputs:
    public_release/charts/cost_vs_quality.png
    public_release/charts/per_category_radar.png
    public_release/charts/strict_distribution_v4_vs_v3t1.png
    public_release/charts/strict_vs_audit_probe.png

Run from repo root:
    .venv/bin/python public_release/charts/generate_charts.py
"""
from __future__ import annotations
import os
from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.image as mpimg
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
import numpy as np

LOGO_DIR = Path(__file__).parent / "logos"


def _load_logo(name):
    """Load a logo PNG, normalizing to RGBA for clean alpha-blending."""
    img = mpimg.imread(LOGO_DIR / f"{name}.png")
    return img

# ---- Style ----
plt.rcParams.update({
    "font.family": "sans-serif",
    "font.size": 11,
    "axes.titlesize": 13,
    "axes.titleweight": "bold",
    "axes.labelsize": 11,
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.grid": True,
    "grid.color": "#e5e5e5",
    "grid.linewidth": 0.6,
    "figure.facecolor": "white",
    "axes.facecolor": "white",
})

# Engine colors — Tetra prominent, others muted
COLORS = {
    "tetra":      "#1f77b4",  # strong blue
    "claude":     "#999999",  # muted gray
    "codex":      "#cc6677",  # muted red
}
ENGINES = [("tetra", "Tetra"), ("claude", "Opus 4.7"), ("codex", "GPT-5.5")]

OUT = Path(__file__).parent


# =========================================================
# Chart 1: Cost vs Quality scatter (the headline chart)
# =========================================================

def chart_cost_vs_quality():
    fig, ax = plt.subplots(figsize=(8.5, 5.5))

    # Per-engine: ($/strict-pp, mean strict % on v4)
    points = {
        "tetra":  (0.006, 50.06, "Tetra"),
        "claude": (0.008, 36.58, "Opus 4.7"),
        "codex":  (0.31,  26.90, "GPT-5.5"),
    }
    for k, (cost, strict, label) in points.items():
        ax.scatter(cost, strict,
                   s=400, color=COLORS[k], edgecolor="white",
                   linewidth=2, zorder=3, label=label)
        # Annotate above
        offset_y = 4 if k != "codex" else -7
        ax.annotate(label, (cost, strict), xytext=(0, offset_y),
                    textcoords="offset points", ha="center",
                    fontsize=12, fontweight="bold",
                    color=COLORS[k])

    ax.set_xscale("log")
    ax.set_xlabel("Cost per percentage point of strict accuracy (log scale, USD)",
                  fontsize=11)
    ax.set_ylabel("Mean strict accuracy (%, n=10)", fontsize=11)
    ax.set_title("Cost-of-quality on real Series F modeling")
    ax.set_ylim(15, 60)
    ax.set_xlim(0.003, 1.0)

    # Highlight the gap with a band
    ax.axhspan(45, 55, alpha=0.05, color="#1f77b4", zorder=1)

    # Annotation for the headline ratio
    ax.annotate("", xy=(0.31, 30), xytext=(0.006, 30),
                arrowprops=dict(arrowstyle="<->", color="#666", lw=1))
    ax.text(0.04, 31, "50× cost gap",
            fontsize=10, ha="center", color="#666",
            bbox=dict(boxstyle="round,pad=0.3",
                      facecolor="white", edgecolor="#ccc", linewidth=0.5))

    fig.tight_layout()
    fig.savefig(OUT / "cost_vs_quality.png", dpi=160, bbox_inches="tight")
    plt.close(fig)
    print("Wrote cost_vs_quality.png")


# =========================================================
# Chart 2: Per-category radar on v4
# =========================================================

def chart_per_category_radar():
    categories = [
        "Ownership", "Anti-dilution", "Convertible\nnotes", "Secondary\ntender",
        "Round\nsolution", "Waterfall\ndollars", "Waterfall\nelections",
        "Sensitivity", "Return\nsolver",
    ]
    # From audit_d3_per_category.md (post-extractor-fix-round-3)
    data = {
        "tetra":  [84.0, 72.5, 43.0, 66.7, 52.5, 31.5, 75.4, 50.0, 62.0],
        "claude": [76.0, 83.3, 48.0, 70.0, 65.0, 11.6, 67.1,  5.0, 20.0],
        "codex":  [55.3, 56.7, 18.0, 20.0, 20.0,  7.1, 30.0, 20.0, 30.0],
    }

    n = len(categories)
    angles = np.linspace(0, 2 * np.pi, n, endpoint=False).tolist()
    angles += angles[:1]  # close

    fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(polar=True))

    for k, label in [("tetra", "Tetra"), ("claude", "Opus 4.7"), ("codex", "GPT-5.5")]:
        vals = data[k] + data[k][:1]
        ax.plot(angles, vals, color=COLORS[k], linewidth=2, label=label)
        ax.fill(angles, vals, color=COLORS[k], alpha=0.12)

    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(categories, fontsize=10)
    ax.set_ylim(0, 100)
    ax.set_yticks([20, 40, 60, 80, 100])
    ax.set_yticklabels(["20%", "40%", "60%", "80%", "100%"], fontsize=8, color="#666")
    ax.set_rlabel_position(90)
    ax.set_title("Per-category strict-pass rate",
                 pad=24, fontsize=13, fontweight="bold")
    ax.legend(loc="upper right", bbox_to_anchor=(1.18, 1.0), fontsize=10)
    ax.grid(color="#dcdcdc", linewidth=0.6)

    fig.tight_layout()
    fig.savefig(OUT / "per_category_radar.png", dpi=160, bbox_inches="tight")
    plt.close(fig)
    print("Wrote per_category_radar.png")


# =========================================================
# Chart 3: Strict-score distribution v4 vs v3_t1 (the regime-conditional chart)
# =========================================================

def chart_strict_distribution():
    # Per-run scores (n=10 each)
    v4 = {
        "tetra":  [54.5, 42.5, 51.5, 78.2, 74.2, 50.3, 25.7, 36.5, 45.5, 73.1],
        "claude": [55.6, 47.6,  7.2, 16.8, 28.1, 53.9, 43.1, 56.9, 47.3, 26.3],
        "codex":  [16.8, 18.6, 43.1, 55.1,  9.0, 23.4, 31.7, 47.3, 59.3, 12.0],
    }
    v3t1 = {
        "tetra":  [70.2, 83.9, 67.7, 83.1, 78.2, 74.2, 88.7, 53.2, 60.5, 59.7],
        "claude": [55.6, 47.6, 79.8, 71.8, 49.2, 79.0, 83.1, 88.7, 88.7, 59.7],
        "codex":  [62.1, 62.1, 92.7, 92.7, 54.0, 54.0, 62.1, 99.2, 99.2, 62.1],
    }

    fig, axes = plt.subplots(1, 2, figsize=(11, 5), sharey=True)

    for i, (data, title) in enumerate([
        (v4,  "Build-from-scratch deal"),
        (v3t1, "Templated task"),
    ]):
        ax = axes[i]
        positions = [1, 2, 3]
        engine_keys = ["tetra", "claude", "codex"]
        engine_labels = ["Tetra", "Opus 4.7", "GPT-5.5"]
        bp = ax.boxplot(
            [data[k] for k in engine_keys],
            positions=positions, widths=0.55, patch_artist=True,
            medianprops=dict(color="white", linewidth=2),
            boxprops=dict(linewidth=1, edgecolor="#333"),
            whiskerprops=dict(linewidth=1, color="#333"),
            capprops=dict(linewidth=1, color="#333"),
            flierprops=dict(marker="o", markersize=4, markerfacecolor="#666",
                            markeredgecolor="none", alpha=0.7),
        )
        for patch, k in zip(bp["boxes"], engine_keys):
            patch.set_facecolor(COLORS[k])
            patch.set_alpha(0.85)
        # overlay individual run dots
        for pos, k in zip(positions, engine_keys):
            xs = np.random.RandomState(42 + pos).uniform(pos - 0.18, pos + 0.18, len(data[k]))
            ax.scatter(xs, data[k], s=20, color="#222",
                       alpha=0.5, zorder=3)

        ax.set_xticks(positions)
        ax.set_xticklabels(engine_labels)
        ax.set_title(title)
        ax.set_ylim(0, 100)
        if i == 0:
            ax.set_ylabel("Strict accuracy (%, n=10)")

    fig.suptitle("Agent edge is conditional on the regime",
                 fontsize=14, fontweight="bold", y=1.02)
    fig.tight_layout()
    fig.savefig(OUT / "strict_distribution_v4_vs_v3t1.png", dpi=160, bbox_inches="tight")
    plt.close(fig)
    print("Wrote strict_distribution_v4_vs_v3t1.png")


# =========================================================
# Chart 4: Strict vs audit-probe scatter (the auditability chart)
# =========================================================

def chart_strict_vs_audit_probe():
    # (engine, org, strict_mean, audit_probe_mean, named_cell_apply_rate)
    data = [
        ("Tetra",   "DealGlass", 50.06, 65.6, "6/10 named-cell coverage"),
        ("Opus 4.7","Anthropic", 36.58, 38.6, "0/10 — could not be probed"),
        ("GPT-5.5", "OpenAI",    26.90, 44.7, "2/10 — math chain breaks when probed"),
    ]
    keys = ["tetra", "claude", "codex"]

    fig, ax = plt.subplots(figsize=(8.5, 6))

    xlim = (0, 80)
    ylim = (0, 80)

    # Shaded zones: above diagonal = auditable (green), below = static (red)
    ax.fill_between([0, 100], [0, 100], [100, 100],
                    color="#2e7d32", alpha=0.07, zorder=0)
    ax.fill_between([0, 100], [0, 0], [0, 100],
                    color="#c62828", alpha=0.06, zorder=0)

    # Diagonal reference (strict == audit-probe)
    ax.plot([0, 100], [0, 100], color="#888888",
            linestyle="--", linewidth=1, zorder=1)

    # Zone labels
    ax.text(8, 72, "AUDITABLE",
            fontsize=13, fontweight="bold", color="#1b5e20", alpha=0.85)
    ax.text(8, 66, "formulas recompute when\nthe analyst changes an input",
            fontsize=9.5, color="#1b5e20", alpha=0.85, style="italic")
    ax.text(60, 10, "STATIC",
            fontsize=13, fontweight="bold", color="#b71c1c", alpha=0.85,
            ha="right")
    ax.text(60, 4, "values look right but\ndon't recompute on perturbation",
            fontsize=9.5, color="#b71c1c", alpha=0.85, style="italic",
            ha="right")

    # Diagonal text label (positioned in upper-right empty space)
    ax.text(73, 76, "strict = recompute",
            fontsize=8.5, color="#666", rotation=37,
            rotation_mode="anchor", ha="center")

    # Use brand logos as markers (all three cropped to circles for visual parity)
    logo_map = {"tetra": "tetra_circle", "claude": "claude_circle", "codex": "openai_circle"}
    LOGO_ZOOM = 0.09  # source is 480px, halved zoom keeps display size constant
    # Per-engine label offsets — Opus 4.7 pushed down to clear GPT-5.5's caption
    label_offsets = {
        "tetra":  {"name": (26, 6),   "desc": (26, -10)},
        "codex":  {"name": (26, 10),  "desc": (26, -4)},
        "claude": {"name": (26, -22), "desc": (26, -36)},
    }
    for (label, org, s, ap, ann), k in zip(data, keys):
        img = _load_logo(logo_map[k])
        oi = OffsetImage(img, zoom=LOGO_ZOOM)
        ab = AnnotationBbox(oi, (s, ap), frameon=False,
                            box_alignment=(0.5, 0.5), zorder=3)
        ax.add_artist(ab)
        off = label_offsets[k]
        # Engine name (bold, in brand color) followed by org (muted)
        ax.annotate(label, (s, ap), xytext=off["name"],
                    textcoords="offset points", fontsize=12, fontweight="bold",
                    color=COLORS[k])
        ax.annotate(f"({org})", (s, ap),
                    xytext=(off["name"][0] + 7 * len(label) + 4, off["name"][1]),
                    textcoords="offset points", fontsize=10,
                    color="#888")
        ax.annotate(ann, (s, ap), xytext=off["desc"],
                    textcoords="offset points", fontsize=9, color="#555",
                    style="italic")

    ax.set_xlabel("Strict accuracy (%) — values match canonical truth", fontsize=11)
    ax.set_ylabel("Audit-probe pass-rate (%) — cells update under input perturbation", fontsize=11)
    ax.set_title("Audit-readiness: do the cells update when an input changes?")
    ax.set_xlim(*xlim)
    ax.set_ylim(*ylim)

    fig.tight_layout()
    fig.savefig(OUT / "strict_vs_audit_probe.png", dpi=160, bbox_inches="tight")
    plt.close(fig)
    print("Wrote strict_vs_audit_probe.png")


# =========================================================
# Chart 5: Time vs strict accuracy scatter (NEW for v4)
# =========================================================

def chart_time_vs_accuracy():
    # Per-run (wall_clock_s, strict_pct) for each engine on v4
    # Pulled from results/v4_strict_scores.csv at n=10 each
    data = {
        "tetra":  [(1310,54.5),(1443,42.5),(1473,51.5),(1610,78.2),(1610,74.2),
                   (1620,50.3),(1840,25.7),(1925,36.5),(1948,45.5),(2745,73.1)],
        "claude": [(1230,55.6),(1242,47.6),(1265,7.2),(1275,16.8),(1310,28.1),
                   (1485,53.9),(1605,43.1),(1763,56.9),(1845,47.3),(2402,26.3)],
        "codex":  [(5,7.8),(28,7.8),(69,7.8),(436,62.1),(497,12.0),(497,9.0),
                   (520,18.6),(692,16.8),(924,47.3),(1469,59.3)],
    }

    fig, ax = plt.subplots(figsize=(9, 5.5))

    for k, label in [("tetra", "Tetra"), ("claude", "Opus 4.7"), ("codex", "GPT-5.5")]:
        xs = [p[0]/60 for p in data[k]]  # convert seconds to minutes
        ys = [p[1] for p in data[k]]
        ax.scatter(xs, ys, s=140, color=COLORS[k], edgecolor="white",
                   linewidth=1.5, zorder=3, label=label, alpha=0.9)

    ax.set_xlabel("Wall-clock time per run (minutes)")
    ax.set_ylabel("Strict accuracy (%)")
    ax.set_title("Time vs accuracy: 10 runs per engine")
    ax.set_xlim(-2, 50)
    ax.set_ylim(0, 90)
    ax.legend(loc="upper right", fontsize=11)
    # Reference: SpreadsheetBench median
    ax.axvline(x=112/60, color="#aaa", linestyle=":", linewidth=1, zorder=1)
    ax.text(112/60 + 0.5, 85, "SpreadsheetBench median (1.9 min)",
            fontsize=9, color="#666", style="italic")

    fig.tight_layout()
    fig.savefig(OUT / "time_vs_accuracy.png", dpi=160, bbox_inches="tight")
    plt.close(fig)
    print("Wrote time_vs_accuracy.png")


# =========================================================

if __name__ == "__main__":
    chart_cost_vs_quality()
    chart_per_category_radar()
    chart_strict_distribution()
    chart_strict_vs_audit_probe()
    chart_time_vs_accuracy()
    print("\nAll 5 charts written to:", OUT)
