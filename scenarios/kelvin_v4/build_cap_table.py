"""Build canonical and perturbed input cap-table workbooks for Kelvin Dynamics v4.

Produces two files:
  - inputs/cap_table.xlsx           (canonical, primary raise = $30M)
  - inputs/cap_table_perturbed.xlsx (audit probe, primary raise = $33M)

Schema follows scenarios/kelvin_v4/cap_table_build_prompt.md (single source of
truth for tab/column/row layout) and CANONICAL_INPUTS.md §14.1.

Plain-value workbook: no formulas, no totals rows, no merged cells, no
defined names. All numbers are stored as native Python int/float.
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook


# -----------------------------------------------------------------------------
# Data definitions (verbatim from cap_table_build_prompt.md)
# -----------------------------------------------------------------------------

STAKEHOLDER_HEADER = ["ID", "Role", "Common shares", "Preferred class", "Preferred shares"]

# Each row: (ID, Role, Common shares, Preferred class, Preferred shares).
# Empty Preferred class is represented as None -> openpyxl writes blank cell.
STAKEHOLDERS: list[tuple[str, str, int, str | None, int]] = [
    ("F1", "Founder (CEO)", 4_000_000, None, 0),
    ("F2", "Founder (CTO)", 2_500_000, None, 0),
    ("F3", "Founder (Chief Scientist)", 1_500_000, None, 0),
    ("E1", "Employee RSA", 100_000, None, 0),
    ("E2", "Employee RSA", 100_000, None, 0),
    ("E3", "Employee RSA", 100_000, None, 0),
    ("E4", "Employee RSA", 100_000, None, 0),
    ("E5", "Employee RSA", 100_000, None, 0),
    ("E6", "Employee RSA", 100_000, None, 0),
    ("E7", "Employee RSA", 100_000, None, 0),
    ("E8", "Employee RSA", 100_000, None, 0),
    ("E9", "Employee RSA", 100_000, None, 0),
    ("E10", "Employee RSA", 100_000, None, 0),
    ("SEED1", "Seed preferred LP", 0, "Seed", 5_000_000),
    ("A1", "Series A preferred LP", 0, "Series A", 5_000_000),
    ("B1", "Series B preferred LP", 0, "Series B", 3_000_000),
    ("C1", "Series C preferred LP", 0, "Series C", 2_400_000),
    ("D1", "Series D preferred LP", 0, "Series D", 1_720_627),
    ("E1_INV", "Series E preferred LP", 0, "Series E", 610_267),
    ("E2_INV", "Series E preferred LP", 0, "Series E", 610_267),
    ("E3_INV", "Series E preferred LP", 0, "Series E", 610_267),
    ("E4_INV", "Series E preferred LP", 0, "Series E", 610_267),
    ("E5_INV", "Series E preferred LP", 0, "Series E", 610_267),
    ("E6_INV", "Series E preferred LP", 0, "Series E", 610_267),
    ("E7_INV", "Series E preferred LP", 0, "Series E", 610_267),
    ("E8_INV", "Series E preferred LP", 0, "Series E", 610_267),
    ("E9_INV", "Series E preferred LP", 0, "Series E", 610_267),
    ("E10_INV", "Series E preferred LP", 0, "Series E", 610_264),
    ("NA1", "Convertible Note A holder", 0, None, 0),
    ("NB1", "Convertible Note B holder", 0, None, 0),
    ("F_LEAD", "Series F lead investor", 0, None, 0),
    ("F_FUND", "Our fund (Series F primary)", 0, None, 0),
]

SECURITIES_HEADER = [
    "Series",
    "Date",
    "OCP-for-AD",
    "PPS-issued-at",
    "Shares",
    "Invested USD",
    "LP-type",
    "AD-type",
]

LP_TYPE = "1x non-participating with conversion election"
AD_TYPE = "broad-based weighted-average"

SECURITIES: list[tuple[str, str, float, float, int, int, str, str]] = [
    ("Seed",     "2020-Q2",  1.00,      1.00,            5_000_000,   5_000_000, LP_TYPE, AD_TYPE),
    ("Series A", "2021-Q1",  3.00,      3.00,            5_000_000,  15_000_000, LP_TYPE, AD_TYPE),
    ("Series B", "2022-Q2", 10.00,     10.00,            3_000_000,  30_000_000, LP_TYPE, AD_TYPE),
    ("Series C", "2023-Q4", 25.00,     25.00,            2_400_000,  60_000_000, LP_TYPE, AD_TYPE),
    ("Series D", "2025-Q2", 58.118347, 60.00,            1_666_667, 100_000_000, LP_TYPE, AD_TYPE),
    ("Series E", "2026-Q1", 49.158311, 49.158311,        6_102_667, 300_000_000, LP_TYPE, AD_TYPE),
]

NOTES_HEADER = ["Note ID", "Holder ID", "Principal USD", "Cap USD", "Discount", "Interest"]

NOTES: list[tuple[str, str, int, int, float, int]] = [
    ("Note A", "NA1", 5_000_000, 1_000_000_000, 0.20, 0),
    ("Note B", "NB1", 3_000_000, 1_600_000_000, 0.15, 0),
]

ROUND_TERMS_HEADER = ["Term name", "Value", "Units", "Notes"]


def round_terms(primary_raise_usd: int) -> list[tuple[str, float | int, str, str]]:
    """Return the RoundTerms rows. The audit-probe perturbation flips
    `series_f_primary_raise_usd` from 30_000_000 to 33_000_000."""
    return [
        ("series_f_pre_money_usd",         1_400_000_000, "USD",      "Series F pre-money valuation"),
        ("series_f_primary_raise_usd",     primary_raise_usd, "USD",  "New Series F primary capital"),
        ("series_f_secondary_raise_usd",      15_000_000, "USD",      "Secondary tender from founders"),
        ("series_f_tender_pps_multiplier",          0.85, "fraction", "Tender PPS = multiplier x F_PPS"),
        ("esop_refresh_target_post_f",              0.13, "fraction", "Post-F option pool as fraction of post-F FD"),
        ("our_fund_check_usd",                10_000_000, "USD",      "Our fund's Series F primary check (= 1/3 of primary)"),
        ("f_lead_primary_check_usd",          20_000_000, "USD",      "F_LEAD's Series F primary check (= 2/3 of primary)"),
        ("option_strike_usd",                       1.00, "USD",      "Option exercise price for issued options at exit"),
        ("pre_f_options_issued",               2_500_000, "shares",   "Issued options outstanding pre-F (carryover from v3 post-E)"),
        ("pre_f_options_unissued",             1_894_667, "shares",   "Unissued option pool pre-F (carryover from v3 post-E refresh)"),
        ("pre_f_fd_total",                    36_617_961, "shares",   "Pre-F fully-diluted share count (post-E T from v3)"),
        ("exit_value_1_usd",                 700_000_000, "USD",      "Exit scenario 1"),
        ("exit_value_2_usd",               1_200_000_000, "USD",      "Exit scenario 2"),
        ("exit_value_3_usd",               2_000_000_000, "USD",      "Exit scenario 3 (sensitivity reference)"),
        ("exit_value_4_usd",               4_000_000_000, "USD",      "Exit scenario 4"),
        ("exit_value_5_usd",               8_000_000_000, "USD",      "Exit scenario 5"),
        ("sensitivity_esop_low",                    0.11, "fraction", "Sensitivity grid axis: low ESOP refresh"),
        ("sensitivity_esop_base",                   0.13, "fraction", "Sensitivity grid axis: base ESOP refresh"),
        ("sensitivity_esop_high",                   0.15, "fraction", "Sensitivity grid axis: high ESOP refresh"),
        ("sensitivity_pre_low_usd",        1_200_000_000, "USD",      "Sensitivity grid axis: low Series F pre-money"),
        ("sensitivity_pre_base_usd",       1_400_000_000, "USD",      "Sensitivity grid axis: base Series F pre-money"),
        ("sensitivity_pre_high_usd",       1_800_000_000, "USD",      "Sensitivity grid axis: high Series F pre-money"),
        ("return_multiple_3x",                       3.0, "multiple", "Return-solver target: 3x net"),
        ("return_multiple_4x",                       4.0, "multiple", "Return-solver target: 4x net"),
    ]


# -----------------------------------------------------------------------------
# Sanity checks
# -----------------------------------------------------------------------------

def sanity_check_stakeholders_vs_securities() -> None:
    """Sum of preferred shares per class on Stakeholders matches Securities.Shares
    for Seed/A/B/C/E. Series D differs by design (Stakeholders shows AD-adjusted
    1,720,627; Securities shows original issuance 1,666,667)."""
    by_class: dict[str, int] = {}
    for _id, _role, _common, pref_class, pref_shares in STAKEHOLDERS:
        if pref_class:
            by_class[pref_class] = by_class.get(pref_class, 0) + pref_shares

    expected_match = {
        "Seed": 5_000_000,
        "Series A": 5_000_000,
        "Series B": 3_000_000,
        "Series C": 2_400_000,
        "Series E": 6_102_667,
    }
    for cls, expected in expected_match.items():
        got = by_class.get(cls, 0)
        assert got == expected, f"{cls}: stakeholders sum {got} != securities {expected}"

    # Series D: stakeholders carries the AD-adjusted count; securities carries
    # the original issuance. Both are intentional.
    assert by_class["Series D"] == 1_720_627, f"Series D stakeholders sum {by_class['Series D']} != 1_720_627"
    securities_d = next(s for s in SECURITIES if s[0] == "Series D")
    assert securities_d[4] == 1_666_667, "Series D Securities.Shares should be 1,666,667"


def sanity_check_types() -> None:
    """Every numeric cell is a native int or float (not a string)."""
    for row in STAKEHOLDERS:
        _id, role, common, pref_class, pref_shares = row
        assert isinstance(_id, str) and isinstance(role, str)
        assert isinstance(common, int)
        assert pref_class is None or isinstance(pref_class, str)
        assert isinstance(pref_shares, int)
    for row in SECURITIES:
        series, date, ocp, pps, shares, invested, lp, ad = row
        assert isinstance(series, str) and isinstance(date, str)
        assert isinstance(ocp, (int, float)) and not isinstance(ocp, bool)
        assert isinstance(pps, (int, float)) and not isinstance(pps, bool)
        assert isinstance(shares, int)
        assert isinstance(invested, int)
        assert isinstance(lp, str) and isinstance(ad, str)
    for row in NOTES:
        nid, holder, principal, cap, discount, interest = row
        assert isinstance(nid, str) and isinstance(holder, str)
        assert isinstance(principal, int)
        assert isinstance(cap, int)
        assert isinstance(discount, float)
        assert isinstance(interest, int)
    for row in round_terms(30_000_000):
        name, value, units, notes = row
        assert isinstance(name, str)
        assert isinstance(value, (int, float)) and not isinstance(value, bool)
        assert isinstance(units, str)
        assert isinstance(notes, str)


# -----------------------------------------------------------------------------
# Workbook builder
# -----------------------------------------------------------------------------

def build_workbook(primary_raise_usd: int, out_path: Path) -> None:
    wb = Workbook()
    # Remove the default sheet to control tab order.
    default = wb.active
    wb.remove(default)

    # Tab 1: Stakeholders
    ws = wb.create_sheet("Stakeholders")
    ws.append(STAKEHOLDER_HEADER)
    for row in STAKEHOLDERS:
        ws.append(list(row))

    # Tab 2: Securities
    ws = wb.create_sheet("Securities")
    ws.append(SECURITIES_HEADER)
    for row in SECURITIES:
        ws.append(list(row))

    # Tab 3: Notes
    ws = wb.create_sheet("Notes")
    ws.append(NOTES_HEADER)
    for row in NOTES:
        ws.append(list(row))

    # Tab 4: RoundTerms
    ws = wb.create_sheet("RoundTerms")
    ws.append(ROUND_TERMS_HEADER)
    for row in round_terms(primary_raise_usd):
        ws.append(list(row))

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def main() -> None:
    sanity_check_stakeholders_vs_securities()
    sanity_check_types()

    here = Path(__file__).resolve().parent
    inputs_dir = here / "inputs"

    canonical_path = inputs_dir / "cap_table.xlsx"
    perturbed_path = inputs_dir / "cap_table_perturbed.xlsx"

    build_workbook(primary_raise_usd=30_000_000, out_path=canonical_path)
    build_workbook(primary_raise_usd=33_000_000, out_path=perturbed_path)

    n_stakeholders = len(STAKEHOLDERS)
    n_securities = len(SECURITIES)
    n_notes = len(NOTES)
    n_terms = len(round_terms(30_000_000))
    print(
        f"Built {n_stakeholders} stakeholders, {n_securities} securities, "
        f"{n_notes} notes, {n_terms} round terms"
    )


if __name__ == "__main__":
    main()
