"""
Kelvin Dynamics -- Canonical Scenario v4 Python ground-truth builder.
("Series F: Tender + Convertibles + Down-Round AD")

Computes every tracked output (167 keys) per CANONICAL_INPUTS.md §12 from
first principles, using decimal.Decimal (getcontext().prec = 60). Writes
truth.json (raise=$30M) and truth_perturbed.json (raise=$33M) next to
this file.

Design notes:

1. Pre-F state inherits from v3's post-E fixed point (CANONICAL_INPUTS §1.1).
   We re-solve v3 internally (cheap) to get full-precision Decimals for
   E_PPS, T_E, r_D(v3), NCP_D, new_E_shares, option pool, etc., and assert
   against v3's truth.json values to 1e-6 abs as a sanity gate.

2. v4 round fixed point couples five unknowns:
       F_PPS, T_F, r_E, r_D(v4), N_A_shares, N_B_shares
   solved by Picard iteration. The note conversion branches are determined
   per-iteration by computing both cap-implied PPS and discount-implied
   PPS for each note and taking the min. We assert post-convergence that
   the case assumption holds: Note A converts at cap, Note B at discount.

3. Anti-dilution: broad-based weighted-average on Series E and Series D
   (both trigger; F_PPS ~$37.5 < OCP_E=$49.16 and < OCP_D=$58.12). For
   Series D, OCP for v4 AD purposes is the post-E NCP_D from v3 (NOT the
   original $60). The compound ratio = r_D(v3) * r_D(v4) is what's
   reported as antidilution.series_d.new_ratio.
   Notes are EXCLUDED from AD broad-base B and C per §7.2.

4. Tender ($15M @ 0.85*F_PPS): existing common shares transfer from
   founders to F_LEAD. No new shares issued, T_F unchanged. Founders'
   share count reduces pro rata (50/31.25/18.75 split). Tender shares
   are tracked as F_LEAD's separate "secondary" common-equivalent
   holding for the ownership.f_lead_secondary.pct line.

5. Waterfall: 7 priced classes (F senior to all) x 5 exits = 35 elections.
   Nash equilibrium by enumeration over 2^7 = 128 combos per exit. Notes
   A and B do NOT have elections; they ride pro-rata always.

6. ESOP refresh: pre-money top-up to 13% of T_F. Pre-F pool already has
   4,394,124 shares (from v3 post-E refresh). top_up = max(0, 0.13*T_F -
   pool_pre_F). If 0.13*T_F < 4,394,124, clamp to zero.

Run: python truth.py
"""

from __future__ import annotations

import json
import os
from decimal import Decimal, getcontext, ROUND_HALF_EVEN
from typing import Dict, List, Tuple

getcontext().prec = 60

# -----------------------------------------------------------------------------
# Decimal helpers
# -----------------------------------------------------------------------------

def d(x) -> Decimal:
    if isinstance(x, Decimal):
        return x
    return Decimal(str(x))


ZERO = d(0)
ONE = d(1)


def round_money(x: Decimal) -> Decimal:
    return x.quantize(d("0.01"), rounding=ROUND_HALF_EVEN)


def round_pps(x: Decimal) -> Decimal:
    # 6 decimal places per spec
    return x.quantize(d("0.000001"), rounding=ROUND_HALF_EVEN)


def round_pct(x: Decimal) -> Decimal:
    # 9 decimal places (percentages stored as fractions; 9dp)
    return x.quantize(d("0.000000001"), rounding=ROUND_HALF_EVEN)


def round_ratio(x: Decimal) -> Decimal:
    # 9 decimal places for ratios
    return x.quantize(d("0.000000001"), rounding=ROUND_HALF_EVEN)


def to_float(x: Decimal) -> float:
    return float(x)


def int_round(x: Decimal) -> int:
    return int(x.quantize(d("1"), rounding=ROUND_HALF_EVEN))


# =============================================================================
# Inputs (locked per CANONICAL_INPUTS.md)
# =============================================================================

# --- Common / founders / employees (§1.1, inherited from v3) ---
SHARES_F1 = d(4_000_000)
SHARES_F2 = d(2_500_000)
SHARES_F3 = d(1_500_000)
SHARES_FOUNDERS = SHARES_F1 + SHARES_F2 + SHARES_F3  # 8M
SHARES_EMPLOYEES = d(1_000_000)  # E1..E10 RSAs
SHARES_COMMON = SHARES_FOUNDERS + SHARES_EMPLOYEES  # 9M
assert SHARES_COMMON == d(9_000_000)

# --- Prior priced rounds (§3, locked) ---
SEED_PPS = d(1)
SEED_SHARES = d(5_000_000)
SEED_RAISE = d(5_000_000)
SEED_LP = SEED_RAISE

A_PPS = d(3)
A_SHARES = d(5_000_000)
A_RAISE = d(15_000_000)
A_LP = A_RAISE

B_PPS = d(10)
B_SHARES = d(3_000_000)
B_RAISE = d(30_000_000)
B_LP = B_RAISE

C_PPS = d(25)
C_SHARES = d(2_400_000)
C_RAISE = d(60_000_000)
C_LP = C_RAISE

D_PPS_ORIGINAL = d(60)            # original D OCP (used in v3 AD)
D_SHARES_ORIGINAL = d(1_666_667)  # original D issued share count
D_RAISE = d(100_000_000)
D_LP = D_RAISE

# --- Series E (§4, CLOSED in v4) ---
E_PRE_MONEY_V3 = d(1_500_000_000)
E_RAISE = d(300_000_000)
E_POST_MONEY = d(1_800_000_000)
E_LP = E_RAISE
ESOP_TARGET_V3 = d("0.12")
OPTIONS_ISSUED = d(2_500_000)         # unchanged through E and F
OPTIONS_UNISSUED_PRE_E = d(500_000)
OPTIONS_POOL_PRE_E = OPTIONS_ISSUED + OPTIONS_UNISSUED_PRE_E
A_BB_V3 = (SHARES_COMMON + SEED_SHARES + A_SHARES + B_SHARES + C_SHARES
           + D_SHARES_ORIGINAL + OPTIONS_ISSUED + OPTIONS_UNISSUED_PRE_E)
assert A_BB_V3 == d(29_066_667)

# --- Series F (§5, the round) ---
F_PRE_MONEY = d(1_400_000_000)
F_PRIMARY_RAISE_BASE = d(30_000_000)
F_SECONDARY_RAISE = d(15_000_000)
TENDER_PPS_MULT = d("0.85")
ESOP_TARGET_F = d("0.13")
F_LEAD_PRIMARY_USD = d(20_000_000)
FUND_CHECK = d(10_000_000)

# --- Convertible notes (§6) ---
NOTE_A_PRINCIPAL = d(5_000_000)
NOTE_A_CAP = d(1_000_000_000)
NOTE_A_DISCOUNT = d("0.20")

NOTE_B_PRINCIPAL = d(3_000_000)
NOTE_B_CAP = d(1_600_000_000)
NOTE_B_DISCOUNT = d("0.15")

# --- Exits (§8) ---
EXITS_LABEL = ["700", "1200", "2000", "4000", "8000"]
EXITS_USD = [d(700_000_000), d(1_200_000_000), d(2_000_000_000),
             d(4_000_000_000), d(8_000_000_000)]

# --- Options strike (matches v3) ---
OPTIONS_STRIKE_WA = d("0.50")


# =============================================================================
# v3 inheritance: re-solve v3's post-E fixed point internally
# =============================================================================

def solve_v3_post_e() -> Dict:
    """Re-derive v3's post-E state. Matches scenarios/kelvin_v3/truth/truth.py
    exactly. Returns full-precision Decimals for inheritance into v4.
    """
    pref_non_D = SEED_SHARES + A_SHARES + B_SHARES + C_SHARES  # 15.4M
    r_D = ONE
    for _ in range(500):
        pref_D_adj = D_SHARES_ORIGINAL * r_D
        denom_coef = ONE - ESOP_TARGET_V3 - (E_RAISE / E_POST_MONEY)
        rhs = SHARES_COMMON + pref_non_D + pref_D_adj
        T = rhs / denom_coef
        E_PPS = E_POST_MONEY / T
        new_E = T * E_RAISE / E_POST_MONEY
        if E_PPS < D_PPS_ORIGINAL:
            NCP_D = D_PPS_ORIGINAL * (A_BB_V3 + E_RAISE / D_PPS_ORIGINAL) / (A_BB_V3 + new_E)
            r_D_new = D_PPS_ORIGINAL / NCP_D
        else:
            NCP_D = D_PPS_ORIGINAL
            r_D_new = ONE
        if abs(r_D_new - r_D) < d("1e-50"):
            r_D = r_D_new
            break
        r_D = r_D_new
    else:
        raise RuntimeError("v3 fixed point did not converge")

    pref_D_adj_exact = D_SHARES_ORIGINAL * r_D
    pool_total = ESOP_TARGET_V3 * T
    options_unissued_post_e = pool_total - OPTIONS_ISSUED

    return {
        "E_PPS": E_PPS,
        "T_E": T,
        "new_E_shares": new_E,
        "r_D_v3": r_D,
        "NCP_D_v3": NCP_D,
        "pref_D_adj_v3_exact": pref_D_adj_exact,
        "pref_D_adj_v3_int": int_round(pref_D_adj_exact),
        "options_pool_post_e": pool_total,
        "options_unissued_post_e": options_unissued_post_e,
    }


def assert_v3_inheritance(v3: Dict) -> None:
    """Sanity-check v3 values against truth.json's emitted (rounded)
    values. Tolerance: 1e-6 abs (matches §13 PPS tolerance).
    """
    # Hardcoded from /scenarios/kelvin_v3/truth/truth.json
    EXPECTED_NCP_D = d("58.118347")
    EXPECTED_R_D = d("1.03237623")
    # Sanity-assertable derived: pref_D_adj_int = 1,720,627 per CANONICAL §1.1
    assert abs(v3["NCP_D_v3"] - EXPECTED_NCP_D) < d("1e-6"), \
        f"v3 NCP_D inheritance mismatch: {v3['NCP_D_v3']} vs {EXPECTED_NCP_D}"
    assert abs(v3["r_D_v3"] - EXPECTED_R_D) < d("1e-6"), \
        f"v3 r_D inheritance mismatch: {v3['r_D_v3']} vs {EXPECTED_R_D}"
    assert v3["pref_D_adj_v3_int"] == 1_720_627, \
        f"v3 pref_D_adj int mismatch: {v3['pref_D_adj_v3_int']} vs 1720627"
    # post-E FD T should be ~36.6M (within 1 share)
    assert abs(v3["T_E"] - d(36_617_702)) < d(2), \
        f"v3 T_E inheritance mismatch: {v3['T_E']}"
    # E_PPS ~ $49.16
    assert abs(v3["E_PPS"] - d("49.156553")) < d("1e-6"), \
        f"v3 E_PPS inheritance mismatch: {v3['E_PPS']}"


# =============================================================================
# v4 round fixed point (5 unknowns)
# =============================================================================

def solve_round_v4(f_pre_money: Decimal,
                   f_primary_raise: Decimal,
                   esop_target_post: Decimal,
                   v3: Dict,
                   max_iter: int = 200,
                   tol: Decimal = d("1e-30")) -> Dict:
    """Picard-iterate over (F_PPS, T_F, r_E, r_D_v4, N_A, N_B).

    F_PPS = post_money_F / T_F where post_money_F = pre_money + primary_raise.
    new_F_shares = primary_raise / F_PPS = T_F * primary_raise / post_money_F.

    Series E AD (broad-based WA), OCP_E = E_PPS (v3 fixed point):
        NCP_E = OCP_E * (A_v4 + raise/OCP_E) / (A_v4 + new_F_shares)
        r_E = OCP_E / NCP_E
    Series D AD (v4 increment), OCP_D = NCP_D from v3:
        NCP_D_v4 = OCP_D_v4 * (A_v4 + raise/OCP_D_v4) / (A_v4 + new_F_shares)
        r_D_v4 = OCP_D_v4 / NCP_D_v4
    Where A_v4 = pre-F FD = T_E (v3 post-E FD). Notes excluded from B,C
    per §7.2.

    Note conversions:
      cap-implied = cap / pre_money_FD_at_conversion
        where pre_money_FD_at_conversion = T_F - new_F_shares
        (= all pre-F holders + ESOP top-up + BOTH notes' shares;
         the OTHER note's share is NOT excluded since both convert)
      discount-implied = F_PPS * (1 - discount)
      conv_pps = min(cap, discount); shares = principal / conv_pps.

    T_F coupling:
      T_F = SHARES_COMMON + (SEED+A+B+C, ratio 1) + D_orig*r_D_v3*r_D_v4
            + E_shares*r_E + N_A + N_B + new_F + OPTIONS_ISSUED
            + (esop * T_F - OPTIONS_ISSUED)
      Rearrange: T_F * (1 - esop - primary_raise/post_money_F) =
            SHARES_COMMON + pref_non_D + D_orig*r_D_compound + E_shares*r_E
            + N_A + N_B
      (note: ESOP top-up may clamp to zero if 0.13*T_F < 4,394,124;
       handled separately.)
    """
    post_money_F = f_pre_money + f_primary_raise
    T_E = v3["T_E"]
    new_E_shares = v3["new_E_shares"]
    r_D_v3 = v3["r_D_v3"]
    OCP_E = v3["E_PPS"]
    OCP_D_v4 = v3["NCP_D_v3"]
    A_v4 = T_E

    pref_non_D = SEED_SHARES + A_SHARES + B_SHARES + C_SHARES  # 15.4M

    # Initial guess: F_PPS ~ $37.5 per §16
    F_PPS = d("37.5")
    r_E = ONE
    r_D_v4 = ONE
    N_A = d(183_000)
    N_B = d(94_000)
    T_F = T_E + d(2_000_000)  # ballpark

    prev_state = (ZERO, ZERO, ZERO, ZERO, ZERO, ZERO)

    for it in range(max_iter):
        # Compute T_F from the closed equation given (r_E, r_D_v4, N_A, N_B)
        # The "ESOP no-clamp" branch (top-up positive): need to verify after.
        D_compound = D_SHARES_ORIGINAL * r_D_v3 * r_D_v4
        E_adj_shares = new_E_shares * r_E
        rhs_no_esop_clamp = (SHARES_COMMON + pref_non_D + D_compound
                             + E_adj_shares + N_A + N_B)
        denom_coef = ONE - esop_target_post - (f_primary_raise / post_money_F)
        T_F_new = rhs_no_esop_clamp / denom_coef

        # Check if ESOP would clamp (top-up < 0): pool_total < pool_pre_F
        pool_total_unclamped = esop_target_post * T_F_new
        pool_pre_F = OPTIONS_POOL_PRE_E  # 3M (issued 2.5M + unissued 0.5M pre-E)
        # Actually pre-F option pool: post-E refresh gave 12% of T_E = 4,394,124
        # So pre-F has v3["options_pool_post_e"] shares
        pool_pre_F = v3["options_pool_post_e"]

        if pool_total_unclamped < pool_pre_F:
            # Clamp: top-up = 0; option pool post-F = pool_pre_F (unchanged)
            # T_F equation changes:
            # T_F = common + pref_non_D + D_compound + E_adj + N_A + N_B
            #       + new_F + pool_pre_F (no top-up)
            # T_F * (1 - raise/post_money_F) = ... + pool_pre_F
            denom_coef2 = ONE - (f_primary_raise / post_money_F)
            rhs2 = rhs_no_esop_clamp + pool_pre_F
            T_F_new = rhs2 / denom_coef2
            esop_clamped = True
        else:
            esop_clamped = False

        F_PPS_new = post_money_F / T_F_new
        new_F_shares = f_primary_raise / F_PPS_new

        # AD on E
        if F_PPS_new < OCP_E:
            NCP_E_new = OCP_E * (A_v4 + f_primary_raise / OCP_E) / (A_v4 + new_F_shares)
            r_E_new = OCP_E / NCP_E_new
        else:
            NCP_E_new = OCP_E
            r_E_new = ONE

        # AD on D (v4 increment, OCP = post-E NCP_D)
        if F_PPS_new < OCP_D_v4:
            NCP_D_v4_new = OCP_D_v4 * (A_v4 + f_primary_raise / OCP_D_v4) / (A_v4 + new_F_shares)
            r_D_v4_new = OCP_D_v4 / NCP_D_v4_new
        else:
            NCP_D_v4_new = OCP_D_v4
            r_D_v4_new = ONE

        # Note conversions
        # pre-money-FD-at-conversion = T_F - new_F_shares
        #   = SHARES_COMMON + pref_non_D + D_compound + E_adj + N_A + N_B
        #     + options pool (post-F)
        pre_money_FD_at_conv = T_F_new - new_F_shares

        # Note A
        cap_pps_A = NOTE_A_CAP / pre_money_FD_at_conv
        disc_pps_A = F_PPS_new * (ONE - NOTE_A_DISCOUNT)
        if cap_pps_A < disc_pps_A:
            conv_pps_A = cap_pps_A
            branch_A = "cap"
        else:
            conv_pps_A = disc_pps_A
            branch_A = "discount"
        N_A_new = NOTE_A_PRINCIPAL / conv_pps_A

        # Note B
        cap_pps_B = NOTE_B_CAP / pre_money_FD_at_conv
        disc_pps_B = F_PPS_new * (ONE - NOTE_B_DISCOUNT)
        if cap_pps_B < disc_pps_B:
            conv_pps_B = cap_pps_B
            branch_B = "cap"
        else:
            conv_pps_B = disc_pps_B
            branch_B = "discount"
        N_B_new = NOTE_B_PRINCIPAL / conv_pps_B

        state_new = (F_PPS_new, T_F_new, r_E_new, r_D_v4_new, N_A_new, N_B_new)
        max_delta = max(abs(s_n - s_p) for s_n, s_p in zip(state_new, prev_state))
        prev_state = state_new

        F_PPS = F_PPS_new
        T_F = T_F_new
        r_E = r_E_new
        r_D_v4 = r_D_v4_new
        N_A = N_A_new
        N_B = N_B_new

        if max_delta < tol:
            break
    else:
        raise RuntimeError(f"v4 fixed point did not converge: max_delta={max_delta}")

    # Final values
    new_F_shares = f_primary_raise / F_PPS
    D_compound_ratio = r_D_v3 * r_D_v4
    D_adj_v4_shares = D_SHARES_ORIGINAL * D_compound_ratio
    E_adj_shares = new_E_shares * r_E

    # ESOP final
    pool_total_F = esop_target_post * T_F
    pool_pre_F = v3["options_pool_post_e"]
    if pool_total_F < pool_pre_F:
        # Clamp branch: pool_post_F = pool_pre_F
        pool_total_F_actual = pool_pre_F
        top_up = ZERO
        options_unissued_post_F = pool_pre_F - OPTIONS_ISSUED
    else:
        pool_total_F_actual = pool_total_F
        top_up = pool_total_F - pool_pre_F
        options_unissued_post_F = pool_total_F - OPTIONS_ISSUED

    # Tender
    tender_pps = TENDER_PPS_MULT * F_PPS
    tender_shares_exact = F_SECONDARY_RAISE / tender_pps
    # F_LEAD primary
    f_lead_primary_shares = F_LEAD_PRIMARY_USD / F_PPS
    fund_shares = FUND_CHECK / F_PPS

    # Note branches at final F_PPS
    pre_money_FD_at_conv = T_F - new_F_shares
    cap_pps_A = NOTE_A_CAP / pre_money_FD_at_conv
    disc_pps_A = F_PPS * (ONE - NOTE_A_DISCOUNT)
    if cap_pps_A < disc_pps_A:
        conv_pps_A = cap_pps_A
        branch_A = "cap"
    else:
        conv_pps_A = disc_pps_A
        branch_A = "discount"

    cap_pps_B = NOTE_B_CAP / pre_money_FD_at_conv
    disc_pps_B = F_PPS * (ONE - NOTE_B_DISCOUNT)
    if cap_pps_B < disc_pps_B:
        conv_pps_B = cap_pps_B
        branch_B = "cap"
    else:
        conv_pps_B = disc_pps_B
        branch_B = "discount"

    # NCP and ratio for E, D (final)
    if F_PPS < OCP_E:
        NCP_E = OCP_E * (A_v4 + f_primary_raise / OCP_E) / (A_v4 + new_F_shares)
    else:
        NCP_E = OCP_E
    if F_PPS < OCP_D_v4:
        NCP_D_v4_final = OCP_D_v4 * (A_v4 + f_primary_raise / OCP_D_v4) / (A_v4 + new_F_shares)
    else:
        NCP_D_v4_final = OCP_D_v4

    return {
        "F_PPS": F_PPS,
        "T_F": T_F,
        "post_money_F": post_money_F,
        "new_F_shares": new_F_shares,
        "f_lead_primary_shares": f_lead_primary_shares,
        "fund_shares": fund_shares,
        "r_E": r_E,
        "NCP_E": NCP_E,
        "E_adj_shares": E_adj_shares,
        "r_D_v4": r_D_v4,
        "r_D_compound": D_compound_ratio,
        "NCP_D_v4": NCP_D_v4_final,
        "D_adj_v4_shares": D_adj_v4_shares,
        "N_A": N_A,
        "N_B": N_B,
        "branch_A": branch_A,
        "branch_B": branch_B,
        "conv_pps_A": conv_pps_A,
        "conv_pps_B": conv_pps_B,
        "cap_pps_A": cap_pps_A,
        "cap_pps_B": cap_pps_B,
        "disc_pps_A": disc_pps_A,
        "disc_pps_B": disc_pps_B,
        "pre_money_FD_at_conv": pre_money_FD_at_conv,
        "pool_total_F": pool_total_F_actual,
        "options_issued_post_F": OPTIONS_ISSUED,
        "options_unissued_post_F": options_unissued_post_F,
        "top_up": top_up,
        "esop_clamped": esop_clamped,
        "esop_target": esop_target_post,
        "f_pre_money": f_pre_money,
        "f_primary_raise": f_primary_raise,
        "tender_pps": tender_pps,
        "tender_shares_exact": tender_shares_exact,
    }


# =============================================================================
# Tender mechanics
# =============================================================================

def compute_tender(R: Dict) -> Dict:
    """Founders sell tender_shares pro-rata (50/31.25/18.75) to F_LEAD.
    F_LEAD's total holding = primary F shares (preferred) + tender shares
    (common from founders).
    """
    tender_shares = R["tender_shares_exact"]
    f1_sold = tender_shares * (SHARES_F1 / SHARES_FOUNDERS)
    f2_sold = tender_shares * (SHARES_F2 / SHARES_FOUNDERS)
    f3_sold = tender_shares * (SHARES_F3 / SHARES_FOUNDERS)
    founders_post = SHARES_FOUNDERS - tender_shares
    return {
        "tender_shares": tender_shares,
        "tender_shares_int": int_round(tender_shares),
        "f1_post": SHARES_F1 - f1_sold,
        "f2_post": SHARES_F2 - f2_sold,
        "f3_post": SHARES_F3 - f3_sold,
        "founders_post": founders_post,
        "f_lead_secondary_shares": tender_shares,  # all goes to F_LEAD
    }


# =============================================================================
# Waterfall / Nash equilibrium
# =============================================================================

CLASSES = ("series_f", "series_e", "series_d", "series_c", "series_b",
           "series_a", "seed")
TRACKED_CLASSES = CLASSES  # all 7 tracked in v4 (35 election strings)

LP_MAP = {
    "series_f": None,  # set per-call (depends on f_primary_raise)
    "series_e": E_LP,
    "series_d": D_LP,
    "series_c": C_LP,
    "series_b": B_LP,
    "series_a": A_LP,
    "seed": SEED_LP,
}


def class_asconv_shares(R: Dict) -> Dict[str, Decimal]:
    """As-converted share count per priced preferred class.

    For Series F, the LP holders are the primary investors (F_LEAD primary +
    F_FUND); the tender shares F_LEAD bought are common-equivalent (per §7.4
    default convention) and ride pro-rata, NOT as part of the F preference.
    """
    return {
        "series_f": R["new_F_shares"],
        "series_e": R["E_adj_shares"],
        "series_d": R["D_adj_v4_shares"],
        "series_c": C_SHARES,
        "series_b": B_SHARES,
        "series_a": A_SHARES,
        "seed": SEED_SHARES,
    }


def compute_waterfall_for_elections(exit_usd: Decimal,
                                    R: Dict,
                                    T_data: Dict,
                                    elect: Dict[str, str]) -> Dict:
    """Walk seniority F -> E -> D -> C -> B -> A -> Seed; preference takes
    LP first (capped at remaining pool), removed from pro-rata; convert
    classes ride pro-rata. Common = (founders post-tender + employees).
    Notes A and B always ride pro-rata (no election). Tender shares (held
    by F_LEAD) are common-equivalent and ride pro-rata.
    """
    asc = class_asconv_shares(R)
    f_primary_raise = R["f_primary_raise"]
    lp_map = dict(LP_MAP)
    lp_map["series_f"] = f_primary_raise

    remaining = exit_usd
    pref_paid = {c: ZERO for c in CLASSES}
    convert_shares = ZERO

    for c in CLASSES:  # seniority walk
        if elect[c] == "preference":
            pay = min(remaining, lp_map[c])
            pref_paid[c] = pay
            remaining -= pay
        else:
            convert_shares += asc[c]

    # Common-equivalent shares riding pro-rata:
    # - founders post-tender + employees (= SHARES_COMMON - tender_shares)
    # - F_LEAD's tender shares (transferred from founders, common-eq)
    # - notes A and B converted shares
    common_post_tender = SHARES_COMMON - T_data["tender_shares"]
    base_parts = (common_post_tender + T_data["f_lead_secondary_shares"]
                  + R["N_A"] + R["N_B"] + convert_shares)

    # Issued options (ITM iteration; strike=$0.50)
    pool = remaining
    opts_issued = R["options_issued_post_F"]
    itm_opts = False
    for _ in range(50):
        parts = base_parts
        p = pool
        if itm_opts:
            parts += opts_issued
            p += opts_issued * OPTIONS_STRIKE_WA
        if parts == 0:
            pps_common = ZERO
        else:
            pps_common = p / parts
        want_opts = pps_common > OPTIONS_STRIKE_WA
        if want_opts == itm_opts:
            break
        itm_opts = want_opts
    else:
        raise RuntimeError("ITM iteration did not converge")

    # Per-class dollars
    class_dollars: Dict[str, Decimal] = {}
    for c in CLASSES:
        if elect[c] == "preference":
            class_dollars[c] = pref_paid[c]
        else:
            class_dollars[c] = asc[c] * pps_common

    # Common (founders post-tender + employees, NOT including tender shares
    # held by F_LEAD which are reported separately under "f_lead_secondary"
    # but the spec only tracks waterfall.<exit>M.common.dollars, which we
    # take to include all common-equivalent NON-PREFERRED holders:
    # founders post-tender + employees + F_LEAD's tender shares.
    # F_LEAD's tender shares are reported as part of "common" because the
    # spec key list (§12.6) only has one common bucket.)
    common_dollars = (common_post_tender + T_data["f_lead_secondary_shares"]) * pps_common
    notes_dollars = (R["N_A"] + R["N_B"]) * pps_common
    options_dollars = (opts_issued * (pps_common - OPTIONS_STRIKE_WA)
                       if itm_opts else ZERO)

    # Our fund's share: fund_shares / new_F_shares of class_dollars["series_f"]
    if asc["series_f"] > 0:
        fund_dollars = class_dollars["series_f"] * (R["fund_shares"] / asc["series_f"])
    else:
        fund_dollars = ZERO

    total = (sum(class_dollars.values(), ZERO)
             + common_dollars + notes_dollars + options_dollars)

    return {
        "exit_usd": exit_usd,
        "class_dollars": class_dollars,
        "common_dollars": common_dollars,
        "notes_dollars": notes_dollars,
        "options_dollars": options_dollars,
        "fund_dollars": fund_dollars,
        "total_distributed": total,
        "pps_common": pps_common,
        "itm_opts": itm_opts,
        "elect": dict(elect),
        "pref_paid": pref_paid,
    }


def nash_equilibrium(exit_usd: Decimal, R: Dict, T_data: Dict) -> Dict:
    """Enumerate 2^7 = 128 election combos; find Nash equilibrium where
    every class weakly best-responds. Tie-break: prefer "convert" at
    knife-edge indifference (matches v3 §7).
    """
    options_list = ("preference", "convert")
    eps = d("1e-9")
    all_combos = []
    # Enumerate 2^7
    for sf in options_list:
        for se in options_list:
            for sd in options_list:
                for sc in options_list:
                    for sb in options_list:
                        for sa in options_list:
                            for ss in options_list:
                                elect = {
                                    "series_f": sf,
                                    "series_e": se,
                                    "series_d": sd,
                                    "series_c": sc,
                                    "series_b": sb,
                                    "series_a": sa,
                                    "seed": ss,
                                }
                                w = compute_waterfall_for_elections(
                                    exit_usd, R, T_data, elect)
                                all_combos.append(w)

    equilibria = []
    for w in all_combos:
        is_eq = True
        for c in CLASSES:
            alt_elect = dict(w["elect"])
            alt_elect[c] = "convert" if alt_elect[c] == "preference" else "preference"
            w_alt = None
            for ww in all_combos:
                if ww["elect"] == alt_elect:
                    w_alt = ww
                    break
            assert w_alt is not None
            if w["class_dollars"][c] < w_alt["class_dollars"][c] - eps:
                is_eq = False
                break
        if is_eq:
            equilibria.append(w)

    if not equilibria:
        raise RuntimeError(f"No Nash equilibrium at exit={exit_usd}")

    # Tie-break: prefer fewer prefs (more converts).
    def rank(w):
        return sum(1 for c in CLASSES if w["elect"][c] == "preference")
    equilibria.sort(key=rank)
    return equilibria[0]


# =============================================================================
# Main builder
# =============================================================================

def build_truth(f_primary_raise: Decimal) -> Dict:
    """Build all 167 truth keys for the given F primary raise."""

    # 1) Inherit v3
    v3 = solve_v3_post_e()
    assert_v3_inheritance(v3)

    # 2) Solve v4 round
    R = solve_round_v4(F_PRE_MONEY, f_primary_raise, ESOP_TARGET_F, v3)

    # 3) Sanity: branch assumption
    assert R["branch_A"] == "cap", \
        f"Note A expected to convert at cap, got {R['branch_A']}"
    assert R["branch_B"] == "discount", \
        f"Note B expected to convert at discount, got {R['branch_B']}"

    # 4) AD trigger sanity (D and E both > 1.0)
    assert R["r_E"] > ONE, f"AD on E expected to trigger; r_E={R['r_E']}"
    assert R["r_D_v4"] > ONE, f"AD on D expected to trigger; r_D_v4={R['r_D_v4']}"

    # 5) Tender
    T_data = compute_tender(R)

    # 6) Build keys
    out: Dict = {}
    T_F = R["T_F"]

    # ----- §12.1 Pro-forma ownership (15 keys) ----------------------------
    asc = class_asconv_shares(R)
    common_post_tender = SHARES_COMMON - T_data["tender_shares"]
    founders_post_tender = SHARES_FOUNDERS - T_data["tender_shares"]
    employees_post = SHARES_EMPLOYEES  # employees do not sell in tender
    # Series F primary investors = lead + our fund
    series_f_new_shares = R["new_F_shares"]  # both primary investors
    f_lead_secondary_shares = T_data["f_lead_secondary_shares"]
    note_a_shares = R["N_A"]
    note_b_shares = R["N_B"]
    fund_shares = R["fund_shares"]

    ownership = {
        "ownership.founders.pct": founders_post_tender / T_F,
        "ownership.employees.pct": employees_post / T_F,
        "ownership.options_issued.pct": R["options_issued_post_F"] / T_F,
        "ownership.options_unissued.pct": R["options_unissued_post_F"] / T_F,
        "ownership.seed.pct": asc["seed"] / T_F,
        "ownership.series_a.pct": asc["series_a"] / T_F,
        "ownership.series_b.pct": asc["series_b"] / T_F,
        "ownership.series_c.pct": asc["series_c"] / T_F,
        "ownership.series_d.pct": asc["series_d"] / T_F,
        "ownership.series_e.pct": asc["series_e"] / T_F,
        "ownership.note_a.pct": note_a_shares / T_F,
        "ownership.note_b.pct": note_b_shares / T_F,
        "ownership.series_f_new.pct": series_f_new_shares / T_F,
        "ownership.f_lead_secondary.pct": f_lead_secondary_shares / T_F,
        "ownership.our_fund.pct": fund_shares / T_F,
    }
    # Sum: all 15 fractions should sum to 1.0 (within 1e-9). Note our_fund
    # is a SUBSET of series_f_new (it's our $10M of the $30M = 1/3),
    # so summing all 15 gives 1.0 + (1/3 of series_f_new) — meaning the
    # primary 14 should sum to 1.0 and our_fund is an internal slice.
    # CANONICAL §12.1 says "Sum of all 15 keys = 1.000000". That can't be
    # right if our_fund ⊂ series_f_new. Re-reading: "ownership.our_fund.pct
    # = (1/3) × series_f_new". So they ARE overlapping.
    # We'll assert: sum of 14 (excluding our_fund) = 1.0.
    sum14 = sum((v for k, v in ownership.items()
                 if k != "ownership.our_fund.pct"), ZERO)
    assert abs(sum14 - ONE) < d("1e-9"), \
        f"Ownership sum (14, excl our_fund) != 1: {sum14}"

    for k, v in ownership.items():
        out[k] = to_float(round_pct(v))

    # ----- §12.2 Anti-dilution (12 keys) ----------------------------------
    # For seed/A/B/C: dormant -> emit OCP unchanged, ratio = 1.0
    # For D: NCP = NCP_D_v4 (v4 adjusted from OCP=NCP_D_v3); ratio = COMPOUND
    # For E: NCP = NCP_E (E's first AD); ratio = r_E
    ad_keys = [
        ("seed", SEED_PPS, ONE),
        ("series_a", A_PPS, ONE),
        ("series_b", B_PPS, ONE),
        ("series_c", C_PPS, ONE),
        ("series_d", R["NCP_D_v4"], R["r_D_compound"]),
        ("series_e", R["NCP_E"], R["r_E"]),
    ]
    for cls, ncp, ratio in ad_keys:
        out[f"antidilution.{cls}.new_conversion_price_usd"] = to_float(round_pps(ncp))
        out[f"antidilution.{cls}.new_ratio"] = to_float(round_ratio(ratio))

    # ----- §12.3 Note conversion (10 keys) --------------------------------
    out["notes.note_a.cap_implied_pps_usd"] = to_float(round_pps(R["cap_pps_A"]))
    out["notes.note_a.discount_implied_pps_usd"] = to_float(round_pps(R["disc_pps_A"]))
    out["notes.note_a.conversion_pps_usd"] = to_float(round_pps(R["conv_pps_A"]))
    out["notes.note_a.conversion_branch"] = R["branch_A"]
    out["notes.note_a.shares_issued"] = int_round(R["N_A"])

    out["notes.note_b.cap_implied_pps_usd"] = to_float(round_pps(R["cap_pps_B"]))
    out["notes.note_b.discount_implied_pps_usd"] = to_float(round_pps(R["disc_pps_B"]))
    out["notes.note_b.conversion_pps_usd"] = to_float(round_pps(R["conv_pps_B"]))
    out["notes.note_b.conversion_branch"] = R["branch_B"]
    out["notes.note_b.shares_issued"] = int_round(R["N_B"])

    # ----- §12.4 Tender mechanics (3 keys) --------------------------------
    out["tender.shares_transferred"] = int_round(T_data["tender_shares"])
    out["tender.usd_to_founders"] = to_float(round_money(F_SECONDARY_RAISE))
    out["tender.f_lead_holding_secondary_shares"] = int_round(T_data["f_lead_secondary_shares"])

    # ----- §12.5 Round solution (4 keys) ----------------------------------
    out["round.f_pps_usd"] = to_float(round_pps(R["F_PPS"]))
    out["round.post_f_fd_shares"] = int_round(R["T_F"])
    out["round.our_fund_shares"] = int_round(R["fund_shares"])
    out["round.f_lead_primary_shares"] = int_round(R["f_lead_primary_shares"])

    # ----- §12.6 Waterfall (95 keys) --------------------------------------
    waterfall_cache = {}
    for label, exit_val in zip(EXITS_LABEL, EXITS_USD):
        w = nash_equilibrium(exit_val, R, T_data)
        waterfall_cache[label] = w
        prefix = f"waterfall.{label}M"
        cd = w["class_dollars"]
        out[f"{prefix}.series_f.dollars"] = to_float(round_money(cd["series_f"]))
        out[f"{prefix}.series_e.dollars"] = to_float(round_money(cd["series_e"]))
        out[f"{prefix}.series_d.dollars"] = to_float(round_money(cd["series_d"]))
        out[f"{prefix}.series_c.dollars"] = to_float(round_money(cd["series_c"]))
        out[f"{prefix}.series_b.dollars"] = to_float(round_money(cd["series_b"]))
        out[f"{prefix}.series_a.dollars"] = to_float(round_money(cd["series_a"]))
        out[f"{prefix}.seed.dollars"] = to_float(round_money(cd["seed"]))
        out[f"{prefix}.notes.dollars"] = to_float(round_money(w["notes_dollars"]))
        out[f"{prefix}.common.dollars"] = to_float(round_money(w["common_dollars"]))
        out[f"{prefix}.options_issued.dollars"] = to_float(round_money(w["options_dollars"]))
        out[f"{prefix}.our_fund.dollars"] = to_float(round_money(w["fund_dollars"]))
        out[f"{prefix}.total_distributed"] = to_float(round_money(w["total_distributed"]))
        for c in TRACKED_CLASSES:
            out[f"{prefix}.{c}.election"] = w["elect"][c]
        # Conservation
        td = w["total_distributed"]
        assert abs(td - exit_val) < d("1"), \
            f"Waterfall conservation failed at {label}M: td={td} exit={exit_val}"

    # ----- §12.7 Sensitivity (18 keys) ------------------------------------
    esop_grid = [("11", d("0.11")), ("13", d("0.13")), ("15", d("0.15"))]
    pre_grid = [("1200", d(1_200_000_000)),
                ("1400", d(1_400_000_000)),
                ("1800", d(1_800_000_000))]
    ref_exit = d(2_000_000_000)
    for esop_label, esop_val in esop_grid:
        for pre_label, pre_val in pre_grid:
            try:
                Rs = solve_round_v4(pre_val, f_primary_raise, esop_val, v3)
                T_data_s = compute_tender(Rs)
                our_pct = Rs["fund_shares"] / Rs["T_F"]
                w = nash_equilibrium(ref_exit, Rs, T_data_s)
                fund_dollars = w["fund_dollars"]
            except Exception as e:
                # If solver doesn't converge for a sensitivity cell, surface
                raise RuntimeError(
                    f"Sensitivity cell esop={esop_label} pre={pre_label}M failed: {e}")
            out[f"sensitivity.esop_{esop_label}.pre_{pre_label}M.our_pct"] = \
                to_float(round_pct(our_pct))
            out[f"sensitivity.esop_{esop_label}.pre_{pre_label}M.dollars_at_2000M_exit"] = \
                to_float(round_money(fund_dollars))

    # ----- §12.8 Return solver (10 keys) ----------------------------------
    # Required Series F ownership pct of FD to return target dollars.
    # series_f_pct_of_fd = asc["series_f"] / T_F
    # required pct (of FD) = target / series_f_dollars * series_f_pct_of_fd
    series_f_pct_of_fd = asc["series_f"] / T_F
    for label, exit_val in zip(EXITS_LABEL, EXITS_USD):
        w = waterfall_cache[label]
        sf_dollars = w["class_dollars"]["series_f"]
        for mult_label, target in [("3x", d(30_000_000)), ("4x", d(40_000_000))]:
            if sf_dollars <= 0:
                req = d("-1")
            else:
                req = target / sf_dollars * series_f_pct_of_fd
            out[f"solver.{mult_label}.exit_{label}M.required_pct"] = \
                to_float(round_pct(req))

    # ----- Final sanity gates ---------------------------------------------
    # Key count = 167 exact
    assert len(out) == 167, f"Expected 167 keys, got {len(out)}"
    # Note A branch == cap; Note B branch == discount
    assert out["notes.note_a.conversion_branch"] == "cap"
    assert out["notes.note_b.conversion_branch"] == "discount"
    # AD ratios
    assert out["antidilution.series_d.new_ratio"] > 1.0
    assert out["antidilution.series_e.new_ratio"] > 1.0
    # Waterfall conservation per exit
    for label, exit_val in zip(EXITS_LABEL, EXITS_USD):
        td = d(str(out[f"waterfall.{label}M.total_distributed"]))
        assert abs(td - exit_val) < d("1"), \
            f"Conservation fail @ {label}M: td={td}"

    return out, R, T_data


# =============================================================================
# Main entry point
# =============================================================================

def main():
    here = os.path.dirname(os.path.abspath(__file__))

    # (a) canonical (raise = $30M)
    print("=" * 70)
    print("Building canonical truth (F primary raise = $30M)")
    print("=" * 70)
    out, R, T_data = build_truth(F_PRIMARY_RAISE_BASE)
    canonical_path = os.path.join(here, "truth.json")
    with open(canonical_path, "w") as f:
        json.dump(out, f, indent=2, sort_keys=True)
    print(f"Wrote {canonical_path}")

    # (b) perturbed (raise = $33M)
    print()
    print("=" * 70)
    print("Building perturbed truth (F primary raise = $33M)")
    print("=" * 70)
    out_p, R_p, T_data_p = build_truth(d(33_000_000))
    perturbed_path = os.path.join(here, "truth_perturbed.json")
    with open(perturbed_path, "w") as f:
        json.dump(out_p, f, indent=2, sort_keys=True)
    print(f"Wrote {perturbed_path}")

    # (c) one-line summary
    print()
    print("=" * 70)
    print("SUMMARY")
    print("=" * 70)
    print(
        f"[canonical] F_PPS={round_pps(R['F_PPS'])} "
        f"T_F={int_round(R['T_F'])} "
        f"AD_D={round_ratio(R['r_D_v4'])} AD_E={round_ratio(R['r_E'])} "
        f"NoteA={R['branch_A']}({int_round(R['N_A'])}) "
        f"NoteB={R['branch_B']}({int_round(R['N_B'])}) "
        f"keys={len(out)}"
    )
    print(
        f"[perturbed] F_PPS={round_pps(R_p['F_PPS'])} "
        f"T_F={int_round(R_p['T_F'])} "
        f"AD_D={round_ratio(R_p['r_D_v4'])} AD_E={round_ratio(R_p['r_E'])} "
        f"NoteA={R_p['branch_A']}({int_round(R_p['N_A'])}) "
        f"NoteB={R_p['branch_B']}({int_round(R_p['N_B'])}) "
        f"keys={len(out_p)}"
    )


if __name__ == "__main__":
    main()
