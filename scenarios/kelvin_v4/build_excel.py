"""
Excel ground-truth builder for Kelvin Dynamics scenario kelvin_v4.

Produces:
  truth/ground_truth.xlsx

Tabs (in this exact order):
  1. Assumptions  - all locked inputs from CANONICAL_INPUTS.md sections 3-6,
                    plus a 50-row unrolled fixed-point iteration block that
                    converges (F_PPS, T_F, r_E_v4, r_D_v4, N_A_shares,
                    N_B_shares).  No circular references.
  2. Pro-Forma    - post-F ownership lines (15), AD block (6 classes x 3),
                    Notes block (2 notes x 5), Tender block (3), and the
                    Round-solution block (4).
  3. Waterfall    - 5 exit columns x 7 priced classes; election cells use
                    IF(pref > convert, "preference", "convert"); dollar cells
                    are IF-driven by election.  Reverse-chronological cascade
                    F -> E -> D -> C -> B -> A -> Seed -> notes -> common ->
                    options.
  4. Sensitivity  - 3 x 3 grid (ESOP refresh x pre-money F).  Each cell owns
                    its own ESOP/pre inputs and re-runs the full round
                    fixed-point.  Outputs our $10M check post-F % and $ at
                    $2B exit.
  5. Return-Solver - 5 exits x 2 multiples (3x, 4x).  Required Series F
                    ownership %.

Every output cell is a formula referencing Assumptions inputs or another
formula cell.  No hardcoded numeric outputs outside Assumptions.

Defined names: 167+ workbook-level names per CANONICAL_INPUTS.md sec 14.5
convention (dot -> underscore).
"""
from __future__ import annotations

from pathlib import Path

import openpyxl
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter


HERE = Path(__file__).resolve().parent
TRUTH_DIR = HERE / "truth"
TRUTH_DIR.mkdir(exist_ok=True)


# ---------------------------------------------------------------------------
# Locked inputs (from CANONICAL_INPUTS.md)
# ---------------------------------------------------------------------------

# Pre-F (post-E) state  -- v3's solved fixed point
PRE_F_FOUNDERS_F1 = 4_000_000
PRE_F_FOUNDERS_F2 = 2_500_000
PRE_F_FOUNDERS_F3 = 1_500_000
PRE_F_FOUNDERS_TOTAL = 8_000_000
PRE_F_RSAS = 1_000_000
PRE_F_COMMON_TOTAL = 9_000_000  # founders + RSAs

PRE_F_SEED_SHARES = 5_000_000
PRE_F_A_SHARES = 5_000_000
PRE_F_B_SHARES = 3_000_000
PRE_F_C_SHARES = 2_400_000
# v3 post-E full-precision values (matches scenarios/kelvin_v3 truth.py
# solver to 15+ significant digits). Storing rounded integers here
# previously injected ~250-share precision drift into v4's fixed point,
# which cascaded into 1e-6+ pct deltas and ~1.7-share Note divergence.
# These are the Decimal solutions truncated to float64 precision.
PRE_F_D_SHARES_ADJ = 1720627.3980650043  # = 1,666,667 * r_D(v3)
PRE_F_E_SHARES = 6102950.3266507019      # v3 fixed-point E shares (full)
PRE_F_OPTIONS_ISSUED = 2_500_000
# pool_total_v3 = 0.12 * T_E; unissued = pool - issued
PRE_F_OPTIONS_UNISSUED = 1894124.235188505  # = 0.12 * T_E - 2.5M
PRE_F_FD_TOTAL = 36617701.9599042116  # T_E full-precision

# Original conversion prices (OCPs) for AD trigger tests
OCP_SEED = 1.00
OCP_A = 3.00
OCP_B = 10.00
OCP_C = 25.00
OCP_D_V4 = 58.11834689628838  # post-E NCP_D, full precision
OCP_E = 49.15655280527901      # E_PPS = post_money_E / T_E, full precision
E_PPS_LOCK = 49.15655280527901

# Original D shares before any adjustment (for "compound ratio" math)
D_SHARES_ORIGINAL = 1_666_667

# Per-class invested $ (preferences)
PREF_SEED = 5_000_000
PREF_A = 15_000_000
PREF_B = 30_000_000
PREF_C = 60_000_000
PREF_D = 100_000_000
PREF_E = 300_000_000

# Series F terms
F_PRE_MONEY = 1_400_000_000
F_PRIMARY_RAISE = 30_000_000
F_SECONDARY_RAISE = 15_000_000
F_TENDER_MULT = 0.85
F_ESOP_TARGET = 0.13
F_LEAD_PRIMARY_USD = 20_000_000
OUR_FUND_USD = 10_000_000

# Convertible note terms
NOTE_A_PRINCIPAL = 5_000_000
NOTE_A_CAP = 1_000_000_000
NOTE_A_DISCOUNT = 0.20

NOTE_B_PRINCIPAL = 3_000_000
NOTE_B_CAP = 1_600_000_000
NOTE_B_DISCOUNT = 0.15

# Options strike (carried from v3)
OPTIONS_STRIKE = 0.50

# Exits
EXIT_VALUES = [700_000_000, 1_200_000_000, 2_000_000_000, 4_000_000_000, 8_000_000_000]
EXIT_TAGS = ["700M", "1200M", "2000M", "4000M", "8000M"]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def col(n):
    return get_column_letter(n)


def addr(sheet: str, cell: str) -> str:
    """Absolute reference to cell on sheet."""
    return f"{sheet}!${cell[0]}${cell[1:]}"


# ---------------------------------------------------------------------------
# Build ground_truth.xlsx
# ---------------------------------------------------------------------------

def build_ground_truth():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws_as = wb.create_sheet("Assumptions")
    ws_pf = wb.create_sheet("Pro-Forma")
    ws_wf = wb.create_sheet("Waterfall")
    ws_se = wb.create_sheet("Sensitivity")
    ws_rs = wb.create_sheet("Return-Solver")

    defined_names: list[tuple[str, str]] = []  # (name, "Sheet!$C$R")

    def add_name(name: str, sheet: str, cell: str):
        ref = f"'{sheet}'!${cell[0]}${cell[1:]}" if "-" in sheet else f"{sheet}!${cell[0]}${cell[1:]}"
        defined_names.append((name, ref))

    # =======================================================================
    # 1) Assumptions tab
    # =======================================================================
    ws_as["A1"] = "Kelvin Dynamics v4 - locked inputs and derived constants"

    # ----- Pre-F state (v3 inheritance, sec 3) -----
    ws_as["A3"] = "Pre-F (post-E) state, locked from v3 truth.json"
    ws_as["A4"] = "F1 founder shares";       ws_as["B4"] = PRE_F_FOUNDERS_F1
    ws_as["A5"] = "F2 founder shares";       ws_as["B5"] = PRE_F_FOUNDERS_F2
    ws_as["A6"] = "F3 founder shares";       ws_as["B6"] = PRE_F_FOUNDERS_F3
    ws_as["A7"] = "Founders total";          ws_as["B7"] = "=B4+B5+B6"
    ws_as["A8"] = "RSAs (E1-E10)";           ws_as["B8"] = PRE_F_RSAS
    ws_as["A9"] = "Common total";            ws_as["B9"] = "=B7+B8"

    ws_as["A11"] = "Seed shares (as-conv ratio=1)";       ws_as["B11"] = PRE_F_SEED_SHARES
    ws_as["A12"] = "Series A shares";                     ws_as["B12"] = PRE_F_A_SHARES
    ws_as["A13"] = "Series B shares";                     ws_as["B13"] = PRE_F_B_SHARES
    ws_as["A14"] = "Series C shares";                     ws_as["B14"] = PRE_F_C_SHARES
    ws_as["A15"] = "Series D shares (v3-adj base)";       ws_as["B15"] = PRE_F_D_SHARES_ADJ
    ws_as["A16"] = "Series E shares (post-v3 fixed pt)";  ws_as["B16"] = PRE_F_E_SHARES
    ws_as["A17"] = "Options issued (post-E)";             ws_as["B17"] = PRE_F_OPTIONS_ISSUED
    ws_as["A18"] = "Options unissued (post-E)";           ws_as["B18"] = PRE_F_OPTIONS_UNISSUED
    ws_as["A19"] = "Pre-F FD total (T_E)";                ws_as["B19"] = PRE_F_FD_TOTAL

    # ----- Prior-round terms (sec 3) -----
    ws_as["A21"] = "Prior-round OCPs and preferences"
    ws_as["A22"] = "Seed OCP";       ws_as["B22"] = OCP_SEED
    ws_as["A23"] = "A OCP";          ws_as["B23"] = OCP_A
    ws_as["A24"] = "B OCP";          ws_as["B24"] = OCP_B
    ws_as["A25"] = "C OCP";          ws_as["B25"] = OCP_C
    ws_as["A26"] = "D OCP (v4: post-E NCP_D)"; ws_as["B26"] = OCP_D_V4
    ws_as["A27"] = "E OCP (= E PPS)"; ws_as["B27"] = OCP_E
    ws_as["A28"] = "Seed preference $"; ws_as["B28"] = PREF_SEED
    ws_as["A29"] = "A preference $";    ws_as["B29"] = PREF_A
    ws_as["A30"] = "B preference $";    ws_as["B30"] = PREF_B
    ws_as["A31"] = "C preference $";    ws_as["B31"] = PREF_C
    ws_as["A32"] = "D preference $";    ws_as["B32"] = PREF_D
    ws_as["A33"] = "E preference $";    ws_as["B33"] = PREF_E

    ws_as["A35"] = "D shares (original, pre any AD)"; ws_as["B35"] = D_SHARES_ORIGINAL
    ws_as["A36"] = "Compound r_D(v3) = adj/orig";     ws_as["B36"] = "=B15/B35"
    # v3's r_D is structural, not solved here

    ws_as["A37"] = "Options strike"; ws_as["B37"] = OPTIONS_STRIKE

    # ----- Series F terms (sec 5) -----
    ws_as["A39"] = "Series F terms"
    ws_as["A40"] = "F pre-money $";          ws_as["B40"] = F_PRE_MONEY
    ws_as["A41"] = "F primary raise $";      ws_as["B41"] = F_PRIMARY_RAISE
    ws_as["A42"] = "F post-money $";         ws_as["B42"] = "=B40+B41"
    ws_as["A43"] = "F secondary tender $";   ws_as["B43"] = F_SECONDARY_RAISE
    ws_as["A44"] = "Tender PPS multiplier";  ws_as["B44"] = F_TENDER_MULT
    ws_as["A45"] = "ESOP target post-F";     ws_as["B45"] = F_ESOP_TARGET
    ws_as["A46"] = "F_LEAD primary $";       ws_as["B46"] = F_LEAD_PRIMARY_USD
    ws_as["A47"] = "Our fund check $";       ws_as["B47"] = OUR_FUND_USD

    # ----- Note terms (sec 6) -----
    ws_as["A49"] = "Convertible notes"
    ws_as["A50"] = "Note A principal";       ws_as["B50"] = NOTE_A_PRINCIPAL
    ws_as["A51"] = "Note A cap";             ws_as["B51"] = NOTE_A_CAP
    ws_as["A52"] = "Note A discount";        ws_as["B52"] = NOTE_A_DISCOUNT
    ws_as["A53"] = "Note B principal";       ws_as["B53"] = NOTE_B_PRINCIPAL
    ws_as["A54"] = "Note B cap";             ws_as["B54"] = NOTE_B_CAP
    ws_as["A55"] = "Note B discount";        ws_as["B55"] = NOTE_B_DISCOUNT

    # ----- Exit values (sec 8) -----
    ws_as["A57"] = "Exit values"
    for i, (tag, ev) in enumerate(zip(EXIT_TAGS, EXIT_VALUES)):
        r = 58 + i
        ws_as[f"A{r}"] = f"Exit {tag}"
        ws_as[f"B{r}"] = ev
    # Exit cell rows: 58..62

    # ----- Pre-F broad-base for AD (sec 7.2) -----
    ws_as["A64"] = "Broad-base A (pre-F FD) for AD"; ws_as["B64"] = "=B19"
    # Pre-F option pool (issued + unissued post-E) — used by sensitivity to
    # detect the ESOP clamp branch when esop_target * T_F < pool_pre_F.
    ws_as["A65"] = "Pre-F option pool (= post-E pool)"
    ws_as["B65"] = "=B17+B18"

    # ----- Round fixed-point block (50-iter unroll) -----
    # Five unknowns: F_PPS, T_F, r_E_v4, r_D_v4, N_A_shares, N_B_shares.
    # Reduction: T_F = post / F_PPS (= 1.43B / F_PPS).
    # Per pass given prev (r_E, r_D, N_A, N_B):
    #   E_adj = 6,102,667 * r_E
    #   D_adj = 1,720,627 * r_D   (this base is already-once-adjusted)
    #   sum_pre_classes_static = 9M (common) + 5M (Seed) + 5M (A) + 3M (B) + 2.4M (C) = 24.4M
    #   F_PPS = (post*(1-ESOP) - primary_raise) / (24.4M + D_adj + E_adj + N_A + N_B)
    #   T_F   = post / F_PPS
    #   new_F = primary_raise / F_PPS
    #   pre_money_FD = T_F - new_F
    #   For Note A: cap_implied = cap_A / pre_money_FD; disc_implied = F_PPS*(1-disc_A)
    #               conv_pps_A = MIN(cap_implied, disc_implied)
    #               N_A = principal_A / conv_pps_A
    #   For Note B: similar
    #   r_E_new = IF(F_PPS < OCP_E, OCP_E*(A_BB + raise/OCP_E)/(A_BB + new_F), OCP_E) ; ratio = OCP_E / NCP_E or 1
    #   r_D_new = IF(F_PPS < OCP_D, ...)
    #   Note: only new_F is in C of the AD formula (notes excluded per sec 7.2)
    # Initial guess: r_E = 1, r_D = 1, N_A = 0, N_B = 0
    # 50 iterations is plenty; it converges in ~10-15 passes.
    ws_as["A66"] = "Fixed-point iteration (50 unrolled passes)"
    ws_as["A67"] = "iter"
    ws_as["B67"] = "r_E_prev"
    ws_as["C67"] = "r_D_prev"
    ws_as["D67"] = "N_A_prev"
    ws_as["E67"] = "N_B_prev"
    ws_as["F67"] = "E_adj"
    ws_as["G67"] = "D_adj"
    ws_as["H67"] = "F_PPS"
    ws_as["I67"] = "T_F"
    ws_as["J67"] = "new_F"
    ws_as["K67"] = "pre_money_FD"
    ws_as["L67"] = "cap_impl_A"
    ws_as["M67"] = "disc_impl_A"
    ws_as["N67"] = "conv_pps_A"
    ws_as["O67"] = "N_A_new"
    ws_as["P67"] = "cap_impl_B"
    ws_as["Q67"] = "disc_impl_B"
    ws_as["R67"] = "conv_pps_B"
    ws_as["S67"] = "N_B_new"
    ws_as["T67"] = "NCP_E_new"
    ws_as["U67"] = "r_E_new"
    ws_as["V67"] = "NCP_D_new"
    ws_as["W67"] = "r_D_new"

    IT_START = 68
    N_ITERS = 50
    # absolute references to inputs
    POST = "$B$42"            # post-money 1.43B
    PRIMARY = "$B$41"         # 30M
    ESOP = "$B$45"            # 0.13
    A_BB_REF = "$B$64"        # 36,617,961
    OCP_E_REF = "$B$27"
    OCP_D_REF = "$B$26"
    SUM_PRE_STATIC = "($B$9+$B$11+$B$12+$B$13+$B$14)"  # 9+5+5+3+2.4 = 24.4M
    E_BASE = "$B$16"          # 6,102,667
    D_BASE_ADJ = "$B$15"      # 1,720,627
    NOTE_A_CAP_REF = "$B$51"
    NOTE_A_DISC_REF = "$B$52"
    NOTE_A_PRIN_REF = "$B$50"
    NOTE_B_CAP_REF = "$B$54"
    NOTE_B_DISC_REF = "$B$55"
    NOTE_B_PRIN_REF = "$B$53"

    for i in range(N_ITERS):
        r = IT_START + i
        ws_as[f"A{r}"] = i
        if i == 0:
            ws_as[f"B{r}"] = 1.0     # r_E init
            ws_as[f"C{r}"] = 1.0     # r_D init
            ws_as[f"D{r}"] = 0       # N_A init
            ws_as[f"E{r}"] = 0       # N_B init
        else:
            ws_as[f"B{r}"] = f"=U{r-1}"
            ws_as[f"C{r}"] = f"=W{r-1}"
            ws_as[f"D{r}"] = f"=O{r-1}"
            ws_as[f"E{r}"] = f"=S{r-1}"
        # E_adj, D_adj
        ws_as[f"F{r}"] = f"={E_BASE}*B{r}"
        ws_as[f"G{r}"] = f"={D_BASE_ADJ}*C{r}"
        # F_PPS = (post*(1-ESOP) - primary) / (24.4M + D_adj + E_adj + N_A + N_B)
        ws_as[f"H{r}"] = (
            f"=({POST}*(1-{ESOP})-{PRIMARY})/"
            f"({SUM_PRE_STATIC}+G{r}+F{r}+D{r}+E{r})"
        )
        # T_F = post / F_PPS
        ws_as[f"I{r}"] = f"={POST}/H{r}"
        # new_F = primary / F_PPS
        ws_as[f"J{r}"] = f"={PRIMARY}/H{r}"
        # pre_money_FD = T_F - new_F
        ws_as[f"K{r}"] = f"=I{r}-J{r}"
        # Note A
        ws_as[f"L{r}"] = f"={NOTE_A_CAP_REF}/K{r}"          # cap-implied
        ws_as[f"M{r}"] = f"=H{r}*(1-{NOTE_A_DISC_REF})"     # disc-implied
        ws_as[f"N{r}"] = f"=MIN(L{r},M{r})"                  # conv pps
        ws_as[f"O{r}"] = f"={NOTE_A_PRIN_REF}/N{r}"          # N_A shares
        # Note B
        ws_as[f"P{r}"] = f"={NOTE_B_CAP_REF}/K{r}"
        ws_as[f"Q{r}"] = f"=H{r}*(1-{NOTE_B_DISC_REF})"
        ws_as[f"R{r}"] = f"=MIN(P{r},Q{r})"
        ws_as[f"S{r}"] = f"={NOTE_B_PRIN_REF}/R{r}"
        # AD on E
        # NCP_E = OCP_E * (A_BB + raise/OCP_E) / (A_BB + C); C = new_F shares
        ws_as[f"T{r}"] = f"={OCP_E_REF}*({A_BB_REF}+{PRIMARY}/{OCP_E_REF})/({A_BB_REF}+J{r})"
        ws_as[f"U{r}"] = f"=IF(H{r}<{OCP_E_REF},{OCP_E_REF}/T{r},1)"
        # AD on D
        ws_as[f"V{r}"] = f"={OCP_D_REF}*({A_BB_REF}+{PRIMARY}/{OCP_D_REF})/({A_BB_REF}+J{r})"
        ws_as[f"W{r}"] = f"=IF(H{r}<{OCP_D_REF},{OCP_D_REF}/V{r},1)"

    CONV_ROW = IT_START + N_ITERS - 1  # row 117

    # Converged values pulled out for clean references
    ws_as["A119"] = "Converged round solution (row 117)"
    ws_as["A120"] = "F_PPS";              ws_as["B120"] = f"=H{CONV_ROW}"
    ws_as["A121"] = "T_F";                ws_as["B121"] = f"=I{CONV_ROW}"
    ws_as["A122"] = "new_F";              ws_as["B122"] = f"=J{CONV_ROW}"
    ws_as["A123"] = "pre_money_FD";       ws_as["B123"] = f"=K{CONV_ROW}"
    ws_as["A124"] = "r_E_v4";             ws_as["B124"] = f"=U{CONV_ROW}"
    ws_as["A125"] = "r_D_v4";             ws_as["B125"] = f"=W{CONV_ROW}"
    ws_as["A126"] = "NCP_E_v4";           ws_as["B126"] = f"=T{CONV_ROW}"
    ws_as["A127"] = "NCP_D_v4";           ws_as["B127"] = f"=V{CONV_ROW}"
    ws_as["A128"] = "E_adj shares";       ws_as["B128"] = f"=F{CONV_ROW}"
    ws_as["A129"] = "D_adj shares";       ws_as["B129"] = f"=G{CONV_ROW}"
    ws_as["A130"] = "Note A cap-impl PPS";    ws_as["B130"] = f"=L{CONV_ROW}"
    ws_as["A131"] = "Note A disc-impl PPS";   ws_as["B131"] = f"=M{CONV_ROW}"
    ws_as["A132"] = "Note A conv PPS";        ws_as["B132"] = f"=N{CONV_ROW}"
    ws_as["A133"] = "Note A shares";          ws_as["B133"] = f"=O{CONV_ROW}"
    ws_as["A134"] = "Note B cap-impl PPS";    ws_as["B134"] = f"=P{CONV_ROW}"
    ws_as["A135"] = "Note B disc-impl PPS";   ws_as["B135"] = f"=Q{CONV_ROW}"
    ws_as["A136"] = "Note B conv PPS";        ws_as["B136"] = f"=R{CONV_ROW}"
    ws_as["A137"] = "Note B shares";          ws_as["B137"] = f"=S{CONV_ROW}"

    # F primary investor shares
    ws_as["A139"] = "F_LEAD primary shares"; ws_as["B139"] = "=$B$46/$B$120"
    ws_as["A140"] = "Our fund shares";       ws_as["B140"] = "=$B$47/$B$120"
    ws_as["A141"] = "F primary total shares (= new_F)"; ws_as["B141"] = "=$B$122"

    # Tender shares
    ws_as["A143"] = "Tender shares transferred";
    ws_as["B143"] = "=$B$43/($B$44*$B$120)"
    ws_as["A144"] = "Tender $ to founders"; ws_as["B144"] = "=$B$43"
    # Tender comes pro rata across F1/F2/F3 by their pre-F holdings
    ws_as["A145"] = "Tender shares from F1 (50%)";    ws_as["B145"] = "=$B$143*0.5"
    ws_as["A146"] = "Tender shares from F2 (31.25%)"; ws_as["B146"] = "=$B$143*0.3125"
    ws_as["A147"] = "Tender shares from F3 (18.75%)"; ws_as["B147"] = "=$B$143*0.1875"

    # Post-tender founders total
    ws_as["A149"] = "Founders post-tender total"; ws_as["B149"] = "=$B$7-$B$143"

    # ESOP top-up (clamped)
    ws_as["A151"] = "Options pool target ($ESOP * T_F)"; ws_as["B151"] = "=$B$45*$B$121"
    ws_as["A152"] = "Options unissued post-F (clamped)"
    ws_as["B152"] = "=MAX(0, $B$151-$B$17)"
    ws_as["A153"] = "Options issued post-F (unchanged)"; ws_as["B153"] = "=$B$17"
    ws_as["A154"] = "Options total post-F"; ws_as["B154"] = "=$B$152+$B$153"

    # F_LEAD post-tender holdings
    ws_as["A156"] = "F_LEAD primary shares (preferred)"; ws_as["B156"] = "=$B$139"
    ws_as["A157"] = "F_LEAD secondary shares (common from founders)"; ws_as["B157"] = "=$B$143"

    # Workbook-level defined names: round solution
    add_name("round_f_pps_usd", "Assumptions", "B120")
    add_name("round_post_f_fd_shares", "Assumptions", "B121")
    add_name("round_our_fund_shares", "Assumptions", "B140")
    add_name("round_f_lead_primary_shares", "Assumptions", "B139")

    # =======================================================================
    # 2) Pro-Forma tab
    # =======================================================================
    ws_pf["A1"] = "Pro-Forma cap table (post Series F close)"
    ws_pf["A3"] = "Class"
    ws_pf["B3"] = "Shares (FD)"
    ws_pf["C3"] = "%"

    # 15 ownership lines
    # Founders post-tender
    ws_pf["A4"] = "Founders (F1+F2+F3, post-tender)"
    ws_pf["B4"] = "=Assumptions!$B$149"
    ws_pf["C4"] = "=B4/Assumptions!$B$121"

    ws_pf["A5"] = "Employees (RSAs)"
    ws_pf["B5"] = "=Assumptions!$B$8"
    ws_pf["C5"] = "=B5/Assumptions!$B$121"

    ws_pf["A6"] = "Options issued"
    ws_pf["B6"] = "=Assumptions!$B$153"
    ws_pf["C6"] = "=B6/Assumptions!$B$121"

    ws_pf["A7"] = "Options unissued"
    ws_pf["B7"] = "=Assumptions!$B$152"
    ws_pf["C7"] = "=B7/Assumptions!$B$121"

    ws_pf["A8"] = "Seed (as-conv, ratio=1)"
    ws_pf["B8"] = "=Assumptions!$B$11"
    ws_pf["C8"] = "=B8/Assumptions!$B$121"

    ws_pf["A9"] = "Series A (as-conv, ratio=1)"
    ws_pf["B9"] = "=Assumptions!$B$12"
    ws_pf["C9"] = "=B9/Assumptions!$B$121"

    ws_pf["A10"] = "Series B (as-conv, ratio=1)"
    ws_pf["B10"] = "=Assumptions!$B$13"
    ws_pf["C10"] = "=B10/Assumptions!$B$121"

    ws_pf["A11"] = "Series C (as-conv, ratio=1)"
    ws_pf["B11"] = "=Assumptions!$B$14"
    ws_pf["C11"] = "=B11/Assumptions!$B$121"

    ws_pf["A12"] = "Series D (compound v3*v4 ratio)"
    ws_pf["B12"] = "=Assumptions!$B$129"
    ws_pf["C12"] = "=B12/Assumptions!$B$121"

    ws_pf["A13"] = "Series E (v4-adjusted)"
    ws_pf["B13"] = "=Assumptions!$B$128"
    ws_pf["C13"] = "=B13/Assumptions!$B$121"

    ws_pf["A14"] = "Note A converted"
    ws_pf["B14"] = "=Assumptions!$B$133"
    ws_pf["C14"] = "=B14/Assumptions!$B$121"

    ws_pf["A15"] = "Note B converted"
    ws_pf["B15"] = "=Assumptions!$B$137"
    ws_pf["C15"] = "=B15/Assumptions!$B$121"

    ws_pf["A16"] = "Series F primary new (lead + our fund)"
    ws_pf["B16"] = "=Assumptions!$B$141"
    ws_pf["C16"] = "=B16/Assumptions!$B$121"

    ws_pf["A17"] = "F_LEAD secondary (common from founders)"
    ws_pf["B17"] = "=Assumptions!$B$157"
    ws_pf["C17"] = "=B17/Assumptions!$B$121"

    ws_pf["A18"] = "Our fund (subset of F primary)"
    ws_pf["B18"] = "=Assumptions!$B$140"
    ws_pf["C18"] = "=B18/Assumptions!$B$121"

    # Sum check (ex our_fund and ex f_lead_secondary which is included in F primary count?)
    # Sum: founders + RSAs + options_issued + options_unissued + seed + A + B + C + D + E + N_A
    #      + N_B + F_primary_new + F_LEAD_secondary
    # Note: F_LEAD_secondary shares are NOT in T_F (they're transferred from founders, not new).
    # T_F should = founders_post_tender + RSAs + opt_iss + opt_uniss + seed + A + B + C + D_adj
    #             + E_adj + N_A + N_B + F_primary_new
    # So including F_LEAD_secondary in the sum would double-count: founders_post_tender already
    # excludes those shares, but secondary holds them as common. Let's verify:
    #   founders_post_tender = 8M - tender_shares
    #   F_LEAD_secondary = tender_shares
    # Sum: (8M - tender) + tender = 8M (the total founders' original common). So including both
    # F_LEAD_secondary AND founders_post_tender DOES sum correctly to 8M founder common shares.
    # Therefore C4..C17 should sum to 1.0 (excluding our_fund row 18 since it's subset).
    ws_pf["A20"] = "Sum check (rows 4-17, excludes our_fund subset)"
    ws_pf["B20"] = "=SUM(B4:B17)"
    ws_pf["C20"] = "=SUM(C4:C17)"

    add_name("ownership_founders_pct",         "Pro-Forma", "C4")
    add_name("ownership_employees_pct",        "Pro-Forma", "C5")
    add_name("ownership_options_issued_pct",   "Pro-Forma", "C6")
    add_name("ownership_options_unissued_pct", "Pro-Forma", "C7")
    add_name("ownership_seed_pct",             "Pro-Forma", "C8")
    add_name("ownership_series_a_pct",         "Pro-Forma", "C9")
    add_name("ownership_series_b_pct",         "Pro-Forma", "C10")
    add_name("ownership_series_c_pct",         "Pro-Forma", "C11")
    add_name("ownership_series_d_pct",         "Pro-Forma", "C12")
    add_name("ownership_series_e_pct",         "Pro-Forma", "C13")
    add_name("ownership_note_a_pct",           "Pro-Forma", "C14")
    add_name("ownership_note_b_pct",           "Pro-Forma", "C15")
    add_name("ownership_series_f_new_pct",     "Pro-Forma", "C16")
    add_name("ownership_f_lead_secondary_pct", "Pro-Forma", "C17")
    add_name("ownership_our_fund_pct",         "Pro-Forma", "C18")

    # ----- AD block (6 classes x 3 fields: trigger, NCP, ratio) -----
    # Per spec: AD trigger cells must be IF formulas comparing F_PPS to OCP - NOT hardcoded ratios.
    ws_pf["A22"] = "Anti-dilution adjustments (F_PPS vs OCP_class)"
    ws_pf["A23"] = "Class"
    ws_pf["B23"] = "Trigger (1=fired)"
    ws_pf["C23"] = "new_conversion_price_usd"
    ws_pf["D23"] = "new_ratio"

    # Seed (dormant)
    ws_pf["A24"] = "Seed"
    ws_pf["B24"] = "=IF(Assumptions!$B$120<Assumptions!$B$22,1,0)"
    ws_pf["C24"] = ("=IF(Assumptions!$B$120<Assumptions!$B$22,"
                    "Assumptions!$B$22*(Assumptions!$B$64+Assumptions!$B$41/Assumptions!$B$22)/"
                    "(Assumptions!$B$64+Assumptions!$B$122),"
                    "Assumptions!$B$22)")
    ws_pf["D24"] = "=Assumptions!$B$22/C24"

    # Series A (dormant)
    ws_pf["A25"] = "Series A"
    ws_pf["B25"] = "=IF(Assumptions!$B$120<Assumptions!$B$23,1,0)"
    ws_pf["C25"] = ("=IF(Assumptions!$B$120<Assumptions!$B$23,"
                    "Assumptions!$B$23*(Assumptions!$B$64+Assumptions!$B$41/Assumptions!$B$23)/"
                    "(Assumptions!$B$64+Assumptions!$B$122),"
                    "Assumptions!$B$23)")
    ws_pf["D25"] = "=Assumptions!$B$23/C25"

    # Series B (dormant)
    ws_pf["A26"] = "Series B"
    ws_pf["B26"] = "=IF(Assumptions!$B$120<Assumptions!$B$24,1,0)"
    ws_pf["C26"] = ("=IF(Assumptions!$B$120<Assumptions!$B$24,"
                    "Assumptions!$B$24*(Assumptions!$B$64+Assumptions!$B$41/Assumptions!$B$24)/"
                    "(Assumptions!$B$64+Assumptions!$B$122),"
                    "Assumptions!$B$24)")
    ws_pf["D26"] = "=Assumptions!$B$24/C26"

    # Series C (dormant)
    ws_pf["A27"] = "Series C"
    ws_pf["B27"] = "=IF(Assumptions!$B$120<Assumptions!$B$25,1,0)"
    ws_pf["C27"] = ("=IF(Assumptions!$B$120<Assumptions!$B$25,"
                    "Assumptions!$B$25*(Assumptions!$B$64+Assumptions!$B$41/Assumptions!$B$25)/"
                    "(Assumptions!$B$64+Assumptions!$B$122),"
                    "Assumptions!$B$25)")
    ws_pf["D27"] = "=Assumptions!$B$25/C27"

    # Series D (triggers in v4)
    # NCP_D: refs Assumptions row 127 (NCP_D_v4 from converged iter)
    # ratio: COMPOUND ratio = r_D(v3) * r_D(v4)  (per spec sec 12.2)
    ws_pf["A28"] = "Series D"
    ws_pf["B28"] = "=IF(Assumptions!$B$120<Assumptions!$B$26,1,0)"
    ws_pf["C28"] = ("=IF(Assumptions!$B$120<Assumptions!$B$26,"
                    "Assumptions!$B$127,Assumptions!$B$26)")
    ws_pf["D28"] = "=Assumptions!$B$36*Assumptions!$B$125"  # compound = r_D(v3) * r_D(v4)

    # Series E (triggers in v4)
    ws_pf["A29"] = "Series E"
    ws_pf["B29"] = "=IF(Assumptions!$B$120<Assumptions!$B$27,1,0)"
    ws_pf["C29"] = ("=IF(Assumptions!$B$120<Assumptions!$B$27,"
                    "Assumptions!$B$126,Assumptions!$B$27)")
    ws_pf["D29"] = "=Assumptions!$B$124"

    add_name("antidilution_seed_new_conversion_price_usd",     "Pro-Forma", "C24")
    add_name("antidilution_seed_new_ratio",                    "Pro-Forma", "D24")
    add_name("antidilution_series_a_new_conversion_price_usd", "Pro-Forma", "C25")
    add_name("antidilution_series_a_new_ratio",                "Pro-Forma", "D25")
    add_name("antidilution_series_b_new_conversion_price_usd", "Pro-Forma", "C26")
    add_name("antidilution_series_b_new_ratio",                "Pro-Forma", "D26")
    add_name("antidilution_series_c_new_conversion_price_usd", "Pro-Forma", "C27")
    add_name("antidilution_series_c_new_ratio",                "Pro-Forma", "D27")
    add_name("antidilution_series_d_new_conversion_price_usd", "Pro-Forma", "C28")
    add_name("antidilution_series_d_new_ratio",                "Pro-Forma", "D28")
    add_name("antidilution_series_e_new_conversion_price_usd", "Pro-Forma", "C29")
    add_name("antidilution_series_e_new_ratio",                "Pro-Forma", "D29")

    # ----- Notes block (2 notes x 5 fields) -----
    # Per spec: conversion-branch cells must use IF formulas comparing cap to discount.
    ws_pf["A31"] = "Convertible note conversions"
    ws_pf["A32"] = "Note"
    ws_pf["B32"] = "cap_implied_pps_usd"
    ws_pf["C32"] = "discount_implied_pps_usd"
    ws_pf["D32"] = "conversion_pps_usd"
    ws_pf["E32"] = "conversion_branch"
    ws_pf["F32"] = "shares_issued"

    ws_pf["A33"] = "Note A"
    ws_pf["B33"] = "=Assumptions!$B$130"
    ws_pf["C33"] = "=Assumptions!$B$131"
    ws_pf["D33"] = "=MIN(B33,C33)"
    ws_pf["E33"] = '=IF(B33<C33,"cap","discount")'
    # shares_issued is an integer per spec sec 13: round to whole shares
    ws_pf["F33"] = "=ROUND(Assumptions!$B$50/D33,0)"

    ws_pf["A34"] = "Note B"
    ws_pf["B34"] = "=Assumptions!$B$134"
    ws_pf["C34"] = "=Assumptions!$B$135"
    ws_pf["D34"] = "=MIN(B34,C34)"
    ws_pf["E34"] = '=IF(B34<C34,"cap","discount")'
    # shares_issued is an integer per spec sec 13: round to whole shares
    ws_pf["F34"] = "=ROUND(Assumptions!$B$53/D34,0)"

    add_name("notes_note_a_cap_implied_pps_usd",      "Pro-Forma", "B33")
    add_name("notes_note_a_discount_implied_pps_usd", "Pro-Forma", "C33")
    add_name("notes_note_a_conversion_pps_usd",       "Pro-Forma", "D33")
    add_name("notes_note_a_conversion_branch",        "Pro-Forma", "E33")
    add_name("notes_note_a_shares_issued",            "Pro-Forma", "F33")
    add_name("notes_note_b_cap_implied_pps_usd",      "Pro-Forma", "B34")
    add_name("notes_note_b_discount_implied_pps_usd", "Pro-Forma", "C34")
    add_name("notes_note_b_conversion_pps_usd",       "Pro-Forma", "D34")
    add_name("notes_note_b_conversion_branch",        "Pro-Forma", "E34")
    add_name("notes_note_b_shares_issued",            "Pro-Forma", "F34")

    # ----- Tender block (3 cells, sec 12.4) -----
    ws_pf["A36"] = "Tender mechanics"
    ws_pf["A37"] = "shares_transferred";              ws_pf["B37"] = "=Assumptions!$B$143"
    ws_pf["A38"] = "usd_to_founders";                 ws_pf["B38"] = "=Assumptions!$B$43"
    ws_pf["A39"] = "f_lead_holding_secondary_shares"; ws_pf["B39"] = "=Assumptions!$B$143"

    add_name("tender_shares_transferred",              "Pro-Forma", "B37")
    add_name("tender_usd_to_founders",                 "Pro-Forma", "B38")
    add_name("tender_f_lead_holding_secondary_shares", "Pro-Forma", "B39")

    # =======================================================================
    # 3) Waterfall tab
    # =======================================================================
    # Reverse-chronological cascade: F -> E -> D -> C -> B -> A -> Seed ->
    # notes -> common+RSAs -> options.  Per priced class per exit, an
    # election cell using IF(pref > convert, "preference", "convert").
    # Dollar cells IF-driven by election.
    #
    # Approach: per exit, use an iterative-elections block on the Waterfall
    # tab (or a helper region within Waterfall) that converges via 8
    # unrolled passes.  At each pass, we:
    #   1. Given current election bits e_F, e_E, ..., e_Seed:
    #      pref_pool_taken = sum_{c not converting} pref_c
    #      conv_shares = sum_{c converting} shares_ac_c + N_A + N_B + common
    #                    + (if itm: opt_issued)
    #      pool = exit - pref_pool_taken (+ opts strike receipts if itm)
    #      pps = pool / conv_shares
    #      pref_value_c = pref_c
    #      conv_value_c = pps * shares_ac_c
    #   2. Update election bit: e_c = IF(conv > pref, 1, 0)
    # Convergence in 6-8 passes for these inputs.
    #
    # Layout per exit: 5 columns side-by-side.  Each column hosts iter rows.

    ws_wf["A1"] = "Liquidation waterfall (election iteration converges per exit)"
    EXIT_COLS = ["B", "C", "D", "E", "F"]

    # Row 3: exit tag header
    ws_wf["A3"] = "Exit"
    for i, tag in enumerate(EXIT_TAGS):
        ws_wf.cell(row=3, column=2+i, value=tag)
    # Row 4: exit value
    ws_wf["A4"] = "Exit value $"
    for i in range(5):
        ws_wf.cell(row=4, column=2+i, value=f"=Assumptions!$B${58+i}")

    # Common references
    SHARES = {
        "F":    "Assumptions!$B$141",   # new_F = F primary
        "E":    "Assumptions!$B$128",   # E_adj
        "D":    "Assumptions!$B$129",   # D_adj
        "C":    "Assumptions!$B$14",    # C
        "B":    "Assumptions!$B$13",    # B
        "A":    "Assumptions!$B$12",    # A
        "Seed": "Assumptions!$B$11",
    }
    PREFS = {
        "F":    "Assumptions!$B$41",    # F primary raise = pref
        "E":    "Assumptions!$B$33",
        "D":    "Assumptions!$B$32",
        "C":    "Assumptions!$B$31",
        "B":    "Assumptions!$B$30",
        "A":    "Assumptions!$B$29",
        "Seed": "Assumptions!$B$28",
    }
    NOTE_A_SH = "Assumptions!$B$133"
    NOTE_B_SH = "Assumptions!$B$137"
    COMMON_SH = "Assumptions!$B$9"  # 9M (founders + RSAs); for waterfall the 9M splits across
                                     # founders_post_tender and F_LEAD_secondary, both common-equiv.
                                     # Total = 9M, so use 9M as the "common pool" for the waterfall.
    OPT_ISS = "Assumptions!$B$153"   # 2,500,000 (unchanged)
    OPT_UN = "Assumptions!$B$152"    # unissued (no exit value)
    STRIKE = "Assumptions!$B$37"     # 0.50

    PRICED_CLASSES = ["F", "E", "D", "C", "B", "A", "Seed"]

    # We layout iterations vertically.  Per exit column, rows 7..N_ITER+6
    # are iterations.  At each iter row, store election bits in 7 sub-rows
    # is not feasible per cell.  Instead we flatten per iter:
    #   For each iter k in 0..K-1:
    #     row 7 + k*ITER_HEIGHT : header for this iter
    #     rows 8..8+rowsperiter : per-class subrows containing pref_c, conv_c,
    #                             elect_c bit (1=convert), in dedicated cells
    # That layout is unwieldy. Better: use a separate helper sheet OR use a
    # wide multi-column block per exit on the Waterfall tab.
    #
    # Cleaner approach: per exit, build a vertical iteration block where each
    # iter has a fixed set of 1+7+1=9 rows.  Across 5 exits, place each block
    # in its own column-group (cols B..F as the data column for each exit).
    # We'll layout vertically with rows for: pool, pps, then per-class pref
    # & conv (so 2 rows per class), then per-class election bit (1 row).
    #
    # For simplicity and to match v3 style (which used a brute-force combo
    # tab), let's use a brute-force 2^7 = 128 combo enumeration over priced
    # classes per exit, on a hidden helper region of the Waterfall tab.
    # Then the final election cells use IF(pref>conv, ...) referenced from
    # the Nash row of that helper.
    #
    # 5 exits * 128 = 640 combo rows is manageable.

    # Layout the 128-combo helper starting at row 50 of the Waterfall tab.
    # Columns:
    #   A: exit_tag, B: exit_val, C..I: e_F, e_E, e_D, e_C, e_B, e_A, e_Seed
    #   J: pref_pool, K: conv_shares_total, L: pool_after_pref, M: opts_itm
    #   N: denom (with opts), O: pool_adj, P: pps
    #   Q..W: dollars per class F,E,D,C,B,A,Seed (7 cols)
    #   X: notes_dollars (combined N_A + N_B at pps)
    #   Y: common_dollars (= pps * 9M)
    #   Z: options_issued_dollars
    #   AA: total_distributed
    #   AB: row_idx_in_block (0..127)
    #   AC..AI: flip indices for F,E,D,C,B,A,Seed (XOR with bit)
    #   AJ..AP: alt payoff for each class under that flipped combo
    #   AQ..AW: ok flags (current >= alt - tol)
    #   AX: all_ok (Nash flag)
    #   AY..BE: pref_value_c per priced class (the "preference" payoff)
    #   BF..BL: conv_value_c per priced class (the "convert" payoff)
    HELPER_HEADER_ROW = 50
    HELPER_DATA_START = 51
    HELPER_HEADERS = [
        "exit_tag", "exit_val",
        "e_F", "e_E", "e_D", "e_C", "e_B", "e_A", "e_Seed",
        "pref_pool", "conv_shares", "pool_after_pref",
        "opts_itm", "denom", "pool_adj", "pps",
        "F_usd", "E_usd", "D_usd", "C_usd", "B_usd", "A_usd", "Seed_usd",
        "notes_usd", "common_usd", "opts_issued_usd", "total_dist",
        "row_idx",
        "fF", "fE", "fD", "fC", "fB", "fA", "fSeed",
        "altF", "altE", "altD", "altC", "altB", "altA", "altSeed",
        "okF", "okE", "okD", "okC", "okB", "okA", "okSeed",
        "all_ok",
        "prefF_v", "prefE_v", "prefD_v", "prefC_v", "prefB_v", "prefA_v", "prefSeed_v",
        "convF_v", "convE_v", "convD_v", "convC_v", "convB_v", "convA_v", "convSeed_v",
    ]
    for i, h in enumerate(HELPER_HEADERS, start=1):
        ws_wf.cell(row=HELPER_HEADER_ROW, column=i, value=h)

    # Column index helpers (1-based to letter)
    def c2l(idx):
        return get_column_letter(idx)

    # Sequential column indices
    # 1=A exit_tag, 2=B exit_val, 3=C e_F ... 9=I e_Seed
    # 10=J pref_pool ... 16=P pps
    # 17=Q F_usd ... 23=W Seed_usd
    # 24=X notes_usd, 25=Y common_usd, 26=Z opts_usd, 27=AA total
    # 28=AB row_idx, 29=AC fF ... 35=AI fSeed
    # 36=AJ altF ... 42=AP altSeed
    # 43=AQ okF ... 49=AW okSeed
    # 50=AX all_ok
    # 51=AY prefF_v ... 57=BE prefSeed_v
    # 58=BF convF_v ... 64=BL convSeed_v

    block_starts = {}  # exit_tag -> first row
    block_ends = {}

    for exit_i in range(5):
        block_start = HELPER_DATA_START + exit_i * 128
        block_end = block_start + 127
        block_starts[EXIT_TAGS[exit_i]] = block_start
        block_ends[EXIT_TAGS[exit_i]] = block_end

        for combo in range(128):
            r = block_start + combo
            # decode combo bits: MSB=F, then E, D, C, B, A, LSB=Seed
            e_F = (combo >> 6) & 1
            e_E = (combo >> 5) & 1
            e_D = (combo >> 4) & 1
            e_C = (combo >> 3) & 1
            e_B = (combo >> 2) & 1
            e_A = (combo >> 1) & 1
            e_S = (combo >> 0) & 1
            ws_wf.cell(row=r, column=1, value=EXIT_TAGS[exit_i])
            ws_wf.cell(row=r, column=2, value=f"=Assumptions!$B${58+exit_i}")
            ws_wf.cell(row=r, column=3, value=e_F)
            ws_wf.cell(row=r, column=4, value=e_E)
            ws_wf.cell(row=r, column=5, value=e_D)
            ws_wf.cell(row=r, column=6, value=e_C)
            ws_wf.cell(row=r, column=7, value=e_B)
            ws_wf.cell(row=r, column=8, value=e_A)
            ws_wf.cell(row=r, column=9, value=e_S)
            # pref_pool (col J=10)
            ws_wf.cell(row=r, column=10, value=
                f"=(1-C{r})*{PREFS['F']}+(1-D{r})*{PREFS['E']}+(1-E{r})*{PREFS['D']}"
                f"+(1-F{r})*{PREFS['C']}+(1-G{r})*{PREFS['B']}+(1-H{r})*{PREFS['A']}"
                f"+(1-I{r})*{PREFS['Seed']}")
            # conv_shares (col K=11): elect priced classes + always notes + always common
            ws_wf.cell(row=r, column=11, value=
                f"=C{r}*{SHARES['F']}+D{r}*{SHARES['E']}+E{r}*{SHARES['D']}"
                f"+F{r}*{SHARES['C']}+G{r}*{SHARES['B']}+H{r}*{SHARES['A']}"
                f"+I{r}*{SHARES['Seed']}+{NOTE_A_SH}+{NOTE_B_SH}+{COMMON_SH}")
            # pool_after_pref (L=12): exit - pref_pool
            ws_wf.cell(row=r, column=12, value=f"=B{r}-J{r}")
            # opts_itm (M=13): trial pps without options >= strike?
            #   trial_pps = pool_after_pref / conv_shares  (no opts)
            ws_wf.cell(row=r, column=13, value=
                f"=IF(K{r}>0,IF(L{r}/K{r}>={STRIKE},1,0),0)")
            # denom (N=14): conv_shares + opts_itm * opt_issued
            ws_wf.cell(row=r, column=14, value=f"=K{r}+M{r}*{OPT_ISS}")
            # pool_adj (O=15): pool_after_pref + opts_itm * opt_issued * strike
            ws_wf.cell(row=r, column=15, value=f"=L{r}+M{r}*{OPT_ISS}*{STRIKE}")
            # pps (P=16)
            ws_wf.cell(row=r, column=16, value=f"=IF(N{r}>0,O{r}/N{r},0)")
            # Class dollars (Q..W)
            ws_wf.cell(row=r, column=17, value=f"=IF(C{r}=1,P{r}*{SHARES['F']},{PREFS['F']})")
            ws_wf.cell(row=r, column=18, value=f"=IF(D{r}=1,P{r}*{SHARES['E']},{PREFS['E']})")
            ws_wf.cell(row=r, column=19, value=f"=IF(E{r}=1,P{r}*{SHARES['D']},{PREFS['D']})")
            ws_wf.cell(row=r, column=20, value=f"=IF(F{r}=1,P{r}*{SHARES['C']},{PREFS['C']})")
            ws_wf.cell(row=r, column=21, value=f"=IF(G{r}=1,P{r}*{SHARES['B']},{PREFS['B']})")
            ws_wf.cell(row=r, column=22, value=f"=IF(H{r}=1,P{r}*{SHARES['A']},{PREFS['A']})")
            ws_wf.cell(row=r, column=23, value=f"=IF(I{r}=1,P{r}*{SHARES['Seed']},{PREFS['Seed']})")
            # notes_usd (X=24)
            ws_wf.cell(row=r, column=24, value=f"=P{r}*({NOTE_A_SH}+{NOTE_B_SH})")
            # common_usd (Y=25)
            ws_wf.cell(row=r, column=25, value=f"=P{r}*{COMMON_SH}")
            # opts_issued_usd (Z=26)
            ws_wf.cell(row=r, column=26, value=f"=IF(M{r}=1,(P{r}-{STRIKE})*{OPT_ISS},0)")
            # total_dist (AA=27)
            ws_wf.cell(row=r, column=27, value=
                f"=Q{r}+R{r}+S{r}+T{r}+U{r}+V{r}+W{r}+X{r}+Y{r}+Z{r}")
            # row_idx (AB=28)
            ws_wf.cell(row=r, column=28, value=combo)
            # flip indices (AC..AI = 29..35); XOR with bit per class
            ws_wf.cell(row=r, column=29, value=combo ^ 64)   # flip F
            ws_wf.cell(row=r, column=30, value=combo ^ 32)   # flip E
            ws_wf.cell(row=r, column=31, value=combo ^ 16)   # flip D
            ws_wf.cell(row=r, column=32, value=combo ^ 8)    # flip C
            ws_wf.cell(row=r, column=33, value=combo ^ 4)    # flip B
            ws_wf.cell(row=r, column=34, value=combo ^ 2)    # flip A
            ws_wf.cell(row=r, column=35, value=combo ^ 1)    # flip Seed

            # pref_value per class (AY..BE = 51..57): the "preference" payoff
            # = always the class's pref dollar amount (constant)
            ws_wf.cell(row=r, column=51, value=f"={PREFS['F']}")
            ws_wf.cell(row=r, column=52, value=f"={PREFS['E']}")
            ws_wf.cell(row=r, column=53, value=f"={PREFS['D']}")
            ws_wf.cell(row=r, column=54, value=f"={PREFS['C']}")
            ws_wf.cell(row=r, column=55, value=f"={PREFS['B']}")
            ws_wf.cell(row=r, column=56, value=f"={PREFS['A']}")
            ws_wf.cell(row=r, column=57, value=f"={PREFS['Seed']}")
            # conv_value per class (BF..BL = 58..64): pps * shares
            ws_wf.cell(row=r, column=58, value=f"=P{r}*{SHARES['F']}")
            ws_wf.cell(row=r, column=59, value=f"=P{r}*{SHARES['E']}")
            ws_wf.cell(row=r, column=60, value=f"=P{r}*{SHARES['D']}")
            ws_wf.cell(row=r, column=61, value=f"=P{r}*{SHARES['C']}")
            ws_wf.cell(row=r, column=62, value=f"=P{r}*{SHARES['B']}")
            ws_wf.cell(row=r, column=63, value=f"=P{r}*{SHARES['A']}")
            ws_wf.cell(row=r, column=64, value=f"=P{r}*{SHARES['Seed']}")

        # Pass 2: alt payoffs and Nash flag (per combo within block)
        for combo in range(128):
            r = block_start + combo
            # alt_X = INDEX(class_usd_col_X, flip_X_idx + 1)
            # F_usd = col Q (17); E_usd=R(18); D_usd=S(19); C_usd=T(20); B_usd=U(21); A_usd=V(22); Seed_usd=W(23)
            ws_wf.cell(row=r, column=36, value=f"=INDEX($Q${block_start}:$Q${block_end},AC{r}+1)")
            ws_wf.cell(row=r, column=37, value=f"=INDEX($R${block_start}:$R${block_end},AD{r}+1)")
            ws_wf.cell(row=r, column=38, value=f"=INDEX($S${block_start}:$S${block_end},AE{r}+1)")
            ws_wf.cell(row=r, column=39, value=f"=INDEX($T${block_start}:$T${block_end},AF{r}+1)")
            ws_wf.cell(row=r, column=40, value=f"=INDEX($U${block_start}:$U${block_end},AG{r}+1)")
            ws_wf.cell(row=r, column=41, value=f"=INDEX($V${block_start}:$V${block_end},AH{r}+1)")
            ws_wf.cell(row=r, column=42, value=f"=INDEX($W${block_start}:$W${block_end},AI{r}+1)")
            TOL = "0.01"
            # ok flags AQ..AW (43..49)
            ws_wf.cell(row=r, column=43, value=f"=IF(Q{r}>=AJ{r}-{TOL},1,0)")
            ws_wf.cell(row=r, column=44, value=f"=IF(R{r}>=AK{r}-{TOL},1,0)")
            ws_wf.cell(row=r, column=45, value=f"=IF(S{r}>=AL{r}-{TOL},1,0)")
            ws_wf.cell(row=r, column=46, value=f"=IF(T{r}>=AM{r}-{TOL},1,0)")
            ws_wf.cell(row=r, column=47, value=f"=IF(U{r}>=AN{r}-{TOL},1,0)")
            ws_wf.cell(row=r, column=48, value=f"=IF(V{r}>=AO{r}-{TOL},1,0)")
            ws_wf.cell(row=r, column=49, value=f"=IF(W{r}>=AP{r}-{TOL},1,0)")
            # all_ok AX (50)
            ws_wf.cell(row=r, column=50, value=
                f"=AQ{r}*AR{r}*AS{r}*AT{r}*AU{r}*AV{r}*AW{r}")

    # -----------------------------------------------------------------
    # Final Waterfall results table (top of sheet, rows 6-30)
    # -----------------------------------------------------------------
    # Row 6: matched Nash row offset (1..128) per exit
    ws_wf["A6"] = "Nash combo offset"
    for exit_i in range(5):
        bs = block_starts[EXIT_TAGS[exit_i]]
        be = block_ends[EXIT_TAGS[exit_i]]
        col_letter = EXIT_COLS[exit_i]
        ws_wf.cell(row=6, column=2+exit_i, value=
            f"=MATCH(1,$AX${bs}:$AX${be},0)")

    # Row 7-13: election cells (F, E, D, C, B, A, Seed) using IF on pref vs conv
    # Election formula: IF(pref_at_nash > conv_at_nash, "preference", "convert")
    # pref_v_c is in cols AY..BE (51..57); conv_v_c in BF..BL (58..64)
    elect_rows = [
        ("series_f",    7,  "AY", "BF"),
        ("series_e",    8,  "AZ", "BG"),
        ("series_d",    9,  "BA", "BH"),
        ("series_c",    10, "BB", "BI"),
        ("series_b",    11, "BC", "BJ"),
        ("series_a",    12, "BD", "BK"),
        ("seed",        13, "BE", "BL"),
    ]
    for name, row_n, pref_col, conv_col in elect_rows:
        ws_wf.cell(row=row_n, column=1, value=f"{name}_election")
        for exit_i in range(5):
            bs = block_starts[EXIT_TAGS[exit_i]]
            be = block_ends[EXIT_TAGS[exit_i]]
            col_letter = EXIT_COLS[exit_i]
            # IF(INDEX(pref_col, nash_offset) > INDEX(conv_col, nash_offset), "preference", "convert")
            fml = (
                f'=IF(INDEX(${pref_col}${bs}:${pref_col}${be},{col_letter}$6)>'
                f'INDEX(${conv_col}${bs}:${conv_col}${be},{col_letter}$6),'
                f'"preference","convert")'
            )
            ws_wf.cell(row=row_n, column=2+exit_i, value=fml)

    # Row 16-25: dollar cells (F, E, D, C, B, A, Seed, notes, common, options_issued)
    # Each is INDEX into the helper at the Nash row.
    dollar_rows = [
        ("series_f", 16, "Q"),
        ("series_e", 17, "R"),
        ("series_d", 18, "S"),
        ("series_c", 19, "T"),
        ("series_b", 20, "U"),
        ("series_a", 21, "V"),
        ("seed",     22, "W"),
        ("notes",    23, "X"),
        ("common",   24, "Y"),
        ("options_issued", 25, "Z"),
    ]
    for name, row_n, dcol in dollar_rows:
        ws_wf.cell(row=row_n, column=1, value=f"{name}_dollars")
        for exit_i in range(5):
            bs = block_starts[EXIT_TAGS[exit_i]]
            be = block_ends[EXIT_TAGS[exit_i]]
            col_letter = EXIT_COLS[exit_i]
            fml = f"=INDEX(${dcol}${bs}:${dcol}${be},{col_letter}$6)"
            ws_wf.cell(row=row_n, column=2+exit_i, value=fml)

    # Row 26: our_fund_dollars = series_f_dollars * (our_fund_shares / new_F_shares)
    ws_wf.cell(row=26, column=1, value="our_fund_dollars")
    for exit_i in range(5):
        col_letter = EXIT_COLS[exit_i]
        ws_wf.cell(row=26, column=2+exit_i, value=
            f"={col_letter}16*Assumptions!$B$140/Assumptions!$B$141")

    # Row 27: total_distributed (sum of all 10 dollar lines)
    ws_wf.cell(row=27, column=1, value="total_distributed")
    for exit_i in range(5):
        col_letter = EXIT_COLS[exit_i]
        ws_wf.cell(row=27, column=2+exit_i, value=
            f"={col_letter}16+{col_letter}17+{col_letter}18+{col_letter}19+{col_letter}20"
            f"+{col_letter}21+{col_letter}22+{col_letter}23+{col_letter}24+{col_letter}25")

    # Defined names for Waterfall (95 keys: 5 x (12 dollars + 7 elections))
    for exit_i, tag in enumerate(EXIT_TAGS):
        col_letter = EXIT_COLS[exit_i]
        # Elections (7 per exit)
        add_name(f"waterfall_{tag}_series_f_election", "Waterfall", f"{col_letter}7")
        add_name(f"waterfall_{tag}_series_e_election", "Waterfall", f"{col_letter}8")
        add_name(f"waterfall_{tag}_series_d_election", "Waterfall", f"{col_letter}9")
        add_name(f"waterfall_{tag}_series_c_election", "Waterfall", f"{col_letter}10")
        add_name(f"waterfall_{tag}_series_b_election", "Waterfall", f"{col_letter}11")
        add_name(f"waterfall_{tag}_series_a_election", "Waterfall", f"{col_letter}12")
        add_name(f"waterfall_{tag}_seed_election",     "Waterfall", f"{col_letter}13")
        # Dollars (12 per exit)
        add_name(f"waterfall_{tag}_series_f_dollars", "Waterfall", f"{col_letter}16")
        add_name(f"waterfall_{tag}_series_e_dollars", "Waterfall", f"{col_letter}17")
        add_name(f"waterfall_{tag}_series_d_dollars", "Waterfall", f"{col_letter}18")
        add_name(f"waterfall_{tag}_series_c_dollars", "Waterfall", f"{col_letter}19")
        add_name(f"waterfall_{tag}_series_b_dollars", "Waterfall", f"{col_letter}20")
        add_name(f"waterfall_{tag}_series_a_dollars", "Waterfall", f"{col_letter}21")
        add_name(f"waterfall_{tag}_seed_dollars",     "Waterfall", f"{col_letter}22")
        add_name(f"waterfall_{tag}_notes_dollars",    "Waterfall", f"{col_letter}23")
        add_name(f"waterfall_{tag}_common_dollars",   "Waterfall", f"{col_letter}24")
        add_name(f"waterfall_{tag}_options_issued_dollars", "Waterfall", f"{col_letter}25")
        add_name(f"waterfall_{tag}_our_fund_dollars", "Waterfall", f"{col_letter}26")
        add_name(f"waterfall_{tag}_total_distributed", "Waterfall", f"{col_letter}27")

    # =======================================================================
    # 4) Sensitivity tab (3 x 3 grid, 9 cells, each re-solves the round)
    # =======================================================================
    # For each cell, we own its own ESOP and pre-money inputs and replicate
    # the 50-iter fixed-point block.  Per cell output: our $10M post-F % and
    # $ at $2B exit.
    #
    # Layout:
    #   Top table (rows 3-13): 9 cells, each 1 row, with output our_pct and
    #   dollars_at_2000M_exit.
    #   Per-cell inputs and iteration block: rows 30 + idx * 75 .. + 70.
    #   Per-cell waterfall at $2B: 9 cells x 128 combos, starting at much
    #   later row to avoid collision.

    ws_se["A1"] = "Sensitivity (3x3 grid: ESOP refresh x pre-money F)"
    ws_se["A3"] = "label"
    ws_se["B3"] = "esop"
    ws_se["C3"] = "pre_M"
    ws_se["D3"] = "our_pct"
    ws_se["E3"] = "dollars_at_2000M_exit"

    SENS_CELLS = []  # list of (label, esop_pct, pre_M, label_pretty)
    for esop in [0.11, 0.13, 0.15]:
        for pre_M in [1200, 1400, 1800]:
            label = f"esop_{int(esop*100)}_pre_{pre_M}M"
            SENS_CELLS.append((esop, pre_M, label))

    # Per-cell iteration blocks
    # Each cell has:
    #   Inputs row (esop, pre_money, post_money)
    #   50-iter fixed-point iteration table (50 rows + 1 header)
    #   Converged values pulled into the top table
    # Layout each cell block as 75 rows starting at row 100 + idx*75.
    SENS_BLOCK_START = 100
    SENS_BLOCK_HEIGHT = 75

    cell_cv_rows = {}    # label -> conv row number (last iter row)
    cell_input_rows = {} # label -> row number where esop/pre/post live

    for idx, (esop, pre_M, label) in enumerate(SENS_CELLS):
        block_start = SENS_BLOCK_START + idx * SENS_BLOCK_HEIGHT
        # Header
        ws_se.cell(row=block_start, column=1, value=f"Cell {label}")
        ws_se.cell(row=block_start+1, column=1, value="esop")
        ws_se.cell(row=block_start+1, column=2, value=esop)
        ws_se.cell(row=block_start+2, column=1, value="pre_M ($M)")
        ws_se.cell(row=block_start+2, column=2, value=pre_M)
        ws_se.cell(row=block_start+3, column=1, value="pre_money $")
        ws_se.cell(row=block_start+3, column=2, value=f"=B{block_start+2}*1000000")
        ws_se.cell(row=block_start+4, column=1, value="post_money $")
        ws_se.cell(row=block_start+4, column=2, value=f"=B{block_start+3}+Assumptions!$B$41")
        ws_se.cell(row=block_start+5, column=1, value="primary raise $")
        ws_se.cell(row=block_start+5, column=2, value="=Assumptions!$B$41")
        cell_input_rows[label] = block_start  # esop@+1, pre@+2, pre$@+3, post$@+4, raise$@+5

        # Iter header at block_start + 7
        hdr = block_start + 7
        headers = ["iter","r_E_p","r_D_p","N_A_p","N_B_p",
                   "E_adj","D_adj","F_PPS","T_F","new_F","pre_FD",
                   "capA","discA","cpA","NA",
                   "capB","discB","cpB","NB",
                   "NCP_E","r_E_n","NCP_D","r_D_n"]
        for i, h in enumerate(headers, start=1):
            ws_se.cell(row=hdr, column=i, value=h)

        # 50 iter rows
        N_IT = 50
        it_start = hdr + 1
        # POST/PRIMARY/ESOP for this cell
        post_ref = f"$B${block_start+4}"
        pri_ref = f"$B${block_start+5}"
        esop_ref = f"$B${block_start+1}"
        for i in range(N_IT):
            r = it_start + i
            ws_se.cell(row=r, column=1, value=i)
            if i == 0:
                ws_se.cell(row=r, column=2, value=1.0)
                ws_se.cell(row=r, column=3, value=1.0)
                ws_se.cell(row=r, column=4, value=0)
                ws_se.cell(row=r, column=5, value=0)
            else:
                ws_se.cell(row=r, column=2, value=f"=U{r-1}")
                ws_se.cell(row=r, column=3, value=f"=W{r-1}")
                ws_se.cell(row=r, column=4, value=f"=O{r-1}")
                ws_se.cell(row=r, column=5, value=f"=S{r-1}")
            # E_adj, D_adj
            ws_se.cell(row=r, column=6, value=f"=Assumptions!$B$16*B{r}")
            ws_se.cell(row=r, column=7, value=f"=Assumptions!$B$15*C{r}")
            # F_PPS:
            #   RHS = SUM_PRE_STATIC + D_adj + E_adj + N_A + N_B
            #   F_PPS_unclamped = (post*(1-esop) - primary) / RHS
            #   T_F_unclamped   = post / F_PPS_unclamped
            #   if esop * T_F_unclamped < pool_pre_F: clamp branch
            #     F_PPS = (post - primary) / (RHS + pool_pre_F)
            #   else: F_PPS = F_PPS_unclamped
            # Spec sec 5.4: ESOP top-up = max(0, esop*T_F - pool_pre_F).
            # When the target pool is below the inherited pool, top-up=0
            # and T_F is solved without the esop term. esop=11% triggers
            # clamp at all 3 pre-money levels in the sensitivity grid.
            sum_pre = ("(Assumptions!$B$9+Assumptions!$B$11+Assumptions!$B$12"
                       "+Assumptions!$B$13+Assumptions!$B$14)")
            rhs = f"({sum_pre}+G{r}+F{r}+D{r}+E{r})"
            f_pps_uncl = f"({post_ref}*(1-{esop_ref})-{pri_ref})/{rhs}"
            f_pps_cl = f"({post_ref}-{pri_ref})/({rhs}+Assumptions!$B$65)"
            t_f_uncl = f"{post_ref}/({f_pps_uncl})"
            ws_se.cell(row=r, column=8, value=
                f"=IF({esop_ref}*({t_f_uncl})<Assumptions!$B$65,"
                f"{f_pps_cl},{f_pps_uncl})")
            # T_F = post / F_PPS
            ws_se.cell(row=r, column=9, value=f"={post_ref}/H{r}")
            # new_F
            ws_se.cell(row=r, column=10, value=f"={pri_ref}/H{r}")
            # pre_FD
            ws_se.cell(row=r, column=11, value=f"=I{r}-J{r}")
            # Note A: cap_impl, disc_impl, conv_pps, N_A
            ws_se.cell(row=r, column=12, value=f"=Assumptions!$B$51/K{r}")
            ws_se.cell(row=r, column=13, value=f"=H{r}*(1-Assumptions!$B$52)")
            ws_se.cell(row=r, column=14, value=f"=MIN(L{r},M{r})")
            ws_se.cell(row=r, column=15, value=f"=Assumptions!$B$50/N{r}")
            # Note B
            ws_se.cell(row=r, column=16, value=f"=Assumptions!$B$54/K{r}")
            ws_se.cell(row=r, column=17, value=f"=H{r}*(1-Assumptions!$B$55)")
            ws_se.cell(row=r, column=18, value=f"=MIN(P{r},Q{r})")
            ws_se.cell(row=r, column=19, value=f"=Assumptions!$B$53/R{r}")
            # AD on E
            ws_se.cell(row=r, column=20, value=
                f"=Assumptions!$B$27*(Assumptions!$B$64+{pri_ref}/Assumptions!$B$27)/(Assumptions!$B$64+J{r})")
            ws_se.cell(row=r, column=21, value=f"=IF(H{r}<Assumptions!$B$27,Assumptions!$B$27/T{r},1)")
            # AD on D
            ws_se.cell(row=r, column=22, value=
                f"=Assumptions!$B$26*(Assumptions!$B$64+{pri_ref}/Assumptions!$B$26)/(Assumptions!$B$64+J{r})")
            ws_se.cell(row=r, column=23, value=f"=IF(H{r}<Assumptions!$B$26,Assumptions!$B$26/V{r},1)")

        cv = it_start + N_IT - 1
        cell_cv_rows[label] = cv

    # Per-cell waterfall at $2B exit (128 combos each, 9 cells = 1152 rows)
    # Place starting at row 5000 to avoid collisions.
    #
    # Layout note: each block gets EXACTLY 128 contiguous data rows. Earlier
    # versions placed a header at `block_start - 1`, which overwrote the
    # combo=127 row of the PREVIOUS block (since SENS_WF_START + idx*128
    # spaces blocks 128 apart). That broke 8 of 9 sensitivity cells: the
    # all-converts equilibrium (combo 127) is the Nash for these cells, and
    # without combo 127 in the lookup range, MATCH(1, AX...) returned #N/A.
    # We now write a single shared header at SENS_WF_START - 1 (= row 4999)
    # for documentation only; the formulas don't reference it.
    SENS_WF_START = 5000

    wf_hdr = ["cell","exit","eF","eE","eD","eC","eB","eA","eS",
              "pref_pool","conv_sh","pool_pref","opts_itm","denom",
              "pool_adj","pps","F_usd","E_usd","D_usd","C_usd","B_usd",
              "A_usd","S_usd","notes_usd","common_usd","opt_usd","total",
              "row_idx","fF","fE","fD","fC","fB","fA","fS",
              "aF","aE","aD","aC","aB","aA","aS",
              "okF","okE","okD","okC","okB","okA","okS","all_ok"]
    for i, h in enumerate(wf_hdr, start=1):
        ws_se.cell(row=SENS_WF_START - 1, column=i, value=h)

    for idx, (esop, pre_M, label) in enumerate(SENS_CELLS):
        cv = cell_cv_rows[label]   # converged iter row
        # Per-cell shares:
        SH_F_cell    = f"$J${cv}"          # new_F
        SH_E_cell    = f"$F${cv}"          # E_adj
        SH_D_cell    = f"$G${cv}"          # D_adj
        N_A_cell_sh  = f"$O${cv}"
        N_B_cell_sh  = f"$S${cv}"
        EXIT_2B_REF = "Assumptions!$B$60"  # exit row for 2000M (idx 2)

        block_start = SENS_WF_START + idx * 128
        block_end = block_start + 127

        for combo in range(128):
            r = block_start + combo
            e_F = (combo >> 6) & 1
            e_E = (combo >> 5) & 1
            e_D = (combo >> 4) & 1
            e_C = (combo >> 3) & 1
            e_B = (combo >> 2) & 1
            e_A = (combo >> 1) & 1
            e_S = (combo >> 0) & 1
            ws_se.cell(row=r, column=1, value=label if combo == 0 else "")
            ws_se.cell(row=r, column=2, value=f"={EXIT_2B_REF}")
            ws_se.cell(row=r, column=3, value=e_F)
            ws_se.cell(row=r, column=4, value=e_E)
            ws_se.cell(row=r, column=5, value=e_D)
            ws_se.cell(row=r, column=6, value=e_C)
            ws_se.cell(row=r, column=7, value=e_B)
            ws_se.cell(row=r, column=8, value=e_A)
            ws_se.cell(row=r, column=9, value=e_S)
            # pref_pool (col J=10)
            ws_se.cell(row=r, column=10, value=
                f"=(1-C{r})*{PREFS['F']}+(1-D{r})*{PREFS['E']}+(1-E{r})*{PREFS['D']}"
                f"+(1-F{r})*{PREFS['C']}+(1-G{r})*{PREFS['B']}+(1-H{r})*{PREFS['A']}"
                f"+(1-I{r})*{PREFS['Seed']}")
            # conv_sh (K=11): elect F + elect E + elect D + Seed/A/B/C + notes + common
            ws_se.cell(row=r, column=11, value=
                f"=C{r}*{SH_F_cell}+D{r}*{SH_E_cell}+E{r}*{SH_D_cell}"
                f"+F{r}*{SHARES['C']}+G{r}*{SHARES['B']}+H{r}*{SHARES['A']}"
                f"+I{r}*{SHARES['Seed']}+{N_A_cell_sh}+{N_B_cell_sh}+{COMMON_SH}")
            # pool_pref (L=12)
            ws_se.cell(row=r, column=12, value=f"=B{r}-J{r}")
            # opts_itm (M=13)
            ws_se.cell(row=r, column=13, value=
                f"=IF(K{r}>0,IF(L{r}/K{r}>={STRIKE},1,0),0)")
            # denom (N=14)
            ws_se.cell(row=r, column=14, value=f"=K{r}+M{r}*{OPT_ISS}")
            # pool_adj (O=15)
            ws_se.cell(row=r, column=15, value=f"=L{r}+M{r}*{OPT_ISS}*{STRIKE}")
            # pps (P=16)
            ws_se.cell(row=r, column=16, value=f"=IF(N{r}>0,O{r}/N{r},0)")
            # class dollars
            ws_se.cell(row=r, column=17, value=f"=IF(C{r}=1,P{r}*{SH_F_cell},{PREFS['F']})")
            ws_se.cell(row=r, column=18, value=f"=IF(D{r}=1,P{r}*{SH_E_cell},{PREFS['E']})")
            ws_se.cell(row=r, column=19, value=f"=IF(E{r}=1,P{r}*{SH_D_cell},{PREFS['D']})")
            ws_se.cell(row=r, column=20, value=f"=IF(F{r}=1,P{r}*{SHARES['C']},{PREFS['C']})")
            ws_se.cell(row=r, column=21, value=f"=IF(G{r}=1,P{r}*{SHARES['B']},{PREFS['B']})")
            ws_se.cell(row=r, column=22, value=f"=IF(H{r}=1,P{r}*{SHARES['A']},{PREFS['A']})")
            ws_se.cell(row=r, column=23, value=f"=IF(I{r}=1,P{r}*{SHARES['Seed']},{PREFS['Seed']})")
            # notes
            ws_se.cell(row=r, column=24, value=f"=P{r}*({N_A_cell_sh}+{N_B_cell_sh})")
            # common
            ws_se.cell(row=r, column=25, value=f"=P{r}*{COMMON_SH}")
            # opts
            ws_se.cell(row=r, column=26, value=f"=IF(M{r}=1,(P{r}-{STRIKE})*{OPT_ISS},0)")
            # total
            ws_se.cell(row=r, column=27, value=
                f"=Q{r}+R{r}+S{r}+T{r}+U{r}+V{r}+W{r}+X{r}+Y{r}+Z{r}")
            # row_idx
            ws_se.cell(row=r, column=28, value=combo)
            # flips
            ws_se.cell(row=r, column=29, value=combo ^ 64)
            ws_se.cell(row=r, column=30, value=combo ^ 32)
            ws_se.cell(row=r, column=31, value=combo ^ 16)
            ws_se.cell(row=r, column=32, value=combo ^ 8)
            ws_se.cell(row=r, column=33, value=combo ^ 4)
            ws_se.cell(row=r, column=34, value=combo ^ 2)
            ws_se.cell(row=r, column=35, value=combo ^ 1)

        # alt cols and Nash flag
        for combo in range(128):
            r = block_start + combo
            ws_se.cell(row=r, column=36, value=f"=INDEX($Q${block_start}:$Q${block_end},AC{r}+1)")
            ws_se.cell(row=r, column=37, value=f"=INDEX($R${block_start}:$R${block_end},AD{r}+1)")
            ws_se.cell(row=r, column=38, value=f"=INDEX($S${block_start}:$S${block_end},AE{r}+1)")
            ws_se.cell(row=r, column=39, value=f"=INDEX($T${block_start}:$T${block_end},AF{r}+1)")
            ws_se.cell(row=r, column=40, value=f"=INDEX($U${block_start}:$U${block_end},AG{r}+1)")
            ws_se.cell(row=r, column=41, value=f"=INDEX($V${block_start}:$V${block_end},AH{r}+1)")
            ws_se.cell(row=r, column=42, value=f"=INDEX($W${block_start}:$W${block_end},AI{r}+1)")
            TOL = "0.01"
            ws_se.cell(row=r, column=43, value=f"=IF(Q{r}>=AJ{r}-{TOL},1,0)")
            ws_se.cell(row=r, column=44, value=f"=IF(R{r}>=AK{r}-{TOL},1,0)")
            ws_se.cell(row=r, column=45, value=f"=IF(S{r}>=AL{r}-{TOL},1,0)")
            ws_se.cell(row=r, column=46, value=f"=IF(T{r}>=AM{r}-{TOL},1,0)")
            ws_se.cell(row=r, column=47, value=f"=IF(U{r}>=AN{r}-{TOL},1,0)")
            ws_se.cell(row=r, column=48, value=f"=IF(V{r}>=AO{r}-{TOL},1,0)")
            ws_se.cell(row=r, column=49, value=f"=IF(W{r}>=AP{r}-{TOL},1,0)")
            ws_se.cell(row=r, column=50, value=
                f"=AQ{r}*AR{r}*AS{r}*AT{r}*AU{r}*AV{r}*AW{r}")

    # Top table populating
    for idx, (esop, pre_M, label) in enumerate(SENS_CELLS):
        se_row = 4 + idx
        cv = cell_cv_rows[label]
        wf_bs = SENS_WF_START + idx * 128
        wf_be = wf_bs + 127
        ws_se.cell(row=se_row, column=1, value=label)
        ws_se.cell(row=se_row, column=2, value=esop)
        ws_se.cell(row=se_row, column=3, value=pre_M)
        # our_pct = our_fund_shares / T_F = (our_check / F_PPS) / T_F
        # our_check = $10M, F_PPS = H_cv, T_F = I_cv
        ws_se.cell(row=se_row, column=4, value=
            f"=(Assumptions!$B$47/H{cv})/I{cv}")
        # dollars_at_2000M_exit = F_usd_at_nash * (our_check / F_RAISE)
        # F_usd is col Q, Nash row is MATCH(1, AX block, 0) within wf block
        ws_se.cell(row=se_row, column=5, value=
            f"=INDEX($Q${wf_bs}:$Q${wf_be},MATCH(1,$AX${wf_bs}:$AX${wf_be},0))"
            f"*Assumptions!$B$47/Assumptions!$B$41")
        add_name(f"sensitivity_{label}_our_pct",                  "Sensitivity", f"D{se_row}")
        add_name(f"sensitivity_{label}_dollars_at_2000M_exit",    "Sensitivity", f"E{se_row}")

    # =======================================================================
    # 5) Return-Solver tab
    # =======================================================================
    # required_pct_3x = (3 * our_check) * new_F / (series_f_dollars * T_F)
    # required_pct_4x = (4 * our_check) * new_F / (series_f_dollars * T_F)
    # Logic: if our fractional stake in Series F primary = k (vs base 1/3),
    # our $ at exit = k * series_f_dollars; setting k * series_f_dollars =
    # target gives k = target / series_f_dollars; required_pct = k * new_F / T_F
    # Substituting: required_pct = (target * new_F) / (series_f_dollars * T_F).

    ws_rs["A1"] = "Return solver"
    ws_rs["A3"] = "exit"
    ws_rs["B3"] = "exit_val"
    ws_rs["C3"] = "solver_3x_required_pct"
    ws_rs["D3"] = "solver_4x_required_pct"

    for i, tag in enumerate(EXIT_TAGS):
        r = 4 + i
        col_letter = EXIT_COLS[i]
        ws_rs.cell(row=r, column=1, value=tag)
        ws_rs.cell(row=r, column=2, value=f"=Assumptions!$B${58+i}")
        ws_rs.cell(row=r, column=3, value=
            f"=(3*Assumptions!$B$47)*Assumptions!$B$141/(Waterfall!${col_letter}$16*Assumptions!$B$121)")
        ws_rs.cell(row=r, column=4, value=
            f"=(4*Assumptions!$B$47)*Assumptions!$B$141/(Waterfall!${col_letter}$16*Assumptions!$B$121)")
        add_name(f"solver_3x_exit_{tag}_required_pct", "Return-Solver", f"C{r}")
        add_name(f"solver_4x_exit_{tag}_required_pct", "Return-Solver", f"D{r}")

    # =======================================================================
    # Register all defined names (workbook-level)
    # =======================================================================
    for name, ref in defined_names:
        dn = DefinedName(name=name, attr_text=ref)
        wb.defined_names[name] = dn

    out = TRUTH_DIR / "ground_truth.xlsx"
    wb.save(out)
    print(f"wrote {out}")
    print(f"defined names registered: {len(defined_names)}")
    return out, defined_names


# ---------------------------------------------------------------------------
# Verification: re-open workbook and check every expected defined name exists
# ---------------------------------------------------------------------------

def expected_defined_names() -> list[str]:
    """Per CANONICAL_INPUTS.md sec 12 - the 167 truth keys, dot->underscore."""
    names: list[str] = []

    # 12.1 Ownership (15)
    names.extend([
        "ownership_founders_pct",
        "ownership_employees_pct",
        "ownership_options_issued_pct",
        "ownership_options_unissued_pct",
        "ownership_seed_pct",
        "ownership_series_a_pct",
        "ownership_series_b_pct",
        "ownership_series_c_pct",
        "ownership_series_d_pct",
        "ownership_series_e_pct",
        "ownership_note_a_pct",
        "ownership_note_b_pct",
        "ownership_series_f_new_pct",
        "ownership_f_lead_secondary_pct",
        "ownership_our_fund_pct",
    ])

    # 12.2 Antidilution (12) - 6 classes x (NCP, ratio)
    for cls in ["seed", "series_a", "series_b", "series_c", "series_d", "series_e"]:
        names.append(f"antidilution_{cls}_new_conversion_price_usd")
        names.append(f"antidilution_{cls}_new_ratio")

    # 12.3 Notes (10) - 2 notes x 5 fields
    for note in ["note_a", "note_b"]:
        names.extend([
            f"notes_{note}_cap_implied_pps_usd",
            f"notes_{note}_discount_implied_pps_usd",
            f"notes_{note}_conversion_pps_usd",
            f"notes_{note}_conversion_branch",
            f"notes_{note}_shares_issued",
        ])

    # 12.4 Tender (3)
    names.extend([
        "tender_shares_transferred",
        "tender_usd_to_founders",
        "tender_f_lead_holding_secondary_shares",
    ])

    # 12.5 Round (4)
    names.extend([
        "round_f_pps_usd",
        "round_post_f_fd_shares",
        "round_our_fund_shares",
        "round_f_lead_primary_shares",
    ])

    # 12.6 Waterfall (95) - 5 exits x (12 dollars + 7 elections)
    for tag in EXIT_TAGS:
        # dollars (12 per exit)
        for cls in ["series_f", "series_e", "series_d", "series_c", "series_b",
                    "series_a", "seed", "notes", "common", "options_issued",
                    "our_fund", "total_distributed"]:
            if cls == "total_distributed":
                names.append(f"waterfall_{tag}_total_distributed")
            else:
                names.append(f"waterfall_{tag}_{cls}_dollars")
        # elections (7 per exit)
        for cls in ["series_f", "series_e", "series_d", "series_c", "series_b",
                    "series_a", "seed"]:
            names.append(f"waterfall_{tag}_{cls}_election")

    # 12.7 Sensitivity (18) - 9 cells x 2
    for esop in [11, 13, 15]:
        for pre_M in [1200, 1400, 1800]:
            names.append(f"sensitivity_esop_{esop}_pre_{pre_M}M_our_pct")
            names.append(f"sensitivity_esop_{esop}_pre_{pre_M}M_dollars_at_2000M_exit")

    # 12.8 Return solver (10) - 2 multiples x 5 exits
    for mult in ["3x", "4x"]:
        for tag in EXIT_TAGS:
            names.append(f"solver_{mult}_exit_{tag}_required_pct")

    return names


def verify_workbook(path: Path):
    expected = expected_defined_names()
    print(f"\nExpected defined-name count (per spec): {len(expected)}")

    wb = openpyxl.load_workbook(path, data_only=False)
    actual = set(wb.defined_names)
    print(f"Actual defined-name count in workbook: {len(actual)}")

    missing = [n for n in expected if n not in actual]
    extras = [n for n in actual if n not in expected]
    if missing:
        print(f"\nMISSING defined names ({len(missing)}):")
        for n in missing:
            print(f"  - {n}")
    else:
        print("\nAll 167 expected defined names present.")
    if extras:
        print(f"\nExtra (helper) defined names not in spec ({len(extras)}):")
        for n in extras:
            print(f"  + {n}")


if __name__ == "__main__":
    path, _names = build_ground_truth()
    verify_workbook(path)
    file_size = path.stat().st_size
    print(f"\nFile size: {file_size:,} bytes")
