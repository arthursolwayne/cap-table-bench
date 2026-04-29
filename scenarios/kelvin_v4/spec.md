# Kelvin Dynamics — Canonical Scenario v4 ("Series F: Tender + Convertibles + Down-Round AD")

Scenario specification for `kelvin_v4`. Spec-writer's render of
`CANONICAL_INPUTS.md` — every numerical claim cross-references that
file by section. Four downstream agents consume v4 (Python builder,
Excel builder, harness/Tetra-prompt agent, spec reader); all must
agree on every number, name, and key. spec.md is the human-readable
contract; CANONICAL_INPUTS.md is the machine-locked input layer.

Disk identifier `kelvin_v4`; company **Kelvin Dynamics**, the same
fictional climate-hardware startup as v3, one round further along
(Series E closed; Series F proposed).

---

## 0. Design intent

This scenario is the **second observation in the templated-complex
regime**, paired with kelvin_v3_t1 to validate Tetra's +12.5pp lead
over Opus on dense templated cap-table modeling. Complexity is turned
**up** along three orthogonal axes — each exercising a primitive that
the v3 scenario stripped.

**The three modeling discriminators.**

1. **Secondary tender alongside primary.** Series F raises $30M
   primary AND $15M secondary tender at 0.85 × F_PPS. The tender
   transfers existing shares (founders → F_LEAD) without issuing new
   shares — does NOT change T_F, but DOES shift
   `ownership.founders.pct` and adds an `ownership.f_lead_secondary.pct`
   line. A model that conflates primary and secondary will mis-allocate
   founder dilution and double-count F_LEAD. Per CANONICAL_INPUTS §5.1,
   §7.4.

2. **Dual convertible-note conversion with cap-vs-discount divergence.**
   Note A's $1.0B cap is below F's $1.4B pre — Note A converts at cap.
   Note B's $1.6B cap is above F's $1.4B pre — Note B converts at
   discount. **Opposite conclusions at the same F_PPS.** A solver
   hardcoding either branch silently mis-prices one note. Per
   CANONICAL_INPUTS §6.

3. **Two simultaneous AD triggers.** F is a slight down round vs E
   (F_PPS ≈ $37.5 < E_PPS ≈ $49.16). AD triggers on **both** Series E
   (first adjustment) and Series D (SECOND adjustment, layered atop v3's
   adjusted ratio). v3 had one trigger; v4 has two, and D's adjustment
   is **compound** across rounds. `series_d.new_ratio` reports
   r_D(v3) × r_D(v4). Per CANONICAL_INPUTS §7.2.

**Discriminator surface.** v4 keeps every v3 primitive (per-class
Nash election, broad-based WA AD, 1× non-participating preferred, ESOP
refresh, 5-exit waterfall, 3×3 sensitivity, return solver) and layers
three more (tender, convertibles, compound AD). Result: **167 tracked
keys** (vs v3's 124), **35 elections** (vs 25), and a fixed-point
solver in **five unknowns** instead of three.

**Stripped primitives.** No SAFEs, no venture debt / PIK, no warrants,
no MFN, no participating preferred, no note interest accrual. Cap
table is otherwise clean growth-stage extending v3's: 9M common, six
priced rounds (Seed–E) totalling $510M raised pre-F, two outstanding
notes ($8M principal), proposed $30M+$15M Series F at $1.4B pre.

**Scorer dimensions** (per CANONICAL_INPUTS §12.10):

- `cells_correct_pct` — fraction of 167 keys within tolerance.
- `elections_correct_pct` — fraction of 35 election strings matching.
- `note_branches_correct_pct` — fraction of 2 note `conversion_branch`
  strings matching (isolates cap-vs-discount primitive).
- `audit_probe_pass` — boolean; perturbed-input recomputation
  reconciles against truth_perturbed.json.

Tetra-prompt agent should reference "167 tracked keys, 35 elections,
2 note branches" rather than rounded numbers.

---

## 1. Company timeline

| Date | Event |
|---|---|
| 2019-Q4 | Incorporation. 3 founders + 10 early employees issued Common / RSAs. |
| 2020-Q2 | Seed priced round ($5M) |
| 2021-Q1 | Series A priced round ($15M) |
| 2022-Q2 | Series B priced round ($30M) |
| 2023-Q4 | Series C priced round ($60M) |
| 2025-Q2 | Series D priced round ($100M) |
| 2026-Q2 | Series E priced round ($300M, **closed**) |
| 2026-Q3 | Convertible Note A issued ($5M principal) |
| 2026-Q4 | Convertible Note B issued ($3M principal) |
| 2027-Q2 | **Proposed Series F — slight down round, with secondary tender** |

Fund's check: **$10,000,000 in Series F primary** (= 1/3 of $30M
primary; no tender participation). Target: **3–4× net return**.

v4's pre-F state IS v3's post-E state (per CANONICAL_INPUTS §1). Every
share count, PPS, OCP, and dollar amount on rounds Seed–E is locked to
v3's fixed point and MUST NOT be re-derived. If a v4 builder
disagrees with v3, the v4 builder is wrong.

---

## 2. Stakeholders

Extends v3's `F{n}` / `E{1..10}` naming. New v4-only holders carry the
prefix of the round/instrument that created them. Per CANONICAL_INPUTS §2.

| ID | Role | Source |
|---|---|---|
| F1 | Founder, CEO            | v3 |
| F2 | Founder, CTO            | v3 |
| F3 | Founder, Chief Scientist | v3 |
| E1–E10 | Early-employee RSAs (10 people, 100k each) | v3 |
| SEED1 | Seed preferred holder (single LP) | v3 |
| A1    | Series A preferred holder | v3 |
| B1    | Series B preferred holder | v3 |
| C1    | Series C preferred holder | v3 |
| D1    | Series D preferred holder | v3 |
| E1_INV–E10_INV | Series E preferred holders (10 LPs, equal pro rata of 6.10M E shares) | v3 |
| **NA1** | **Convertible Note A holder** (new in v4) | §4 |
| **NB1** | **Convertible Note B holder** (new in v4) | §4 |
| **F_LEAD** | **Series F lead investor** (new in v4) | §6 |
| **F_FUND** | **Our fund** ($10M of $30M F primary, new in v4) | §6 |

`E1_INV` vs `E1`: `E1` is an early-employee RSA holder (carries from
v3); `E1_INV` is a Series E preferred LP. Disambiguated with `_INV`.

### 2.1 Pre-F (post-E) share counts — locked to v3's fixed point

Starting state for all v4 math. Builders MUST use these verbatim per
CANONICAL_INPUTS §1.1; do NOT re-solve the v3 fixed point.

| Holder / Class | Pre-F shares |
|---|---|
| F1 (CEO)                          | 4,000,000 |
| F2 (CTO)                          | 2,500,000 |
| F3 (Chief Scientist)              | 1,500,000 |
| Founders subtotal                 | 8,000,000 |
| E1–E10 (RSAs, 100k each)          | 1,000,000 |
| Common subtotal                   | 9,000,000 |
| Seed preferred (as-conv, ratio 1) | 5,000,000 |
| Series A (as-conv, ratio 1)       | 5,000,000 |
| Series B (as-conv, ratio 1)       | 3,000,000 |
| Series C (as-conv, ratio 1)       | 2,400,000 |
| Series D **adjusted** (ratio 1.03237623…) | 1,720,627 |
| Series E (new, base case)         | 6,102,667 |
| Options issued (post-E)           | 2,500,000 |
| Options unissued (post-E refresh) | 1,894,667 |
| **Pre-F FD total `T_E`**          | **36,617,961** |

Builders MAY round D/E share counts to integer for display, but MUST
use full Decimal precision in denominators. `T_E = 1,800,000,000 /
49.158311… ≈ 36,617,960.84`; use the literal Decimal value, not the
rounded integer, in the v4 fixed point.

### 2.2 Aggregation convention (per CANONICAL_INPUTS §2)

- Founders F1+F2+F3 → one `ownership.founders.pct` (post-tender).
- Employees E1–E10 → one `ownership.employees.pct`.
- Series E investors E1_INV–E10_INV → one `ownership.series_e.pct`
  (Series E is now an EXISTING class for v4, not the new round).
- F lead + our fund + any other new F primaries → one
  `ownership.series_f_new.pct`.
- Our fund alone → `ownership.our_fund.pct` (= 1/3 of `series_f_new`).
- F_LEAD's tender holding (common bought from founders) →
  `ownership.f_lead_secondary.pct`.
- Note A → `ownership.note_a.pct`. Note B → `ownership.note_b.pct`.

---

## 3. Priced equity rounds pre-F (Seed–E, locked)

All six rounds prior to F are locked from v3 §3 + §7. No re-derivation.

| Series | Date    | OCP ($) | Shares    | Invested ($) | LP type            | AD              |
|---|---|---|---|---|---|---|
| Seed     | 2020-Q2 | 1.00         | 5,000,000 | 5,000,000    | 1× non-part, conv  | broad-based WA |
| Series A | 2021-Q1 | 3.00         | 5,000,000 | 15,000,000   | 1× non-part, conv  | broad-based WA |
| Series B | 2022-Q2 | 10.00        | 3,000,000 | 30,000,000   | 1× non-part, conv  | broad-based WA |
| Series C | 2023-Q4 | 25.00        | 2,400,000 | 60,000,000   | 1× non-part, conv  | broad-based WA |
| Series D | 2025-Q2 | **58.118347** | 1,666,667 | 100,000,000  | 1× non-part, conv  | broad-based WA |
| Series E | 2026-Q2 | 49.158311…   | 6,102,667 | 300,000,000  | 1× non-part, conv  | broad-based WA |

Seniority: reverse-chronological **E > D > C > B > A > Seed**, with F
senior to all (see §6.4).

**Series D OCP field (v4-specific subtlety).** Series D's OCP for v4
AD is the **post-E NCP_D = $58.118347**, NOT $60.00 (per
CANONICAL_INPUTS §3). The trigger test reads "F PPS < $58.118347?",
not "< $60". Series D's share count is v3-adjusted = 1,666,667 ×
1.03237623… ≈ 1,720,627; v4's ratio compounds on top.

**Series E OCP for v4 AD** = `E_PPS = $49.158311` (per CANONICAL_INPUTS
§4). v4 is a down round → Series E AD TRIGGERS (its first adjustment).

### 3.1 v3 reference values (sanity asserts only)

Per CANONICAL_INPUTS §1.2: E_PPS ≈ $49.158311; NCP_D(post-E) ≈
$58.118347; r_D(post-E) ≈ 1.03237623; T_E ≈ 36,617,961; post-E option
pool = 12.0% of T_E = 4,394,667 (2,500,000 issued + 1,894,667
unissued). NOT v4 truth keys; sanity asserts only.

---

## 4. Convertible notes outstanding

Two notes outstanding pre-F. Both convert in full at F closing per
CANONICAL_INPUTS §6. Input constants: principal, valuation cap, and
discount are locked. Standard `min(cap-implied PPS, F_PPS × (1 −
discount))` mechanic.

### 4.1 Note terms

| Field | Note A (NA1) | Note B (NB1) |
|---|---|---|
| Principal              | $5,000,000          | $3,000,000 |
| Valuation cap          | $1,000,000,000      | $1,600,000,000 |
| Discount               | 20%                 | 15% |
| Cap-implied PPS        | cap / pre-F-FD-at-conversion | cap / pre-F-FD-at-conversion |
| Discount-implied PPS   | F_PPS × (1 − 0.20)  | F_PPS × (1 − 0.15) |
| Conversion price       | min(cap-implied, discount-implied) | min(cap-implied, discount-implied) |
| Expected outcome       | **CAP CONVERTS**    | **DISCOUNT CONVERTS** |
| Expected PPS           | ≈ $27.32            | ≈ $31.875 (= F_PPS × 0.85) |
| Expected shares        | ≈ 183,000           | ≈ 92,000–94,000 |
| Interest accrued       | $0 (zero-coupon)    | $0 (zero-coupon) |

### 4.2 Why these cap/discount values

Per CANONICAL_INPUTS §6.1 — chosen so the **two notes resolve to
opposite conclusions** at the expected F_PPS. The modeling
discriminator: solver must independently compute both `cap-implied`
and `discount-implied` per note and pick the lower.

- Note A: $1.0B cap below $1.4B pre → cap-implied dominates 20%
  discount → CAP CONVERTS.
- Note B: $1.6B cap above $1.4B pre → cap non-binding → DISCOUNT
  CONVERTS.

Hardcoding either branch fails one note.

### 4.3 Cap-implied PPS denominator

Per CANONICAL_INPUTS §6.2: `cap-implied PPS = cap /
pre-money-FD-at-conversion`. Denominator EXCLUDES new F primary shares
but INCLUDES all pre-F holders, ESOP top-up, AND the OTHER note's
converted shares (standard "pre-money-shares" definition: includes
round's other convertibles, excludes new priced equity).

Base case (approximate): cap-PPS_A ≈ $27.32; cap-PPS_B ≈ $42.50;
discount-PPS_A ≈ $30.00; discount-PPS_B ≈ $31.875. → Note A CAP ≈
$27.32 → ~183K shares; Note B DISCOUNT ≈ $31.875 → ~94K shares.

### 4.4 FD-count placement

Per CANONICAL_INPUTS §6.3 — converted shares enter `T_F`, ride pro
rata with common (convert side). NOT new preferred classes; no AD
keys; no preference. Appear in §11.1 as ownership lines.

### 4.5 Fixed-point coupling

Per CANONICAL_INPUTS §6.4 — note conversions, ESOP refresh, F_PPS, AD
on E and D, and T_F are **coupled**. Five unknowns: `F_PPS`, `T_F`,
`r_E`, `r_D(v4)`, `N_A_shares`, `N_B_shares`. System contracts <<
0.05/pass at v3-style inputs; 30–50 iterations at Decimal precision
60 → 1e-30. Closed-form by case analysis OK, but builders MUST verify
case assumption at converged F_PPS. See §13.

### 4.6 No interest, no MFN

Per CANONICAL_INPUTS §6.5 — zero-coupon, no MFN, no senior preference
(common-equivalent post-conversion), full conversion at closing.

---

## 5. SAFEs / venture debt / warrants

**All dropped, like v3.** No SAFEs, no venture debt, no warrants /
VDW / lender sweeteners. All `safe.*`, `venture_debt.*`, `warrants.*`
keys absent from v4 tracked set.

---

## 6. Series F base-case terms (per CANONICAL_INPUTS §5)

| Field | Value |
|---|---|
| Pre-money                  | **$1,400,000,000** (slight down round vs E's $1.8B post) |
| Primary raise              | **$30,000,000** |
| Secondary tender raise     | **$15,000,000** |
| Tender PPS multiplier      | **0.85× of F PPS** |
| Total cash in              | $45,000,000 ($30M dilutive primary) |
| Post-money (priced equity) | $1,430,000,000 (= pre + primary; tender excluded) |
| ESOP refresh target post-F | **13.0% of post-F FD** (pre-money refresh) |
| LP                         | 1× non-participating, conversion election |
| AD                         | broad-based WA |
| Seniority                  | senior to E, D, C, B, A, Seed; senior to all notes |
| Lead investor              | F_LEAD writes $20M of $30M primary |
| **Our fund's check**       | **$10,000,000 = 1/3 of $30M F primary** |

### 6.1 Primary vs secondary mechanics (per CANONICAL_INPUTS §5.1)

- **Primary ($30M)**: new Series F shares issued; adds to FD.
  Allocated $20M F_LEAD + $10M F_FUND.
- **Secondary tender ($15M @ 0.85·F_PPS)**: F_LEAD buys **existing
  shares from founders** (F1+F2+F3 sell pro rata). **No new shares
  issued.**
  - Tender shares = $15,000,000 / (0.85 × F_PPS).
  - Founders reduced pro rata by pre-F holdings (4M/2.5M/1.5M →
    0.50/0.3125/0.1875).
  - F_LEAD post-F holding = primary ($20M/F_PPS) + tender
    ($15M/(0.85·F_PPS)).
- **Our fund**: $10M/F_PPS Series F primary only. No tender.

### 6.2 ESOP refresh (per CANONICAL_INPUTS §5.2, §7.3)

Same as v3: pre-money top-up, target 13.0% of post-F FD. Dilutes pre-F
holders (founders, RSAs, all classes Seed–E, both note holders); does
NOT dilute F primary. Top-up = `max(0, 0.13 × T_F − 4,394,667)`
(pre-F pool = 2,500,000 issued + 1,894,667 unissued = 4,394,667).
Post-F unissued = `0.13 × T_F − 2,500,000`. If 13% × T_F < 4,394,667
(extreme low-pre cells), clamp top-up to zero ("no top-up").

### 6.3 Tender = common-purchase (locked default, per CANONICAL_INPUTS §7.4)

- Tender shares ≈ $15M / $31.875 ≈ 470,588.
- Founders → F_LEAD pro rata (F1 50%, F2 31.25%, F3 18.75%).
- T_F UNCHANGED (no issuance).
- `ownership.founders.pct` and `ownership.f_lead_secondary.pct`
  reflect post-tender.
- **Locked convention**: tender purchases **common** from founders at
  $0.85·F_PPS; F_LEAD holds a mix of (F preferred from primary) +
  (common from secondary). Standard structured-secondary mechanic.
- **Waterfall consequence**: F_LEAD's tender shares ride pro rata with
  common, NOT with F preferred. Do not contribute to F preference
  dollars; only to convert-side pool.

### 6.4 Seniority (per CANONICAL_INPUTS §5.3)

Reverse-chronological:

    F  >  E  >  D  >  C  >  B  >  A  >  Seed  >  common+RSAs+options+notes

Notes A and B convert to F-priced common-equivalent at closing
(§4.4); ride pro rata with common.

---

## 7. Anti-dilution (per CANONICAL_INPUTS §7.2)

Broad-based weighted-average on every prior class. Trigger test:

| Class | Trigger test | Status at base case |
|---|---|---|
| Seed     | F_PPS < $1.00? | dormant (test fails at any F_PPS > $1) |
| Series A | F_PPS < $3.00? | dormant |
| Series B | F_PPS < $10.00? | dormant |
| Series C | F_PPS < $25.00? | dormant |
| Series D | F_PPS < **$58.118347** (post-E NCP_D)? | **TRIGGERS** at F_PPS ≈ $37.5 |
| Series E | F_PPS < **$49.158311** (E_PPS)? | **TRIGGERS** at F_PPS ≈ $37.5 |

**Two AD triggers in v4** (vs one in v3). Each is its own broad-based
WA solve, layered into the fixed point.

### 7.1 Broad-based WA formula (NVCA form)

For triggering class X:

    NCP_X(v4) = OCP_X × (A + B_X) / (A + C)
    r_X(v4)   = OCP_X / NCP_X(v4)

- `A` = pre-F broad base = 36,617,961 (= T_E per §2.1).
- `B_X` = $30,000,000 / OCP_X (shares that would have issued at OCP_X
  for the F raise; class-dependent).
- `C` = $30,000,000 / F_PPS (actual F primary shares).

Note conversions EXCLUDED from `B` and `C` per standard NVCA
convention (notes are converting prior instruments, not new equity at
a lower price). Tender also EXCLUDED from `A`, `B`, `C` (no new
shares).

### 7.2 Compound ratio on Series D (load-bearing v4-vs-v3 distinction)

Per CANONICAL_INPUTS §12.2, §7.2:

- D's OCP for v4 AD test = post-E NCP_D = $58.118347 (NOT $60.00).
- r_D(v4) = WA formula at OCP_X = $58.118347.
- D's effective as-converted post-F = `1,666,667 × r_D(v3) × r_D(v4)
  = 1,720,627 × r_D(v4)`.
- Truth `antidilution.series_d.new_ratio` = **COMPOUND** r_D(v3) ×
  r_D(v4). `antidilution.series_d.new_conversion_price_usd` =
  v4-adjusted NCP_D (computed from OCP $58.118347).

### 7.3 Series E first adjustment

r_E(v4) = WA formula at OCP_X = $49.158311 (= E_PPS). E has not been
adjusted before; truth `series_e.new_ratio` = r_E(v4) directly (no
compounding).

### 7.4 Dormant classes (Seed / A / B / C)

Trigger fails → emit OCP unchanged, ratio = 1.0. Trigger test MUST be
performed (Excel formula, not hardcode); dormant branches must be
derived, not asserted.

### 7.5 Expected base-case AD outputs (per CANONICAL_INPUTS §16)

- AD on E: NCP_E ≈ $47.7; r_E ≈ 1.03.
- AD on D: NCP_D(v4) ≈ $55; r_D(v4) ≈ 1.05; compound ≈ 1.087.

Builders match full-precision Decimal outputs.

---

## 8. Exit scenarios and waterfall (per CANONICAL_INPUTS §8)

Five exits: **$700M, $1,200M, $2,000M, $4,000M, $8,000M**. Labels
`700M`, `1200M`, `2000M`, `4000M`, `8000M`. Bottom two put F (and
probably E, D) into preference; top two push everyone into convert;
$2B is sensitivity reference.

### 8.1 Waterfall seniority

1. Series F (1× non-part, conv election; new in v4)
2. Series E (1× non-part, conv election)
3. Series D (1× non-part, conv election; compound r_D(v3) × r_D(v4) ratio)
4. Series C (1× non-part, conv election)
5. Series B (1× non-part, conv election)
6. Series A (1× non-part, conv election)
7. Seed (1× non-part, conv election)
8. Common + RSAs + options + tender shares + converted notes (pro rata
   as-if-exercised, after all elections, after ITM options pay strike)

### 8.2 Treatment rules

- **Issued options (2,500,000 @ $0.50 strike)**: exercise if per-share
  exit ≥ $0.50 (net of strike); else forfeit.
- **Unissued options**: forfeit on acquisition.
- **Founders + RSAs**: pro rata with convert side (founders post-tender).
- **F_LEAD tender shares**: pro rata with common (§6.3).
- **Notes A, B (converted)**: pro rata with convert side. No election;
  no preference. Shares = `principal / conversion_pps` per §4.
- Each class electing **preference**: shares removed from pro-rata
  denominator, preference subtracted from pool. Each class electing
  **convert**: AD-adjusted shares included in denominator.
- `total_distributed` = exit ± $1.

### 8.3 Per-class preferences (aggregate $)

| Class | Aggregate preference ($) |
|---|---|
| Series F | $30,000,000 (primary only; tender does not create preference) |
| Series E | $300,000,000 |
| Series D | $100,000,000 |
| Series C | $60,000,000 |
| Series B | $30,000,000 |
| Series A | $15,000,000 |
| Seed     | $5,000,000 |
| **Total** | **$540,000,000** |

$700M sits above $540M sum-of-prefs but well below full-convert →
mixed-election waterfall.

### 8.4 Series F election tracked (per CANONICAL_INPUTS §7.1)

**Series F's election IS tracked in v4** (v3 did not track E's).
Rationale: F participates as a regular priced class post-close;
tracking keeps v4 symmetric with the other six classes and adds 5
more observations on the senior-most class. **35 tracked elections**
= 7 classes × 5 exits (vs v3's 25 = 5 × 5).

Notes A, B have no elections (common-equivalent post-conversion); ride
pro-rata pool.

### 8.5 Nash equilibrium

Same convention as v3. For each exit, the 7 elections must form a
simultaneous best-response. Equilibrium is unique under
reverse-chronological seniority and computable by seniority walk:
start all-convert, walk from F down, flip classes whose aggregate
preference exceeds pro-rata share of remaining pool, re-compute
juniors, iterate until stable (≤ 8 passes for 7 classes).

Builders may use seniority walk OR exhaustive max() over 2^7 = 128
combinations per exit. Exhaustive recommended for Excel clarity.

### 8.6 Our fund's proceeds

`waterfall.{exit}M.our_fund.dollars` = (1/3) × Series F primary class
proceeds. Fund holds F primary only; no tender; no other class.
F_LEAD's tender shares are NOT part of F preferred for waterfall (§6.3
common-purchase convention); fund's 1/3 is of F primary only.

---

## 9. Sensitivity grid (per CANONICAL_INPUTS §9)

ESOP refresh × Series F pre-money. **3 × 3 = 9 cells.** Outputs: our
$10M check's post-F ownership %, $ proceeds at $2B exit.

| ESOP refresh \ Pre-money F | $1,200M | $1,400M (base) | $1,800M |
|---|---|---|---|
| 11% | cell | cell | cell |
| 13% (base) | cell | **base case** | cell |
| 15% | cell | cell | cell |

Each cell re-solves the full round fixed point. F_PPS varies → AD on
D and E may retrigger or go dormant per cell → notes' cap/discount
branches must be re-evaluated per cell.

### 9.1 Cap-vs-discount may flip across cells (load-bearing)

Per CANONICAL_INPUTS §9: at $1.8B pre, **Note A may flip from CAP to
DISCOUNT**. The sensitivity grid genuinely exercises the
cap-vs-discount logic at multiple operating points.

Mechanism: at $1.8B pre, F_PPS rises (closer to E_PPS). Note A's
cap-PPS ≈ $27 (still binding); discount-PPS = F_PPS × 0.80 rises. If
0.80 × F_PPS < $27, A flips to discount. At $1.2B pre, F_PPS falls;
notes push deeper into cap. The 9 cells observe both branches.
Builders verify per cell — do not short-circuit.

### 9.2 AD per-cell retrigger / dormant

At $1.2B pre, both AD triggers fire harder; compound D ratio grows.
At $1.8B pre, F_PPS rises toward E_PPS; AD on E may go dormant if
F_PPS ≥ E_PPS. Builders verify per cell.

### 9.3 ESOP clamping

If 0.13 × T_F < 4,394,667 (or 0.11 × at low-ESOP cells), top-up
clamps to zero per §6.2. Document any clamping cell.

---

## 10. Return solver (per CANONICAL_INPUTS §10)

For each of 5 exits, solve for Series F ownership % our $10M check
needs to return 3× net ($30M) or 4× net ($40M). 10 keys total. Same
flagging as v3: infeasible cells emit a value > 1.0 or unbounded
verbatim — scorer compares verbatim. At $700M and $1,200M the F class
proceeds under preference election may cap below 3× regardless of
ownership %; flag those cells.

---

## 11. Tracked outputs (per CANONICAL_INPUTS §12)

Both builders produce these named values. Cell-level comparison uses
these keys. Dot-notation keys → Python JSON keys identical → Excel
defined names identical (`.` → `_`).

**Total tracked keys: 167** (vs v3's 124).

### 11.1 Pro-forma ownership — 15 keys

    ownership.founders.pct                  # F1+F2+F3, post-F, AFTER tender
    ownership.employees.pct                 # E1–E10 RSAs, post-F
    ownership.options_issued.pct            # issued options post-F-refresh
    ownership.options_unissued.pct          # unissued pool post-F-refresh
    ownership.seed.pct                      # Seed pref, as-conv (ratio = 1.0)
    ownership.series_a.pct                  # A pref, as-conv (ratio = 1.0)
    ownership.series_b.pct                  # B pref, as-conv (ratio = 1.0)
    ownership.series_c.pct                  # C pref, as-conv (ratio = 1.0)
    ownership.series_d.pct                  # D pref, as-conv at COMPOUND ratio (v3·v4)
    ownership.series_e.pct                  # E pref, as-conv at v4-adjusted ratio
    ownership.note_a.pct                    # NEW: Note A converted shares
    ownership.note_b.pct                    # NEW: Note B converted shares
    ownership.series_f_new.pct              # F primary investors (lead + our fund)
    ownership.f_lead_secondary.pct          # NEW: F_LEAD's tender holding (common from founders)
    ownership.our_fund.pct                  # = (1/3) × series_f_new

Sum of all 15 = 1.000000 (to 1e-9 absolute).

### 11.2 Anti-dilution — 12 keys (6 × 2)

For class in {seed, series_a, series_b, series_c, series_d, series_e}:

    antidilution.{class}.new_conversion_price_usd
    antidilution.{class}.new_ratio

For seed / A / B / C: dormant → emit OCP unchanged, ratio = 1.0.

For series_d: emit v4-adjusted NCP_D and **COMPOUND ratio = r_D(v3)
× r_D(v4)** per §7.2. Load-bearing v4-vs-v3 distinction: NCP is
v4-adjusted only (computed from OCP $58.118347); ratio is total
compound adjustment from D's original $60.00 OCP.

For series_e: emit v4-adjusted NCP_E and r_E(v4) (first AD).

### 11.3 Note conversion — 10 keys (2 × 5)

For note in {note_a, note_b}:

    notes.{note}.cap_implied_pps_usd        # cap / pre-money-FD-at-conversion
    notes.{note}.discount_implied_pps_usd   # F_PPS × (1 − discount)
    notes.{note}.conversion_pps_usd         # min of the two
    notes.{note}.conversion_branch          # "cap" | "discount"
    notes.{note}.shares_issued              # principal / conversion_pps

Expected branches:

- `notes.note_a.conversion_branch = "cap"`
- `notes.note_b.conversion_branch = "discount"`

`shares_issued` is an integer share count.

### 11.4 Tender mechanics — 3 keys

    tender.shares_transferred               # int: $15M / (0.85·F_PPS)
    tender.usd_to_founders                  # = $15,000,000
    tender.f_lead_holding_secondary_shares  # = shares_transferred; cross-check

`tender.usd_to_founders` is fixed ($15M) but tracked — Excel cell must
be a formula referencing the secondary_raise input, not a hardcode.

### 11.5 Round solution — 4 keys

    round.f_pps_usd                         # solved F PPS
    round.post_f_fd_shares                  # T_F, integer
    round.our_fund_shares                   # $10M / F_PPS, integer
    round.f_lead_primary_shares             # $20M / F_PPS, integer

Exposed so builders agree on fixed-point convergence; v3 did not
expose these because its fixed point was simpler.

### 11.6 Waterfall — 95 keys (5 exits × (12 dollars + 7 elections))

For exit in {700, 1200, 2000, 4000, 8000} (units: $M):

Dollar fields (12 per exit = 60 total):

    waterfall.{exit}M.series_f.dollars
    waterfall.{exit}M.series_e.dollars
    waterfall.{exit}M.series_d.dollars
    waterfall.{exit}M.series_c.dollars
    waterfall.{exit}M.series_b.dollars
    waterfall.{exit}M.series_a.dollars
    waterfall.{exit}M.seed.dollars
    waterfall.{exit}M.notes.dollars                  # combined Note A + Note B pro-rata
    waterfall.{exit}M.common.dollars                 # founders (post-tender) + RSAs
    waterfall.{exit}M.options_issued.dollars         # net of strike if ITM
    waterfall.{exit}M.our_fund.dollars               # our $10M's share of series_f
    waterfall.{exit}M.total_distributed              # = exit ± $1

Election fields (7 per exit = 35 total):

    waterfall.{exit}M.seed.election                  # "preference" | "convert"
    waterfall.{exit}M.series_a.election
    waterfall.{exit}M.series_b.election
    waterfall.{exit}M.series_c.election
    waterfall.{exit}M.series_d.election
    waterfall.{exit}M.series_e.election
    waterfall.{exit}M.series_f.election

Notes have no elections (§8.4); pro-rata rolled into
`waterfall.{exit}M.notes.dollars`. F_LEAD's tender shares rolled into
`waterfall.{exit}M.common.dollars` (§6.3).

### 11.7 Sensitivity — 18 keys (9 × 2)

For esop in {0.11, 0.13, 0.15}, pre in {1200, 1400, 1800} $M:

    sensitivity.esop_{11,13,15}.pre_{1200,1400,1800}M.our_pct
    sensitivity.esop_{11,13,15}.pre_{1200,1400,1800}M.dollars_at_2000M_exit

Underscored keys without decimals: `esop_11`, `esop_13`, `esop_15`;
`pre_1200M`, `pre_1400M`, `pre_1800M`.

### 11.8 Return solver — 10 keys (2 × 5)

    solver.3x.exit_{700,1200,2000,4000,8000}M.required_pct
    solver.4x.exit_{700,1200,2000,4000,8000}M.required_pct

### 11.9 Total count

| Section | Keys |
|---|---|
| 11.1 Ownership                                  | 15  |
| 11.2 Anti-dilution (6 × 2)                      | 12  |
| 11.3 Note conversion (2 × 5)                    | 10  |
| 11.4 Tender mechanics                           | 3   |
| 11.5 Round solution                             | 4   |
| 11.6 Waterfall (5 × 12 dollars + 5 × 7 elects)  | 95  |
| 11.7 Sensitivity (9 × 2)                        | 18  |
| 11.8 Return solver (2 × 5)                      | 10  |
| **Total**                                       | **167** |

### 11.10 Scorer dimensions (per CANONICAL_INPUTS §12.10)

- `cells_correct_pct` — 167 keys within tolerance.
- `elections_correct_pct` — 35 election strings (v4) vs 25 (v3).
- `note_branches_correct_pct` — 2 note `conversion_branch` strings
  (new in v4; isolates cap-vs-discount primitive).
- `audit_probe_pass` — perturbed-input recomputation reconciles
  against truth_perturbed.json.

---

## 12. Reconciliation tolerance (per CANONICAL_INPUTS §13)

| Field type | Tolerance |
|---|---|
| Ratios (`*.new_ratio`)                         | abs ≤ 1e-9                        |
| Percentages (`*.pct`, `*.required_pct`)        | abs ≤ 1e-6                        |
| Dollar amounts (`*.dollars`, `*.usd`)          | abs ≤ $1 OR rel ≤ 1e-4 (greater)  |
| Share counts (`*.shares*`, `shares_issued`)    | abs ≤ 1e-2 (effectively integer)  |
| PPS / conversion prices (`*.pps_usd`, `*.conversion_price_usd`) | abs ≤ $1e-4 OR rel ≤ 1e-6 |
| Election strings (`*.election`)                | exact match                       |
| Conversion-branch strings (`*.conversion_branch`) | exact match                    |

Any violation blocks the benchmark.

---

## 13. Fixed-point math sketch (F_PPS solver)

Coupled fixed point in **five unknowns** (per §4.5): `F_PPS`, `T_F`,
`r_E`, `r_D(v4)`, `N_A_shares`, `N_B_shares`.

### 13.1 Coupling

1. **`F_PPS` ↔ `T_F`**: pre-money / post-money identity.
   `F_PPS × pre_money_FD_at_conversion = pre_money_F = $1,400,000,000`,
   where the denominator INCLUDES ESOP top-up and converted notes but
   EXCLUDES new F primary shares.

2. **`T_F` ↔ all other unknowns**: every class adjustment changes T_F.
   - `T_F = common(post-tender) + RSAs + options_issued +
     options_unissued_post_F + Seed + A + B + C + D_compound + E_adj
     + new_F_primary + N_A + N_B`
   - `D_compound = 1,720,627 × r_D(v4)`
   - `E_adj = 6,102,667 × r_E`
   - `options_unissued_post_F = max(0, 0.13 × T_F − 2,500,000)`
   - `new_F_primary = $30M / F_PPS`
   - Founders(post-tender) = `8,000,000 − $15M / (0.85·F_PPS)`. Tender
     does NOT change T_F (shares move within FD).

3. **`r_E`, `r_D(v4)` depend on `F_PPS`** via WA formula, with
   `C = $30M / F_PPS`. r_D uses OCP $58.118347 (NOT $60).

4. **`N_A_shares`, `N_B_shares` depend on both `F_PPS` and `T_F`**:
   - `cap_PPS_X = cap_X / pre_money_FD_at_conversion`
   - `discount_PPS_X = F_PPS × (1 − discount_X)`
   - `conversion_PPS_X = min(cap_PPS_X, discount_PPS_X)`
   - `N_X_shares = principal_X / conversion_PPS_X`

5. `pre_money_FD_at_conversion = T_F − new_F_primary` (includes ESOP
   top-up and BOTH notes' converted shares).

### 13.2 Convergence

System contracts at < 0.05/pass at v3-style inputs. 30–50 iterations
at Decimal precision 60 → 1e-30. Natural order:

1. Init `F_PPS = $37.5`, `T_F = 38.3M`, `r_E = r_D(v4) = 1.0`,
   `N_A = N_B = 0`.
2. Compute `cap_PPS_A`, `cap_PPS_B` from current T_F.
3. Compute `discount_PPS_A`, `discount_PPS_B` from current F_PPS.
4. Pick `min` per note → `N_A_shares`, `N_B_shares`.
5. Compute `r_E`, `r_D(v4)` from WA formula at current F_PPS.
6. Compute `T_F` from FD-sum equation.
7. Compute new `F_PPS = pre_money_F / pre_money_FD_at_conversion`.
8. Repeat until ‖Δ‖ < 1e-30.

### 13.3 Branch verification gate (per CANONICAL_INPUTS §6.4)

Closed-form solvers (assuming A=cap, B=discount) MUST verify at
converged F_PPS:

- `cap_PPS_A < discount_PPS_A` ($27.32 < $30.00 ✓ at base)
- `cap_PPS_B > discount_PPS_B` ($42.50 > $31.875 ✓ at base)

If branches fail at converged F_PPS, closed-form assumption is wrong
— iterate.

### 13.4 Tender excluded from fixed point

The $15M secondary does NOT enter the F_PPS solver. F_PPS is
determined by primary post-money / FD; tender executes AFTER F_PPS is
known (at 0.85·F_PPS, founders → F_LEAD, T_F unchanged). Tender's
only schema effects: `ownership.founders.pct`,
`ownership.f_lead_secondary.pct`, `tender.*`, and the `common.dollars`
waterfall line.

### 13.5 Sanity asserts (per CANONICAL_INPUTS §16)

Builders `assert` final outputs within ±5% of:

| Expected output | Approximate value |
|---|---|
| F_PPS                                 | $37–$38 |
| AD on E: NCP_E                        | ≈ $47.7 (r_E ≈ 1.03) |
| AD on D: NCP_D(v4)                    | ≈ $55 (r_D(v4) ≈ 1.05; compound ≈ 1.087) |
| Note A: CAP CONVERTS at PPS           | ≈ $27.32 (≈ 183K shares) |
| Note B: DISCOUNT CONVERTS at PPS      | ≈ $31.875 (≈ 92–94K shares) |
| Post-F FD T_F                         | ≈ 38.3M shares |
| Tender shares                         | ≈ 470K |

Truth values come from full-precision Decimal arithmetic; these are
correctness gates, not targets.

---

## 14. Audit-probe sub-test (per CANONICAL_INPUTS §11)

A second file `inputs/cap_table_perturbed.xlsx` is shipped alongside
canonical `cap_table.xlsx`. **Identical** except:

- `RoundTerms.series_f_primary_raise_usd` = **$33,000,000** (vs $30M)

### 14.1 Purpose

Verifies the API's **output workbook has live formulas**. When the
harness swaps perturbed inputs into the API's output workbook, every
dependent cell (F_PPS, ownership %s, waterfall dollars, sensitivity
cells) must recompute. A hardcoded-output workbook fails audit-probe
even if it passes the canonical numeric check.

### 14.2 Truth provenance

Audit-probe truth built by running `truth.py` with
`F_PRIMARY_RAISE = d(33_000_000)`, emits to `truth/truth_perturbed.json`.

### 14.3 Harness extension

Single new comparison step: "re-evaluate workbook with perturbed input,
reconcile against truth_perturbed.json". Surfaces `audit_probe_pass`
boolean (or the same `cells_correct_pct` over perturbed truth set).

### 14.4 Audit-probe key set

Same 167 keys as canonical truth (§11). Same tolerances (§12).

### 14.5 Expected differences under perturbation

A correctly-formula-driven workbook shows different values for every
key dependent on primary raise: F_PPS shifts (depends on FD
denominator); T_F rises (more F shares); F ownership rises; older
classes dilute slightly; both notes' shares change (cap-PPS depends on
T_F; discount-PPS depends on F_PPS); both AD ratios change (B and C
in WA formula shift); every waterfall dollar shifts; every sensitivity
cell shifts; every solver cell shifts.

A hardcoded workbook shows ZERO shifts. Audit-probe is a binary
"live-formula vs hardcode" test with strong gradient between "all 167
shift correctly" (pass) and "none shift" (catastrophic fail).

---

## 15. What builders must NOT do

- NOT read each other's output files.
- NOT consult external "ground truth" sources (Carta, published
  waterfalls).
- NOT use LLMs to compute numbers at runtime.
- Build every number from first principles, from this spec and
  CANONICAL_INPUTS.md.
- NOT re-derive v3's pre-F state. Use locked values in §2.1.
- NOT hardcode the 35 election strings or 2 note-branch strings.
  Election cells must derive from max-selection formula (Excel) or
  Nash solver (Python) reading per-class preference and pro-rata.
  Note-branch cells must derive from `min`-comparison formula (Excel)
  or `min` (Python) reading cap-implied and discount-implied PPS.
  Scorer checks strings, but strings must be *computed*, not asserted.
- NOT alter any locked input (CANONICAL_INPUTS §3–§9). If a builder
  finds arithmetic infeasibility, STOP and surface to human; do not
  silently "fix" by adjusting an input.

---

End of spec.md.
