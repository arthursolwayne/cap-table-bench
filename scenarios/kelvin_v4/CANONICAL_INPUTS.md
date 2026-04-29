# Kelvin Dynamics — Canonical Inputs v4 ("Series F: Tender + Convertibles + Down-Round AD")

This document is the **single source of truth for INPUTS** to scenario
`kelvin_v4`. Four downstream agents consume it:

1. **Spec-writer agent** → produces `scenarios/kelvin_v4/spec.md`
2. **Python builder agent** → produces `truth/truth.py` → emits `truth/truth.json`
3. **Excel builder agent** → produces `inputs/cap_table.xlsx`,
   `inputs/cap_table_perturbed.xlsx`, and `truth/ground_truth.xlsx`
4. **Harness/Tetra-prompt agent** → extends the harness, writes the
   construction prompt the API consumes

All four agents must agree on every number, name, and key in this file.
This file is INPUT-ONLY: it locks the cap-table facts, the deal terms, the
modeling rules, and the truth-key schema. It does **not** specify how to
solve the system, what tabs to lay out, or what the prompt should say.
Those decisions belong to the spec-writer / builders.

The scenario identifier on disk is `kelvin_v4`; the company inside the
model is **Kelvin Dynamics**, the same fictional climate-hardware startup
as v3, now one round further along (Series E has closed; Series F is
proposed).

This scenario serves as a **second observation in the templated-complex
regime** alongside kelvin_v3_t1, validating Tetra's +12.5pp lead over
Opus on dense templated cap-table modeling. The complexity dial is turned
**up** vs v3: tender (secondary), two converting notes with mutually
inverted cap-vs-discount conclusions, and a slight down-round triggering
broad-based weighted-average AD on **two** prior classes (E and D).

---

## 1. Inheritance from v3 (locked, do not re-derive)

v4's pre-Series-F state IS v3's post-Series-E state, with one structural
addition (two convertible notes outstanding, see §6). Every share count,
PPS, OCP, and dollar amount in §2–§5 below MUST match the corresponding
v3 number to the digit. If a v4 builder re-derives one of these and
disagrees with v3, the v4 builder is wrong.

- All v3 stakeholders carry over (founders F1/F2/F3, employees E1–E10).
- All five v3 priced rounds (Seed, A, B, C, D, E) are immutable history.
- v3's Series E has CLOSED. The post-E cap table from v3's truth.json IS
  the pre-F cap table here.
- Series E's anti-dilution-adjusted Series D ratio (~1.0324) carries
  forward as Series D's "current" ratio entering Series F.

### 1.1 Pre-F (post-E) share counts — locked to v3's solved fixed point

These are the **starting state** for all v4 math. They come from v3's
truth.json. Builders MUST use these values verbatim; do NOT re-solve
the v3 fixed point inside v4.

| Holder / Class | Pre-F shares | Origin |
|---|---|---|
| F1 (CEO)                          | 4,000,000  | v3 §2 |
| F2 (CTO)                          | 2,500,000  | v3 §2 |
| F3 (Chief Scientist)              | 1,500,000  | v3 §2 |
| Founders subtotal                 | 8,000,000  | sum |
| E1–E10 (RSAs, 100k each)          | 1,000,000  | v3 §2 |
| Common subtotal                   | 9,000,000  |       |
| Seed preferred (as-conv, ratio 1) | 5,000,000  | v3 §3 |
| Series A (as-conv, ratio 1)       | 5,000,000  | v3 §3 |
| Series B (as-conv, ratio 1)       | 3,000,000  | v3 §3 |
| Series C (as-conv, ratio 1)       | 2,400,000  | v3 §3 |
| Series D **adjusted** (ratio 1.03237623…) | 1,720,627 | v3 §7 / truth.json: pref_D_adj_int |
| Series E (new, base case)         | 6,102,667  | v3 §13 / 1.5B post-money fixed point |
| Options issued (post-E)           | 2,500,000  | v3 §7 (unchanged through E) |
| Options unissued (post-E refresh) | 1,894,667  | v3 §7 / 0.12·T − 2,500,000 at v3's T |
| **Pre-F FD total**                | **36,617,961** | T from v3 §13 |

Builders implementing v4 MAY round Series D and Series E share counts to
nearest integer for display, but MUST use full Decimal precision in
denominators (matches v3 §13 convention). The pre-F FD total `T_E`
(v3's post-money FD) is **`T_E = 36,617,961`** to nearest integer; for
Decimal arithmetic, `T_E = post_money_E / E_PPS = 1,800,000,000 /
49.158311… ≈ 36,617,960.84`. Use `T_E` as the literal Decimal value,
not the rounded integer, in the v4 fixed point.

### 1.2 Reference v3 numbers (informational, for sanity asserts only)

- E_PPS ≈ $49.158311
- NCP_D (post-E) ≈ $58.118347
- r_D (post-E) ≈ 1.03237623
- post-E FD T ≈ 36,617,961
- post-E option pool = 12.0% of T = 4,394,667 shares (= 2,500,000 issued
  + 1,894,667 unissued)

These are **not** v4 truth keys. They are sanity checks the v4 builders
should `assert` against to confirm they inherited v3's state correctly.

---

## 2. Stakeholders (full list, post-F)

Naming convention extends v3's `F{n}` / `E{1..10}` pattern. New v4-only
holders carry the prefix of the round that created them.

| ID | Role | Source |
|---|---|---|
| F1, F2, F3            | Founders (carry from v3 §2) | v3 |
| E1–E10                | Early-employee RSAs (carry from v3) | v3 |
| SEED1                 | Seed preferred holder (single LP) | v3 |
| A1                    | Series A preferred holder | v3 |
| B1                    | Series B preferred holder | v3 |
| C1                    | Series C preferred holder | v3 |
| D1                    | Series D preferred holder | v3 |
| E1_INV–E10_INV (10)   | Series E preferred holders (10 LPs, equal pro rata of 6.10M E shares) | v3 |
| **NA1**               | **Convertible Note A holder** (new in v4) | v4 §6 |
| **NB1**               | **Convertible Note B holder** (new in v4) | v4 §6 |
| **F_LEAD**            | **Series F lead investor** (new in v4) | v4 §5 |
| **F_FUND**            | **Our fund** ($10M of $30M Series F primary, new in v4) | v4 §5 |

Note on `E1_INV` vs `E1`: `E1` is an early-employee RSA holder (from v3);
`E1_INV` is a Series E preferred LP. The naming collision is
deliberately disambiguated with the `_INV` suffix.

**Aggregation convention for tracked keys** (matches v3 §2):

- Founders F1+F2+F3 → one `ownership.founders.pct` value
- Employees E1–E10 → one `ownership.employees.pct` value
- Series E investors E1_INV–E10_INV → one `ownership.series_e.pct` value
  (Series E is now an EXISTING class for v4, not the new round)
- F lead + our fund + any other new F primaries → one
  `ownership.series_f_new.pct` value
- Our fund alone → `ownership.our_fund.pct` (= 1/3 of `series_f_new`)
- Note A holder → `ownership.note_a.pct`
- Note B holder → `ownership.note_b.pct`

---

## 3. Prior-round terms (verbatim from v3, locked)

All five priced rounds prior to Series E are unchanged from v3 §3.

| Series | Date    | OCP ($) | Shares    | Invested ($)  | LP type            | AD                           |
|---|---|---|---|---|---|---|
| Seed     | 2020-Q2 | 1.00    | 5,000,000 | 5,000,000     | 1× non-part, conv  | broad-based WA |
| Series A | 2021-Q1 | 3.00    | 5,000,000 | 15,000,000    | 1× non-part, conv  | broad-based WA |
| Series B | 2022-Q2 | 10.00   | 3,000,000 | 30,000,000    | 1× non-part, conv  | broad-based WA |
| Series C | 2023-Q4 | 25.00   | 2,400,000 | 60,000,000    | 1× non-part, conv  | broad-based WA |
| Series D | 2025-Q2 | 60.00   | 1,666,667 | 100,000,000   | 1× non-part, conv  | broad-based WA |

Seniority on these classes: standard reverse-chronological,
**E > D > C > B > A > Seed**, with F now senior to all of them (see §5).

Series D has already been adjusted once (post-E); its `OCP` for v4 AD
purposes is **the post-E NCP_D = $58.118347**, not the original $60.00.
This is the v4-specific subtlety: the AD test on D in v4 reads "is F PPS
< $58.118347?", not "< $60". The "OCP" field on Series D in any v4
input artifact is therefore $58.118347.

---

## 4. Series E terms (verbatim from v3 §7, now CLOSED)

Series E is no longer a "proposed" round in v4 — it has closed. Its terms
are still tracked because Series E participates in the v4 waterfall as a
priced preferred class with conversion election.

| Field | Value | Source |
|---|---|---|
| Pre-money       | $1,500,000,000 | v3 §7 |
| Raise           | $300,000,000   | v3 §7 |
| Post-money      | $1,800,000,000 | v3 §7 |
| PPS (E_PPS)     | $49.158311…    | v3 §13 fixed point |
| Shares issued   | 6,102,666.84… (Decimal); 6,102,667 (int) | derived |
| LP              | 1× non-participating with conversion election | v3 §7 |
| AD              | broad-based WA  | v3 §7 |
| Seniority       | senior to D / C / B / A / Seed; junior to F | v3 §7, v4 §5 |
| OCP (for v4 AD) | $49.158311 (= original E PPS) | derived |

E_PPS is treated as Series E's OCP for purposes of v4's anti-dilution
trigger test on Series E. Because v4 is a slight down round (F PPS ≈
$37.5 < $49.16), Series E's AD will TRIGGER in v4.

---

## 5. Series F terms (NEW in v4 — the round being modeled)

| Field | Value |
|---|---|
| Pre-money                 | **$1,400,000,000** (a slight down round vs E's $1.8B post) |
| Primary raise             | **$30,000,000** |
| Secondary tender raise    | **$15,000,000** |
| Tender PPS multiplier     | **0.85× of F PPS** (secondary buys existing shares at a discount) |
| Total cash in             | $45,000,000 (only $30M is dilutive primary) |
| ESOP refresh target post-F| **13.0% of post-F FD** (pre-money refresh, borne by pre-F holders) |
| LP                        | 1× non-participating with conversion election |
| AD                        | broad-based WA |
| Seniority                 | **senior to E, D, C, B, A, Seed**; senior to all converted notes |
| Lead investor             | F_LEAD writes $20M of the $30M primary |
| **Our fund's check**      | **$10,000,000 = 1/3 of the $30M Series F primary** |

### 5.1 Primary vs secondary mechanics

- **Primary ($30M)**: new Series F shares are issued. Adds to FD count.
  Allocated to F_LEAD ($20M) and F_FUND/our fund ($10M).
- **Secondary tender ($15M @ 0.85·F_PPS)**: F_LEAD additionally buys
  existing shares from founders (F1+F2+F3 sell pro rata across the three
  founders). **No new shares are issued** for the secondary; share
  ownership transfers from founders to F_LEAD at the tender price.
  - Tender share count = $15,000,000 / (0.85 × F_PPS).
  - Founders' pre-F balance is reduced by this share count, allocated
    pro rata across F1/F2/F3 by their pre-F holdings (4M / 2.5M / 1.5M
    → 0.50 / 0.3125 / 0.1875 of the tender).
  - F_LEAD's post-F holding = primary F shares ($20M / F_PPS) + tender
    shares ($15M / (0.85·F_PPS)).
- **Our fund's holding**: $10M / F_PPS new Series F primary shares
  only. Our fund does NOT participate in the tender.

### 5.2 ESOP refresh

- Same convention as v3: pre-money top-up, target 13.0% of post-F FD.
- Top-up dilutes pre-F holders (founders, RSAs, all priced classes
  through E, both note holders); does NOT dilute Series F primary
  investors.
- Top-up amount = `0.13 × T_F − options_pool_pre_F`, where the
  pre-F option pool already has 4,394,667 shares (post-E ESOP refresh)
  and `T_F` is the post-F FD share count (the unknown in the round
  fixed point).

### 5.3 Seniority and waterfall position

Reverse-chronological after F is added:

    F  >  E  >  D  >  C  >  B  >  A  >  Seed  >  common+RSAs+options

Notes A and B convert to F-priced common-equivalent shares at closing
(per §6.5) and ride pro rata with common in the waterfall convert side.

---

## 6. Convertible notes (NEW in v4 — both convert at F)

Two convertible notes are outstanding pre-F, issued at unspecified
historical dates between v3's E close and v4's F close. They are **input
constants**: principal, valuation cap, and discount are locked here. Both
convert at F closing using the standard `min(cap-implied PPS, F_PPS ×
(1 − discount))` mechanic.

| Field | Note A (NA1) | Note B (NB1) |
|---|---|---|
| Principal              | $5,000,000          | $3,000,000 |
| Valuation cap          | $1,000,000,000      | $1,600,000,000 |
| Discount               | 20%                 | 15% |
| Cap-implied PPS        | cap / pre-F-FD-at-conversion (see §6.4) | cap / pre-F-FD-at-conversion |
| Discount-implied PPS   | F_PPS × (1 − 0.20)  | F_PPS × (1 − 0.15) |
| Conversion price       | min(cap-implied, discount-implied) | min(cap-implied, discount-implied) |
| Expected outcome       | **CAP CONVERTS** (cap-implied < discount-implied) | **DISCOUNT CONVERTS** (discount-implied < cap-implied) |
| Expected PPS           | ≈ $27.32            | ≈ $32.5 (= F_PPS × 0.85) |
| Expected shares        | ≈ 183,000           | ≈ 92,000 |
| Interest accrued       | **$0** (zero-coupon, ignore in this scenario) | $0 |

### 6.1 Why these specific cap/discount values

The cap and discount on each note were chosen so the **two notes resolve
to opposite conclusions** at the expected F PPS. This is the modeling
discriminator: the solver must independently compute both `cap-implied`
and `discount-implied` for each note and pick the lower.

- Note A's $1.0B cap is well below F's $1.4B pre — so Note A converts at
  cap and the cap-implied PPS dominates the 20% discount.
- Note B's $1.6B cap is above F's $1.4B pre — so the cap is non-binding
  and Note B converts at discount.

The two notes share a structural form but flip on the cap-vs-discount
test. A model that hardcodes either branch will fail one note.

### 6.2 Cap-implied PPS denominator

The convention used: `cap-implied PPS = cap / pre-money-FD-at-conversion`,
where pre-money-FD-at-conversion is the FD share count **excluding** the
new F primary shares but **including** all pre-F holders, the ESOP
top-up, and the OTHER note's converted shares (i.e., the standard
"pre-money-shares" definition that includes the round's other
convertibles but excludes the round's new priced equity).

For numerical concreteness, in the base case the cap-implied PPS for
Note A solves to approximately **$27.32**, and for Note B to
approximately **$42.50**, against an F_PPS of ≈ $37.5 → discount-implied
PPS_A ≈ $30, PPS_B ≈ $31.875. Therefore:

- min($27.32, $30) → Note A **converts at cap** ≈ $27.32 → ~183,000 shares
- min($42.50, $31.875) → Note B **converts at discount** ≈ $31.875 → ~94,000 shares

(Solver will refine; the document states ~$32.5/~92K as the locked
expectation. Builders match the solver, not these approximations.)

### 6.3 Where the converted shares go in the FD count

Both notes' converted shares appear in `T_F`, the post-F FD count. They
join the pro-rata pool on the convert side of the waterfall and never
take preference (they are common-equivalent post-conversion). They are
NOT new preferred classes; they do not have AD; they do not appear in
§11.4. They DO appear in §11.1 as their own ownership lines.

### 6.4 Coupling with the round fixed point

The note conversions, the ESOP refresh, the F PPS, the AD adjustments to
E and D, and the post-F FD count are all **coupled**. The fixed-point
solver must iterate (or solve analytically) over five unknowns:

- F_PPS
- T_F (post-F FD share count)
- r_E (Series E adjusted ratio post-F AD)
- r_D (Series D adjusted ratio post-F AD; layered atop v3's already-
  adjusted ratio — see §7.4)
- N_A_shares, N_B_shares (the two notes' converted share counts; one
  caps-out, the other discounts-out)

The system is well-conditioned (contraction << 0.05 per pass at v3-style
inputs); 30–50 iterations at Decimal precision 60 should converge to
1e-30. Builders may also derive a closed-form by case analysis on the
note conversion branches — but they MUST verify the case assumption
holds at the converged F_PPS (i.e., Note A's cap-implied < Note A's
discount-implied; Note B's cap-implied > Note B's discount-implied).

### 6.5 No interest, no MFN, no most-favored-nation triggers

Both notes are zero-coupon for the purpose of this scenario; no MFN
clause; no senior preference at exit (they are common-equivalent
post-conversion); they convert in full at F closing — no roll-forward.

---

## 7. Modeling rules (locked across all builders)

### 7.1 Liquidation preferences and conversion election

Every priced preferred class — Seed, A, B, C, D, E, AND F — is **1×
non-participating with conversion election**. Per exit, each class
chooses preference vs convert to maximize its own proceeds, given the
other classes' Nash-optimal choices. This is exactly v3's mechanic
extended one class up.

- 35 tracked elections in the base waterfall: 7 priced classes × 5 exits
  = 35. (v3 had 5 × 5 = 25.)
- Series F's election IS tracked in v4 (unlike v3, which did not track
  Series E's). Rationale: F is the new round being modeled, but F also
  participates in the waterfall as a regular priced class once the deal
  closes. Tracking F's election keeps v4 symmetric with the other six
  classes — and adds 5 more observations in the templated regime
  for the senior-most class's behavior.
- Notes A and B do NOT have elections; they are common-equivalent
  post-conversion and ride the pro-rata pool. Their share count goes
  into the convert-side denominator.

### 7.2 Anti-dilution

Broad-based weighted-average on every prior class. Trigger test per
class in v4:

- Seed: F_PPS < $1.00? — dormant (test fails at any F_PPS > $1)
- A:    F_PPS < $3.00? — dormant
- B:    F_PPS < $10.00? — dormant
- C:    F_PPS < $25.00? — dormant
- D:    F_PPS < $58.118347 (post-E NCP_D)? — **TRIGGERS** at F_PPS ≈ $37.5
- E:    F_PPS < $49.158311 (E_PPS)? — **TRIGGERS** at F_PPS ≈ $37.5

**Two AD triggers in v4** (vs one in v3). Each requires its own broad-
based WA solve, layered into the fixed point. Series D's adjustment in
v4 is **on top of** v3's adjustment — so D's effective new ratio in the
truth schema is r_D(v4) computed with OCP = $58.118347 (NOT $60), and
D's as-converted share count is `1,666,667 × r_D(v3) × r_D(v4) =
1,720,627 × r_D(v4)` to nearest integer.

The broad-base `A` for v4's AD computation = pre-F FD = 36,617,961
(post-E T from §1.1).
The `B` for class X = raise / OCP_X (with OCP_E = $49.158311, OCP_D =
$58.118347).
The `C` for class X = new F shares actually issued = $30M / F_PPS.

The note conversions are EXCLUDED from `B` and from `C` per
standard NVCA broad-based WA convention (notes are not "new equity at a
lower price"; they are converting prior-issued instruments). Builders
should document this choice in spec.md.

### 7.3 ESOP refresh (pre-money, target 13%)

Same mechanic as v3 but at 13% (vs v3's 12%), and the pre-F option pool
already contains 4,394,667 shares (issued 2,500,000 + unissued
1,894,667). Top-up = `max(0, 0.13 × T_F − 4,394,667)`. The unissued
post-F = `0.13 × T_F − 2,500,000` (issued unchanged).

If 13% × T_F is **less** than 4,394,667 (e.g., at very low pre-money
sensitivity cells), the top-up clamps to zero rather than going
negative — emit "no top-up" in spec.md as the convention.

### 7.4 Secondary tender (no new shares)

The $15M secondary at 0.85 × F_PPS:

- Tender shares = $15,000,000 / (0.85 × F_PPS) ≈ $15M / $31.875 ≈ 470,588
- These shares move from founders (pro rata: F1 50%, F2 31.25%, F3
  18.75%) to F_LEAD.
- T_F is **unchanged** by the tender (no issuance).
- The ownership lines `ownership.founders.pct` and `ownership.f_lead.pct`
  reflect post-tender holdings.
- F_LEAD's primary shares + tender shares both ride as Series F
  preferred for waterfall purposes (i.e., the tender shares ARE Series F
  preferred — the founders sold preferred-equivalent rights, not common,
  per the typical structured-secondary convention; document explicitly
  in spec.md).

For a cleaner alternative convention (tender = common purchase, not
preferred), spec-writer may specify either way as long as it's locked
and the truth keys reflect it. **Default for v4: tender purchases
common shares from founders at $0.85·F_PPS; F_LEAD therefore holds a
mix of (Series F preferred from primary) + (common from secondary).**
This matches the most common real-world structured-secondary mechanic
and makes the waterfall cleaner.

### 7.5 Convertible note conversion at closing

Both notes convert in full at F closing per §6. Their converted shares
become common-equivalent (NOT a new preferred class) and join the
convert-side pro-rata pool in every waterfall scenario. They do not
have a preference election; they cannot take preference at any exit.

---

## 8. Exit scenarios (matches v3 §8)

Five exits: **$700M, $1,200M, $2,000M, $4,000M, $8,000M**.

Same labels as v3 (`700M`, `1200M`, `2000M`, `4000M`, `8000M`). Same
spread rationale: the bottom two ($700M, $1.2B) put F (and probably E
and D) into preference; the top two ($4B, $8B) push everyone into
convert; $2B is the sensitivity reference.

---

## 9. Sensitivity grid (axes match v3 §9, values shifted)

ESOP refresh × Series F pre-money. **3 × 3 = 9 cells.** Output: our $10M
check's post-F ownership %, and $ proceeds at $2B exit.

| ESOP refresh \ Pre-money F | $1,200M | $1,400M (base) | $1,800M |
|---|---|---|---|
| 11% | cell | cell | cell |
| 13% (base) | cell | **base case** | cell |
| 15% | cell | cell | cell |

Each cell re-solves the full round fixed point (F_PPS varies → AD on D
and E retriggers per cell → notes' cap/discount branches must be re-
evaluated per cell — at $1.8B pre, Note A may flip from CAP to DISCOUNT,
which makes the sensitivity grid genuinely test the cap-vs-discount
logic at multiple operating points, not just the base case).

---

## 10. Return solver (matches v3 §10)

For each of 5 exits, solve for the Series F ownership % our $10M check
would need to return:

- 3× net ($30M back)
- 4× net ($40M back)

10 keys total. Same flagging rule as v3 (infeasible cells emit a value
> 1.0 or as the unbounded numeric verbatim — the scorer compares
verbatim).

---

## 11. Audit-probe perturbed input

A second file `inputs/cap_table_perturbed.xlsx` is shipped alongside
the canonical `cap_table.xlsx`. It is **identical** to the canonical
input except:

- `RoundTerms.series_f_primary_raise_usd` = **$33,000,000** (vs $30M)

All other inputs (pre-money, tender, notes, ESOP target, etc.) are
unchanged. The audit-probe sub-test exists to verify that the
**output** workbook the API produces under the canonical input has
**live formulas** — i.e., when the harness re-opens the API's output
workbook with the perturbed inputs swapped in, every dependent cell
(F_PPS, ownership %s, waterfall dollars, sensitivity cells, etc.)
recomputes. A workbook that hardcodes outputs will fail the audit-probe
even if it passes the canonical-input numeric check.

The audit-probe truth.json is built by running the same `truth.py`
with `F_PRIMARY_RAISE = d(33_000_000)` and emits to
`truth/truth_perturbed.json`. The harness compares the API output's
recomputed cells against this perturbed truth.

The harness extension required for v4 is a single new comparison step
("re-evaluate workbook with perturbed input, reconcile against
truth_perturbed.json"). The scorer dimension surfaced is
`audit_probe_pass` (boolean, or the same `cells_correct_pct` over the
perturbed truth set).

The audit-probe truth set IS the same key list as the canonical truth
(same 167 keys, see §12). Tolerances are the same.

---

## 12. Truth-key schema (full enumeration)

Dot-notation keys → Python JSON keys identical → Excel defined names
identical (with `.` → `_` substitution). All sections below match v3
§11's structure where applicable, with v4 additions/extensions
explicitly marked.

### 12.1 Pro-forma ownership — 15 keys (v3 had 11)

    ownership.founders.pct                  # F1+F2+F3, post-F, AFTER tender
    ownership.employees.pct                 # E1–E10 RSAs, post-F
    ownership.options_issued.pct            # issued options post-F-refresh
    ownership.options_unissued.pct          # unissued pool post-F-refresh
    ownership.seed.pct                      # Seed pref, as-conv (ratio = 1.0)
    ownership.series_a.pct                  # A pref, as-conv (ratio = 1.0)
    ownership.series_b.pct                  # B pref, as-conv (ratio = 1.0)
    ownership.series_c.pct                  # C pref, as-conv (ratio = 1.0)
    ownership.series_d.pct                  # D pref, as-conv at COMPOUND adjusted ratio (v3·v4)
    ownership.series_e.pct                  # E pref, as-conv at v4-adjusted ratio
    ownership.note_a.pct                    # NEW: Note A converted shares
    ownership.note_b.pct                    # NEW: Note B converted shares
    ownership.series_f_new.pct              # F primary investors (lead + our fund)
    ownership.f_lead_secondary.pct          # NEW: F_LEAD's tender holding (common from founders)
    ownership.our_fund.pct                  # = (1/3) × series_f_new

Sum of all 15 keys = 1.000000 (to 1e-9).

**Locked count for §12.1: 15 keys.**

### 12.2 Anti-dilution — 6 × 2 = 12 keys (v3 had 10)

For class in {seed, series_a, series_b, series_c, series_d, series_e}:

    antidilution.{class}.new_conversion_price_usd
    antidilution.{class}.new_ratio

For seed/A/B/C: dormant → emit OCP unchanged, ratio = 1.0.
For D: emit v4-adjusted NCP_D and v4-adjusted r_D. **The
"new_conversion_price_usd" is the v4-adjusted NCP; the "new_ratio" is
the COMPOUND ratio = r_D(v3) × r_D(v4)** (i.e., the total adjustment
from D's original $60 OCP to its post-F effective conversion price).
Document this convention in spec.md — it is the load-bearing detail
that distinguishes v4 from v3.
For E: emit v4-adjusted NCP_E and r_E (E's first AD adjustment).

**Locked count for §12.2: 12 keys.**

### 12.3 Note conversion — 2 × 5 = 10 keys (NEW in v4)

For note in {note_a, note_b}:

    notes.{note}.cap_implied_pps_usd        # cap / pre-money-FD-at-conversion
    notes.{note}.discount_implied_pps_usd   # F_PPS × (1 − discount)
    notes.{note}.conversion_pps_usd         # min of the two
    notes.{note}.conversion_branch          # "cap" | "discount"
    notes.{note}.shares_issued              # principal / conversion_pps

Expected branches: `notes.note_a.conversion_branch = "cap"`;
`notes.note_b.conversion_branch = "discount"`.

`shares_issued` is an integer share count.

**Locked count for §12.3: 10 keys.**

### 12.4 Tender mechanics — 3 keys (NEW in v4)

    tender.shares_transferred               # int: $15M / (0.85·F_PPS)
    tender.usd_to_founders                  # = $15,000,000 (cash to founders)
    tender.f_lead_holding_secondary_shares  # = same as shares_transferred; tracked for cross-check

`tender.usd_to_founders` is a fixed input (= $15M), but it appears in
the tracked set for cross-builder verification (the Excel cell must be
a formula = secondary_raise input, not a hardcode in the output tab).

**Locked count for §12.4: 3 keys.**

### 12.5 Round solution — 4 keys (NEW in v4)

    round.f_pps_usd                         # solved F PPS
    round.post_f_fd_shares                  # T_F, post-F fully-diluted shares (integer)
    round.our_fund_shares                   # $10M / F_PPS, integer
    round.f_lead_primary_shares             # $20M / F_PPS, integer

These are exposed because the solver must agree across builders on the
fixed-point convergence; without them we cannot diagnose disagreements
in downstream pro-forma keys. v3 did not expose these explicitly
because v3's fixed point was simpler.

**Locked count for §12.5: 4 keys.**

### 12.6 Waterfall — 5 exits × (12 dollar fields + 7 election fields) = 95 keys

For exit in {700, 1200, 2000, 4000, 8000} ($M):

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

Note: Notes A and B do NOT have elections (per §7.1). Their pro-rata
share is rolled into `waterfall.{exit}M.notes.dollars` for tracking,
but they are part of the convert-side pool mechanically.

**Locked count for §12.6: 95 keys (60 dollars + 35 elections).**

### 12.7 Sensitivity — 9 × 2 = 18 keys (matches v3)

For esop in {0.11, 0.13, 0.15}, pre in {1200, 1400, 1800} $M:

    sensitivity.esop_{11,13,15}.pre_{1200,1400,1800}M.our_pct
    sensitivity.esop_{11,13,15}.pre_{1200,1400,1800}M.dollars_at_2000M_exit

**Locked count for §12.7: 18 keys.**

### 12.8 Return solver — 2 × 5 = 10 keys (matches v3)

    solver.3x.exit_{700,1200,2000,4000,8000}M.required_pct
    solver.4x.exit_{700,1200,2000,4000,8000}M.required_pct

**Locked count for §12.8: 10 keys.**

### 12.9 Total tracked keys

| Section | Keys |
|---|---|
| 12.1 Ownership                                  | 15  |
| 12.2 Anti-dilution (6 × 2)                      | 12  |
| 12.3 Note conversion (2 × 5)                    | 10  |
| 12.4 Tender mechanics                           | 3   |
| 12.5 Round solution                             | 4   |
| 12.6 Waterfall (5 × 12 dollars + 5 × 7 elects)  | 95  |
| 12.7 Sensitivity (9 × 2)                        | 18  |
| 12.8 Return solver (2 × 5)                      | 10  |
| **Total**                                       | **167** |

The Tetra-prompt agent should reference "167 tracked keys" in the
construction prompt rather than the rounded "~165."

### 12.10 Scorer dimensions

- `cells_correct_pct` — fraction of the 167 keys within tolerance.
- `elections_correct_pct` — fraction of 35 election strings exactly
  matching truth (v4 has 35, v3 had 25).
- `note_branches_correct_pct` — fraction of 2 note `conversion_branch`
  strings exactly matching truth (new in v4; isolates the cap-vs-
  discount reasoning primitive).
- `audit_probe_pass` — boolean; truthy iff the perturbed-input
  recomputation reconciles within tolerance against truth_perturbed.json.

---

## 13. Tolerances (matches v3 §12)

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

## 14. Cross-builder contract: which keys come from which artifact

This table tells each downstream agent where each key originates and
what its formula provenance is. The Excel builder produces formulas;
the Python builder produces Decimals; the harness reconciles them.

### 14.1 INPUT artifacts

`inputs/cap_table.xlsx` (canonical) and `inputs/cap_table_perturbed.xlsx`
(audit-probe). Both are clean Carta-style exports with these tabs:

| Tab | Rows / contents |
|---|---|
| **Stakeholders**  | F1, F2, F3, E1–E10, SEED1, A1, B1, C1, D1, E1_INV–E10_INV, NA1, NB1, F_LEAD, F_FUND. Each with role, common shares (where applicable), preferred shares (where applicable). |
| **Securities**    | Seed, A, B, C, D, E rows (one per priced class): PPS, OCP-for-AD-purposes, shares issued, invested, LP type, AD type. |
| **Notes**         | Note A, Note B rows: principal, valuation cap, discount, accrued interest (= 0). |
| **RoundTerms**    | Series F pre-money, primary raise, secondary raise, tender PPS multiplier (0.85), ESOP target post-F (0.13), our check size ($10M), F_LEAD primary check ($20M), 5 exit values. |

The perturbed file differs in **one cell**: `RoundTerms.primary_raise_usd
= 33000000` instead of `30000000`.

### 14.2 Truth artifacts

- `truth/truth.py` (Python builder; reads no xlsx)
- `truth/truth.json` (canonical truth, output of truth.py)
- `truth/truth_perturbed.json` (audit-probe truth, output of truth.py
  with one input flag flipped)
- `truth/ground_truth.xlsx` (Excel builder; formula-driven; defined
  names match keys per §12)

### 14.3 Formula provenance per key section

| Key prefix | Excel tab origin | Python function origin | Notes |
|---|---|---|---|
| `ownership.*`                      | Pro-Forma   | `solve_round` + `compute_ownership` | Sums to 1.0 |
| `antidilution.{seed,a,b,c}.*`      | Pro-Forma "AD block" | `compute_antidilution` (dormant branch) | Trigger test must be a formula, not a hardcode |
| `antidilution.series_d.*`          | Pro-Forma "AD block" | `solve_round` (compound ratio) | NCP formula references the v3-adjusted OCP |
| `antidilution.series_e.*`          | Pro-Forma "AD block" | `solve_round` (E's first AD) | NCP_E formula = `49.158311 × (A+B)/(A+C)` |
| `notes.note_a.*`, `notes.note_b.*` | Pro-Forma "Notes block" | `compute_note_conversion` | The `min(cap, discount)` branch must be an `IF` formula |
| `tender.*`                         | Pro-Forma "Tender block" | direct from inputs | Shares transferred = formula on F_PPS |
| `round.*`                          | Assumptions / Pro-Forma | `solve_round` outputs | The fixed-point output values |
| `waterfall.{exit}M.*.dollars`      | Waterfall   | `compute_waterfall` | Dollar cells are `IF`-driven by election cells |
| `waterfall.{exit}M.*.election`     | Waterfall   | `solve_nash_elections` | Election strings come from `IF(pref > convert, "preference", "convert")` cell formulas; never hardcoded |
| `sensitivity.*`                    | Sensitivity | `solve_round` × 9 + `compute_waterfall` × 9 | Each cell re-runs the full round solver; AD on E and D may retrigger or dormant per cell |
| `solver.{3x,4x}.*`                 | Return-Solver | `solve_required_pct` | Solver inversion of base-case waterfall |

### 14.4 Output tabs of `cap_table.xlsx`

The API-facing input file has **no output tabs**. It is purely
Stakeholders / Securities / Notes / RoundTerms as listed in §14.1.
The five output tabs (Assumptions / Pro-Forma / Waterfall / Sensitivity
/ Return-Solver) live in `ground_truth.xlsx` and in the API's
**produced** workbook — not in `cap_table.xlsx`.

### 14.5 Defined-name conventions for Excel

Excel defined names = §12 keys with `.` → `_`. Examples:

- `ownership_founders_pct`
- `antidilution_series_e_new_ratio`
- `notes_note_a_conversion_branch`
- `waterfall_2000M_series_f_election`
- `tender_shares_transferred`
- `sensitivity_esop_13_pre_1400M_our_pct`
- `round_f_pps_usd`

Defined names are workbook-scoped, not sheet-scoped. The reconciler
reads workbook-level names exclusively.

---

## 15. Things this document deliberately does NOT decide

The following are decisions for downstream agents, NOT input constants:

- The exact Excel tab layout, cell positions, or scratch-tab structure.
  (Excel-builder's call.)
- The Python solver algorithm — direct iteration vs closed-form vs
  Newton's method. (Python-builder's call.)
- The harness perturbation mechanism — file swap vs in-process input
  override vs both. (Harness agent's call.)
- The agent construction prompt's wording, output-tab
  specification language, or formula-only constraint phrasing. (Prompt
  agent's call. May reference v3's `prompt.md` for tone.)
- Whether the convertible notes' interest accrues. (This document says
  zero; spec.md may elaborate but cannot change the input value
  without invalidating cross-builder agreement.)

---

## 16. Sanity-check expected outputs (informational; for unit asserts)

These are NOT truth-key targets — they are approximate solver outputs
the user provided in the v4 design brief. Builders should `assert` that
their final solver outputs are within these ballparks (±5%) as a
correctness gate; the actual truth values come from full-precision
Decimal arithmetic.

| Expected output | Approximate value | Notes |
|---|---|---|
| F_PPS                                 | $37–$38                 | down from E_PPS ≈ $49.16 |
| AD on E: triggers (NCP_E)             | ≈ E_PPS × 0.97 ≈ $47.7  | r_E ≈ 1.03 |
| AD on D: triggers (NCP_D in v4)       | ≈ $58.12 × 0.95 ≈ $55   | r_D(v4) ≈ 1.05; compound ratio ≈ 1.087 |
| Note A: CAP CONVERTS                  | ≈ $27.32                | ≈ 183K shares |
| Note B: DISCOUNT CONVERTS             | ≈ $32.5 (= F_PPS × 0.85)| ≈ 92K shares |
| Post-F FD T_F                         | ≈ 38.3M shares          | up from 36.6M post-E |
| Tender shares                         | ≈ 470K                  | $15M / (0.85 × $37.5) |
| Tracked keys                          | 167                     | per §12.9 |
| Tracked elections                     | 35 strings              | 7 classes × 5 exits |

---

## 17. Versioning and change discipline

This file IS frozen for v4. If a downstream builder finds a real
inconsistency (e.g., an arithmetic infeasibility with the stated
constants), they must STOP and surface it to the human, not silently
"fix" by adjusting an input. Permitted changes:

- Adding a clarifying comment in any §
- Tightening a tolerance in §13
- Adding (not removing) a sanity check in §16

Forbidden changes without re-cutting v4:

- Changing any locked numerical input (§3, §4, §5, §6, §7, §8, §9)
- Changing the truth-key list or count (§12)
- Changing the audit-probe perturbation (§11)

---

End of CANONICAL_INPUTS.md.
