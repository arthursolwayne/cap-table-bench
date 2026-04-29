# Kelvin Dynamics — Series F Analysis

A venture fund is evaluating a $10,000,000 participation in Kelvin Dynamics' proposed Series F. **Our $10M is one-third of the $30M primary**; F_LEAD writes the other $20M plus a $15M secondary tender at 0.85× F PPS from the founders. Build a formula-driven Excel workbook from the attached `cap_table.xlsx` — Carta-style tabs (Stakeholders, Securities, Notes, RoundTerms), the sole source of cap-table facts — with the five output tabs below.

## Required output tabs (name exactly)

1. **Assumptions** — every knob on labeled rows: Series F pre-money, primary raise, secondary raise, tender PPS multiplier, ESOP refresh target % (post-F FD), five exit values, our check size, six prior-round OCPs (Seed, A, B, C, D-as-post-E-NCP, E), option strike, both notes' principals / caps / discounts, broad-based WA AD inputs. Downstream tabs read only from here.

2. **Pro-Forma** — post-F cap table by stakeholder and security class. Include (a) **AD block** for Seed / A / B / C / D / E: trigger test (F PPS vs. OCP), new conversion price, new ratio (D's reported ratio is the **compound** v3·v4 adjustment off D's original $60 OCP), as-converted share count; (b) **notes block** for Note A and Note B: cap-implied PPS, discount-implied PPS, conversion PPS = min, conversion-branch label ("cap" | "discount"), shares issued; (c) **tender block**: tender shares = $15M / (0.85·F_PPS), founders' pro-rata share-out, F_LEAD secondary holding. Ownership sums to 100.000000%.

3. **Waterfall** — five exits ($700M, $1.2B, $2B, $4B, $8B) as columns. Rows in seniority: F, E, D, C, B, A, Seed, notes (combined Note A + Note B pro rata), common + RSAs, issued options (net of strike if ITM). **For each of the seven priced classes (Seed, A, B, C, D, E, F), at each of the five exits, show the class's elected treatment ("preference" or "convert") as a separate cell**, chosen to maximize that class's proceeds given the other classes' Nash-optimal choices. **35 election cells total.** Each column reconciles to its exit within $1.

4. **Sensitivity** — 3×3 grid: ESOP refresh {11%, 13%, 15%} × Series F pre-money {$1.2B, $1.4B, $1.8B}. Each cell reports our $10M check's post-F ownership % and $ proceeds at a **$2,000M exit**. Each cell re-solves the full round fixed point — F PPS, ESOP top-up, AD on E and D, BOTH note conversion branches — from scratch; at $1.8B pre, Note A's branch may flip from cap to discount, so per-cell branch evaluation is required.

5. **Return-Solver** — for each exit, solve for the Series F ownership % our $10M check would need to return 3× ($30M) and 4× ($40M), under base-case ESOP and the full waterfall (including F's own election). Flag infeasible targets (≤ 0% or > 100%).

## Hard requirement

Every output-tab cell is a formula referencing Assumptions or another formula cell. No hardcoded numeric outputs outside Assumptions. **No election string is hardcoded** — every election cell is an `IF`-formula output comparing aggregate preference to convert-pro-rata. **No conversion-branch string is hardcoded** — both notes' branch labels are `IF`-formula outputs comparing cap-implied PPS to discount-implied PPS.

## Modeling choices to get right

- **1× non-participating with conversion election on every priced class** (Seed, A, B, C, D, E, F — seven classes). Per exit, each class elects max of aggregate preference vs. convert. The 35 elections form a simultaneous best-response under reverse-chronological seniority (F > E > D > C > B > A > Seed). Notes A and B have NO election — common-equivalent post-conversion, always ride the convert pool.

- **Broad-based WA anti-dilution on every prior class.** Test `F_PPS < OCP` for each. Seed/A/B/C dormant. **D and E both trigger** (F_PPS ≈ $37.5 < D's post-E NCP $58.118347 and < E's $49.158311 — two separate broad-based WA solves). Apply `NCP = OCP × (A + B) ÷ (A + C)`: `A` = pre-F broad-based FD = 36,617,961; `B` = $30M / OCP; `C` = new F primary = $30M / F_PPS. Notes EXCLUDED from B and C. Series D's reported new_ratio is the **compound** v3·v4 ratio; D's effective OCP is $58.118347, NOT $60.

- **Convertible notes A and B convert at F closing.** Per note, conversion PPS = `min(cap-implied PPS, F_PPS × (1 − discount))`, where cap-implied PPS = cap / pre-money-FD-at-conversion (FD includes ESOP top-up and the OTHER note's converted shares; excludes new F primary). **Compute both branches independently per note and select the lower via `IF`.** Note A: $5M, $1.0B cap, 20% discount. Note B: $3M, $1.6B cap, 15% discount. Zero accrued interest.

- **Pre-money ESOP top-up to 13.0% of post-F FD**, borne by pre-F holders. Pre-F pool already 4,394,667 (2,500,000 issued + 1,894,667 unissued); top-up = `max(0, 0.13 × T_F − 4,394,667)`. Couples to fixed point and to AD on E and D.

- **Secondary tender ($15M @ 0.85·F_PPS)** transfers existing common from founders to F_LEAD. Tender shares = $15M / (0.85 × F_PPS). **No new shares issued** — T_F unchanged. Founders reduce pro rata across F1/F2/F3 (50% / 31.25% / 18.75%). F_LEAD post-F = primary F preferred ($20M / F_PPS) + secondary common.

- **Our $10M check** is Series F primary only — $10M / F_PPS — no secondary, no side vehicle. Our fund's % = 1/3 of `series_f_new`.

- **Issued options** exercise at strike, contribute net if ITM; else forfeit. **Unissued options** forfeit on acquisition.

Derive every number from first principles using only `cap_table.xlsx` and the conventions above. Produce the workbook, not a narrative walkthrough.
