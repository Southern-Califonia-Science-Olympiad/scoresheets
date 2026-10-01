# Roller Coaster scoresheet tests

Targeted regression tests for `../scoresheet_b.xlsx`.

Shared machinery lives in [`../../test-utils`](../../test-utils) — see its
README for how the harness works and how to wire up another event. This file
holds only the Roller Coaster column maps and scenarios.

## Running

From the repo root:

```sh
make test-2027RollerCoaster                 # just this one
make test-2027RollerCoaster ONLY=tiebreaks  # one scenario
make test-2027RollerCoaster SEED=42         # replay with a given row placement
```

Or directly:

```sh
python3 test_scoresheet.py                     # the sheet next door
python3 test_scoresheet.py path/to/other.xlsx  # a future iteration
python3 test_scoresheet.py --only tiebreaks    # one scenario
```

Teams are placed on randomly chosen rows in 9–508. The Target Time is a global
input in `D7`; every scenario sets it through `cells`. Nothing is normalised
against a field maximum, so there are no zero-maximum cases.

## Layout

Built from the Scrambler workbook (same band/style scheme), with Mission
Possible's check-in boxes and Target Time cell.

| Columns | Purpose |
|---|---|
| `E`–`I` | Boxes 1–5: impounded, safety met, device width, length, height |
| `J`–`O` | Run 1, boxes 6–11: const params, comp params, Finish Line crossed, run time, elevators, loop height |
| `P`–`U` | Run 2, boxes 12–17 |
| `V` | Box 18: disqualify |
| `X`–`AO` | Sanitised inputs (hidden) |
| `AP`–`AR` | Target Time, dimension error, Target Time message (hidden) |
| `AS`–`BK` | Run 1 block, scored-run arrow `BB`, run 2 block (hidden) |
| `BL`–`CA` | Status, tier, scored run, final, ranks and tiebreaks (hidden) |
| `CC` | Errors (visible; Target Time messages only) |
| `CD`–`CM` | Breakdown and export block |

## Scoring rules under test (2027 B rules, section 5)

Rules source: the web edition,
<https://www.soinc.org/sites/default/files/uploaded_files/Science_Olympiad_Div_B_Rules_2027_for_Web_Secured.pdf>.

| Rule | Where |
|---|---|
| Run Score = Size + Time + Elevator + Loop; high score wins; best run is the Final Score (5.a–c) | `AU`–`BA`, `BE`–`BK`, `BB` |
| Size = (60 − h) + (30 − w) + (80 − l), 0.1 cm (5.d) | `AU`, `BE` |
| Time = 5/s up to the Target Time − 5/s past it, to 2 × target; 0 if the ball did not cross the Finish Line (5.e, 4.k) | `AV`, `BF` |
| Elevator = 50 each, 0–2; Loop = 3 × whole cm (5.f–g) | `AW`/`AX`, `BG`/`BH` |
| Tier 1 clean, Tier 2 a construction or competition violation (box 6/7 False, or a device over 30 × 80 × 60), Tier 3 not impounded; tier ranks before score (5.h) | `AY`, `BI`, `BM`, `BQ` |
| Scored run: lower tier, higher score, higher Loop, Size, Time Score, then longer Run Time | `BB` |
| Ties (5.i): Loop, Size, Time Score, longest Run Time, of the scored run; unbroken ties stay tied | `BR`–`CA` |
| No safety, or no run started → `P`; DQ / NS / P points | `BL`, `CM` |
| A used run with a missing or unusable time (needed unless the ball missed the Finish Line), a bad Elevator count (not 0–2) or Loop height (negative or not a number), or missing device dimensions once competing → `ERR`; the cell turns red, no message | `AT`, `BD`, `AQ` |
| Missing or invalid Target Time → `ERR` for teams that would be scored, with one message per check in `AR5`/`AR6` | `AR`, `CC` |

Conditional formatting (red input cell on an `ERR` run, pink `F` boxes and
oversize dimensions) is not covered; the harness reads values, not styles.

## Scenarios

| Name | Covers |
|---|---|
| `run_scores` | Size, Time, Elevators, Loop (whole cm) and the legal limits; run 2 unused |
| `time_score` | Early, on target, fractional seconds, late, at and beyond 2 × target |
| `target_60`, `target_30` | Time Score follows the Target Time in `D7` |
| `finish_line` | Missed Finish Line: Time 0 but size/Elevator/Loop kept, run time kept, no time needed; blank box defaults to crossed |
| `tiers` | Each Tier 2 trigger, Tier 3 over Tier 2, tier ranking ahead of score |
| `best_run` | Run choice by tier, score, Loop, Time Score, Run Time; run 2 only; breakdown follows |
| `tiebreaks` | Each tiebreak level deciding a tie and an unbroken tie |
| `statuses` | `C` / `DQ` / `NS` / `P` (no safety, no runs) / `ERR`, points |
| `input_errors` | Each check on each run, dimensions, a clean row, no messages; dimensions ignored when not competing |
| `error_containment` | Errored teams not scored, ranked or counted; others' ranks intact |
| `target_missing`, `target_invalid_*`, `target_bounds` | Target Time messages; 30 and 60 valid |

## Status as of 2026-09-30

683 checks, all passing.

## Decisions and open points

- **Dimensions are entered once** (boxes 3–5), as on the checklist, so both
  runs share a Size Score. Rule 4.l measures height after *each* scorable run,
  and 4.e lets teams modify the device between runs, so the two runs could
  legitimately differ. The checklist has no second set of boxes.
- **An oversize device is Tier 2 automatically** (rule 3.b), in addition to
  box 6/12 being marked False. Size Score is not clamped, so it can go
  negative.
- **Loop height has no upper limit**, since the rules set none: any
  non-negative height scores 3 points per whole cm. An Elevator count other
  than 0, 1 or 2 is an input error. Both boxes are "only fill out if
  successful", so blank is 0.
- **Tier is per run**, as the rules word it; the team's tier is the scored
  run's (3 if not impounded). The tier column is visible, unlike the tierless
  Scrambler.
- A missed Finish Line with no time entered is not an error (time is only a
  last tiebreak); an unusable value there is.
