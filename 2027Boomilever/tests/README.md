# Boomilever scoresheet tests

Targeted regression tests for `../scoresheet_b.xlsx` and `../scoresheet_c.xlsx`.

Shared machinery lives in [`../../test-utils`](../../test-utils) — see its
README for how the harness works and how to wire up another event. This file
holds only the Boomilever column maps and scenarios.

## Two divisions, one suite

Boomilever B and C score identically except for the Load Score Bonus on
`Constants!B5` — **5000 in Division B, 7500 in Division C** — so the event
keeps one workbook per division and this suite runs both, building its
expectations from that one number. A run reports each division separately and
exits non-zero if either fails.

That is the only difference between the two workbooks; `Info.!B4` (Division)
is the only other cell that differs.

## Running

From the repo root:

```sh
make                                      # every event's suite
make test-2027Boomilever                  # both divisions
make test-2027Boomilever ONLY=tiebreaks   # one scenario, both divisions
make test-2027Boomilever SEED=42          # replay with a given row placement
make test-2027Boomilever SHEET=2027Boomilever/scoresheet_b.xlsx   # one workbook
```

A workbook named on the command line replaces the pair, and the division is
inferred from its filename (`…_b.xlsx` / `…_c.xlsx`) to pick the expectations.
A name ending in neither is an error rather than a guess — otherwise a Division
B sheet would be silently checked against Division C's bonus.

Teams are placed on randomly chosen rows in 8–507, which differ from run to
run so formulas are checked across the whole fill-down range. Each run prints
its seed; a failure that only appears under some seeds is a fill-down defect.

Runs from any working directory. Requires Python 3 (stdlib only) and `soffice`
on `PATH`.
