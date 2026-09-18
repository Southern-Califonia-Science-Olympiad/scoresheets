# Thermodynamics scoresheet tests

Targeted regression tests for `../scoresheet_bc.xlsx`.

Shared machinery lives in [`../../test-utils`](../../test-utils) — see its
README for how the harness works and how to wire up another event. This file
holds only the Thermodynamics column maps and scenarios.

## Running

From the repo root:

```sh
make                                          # every event's suite
make test-2027Thermodynamics                 # just this one
make test-2027Thermodynamics ONLY=tiebreaks  # one scenario
make test-2027Thermodynamics SEED=42         # replay with a given row placement
```

Teams are placed on randomly chosen rows in 8–507, which differ from run to
run so formulas are checked across the whole fill-down range. Each run prints
its seed; a failure that only appears under some seeds is a fill-down defect.

Runs from any working directory. Exit status is non-zero if any check fails.
Requires Python 3 (stdlib only) and `soffice` on `PATH`.
