# Hovercraft scoresheet tests

Targeted regression tests for `../scoresheet_bc.xlsx`.

Shared machinery lives in [`../../test-utils`](../../test-utils) — see its
README for how the harness works and how to wire up another event. This file
holds only the Hovercraft column maps and scenarios.

## Running

From the repo root:

```sh
make test-2027Hovercraft                 # just this one
make test-2027Hovercraft ONLY=tiebreaks  # one scenario
make test-2027Hovercraft SEED=42         # replay with a given row placement
```

Teams are placed on randomly chosen rows in 11–510, which differ from run to
run so formulas are checked across the whole fill-down range. Each run prints
its seed; a failure that only appears under some seeds is a fill-down defect.

Every team is scored against the Target Time in `D8`, which is outside the team
rows, so each scenario sets it through `Scenario(cells={"D8": ...})`.
