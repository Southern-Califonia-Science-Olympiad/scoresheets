# Elastic Launched Glider scoresheet tests

Targeted regression tests for `../scoresheet_b.xlsx`.

Shared machinery lives in [`../../test-utils`](../../test-utils) — see its
README for how the harness works and how to wire up another event. This file
holds only the Elastic Launched Glider column maps and scenarios.

## Running

From the repo root:

```sh
make test-2027ElasticLaunchedGlider                 # just this one
make test-2027ElasticLaunchedGlider ONLY=tiebreaks  # one scenario
make test-2027ElasticLaunchedGlider SEED=42         # replay with a given row placement
```

Teams are placed on randomly chosen rows in 8–507, which differ from run to
run so formulas are checked across the whole fill-down range. Each run prints
its seed; a failure that only appears under some seeds is a fill-down defect.
