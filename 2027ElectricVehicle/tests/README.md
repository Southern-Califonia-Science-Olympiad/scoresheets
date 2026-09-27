# Electric Vehicle scoresheet tests

Targeted regression tests for `../scoresheet_c.xlsx`.

Shared machinery lives in [`../../test-utils`](../../test-utils) — see its
README for how the harness works and how to wire up another event. This file
holds only the Electric Vehicle column maps and scenarios.

## Running

From the repo root:

```sh
make test-2027ElectricVehicle                 # just this one
make test-2027ElectricVehicle ONLY=tiebreaks  # one scenario
make test-2027ElectricVehicle SEED=42         # replay with a given row placement
```

Teams are placed on randomly chosen rows in 10–509. Every team is scored
against the Target Time in `D8`, which sits outside the team rows, so each
scenario sets it through `cells`.
