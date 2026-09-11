# Thermodynamics scoresheet tests

Targeted regression tests for `../Thermodynamics B_C - Scoresheet.xlsx`.

Shared machinery lives in [`../../test-utils`](../../test-utils) — see its
README for how the harness works and how to wire up another event. This file
holds only the Thermodynamics column maps and scenarios.

## Running

From the repo root:

```sh
make                                   # every event's suite
make test-Thermodynamics               # just this one
make test-Thermodynamics ONLY=tiebreaks
```
