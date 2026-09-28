# Scrambler scoresheet tests

Targeted regression tests for `../scoresheet_b.xlsx`.

Shared machinery lives in [`../../test-utils`](../../test-utils) — see its
README for how the harness works and how to wire up another event. This file
holds only the Scrambler column maps and scenarios.

## Running

From the repo root:

```sh
make test-2027Scrambler                 # just this one
make test-2027Scrambler ONLY=tiebreaks  # one scenario
make test-2027Scrambler SEED=42         # replay with a given row placement
```

Teams are placed on randomly chosen rows in 9–508. Unlike Electric Vehicle there
is no global input: the Target Distance doesn't enter the score, and Time Score
is the Run Time itself.
