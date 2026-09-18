# test-utils

Shared machinery for scoresheet regression tests. Event-agnostic — each event
keeps its own cases in `<Event>/tests/`.

| File | Purpose |
|---|---|
| `harness.py` | xlsx mechanics: write input cells, strip cached values, recalculate via LibreOffice, read results back |
| `runner.py` | `SheetSpec`, `Scenario`, the run loop, comparison and reporting, and the CLI entry point |

Python 3 stdlib only, plus `soffice` on `PATH`. The runner uses its own
LibreOffice profile directory, so it is safe to run while a sheet is open.

The root `Makefile` discovers suites as `*/tests/test_scoresheet.py`, so a new
event folder following the layout below is picked up by `make` automatically.

## Why LibreOffice

The scoring logic lives in formulas, so testing it means recalculating with real
inputs. Driving LibreOffice headless puts the workbook's own formulas under test
rather than a reimplementation of them. A full recalculation takes about a
second.

Python formula-evaluation libraries (`pycel`, `formulas`) were not used: these
sheets rely on `SWITCH`, `COUNTIFS` with concatenated criteria, `RANK` and
`IFERROR`, and partial function coverage produces false failures on a correct
sheet. `openpyxl` alone cannot help either — with `data_only=True` it reads
*cached* values, so it cannot evaluate new inputs at all.

## Wiring up a new event

Create `<Event>/tests/test_scoresheet.py`:

```python
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "test-utils"))

from runner import Scenario, SheetSpec, main

DEFAULT_SHEET = Path(__file__).resolve().parent.parent / "<Event> B_C - Scoresheet.xlsx"

INPUT_COLS = {"team_no": "B", "school": "C", "team": "D", ...}
OUT_COLS = {"score": "...", "rank": "...", "points": "...", ...}
SPEC = SheetSpec(INPUT_COLS, OUT_COLS, first_row=8, last_row=507)

SCENARIOS = [Scenario("name", "why", teams=[...], expect={...})]

if __name__ == "__main__":
    sys.exit(main(DEFAULT_SHEET, SPEC, SCENARIOS))
```

`test-utils` is not a Python package — the name has a hyphen — so the
`sys.path.insert` line is how the import resolves. It is anchored to `__file__`,
so the suite runs from any working directory.

`INPUT_COLS` may define `team_no` and `team`; the runner fills those in
automatically per row so cases only specify what matters. Every team dict needs
a `school` key, which is what `expect` is keyed by.

## Writing scenarios

**Cases are not independent.** Ranks come from `RANK`/`COUNTIFS` over the whole
data range, and scores are typically normalised against field maxima. Adding a
team changes everyone's result. So each `Scenario` is a self-contained mock
tournament, run as its own workbook, rather than a row in one shared sheet.

Pin the field maxima with an anchor team so the normalised components are
predictable, then vary only what the case is about.

Expectations encode **intended** behaviour. If a check fails, the sheet
disagrees with the rules as specified — fix the sheet, not the expectation.

**Teams land on random rows.** Each run picks a pool of rows from
`first_row`–`last_row` and places each scenario's teams on them, in ascending
order so teams keep the order they're listed in. A formula that is right in
row 8 but broken further down only fails when a team sits on a broken row, so
fixed rows at the top of the range would never see it. Scenarios share the
pool: one with *n* teams uses its first *n* rows. The run prints its seed;
replay it with `--seed N` (or `make SEED=N`). A failure that appears under one
seed and not another is a fill-down defect on the rows that seed chose.

Because rows vary, `extra` refs name rows by placeholder: `"BM{MissingStart}"`
is column `BM` on that school's row, and `"AE{unused}"` is a data row with no
team on it. A ref with no placeholder is absolute, for cells that don't move
with the teams (field maxima in row 6, a rankings list filled from the top).

Floats compare with a tolerance (`SheetSpec(tolerance=...)`, default `1e-6`).
An expected value of `""` matches a blank or absent cell. Error values come back
as their text, so `#DIV/0!` compares as a string.

For invariants where no single value is intended, pass a `Check` instead of a
literal. `runner` provides `is_number` (computed, not an error) and `nonzero`;
build others as `Check("description", predicate)`. To assert on a cell outside
the per-team output columns — a field maximum, say — add a labeled entry to
`extra`: `("max PE", "AL6", nonzero)`.

Every input cell in the data range is blanked before a scenario is written, so
leftover teams in the workbook don't leak into results.
