"""Generic scenario runner for scoresheet regression tests.

Event-specific test files supply a SheetSpec (where inputs and outputs live)
and a list of Scenarios. Everything in this module is event-agnostic.

Typical use from `<Event>/tests/test_scoresheet.py`:

    import sys
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "test-utils"))

    from runner import Scenario, SheetSpec, main

    SPEC = SheetSpec(input_cols=..., out_cols=...)
    SCENARIOS = [...]

    if __name__ == "__main__":
        sys.exit(main(DEFAULT_SHEET, SPEC, SCENARIOS))
"""

import random
import sys
import tempfile
from pathlib import Path

from harness import Workbook


class SheetSpec:
    """Where a given event's scoresheet keeps its inputs and results.

    input_cols / out_cols map friendly names to column letters. first_row
    and last_row bound the team data range. Every input cell in that range is
    blanked before a scenario is written, and teams land on randomly chosen
    rows within it (see pick_rows).
    """

    def __init__(self, input_cols, out_cols,
                 first_row=8, last_row=507, tolerance=1e-6):
        self.input_cols = input_cols
        self.out_cols = out_cols
        self.first_row = first_row
        self.last_row = last_row
        self.tolerance = tolerance


class Scenario:
    """A self-contained mock tournament plus the assertions it must satisfy.

    teams   list of dicts keyed by SheetSpec.input_cols names; each needs a
            "school" key, which is what `expect` is keyed by.
    expect  {school: {output_name: expected}}
    extra   [(cell_ref, expected)] or [(label, cell_ref, expected)] for cells
            outside the per-team output columns. Teams don't sit on fixed
            rows, so a ref names its row by placeholder: "BM{MissingStart}"
            is column BM on that school's row, and "AE{unused}" is a data row
            with no team on it. Refs without a placeholder are absolute.

    `expected` is a literal value or a Check such as is_number / nonzero.
    """

    def __init__(self, name, why, teams, expect, extra=None):
        self.name = name
        self.why = why
        self.teams = teams
        self.expect = expect
        self.extra = extra or []


def pick_rows(spec, scenarios, seed):
    """One shared pool of randomly chosen data rows for the whole run.

    Formulas are filled down the data range, and a fill-down mistake shows up
    only on the rows it affects -- row 8 being right proves nothing. Placing
    teams at random rows exercises the formulas across the whole range.

    Every scenario draws from the same pool: a scenario with n teams uses the
    first n rows of it, so scenarios share rows as far as their sizes allow.
    The pool is sized from all scenarios, not just the selected ones, so a
    seed reproduces the same rows under --only. One spare row is kept for the
    "{unused}" placeholder.
    """
    need = max(len(s.teams) for s in scenarios) + 1
    span = range(spec.first_row, spec.last_row + 1)
    if need > len(span):
        sys.exit("scenarios need %d rows, data range has %d" % (need, len(span)))
    return random.Random(seed).sample(span, need)


def layout(pool, scenario):
    """Rows for this scenario's teams, plus the {placeholder} row map.

    Rows are kept ascending so teams sit in the order they are listed, which
    anything that breaks ties by row position relies on.
    """
    n = len(scenario.teams)
    rows = sorted(pool[:n])
    names = {t["school"]: r for t, r in zip(scenario.teams, rows)}
    names["unused"] = pool[n]
    return rows, names


def build(path, spec, scenario, rows):
    """Copy the workbook and write the scenario's teams into it."""
    wb = Workbook(path)
    wb.clear(spec.input_cols.values(), spec.first_row, spec.last_row)

    for i, (t, row) in enumerate(zip(scenario.teams, rows)):
        values = {}
        if "team_no" in spec.input_cols:
            values["team_no"] = i + 1
        if "team" in spec.input_cols:
            values["team"] = "Team %d" % (i + 1)
        values.update(t)
        for field, value in values.items():
            wb.set("%s%d" % (spec.input_cols[field], row), value)

    return wb


class Check:
    """An expectation that is a property rather than a single value.

    Use for invariants such as "never zero", where no specific value is
    intended. Shows as its description in failure output.
    """

    def __init__(self, description, predicate):
        self.description = description
        self.predicate = predicate

    def __call__(self, actual):
        return self.predicate(actual)

    def __str__(self):
        return self.description


# Error values come back as text (#DIV/0!), so "is a float" means "computed".
is_number = Check("a number", lambda a: isinstance(a, float))
nonzero = Check("a nonzero number", lambda a: isinstance(a, float) and a != 0)


def matches(expected, actual, tolerance):
    if isinstance(expected, Check):
        return expected(actual)
    if isinstance(expected, (int, float)) and not isinstance(expected, bool):
        return isinstance(actual, float) and abs(actual - expected) <= tolerance
    return str(actual) == str(expected)


def show(value):
    if value == "":
        return "(blank)"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def run(path, spec, scenarios, profile, only=None, seed=None):
    total = failed = 0
    problems = []

    selected = [s for s in scenarios if only is None or s.name == only]
    if not selected:
        sys.exit("no scenario named %r (have: %s)"
                 % (only, ", ".join(s.name for s in scenarios)))

    if seed is None:
        seed = random.randrange(10 ** 6)
    pool = pick_rows(spec, scenarios, seed)
    print("Seed: %d (rerun with --seed %d, or SEED=%d under make)"
          % (seed, seed, seed))
    print("Rows: %s" % ", ".join(str(r) for r in sorted(pool)))

    for sc in selected:
        print("\n\033[1m%s\033[0m -- %s" % (sc.name, sc.why))
        rows, names = layout(pool, sc)
        values = build(path, spec, sc, rows).recalc(profile)

        checks = []
        for t, row in zip(sc.teams, rows):
            for field, expected in sc.expect.get(t["school"], {}).items():
                ref = "%s%d" % (spec.out_cols[field], row)
                checks.append(("%s.%s" % (t["school"], field), ref, expected))
        for entry in sc.extra:
            if len(entry) == 3:
                label, ref, expected = entry
            else:
                ref, expected = entry
                label = ref
            checks.append((label, ref.format_map(names), expected))

        for label, ref, expected in checks:
            actual = values.get(ref, "")
            total += 1
            if matches(expected, actual, spec.tolerance):
                print("  \033[32mPASS\033[0m %-28s %s = %s"
                      % (label, ref, show(actual)))
                continue
            print("  \033[31mFAIL\033[0m %-28s %s  expected %s, got %s"
                  % (label, ref, show(expected), show(actual)))
            failed += 1
            problems.append((sc.name, label, ref, expected, actual))

    print("\n" + "=" * 72)
    print("%d checks, %d failed" % (total, failed))
    if problems:
        print("\nFailures:")
        for name, label, ref, expected, actual in problems:
            print("  %-22s %-28s %s  expected %s, got %s"
                  % (name, label, ref, show(expected), show(actual)))
    return 1 if failed else 0


def main(default_sheet, spec, scenarios, argv=None):
    args = list(sys.argv[1:] if argv is None else argv)

    only = None
    if "--only" in args:
        i = args.index("--only")
        only = args[i + 1]
        del args[i:i + 2]

    seed = None
    if "--seed" in args:
        i = args.index("--seed")
        seed = int(args[i + 1])
        del args[i:i + 2]

    path = Path(args[0]) if args else Path(default_sheet)
    if not path.exists():
        sys.exit("scoresheet not found: %s" % path)

    print("Testing: %s" % path)
    profile = tempfile.mkdtemp(prefix="scoresheet-lo-profile-")
    return run(path, spec, scenarios, profile, only, seed)
