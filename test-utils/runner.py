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

import sys
import tempfile
from pathlib import Path

from harness import Workbook


class SheetSpec:
    """Where a given event's scoresheet keeps its inputs and results.

    input_cols / out_cols map friendly names to column letters. Teams are
    written into consecutive rows starting at first_row; rows up to
    clear_through are blanked first so each run is deterministic.
    """

    def __init__(self, input_cols, out_cols,
                 first_row=8, clear_through=40, tolerance=1e-6):
        self.input_cols = input_cols
        self.out_cols = out_cols
        self.first_row = first_row
        self.clear_through = clear_through
        self.tolerance = tolerance


class Scenario:
    """A self-contained mock tournament plus the assertions it must satisfy.

    teams   list of dicts keyed by SheetSpec.input_cols names; each needs a
            "school" key, which is what `expect` is keyed by.
    expect  {school: {output_name: expected_value}}
    extra   [(absolute_cell_ref, expected_value)] for whole-row checks
    """

    def __init__(self, name, why, teams, expect, extra=None):
        self.name = name
        self.why = why
        self.teams = teams
        self.expect = expect
        self.extra = extra or []


def build(path, spec, scenario):
    """Copy the workbook and write the scenario's teams into it."""
    wb = Workbook(path)

    for row in range(spec.first_row, spec.clear_through + 1):
        for col in spec.input_cols.values():
            wb.set("%s%d" % (col, row), None)

    for i, t in enumerate(scenario.teams):
        row = spec.first_row + i
        values = {}
        if "team_no" in spec.input_cols:
            values["team_no"] = i + 1
        if "team" in spec.input_cols:
            values["team"] = "Team %d" % (i + 1)
        values.update(t)
        for field, value in values.items():
            wb.set("%s%d" % (spec.input_cols[field], row), value)

    return wb


def matches(expected, actual, tolerance):
    if isinstance(expected, (int, float)) and not isinstance(expected, bool):
        return isinstance(actual, float) and abs(actual - expected) <= tolerance
    return str(actual) == str(expected)


def show(value):
    if value == "":
        return "(blank)"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def run(path, spec, scenarios, profile, only=None):
    total = failed = 0
    problems = []

    selected = [s for s in scenarios if only is None or s.name == only]
    if not selected:
        sys.exit("no scenario named %r (have: %s)"
                 % (only, ", ".join(s.name for s in scenarios)))

    for sc in selected:
        print("\n\033[1m%s\033[0m -- %s" % (sc.name, sc.why))
        values = build(path, spec, sc).recalc(profile)

        checks = []
        for i, t in enumerate(sc.teams):
            row = spec.first_row + i
            for field, expected in sc.expect.get(t["school"], {}).items():
                ref = "%s%d" % (spec.out_cols[field], row)
                checks.append(("%s.%s" % (t["school"], field), ref, expected))
        for ref, expected in sc.extra:
            checks.append(("row %s empty" % ref[2:], ref, expected))

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

    path = Path(args[0]) if args else Path(default_sheet)
    if not path.exists():
        sys.exit("scoresheet not found: %s" % path)

    print("Testing: %s" % path)
    profile = tempfile.mkdtemp(prefix="scoresheet-lo-profile-")
    return run(path, spec, scenarios, profile, only)
