"""Boomilever B/C scoresheet regression tests.

Both divisions share one scoring model and differ only in the Load Score
Bonus on Constants!B5 -- 5000 in Division B, 7500 in Division C -- so the
scenarios are built from that one number and run against both workbooks.

    python3 tests/test_scoresheet.py                 # both divisions
    python3 tests/test_scoresheet.py ../path.xlsx    # one workbook (division
                                                     # inferred from its name)
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "test-utils"))

from runner import Scenario, SheetSpec, main, table

EVENT = Path(__file__).resolve().parent.parent

# division -> (workbook, Load Score Bonus)
DIVISIONS = {
    "b": (EVENT / "scoresheet_b.xlsx", 5000),
    "c": (EVENT / "scoresheet_c.xlsx", 7500),
}

MAX_LOAD = 15000        # Constants!B6, the load at which a Bonus claim counts

INPUT_COLS = {
    "team_no": "B", "school": "C", "team": "D",
    "estimate": "E",    # 1. Team's Estimated Load Supported (g)
    "mass": "F",        # 2. Mass of Structure (g)
    "eye": "G",         # 3. Eye protection is worn (T/F)
    "params": "H",      # 4. All const and comp params met (T/F)
    "option": "I",      # 5. Option Met (None/Base/Bonus)
    "load": "J",        # 6. Load Supported (g)
    "dq": "K",          # 7. Disqualify (T/F)
}

OUT_COLS = {
    "flag": "O",            # Box 2 Error Check
    "status": "V",          # C / P / NS / DQ / ERR
    "load_scored": "AL",    # breakdown: load plus bonus, if earned
    "mass_out": "AM",       # breakdown: structure mass
    "score": "AN",          # export: efficiency (load scored per gram)
    "tier": "AO",
    "tiebreak": "AP",       # places gained/lost to the tiebreak cascade
    "rank": "AQ",
    "points": "AR",
}

SPEC = SheetSpec(INPUT_COLS, OUT_COLS, first_row=8, last_row=507)


def team(school, estimate=None, mass=None, eye="T", params="T",
         option="Base", load=None, dq="F"):
    """A team row. Defaults are a clean, competing, Tier 1 entry."""
    t = {"school": school, "eye": eye, "params": params, "dq": dq}
    if estimate is not None:
        t["estimate"] = estimate
    if mass is not None:
        t["mass"] = mass
    if option is not None:
        t["option"] = option
    if load is not None:
        t["load"] = load
    return t


def scenarios(bonus):
    """Every scenario, with the division's bonus folded into expectations."""

    # A Bonus claim only earns the bonus once the load reaches MAX_LOAD.
    bonus_scored = MAX_LOAD + bonus

    return [
        Scenario(
            "bonus",
            "the Bonus option adds the division's bonus, but only at max load",
            teams=[
                team("Earned", estimate=15000, mass=20, option="Bonus", load=MAX_LOAD),
                team("Base", estimate=15000, mass=20, option="Base", load=MAX_LOAD),
                team("Short", estimate=14000, mass=20, option="Bonus", load=14000),
            ],
            expect=table(("load_scored", "score", "rank"), {
                # Bonus earned: 15000 + bonus, over 20 g
                "Earned": (bonus_scored, bonus_scored / 20, 1),
                # Same load, no Bonus claimed
                "Base": (15000, 750, 2),
                # Bonus claimed but load short of max: no bonus
                "Short": (14000, 700, 3),
            }),
        ),

        Scenario(
            "load_cap",
            "load above the maximum is capped, and still earns the bonus",
            teams=[
                team("Over", estimate=20000, mass=20, option="Bonus", load=20000),
                team("Exact", estimate=15000, mass=20, option="Bonus", load=MAX_LOAD),
            ],
            expect=table(("load_scored", "score"), {
                "Over": (bonus_scored, bonus_scored / 20),
                "Exact": (bonus_scored, bonus_scored / 20),
            }),
        ),

        Scenario(
            "tiers",
            "every Tier 1 team outranks every Tier 2 team, whatever the score",
            teams=[
                # Tier 1, deliberately poor efficiency
                team("T1 low", estimate=1000, mass=20, load=1000),
                # Tier 2 (params not met), far better efficiency
                team("T2 high", estimate=15000, mass=10, params="F", load=MAX_LOAD),
            ],
            expect=table(("tier", "score", "rank"), {
                "T1 low": (1, 50, 1),
                "T2 high": (2, 1500, 2),
            }),
        ),

        Scenario(
            "tiebreaks",
            "equal scores break on estimate accuracy, then mass; a full tie holds",
            teams=[
                # All four score 600; Base option, so no division difference.
                team("Exact", estimate=6000, mass=10, load=6000),    # TB1 0
                team("Light", estimate=5000, mass=10, load=6000),    # TB1 1000, mass 10
                team("Heavy", estimate=11000, mass=20, load=12000),  # TB1 1000, mass 20
                team("Twin", estimate=5000, mass=10, load=6000),     # identical to Light
            ],
            # All four tie on score, so each starts at rank 1 and the
            # tiebreak column reports the places the cascade cost them.
            expect=table(("score", "tiebreak", "rank"), {
                "Exact": (600, 0, 1),
                # Light and Twin are identical on score, estimate and mass:
                # a genuine tie, so they share the rank...
                "Light": (600, -1, 2),
                "Twin": (600, -1, 2),
                # ...and the tie consumes both places.
                "Heavy": (600, -3, 4),
            }),
        ),

        Scenario(
            "statuses",
            "DQ, No-Show and Participation-only sit outside the ranking",
            teams=[
                team("Scored", estimate=6000, mass=10, load=6000),
                team("DQd", estimate=6000, mass=10, load=6000, dq="T"),
                # No inputs at all beyond the school name
                {"school": "NoShow"},
                # Load of 0 is participation only
                team("NoLoad", estimate=6000, mass=10, load=0),
            ],
            expect=table(("status", "score", "rank", "points"), {
                "Scored": ("C", 600, 1, 1),
                # points: DQ = teams + 2, NS = teams + 1, P = teams
                "DQd": ("DQ", "DQ", "DQ", 6),
                "NoShow": ("NS", "NS", "NS", 5),
                "NoLoad": ("P", "P", "P", 4),
            }),
        ),

        Scenario(
            "blank_mass_errors",
            "a blank Box 2 flags the row and sets ERR; a clean row does not",
            teams=[
                team("Clean", estimate=6000, mass=10, load=6000),
                team("NoMass", estimate=6000, mass=None, load=6000),
            ],
            expect=table(("flag", "status", "score", "rank"), {
                "Clean": (False, "C", 600, 1),
                "NoMass": (True, "ERR", "ERR", "ERR"),
            }),
        ),

        # The pair below is the error-containment check: same three teams,
        # once alone and once beside a row the sheet cannot score. Their
        # scores and ranks must be identical in both, and the errored team
        # must not consume a rank.
        Scenario(
            "field_clean",
            "control: three teams alone in the field",
            teams=[
                team("First", estimate=12000, mass=10, load=12000),
                team("Second", estimate=6000, mass=10, load=6000),
                team("Third", estimate=3000, mass=10, load=3000),
            ],
            expect=table(("score", "rank"), {
                "First": (1200, 1),
                "Second": (600, 2),
                "Third": (300, 3),
            }),
        ),

        Scenario(
            "field_with_error",
            "an errored row leaves every other team's score and rank untouched",
            teams=[
                team("First", estimate=12000, mass=10, load=12000),
                team("Broken", estimate=9000, mass=None, load=9000),
                team("Second", estimate=6000, mass=10, load=6000),
                team("Third", estimate=3000, mass=10, load=3000),
            ],
            expect=table(("score", "rank"), {
                "First": (1200, 1),
                # No rank slot consumed: the others keep 2 and 3.
                "Broken": ("ERR", "ERR"),
                "Second": (600, 2),
                "Third": (300, 3),
            }),
        ),
    ]


def division_of(path):
    """Infer which division a workbook is, from its filename."""
    stem = Path(path).stem.lower()
    for div in DIVISIONS:
        if stem.endswith("_" + div):
            return div
    sys.exit("cannot tell which division %s is; expected a name ending in %s"
             % (path, " or ".join("_" + d for d in DIVISIONS)))


def positional(args):
    """The non-flag arguments, skipping the values belonging to flags."""
    out, i = [], 0
    while i < len(args):
        if args[i] in ("--only", "--seed"):
            i += 2
        elif args[i].startswith("--"):
            i += 1
        else:
            out.append(args[i])
            i += 1
    return out


if __name__ == "__main__":
    args = sys.argv[1:]

    # A workbook given on the command line (make test SHEET=...) replaces the
    # pair, and its own division decides which expectations apply.
    given = positional(args)
    if len(given) > 1:
        sys.exit("expected at most one workbook, got: %s" % ", ".join(given))

    if given:
        override = given[0]
        div = division_of(override)
        sys.exit(main(override, SPEC, scenarios(DIVISIONS[div][1]), argv=args))

    status = 0
    for div, (sheet, bonus) in DIVISIONS.items():
        print("\n\033[1m===== Division %s (bonus %d) =====\033[0m" % (div.upper(), bonus))
        if not sheet.exists():
            sys.exit("scoresheet not found: %s" % sheet)
        status |= main(sheet, SPEC, scenarios(bonus), argv=args)
    sys.exit(status)
