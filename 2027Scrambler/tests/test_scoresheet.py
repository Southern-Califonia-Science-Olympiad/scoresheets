#!/usr/bin/env python3
"""Targeted regression tests for the Scrambler B scoresheet.

Each scenario builds a self-contained mock tournament, recalculates the real
workbook with LibreOffice, and asserts on named output columns.

    python3 test_scoresheet.py [path/to/scoresheet.xlsx] [--only SCENARIO]

Expectations encode INTENDED behaviour. A failure means the sheet disagrees
with the rules as specified -- not that the test needs adjusting to match.

Rules under test (2027 B rules, section 5; low score wins):
  - Run Score = 100 + Distance Score + Time Score + Can Bonus + Run Penalties.
  - Distance Score = 2 x Vehicle Distance; 2500 for a Failed Run.
  - Time Score = Run Time; Run Time is 0 for a Failed Run.
  - Can Bonus = -0.5 x (110 - Inside Can Distance). A blank Inside Can
    Distance means no bonus; over 100 cm counts as 100; no bonus on a Failed
    Run.
  - Run Penalties: +150 competition, +300 construction, +50 non-modification
    (kit) -- on every run, Failed Runs included.
  - Final Score = better Run Score + 5000 if not impounded.
  - A missing run is a Failed Run (8.h.v). A blank Successful/Failed box is
    Successful when the run has any measurement, otherwise Failed.
  - No proper eyewear: participation only (P).
  - Ties (5.j): shorter Vehicle Distance of the scored run, lower Time Score of
    the scored run, then the same for the non-scored run. A Failed Run has no
    Vehicle Distance and loses that tiebreak. Two Failed Runs without
    violations stay tied (5.i).
  - A Successful run missing its Run Time or Vehicle Distance, or with an
    Inside Can Distance that is negative or not a number, makes the team ERR:
    not scored, ranked or listed. The input cell turns red; there is no
    Errors-column message.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "test-utils"))

from runner import Scenario, SheetSpec, main  # noqa: E402

HERE = Path(__file__).resolve().parent
DEFAULT_SHEET = HERE.parent / "scoresheet_b.xlsx"

# Input boxes, by the numbering printed in rows 2-4 of the Scoring sheet.
INPUT_COLS = {
    "team_no": "B", "school": "C", "team": "D",
    "impound": "E",      # 1. Impounded (T/F)
    "eyewear": "F",      # 2. Participants use proper eyewear (T/F)
    "kit": "G",          # 3. Conforms to kit modification requirements (T/F)
    "dq": "T",           # 16. Disqualify (T/F)
}
# Per run (boxes 4-9 and 10-15): const params, comp params, Successful/Failed,
# run time, vehicle distance, inside can distance.
RUN_FIELDS = ["const", "comp", "success_failed", "time", "vehicle_distance", "can_distance"]
RUN_COLS = {1: "H I J K L M", 2: "N O P Q R S"}
for _n, _cols in RUN_COLS.items():
    for _field, _col in zip(RUN_FIELDS, _cols.split()):
        INPUT_COLS["%s%d" % (_field, _n)] = _col

OUT_COLS = {
    "r1_err": "AM", "r1_ds": "AN", "r1_ts": "AO", "r1_can": "AP", "r1_pen": "AQ", "r1": "AR",
    "scored": "AU",
    "r2_err": "AV", "r2_ds": "AW", "r2_ts": "AX", "r2_can": "AY", "r2_pen": "AZ", "r2": "BA",
    # each run's own tiebreaks: vehicle distance, time score
    "r1_tb1": "AS", "r1_tb2": "AT",
    "r2_tb1": "BB", "r2_tb2": "BC",
    "status": "BD", "tier": "BE", "run_score": "BF", "final_pen": "BG",
    "final": "BH", "rank": "BI",
    # scored run VD, TS; non-scored run VD, TS
    "tb1": "BJ", "tb2": "BL", "tb3": "BN", "tb4": "BP",
    "rank_tb": "BR", "rank_diff": "BS",
    # Breakdown of the scored run: run #, DS, TS, Can Bonus, penalties (run +
    # impound). Final = 100 + DS + TS + bonus + penalties.
    "bd_run": "BV", "bd_ds": "BW", "bd_ts": "BX", "bd_bonus": "BY", "bd_pen": "BZ",
    "exp_score": "CA", "exp_tier": "CB", "exp_tiebreak": "CC", "exp_rank": "CD",
    "points": "CE",
}

LISTED = "CG8"          # teams on the final rankings list
LIST_SCHOOL = "CS%d"    # final rankings: school, list starting on row 9

SPEC = SheetSpec(INPUT_COLS, OUT_COLS, first_row=9, last_row=508)

# A Failed Run: 100 + 2500 + 0.
FAILED = 2600


def keyed(rows):
    """A Scenario's `expect`, keyed by output name: {school: dict(...)}.

    Every team in a scenario must check the same outputs, so a team can't pass
    by leaving one out; a mismatch raises when the suite loads.
    """
    fields = None
    for school, row in rows.items():
        if fields is None:
            fields = set(row)
        elif set(row) != fields:
            raise ValueError("%s checks %s, others check %s"
                             % (school, sorted(row), sorted(fields)))
    return rows


def run(n, **kw):
    """Inputs for run n. Boxes left out stay blank."""
    return {"%s%d" % (k, n): v for k, v in kw.items() if v is not None}


S, F = "Successful", "Failed"


def ok(n, time, vehicle_distance, **kw):
    """A Successful run with its Run Time and Vehicle Distance."""
    return run(n, success_failed=S, time=time, vehicle_distance=vehicle_distance, **kw)


def team(school, *runs, **kw):
    """A team that impounded and wore eyewear. Override with impound="F"
    etc., or impound=None to leave a box blank."""
    row = {"school": school, "impound": "T", "eyewear": "T"}
    for r in runs:
        row.update(r)
    row.update(kw)
    return {k: v for k, v in row.items() if v is not None}


def bare(school):
    """No boxes at all: a No-Show."""
    return {"school": school}


# --------------------------------------------------------------------------
# Scenarios
# --------------------------------------------------------------------------


RULES_EXAMPLE = Scenario(
    "rules_example",
    "Section 6 example (Final Score 133.87).",
    teams=[
        team("Example",
             run(1, success_failed=S, time=7.27, vehicle_distance=67.6),
             run(2, success_failed=S, time=8.67, vehicle_distance=27.6, can_distance=50)),
    ],
    expect=keyed({
        "Example": dict(
            r1_ds=135.2, r1_ts=7.27, r1_can=0, r1_pen=0, r1=242.47, r2_ds=55.2, r2_ts=8.67,
            r2_can=-30, r2_pen=0, r2=133.87, scored="->", run_score=133.87, final=133.87,
            exp_score=133.87, status="C",
        ),
    }),
)

FAILED_RUNS = Scenario(
    "failed_runs",
    "Failed Runs score DS 2500 and TS 0 with no Can Bonus; missing runs fail; "
    "S/F inference; double-failed teams stay tied.",
    teams=[
        team("BothFailed", run(1, success_failed=F), run(2, success_failed=F)),
        # Only check-in boxes: two Failed Runs, but competed.
        team("NoRuns"),
        # An explicit F wins over measurements typed in anyway.
        team("FailedWithData",
             run(1, success_failed=F, time=5, vehicle_distance=5, can_distance=50)),
        # No S/F but measurements: Successful.
        team("InferredSuccess", run(1, time=5, vehicle_distance=10)),
        # Double-failed with the same violation: tied on score, and a Failed
        # Run gives the tiebreaks nothing to split them with.
        team("FailedViolA", run(1, success_failed=F, comp="F"),
             run(2, success_failed=F, comp="F")),
        team("FailedViolB", run(1, success_failed=F, comp="F"),
             run(2, success_failed=F, comp="F")),
    ],
    expect=keyed({
        "BothFailed": dict(
            r1_ds=2500, r1_ts=0, r1_can=0, r1_pen=0, r1=FAILED, r2=FAILED, final=FAILED,
            exp_score=FAILED, tb1="", tb2=0, tb3="", tb4=0, rank=2, rank_tb=2,
        ),
        "NoRuns": dict(
            r1_ds=2500, r1_ts=0, r1_can=0, r1_pen=0, r1=FAILED, r2=FAILED, final=FAILED,
            exp_score=FAILED, tb1="", tb2=0, tb3="", tb4=0, rank=2, rank_tb=2,
        ),
        "FailedWithData": dict(
            r1_ds=2500, r1_ts=0, r1_can=0, r1_pen=0, r1=FAILED, r2=FAILED, final=FAILED,
            exp_score=FAILED, tb1="", tb2=0, tb3="", tb4=0, rank=2, rank_tb=2,
        ),
        "InferredSuccess": dict(
            r1_ds=20, r1_ts=5, r1_can=0, r1_pen=0, r1=125, r2=FAILED, final=125, exp_score=125,
            tb1=10, tb2=5, tb3="", tb4=0, rank=1, rank_tb=1,
        ),
        "FailedViolA": dict(
            r1_ds=2500, r1_ts=0, r1_can=0, r1_pen=150, r1=FAILED + 150, r2=FAILED + 150,
            final=FAILED + 150, exp_score=FAILED + 150, tb1="", tb2=0, tb3="", tb4=0, rank=5,
            rank_tb=5,
        ),
        "FailedViolB": dict(
            r1_ds=2500, r1_ts=0, r1_can=0, r1_pen=150, r1=FAILED + 150, r2=FAILED + 150,
            final=FAILED + 150, exp_score=FAILED + 150, tb1="", tb2=0, tb3="", tb4=0, rank=5,
            rank_tb=5,
        ),
    }),
)

CAN_BONUS = Scenario(
    "can_bonus",
    "Can Bonus = -0.5 x (110 - Inside Can Distance); a blank box means no "
    "bonus, over 100 cm counts as 100, and a Failed Run gets none.",
    teams=[
        team("NoCans", ok(1, time=5, vehicle_distance=10)),
        team("Can50", ok(1, time=5, vehicle_distance=10, can_distance=50)),
        team("Can0", ok(1, time=5, vehicle_distance=10, can_distance=0)),
        team("Can100", ok(1, time=5, vehicle_distance=10, can_distance=100)),
        team("CanOver", ok(1, time=5, vehicle_distance=10, can_distance=101)),
        team("Run2CanOver", ok(1, time=5, vehicle_distance=10),
             ok(2, time=5, vehicle_distance=10, can_distance=250)),
        team("CanDecimal", ok(1, time=5, vehicle_distance=10, can_distance=37.5)),
        team("CanOnFailed", run(1, success_failed=F, can_distance=50)),
        # Negative, but on a Failed Run it is never used: no error.
        team("BadCanOnFailed", run(1, success_failed=F, can_distance=-5)),
        team("Run2Can", ok(1, time=5, vehicle_distance=10),
             ok(2, time=5, vehicle_distance=10, can_distance=20)),
    ],
    expect=keyed({
        "NoCans": dict(
            r1_can=0, r1=125, r2_can=0, r2=FAILED, final=125, exp_score=125, status="C",
        ),
        "Can50": dict(
            r1_can=-30, r1=95, r2_can=0, r2=FAILED, final=95, exp_score=95, status="C",
        ),
        "Can0": dict(
            r1_can=-55, r1=70, r2_can=0, r2=FAILED, final=70, exp_score=70, status="C",
        ),
        "Can100": dict(
            r1_can=-5, r1=120, r2_can=0, r2=FAILED, final=120, exp_score=120, status="C",
        ),
        "CanOver": dict(
            r1_can=-5, r1=120, r2_can=0, r2=FAILED, final=120, exp_score=120, status="C",
        ),
        "Run2CanOver": dict(
            r1_can=0, r1=125, r2_can=-5, r2=120, final=120, exp_score=120, status="C",
        ),
        "CanDecimal": dict(
            r1_can=-36.25, r1=88.75, r2_can=0, r2=FAILED, final=88.75, exp_score=88.75,
            status="C",
        ),
        "CanOnFailed": dict(
            r1_can=0, r1=FAILED, r2_can=0, r2=FAILED, final=FAILED, exp_score=FAILED, status="C",
        ),
        "BadCanOnFailed": dict(
            r1_can=0, r1=FAILED, r2_can=0, r2=FAILED, final=FAILED, exp_score=FAILED, status="C",
        ),
        "Run2Can": dict(
            r1_can=0, r1=125, r2_can=-45, r2=80, final=80, exp_score=80, status="C",
        ),
    }),
)

PENALTIES = Scenario(
    "penalties",
    "Run penalties on each run (Failed Runs too), impound on the Final Score, "
    "no eyewear participation only.",
    teams=[
        team("Clean", ok(1, time=5, vehicle_distance=10)),
        team("Comp", ok(1, time=5, vehicle_distance=10, comp="F")),
        team("Const", ok(1, time=5, vehicle_distance=10, const="F")),
        team("CompConst", ok(1, time=5, vehicle_distance=10, const="F", comp="F")),
        team("Kit", ok(1, time=5, vehicle_distance=10), kit="F"),
        team("NoImpound", ok(1, time=5, vehicle_distance=10), impound="F"),
        team("FailedComp", run(1, success_failed=F, comp="F")),
        team("Run2Comp", ok(1, time=5, vehicle_distance=10),
             ok(2, time=5, vehicle_distance=10, comp="F")),
        team("NoEyewear", ok(1, time=5, vehicle_distance=10), eyewear="F"),
    ],
    expect=keyed({
        "Clean": dict(
            r1_pen=0, r1=125, r2_pen=0, r2=FAILED, scored="<-", final_pen=0, final=125,
            exp_score=125, status="C",
        ),
        "Comp": dict(
            r1_pen=150, r1=275, r2_pen=0, r2=FAILED, scored="<-", final_pen=0, final=275,
            exp_score=275, status="C",
        ),
        "Const": dict(
            r1_pen=300, r1=425, r2_pen=0, r2=FAILED, scored="<-", final_pen=0, final=425,
            exp_score=425, status="C",
        ),
        "CompConst": dict(
            r1_pen=450, r1=575, r2_pen=0, r2=FAILED, scored="<-", final_pen=0, final=575,
            exp_score=575, status="C",
        ),
        "Kit": dict(
            r1_pen=50, r1=175, r2_pen=50, r2=FAILED + 50, scored="<-", final_pen=0, final=175,
            exp_score=175, status="C",
        ),
        "NoImpound": dict(
            r1_pen=0, r1=125, r2_pen=0, r2=FAILED, scored="<-", final_pen=5000, final=5125,
            exp_score=5125, status="C",
        ),
        "FailedComp": dict(
            r1_pen=150, r1=FAILED + 150, r2_pen=0, r2=FAILED, scored="->", final_pen=0,
            final=FAILED, exp_score=FAILED, status="C",
        ),
        "Run2Comp": dict(
            r1_pen=0, r1=125, r2_pen=150, r2=275, scored="<-", final_pen=0, final=125,
            exp_score=125, status="C",
        ),
        "NoEyewear": dict(
            r1_pen="", r1="", r2_pen="", r2="", scored="", final_pen="", final="P", exp_score="P",
            status="P",
        ),
    }),
)

TIEBREAKS = Scenario(
    "tiebreaks",
    "Each tiebreak level deciding a tie (loser listed first), and an unbroken tie.",
    teams=[
        # 125: TB1, shorter Vehicle Distance of the scored run (5 vs 10).
        team("TB1Loser", ok(1, time=5, vehicle_distance=10)),
        team("TB1Winner", ok(1, time=15, vehicle_distance=5)),
        # 145: TB2 level, same scored run; TB3, the non-scored run is shorter.
        team("TB3Loser", ok(1, time=5, vehicle_distance=20), ok(2, time=5, vehicle_distance=30)),
        team("TB3Winner", ok(1, time=5, vehicle_distance=20), ok(2, time=5, vehicle_distance=25)),
        # 165: TB4, same scored run and non-scored VD; faster non-scored run.
        team("TB4Loser", ok(1, time=5, vehicle_distance=30), ok(2, time=9, vehicle_distance=40)),
        team("TB4Winner", ok(1, time=5, vehicle_distance=30), ok(2, time=7, vehicle_distance=40)),
        # 175: TB2, same VD; lower Time Score (5 + kit penalty vs 55).
        team("TB2Loser", ok(1, time=55, vehicle_distance=10)),
        team("TB2Winner", ok(1, time=5, vehicle_distance=10), kit="F"),
        # 195: TB3, a measured non-scored run beats a Failed one.
        team("TB3FailedLoser", ok(1, time=5, vehicle_distance=45)),
        team("TB3FailedWinner", ok(1, time=5, vehicle_distance=45),
             ok(2, time=5, vehicle_distance=100)),
        # 205: identical.
        team("TieA", ok(1, time=5, vehicle_distance=50)),
        team("TieB", ok(1, time=5, vehicle_distance=50)),
    ],
    expect=keyed({
        "TB1Loser": dict(
            final=125, exp_score=125, tb1=10, tb2=5, tb3="", tb4=0, rank=1, rank_tb=2,
            rank_diff=-1, points=2,
        ),
        "TB1Winner": dict(
            final=125, exp_score=125, tb1=5, tb2=15, tb3="", tb4=0, rank=1, rank_tb=1,
            rank_diff=0, points=1,
        ),
        "TB3Loser": dict(
            final=145, exp_score=145, tb1=20, tb2=5, tb3=30, tb4=5, rank=3, rank_tb=4,
            rank_diff=-1, points=4,
        ),
        "TB3Winner": dict(
            final=145, exp_score=145, tb1=20, tb2=5, tb3=25, tb4=5, rank=3, rank_tb=3,
            rank_diff=0, points=3,
        ),
        "TB4Loser": dict(
            final=165, exp_score=165, tb1=30, tb2=5, tb3=40, tb4=9, rank=5, rank_tb=6,
            rank_diff=-1, points=6,
        ),
        "TB4Winner": dict(
            final=165, exp_score=165, tb1=30, tb2=5, tb3=40, tb4=7, rank=5, rank_tb=5,
            rank_diff=0, points=5,
        ),
        "TB2Loser": dict(
            final=175, exp_score=175, tb1=10, tb2=55, tb3="", tb4=0, rank=7, rank_tb=8,
            rank_diff=-1, points=8,
        ),
        "TB2Winner": dict(
            final=175, exp_score=175, tb1=10, tb2=5, tb3="", tb4=0, rank=7, rank_tb=7,
            rank_diff=0, points=7,
        ),
        "TB3FailedLoser": dict(
            final=195, exp_score=195, tb1=45, tb2=5, tb3="", tb4=0, rank=9, rank_tb=10,
            rank_diff=-1, points=10,
        ),
        "TB3FailedWinner": dict(
            final=195, exp_score=195, tb1=45, tb2=5, tb3=100, tb4=5, rank=9, rank_tb=9,
            rank_diff=0, points=9,
        ),
        "TieA": dict(
            final=205, exp_score=205, tb1=50, tb2=5, tb3="", tb4=0, rank=11, rank_tb=11,
            rank_diff=0, points=11,
        ),
        "TieB": dict(
            final=205, exp_score=205, tb1=50, tb2=5, tb3="", tb4=0, rank=11, rank_tb=11,
            rank_diff=0, points=11,
        ),
    }),
)

SCORED_RUN = Scenario(
    "scored_run",
    "The better Run Score counts; equal runs fall back to each run's own "
    "tiebreaks: shorter Vehicle Distance (none for a Failed Run), then lower "
    "Time Score.",
    teams=[
        team("Run2Better", ok(1, time=5, vehicle_distance=15), ok(2, time=5, vehicle_distance=10)),
        team("Run1Better", ok(1, time=5, vehicle_distance=10), ok(2, time=5, vehicle_distance=15)),
        # 145 each: the shorter Vehicle Distance picks.
        team("EqualShorterVD1", ok(1, time=25, vehicle_distance=10),
             ok(2, time=5, vehicle_distance=20)),
        team("EqualShorterVD2", ok(1, time=5, vehicle_distance=20),
             ok(2, time=25, vehicle_distance=10)),
        # 125 each with the same VD: run 2 has the lower Time Score.
        team("EqualLowerTS", ok(1, time=35, vehicle_distance=10, can_distance=50),
             ok(2, time=5, vehicle_distance=10)),
        team("Identical", ok(1, time=5, vehicle_distance=10), ok(2, time=5, vehicle_distance=10)),
        team("FailThenOK", run(1, success_failed=F), ok(2, time=5, vehicle_distance=10)),
        # 2600 each: the Successful run has a Vehicle Distance, the Failed one doesn't.
        team("FailedEqualsOK", run(1, success_failed=F), ok(2, time=5, vehicle_distance=1247.5)),
        team("OKEqualsFailed", ok(1, time=5, vehicle_distance=1247.5), run(2, success_failed=F)),
    ],
    expect=keyed({
        "Run2Better": dict(
            r1=135, r2=125, r1_tb1=15, r1_tb2=5, r2_tb1=10, r2_tb2=5, scored="->", run_score=125,
            exp_score=125, tb1=10, tb2=5, tb3=15, tb4=5,
        ),
        "Run1Better": dict(
            r1=125, r2=135, r1_tb1=10, r1_tb2=5, r2_tb1=15, r2_tb2=5, scored="<-", run_score=125,
            exp_score=125, tb1=10, tb2=5, tb3=15, tb4=5,
        ),
        "EqualShorterVD1": dict(
            r1=145, r2=145, r1_tb1=10, r1_tb2=25, r2_tb1=20, r2_tb2=5, scored="<-", run_score=145,
            exp_score=145, tb1=10, tb2=25, tb3=20, tb4=5,
        ),
        "EqualShorterVD2": dict(
            r1=145, r2=145, r1_tb1=20, r1_tb2=5, r2_tb1=10, r2_tb2=25, scored="->", run_score=145,
            exp_score=145, tb1=10, tb2=25, tb3=20, tb4=5,
        ),
        "EqualLowerTS": dict(
            r1=125, r2=125, r1_tb1=10, r1_tb2=35, r2_tb1=10, r2_tb2=5, scored="->", run_score=125,
            exp_score=125, tb1=10, tb2=5, tb3=10, tb4=35,
        ),
        "Identical": dict(
            r1=125, r2=125, r1_tb1=10, r1_tb2=5, r2_tb1=10, r2_tb2=5, scored="<-", run_score=125,
            exp_score=125, tb1=10, tb2=5, tb3=10, tb4=5,
        ),
        "FailThenOK": dict(
            r1=FAILED, r2=125, r1_tb1="", r1_tb2=0, r2_tb1=10, r2_tb2=5, scored="->",
            run_score=125, exp_score=125, tb1=10, tb2=5, tb3="", tb4=0,
        ),
        "FailedEqualsOK": dict(
            r1=FAILED, r2=FAILED, r1_tb1="", r1_tb2=0, r2_tb1=1247.5, r2_tb2=5, scored="->",
            run_score=FAILED, exp_score=FAILED, tb1=1247.5, tb2=5, tb3="", tb4=0,
        ),
        "OKEqualsFailed": dict(
            r1=FAILED, r2=FAILED, r1_tb1=1247.5, r1_tb2=5, r2_tb1="", r2_tb2=0, scored="<-",
            run_score=FAILED, exp_score=FAILED, tb1=1247.5, tb2=5, tb3="", tb4=0,
        ),
    }),
)

STATUSES = Scenario(
    "statuses",
    "C / DQ / NS / P / ERR: score, tier, rank, points and the final rankings list.",
    teams=[
        team("Competitor", ok(1, time=5, vehicle_distance=10)),
        team("Disqualified", ok(1, time=5, vehicle_distance=10), dq="T"),
        bare("NoShow"),
        team("NoEyewear", eyewear="F"),
        team("Errored", run(1, success_failed=S)),
    ],
    expect=keyed({
        "Competitor":   dict(status="C", exp_score=125, tier=1, exp_tier=1, exp_rank=1, points=1),
        "Disqualified": dict(
            status="DQ", exp_score="DQ", tier="DQ", exp_tier="DQ", exp_rank="DQ", points=7,
        ),
        "NoShow": dict(
            status="NS", exp_score="NS", tier="NS", exp_tier="NS", exp_rank="NS", points=6,
        ),
        "NoEyewear": dict(
            status="P", exp_score="P", tier="P", exp_tier="P", exp_rank="P", points=5,
        ),
        "Errored": dict(
            status="ERR", exp_score="ERR", tier="ERR", exp_tier="ERR", exp_rank="ERR",
            points="ERR",
        ),
    }),
    extra=[
        ("blank row status", "BD{unused}", ""),
        ("blank row score", "CA{unused}", ""),
        ("listed teams", LISTED, 4),
        ("1st", LIST_SCHOOL % 9, "Competitor"),
        ("2nd", LIST_SCHOOL % 10, "NoEyewear"),
        ("3rd", LIST_SCHOOL % 11, "NoShow"),
        ("4th", LIST_SCHOOL % 12, "Disqualified"),
        ("no 5th (ERR unlisted)", LIST_SCHOOL % 13, ""),
    ],
)

INPUT_ERRORS = Scenario(
    "input_errors",
    "A Successful run missing a measurement, or with a negative or text "
    "Inside Can Distance, is ERR; errored teams don't take rank slots.",
    teams=[
        team("SOnly", run(1, success_failed=S)),
        team("MissingTime", run(1, success_failed=S, vehicle_distance=10)),
        team("MissingVD", run(1, success_failed=S, time=5)),
        team("InferredPartial", run(1, time=5)),
        # A can distance alone makes the run Successful, with nothing measured.
        team("InferredCanOnly", run(1, can_distance=50)),
        team("Negative", ok(1, time=5, vehicle_distance=-5)),
        team("CanNegative", ok(1, time=5, vehicle_distance=10, can_distance=-1)),
        team("CanText", ok(1, time=5, vehicle_distance=10, can_distance="x")),
        team("Run2Only", ok(1, time=5, vehicle_distance=10),
             ok(2, time=5, vehicle_distance=10, can_distance=-3)),
        team("FailedEmpty", run(1, success_failed=F)),
        team("Clean", ok(1, time=5, vehicle_distance=10)),
        team("Clean2", ok(1, time=5, vehicle_distance=15)),
    ],
    expect=keyed({
        "SOnly": dict(
            r1_err=True, r2_err=False, status="ERR", exp_score="ERR", exp_rank="ERR",
            points="ERR",
        ),
        "MissingTime": dict(
            r1_err=True, r2_err=False, status="ERR", exp_score="ERR", exp_rank="ERR",
            points="ERR",
        ),
        "MissingVD": dict(
            r1_err=True, r2_err=False, status="ERR", exp_score="ERR", exp_rank="ERR",
            points="ERR",
        ),
        "InferredPartial": dict(
            r1_err=True, r2_err=False, status="ERR", exp_score="ERR", exp_rank="ERR",
            points="ERR",
        ),
        "InferredCanOnly": dict(
            r1_err=True, r2_err=False, status="ERR", exp_score="ERR", exp_rank="ERR",
            points="ERR",
        ),
        "Negative": dict(
            r1_err=True, r2_err=False, status="ERR", exp_score="ERR", exp_rank="ERR",
            points="ERR",
        ),
        "CanNegative": dict(
            r1_err=True, r2_err=False, status="ERR", exp_score="ERR", exp_rank="ERR",
            points="ERR",
        ),
        "CanText": dict(
            r1_err=True, r2_err=False, status="ERR", exp_score="ERR", exp_rank="ERR",
            points="ERR",
        ),
        "Run2Only": dict(
            r1_err=False, r2_err=True, status="ERR", exp_score="ERR", exp_rank="ERR",
            points="ERR",
        ),
        "FailedEmpty": dict(
            r1_err=False, r2_err=False, status="C", exp_score=FAILED, exp_rank=3, points=3,
        ),
        "Clean": dict(r1_err=False, r2_err=False, status="C", exp_score=125, exp_rank=1, points=1),
        "Clean2": dict(
            r1_err=False, r2_err=False, status="C", exp_score=135, exp_rank=2, points=2,
        ),
    }),
    extra=[
        ("listed teams", LISTED, 3),
        ("1st", LIST_SCHOOL % 9, "Clean"),
        ("2nd", LIST_SCHOOL % 10, "Clean2"),
        ("3rd", LIST_SCHOOL % 11, "FailedEmpty"),
        ("no 4th", LIST_SCHOOL % 12, ""),
    ],
)

BREAKDOWN = Scenario(
    "breakdown",
    "Breakdown shows the scored run's DS, TS, Can Bonus and penalties (incl. "
    "impound); blank when not competing.",
    teams=[
        team("Example",
             run(1, success_failed=S, time=7.27, vehicle_distance=67.6),
             run(2, success_failed=S, time=8.67, vehicle_distance=27.6, can_distance=50)),
        team("Penalised", ok(1, time=5, vehicle_distance=10, comp="F"),
             ok(2, time=5, vehicle_distance=15), kit="F", impound="F"),
        team("DoubleFailed", run(1, success_failed=F, const="F"), run(2, success_failed=F)),
        team("Disqualified", ok(1, time=5, vehicle_distance=10), dq="T"),
        team("Errored", run(1, success_failed=S)),
    ],
    expect=keyed({
        "Example": dict(
            bd_run=2, bd_ds=55.2, bd_ts=8.67, bd_bonus=-30, bd_pen=0, final=133.87,
            exp_score=133.87,
        ),
        # run 1: 125 + 150 + 50 = 325; run 2: 135 + 50 = 185 -> run 2, + 5000
        "Penalised": dict(
            bd_run=2, bd_ds=30, bd_ts=5, bd_bonus=0, bd_pen=5050, final=5185, exp_score=5185,
        ),
        # run 2 (2600) beats run 1 (2900)
        "DoubleFailed": dict(
            bd_run=2, bd_ds=2500, bd_ts=0, bd_bonus=0, bd_pen=0, final=FAILED, exp_score=FAILED,
        ),
        "Disqualified": dict(
            bd_run="", bd_ds="", bd_ts="", bd_bonus="", bd_pen="", final="DQ", exp_score="DQ",
        ),
        "Errored": dict(
            bd_run="", bd_ds="", bd_ts="", bd_bonus="", bd_pen="", final="ERR", exp_score="ERR",
        ),
    }),
)

SCENARIOS = [RULES_EXAMPLE, FAILED_RUNS, CAN_BONUS, PENALTIES, TIEBREAKS, SCORED_RUN,
             STATUSES, INPUT_ERRORS, BREAKDOWN]

# The export columns must mirror their working columns: asserting a working
# value also asserts its export twin wherever a scenario checks the former.
_TWINS = {"tier": "exp_tier", "rank_tb": "exp_rank", "rank_diff": "exp_tiebreak"}
for _s in SCENARIOS:
    for _school, _exp in _s.expect.items():
        for _src, _dst in _TWINS.items():
            if _src in _exp and _dst not in _exp and _exp.get("status", "C") == "C":
                _exp[_dst] = _exp[_src]

if __name__ == "__main__":
    sys.exit(main(DEFAULT_SHEET, SPEC, SCENARIOS))
