#!/usr/bin/env python3
"""Targeted regression tests for the Electric Vehicle C scoresheet.

Each scenario builds a self-contained mock tournament, recalculates the real
workbook with LibreOffice, and asserts on named output columns.

    python3 test_scoresheet.py [path/to/scoresheet.xlsx] [--only SCENARIO]

Expectations encode INTENDED behaviour. A failure means the sheet disagrees
with the rules as specified -- not that the test needs adjusting to match.

Rules under test (2027 C rules, section 5; low score wins):
  - Run Score = 100 + Distance Score + Time Score + Bonuses + Run Penalties.
  - Distance Score = 2 x Vehicle Distance + 1 x Bottle Distance; 2500 for a
    Failed Run. Bottle Distance is 400 when the bottle is not past the Target
    Point or the Moving Water Bottle Requirements were violated (box = F),
    and that run gets no Bottle Bonus.
  - Time Score = 0.5 x |Target Time - Run Time|; Run Time is 0 for a Failed
    Run.
  - Bottle Bonus -20 if the bottle is beyond the Bottle Line; Pusher Bonus
    (opening - 35.0) x 1.5, with no pusher measured as 35.0. Neither applies
    to a Failed Run.
  - Run Penalties: +150 competition, +300 construction, +50 non-modification
    (kit) -- on every run, Failed Runs included.
  - Final Score = better Run Score + 5000 if not impounded + Event Time Bonus
    = (Event Time Used - 480) / 30. Event Time Left is only entered at
    Nationals; blank means no bonus.
  - A missing run is a Failed Run (4.r.v). A blank Successful/Failed box is
    Successful when the run has any measurement, otherwise Failed.
  - Lithium/lead batteries: participation only (P).
  - Ties: lower Distance Score of the scored run, lower Time Score, shorter
    Vehicle Distance, lower Bottle Pusher measurement. The pusher tiebreak
    applies to every team, double-failed ones included.
  - A Successful run missing its Run Time, Vehicle or Bottle Distance, or no
    Target Time, makes the team ERR: not scored, ranked or listed. Only the
    Target Time has an Errors-column message; a missing measurement is shown
    by its red input cell.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "test-utils"))

from runner import Scenario, SheetSpec, main  # noqa: E402

HERE = Path(__file__).resolve().parent
DEFAULT_SHEET = HERE.parent / "scoresheet_c.xlsx"

TARGET_TIME = "D7"

# Input boxes, by the numbering printed in rows 2-4 of the Scoring sheet.
INPUT_COLS = {
    "team_no": "B", "school": "C", "team": "D",
    "impound": "E",      # 1. Impounded (T/F)
    "battery": "F",      # 2. No batteries containing lithium or lead (T/F)
    "kit": "G",          # 3. Conforms to kit modification requirements (T/F)
    "pusher": "H",       # 4. Bottle Pusher Width (cm)
    "left_min": "Y", "left_sec": "Z",   # 21. Event Time Left
    "dq": "AA",          # 22. Disqualify (T/F)
}
# Per run (boxes 5-12 and 13-20): const params, comp params,
# Successful/Failed, run time, vehicle distance, bottle past Target Point with
# the Moving Water Bottle Requirements met, bottle distance, bottle bonus.
RUN_FIELDS = ["const", "comp", "sf", "time", "vd", "past", "bd", "bonus"]
RUN_COLS = {1: "I J K L M N O P", 2: "Q R S T U V W X"}
for _n, _cols in RUN_COLS.items():
    for _field, _col in zip(RUN_FIELDS, _cols.split()):
        INPUT_COLS["%s%d" % (_field, _n)] = _col

OUT_COLS = {
    "r1_err": "BB", "r1_ds": "BC", "r1_ts": "BD", "r1_bottle": "BE",
    "r1_pusher": "BF", "r1_pen": "BG", "r1": "BH",
    "scored": "BL",
    "r2_err": "BM", "r2_ds": "BN", "r2_ts": "BO", "r2_bottle": "BP",
    "r2_pusher": "BQ", "r2_pen": "BR", "r2": "BS",
    # each run's own tiebreaks: DS, TS, vehicle distance (the pusher measurement
    # is the same on both runs, so it can't pick between them)
    "r1_tb1": "BI", "r1_tb2": "BJ", "r1_tb3": "BK",
    "r2_tb1": "BT", "r2_tb2": "BU", "r2_tb3": "BV",
    "status": "BX", "tier": "BY", "run_score": "BZ", "final_pen": "CA",
    "etb": "CB", "final": "CC", "rank": "CD",
    "tb1": "CE", "tb2": "CG", "tb3": "CI", "tb4": "CK",
    "rank_tb": "CM", "rank_diff": "CN",
    "errors": "CQ",
    # Breakdown of the scored run: run #, DS, TS, bonuses (bottle + pusher +
    # event time), penalties (run + impound). Final = 100 + DS + TS + B + P.
    "bd_run": "CR", "bd_ds": "CS", "bd_ts": "CT", "bd_bonus": "CU", "bd_pen": "CV",
    "exp_score": "CW", "exp_tier": "CX", "exp_tiebreak": "CY", "exp_rank": "CZ",
    "points": "DA",
}

LISTED = "DC8"          # teams on the final rankings list
LIST_SCHOOL = "DO%d"    # final rankings: school, list starting on row 9

SPEC = SheetSpec(INPUT_COLS, OUT_COLS, first_row=9, last_row=508)

ERR_TT = "ERROR: Set Target Time in cell D7"
ERR_TT_TEXT = "ERROR: Target Time in cell D7 must be a number"

TT14 = {TARGET_TIME: 14}

# A Failed Run against a 14 s Target Time: 100 + 2500 + 0.5 x 14.
FAILED = 2607


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


def ok(n, time=14, vd=10, bd=10, **kw):
    """A Successful run, by default on time with DS 30: Run Score 130."""
    return run(n, sf=S, time=time, vd=vd, bd=bd, **kw)


def team(school, *runs, **kw):
    """A team that impounded with legal batteries. Override with impound="F"
    etc., or impound=None to leave a box blank."""
    row = {"school": school, "impound": "T", "battery": "T"}
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
    "Section 6 example from the web rules (Final Score 109.79).",
    teams=[
        team("Example",
             run(1, sf=S, time=12.27, vd=17.6, bd=10.6, bonus="F"),
             run(2, sf=S, time=16.37, vd=22.5, bd=12.1, bonus="T"),
             pusher=16),
    ],
    expect=keyed({
        "Example": dict(
            r1_ds=45.8, r1_ts=0.865, r1_bottle=0, r1_pusher=-28.5, r1_pen=0, r1=118.165,
            r2_ds=57.1, r2_ts=1.185, r2_bottle=-20, r2_pusher=-28.5, r2_pen=0, r2=109.785,
            scored="->", run_score=109.785, final=109.785, status="C",
        ),
    }),
    cells=TT14,
)

FAILED_RUNS = Scenario(
    "failed_runs",
    "Failed Runs score DS 2500 and TS 0.5 x Target; no bonuses; missing runs "
    "fail; S/F inference; double-failed teams split by pusher measurement.",
    teams=[
        team("BothFailed", run(1, sf=F), run(2, sf=F), pusher=16),
        # Only check-in boxes: two Failed Runs, but competed.
        team("NoRuns"),
        # An explicit F wins over measurements typed in anyway.
        team("FailedWithData", run(1, sf=F, time=12, vd=5, bd=5, bonus="T")),
        # No S/F but measurements: Successful.
        team("InferredSuccess", run(1, time=14, vd=10, bd=10)),
        # Double-failed teams tie on DS/TS/VD, so the pusher measurement (TB4)
        # decides even though no Pusher Bonus applies: 10 < 16 < 30 < 35 (none).
        team("FailedPusher10", run(1, sf=F), run(2, sf=F), pusher=10),
        team("FailedPusher30", run(1, sf=F), run(2, sf=F), pusher=30),
        # The same with violations.
        team("FailedViolLoser", run(1, sf=F, comp="F"), run(2, sf=F, comp="F"), pusher=30),
        team("FailedViolWinner", run(1, sf=F, comp="F"), run(2, sf=F, comp="F"), pusher=10),
    ],
    expect=keyed({
        "BothFailed": dict(
            r1_ds=2500, r1_ts=7, r1_bottle=0, r1_pusher=0, r1_pen=0, r1=FAILED, r2=FAILED,
            final=FAILED, tb4=16, rank=2, rank_tb=3, errors="",
        ),
        "NoRuns": dict(
            r1_ds=2500, r1_ts=7, r1_bottle=0, r1_pusher=0, r1_pen=0, r1=FAILED, r2=FAILED,
            final=FAILED, tb4=35, rank=2, rank_tb=5, errors="",
        ),
        "FailedWithData": dict(
            r1_ds=2500, r1_ts=7, r1_bottle=0, r1_pusher=0, r1_pen=0, r1=FAILED, r2=FAILED,
            final=FAILED, tb4=35, rank=2, rank_tb=5, errors="",
        ),
        "InferredSuccess": dict(
            r1_ds=30, r1_ts=0, r1_bottle=0, r1_pusher=0, r1_pen=0, r1=130, r2=FAILED, final=130,
            tb4=35, rank=1, rank_tb=1, errors="",
        ),
        "FailedPusher10": dict(
            r1_ds=2500, r1_ts=7, r1_bottle=0, r1_pusher=0, r1_pen=0, r1=FAILED, r2=FAILED,
            final=FAILED, tb4=10, rank=2, rank_tb=2, errors="",
        ),
        "FailedPusher30": dict(
            r1_ds=2500, r1_ts=7, r1_bottle=0, r1_pusher=0, r1_pen=0, r1=FAILED, r2=FAILED,
            final=FAILED, tb4=30, rank=2, rank_tb=4, errors="",
        ),
        "FailedViolLoser": dict(
            r1_ds=2500, r1_ts=7, r1_bottle=0, r1_pusher=0, r1_pen=150, r1=FAILED + 150,
            r2=FAILED + 150, final=FAILED + 150, tb4=30, rank=7, rank_tb=8, errors="",
        ),
        "FailedViolWinner": dict(
            r1_ds=2500, r1_ts=7, r1_bottle=0, r1_pusher=0, r1_pen=150, r1=FAILED + 150,
            r2=FAILED + 150, final=FAILED + 150, tb4=10, rank=7, rank_tb=7, errors="",
        ),
    }),
    cells=TT14,
)

BOTTLE_PAST_TARGET = Scenario(
    "bottle_past_target",
    "Box 10/18 = F (bottle not past the Target Point, or Moving Water Bottle "
    "Requirements violated): Bottle Distance counts as 400, none need be "
    "entered, and no Bottle Bonus.",
    teams=[
        team("Past", ok(1, past="T")),
        team("PastBlank", ok(1)),
        team("NotPast", run(1, sf=S, time=14, vd=10, past="F")),
        team("NotPastWithBd", ok(1, past="F")),
        # No Bottle Bonus once the bottle failed box 10 (box 12 shows red).
        team("NotPastBonus", ok(1, past="F", bonus="T")),
        team("Run2NotPastBonus", ok(1, bd=390), ok(2, past="F", bonus="T")),
        team("Run2NotPast", run(1, sf=F), run(2, sf=S, time=14, vd=10, past="F")),
        team("PastMissingBd", run(1, sf=S, time=14, vd=10, past="T")),
    ],
    expect=keyed({
        "Past": dict(
            r1_ds=30, r1_bottle=0, r1=130, r2_ds=2500, r2_bottle=0, r2=FAILED, final=130,
            status="C", errors="",
        ),
        "PastBlank": dict(
            r1_ds=30, r1_bottle=0, r1=130, r2_ds=2500, r2_bottle=0, r2=FAILED, final=130,
            status="C", errors="",
        ),
        "NotPast": dict(
            r1_ds=420, r1_bottle=0, r1=520, r2_ds=2500, r2_bottle=0, r2=FAILED, final=520,
            status="C", errors="",
        ),
        "NotPastWithBd": dict(
            r1_ds=420, r1_bottle=0, r1=520, r2_ds=2500, r2_bottle=0, r2=FAILED, final=520,
            status="C", errors="",
        ),
        "NotPastBonus": dict(
            r1_ds=420, r1_bottle=0, r1=520, r2_ds=2500, r2_bottle=0, r2=FAILED, final=520,
            status="C", errors="",
        ),
        "Run2NotPastBonus": dict(
            r1_ds=410, r1_bottle=0, r1=510, r2_ds=420, r2_bottle=0, r2=520, final=510, status="C",
            errors="",
        ),
        "Run2NotPast": dict(
            r1_ds=2500, r1_bottle=0, r1=FAILED, r2_ds=420, r2_bottle=0, r2=520, final=520,
            status="C", errors="",
        ),
        "PastMissingBd": dict(
            r1_ds="", r1_bottle="", r1="", r2_ds="", r2_bottle="", r2="", final="ERR",
            status="ERR", errors="",
        ),
    }),
    cells=TT14,
)

PENALTIES = Scenario(
    "penalties",
    "Run penalties on each run (Failed Runs too), impound on the Final Score, "
    "lithium batteries participation only.",
    teams=[
        team("Clean", ok(1)),
        team("Comp", ok(1, comp="F")),
        team("Const", ok(1, const="F")),
        team("CompConst", ok(1, const="F", comp="F")),
        team("Kit", ok(1), kit="F"),
        team("NoImpound", ok(1), impound="F"),
        team("FailedComp", run(1, sf=F, comp="F")),
        team("Run2Comp", ok(1), ok(2, comp="F")),
        team("Lithium", ok(1), battery="F"),
    ],
    expect=keyed({
        "Clean": dict(
            r1_pen=0, r1=130, r2_pen=0, r2=FAILED, scored="<-", final_pen=0, final=130, status="C",
        ),
        "Comp": dict(
            r1_pen=150, r1=280, r2_pen=0, r2=FAILED, scored="<-", final_pen=0, final=280,
            status="C",
        ),
        "Const": dict(
            r1_pen=300, r1=430, r2_pen=0, r2=FAILED, scored="<-", final_pen=0, final=430,
            status="C",
        ),
        "CompConst": dict(
            r1_pen=450, r1=580, r2_pen=0, r2=FAILED, scored="<-", final_pen=0, final=580,
            status="C",
        ),
        "Kit": dict(
            r1_pen=50, r1=180, r2_pen=50, r2=FAILED + 50, scored="<-", final_pen=0, final=180,
            status="C",
        ),
        "NoImpound": dict(
            r1_pen=0, r1=130, r2_pen=0, r2=FAILED, scored="<-", final_pen=5000, final=5130,
            status="C",
        ),
        "FailedComp": dict(
            r1_pen=150, r1=FAILED + 150, r2_pen=0, r2=FAILED, scored="->", final_pen=0,
            final=FAILED, status="C",
        ),
        "Run2Comp": dict(
            r1_pen=0, r1=130, r2_pen=150, r2=280, scored="<-", final_pen=0, final=130, status="C",
        ),
        "Lithium": dict(
            r1_pen="", r1="", r2_pen="", r2="", scored="", final_pen="", final="P", status="P",
        ),
    }),
    cells=TT14,
)

EVENT_TIME = Scenario(
    "event_time",
    "Event Time Bonus = (used - 480) / 30 whenever Event Time Left is entered "
    "(Nationals only); blank means none; capped at 8 min.",
    teams=[
        team("TwoMinLeft", ok(1), left_min=2, left_sec=0),
        team("ThirtySecLeft", ok(1), left_sec=30),
        team("NoneLeft", ok(1)),
        team("OverEight", ok(1), left_min=9),
    ],
    expect=keyed({
        "TwoMinLeft":    dict(etb=-4, final=126, rank=2),
        "ThirtySecLeft": dict(etb=-1, final=129, rank=3),
        "NoneLeft":      dict(etb=0, final=130, rank=4),
        "OverEight":     dict(etb=-16, final=114, rank=1),
    }),
    cells=TT14,
)

TIEBREAKS = Scenario(
    "tiebreaks",
    "Each tiebreak level deciding a tie (loser listed first), and an unbroken tie.",
    teams=[
        # 130: TB1, lower Distance Score (28 vs 30).
        team("TB1Loser", ok(1, time=14, vd=10, bd=10)),
        team("TB1Winner", ok(1, time=18, vd=10, bd=8)),
        # 186: TB2, same DS 36, lower Time Score (0 + kit penalty vs 50).
        team("TB2Loser", ok(1, time=114, vd=10, bd=16)),
        team("TB2Winner", ok(1, time=14, vd=13, bd=10), kit="F"),
        # 140: TB3, same DS 40 and TS 0, shorter Vehicle Distance.
        team("TB3Loser", ok(1, vd=15, bd=10)),
        team("TB3Winner", ok(1, vd=10, bd=20)),
        # 150: TB4, same DS/TS/VD; pusher 15 (-30) + kit (+50) + bottle (-20)
        # nets 0, so the smaller opening decides.
        team("TB4Loser", ok(1, vd=10, bd=30)),
        team("TB4Winner", ok(1, vd=10, bd=30, bonus="T"), pusher=15, kit="F"),
        # 160: identical.
        team("TieA", ok(1, vd=10, bd=40)),
        team("TieB", ok(1, vd=10, bd=40)),
    ],
    expect=keyed({
        "TB1Loser": dict(
            final=130, tb1=30, tb2=0, tb3=10, tb4=35, rank=1, rank_tb=2, rank_diff=-1, exp_rank=2,
            exp_tiebreak=-1, points=2,
        ),
        "TB1Winner": dict(
            final=130, tb1=28, tb2=2, tb3=10, tb4=35, rank=1, rank_tb=1, rank_diff=0, exp_rank=1,
            exp_tiebreak=0, points=1,
        ),
        "TB2Loser": dict(
            final=186, tb1=36, tb2=50, tb3=10, tb4=35, rank=9, rank_tb=10, rank_diff=-1,
            exp_rank=10, exp_tiebreak=-1, points=10,
        ),
        "TB2Winner": dict(
            final=186, tb1=36, tb2=0, tb3=13, tb4=35, rank=9, rank_tb=9, rank_diff=0, exp_rank=9,
            exp_tiebreak=0, points=9,
        ),
        "TB3Loser": dict(
            final=140, tb1=40, tb2=0, tb3=15, tb4=35, rank=3, rank_tb=4, rank_diff=-1, exp_rank=4,
            exp_tiebreak=-1, points=4,
        ),
        "TB3Winner": dict(
            final=140, tb1=40, tb2=0, tb3=10, tb4=35, rank=3, rank_tb=3, rank_diff=0, exp_rank=3,
            exp_tiebreak=0, points=3,
        ),
        "TB4Loser": dict(
            final=150, tb1=50, tb2=0, tb3=10, tb4=35, rank=5, rank_tb=6, rank_diff=-1, exp_rank=6,
            exp_tiebreak=-1, points=6,
        ),
        "TB4Winner": dict(
            final=150, tb1=50, tb2=0, tb3=10, tb4=15, rank=5, rank_tb=5, rank_diff=0, exp_rank=5,
            exp_tiebreak=0, points=5,
        ),
        "TieA": dict(
            final=160, tb1=60, tb2=0, tb3=10, tb4=35, rank=7, rank_tb=7, rank_diff=0,
            exp_rank=7, exp_tiebreak=0, points=7,
        ),
        "TieB": dict(
            final=160, tb1=60, tb2=0, tb3=10, tb4=35, rank=7, rank_tb=7, rank_diff=0,
            exp_rank=7, exp_tiebreak=0, points=7,
        ),
    }),
    cells=TT14,
)

SCORED_RUN = Scenario(
    "scored_run",
    "The better Run Score counts; equal runs fall back to each run's own "
    "tiebreaks: lower DS, then TS, then Vehicle Distance.",
    teams=[
        team("Run2Better", ok(1, bd=20), ok(2)),
        team("Run1Better", ok(1), ok(2, bd=20)),
        # 140 each: run 2 has the lower Distance Score (30 vs 40).
        team("EqualLowerDS", ok(1, bd=20), ok(2, time=34)),
        team("Identical", ok(1), ok(2)),
        team("FailThenOK", run(1, sf=F), ok(2)),
        # 140 each with DS 40 and TS 0: TB3, the shorter Vehicle Distance, picks.
        team("EqualLowerVD1", ok(1, vd=10, bd=20), ok(2, vd=15, bd=10)),
        team("EqualLowerVD2", ok(1, vd=15, bd=10), ok(2, vd=10, bd=20)),
    ],
    expect=keyed({
        "Run2Better": dict(
            r1=140, r2=130, r1_tb1=40, r1_tb2=0, r1_tb3=10, r2_tb1=30, r2_tb2=0,
            r2_tb3=10, scored="->", run_score=130, tb1=30, tb2=0, tb3=10, tb4=35,
        ),
        "Run1Better": dict(
            r1=130, r2=140, r1_tb1=30, r1_tb2=0, r1_tb3=10, r2_tb1=40, r2_tb2=0,
            r2_tb3=10, scored="<-", run_score=130, tb1=30, tb2=0, tb3=10, tb4=35,
        ),
        "EqualLowerDS": dict(
            r1=140, r2=140, r1_tb1=40, r1_tb2=0, r1_tb3=10, r2_tb1=30, r2_tb2=10,
            r2_tb3=10, scored="->", run_score=140, tb1=30, tb2=10, tb3=10, tb4=35,
        ),
        "Identical": dict(
            r1=130, r2=130, r1_tb1=30, r1_tb2=0, r1_tb3=10, r2_tb1=30, r2_tb2=0,
            r2_tb3=10, scored="<-", run_score=130, tb1=30, tb2=0, tb3=10, tb4=35,
        ),
        "FailThenOK": dict(
            r1=FAILED, r2=130, r1_tb1=2500, r1_tb2=7, r1_tb3="", r2_tb1=30, r2_tb2=0,
            r2_tb3=10, scored="->", run_score=130, tb1=30, tb2=0, tb3=10, tb4=35,
        ),
        "EqualLowerVD1": dict(
            r1=140, r2=140, r1_tb1=40, r1_tb2=0, r1_tb3=10, r2_tb1=40, r2_tb2=0,
            r2_tb3=15, scored="<-", run_score=140, tb1=40, tb2=0, tb3=10, tb4=35,
        ),
        "EqualLowerVD2": dict(
            r1=140, r2=140, r1_tb1=40, r1_tb2=0, r1_tb3=15, r2_tb1=40, r2_tb2=0,
            r2_tb3=10, scored="->", run_score=140, tb1=40, tb2=0, tb3=10, tb4=35,
        ),
    }),
    cells=TT14,
)

STATUSES = Scenario(
    "statuses",
    "C / DQ / NS / P / ERR: score, tier, rank, points and the final rankings list.",
    teams=[
        team("Competitor", ok(1)),
        team("Disqualified", ok(1), dq="T"),
        bare("NoShow"),
        team("Lithium", battery="F"),
        team("Errored", run(1, sf=S)),
    ],
    expect=keyed({
        "Competitor":   dict(status="C", exp_score=130, tier=1, exp_tier=1, exp_rank=1, points=1),
        "Disqualified": dict(
            status="DQ", exp_score="DQ", tier="DQ", exp_tier="DQ", exp_rank="DQ", points=7,
        ),
        "NoShow": dict(
            status="NS", exp_score="NS", tier="NS", exp_tier="NS", exp_rank="NS", points=6,
        ),
        "Lithium": dict(
            status="P", exp_score="P", tier="P", exp_tier="P", exp_rank="P", points=5,
        ),
        "Errored": dict(
            status="ERR", exp_score="ERR", tier="ERR", exp_tier="ERR", exp_rank="ERR",
            points="ERR",
        ),
    }),
    extra=[
        ("blank row status", "BX{unused}", ""),
        ("blank row score", "CW{unused}", ""),
        ("listed teams", LISTED, 4),
        ("1st", LIST_SCHOOL % 9, "Competitor"),
        ("2nd", LIST_SCHOOL % 10, "Lithium"),
        ("3rd", LIST_SCHOOL % 11, "NoShow"),
        ("4th", LIST_SCHOOL % 12, "Disqualified"),
        ("no 5th (ERR unlisted)", LIST_SCHOOL % 13, ""),
    ],
    cells=TT14,
)

INPUT_ERRORS = Scenario(
    "input_errors",
    "A Successful run missing a measurement is ERR -- shown by red input "
    "cells, not an Errors message; errored teams don't take rank slots.",
    teams=[
        team("MissingBottle", run(1, sf=S, time=14, vd=10)),
        team("SOnly", run(1, sf=S)),
        team("InferredPartial", run(1, time=14)),
        team("Negative", run(1, sf=S, time=14, vd=-5, bd=10)),
        team("BothRuns", run(1, sf=S), run(2, sf=S)),
        team("Run2Only", ok(1), run(2, sf=S, time=14, bd=3)),
        team("FailedEmpty", run(1, sf=F)),
        team("Clean", ok(1)),
        team("Clean2", ok(1, bd=20)),
    ],
    expect=keyed({
        "MissingBottle": dict(
            status="ERR", errors="", exp_score="ERR", exp_rank="ERR", points="ERR",
        ),
        "SOnly": dict(
            status="ERR", errors="", exp_score="ERR", exp_rank="ERR", points="ERR",
        ),
        "InferredPartial": dict(
            status="ERR", errors="", exp_score="ERR", exp_rank="ERR", points="ERR",
        ),
        "Negative": dict(
            status="ERR", errors="", exp_score="ERR", exp_rank="ERR", points="ERR",
        ),
        "BothRuns": dict(
            status="ERR", errors="", exp_score="ERR", exp_rank="ERR", points="ERR",
        ),
        "Run2Only": dict(
            status="ERR", errors="", exp_score="ERR", exp_rank="ERR", points="ERR",
        ),
        "FailedEmpty":     dict(status="C", errors="", exp_score=FAILED, exp_rank=3, points=3),
        "Clean":           dict(status="C", errors="", exp_score=130, exp_rank=1, points=1),
        "Clean2":          dict(status="C", errors="", exp_score=140, exp_rank=2, points=2),
    }),
    extra=[
        ("errors header (no counter)", "CQ8", "Errors"),
        ("listed teams", LISTED, 3),
        ("1st", LIST_SCHOOL % 9, "Clean"),
        ("2nd", LIST_SCHOOL % 10, "Clean2"),
        ("3rd", LIST_SCHOOL % 11, "FailedEmpty"),
        ("no 4th", LIST_SCHOOL % 12, ""),
    ],
    cells=TT14,
)

TARGET_TIME_MISSING = Scenario(
    "target_time_missing",
    "No Target Time: every competing team ERR; DQ/NS/P keep their status and "
    "show no error.",
    teams=[
        team("Competitor", ok(1)),
        team("Failed", run(1, sf=F)),
        team("Disqualified", ok(1), dq="T"),
        bare("NoShow"),
        team("Lithium", battery="F"),
    ],
    expect=keyed({
        "Competitor":   dict(status="ERR", errors=ERR_TT, exp_rank="ERR", points="ERR"),
        "Failed":       dict(status="ERR", errors=ERR_TT, exp_rank="ERR", points="ERR"),
        "Disqualified": dict(status="DQ", errors="", exp_rank="DQ", points=7),
        "NoShow":       dict(status="NS", errors="", exp_rank="NS", points=6),
        "Lithium":      dict(status="P", errors="", exp_rank="P", points=5),
    }),
    cells={TARGET_TIME: None},
)

TARGET_TIME_TEXT = Scenario(
    "target_time_text",
    "A non-numeric Target Time has its own message.",
    teams=[team("Competitor", ok(1))],
    expect=keyed({"Competitor": dict(status="ERR", errors=ERR_TT_TEXT)}),
    cells={TARGET_TIME: "fourteen"},
)

BREAKDOWN = Scenario(
    "breakdown",
    "Breakdown shows the scored run's DS, TS, bonuses (incl. Event Time Bonus) "
    "and penalties (incl. impound); blank when not competing.",
    teams=[
        team("Example",
             run(1, sf=S, time=12.27, vd=17.6, bd=10.6, bonus="F"),
             run(2, sf=S, time=16.37, vd=22.5, bd=12.1, bonus="T"),
             pusher=16),
        team("Penalised", ok(1, comp="F"), ok(2, bd=20), kit="F", impound="F"),
        team("TimeLeft", ok(1), left_min=1),
        team("DoubleFailed", run(1, sf=F, const="F"), run(2, sf=F)),
        team("Disqualified", ok(1), dq="T"),
        team("Errored", run(1, sf=S)),
    ],
    expect=keyed({
        "Example": dict(
            bd_run=2, bd_ds=57.1, bd_ts=1.185, bd_bonus=-48.5, bd_pen=0, final=109.785,
        ),
        # run 1: 130 + 150 + 50 = 330; run 2: 140 + 50 = 190 -> run 2, + 5000
        "Penalised":    dict(bd_run=2, bd_ds=40, bd_ts=0, bd_bonus=0, bd_pen=5050, final=5190),
        "TimeLeft":     dict(bd_run=1, bd_ds=30, bd_ts=0, bd_bonus=-2, bd_pen=0, final=128),
        # run 2 (2607) beats run 1 (2907)
        "DoubleFailed": dict(bd_run=2, bd_ds=2500, bd_ts=7, bd_bonus=0, bd_pen=0, final=FAILED),
        "Disqualified": dict(bd_run="", bd_ds="", bd_ts="", bd_bonus="", bd_pen="", final="DQ"),
        "Errored":      dict(bd_run="", bd_ds="", bd_ts="", bd_bonus="", bd_pen="", final="ERR"),
    }),
    cells=TT14,
)

SCENARIOS = [RULES_EXAMPLE, FAILED_RUNS, BOTTLE_PAST_TARGET, PENALTIES, EVENT_TIME,
             TIEBREAKS, SCORED_RUN, STATUSES, INPUT_ERRORS,
             TARGET_TIME_MISSING, TARGET_TIME_TEXT, BREAKDOWN]

# The export columns must mirror their working columns: asserting a working
# value also asserts its export twin wherever a scenario checks the former.
_TWINS = {"final": "exp_score", "tier": "exp_tier", "rank_tb": "exp_rank",
          "rank_diff": "exp_tiebreak"}
for _s in SCENARIOS:
    for _school, _exp in _s.expect.items():
        for _src, _dst in _TWINS.items():
            if _src in _exp and _dst not in _exp and _exp.get("status", "C") == "C":
                _exp[_dst] = _exp[_src]

if __name__ == "__main__":
    sys.exit(main(DEFAULT_SHEET, SPEC, SCENARIOS))
