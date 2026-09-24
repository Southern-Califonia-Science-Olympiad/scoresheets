#!/usr/bin/env python3
"""Targeted regression tests for the Ping Pong Parachute B/C scoresheet.

Each scenario builds a self-contained mock tournament, recalculates the real
workbook with LibreOffice, and asserts on named output columns.

    python3 test_scoresheet.py [path/to/scoresheet.xlsx] [--only SCENARIO]

Expectations encode INTENDED behaviour. A failure means the sheet disagrees
with the rules as specified -- not that the test needs adjusting to match.

Rules under test:
  - Two flights, each on Rocket A or B (a blank or unrecognised rocket is A).
  - A flight is scoreable only if it has a time and its rocket passed
    construction (box 1 for A, box 2 for B; blank counts as passed). A flight
    on a failed rocket is ignored: it scores 0 and is never the scored flight
    unless the other flight scores 0 too.
  - A team with no scoreable flight is P (Participated).
  - Flight score = time x Practice Log multiplier (Complete 1, Incomplete
    0.85, Not present or blank 0.7) x 0.9 if the parachute did not separate
    x 0.25 if the rocket or parachute touched the ceiling.
  - Score = the higher flight score; equal scores pick Flight 1.
  - Rank by score, then TB1 = the other flight's score.
  - The event has no tiering: every team that competed is Tier 1.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "test-utils"))

from runner import Scenario, SheetSpec, main  # noqa: E402

HERE = Path(__file__).resolve().parent
DEFAULT_SHEET = HERE.parent / "scoresheet_bc.xlsx"

# Input boxes, by the numbering printed in row 3 of the Scoring sheet.
INPUT_COLS = {
    "team_no": "B", "school": "C", "team": "D",
    "const_a": "E",      # box 1  Rocket A met all construction parameters
    "const_b": "F",      # box 2  Rocket B met all construction parameters
    "log": "G",          # box 3  Practice Log: Complete / Incomplete / Not present
    # Flight 1: box 4 rocket, 5 time, 6 parachute separates, 7 no ceiling touch
    "r1": "H", "t1": "I", "sep1": "J", "ceil1": "K",
    # Flight 2: boxes 8-11, same order
    "r2": "L", "t2": "M", "sep2": "N", "ceil2": "O",
    "dq": "P",           # box 12 Disqualify
}

# Computed columns worth asserting on.
OUT_COLS = {
    "log_mult": "T",
    "f1_ok": "AD", "f2_ok": "AE", "status": "AF", "tier": "AG",
    "f1_time": "AH", "f1_mult": "AI", "f1_score": "AJ",
    "f2_time": "AK", "f2_mult": "AL", "f2_score": "AM",
    "scored": "AN", "non_scored": "AO", "score": "AP", "rank": "AQ",
    "tb1": "AR", "tb1_rank": "AS", "rank_tb": "AT", "rank_diff": "AU",
    # Visible breakdown and export block.
    "exp_scored": "AW", "exp_time": "AX", "exp_mult": "AY",
    "exp_score": "AZ", "exp_tier": "BA", "exp_tiebreak": "BB",
    "exp_rank": "BC", "points": "BD",
}

SPEC = SheetSpec(INPUT_COLS, OUT_COLS, first_row=8, last_row=507)


def team(school, *flights, **kw):
    """A competing team: both rockets pass and the Practice Log is Complete
    (multiplier 1), so a flight's score equals its time unless overridden.

    Each flight is (rocket, seconds); None leaves that box blank. Parachute
    and ceiling boxes are left blank (pass) unless given as sep1=, ceil2=, ...
    Pass const_a=None etc. to leave a box blank.
    """
    row = {"school": school, "const_a": "T", "const_b": "T", "log": "Complete"}
    for i, (rocket, seconds) in enumerate(flights, 1):
        row["r%d" % i] = rocket
        row["t%d" % i] = seconds
    row.update(kw)
    return {k: v for k, v in row.items() if v is not None}


def bare(school, **kw):
    """A team with no box inputs at all -- boxes 1-11 blank means No-Show."""
    row = {"school": school}
    row.update(kw)
    return row


# --------------------------------------------------------------------------
# Scenarios
# --------------------------------------------------------------------------

# 9 teams entered, so P = 9 points.
CONSTRUCTION = Scenario(
    "construction",
    "A flight on a rocket that failed construction is ignored; no scoreable flight is P.",
    teams=[
        team("Clean", ("A", 20), ("B", 30)),
        # Rocket A failed and flew the longer flight: only B's flight counts.
        team("FailedAFlown", ("A", 50), ("B", 30), const_a="F"),
        team("FailedBFlown", ("A", 25), ("B", 60), const_b="F"),
        # Rocket A failed but only B flew: nothing is lost.
        team("FailedANotFlown", ("B", 15), ("B", 18), const_a="F"),
        # Both rockets failed: nothing left to score.
        team("BothFailed", ("A", 40), ("B", 40), const_a="F", const_b="F"),
        # Brought one rocket (A, rocket boxes left blank) and it failed.
        # Box 2 may be left blank or marked F for the rocket not brought.
        team("OneRocketFailed_B2Blank", (None, 40), (None, 35),
             const_a="F", const_b=None),
        team("OneRocketFailed_B2F", (None, 40), (None, 35),
             const_a="F", const_b="F"),
        # Brought one rocket and it passed: scores normally.
        team("OneRocketPassed", (None, 22), (None, 24), const_b=None),
        # A blank rocket box is Rocket A, so it is ignored when A failed.
        team("BlankRocketIsA", (None, 50), ("B", 10), const_a="F"),
    ],
    expect={
        "Clean": dict(f1_ok=True, f2_ok=True, status="C", scored=2, score=30,
                      tb1=20, rank=1, rank_tb=1, points=1),
        # The ignored 50 s flight neither scores nor serves as the tiebreak,
        # so this team loses the tie with Clean on TB1 (0 vs 20).
        "FailedAFlown": dict(f1_ok=False, f2_ok=True, status="C", f1_time=0,
                             f1_score=0, f2_score=30, scored=2, score=30, tb1=0,
                             rank=1, rank_tb=2, points=2),
        "FailedBFlown": dict(f1_ok=True, f2_ok=False, f2_time=0, f2_score=0,
                             scored=1, score=25, tb1=0, rank_tb=3, points=3),
        "OneRocketPassed": dict(f1_ok=True, f2_ok=True, scored=2, score=24,
                                tb1=22, rank_tb=4, points=4),
        "FailedANotFlown": dict(f1_ok=True, f2_ok=True, scored=2, score=18,
                                tb1=15, rank_tb=5, points=5),
        "BlankRocketIsA": dict(f1_ok=False, f2_ok=True, scored=2, score=10,
                               tb1=0, rank_tb=6, points=6),
        "BothFailed": dict(f1_ok=False, f2_ok=False, status="P", tier="P",
                           score="P", rank="P", exp_score="P", exp_rank="P",
                           exp_scored="", exp_time="", exp_mult="", points=9),
        "OneRocketFailed_B2Blank": dict(f1_ok=False, f2_ok=False, status="P",
                                        score="P", rank="P", points=9),
        "OneRocketFailed_B2F": dict(f1_ok=False, f2_ok=False, status="P",
                                    score="P", rank="P", points=9),
    },
)

# One 20 s flight per team unless stated, so each flight score is 20 x mult.
MULTIPLIERS = Scenario(
    "multipliers",
    "Practice Log, parachute and ceiling multipliers, alone and combined, per flight.",
    teams=[
        team("LogComplete", ("A", 20)),
        team("LogIncomplete", ("A", 20), log="Incomplete"),
        team("LogNotPresent", ("A", 20), log="Not present"),
        team("LogBlank", ("A", 20), log=None),
        team("ExplicitT", ("A", 20), sep1="T", ceil1="T"),
        team("NoSeparation", ("A", 20), sep1="F"),
        team("CeilingTouch", ("A", 20), ceil1="F"),
        team("SepAndCeiling", ("A", 20), sep1="F", ceil1="F"),
        team("AllThree", ("A", 20), log="Incomplete", sep1="F", ceil1="F"),
        # The log multiplier applies to both flights.
        team("LogBothFlights", ("A", 20), ("A", 40), log="Incomplete"),
        # Flight 2's penalties come from boxes 10/11, not 6/7.
        team("Flight2Penalty", ("A", 20), ("A", 30), ceil2="F"),
        team("Flight1Penalty", ("A", 20), ("A", 20), sep1="F"),
    ],
    expect={
        "LogComplete": dict(log_mult=1, f1_mult=1, f1_score=20, score=20),
        "LogIncomplete": dict(log_mult=0.85, f1_mult=0.85, f1_score=17, score=17),
        "LogNotPresent": dict(log_mult=0.7, f1_mult=0.7, f1_score=14, score=14),
        # A blank log counts as Not present.
        "LogBlank": dict(log_mult=1, f1_mult=1, f1_score=20, score=20),
        "ExplicitT": dict(f1_mult=1, f1_score=20, score=20),
        "NoSeparation": dict(f1_mult=0.9, f1_score=18, score=18),
        "CeilingTouch": dict(f1_mult=0.25, f1_score=5, score=5),
        "SepAndCeiling": dict(f1_mult=0.225, f1_score=4.5, score=4.5),
        "AllThree": dict(f1_mult=0.19125, f1_score=3.825, score=3.825),
        "LogBothFlights": dict(f1_score=17, f2_mult=0.85, f2_score=34, scored=2,
                               score=34, tb1=17, exp_time=40, exp_mult=0.85),
        # The longer raw flight loses once penalised: 20 x 1 beats 30 x 0.25.
        "Flight2Penalty": dict(f1_mult=1, f2_mult=0.25, f1_score=20,
                               f2_score=7.5, scored=1, non_scored=2, score=20,
                               tb1=7.5, exp_scored=1, exp_time=20, exp_mult=1),
        "Flight1Penalty": dict(f1_mult=0.9, f2_mult=1, f1_score=18, f2_score=20,
                               scored=2, non_scored=1, score=20, tb1=18,
                               exp_scored=2, exp_time=20, exp_mult=1),
    },
)

# Cascade: score -> TB1 (the non-scored flight's score).
# TB1 values across the field: 35, 30, 25, 15, 15, 10, 8.75.
TIEBREAKS = Scenario(
    "tiebreaks",
    "Score then TB1 (the other flight's score), equal flights, and an unbroken tie.",
    teams=[
        team("Anchor", ("A", 50), ("A", 10)),
        # 40 each: split at TB1 (35 vs 30 vs 8.75)
        team("SplitTB1_Win", ("A", 40), ("A", 35)),
        team("SplitTB1_Lose", ("A", 30), ("A", 40)),
        # TB1 is the flight *score*: 35 s with a ceiling touch is only 8.75.
        team("SplitTB1_Penalised", ("A", 40), ("A", 35), ceil2="F"),
        # Equal flights: Flight 1 is scored, Flight 2 is the tiebreak.
        team("EqualFlights", ("A", 25), ("A", 25)),
        # 20 each with TB1 15 each: a genuine tie.
        team("TrueTieA", ("A", 20), ("A", 15)),
        team("TrueTieB", ("A", 15), ("A", 20)),
    ],
    expect={
        "Anchor": dict(score=50, tb1=10, rank=1, tb1_rank=6, rank_tb=1,
                       rank_diff=0, exp_rank=1, points=1),
        "SplitTB1_Win": dict(score=40, scored=1, tb1=35, rank=2, tb1_rank=1,
                             rank_tb=2, rank_diff=0, points=2),
        "SplitTB1_Lose": dict(score=40, scored=2, tb1=30, rank=2, tb1_rank=2,
                              rank_tb=3, rank_diff=-1, exp_tiebreak=-1,
                              exp_rank=3, points=3),
        "SplitTB1_Penalised": dict(score=40, tb1=8.75, rank=2, tb1_rank=7,
                                   rank_tb=4, rank_diff=-2, points=4),
        "EqualFlights": dict(scored=1, non_scored=2, exp_scored=1, score=25,
                             tb1=25, rank=5, rank_tb=5, points=5),
        # Both keep rank 6 -- the sheet must not invent a split.
        "TrueTieA": dict(scored=1, score=20, tb1=15, rank=6, rank_tb=6, points=6),
        "TrueTieB": dict(scored=2, score=20, tb1=15, rank=6, rank_tb=6, points=6),
    },
)

NO_BREAKDOWN = dict(exp_scored="", exp_time="", exp_mult="")

# 5 teams entered, so DQ = 5+2 = 7, NS = 5+1 = 6, P = 5.
STATUSES = Scenario(
    "statuses",
    "Status assignment, Tier 1 for everyone who competed, and DQ / NS / P points.",
    teams=[
        team("Winner", ("A", 30), ("B", 25)),
        team("Runner", ("A", 20), ("B", 15)),
        team("Disqualified", ("A", 60), ("B", 60), dq="T"),
        bare("NoShow"),
        # Checked in, but no flight time recorded.
        team("Participated"),
    ],
    expect={
        "Winner": dict(status="C", tier=1, score=30, rank=1, exp_tier=1,
                       exp_score=30, exp_rank=1, points=1),
        "Runner": dict(status="C", tier=1, score=20, rank=2, exp_tier=1,
                       exp_score=20, exp_rank=2, points=2),
        # The flight breakdown (AW-AY) is blank unless the team competed;
        # the export columns still carry the status code.
        "Disqualified": dict(status="DQ", tier="DQ", score="DQ", rank="DQ",
                             exp_score="DQ", exp_tier="DQ", exp_rank="DQ",
                             points=7, **NO_BREAKDOWN),
        "NoShow": dict(status="NS", tier="NS", score="NS", rank="NS",
                       exp_score="NS", exp_tier="NS", exp_rank="NS", points=6,
                       **NO_BREAKDOWN),
        "Participated": dict(status="P", tier="P", score="P", rank="P",
                             exp_score="P", exp_tier="P", exp_rank="P",
                             points=5, **NO_BREAKDOWN),
    },
    # A row with no team entered must stay empty all the way to Points.
    # BP is the final rankings list, filled from the top, so its rows are
    # fixed: every entered team is listed, DQ/NS/P included, in points order.
    extra=[("AF{unused}", ""), ("AG{unused}", ""), ("AP{unused}", ""),
           ("AW{unused}", ""), ("AZ{unused}", ""), ("BA{unused}", ""),
           ("BD{unused}", ""),
           ("teams in final rankings", "BF7", 5),
           ("1st place in final rankings", "BP8", 1),
           ("2nd place in final rankings", "BP9", 2),
           ("then Participated", "BP10", "P"),
           ("then No-Show", "BP11", "NS"),
           ("Disqualified last", "BP12", "DQ"),
           ("nothing after last place", "BP13", ""),
           ("winner's tier in final rankings", "BT8", 1),
           ("winner's score in final rankings", "BU8", 30)],
)

# Blank or odd boxes take a default rather than raising an error. Data
# validation stops a text or negative time being typed in, but a paste can
# get past it; such a flight is simply not scoreable.
DEFAULTS = Scenario(
    "defaults",
    "Blank construction boxes pass, odd rockets are A, unreadable times don't score.",
    teams=[
        team("BlankConstBoxes", ("A", 15), ("B", 16), const_a=None, const_b=None),
        team("OddRocketIsA", ("C", 30), ("B", 10), const_a="F"),
        team("RocketNoTime", ("B", None), ("A", 12)),
        team("TextTime", ("A", "abc"), ("A", 20)),
        team("NegativeTime", ("A", -5), ("A", 20)),
    ],
    expect={
        "BlankConstBoxes": dict(f1_ok=True, f2_ok=True, score=16, tb1=15),
        "OddRocketIsA": dict(f1_ok=False, f2_ok=True, score=10, tb1=0),
        # A rocket named with no time is not a flight, and not an error.
        "RocketNoTime": dict(f1_ok=False, f2_ok=True, status="C", f1_time=0,
                             score=12, tb1=0),
        "TextTime": dict(f1_ok=False, f1_time=0, status="C", score=20, tb1=0),
        "NegativeTime": dict(f1_ok=False, f1_time=0, status="C", score=20, tb1=0),
    },
)

SCENARIOS = [CONSTRUCTION, MULTIPLIERS, TIEBREAKS, STATUSES, DEFAULTS]

# Visible columns that only display a working column: wherever a scenario
# asserts the working value, assert the same on the column that shows it, so
# the two cannot drift apart.
MIRRORED = {"scored": "exp_scored", "score": "exp_score", "tier": "exp_tier"}
for _sc in SCENARIOS:
    for _expect in _sc.expect.values():
        for _component, _shown in MIRRORED.items():
            if _component in _expect and _shown not in _expect:
                _expect[_shown] = _expect[_component]


if __name__ == "__main__":
    sys.exit(main(DEFAULT_SHEET, SPEC, SCENARIOS))
