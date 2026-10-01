#!/usr/bin/env python3
"""Targeted regression tests for the Ping Pong Parachute B/C scoresheet.

Each scenario builds a self-contained mock tournament, recalculates the real
workbook with LibreOffice, and asserts on named output columns.

    python3 test_scoresheet.py [path/to/scoresheet.xlsx] [--only SCENARIO]

Expectations encode INTENDED behaviour. A failure means the sheet disagrees
with the rules as specified -- not that the test needs adjusting to match.

Rules under test:
  - Two flights per team; there is one construction box (box 1), true when at
    least one rocket meets the construction parameters.
  - Box 1 = F means no rocket can launch: the team is P (Participated).
    A blank box 1 passes.
  - A flight is scoreable only if it has a time. A team with no scoreable
    flight is P.
  - Flight score = time x Practice Log multiplier (Complete 1, Incomplete
    0.85, Not present 0.7; a blank log takes the normal path, 1) x 0.9 if the
    parachute did not separate x 0.25 if the rocket or parachute touched the
    ceiling.
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
    "const": "E",        # box 1  At least one rocket met all construction parameters
    "log": "F",          # box 2  Practice Log: Complete / Incomplete / Not present
    # Flight 1: box 3 time, 4 parachute separates, 5 no ceiling touch
    "t1": "G", "sep1": "H", "ceil1": "I",
    # Flight 2: boxes 6-8, same order
    "t2": "J", "sep2": "K", "ceil2": "L",
    "dq": "M",           # box 9  Disqualify
}

# Computed columns worth asserting on.
OUT_COLS = {
    "log_mult": "P",
    "f1_ok": "X", "f2_ok": "Y", "status": "Z", "tier": "AA",
    "f1_time": "AB", "f1_mult": "AC", "f1_score": "AD",
    "f2_time": "AE", "f2_mult": "AF", "f2_score": "AG",
    "scored": "AH", "non_scored": "AI", "score": "AJ", "rank": "AK",
    "tb1": "AL", "tb1_rank": "AM", "rank_tb": "AN", "rank_diff": "AO",
    # Visible breakdown and export block.
    "exp_scored": "AQ", "exp_time": "AR", "exp_mult": "AS",
    "exp_score": "AT", "exp_tier": "AU", "exp_tiebreak": "AV",
    "exp_rank": "AW", "points": "AX",
}

SPEC = SheetSpec(INPUT_COLS, OUT_COLS, first_row=8, last_row=507)


def team(school, *times, **kw):
    """A competing team: construction passes and the Practice Log is Complete
    (multiplier 1), so a flight's score equals its time unless overridden.

    Each positional argument is a flight time in seconds; None leaves that box
    blank. Parachute and ceiling boxes are left blank (pass) unless given as
    sep1=, ceil2=, ... Pass const=None to leave box 1 blank.
    """
    row = {"school": school, "const": "T", "log": "Complete"}
    for i, seconds in enumerate(times, 1):
        row["t%d" % i] = seconds
    row.update(kw)
    return {k: v for k, v in row.items() if v is not None}


def bare(school, **kw):
    """A team with no box inputs at all -- boxes 1-8 blank means No-Show."""
    row = {"school": school}
    row.update(kw)
    return row


# --------------------------------------------------------------------------
# Scenarios
# --------------------------------------------------------------------------

# 4 teams entered, so P = 4 points.
CONSTRUCTION = Scenario(
    "construction",
    "Box 1 = F means no rocket can launch (P); a blank box 1 passes.",
    teams=[
        team("Clean", 20, 30),
        # Blank box 1 takes the normal path and scores.
        team("BlankConst", 15, 16, const=None),
        # No rocket met construction: nothing to score, whatever was timed.
        team("Failed", 40, 35, const="F"),
        team("FailedNoTimes", const="F"),
    ],
    expect={
        "Clean": dict(f1_ok=True, f2_ok=True, status="C", scored=2, score=30,
                      tb1=20, rank=1, rank_tb=1, points=1),
        "BlankConst": dict(f1_ok=True, f2_ok=True, status="C", scored=2,
                           score=16, tb1=15, rank=2, rank_tb=2, points=2),
        "Failed": dict(f1_ok=False, f2_ok=False, status="P", tier="P",
                       score="P", rank="P", exp_score="P", exp_rank="P",
                       exp_scored="", exp_time="", exp_mult="", points=4),
        "FailedNoTimes": dict(f1_ok=False, f2_ok=False, status="P", tier="P",
                              score="P", rank="P", points=4),
    },
)

# One 20 s flight per team unless stated, so each flight score is 20 x mult.
MULTIPLIERS = Scenario(
    "multipliers",
    "Practice Log, parachute and ceiling multipliers, alone and combined, per flight.",
    teams=[
        team("LogComplete", 20),
        team("LogIncomplete", 20, log="Incomplete"),
        team("LogNotPresent", 20, log="Not present"),
        team("LogBlank", 20, log=None),
        team("ExplicitT", 20, sep1="T", ceil1="T"),
        team("NoSeparation", 20, sep1="F"),
        team("CeilingTouch", 20, ceil1="F"),
        team("SepAndCeiling", 20, sep1="F", ceil1="F"),
        team("AllThree", 20, log="Incomplete", sep1="F", ceil1="F"),
        # The log multiplier applies to both flights.
        team("LogBothFlights", 20, 40, log="Incomplete"),
        # Flight 2's penalties come from boxes 7/8, not 4/5.
        team("Flight2Penalty", 20, 30, ceil2="F"),
        team("Flight1Penalty", 20, 20, sep1="F"),
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
        team("Anchor", 50, 10),
        # 40 each: split at TB1 (35 vs 30 vs 8.75)
        team("SplitTB1_Win", 40, 35),
        team("SplitTB1_Lose", 30, 40),
        # TB1 is the flight *score*: 35 s with a ceiling touch is only 8.75.
        team("SplitTB1_Penalised", 40, 35, ceil2="F"),
        # Equal flights: Flight 1 is scored, Flight 2 is the tiebreak.
        team("EqualFlights", 25, 25),
        # 20 each with TB1 15 each: a genuine tie.
        team("TrueTieA", 20, 15),
        team("TrueTieB", 15, 20),
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
        team("Winner", 30, 25),
        team("Runner", 20, 15),
        team("Disqualified", 60, 60, dq="T"),
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
)

# Blank or odd boxes take a default rather than raising an error. Data
# validation stops a text or negative time being typed in, but a paste can
# get past it; such a flight is simply not scoreable.
DEFAULTS = Scenario(
    "defaults",
    "Blank boxes take the normal path; unreadable times don't score.",
    teams=[
        team("BlankConstBox", 15, 16, const=None),
        team("BlankLog", 20, log=None),
        team("TextTime", "abc", 20),
        team("NegativeTime", -5, 20),
        team("NoTime1", None, 12),
    ],
    expect={
        "BlankConstBox": dict(f1_ok=True, f2_ok=True, score=16, tb1=15),
        "BlankLog": dict(log_mult=1, f1_mult=1, score=20),
        "TextTime": dict(f1_ok=False, f1_time=0, status="C", score=20, tb1=0),
        "NegativeTime": dict(f1_ok=False, f1_time=0, status="C", score=20, tb1=0),
        # A blank time is not a flight, and not an error.
        "NoTime1": dict(f1_ok=False, f2_ok=True, status="C", f1_time=0,
                        score=12, tb1=0),
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
