#!/usr/bin/env python3
"""Targeted regression tests for the Wright Stuff C scoresheet.

Each scenario builds a self-contained mock tournament, recalculates the real
workbook with LibreOffice, and asserts on named output columns.

    python3 test_scoresheet.py [path/to/scoresheet.xlsx] [--only SCENARIO]

Expectations encode INTENDED behaviour. A failure means the sheet disagrees
with the rules as specified -- not that the test needs adjusting to match.

Rules under test (Wright Stuff C 2027, rules 4-6):
  - Up to 2 Airplanes and up to 2 official flights (4.e.vi).
  - Flight time is the median of the timers, entered from the checklist
    (4.e.viii), as minutes and seconds; the mirrors convert MM:SS to total
    seconds. A blank half counts as 0.
  - Flight Score = flight time x Flight Log multiplier (6.c), Complete 1.2,
    Partial 1.1, none 1.0 (6.d).
  - Final Score = the LARGER of the two Flight Scores (6.b) -- not a sum.
    A flight not flown counts as 0 s.
  - A blank Flight Log counts as Not present; a blank plane counts as A.
  - Tier 2 for a construction or competition violation (6.f): box 4 is F, or
    a flight was flown with a plane whose construction box (1 or 2) is F.
    Tier 1 ranks above Tier 2 regardless of score.
  - Ties break on the non-scored flight's Flight Score (6.g).

Every team spells out every input box, so a case can be read without looking
anywhere else. `None` is a box left blank.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "test-utils"))

from runner import Scenario, SheetSpec, main  # noqa: E402

HERE = Path(__file__).resolve().parent
DEFAULT_SHEET = HERE.parent / "scoresheet_c.xlsx"

# Input boxes, by the numbering printed in rows 3-4 of the Scoring sheet.
INPUT_COLS = {
    "team_no": "B", "school": "C", "team": "D",
    "const_a": "E",      # box 1  Plane A met all construction parameters
    "const_b": "F",      # box 2  Plane B met all construction parameters
    "log": "G",          # box 3  Flight Log: Complete / Partial / Not present
    "comp": "H",         # box 4  Met all competition parameters
    "p1": "I", "m1": "J", "s1": "K",    # box 5 plane, box 6 time (mins, secs)
    "p2": "L", "m2": "M", "s2": "N",    # boxes 7, 8
    "dq": "O",           # box 9 Disqualify
}

# Computed columns worth asserting on.
OUT_COLS = {
    "mult": "S",
    # MM:SS converted to total seconds, one mirror per flight.
    "secs1": "V", "secs2": "X",
    "status": "Z", "tier": "AA",
    "f1": "AB", "f2": "AC",
    # Flight number (1 or 2) whose Flight Score is the team's Final Score.
    "scored_flight": "AD",
    "best": "AE", "score": "AF", "score_rank": "AG", "rank": "AH",
    "tb1": "AI", "tb1_rank": "AJ", "rank_tb": "AK", "rank_diff": "AL",
    # Visible score breakdown: which flight was scored, and the multiplier.
    "scored": "AN", "exp_mult": "AO",
    "exp_score": "AP", "exp_tier": "AQ", "exp_tiebreak": "AR",
    "exp_rank": "AS", "points": "AT",
}

SPEC = SheetSpec(INPUT_COLS, OUT_COLS, first_row=8, last_row=507)


# --------------------------------------------------------------------------
# Scenarios
# --------------------------------------------------------------------------

SCORING = Scenario(
    "scoring",
    "Final Score is the larger Flight Score, never a sum, with the log multiplier.",
    teams=[
        # The second flight is the better one: 50 scores, 40 is left for TB1.
        dict(school="TwoFlights",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1="A", m1=None, s1=40,
             p2="A", m2=None, s2=50,
             dq=None),
        dict(school="CompleteLog",
             const_a="T", const_b="T", log="Complete", comp="T",
             p1="A", m1=None, s1=40,
             p2="A", m2=None, s2=50,
             dq=None),
        dict(school="PartialLog",
             const_a="T", const_b="T", log="Partial", comp="T",
             p1="A", m1=None, s1=40,
             p2="A", m2=None, s2=50,
             dq=None),
        # One flight: the flight not flown counts as 0 s, so TB1 is 0.
        dict(school="OneFlight",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1="A", m1=None, s1=45,
             p2=None, m2=None, s2=None,
             dq=None),
        # Equal flights: the earlier one is the scored flight.
        dict(school="EqualFlights",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1="A", m1=None, s1=30,
             p2="A", m2=None, s2=30,
             dq=None),
        # A plane named with no time is not a flight.
        dict(school="PlaneNoTime",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1="A", m1=None, s1=20,
             p2="B", m2=None, s2=None,
             dq=None),
        dict(school="Decimals",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1="A", m1=None, s1=12.25,
             p2="A", m2=None, s2=10.5,
             dq=None),
    ],
    expect={
        # 40 + 50 would be 90: a sum would fail here.
        "TwoFlights": dict(f1=40, f2=50, best=50, mult=1, score=50,
                           scored_flight=2, scored=2, tb1=40),
        "CompleteLog": dict(f1=40, f2=50, best=50, mult=1.2, score=60,
                            scored_flight=2, scored=2, tb1=48),
        "PartialLog": dict(best=50, mult=1.1, score=55, tb1=44),
        "OneFlight": dict(f2=0, best=45, score=45, scored_flight=1, tb1=0),
        "EqualFlights": dict(best=30, score=30, scored_flight=1, tb1=30),
        "PlaneNoTime": dict(f2=0, best=20, score=20, scored_flight=1),
        "Decimals": dict(best=12.25, score=12.25, scored_flight=1, tb1=10.5),
    },
)

TIERS = Scenario(
    "tiers",
    "Tier 2 for a competition violation, or for flying a plane that failed construction.",
    teams=[
        dict(school="Clean",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1="A", m1=None, s1=10,
             p2=None, m2=None, s2=None,
             dq=None),
        dict(school="CompFail",
             const_a="T", const_b="T", log="Not present", comp="F",
             p1="A", m1=None, s1=100,
             p2=None, m2=None, s2=None,
             dq=None),
        # Plane A failed construction and was flown -> Tier 2.
        dict(school="FlewFailedA",
             const_a="F", const_b="T", log="Not present", comp="T",
             p1="A", m1=None, s1=90,
             p2=None, m2=None, s2=None,
             dq=None),
        # Plane A failed construction but only B flew -> still Tier 1.
        dict(school="FailedANotFlown",
             const_a="F", const_b="T", log="Not present", comp="T",
             p1="B", m1=None, s1=20,
             p2=None, m2=None, s2=None,
             dq=None),
        # Plane B failed and flew as the second flight -> Tier 2.
        dict(school="FlewFailedBOnce",
             const_a="T", const_b="F", log="Not present", comp="T",
             p1="A", m1=None, s1=80,
             p2="B", m2=None, s2=5,
             dq=None),
        # Failed plane A named on a flight that was never timed -> Tier 1.
        dict(school="FailedANamedOnly",
             const_a="F", const_b="T", log="Not present", comp="T",
             p1="A", m1=None, s1=None,
             p2="B", m2=None, s2=30,
             dq=None),
    ],
    expect={
        "Clean": dict(tier=1, score=10, rank=3, exp_tier=1),
        "CompFail": dict(tier=2, score=100, rank=4, exp_tier=2),
        "FlewFailedA": dict(tier=2, score=90, rank=5),
        "FailedANotFlown": dict(tier=1, score=20, rank=2),
        "FlewFailedBOnce": dict(tier=2, score=80, rank=6),
        "FailedANamedOnly": dict(tier=1, score=30, rank=1),
    },
)

# Cascade: tier -> score -> TB1 (the non-scored flight's Flight Score).
TIEBREAKS = Scenario(
    "tiebreaks",
    "Score then TB1 (non-scored Flight Score), plus an unbroken tie, within tiers.",
    teams=[
        dict(school="Anchor",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1="A", m1=None, s1=60,
             p2="A", m2=None, s2=10,
             dq=None),
        # 50 each: split on the non-scored flight (40 vs 20)
        dict(school="SplitTB1_Win",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1="A", m1=None, s1=50,
             p2="A", m2=None, s2=40,
             dq=None),
        dict(school="SplitTB1_Lose",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1="A", m1=None, s1=20,
             p2="A", m2=None, s2=50,
             dq=None),
        # 30 each with the same non-scored flight: genuine tie
        dict(school="TrueTieA",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1="A", m1=None, s1=30,
             p2="A", m2=None, s2=10,
             dq=None),
        dict(school="TrueTieB",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1="A", m1=None, s1=10,
             p2="A", m2=None, s2=30,
             dq=None),
        # Same score and TB1 as the tied pair, but Tier 2: must not join the tie.
        dict(school="TierTwoSame",
             const_a="T", const_b="T", log="Not present", comp="F",
             p1="A", m1=None, s1=30,
             p2="A", m2=None, s2=10,
             dq=None),
    ],
    expect={
        "Anchor": dict(score=60, rank=1, rank_tb=1, rank_diff=0, points=1),
        "SplitTB1_Win": dict(score=50, tb1=40, rank=2, rank_tb=2,
                             rank_diff=0, points=2),
        "SplitTB1_Lose": dict(score=50, tb1=20, rank=2, rank_tb=3,
                              rank_diff=-1, exp_tiebreak=-1, points=3),
        # Both keep rank 4 -- the sheet must not invent a split.
        "TrueTieA": dict(score=30, tb1=10, scored_flight=1, rank=4, rank_tb=4, points=4),
        "TrueTieB": dict(score=30, tb1=10, scored_flight=2, rank=4, rank_tb=4, points=4),
        "TierTwoSame": dict(score=30, tier=2, rank=6, rank_tb=6, points=6),
    },
)

# 5 teams entered, so DQ = 5+2 = 7, NS = 5+1 = 6, P = 5.
STATUSES = Scenario(
    "statuses",
    "Status assignment and the points awarded for DQ / NS / P.",
    teams=[
        dict(school="Winner",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1="A", m1=None, s1=50,
             p2="A", m2=None, s2=40,
             dq=None),
        dict(school="Runner",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1="A", m1=None, s1=40,
             p2="A", m2=None, s2=30,
             dq=None),
        dict(school="Disqualified",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1="A", m1=None, s1=60,
             p2="A", m2=None, s2=60,
             dq="T"),
        # Every box blank: a No-Show.
        dict(school="NoShow",
             const_a=None, const_b=None, log=None, comp=None,
             p1=None, m1=None, s1=None,
             p2=None, m2=None, s2=None,
             dq=None),
        # Checked in, but no flight time recorded.
        dict(school="Participated",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1=None, m1=None, s1=None,
             p2=None, m2=None, s2=None,
             dq=None),
    ],
    expect={
        "Winner": dict(status="C", tier=1, score=50, rank=1,
                       exp_rank=1, exp_score=50, points=1),
        "Runner": dict(status="C", tier=1, score=40, rank=2,
                       exp_rank=2, exp_score=40, points=2),
        "Disqualified": dict(status="DQ", tier="DQ", score="DQ", rank="DQ",
                             exp_rank="DQ", exp_score="DQ", exp_tier="DQ", points=7),
        "NoShow": dict(status="NS", tier="NS", score="NS", rank="NS",
                       exp_rank="NS", exp_score="NS", points=6),
        "Participated": dict(status="P", tier="P", score="P", rank="P",
                             exp_rank="P", exp_score="P", scored="P", points=5),
    },
)

# The flight time box is two columns, minutes and seconds; the mirrors convert
# them to total seconds before anything else uses the time.
MM_SS = Scenario(
    "mm_ss",
    "Minutes and seconds convert to total seconds; a blank half counts as 0.",
    teams=[
        # 1:05 and 2:00 -> 65 s and 120 s, so flight 2 is the better one.
        dict(school="MinutesAndSeconds",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1="A", m1=1, s1=5,
             p2="A", m2=2, s2=None,
             dq=None),
        # Seconds only, including a value past a minute, and decimals.
        dict(school="SecondsOnly",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1="A", m1=None, s1=75,
             p2="A", m2=None, s2=12.5,
             dq=None),
        # Minutes only on both flights.
        dict(school="MinutesOnly",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1="A", m1=3, s1=None,
             p2="A", m2=1, s2=None,
             dq=None),
        # Fractional seconds survive the conversion: 1:05.5 = 65.5 s.
        dict(school="FractionalSeconds",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1="A", m1=1, s1=5.5,
             p2=None, m2=None, s2=None,
             dq=None),
        # Neither half entered is a flight not flown, not a 0 s flight.
        dict(school="SecondFlightBlank",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1="A", m1=2, s1=30,
             p2="B", m2=None, s2=None,
             dq=None),
    ],
    expect={
        "MinutesAndSeconds": dict(secs1=65, secs2=120, f1=65, f2=120,
                                  best=120, score=120, scored_flight=2, tb1=65),
        "SecondsOnly": dict(secs1=75, secs2=12.5, best=75, score=75,
                            scored_flight=1, tb1=12.5),
        "MinutesOnly": dict(secs1=180, secs2=60, best=180, score=180,
                            scored_flight=1, tb1=60),
        "FractionalSeconds": dict(secs1=65.5, best=65.5, score=65.5),
        # The unflown flight's mirror stays blank and only the 0 s default
        # reaches the working column.
        "SecondFlightBlank": dict(secs1=150, secs2="", f2=0, best=150,
                                  score=150, scored_flight=1, tb1=0),
    },
)


# Blank boxes take a default rather than raising an error.
DEFAULTS = Scenario(
    "defaults",
    "A blank Flight Log counts as Not present and a blank plane as Plane A.",
    teams=[
        dict(school="BlankLog",
             const_a="T", const_b="T", log=None, comp="T",
             p1="A", m1=None, s1=30,
             p2="A", m2=None, s2=20,
             dq=None),
        dict(school="BlankPlanes",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1=None, m1=None, s1=30,
             p2=None, m2=None, s2=20,
             dq=None),
        # Unrecognised text is treated the same as blank.
        dict(school="OddPlane",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1="C", m1=None, s1=30,
             p2=None, m2=None, s2=None,
             dq=None),
        # The default is a real Plane A: if A failed construction, flying
        # it by default still drops the team to Tier 2.
        dict(school="BlankPlaneFailedA",
             const_a="F", const_b="T", log="Not present", comp="T",
             p1=None, m1=None, s1=25,
             p2=None, m2=None, s2=None,
             dq=None),
        # ...whereas a failed B is not flown when the box is left blank.
        dict(school="BlankPlaneFailedB",
             const_a="T", const_b="F", log="Not present", comp="T",
             p1=None, m1=None, s1=25,
             p2=None, m2=None, s2=None,
             dq=None),
    ],
    expect={
        "BlankLog": dict(mult=1, best=30, score=30, tier=1),
        "BlankPlanes": dict(best=30, score=30, tier=1),
        "OddPlane": dict(score=30, tier=1),
        "BlankPlaneFailedA": dict(score=25, tier=2),
        "BlankPlaneFailedB": dict(score=25, tier=1),
    },
)

# Data validation stops a text or negative flight time being typed in, but a
# paste can get past it. Such a time counts as a flight not flown (0 s), and
# the row still scores and ranks normally.
UNREADABLE_TIMES = Scenario(
    "unreadable_times",
    "A pasted text or negative flight time counts as 0 s and breaks nothing else.",
    teams=[
        dict(school="Clean",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1="A", m1=None, s1=50,
             p2="A", m2=None, s2=20,
             dq=None),
        dict(school="TextTime",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1="A", m1=None, s1=40,
             p2="A", m2=None, s2="abc",
             dq=None),
        dict(school="NegativeTime",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1="A", m1=None, s1=-5,
             p2="A", m2=None, s2=30,
             dq=None),
        # An unreadable minutes box voids that flight's whole time.
        dict(school="TextMinutes",
             const_a="T", const_b="T", log="Not present", comp="T",
             p1="A", m1="abc", s1=20,
             p2="A", m2=None, s2=25,
             dq=None),
    ],
    expect={
        "Clean": dict(score=50, rank=1, points=1),
        "TextTime": dict(f2=0, best=40, score=40, scored_flight=1, tb1=0,
                         rank=2, points=2),
        "NegativeTime": dict(f1=0, best=30, score=30, scored_flight=2, tb1=0,
                             rank=3, points=3),
        "TextMinutes": dict(secs1="", f1=0, best=25, score=25,
                            scored_flight=2, tb1=0, rank=4, points=4),
    },
)

SCENARIOS = [SCORING, TIERS, TIEBREAKS, STATUSES, MM_SS, DEFAULTS,
             UNREADABLE_TIMES]

# AN/AO are the visible breakdown; they exist only to display the working
# columns AD/S. Wherever a scenario asserts one, assert the same value on the
# column that shows it, so the two cannot drift apart.
MIRRORED = {"mult": "exp_mult", "scored_flight": "scored"}
for _sc in SCENARIOS:
    for _expect in _sc.expect.values():
        for _component, _shown in MIRRORED.items():
            if _component in _expect and _shown not in _expect:
                _expect[_shown] = _expect[_component]


if __name__ == "__main__":
    sys.exit(main(DEFAULT_SHEET, SPEC, SCENARIOS))
