#!/usr/bin/env python3
"""Targeted regression tests for the Elastic Launched Glider B scoresheet.

Each scenario builds a self-contained mock tournament, recalculates the real
workbook with LibreOffice, and asserts on named output columns.

    python3 test_scoresheet.py [path/to/scoresheet.xlsx] [--only SCENARIO]

Expectations encode INTENDED behaviour. A failure means the sheet disagrees
with the rules as specified -- not that the test needs adjusting to match.

Rules under test:
  - Flight score = the flight time entered for that flight (the median of the
    timers, taken from the checklist).
  - Score = sum of the best 3 of 5 flights x Flight Log multiplier
    (Complete 1.2, Partial 1.1, Not present 1). Flights not flown count 0 s.
  - A blank Flight Log counts as Not present; a blank glider counts as A.
  - Tier 2 if box 4 is F, or if any flight was flown with a glider whose
    construction box (1 or 2) is F; otherwise Tier 1. Tier 1 ranks above
    Tier 2 regardless of score.
  - TB1 = the longer non-scored flight, i.e. the 4th best flight time.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "test-utils"))

from runner import Scenario, SheetSpec, main  # noqa: E402

HERE = Path(__file__).resolve().parent
DEFAULT_SHEET = HERE.parent / "scoresheet_b.xlsx"

# Input boxes, by the numbering printed in row 3 of the Scoring sheet.
INPUT_COLS = {
    "team_no": "B", "school": "C", "team": "D",
    "const_a": "E",      # box 1  Glider A met all construction parameters
    "const_b": "F",      # box 2  Glider B met all construction parameters
    "log": "G",          # box 3  Flight Log: Complete / Partial / Not present
    "comp": "H",         # box 4  Met all competition parameters
    "g1": "I", "t1": "J",    # box 5 glider, box 6 flight time
    "g2": "K", "t2": "L",    # boxes 7, 8
    "g3": "M", "t3": "N",    # boxes 9, 10
    "g4": "O", "t4": "P",    # boxes 11, 12
    "g5": "Q", "t5": "R",    # boxes 13, 14
    "dq": "S",           # box 15 Disqualify
}

# Computed columns worth asserting on.
OUT_COLS = {
    "mult": "W", "status": "AJ", "tier": "AK",
    "f1": "AL", "f2": "AM", "f3": "AN", "f4": "AO", "f5": "AP",
    # Flight number (1-5) of the longest, 2nd and 3rd longest flight.
    "first": "AQ", "second": "AR", "third": "AS",
    "best3": "AT", "score": "AU", "score_rank": "AV", "rank": "AW",
    "tb1": "AX", "tb1_rank": "AY", "rank_tb": "AZ", "rank_diff": "BA",
    # Visible score breakdown: which flights were summed, and the multiplier.
    "scored": "BC", "exp_mult": "BD",
    "exp_score": "BE", "exp_tier": "BF", "exp_tiebreak": "BG",
    "exp_rank": "BH", "points": "BI",
}

SPEC = SheetSpec(INPUT_COLS, OUT_COLS, first_row=8, last_row=507)

def team(school, flights=(), **kw):
    """A competing team: boxes 1, 2 and 4 pass and no Flight Log unless overridden.

    `flights` is a list of (glider, seconds) for flights 1, 2, ... in order;
    use None for a box left blank.
    """
    row = {"school": school, "const_a": "T", "const_b": "T", "comp": "T",
           "log": "Not present"}
    for i, (glider, seconds) in enumerate(flights, 1):
        if glider is not None:
            row["g%d" % i] = glider
        if seconds is not None:
            row["t%d" % i] = seconds
    row.update(kw)
    return row


def bare(school, **kw):
    """A team with no box inputs at all -- boxes 1-14 blank means No-Show."""
    row = {"school": school}
    row.update(kw)
    return row


def glider_a(*times):
    return [("A", t) for t in times]


# --------------------------------------------------------------------------
# Scenarios
# --------------------------------------------------------------------------

SCORING = Scenario(
    "scoring",
    "Best 3 of 5 flights, the Flight Log multiplier, and missing flights as 0 s.",
    teams=[
        # Flights out of order on purpose: best 3 are 50+40+30 = 120, TB1 = 20.
        team("FiveFlights", glider_a(30, 10, 50, 20, 40)),
        team("CompleteLog", glider_a(30, 10, 50, 20, 40), log="Complete"),
        team("PartialLog", glider_a(30, 10, 50, 20, 40), log="Partial"),
        # Two flights: the third scored flight is assessed 0 s.
        team("TwoFlights", glider_a(35, 25)),
        team("OneFlight", glider_a(45)),
        # Flight 2 names a glider but has no time: not flown, not an error.
        team("GliderNoTime", [("A", 20), ("B", None), ("A", 15)]),
        # Decimal times survive untouched.
        team("Decimals", glider_a(12.25, 10.5, 11.75)),
        # Four equal best times: the earliest three are the ones summed.
        team("EqualTimes", glider_a(10, 30, 30, 30, 30)),
        # A tie below the top: 2nd and 3rd longest share a time and must
        # still name two different flights, the earlier one first.
        team("TieBelowTop", glider_a(20, 40, 25, 20, 25)),
    ],
    expect={
        "FiveFlights": dict(f1=30, f2=10, f3=50, f4=20, f5=40, best3=120, mult=1,
                            score=120, tb1=20, first=3, second=5, third=1,
                            scored="1, 3, 5"),
        "CompleteLog": dict(best3=120, mult=1.2, score=144, tb1=24),
        "PartialLog": dict(best3=120, mult=1.1, score=132, tb1=22),
        # The 0 s assessed for flight 3 is one of the three summed.
        "TwoFlights": dict(f3=0, f4=0, f5=0, best3=60, score=60, tb1=0,
                           first=1, second=2, third=3, scored="1, 2, 3"),
        "OneFlight": dict(best3=45, score=45, tb1=0),
        "GliderNoTime": dict(f2=0, best3=35, score=35, scored="1, 2, 3"),
        "Decimals": dict(best3=34.5, score=34.5, first=1, second=3, third=2,
                         scored="1, 2, 3"),
        "TieBelowTop": dict(best3=90, first=2, second=3, third=5, tb1=20,
                            scored="2, 3, 5"),
        "EqualTimes": dict(best3=90, score=90, tb1=30, first=2, second=3, third=4,
                           scored="2, 3, 4"),
    },
)

TIERS = Scenario(
    "tiers",
    "Tier 2 for a competition violation, or for flying a glider that failed construction.",
    teams=[
        team("Clean", glider_a(10)),
        team("CompFail", glider_a(100, 100, 100), comp="F"),
        # Glider A failed construction and was flown -> Tier 2.
        team("FlewFailedA", glider_a(90), const_a="F"),
        # Glider A failed construction but only B flew -> still Tier 1.
        team("FailedANotFlown", [("B", 20)], const_a="F"),
        # Glider B failed and flew just once among A flights -> Tier 2.
        team("FlewFailedBOnce", [("A", 80), ("A", 80), ("B", 5)], const_b="F"),
        # Failed glider A named on a flight that was never timed -> Tier 1.
        team("FailedANamedOnly", [("A", None), ("B", 30)], const_a="F"),
    ],
    expect={
        "Clean": dict(tier=1, score=10, rank=3, exp_tier=1),
        "CompFail": dict(tier=2, score=300, rank=4, exp_tier=2),
        "FlewFailedA": dict(tier=2, score=90, rank=6, exp_tier=2),
        "FailedANotFlown": dict(tier=1, score=20, rank=2, exp_tier=1),
        "FlewFailedBOnce": dict(tier=2, score=165, rank=5, exp_tier=2),
        "FailedANamedOnly": dict(tier=1, score=30, rank=1, exp_tier=1),
    },
)

# Cascade: tier -> score -> TB1 (4th best flight).
TIEBREAKS = Scenario(
    "tiebreaks",
    "Score then TB1 (longer non-scored flight), plus an unbroken tie, within tiers.",
    teams=[
        team("Anchor", glider_a(50, 50, 50)),
        # 120 each: split at TB1 (30 vs 20)
        team("SplitTB1_Win", glider_a(40, 30, 40, 40)),
        team("SplitTB1_Lose", glider_a(20, 40, 40, 40, 10)),
        # 90 each and the same TB1: genuine tie
        team("TrueTieA", glider_a(30, 30, 30, 10)),
        team("TrueTieB", glider_a(10, 30, 30, 30, 5)),
        # Same score and TB1 as the tied pair, but Tier 2: must not join the tie.
        team("TierTwoSame", glider_a(30, 30, 30, 10), comp="F"),
    ],
    expect={
        "Anchor": dict(score=150, rank=1, rank_tb=1, rank_diff=0, points=1),
        # Best TB1 in the field (Anchor's is 0 -- it flew only three times).
        "SplitTB1_Win": dict(score=120, tb1=30, rank=2, tb1_rank=1, rank_tb=2,
                             rank_diff=0, points=2),
        "SplitTB1_Lose": dict(score=120, tb1=20, rank=2, rank_tb=3,
                              rank_diff=-1, exp_tiebreak=-1, points=3),
        # Both keep rank 4 -- the sheet must not invent a split.
        "TrueTieA": dict(score=90, tb1=10, rank=4, rank_tb=4, points=4),
        "TrueTieB": dict(score=90, tb1=10, rank=4, rank_tb=4, points=4),
        "TierTwoSame": dict(score=90, tier=2, rank=6, rank_tb=6, points=6),
    },
)

# 5 teams entered, so DQ = 5+2 = 7, NS = 5+1 = 6, P = 5.
STATUSES = Scenario(
    "statuses",
    "Status assignment and the points awarded for DQ / NS / P.",
    teams=[
        team("Winner", glider_a(50, 50, 50)),
        team("Runner", glider_a(40, 40, 40)),
        team("Disqualified", glider_a(60, 60, 60), dq="T"),
        bare("NoShow"),
        # Checked in, but no flight time recorded.
        team("Participated"),
    ],
    expect={
        "Winner": dict(status="C", tier=1, score=150, rank=1,
                       exp_rank=1, exp_score=150, points=1),
        "Runner": dict(status="C", tier=1, score=120, rank=2,
                       exp_rank=2, exp_score=120, points=2),
        "Disqualified": dict(status="DQ", tier="DQ", score="DQ", rank="DQ",
                             exp_rank="DQ", exp_score="DQ", exp_tier="DQ", points=7),
        "NoShow": dict(status="NS", tier="NS", score="NS", rank="NS",
                       exp_rank="NS", exp_score="NS", points=6),
        "Participated": dict(status="P", tier="P", score="P", rank="P",
                             exp_rank="P", exp_score="P", scored="P", points=5),
    },
)

# Blank boxes take a default rather than raising an error.
DEFAULTS = Scenario(
    "defaults",
    "A blank Flight Log counts as Not present and a blank glider as Glider A.",
    teams=[
        team("BlankLog", glider_a(30, 20, 10), log=None),
        team("BlankGliders", [(None, 30), (None, 20), ("B", 10)]),
        # Unrecognised text is treated the same as blank.
        team("OddGlider", [("C", 30)]),
        # The default is a real Glider A: if A failed construction, flying
        # it by default still drops the team to Tier 2.
        team("BlankGliderFailedA", [(None, 25)], const_a="F"),
        # ...whereas a failed B is not flown when the box is left blank.
        team("BlankGliderFailedB", [(None, 25)], const_b="F"),
    ],
    expect={
        "BlankLog": dict(mult=1, best3=60, score=60, tier=1),
        "BlankGliders": dict(best3=60, score=60, tier=1),
        "OddGlider": dict(score=30, tier=1),
        "BlankGliderFailedA": dict(score=25, tier=2),
        "BlankGliderFailedB": dict(score=25, tier=1),
    },
)

# Data validation stops a text or negative flight time being typed in, but a
# paste can get past it. Such a time counts as a flight not flown (0 s), and
# the row still scores and ranks normally.
UNREADABLE_TIMES = Scenario(
    "unreadable_times",
    "A pasted text or negative flight time counts as 0 s and breaks nothing else.",
    teams=[
        team("Clean", glider_a(50, 50, 50)),
        team("TextTime", glider_a(40, 40, 40, "abc")),
        team("NegativeTime", glider_a(30, -5, 30, 30)),
    ],
    expect={
        "Clean": dict(score=150, rank=1, points=1),
        "TextTime": dict(f4=0, best3=120, score=120, tb1=0, rank=2, points=2),
        "NegativeTime": dict(f2=0, best3=90, score=90, first=1, second=3,
                             third=4, scored="1, 3, 4",
                             rank=3, points=3),
    },
)

SCENARIOS = [SCORING, TIERS, TIEBREAKS, STATUSES, DEFAULTS, UNREADABLE_TIMES]

# BD is the visible Log Multiplier; it exists only to display the working
# column W. Wherever a scenario asserts the multiplier, assert the same value
# on the column that shows it, so the two cannot drift apart.
MIRRORED = {"mult": "exp_mult"}
for _sc in SCENARIOS:
    for _expect in _sc.expect.values():
        for _component, _shown in MIRRORED.items():
            if _component in _expect and _shown not in _expect:
                _expect[_shown] = _expect[_component]


if __name__ == "__main__":
    sys.exit(main(DEFAULT_SHEET, SPEC, SCENARIOS))
