#!/usr/bin/env python3
"""Targeted regression tests for the Thermodynamics B/C scoresheet.

Each scenario builds a self-contained mock tournament, recalculates the real
workbook with LibreOffice, and asserts on named output columns.

    python3 test_scoresheet.py [path/to/scoresheet.xlsx] [--only SCENARIO]

Expectations encode INTENDED behaviour. A failure means the sheet disagrees
with the rules as specified -- not that the test needs adjusting to match.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "test-utils"))

from runner import Scenario, SheetSpec, main  # noqa: E402

HERE = Path(__file__).resolve().parent
DEFAULT_SHEET = HERE.parent / "Thermodynamics B_C - Scoresheet.xlsx"

# Input boxes, by the numbering printed in row 3 of the Scoring sheet.
INPUT_COLS = {
    "team_no": "B", "school": "C", "team": "D",
    "impounded": "G",    # box 1  Device Impounded
    "no_hazmat": "H",    # box 2  No hazardous materials
    "const_para": "I",   # box 3  Met all construction parameters
    "no_touch": "J",     # box 4  Did not touch setup during Heating Time
    "comp_para": "K",    # box 5  Met all competition parameters
    "start_temp": "L",   # box 6
    "predicted_temp": "M",    # box 7
    "actual_temp": "N",       # box 8
    "part_i_score": "O",          # box 9  Part I Raw Score
    "tb3": "P",          # box 10 Tiebreak 3 ranking
    "dq": "Q",           # box 11 Disqualify
}

# Computed columns worth asserting on.
OUT_COLS = {
    "status": "AE", "tier": "AF", "mult": "AG", "tg": "AH",
    "raw_ts": "AJ", "ts": "AK", "pe": "AL", "raw_ps": "AN", "ps": "AO",
    "es": "AP", "score": "AQ", "rank": "AR",
    "tb1": "AS", "tb1_rank": "AT", "tb2": "AU", "tb2_rank": "AV",
    "tb3_rank": "AW", "rank_tb": "AX", "rank_diff": "AY",
    "exp_score": "BB", "exp_tier": "BC", "exp_tiebreak": "BD",
    "exp_rank": "BE", "points": "BF",
}

SPEC = SheetSpec(INPUT_COLS, OUT_COLS, first_row=8, clear_through=40)

BOXES_OK = {"impounded": "T", "no_hazmat": "T", "const_para": "T",
            "no_touch": "T", "comp_para": "T"}


def team(school, **kw):
    """A competing team: all five T/F boxes pass unless overridden."""
    row = dict(BOXES_OK)
    row["school"] = school
    row.update(kw)
    return row


def bare(school, **kw):
    """A team with no box inputs at all -- boxes 1-9 blank means No-Show."""
    row = {"school": school}
    row.update(kw)
    return row


# --------------------------------------------------------------------------
# Scenarios
# --------------------------------------------------------------------------

# Every team here has identical numeric inputs, so TG/PE/part I score
# normalisation is constant across the field and only the T/F boxes move
# the result.
#   TG = 40-20 = 20 (field max)   PE = 1-|40-40|/40 = 1   Part I Score = 10 (field max)
#   full marks: TS = 30, PS = 20, ES = 50
SAME_NUMBERS = dict(start_temp=20, predicted_temp=40, actual_temp=40, part_i_score=10)

PENALTIES = Scenario(
    "penalties",
    "Construction/competition penalty multipliers and the T/U/W gates.",
    teams=[
        team("ConstructionPenaltyRow8", const_para="F", **SAME_NUMBERS),
        team("ConstructionPenaltyRow9", const_para="F", **SAME_NUMBERS),
        team("CompetitionPenalty", comp_para="F", **SAME_NUMBERS),
        team("BothPenalties", const_para="F", comp_para="F", **SAME_NUMBERS),
        team("TouchedSetup", no_touch="F", **SAME_NUMBERS),
        team("Hazmat", no_hazmat="F", **SAME_NUMBERS),
        team("NotImpounded", impounded="F", **SAME_NUMBERS),
        team("Clean", **SAME_NUMBERS),
    ],
    expect={
        # Identical inputs to the row below it -- both must yield 0.7.
        "ConstructionPenaltyRow8": dict(mult=0.7, tg=20, ts=21, ps=14, es=50, score=85),
        "ConstructionPenaltyRow9": dict(mult=0.7, tg=20, ts=21, ps=14, es=50, score=85),
        "CompetitionPenalty": dict(mult=0.9, tg=20, ts=27, ps=18, es=50, score=95),
        "BothPenalties": dict(mult=0.63, tg=20, ts=18.9, ps=12.6, es=50, score=81.5),
        # Box 4 gates Raw TS and Raw PS but NOT the TG calculation itself.
        "TouchedSetup": dict(mult=1, tg=20, raw_ts=0, ts=0, ps=0, es=50, score=50),
        # Boxes 1 and 2 gate TG and PE themselves, so both come back blank.
        "Hazmat": dict(tg="", pe="", ts=0, ps=0, es=50, score=50),
        "NotImpounded": dict(tg="", pe="", ts=0, ps=0, es=50, score=50),
        "Clean": dict(mult=1, tg=20, ts=30, ps=20, es=50, score=100),
    },
)

# Cascade: score -> TB1 (best TS) -> TB2 (best ES) -> TB3 (entered ranking).
# Field maxima: TG 20, PE 1, raw 10.
TIEBREAKS = Scenario(
    "tiebreaks",
    "Full TB1/TB2/TB3 cascade plus an unbroken tie.",
    teams=[
        # score 100, outright winner
        team("Anchor", start_temp=20, predicted_temp=40, actual_temp=40, part_i_score=10),
        # 60 each: split at TB1 (TS 30 vs 15)
        team("SplitTB1_Win", start_temp=20, predicted_temp=10, actual_temp=40, part_i_score=5),
        team("SplitTB1_Lose", start_temp=20, predicted_temp=30, actual_temp=30, part_i_score=5),
        # 55 each: tie through TB1 and TB2, split at TB3 (1 vs 2)
        team("SplitTB3_Win", start_temp=20, predicted_temp=30, actual_temp=30, part_i_score=4, tb3=1),
        team("SplitTB3_Lose", start_temp=20, predicted_temp=30, actual_temp=30, part_i_score=4, tb3=2),
        # 50 each: tie at TB1 (TS 15), split at TB2 (ES 25 vs 15)
        team("SplitTB2_Win", start_temp=20, predicted_temp=15, actual_temp=30, part_i_score=5),
        team("SplitTB2_Lose", start_temp=20, predicted_temp=30, actual_temp=30, part_i_score=3),
        # 45 each: identical on every criterion -> genuine tie
        team("TrueTieA", start_temp=20, predicted_temp=30, actual_temp=30, part_i_score=2),
        team("TrueTieB", start_temp=20, predicted_temp=30, actual_temp=30, part_i_score=2),
    ],
    expect={
        "Anchor": dict(score=100, rank=1, rank_tb=1, points=1),
        "SplitTB1_Win": dict(score=60, rank=2, ts=30, tb1_rank=1, rank_tb=2, points=2),
        "SplitTB1_Lose": dict(score=60, rank=2, ts=15, tb1_rank=3, rank_tb=3, points=3),
        "SplitTB3_Win": dict(score=55, rank=4, tb3_rank=1, rank_tb=4, points=4),
        "SplitTB3_Lose": dict(score=55, rank=4, tb3_rank=2, rank_tb=5, points=5),
        "SplitTB2_Win": dict(score=50, rank=6, es=25, tb2_rank=2, rank_tb=6, points=6),
        "SplitTB2_Lose": dict(score=50, rank=6, es=15, tb2_rank=7, rank_tb=7, points=7),
        # Both keep rank 8 -- the sheet must not invent a split.
        "TrueTieA": dict(score=45, rank=8, rank_tb=8, points=8),
        "TrueTieB": dict(score=45, rank=8, rank_tb=8, points=8),
    },
)

# 4 teams entered, so DQ = 4+2 = 6, NS = 4+1 = 5, P = 4.
STATUSES = Scenario(
    "statuses",
    "Status assignment and the points awarded for DQ / NS.",
    teams=[
        team("Winner", start_temp=20, predicted_temp=40, actual_temp=40, part_i_score=10),
        team("Runner", start_temp=20, predicted_temp=30, actual_temp=30, part_i_score=5),
        team("Disqualified", dq="T", start_temp=20, predicted_temp=40, actual_temp=40, part_i_score=10),
        bare("NoShow"),
    ],
    expect={
        "Winner": dict(status="C", tier=1, score=100, rank=1,
                       exp_rank=1, exp_score=100, points=1),
        "Runner": dict(status="C", tier=1, score=60, rank=2,
                       exp_rank=2, exp_score=60, points=2),
        "Disqualified": dict(status="DQ", tier="DQ", score="DQ", rank="DQ",
                             exp_rank="DQ", exp_score="DQ", points=6),
        "NoShow": dict(status="NS", tier="NS", score="NS", rank="NS",
                       exp_rank="NS", exp_score="NS", points=5),
    },
    # A row with no team entered must stay empty all the way to Points.
    extra=[("AE12", ""), ("AQ12", ""), ("BE12", ""), ("BF12", "")],
)

# Part I raw score cannot go negative, so a field max of 0 means everyone tied
# at the top of that component and should receive full ES credit, not #DIV/0!.
NORMALISATION_ES = Scenario(
    "normalisation_es_zero",
    "Whole field scores 0 on Part I -- ES must award full marks, not divide by zero.",
    teams=[
        team("ZeroRaw_High", start_temp=20, predicted_temp=40, actual_temp=40, part_i_score=0),
        team("ZeroRaw_Low", start_temp=20, predicted_temp=30, actual_temp=30, part_i_score=0),
    ],
    expect={
        "ZeroRaw_High": dict(es=50, ts=30, ps=20, score=100),
        "ZeroRaw_Low": dict(es=50, ts=15, ps=20, score=85),
    },
)

# Same rule for Temperature Gain. TG is floored at 0 by MAX(0, ...), so a field
# max of 0 means nobody gained heat and the whole field is tied at the top of
# that component -- everyone takes full TS credit rather than #DIV/0!.
NORMALISATION_TG = Scenario(
    "normalisation_tg_zero",
    "Whole field gains no heat -- TS must award full marks, not divide by zero.",
    teams=[
        team("NoGainA", start_temp=40, predicted_temp=40, actual_temp=40, part_i_score=10),
        team("NoGainB", start_temp=40, predicted_temp=40, actual_temp=40, part_i_score=5),
    ],
    expect={
        "NoGainA": dict(ts=30, es=50),
        "NoGainB": dict(ts=30, es=25),
    },
)

SCENARIOS = [PENALTIES, TIEBREAKS, STATUSES, NORMALISATION_ES, NORMALISATION_TG]


if __name__ == "__main__":
    sys.exit(main(DEFAULT_SHEET, SPEC, SCENARIOS))
