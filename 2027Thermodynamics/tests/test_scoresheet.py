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
DEFAULT_SHEET = HERE.parent / "scoresheet_bc.xlsx"

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
    "errors": "BA",
    # Visible score breakdown, immediately right of Errors, in ES/TS/PS order.
    "exp_es": "BB", "exp_ts": "BC", "exp_ps": "BD",
    "exp_score": "BE", "exp_tier": "BF", "exp_tiebreak": "BG",
    "exp_rank": "BH", "points": "BI",
    # Final-rankings helpers: the sort key (Points + row/1000) and its rank.
    "sort_key": "BM", "sort_rank": "BN",
}

SPEC = SheetSpec(INPUT_COLS, OUT_COLS, first_row=8, last_row=507)

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
        team("ConstructionPenaltyA", const_para="F", **SAME_NUMBERS),
        team("ConstructionPenaltyB", const_para="F", **SAME_NUMBERS),
        team("CompetitionPenalty", comp_para="F", **SAME_NUMBERS),
        team("BothPenalties", const_para="F", comp_para="F", **SAME_NUMBERS),
        team("TouchedSetup", no_touch="F", **SAME_NUMBERS),
        team("Hazmat", no_hazmat="F", **SAME_NUMBERS),
        team("NotImpounded", impounded="F", **SAME_NUMBERS),
        team("Clean", **SAME_NUMBERS),
    ],
    expect={
        # Identical inputs on two different rows -- both must yield 0.7.
        "ConstructionPenaltyA": dict(mult=0.7, tg=20, ts=21, ps=14, es=50, score=85),
        "ConstructionPenaltyB": dict(mult=0.7, tg=20, ts=21, ps=14, es=50, score=85),
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

# Raw PS divides by $AL$6, the field's best PE. PE is 1-|actual-predicted|/actual,
# which is exactly 0 for a prediction of 0 or of double the actual temperature,
# so a field of such predictions puts the best PE at 0. Everyone then takes full
# PS credit rather than #DIV/0!, the same rule as ES and TS.
#   TG = 20 (field max) -> TS 30   raw = 10 (field max) -> ES 50   PS 20
NORMALISATION_PE = Scenario(
    "normalisation_pe_zero",
    "Best prediction in the field scores PE 0 -- PS must award full marks, not divide by zero.",
    teams=[
        team("PredictZero", start_temp=20, predicted_temp=0, actual_temp=40, part_i_score=10),
        team("PredictDouble", start_temp=20, predicted_temp=80, actual_temp=40, part_i_score=10),
    ],
    expect={
        # pe=0 confirms the setup really produces the degenerate field.
        "PredictZero": dict(pe=0, ps=20, score=100),
        "PredictDouble": dict(pe=0, ps=20, score=100),
    },
)

# The one harmless way max PE reaches 0: nobody enters a prediction. PE stays
# blank and Raw PS's ISNUMBER(Z) gate skips the division, so PS is simply 0.
# Pins that gate -- removing it would reintroduce the division.
PS_NO_PREDICTIONS = Scenario(
    "ps_no_predictions",
    "No team enters a prediction -- PS must be 0 without dividing.",
    teams=[
        team("NoPredictionA", start_temp=20, actual_temp=40, part_i_score=10),
        team("NoPredictionB", start_temp=20, actual_temp=30, part_i_score=5),
    ],
    expect={
        "NoPredictionA": dict(pe="", ps=0),
        "NoPredictionB": dict(pe="", ps=0),
    },
)

# The Errors column (BA) surfaces the AI/AM checks to the scorer. Messages are
# derived from the inputs rather than from ISERROR(AH), so they stay correct if
# TG/PE are later changed to degrade instead of erroring.
INPUT_ERRORS = Scenario(
    "input_errors",
    "Missing or zero temperature boxes must name themselves in the Errors column.",
    teams=[
        team("Complete", start_temp=20, predicted_temp=40, actual_temp=40, part_i_score=10),
        # box 6 blank -- TG cannot be computed
        team("NoStartTemp", predicted_temp=40, actual_temp=40, part_i_score=10),
        # box 8 blank -- TG and PE both lose their actual temperature
        team("NoActualTemp", start_temp=20, predicted_temp=40, part_i_score=10),
        # box 8 present but zero -- PE divides by it
        team("ZeroActualTemp", start_temp=20, predicted_temp=40, actual_temp=0, part_i_score=10),
    ],
    expect={
        "Complete": dict(errors=""),
        "NoStartTemp": dict(errors="Box 6 or Box 8 is missing."),
        "NoActualTemp": dict(errors="Box 6 or Box 8 is missing."),
        "ZeroActualTemp": dict(errors="Box 8 cannot be 0."),
    },
)

# TG and PE yield "" rather than an error for unusable input, so the MAX that
# normalises every team's score never sees an error value. Without this, one
# missing box turned every team's TS, Score, Rank and Points into #VALUE!.
ERROR_CONTAINMENT = Scenario(
    "error_containment",
    "An errored row is neither scored nor ranked, and disturbs no other team.",
    teams=[
        team("Complete", start_temp=20, predicted_temp=40, actual_temp=40, part_i_score=10),
        team("MissingStart", predicted_temp=40, actual_temp=40, part_i_score=10),
        team("Runner", start_temp=20, predicted_temp=30, actual_temp=30, part_i_score=5),
    ],
    # The final rankings block sorts on BM (= Points + row/1000) and ranks it
    # with RANK over the whole column. A non-numeric Points value must leave
    # that column error-free, or every other team loses its sort rank.
    expect={
        "Complete": dict(ts=30, ps=20, es=50, score=100, rank=1, points=1, errors="",
                         sort_rank=1),
        # Components still compute (TS is 0 without box 6), but no score, rank
        # or points are issued while the error stands, and the errored row
        # takes no sort key -- which is what keeps the rankings error-free.
        "MissingStart": dict(ts=0, ps=20, es=50, score="ERR", rank="ERR",
                             exp_rank="ERR", points="ERR", sort_key="",
                             errors="Box 6 or Box 8 is missing."),
        # Would have placed 3rd behind MissingStart's 70; the errored row must
        # not consume a rank slot, so this is 2nd.
        "Runner": dict(score=60, rank=2, points=2, errors=""),
    },
)

SCENARIOS = [PENALTIES, TIEBREAKS, STATUSES, NORMALISATION_ES, NORMALISATION_TG,
             NORMALISATION_PE, PS_NO_PREDICTIONS, INPUT_ERRORS, ERROR_CONTAINMENT]

# BB/BC/BD are the visible breakdown beside Errors; they exist only to display
# the working columns AP/AK/AO. Wherever a scenario asserts a component, assert
# the same value on the column that shows it, so the two cannot drift apart.
# Derived rather than written out per team: a mirror is not a separate fact.
MIRRORED = {"es": "exp_es", "ts": "exp_ts", "ps": "exp_ps"}
for _sc in SCENARIOS:
    for _expect in _sc.expect.values():
        for _component, _shown in MIRRORED.items():
            if _component in _expect and _shown not in _expect:
                _expect[_shown] = _expect[_component]


if __name__ == "__main__":
    sys.exit(main(DEFAULT_SHEET, SPEC, SCENARIOS))
