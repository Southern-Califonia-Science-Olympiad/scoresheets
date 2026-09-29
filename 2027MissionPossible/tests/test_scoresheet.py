#!/usr/bin/env python3
"""Targeted regression tests for the Mission Possible C scoresheet.

Each scenario builds a self-contained mock tournament, recalculates the real
workbook with LibreOffice, and asserts on named output columns.

    python3 test_scoresheet.py [path/to/scoresheet.xlsx] [--only SCENARIO]

Expectations encode INTENDED behaviour. A failure means the sheet disagrees
with the rules as specified -- not that the test needs adjusting to match.

Rules under test (2027 C rules, sections 4-9; high score wins):
  - Awards: Set-Up 50 (<= 30 min) or 75 (State/Nats <= 15 min); 25 per ASL
    item (on time, format, accurate, labelled); Start Action 25/50/200/0 for
    1/2/3/4 coins, 0 for a touched start; 50 per scorable action (max 12);
    Final Action 250; Time Score 2 per full second up to the Target Time, 0
    at 2x the Target Time; Water Timer 1 per full second before the Target
    Time; No Adjustments 75; Device Size 0.1 per 0.1 cm under 80 cm, max 30
    per dimension.
  - Penalties: 2 per full second past the Target Time (up to 2x); 25 per
    dimension over 80 cm; 25 top/walls not open; 25 per touch, max 3 (a
    touched start is one of them); 50 solid/liquid leaving; 250 electricity
    after 30 s.
  - A touch leading to the Final Action: no Final Action points, Time Score 0.
  - Tier 3: not impounded on time, no eye protection, can't answer build
    questions. Tier 2: construction violation, or any dimension over 82 cm.
    Tier 1 ranks above Tier 2 above Tier 3 regardless of score.
  - No eye protection: does not run, Tier 3, scored on device size alone
    (Device Size Score less the over-80 cm penalty).
  - Unsafe or remote-controlled device, or impounded but not competing: P.
  - Ties: fewest penalty points, then smallest L+W+H.
  - Input errors (missing dimension, Start Action, Device Time; invalid
    counts) make the row ERR: not scored, ranked or listed. The offending
    input cell turns red; only a missing/invalid Target Time (a global input)
    gets an Errors message.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "test-utils"))

from runner import Scenario, SheetSpec, main, table  # noqa: E402

HERE = Path(__file__).resolve().parent
DEFAULT_SHEET = HERE.parent / "scoresheet_c.xlsx"

TARGET_TIME = "D7"

# Input boxes, by the numbering printed in row 3 of the Scoring sheet.
INPUT_COLS = {
    "team_no": "B", "school": "C", "team": "D",
    "impound": "E",        # 1. State/Nats Only: Device Impounded (T/F)
    "asl_on_time": "F",    # 2. ASL Impounded (T/F)
    "eye": "G",            # 3. Eye Protection (T/F)
    "safe": "H",           # 4. Device is safe (T/F)
    "d1": "I", "d2": "J", "d3": "K",   # 5-7. Dimensions (cm)
    "walls": "L",          # 8. Top & 2+ vertical walls transparent/open (T/F)
    "const": "M",          # 9. All const params met (T/F)
    "remote": "N",         # 10. Remote Controlled (T/F)
    "answer": "O",         # 11. Team able to answer questions (T/F)
    "setup30": "P",        # 12. <= 30 min to set up (T/F)
    "setup15": "Q",        # 13. State/Nats Only: <= 15 min to set up (T/F)
    "asl_format": "R",     # 14. ASL proper format (T/F)
    "asl_accurate": "S",   # 15. All actions included and accurate (T/F)
    "asl_labelled": "T",   # 16. ASL # properly labelled (T/F)
    "start": "U",          # 17. Start Action (# of coins or Touch)
    "actions": "V",        # 18. # of successful actions
    "water": "W",          # 19. Water Timer (sec)
    "no_touch_final": "X",  # 20. No Touches leading to Final Action (T/F)
    "final": "Y",          # 21. Final Action Satisfied (T/F)
    "touches": "Z",        # 22. Adjustments/Touches
    "elec": "AA",          # 23. No electricity used after 30 s (T/F)
    "solid": "AB",         # 24. No Solid/liquid leaving (T/F)
    "time": "AC",          # 25. Device Time (s)
    "dq": "AD",            # 26. Disqualify (T/F)
}

OUT_COLS = {
    "status": "BH", "tier": "BI", "touches": "BJ",
    "setup": "BK", "asl": "BL", "start": "BM", "actions": "BN", "final": "BO",
    "time": "BP", "water": "BQ", "noadj": "BR", "size": "BS",
    "p_over": "BT", "p_dim": "BU", "p_open": "BV", "p_touch": "BW",
    "p_solid": "BX", "p_elec": "BY", "pen": "BZ",
    "score": "CA", "srank": "CB", "trank": "CC", "tb1r": "CD",
    "tb2": "CE", "tb2r": "CF", "rank_tb": "CG", "rank_diff": "CH",
    "e_tt": "CI", "e_dim": "CJ", "e_start": "CK", "e_act": "CL",
    "e_water": "CM", "e_touch": "CN", "e_time": "CO",
    "errors": "CR",
    # Breakdown: Set-Up & ASL, Actions (start + scorable + final), Time Score,
    # Bonus (water timer + no adjustments), Device Size, Penalties. Sums to Score.
    "bd_setup": "CS", "bd_actions": "CT", "bd_time": "CU", "bd_bonus": "CV",
    "bd_size": "CW", "bd_pen": "CX",
    "exp_score": "CY", "exp_tier": "CZ", "exp_tiebreak": "DA", "exp_rank": "DB",
    "points": "DC",
}


SPEC = SheetSpec(INPUT_COLS, OUT_COLS, first_row=9, last_row=508)

TT60 = {TARGET_TIME: 60}

MSG_TT = "ERROR: Set Target Time in cell D7"
MSG_TT_TEXT = "ERROR: Target Time in cell D7 must be a number"
T, F = True, False     # per-row input check flags

# The base team against a 60 s Target Time:
#   Set-Up 50 + ASL 100 + Start 25 + 5 actions 250 + Final 250 + Time 120
#   + No Adjustments 75 + Device Size 3 x 10 = 900, no penalties.
BASE = 900

BASE_INPUTS = dict(
    eye="T", safe="T", d1=70, d2=70, d3=70, setup30="T",
    asl_on_time="T", asl_format="T", asl_accurate="T", asl_labelled="T",
    start=1, actions=5, final="T", time=60,
)


def team(school, **kw):
    """A Tier 1 team scoring BASE at a Regional (box 1 blank). Override any
    box, or pass box=None to leave it blank."""
    row = {"school": school}
    row.update(BASE_INPUTS)
    row.update(kw)
    return {k: v for k, v in row.items() if v is not None}


def bare(school, **kw):
    """Only the boxes given: bare("X") is a No-Show."""
    row = {"school": school}
    row.update(kw)
    return row


def statuses_of(*rows):
    return {school: dict(status=s) for school, s in rows}


# --------------------------------------------------------------------------
# Scenarios
# --------------------------------------------------------------------------

AWARD_FIELDS = ("setup", "asl", "start", "actions", "final", "time", "water",
                "noadj", "size", "touches", "pen", "score", "status", "tier")

SCORING = Scenario(
    "scoring",
    "Each award: Set-Up, ASL items, Start Action coins and Touch, scorable "
    "actions, Final Action, Water Timer (capped at the Target Time).",
    teams=[
        team("Base"),
        team("Setup15", setup15="T"),
        team("Setup15Only", setup30=None, setup15="T"),
        team("NoSetup", setup30="F"),
        team("NoASL", asl_on_time="F", asl_format=None, asl_accurate=None, asl_labelled=None),
        team("TwoASL", asl_format="F", asl_labelled="F"),
        team("Start2", start=2),
        team("Start3", start=3),
        team("Start4", start=4),
        team("StartText2", start="2"),
        team("StartTouch", start="Touch"),
        team("Actions12", actions=12),
        team("Actions0", actions=0),
        team("ActionsBlank", actions=None),
        team("NoFinal", final="F"),
        team("FinalBlank", final=None),
        team("Water", water=20.9),
        team("WaterCapped", water=75),
    ],
    expect=table(AWARD_FIELDS, {
        "Base":         (50, 100, 25, 250, 250, 120, 0, 75, 30, 0, 0, BASE, "C", 1),
        "Setup15":      (75, 100, 25, 250, 250, 120, 0, 75, 30, 0, 0, 925, "C", 1),
        "Setup15Only":  (75, 100, 25, 250, 250, 120, 0, 75, 30, 0, 0, 925, "C", 1),
        "NoSetup":      (0, 100, 25, 250, 250, 120, 0, 75, 30, 0, 0, 850, "C", 1),
        "NoASL":        (50, 0, 25, 250, 250, 120, 0, 75, 30, 0, 0, 800, "C", 1),
        "TwoASL":       (50, 50, 25, 250, 250, 120, 0, 75, 30, 0, 0, 850, "C", 1),
        "Start2":       (50, 100, 50, 250, 250, 120, 0, 75, 30, 0, 0, 925, "C", 1),
        "Start3":       (50, 100, 200, 250, 250, 120, 0, 75, 30, 0, 0, 1075, "C", 1),
        "Start4":       (50, 100, 0, 250, 250, 120, 0, 75, 30, 0, 0, 875, "C", 1),
        "StartText2":   (50, 100, 50, 250, 250, 120, 0, 75, 30, 0, 0, 925, "C", 1),
        # a touched start: no start points, one touch (-25), no 75 bonus
        "StartTouch":   (50, 100, 0, 250, 250, 120, 0, 0, 30, 1, 25, 775, "C", 1),
        "Actions12":    (50, 100, 25, 600, 250, 120, 0, 75, 30, 0, 0, 1250, "C", 1),
        "Actions0":     (50, 100, 25, 0, 250, 120, 0, 75, 30, 0, 0, 650, "C", 1),
        "ActionsBlank": (50, 100, 25, 0, 250, 120, 0, 75, 30, 0, 0, 650, "C", 1),
        "NoFinal":      (50, 100, 25, 250, 0, 120, 0, 75, 30, 0, 0, 650, "C", 1),
        "FinalBlank":   (50, 100, 25, 250, 0, 120, 0, 75, 30, 0, 0, 650, "C", 1),
        "Water":        (50, 100, 25, 250, 250, 120, 20, 75, 30, 0, 0, 920, "C", 1),
        "WaterCapped":  (50, 100, 25, 250, 250, 120, 60, 75, 30, 0, 0, 960, "C", 1),
    }),
    cells=TT60,
)

TIME_FIELDS = ("time", "final", "p_over", "pen", "score")

TIME_SCORE = Scenario(
    "time_score",
    "Time Score up to the Target Time, overtime penalty up to 2x, zero time "
    "points at 2x, and a touch leading to the Final Action.",
    teams=[
        team("Under", time=45.7),
        team("OnTarget", time=60),
        team("JustOver", time=60.9),
        team("Over", time=75.9),
        team("JustUnderDouble", time=119.9),
        team("Double", time=120),
        team("PastDouble", time=130),
        team("Zero", time=0),
        team("TouchedFinal", no_touch_final="F"),
        team("TouchedFinalOver", no_touch_final="F", time=70),
    ],
    expect=table(TIME_FIELDS, {
        "Under":            (90, 250, 0, 0, 870),
        "OnTarget":         (120, 250, 0, 0, 900),
        "JustOver":         (120, 250, 0, 0, 900),
        "Over":             (120, 250, 30, 30, 870),
        "JustUnderDouble":  (120, 250, 118, 118, 782),
        "Double":           (0, 250, 120, 120, 660),
        "PastDouble":       (0, 250, 120, 120, 660),
        "Zero":             (0, 250, 0, 0, 780),
        "TouchedFinal":     (0, 0, 0, 0, 530),
        "TouchedFinalOver": (0, 0, 20, 20, 510),
    }),
    cells=TT60,
)

STATE_TARGET = Scenario(
    "state_target",
    "A 90 s State Target Time: Time Score, overtime and Water Timer cap all "
    "follow D7.",
    teams=[
        team("Sixty"),
        team("Over", time=95.5),
        team("Water", water=100),
    ],
    expect=table(("time", "p_over", "water", "score"), {
        "Sixty": (120, 0, 0, 900),
        "Over":  (180, 10, 0, 950),
        "Water": (120, 0, 90, 990),
    }),
    cells={TARGET_TIME: 90},
)

DEVICE_SIZE = Scenario(
    "device_size",
    "0.1 point per 0.1 cm under 80 cm, max 30 per dimension; 25 per dimension "
    "over 80 cm; Tier 2 only over 82 cm.",
    teams=[
        team("Small", d1=50, d2=50, d3=50),
        team("Tiny", d1=40, d2=45, d3=49.9),
        team("Fraction", d1=79.75, d2=79.7, d3=70),
        team("AtLimit", d1=80, d2=80, d3=80),
        team("Over", d1=80.5),
        team("At82", d1=82),
        team("Tier2", d1=82.1),
        team("AllOver", d1=81, d2=81, d3=81),
    ],
    expect=table(("size", "p_dim", "tb2", "tier", "score"), {
        "Small":    (90, 0, 150, 1, 960),
        "Tiny":     (90, 0, 134.9, 1, 960),
        "Fraction": (10.5, 0, 229.45, 1, 880.5),
        "AtLimit":  (0, 0, 240, 1, 870),
        "Over":     (20, 25, 220.5, 1, 865),
        "At82":     (20, 25, 222, 1, 865),
        "Tier2":    (20, 25, 222.1, 2, 865),
        "AllOver":  (0, 75, 243, 1, 795),
    }),
    cells=TT60,
)

PEN_FIELDS = ("p_over", "p_dim", "p_open", "p_touch", "p_solid", "p_elec",
              "pen", "touches", "noadj", "start", "score")

PENALTIES = Scenario(
    "penalties",
    "Each penalty, touches capped at 3 including a touched start, and all of "
    "them at once.",
    teams=[
        team("Walls", walls="F"),
        team("Touches2", touches=2),
        team("Touches3", touches=3),
        team("TouchStartPlus1", start="Touch", touches=1),
        team("TouchStartPlus3", start="Touch", touches=3),
        team("Solid", solid="F"),
        team("Electric", elec="F"),
        team("Everything", walls="F", touches=1, solid="F", elec="F", time=70, d1=81),
    ],
    expect=table(PEN_FIELDS, {
        "Walls":           (0, 0, 25, 0, 0, 0, 25, 0, 75, 25, 875),
        "Touches2":        (0, 0, 0, 50, 0, 0, 50, 2, 0, 25, 775),
        "Touches3":        (0, 0, 0, 75, 0, 0, 75, 3, 0, 25, 750),
        "TouchStartPlus1": (0, 0, 0, 50, 0, 0, 50, 2, 0, 0, 750),
        "TouchStartPlus3": (0, 0, 0, 75, 0, 0, 75, 3, 0, 0, 725),
        "Solid":           (0, 0, 0, 0, 50, 0, 50, 0, 75, 25, 850),
        "Electric":        (0, 0, 0, 0, 0, 250, 250, 0, 75, 25, 650),
        # 815 in awards (no 75 bonus; size 0 + 10 + 10) less 20+25+25+25+50+250
        "Everything":      (20, 25, 25, 25, 50, 250, 395, 1, 0, 25, 420),
    }),
    cells=TT60,
)

TIERS_EXPECT = table(("tier", "score", "srank", "trank", "rank_tb", "rank_diff", "points",
                      "exp_tier", "exp_rank"), {
    "T1Low":     (1, 650, 6, 1, 1, 0, 1, 1, 1),
    "T2High":    (2, 1250, 1, 2, 2, 0, 2, 2, 2),
    "T2Dim":     (2, 865, 4, 3, 3, 0, 3, 2, 3),
    "T3Impound": (3, 1250, 1, 4, 4, 0, 4, 3, 4),
    "T3Answer":  (3, 900, 3, 5, 5, 0, 5, 3, 5),
    "T3AndT2":   (3, 850, 5, 6, 6, 0, 6, 3, 6),
    "T3Eye":     (3, 60, 7, 7, 7, 0, 7, 3, 7),
    "T3EyeOver": (3, 15, 8, 8, 8, 0, 8, 3, 8),
})
# Without eye protection a team is scored on device size alone, so every other
# component stays 0 -- and an oversized device is still penalised.
TIERS_EXPECT["T3Eye"].update(setup=0, start=0, actions=0, time=0, noadj=0, size=60)
TIERS_EXPECT["T3EyeOver"].update(p_dim=25, errors="")

TIERS = Scenario(
    "tiers",
    "Tier 3 (impound, eye protection, build questions) over Tier 2 "
    "(construction, > 82 cm); tiers rank before score; no eye protection "
    "scores device size only.",
    teams=[
        team("T1Low", actions=0),
        team("T2High", const="F", actions=12),
        team("T2Dim", d1=83),
        team("T3Impound", impound="F", actions=12),
        team("T3Answer", answer="F"),
        team("T3AndT2", impound="F", const="F", actions=4),
        # run boxes filled in anyway: ignored, the team did not run
        team("T3Eye", eye="F", d1=60, d2=60, d3=60, start=3, actions=12),
        bare("T3EyeOver", eye="F", d1=81, d2=60, d3=60),
    ],
    expect=TIERS_EXPECT,
    cells=TT60,
)

TIEBREAKS = Scenario(
    "tiebreaks",
    "Score, then fewest penalty points, then smallest L+W+H; a genuine "
    "unbroken tie shares a rank.",
    teams=[
        team("Higher", d1=50, d2=50, d3=50, actions=6, walls="F"),
        team("Sum140", d1=40, d2=50, d3=50),
        team("Pen0", d1=50, d2=50, d3=50),
        team("TieB", d1=50, d2=50, d3=50),
        team("Pen25Small", d1=30, d2=30, d3=30, walls="F", setup15="T"),
        team("Pen25", d1=50, d2=50, d3=50, walls="F", setup15="T"),
    ],
    expect=table(("score", "pen", "tb2", "srank", "tb1r", "tb2r", "rank_tb", "rank_diff",
                  "exp_tiebreak", "points"), {
        "Higher":     (985, 25, 150, 1, 4, 3, 1, 0, 0, 1),
        "Sum140":     (960, 0, 140, 2, 1, 2, 2, 0, 0, 2),
        "Pen0":       (960, 0, 150, 2, 1, 3, 3, -1, -1, 3),
        "TieB":       (960, 0, 150, 2, 1, 3, 3, -1, -1, 3),
        "Pen25Small": (960, 25, 90, 2, 4, 1, 5, -3, -3, 5),
        "Pen25":      (960, 25, 150, 2, 4, 3, 6, -4, -4, 6),
    }),
    cells=TT60,
)

STATUSES = Scenario(
    "statuses",
    "C / DQ / NS / P (unsafe, remote-controlled, impounded only) / ERR, their "
    "points, a blank row, and the final rankings list.",
    teams=[
        team("Competed"),
        team("Second", actions=4),
        team("Disq", dq="T"),
        bare("NoShow"),
        team("Unsafe", safe="F"),
        team("Remote", remote="T"),
        bare("ImpoundOnly", impound="T", asl_on_time="T"),
        team("Errored", time=None),
        team("Third", actions=3),
    ],
    expect=table(("status", "tier", "exp_score", "exp_tier", "exp_rank", "points"), {
        "Competed":    ("C", 1, BASE, 1, 1, 1),
        "Second":      ("C", 1, 850, 1, 2, 2),
        "Disq":        ("DQ", "DQ", "DQ", "DQ", "DQ", 11),
        "NoShow":      ("NS", "NS", "NS", "NS", "NS", 10),
        "Unsafe":      ("P", "P", "P", "P", "P", 9),
        "Remote":      ("P", "P", "P", "P", "P", 9),
        "ImpoundOnly": ("P", "P", "P", "P", "P", 9),
        "Errored":     ("ERR", "ERR", "ERR", "ERR", "ERR", "ERR"),
        "Third":       ("C", 1, 800, 1, 3, 3),
    }),
    cells=TT60,
)

INPUT_ERRORS = Scenario(
    "input_errors",
    "Each input check flags its row ERR with no Errors message (the input "
    "cell turns red instead); two at once; a clean row; checks that don't "
    "apply to a team that did not run; DQ/P win over ERR.",
    teams=[
        team("Clean"),
        team("NoDim", d1=None),
        team("TextDim", d2="abc"),
        team("NegDim", d3=-5),
        team("NoStart", start=None),
        team("BadStart", start=5),
        team("BadStartText", start="Coin"),
        team("TooManyActions", actions=13),
        team("FracActions", actions=2.5),
        team("NegWater", water=-1),
        team("TextWater", water="ten"),
        team("FracTouches", touches=1.5),
        team("TooManyTouches", touches=4),
        team("NoTime", time=None),
        team("TextTime", time="fast"),
        team("TwoErrors", d1=None, time=None),
        bare("EyeNoRun", eye="F", d1=60, d2=60, d3=60),
        bare("EyeNoDims", eye="F"),
        team("DQWithError", dq="T", time=None),
        team("PWithError", safe="F", time=None),
    ],
    expect=table(("status", "errors", "e_dim", "e_start", "e_act", "e_water", "e_touch", "e_time",
                  "exp_score", "exp_rank", "points"), {
        "Clean":         ("C", "", F, F, F, F, F, F, BASE, 1, 1),
        "NoDim":         ("ERR", "", T, F, F, F, F, F, "ERR", "ERR", "ERR"),
        "TextDim":       ("ERR", "", T, F, F, F, F, F, "ERR", "ERR", "ERR"),
        "NegDim":        ("ERR", "", T, F, F, F, F, F, "ERR", "ERR", "ERR"),
        "NoStart":       ("ERR", "", F, T, F, F, F, F, "ERR", "ERR", "ERR"),
        "BadStart":      ("ERR", "", F, T, F, F, F, F, "ERR", "ERR", "ERR"),
        "BadStartText":  ("ERR", "", F, T, F, F, F, F, "ERR", "ERR", "ERR"),
        "TooManyActions":("ERR", "", F, F, T, F, F, F, "ERR", "ERR", "ERR"),
        "FracActions":   ("ERR", "", F, F, T, F, F, F, "ERR", "ERR", "ERR"),
        "NegWater":      ("ERR", "", F, F, F, T, F, F, "ERR", "ERR", "ERR"),
        "TextWater":     ("ERR", "", F, F, F, T, F, F, "ERR", "ERR", "ERR"),
        "FracTouches":   ("ERR", "", F, F, F, F, T, F, "ERR", "ERR", "ERR"),
        "TooManyTouches":("ERR", "", F, F, F, F, T, F, "ERR", "ERR", "ERR"),
        "NoTime":        ("ERR", "", F, F, F, F, F, T, "ERR", "ERR", "ERR"),
        "TextTime":      ("ERR", "", F, F, F, F, F, T, "ERR", "ERR", "ERR"),
        "TwoErrors":     ("ERR", "", T, F, F, F, F, T, "ERR", "ERR", "ERR"),
        "EyeNoRun":      ("C", "", F, F, F, F, F, F, 60, 2, 2),
        "EyeNoDims":     ("ERR", "", T, F, F, F, F, F, "ERR", "ERR", "ERR"),
        # the flag is still raised, but DQ / P wins over ERR
        "DQWithError":   ("DQ", "", F, F, F, F, F, T, "DQ", "DQ", 22),
        "PWithError":    ("P", "", F, F, F, F, F, T, "P", "P", 20),
    }),
    cells=TT60,
)

CONTAINMENT = Scenario(
    "containment",
    "Errored rows are not scored or ranked, take no rank slot, and leave "
    "everyone else's score, rank and final ordering intact.",
    teams=[
        team("Broken1", actions=12, time=None),
        team("A", actions=7),
        team("Broken2", d1=None, d2=None, d3=None),
        team("B"),
        team("C", actions=4),
    ],
    expect=table(("status", "score", "srank", "rank_tb", "points"), {
        "Broken1": ("ERR", "ERR", "ERR", "ERR", "ERR"),
        "A":       ("C", 1000, 1, 1, 1),
        "Broken2": ("ERR", "ERR", "ERR", "ERR", "ERR"),
        "B":       ("C", BASE, 2, 2, 2),
        "C":       ("C", 850, 3, 3, 3),
    }),
    cells=TT60,
)

TT_TEAMS = [
    team("Runner"),
    bare("EyeF", eye="F", d1=60, d2=60, d3=60),
    team("Disq", dq="T"),
    team("Unsafe", safe="F"),
]

TARGET_TIME_MISSING = Scenario(
    "target_time_missing",
    "No Target Time: every team that ran is ERR with the D7 message; a team "
    "without eye protection still scores; DQ / P unaffected.",
    teams=TT_TEAMS,
    expect=table(("status", "errors", "score", "points"), {
        "Runner": ("ERR", MSG_TT, "ERR", "ERR"),
        "EyeF":   ("C", "", 60, 1),
        "Disq":   ("DQ", "", "DQ", 6),
        "Unsafe": ("P", "", "P", 4),
    }),
    cells={TARGET_TIME: None},
)

TARGET_TIME_TEXT = Scenario(
    "target_time_text",
    "A non-numeric Target Time gets its own message.",
    teams=TT_TEAMS,
    expect=table(("status", "errors", "score"), {
        "Runner": ("ERR", MSG_TT_TEXT, "ERR"),
        "EyeF":   ("C", "", 60),
        "Disq":   ("DQ", "", "DQ"),
        "Unsafe": ("P", "", "P"),
    }),
    cells={TARGET_TIME: "sixty"},
)

BREAKDOWN = Scenario(
    "breakdown",
    "The six breakdown columns add up to the Score; blank for DQ and ERR.",
    teams=[
        team("Full", setup15="T", start=3, actions=12, time=75, water=30, touches=1,
             walls="F", d1=50, d2=60, d3=81),
        team("Disq", dq="T"),
        team("Err", time=None),
        bare("EyeF", eye="F", d1=60, d2=60, d3=60),
    ],
    expect=table(("bd_setup", "bd_actions", "bd_time", "bd_bonus", "bd_size", "bd_pen",
                  "score"), {
        # 75+100 | 200+600+250 | 120 | 30+0 | 30+20+0 | -(30+25+25+25)
        "Full": (175, 1050, 120, 30, 50, -105, 1320),
        "Disq": ("", "", "", "", "", "", "DQ"),
        "Err":  ("", "", "", "", "", "", "ERR"),
        "EyeF": (0, 0, 0, 0, 60, 0, 60),
    }),
    cells=TT60,
)

SCENARIOS = [
    SCORING, TIME_SCORE, STATE_TARGET, DEVICE_SIZE, PENALTIES, TIERS, TIEBREAKS,
    STATUSES, INPUT_ERRORS, CONTAINMENT, TARGET_TIME_MISSING, TARGET_TIME_TEXT,
    BREAKDOWN,
]

if __name__ == "__main__":
    sys.exit(main(DEFAULT_SHEET, SPEC, SCENARIOS))
