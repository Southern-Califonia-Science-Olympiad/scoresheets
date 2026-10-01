#!/usr/bin/env python3
"""Targeted regression tests for the Roller Coaster B scoresheet.

Each scenario builds a self-contained mock tournament, recalculates the real
workbook with LibreOffice, and asserts on named output columns.

    python3 test_scoresheet.py [path/to/scoresheet.xlsx] [--only SCENARIO]

Expectations encode INTENDED behaviour. A failure means the sheet disagrees
with the rules as specified -- not that the test needs adjusting to match.

Rules under test (2027 B rules, section 5; high score wins):
  - Run Score = Size Score + Time Score + Elevator Score + Loop Score.
  - Size Score = (60 - height) + (30 - width) + (80 - length), to 0.1 cm.
  - Time Score = Time Bonus - Time Penalty: 5 per full second of Run Time up
    to the Target Time, less 5 per full second past it (up to 2 x Target
    Time). A run that does not cross the Finish Line has a Time Score of 0.
  - Elevator Score = 50 per Elevator used (0, 1 or 2).
  - Loop Score = 3 per whole cm of Loop height.
  - Final Score = the better run. Tiers: 1 no violations, 2 a construction or
    competition violation in the run, 3 not impounded. Tier ranks first.
  - Ties (5.i): Loop Score, Size Score, Time Score, then the longest Run Time,
    all of the scored run. The scored run is picked by lower tier, higher
    score, then the same tiebreaks.
  - No safety (box 2 False), or no run started: participation only (P).
  - A used run missing a needed measurement, an unusable value, or missing
    device dimensions, or a missing/invalid Target Time, makes the team ERR:
    not scored, ranked or listed. Input cells turn red; only the Target Time
    problems write a message (Errors column).
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "test-utils"))

from runner import Scenario, SheetSpec, main, table  # noqa: E402

HERE = Path(__file__).resolve().parent
DEFAULT_SHEET = HERE.parent / "scoresheet_b.xlsx"

# Input boxes, by the numbering printed in rows 2-4 of the Scoring sheet.
INPUT_COLS = {
    "team_no": "B", "school": "C", "team": "D",
    "imp": "E",        # 1. Impounded (T/F)
    "safe": "F",       # 2. Safety params met (T/F)
    "w": "G",          # 3. Device Width (cm)
    "l": "H",          # 4. Device Length (cm)
    "h": "I",          # 5. Device Height (cm)
    "dq": "V",         # 18. Disqualify (T/F)
}
# Per run (boxes 6-11 and 12-17): const params, comp params, Finish Line
# crossed, run time, elevators, loop height.
RUN_FIELDS = ["c", "p", "f", "t", "e", "lp"]
RUN_COLS = {1: "J K L M N O", 2: "P Q R S T U"}
for _n, _cols in RUN_COLS.items():
    for _field, _col in zip(RUN_FIELDS, _cols.split()):
        INPUT_COLS["%s%d" % (_field, _n)] = _col

OUT_COLS = {
    "dims_err": "AQ", "t_err": "AR",
    # run 1 / run 2: used, error, size, time, elevator, loop, tier, score, run time
    "r1_used": "AS", "r1_err": "AT", "r1_size": "AU", "r1_time": "AV", "r1_elev": "AW",
    "r1_loop": "AX", "r1_tier": "AY", "r1": "AZ", "r1_rt": "BA",
    "scored": "BB",
    "r2_used": "BC", "r2_err": "BD", "r2_size": "BE", "r2_time": "BF", "r2_elev": "BG",
    "r2_loop": "BH", "r2_tier": "BI", "r2": "BJ", "r2_rt": "BK",
    "status": "BL", "tier": "BM", "run_score": "BN", "final": "BO",
    "score_rank": "BP", "rank_ts": "BQ",
    # scored run's tiebreak values and their ranks
    "tb1": "BR", "tb1_rank": "BS", "tb2": "BT", "tb2_rank": "BU",
    "tb3": "BV", "tb3_rank": "BW", "tb4": "BX", "tb4_rank": "BY",
    "rank_tb": "BZ", "rank_diff": "CA",
    "errors": "CC",
    # breakdown of the scored run: run #, size, time, elevator, loop
    "bd_run": "CD", "bd_size": "CE", "bd_time": "CF", "bd_elev": "CG", "bd_loop": "CH",
    "exp_score": "CI", "exp_tier": "CJ", "exp_tiebreak": "CK", "exp_rank": "CL",
    "points": "CM",
}

SPEC = SheetSpec(INPUT_COLS, OUT_COLS, first_row=9, last_row=508)

TARGET = 40
CELLS = {"D7": TARGET}
MSG_MISSING = "ERROR: Set Target Time in cell D7"
MSG_INVALID = "ERROR: Target Time in cell D7 must be a whole number from 30 to 60"

# Default device: 25.0 wide x 70.0 long x 50.0 high -> Size Score 25.
BASE_SIZE = 25


def run(n, **kw):
    """Inputs for run n. Boxes left out (or None) stay blank."""
    return {"%s%d" % (k, n): v for k, v in kw.items() if v is not None}


def team(school, *runs, **kw):
    """A team that impounded, passed safety and measured the default device.
    Override with imp="F", h=None (leave blank), etc."""
    row = {"school": school, "imp": "T", "safe": "T", "w": 25, "l": 70, "h": 50}
    for r in runs:
        row.update(r)
    row.update(kw)
    return {k: v for k, v in row.items() if v is not None}


def bare(school):
    """No boxes at all: a No-Show."""
    return {"school": school}


def scen(name, why, teams, fields, rows, cells=CELLS):
    return Scenario(name, why, teams=teams, expect=table(fields, rows), cells=cells)


# --------------------------------------------------------------------------
# Scenarios
# --------------------------------------------------------------------------

RUN_SCORES = scen(
    "run_scores",
    "Run Score = Size + Time + Elevator + Loop; Loop counts whole cm; run 2 "
    "unused; the device limits themselves are legal.",
    teams=[
        team("Basic", run(1, t=40)),
        team("Elevators", run(1, t=30, e=2)),
        team("Loop", run(1, t=30, lp=12.7)),
        team("All", run(1, t=35, e=1, lp=20)),
        team("SmallDevice", run(1, t=40), w=20, l=60.5, h=45.3),
        team("AtLimits", run(1, t=40), w=30, l=80, h=60),
    ],
    fields=("r1_used", "r2_used", "r1_size", "r1_time", "r1_elev", "r1_loop", "r1_tier", "r1",
            "r2", "scored", "final", "exp_score", "status", "tier"),
    rows={
        "Basic":       (True, False, 25, 200, 0, 0, 1, 225, "", "<-", 225, 225, "C", 1),
        "Elevators":   (True, False, 25, 150, 100, 0, 1, 275, "", "<-", 275, 275, "C", 1),
        "Loop":        (True, False, 25, 150, 0, 36, 1, 211, "", "<-", 211, 211, "C", 1),
        "All":         (True, False, 25, 175, 50, 60, 1, 310, "", "<-", 310, 310, "C", 1),
        "SmallDevice": (True, False, 44.2, 200, 0, 0, 1, 244.2, "", "<-", 244.2, 244.2, "C", 1),
        "AtLimits":    (True, False, 0, 200, 0, 0, 1, 200, "", "<-", 200, 200, "C", 1),
    },
)

TIME_SCORE = scen(
    "time_score",
    "Target Time 40: 5 per full second up to the target, minus 5 per full "
    "second past it, never below 0 (timing ends at 2 x target).",
    teams=[
        team("OnTarget", run(1, t=40)),
        team("Early", run(1, t=12.9)),
        team("JustLate", run(1, t=40.9)),
        team("Late", run(1, t=45.5)),
        team("MaxLate", run(1, t=80)),
        team("BeyondMax", run(1, t=95)),
        team("ZeroTime", run(1, t=0)),
    ],
    fields=("r1_time", "r1", "r1_rt", "final"),
    rows={
        "OnTarget":  (200, 225, 40, 225),
        "Early":     (60, 85, 12.9, 85),
        "JustLate":  (200, 225, 40.9, 225),
        "Late":      (175, 200, 45.5, 200),
        "MaxLate":   (0, 25, 80, 25),
        "BeyondMax": (0, 25, 95, 25),
        "ZeroTime":  (0, 25, 0, 25),
    },
)

TARGET_60 = scen(
    "target_60",
    "The Target Time comes from D7: at 60, full marks are 300.",
    teams=[
        team("OnTarget", run(1, t=60)),
        team("Late30", run(1, t=90)),
        team("MaxLate", run(1, t=120)),
        team("Late5", run(1, t=65.5)),
    ],
    fields=("r1_time", "final"),
    rows={
        "OnTarget": (300, 325),
        "Late30":   (150, 175),
        "MaxLate":  (0, 25),
        "Late5":    (275, 300),
    },
    cells={"D7": 60},
)

TARGET_30 = scen(
    "target_30",
    "Lowest Target Time (30): on-target 150, past it loses 5 a second.",
    teams=[
        team("OnTarget", run(1, t=30)),
        team("Late15", run(1, t=45)),
        team("MaxLate", run(1, t=60)),
        team("Early", run(1, t=12)),
    ],
    fields=("r1_time", "final"),
    rows={
        "OnTarget": (150, 175),
        "Late15":   (75, 100),
        "MaxLate":  (0, 25),
        "Early":    (60, 85),
    },
    cells={"D7": 30},
)

FINISH_LINE = scen(
    "finish_line",
    "A run that does not cross the Finish Line has Time Score 0 but still "
    "scores size, Elevators and Loop; its Run Time is kept as the last "
    "tiebreak, and a missing time is fine. A blank box defaults to crossed.",
    teams=[
        team("NoFinish", run(1, f="F", t=12.3, e=1, lp=10)),
        team("NoFinishNoTime", run(1, f="F", e=1)),
        team("BlankFinish", run(1, t=20)),
        team("ExplicitFinish", run(1, f="T", t=20)),
    ],
    fields=("r1_time", "r1_elev", "r1_loop", "r1", "r1_rt", "r1_err", "r1_tier", "status"),
    rows={
        "NoFinish":       (0, 50, 30, 105, 12.3, False, 1, "C"),
        "NoFinishNoTime": (0, 50, 0, 75, 0, False, 1, "C"),
        "BlankFinish":    (100, 0, 0, 125, 20, False, 1, "C"),
        "ExplicitFinish": (100, 0, 0, 125, 20, False, 1, "C"),
    },
)

TIERS = scen(
    "tiers",
    "Tier 2 for a construction or competition violation or an oversize "
    "device, Tier 3 for not impounded (wins over 2); a worse tier ranks below "
    "a better one whatever the score.",
    teams=[
        team("T1Low", run(1, t=10)),
        team("T1Mid", run(1, t=20)),
        team("ConstViol", run(1, c="F", t=40, e=2)),
        team("CompViol", run(1, p="F", t=35)),
        team("OversizeH", run(1, t=15), h=60.1),
        team("OversizeW", run(1, t=10), w=30.1),
        team("OversizeL", run(1, t=10), l=80.1),
        team("NoImpound", run(1, t=40, e=2, lp=30), imp="F"),
        team("NoImpoundViol", run(1, p="F", t=10), imp="F"),
    ],
    fields=("r1_tier", "tier", "final", "score_rank", "rank_ts", "rank_tb", "points"),
    rows={
        "T1Low":         (1, 1, 75, 6, 2, 2, 2),
        "T1Mid":         (1, 1, 125, 4, 1, 1, 1),
        "ConstViol":     (2, 2, 325, 2, 3, 3, 3),
        "CompViol":      (2, 2, 200, 3, 4, 4, 4),
        "OversizeH":     (2, 2, 89.9, 5, 5, 5, 5),
        "OversizeW":     (2, 2, 69.9, 8, 6, 6, 6),
        "OversizeL":     (2, 2, 64.9, 9, 7, 7, 7),
        "NoImpound":     (3, 3, 415, 1, 8, 8, 8),
        "NoImpoundViol": (3, 3, 75, 6, 9, 9, 9),
    },
)

BEST_RUN = scen(
    "best_run",
    "The scored run has the lower tier, then the higher score, then the "
    "higher Loop, Time Score and Run Time; the breakdown follows it.",
    teams=[
        team("Run2Better", run(1, t=20), run(2, t=30)),
        team("Run1Better", run(1, t=30), run(2, t=20)),
        # run 1 scores more but is Tier 2, so the clean run 2 counts
        team("TierBeatsScore", run(1, c="F", t=40, e=2), run(2, t=20)),
        team("TierBeatsScore2", run(1, t=20), run(2, p="F", t=40, e=2)),
        team("OnlyRun2", run(2, t=20)),
        team("Identical", run(1, t=20), run(2, t=20)),
        # 225 each: Loop breaks it
        team("EqualLoop2", run(1, t=40), run(2, t=37, lp=5)),
        team("EqualLoop1", run(1, t=37, lp=5), run(2, t=40)),
        # 225 each, no Loop: Time Score breaks it (run 2 time 200 v 150)
        team("EqualTime", run(1, t=30, e=1), run(2, t=40)),
        # 75 each, neither crossed: the longer Run Time breaks it
        team("EqualRunTime", run(1, f="F", t=10, e=1), run(2, f="F", t=20, e=1)),
        # both Tier 2: plain higher score
        team("BothTier2", run(1, p="F", t=20), run(2, p="F", t=30)),
    ],
    fields=("r1", "r2", "r1_tier", "r2_tier", "scored", "final", "tier", "bd_run", "bd_size",
            "bd_time", "bd_elev", "bd_loop"),
    rows={
        "Run2Better":      (125, 175, 1, 1, "->", 175, 1, 2, 25, 150, 0, 0),
        "Run1Better":      (175, 125, 1, 1, "<-", 175, 1, 1, 25, 150, 0, 0),
        "TierBeatsScore":  (325, 125, 2, 1, "->", 125, 1, 2, 25, 100, 0, 0),
        "TierBeatsScore2": (125, 325, 1, 2, "<-", 125, 1, 1, 25, 100, 0, 0),
        "OnlyRun2":        ("", 125, "", 1, "->", 125, 1, 2, 25, 100, 0, 0),
        "Identical":       (125, 125, 1, 1, "<-", 125, 1, 1, 25, 100, 0, 0),
        "EqualLoop2":      (225, 225, 1, 1, "->", 225, 1, 2, 25, 185, 0, 15),
        "EqualLoop1":      (225, 225, 1, 1, "<-", 225, 1, 1, 25, 185, 0, 15),
        "EqualTime":       (225, 225, 1, 1, "->", 225, 1, 2, 25, 200, 0, 0),
        "EqualRunTime":    (75, 75, 1, 1, "->", 75, 1, 2, 25, 0, 50, 0),
        "BothTier2":       (125, 175, 2, 2, "->", 175, 2, 2, 25, 150, 0, 0),
    },
)

TIEBREAKS = scen(
    "tiebreaks",
    "Each tiebreak level deciding a tie (loser listed first): Loop, Size, "
    "Time Score, longest Run Time; and an unbroken tie.",
    teams=[
        # 275: Time Score (150 v 200)
        team("TB3Loser", run(1, t=30, e=2)),
        team("TB3Winner", run(1, t=40, e=1)),
        # 225: Loop (0 v 15)
        team("TB1Loser", run(1, t=40)),
        team("TB1Winner", run(1, t=37, lp=5)),
        # 175: Size (25 v 30)
        team("TB2Loser", run(1, t=30)),
        team("TB2Winner", run(1, t=29), w=20),
        # 125: identical
        team("TieA", run(1, t=20)),
        team("TieB", run(1, t=20)),
        # 75: neither crossed, Run Time (10 v 20)
        team("TB4Loser", run(1, f="F", t=10, e=1)),
        team("TB4Winner", run(1, f="F", t=20, e=1)),
    ],
    fields=("final", "tb1", "tb2", "tb3", "tb4", "score_rank", "rank_ts", "rank_tb",
            "rank_diff", "points"),
    rows={
        "TB3Loser":  (275, 0, 25, 150, 30, 1, 1, 2, -1, 2),
        "TB3Winner": (275, 0, 25, 200, 40, 1, 1, 1, 0, 1),
        "TB1Loser":  (225, 0, 25, 200, 40, 3, 3, 4, -1, 4),
        "TB1Winner": (225, 15, 25, 185, 37, 3, 3, 3, 0, 3),
        "TB2Loser":  (175, 0, 25, 150, 30, 5, 5, 6, -1, 6),
        "TB2Winner": (175, 0, 30, 145, 29, 5, 5, 5, 0, 5),
        "TieA":      (125, 0, 25, 100, 20, 7, 7, 7, 0, 7),
        "TieB":      (125, 0, 25, 100, 20, 7, 7, 7, 0, 7),
        "TB4Loser":  (75, 0, 25, 0, 10, 9, 9, 10, -1, 10),
        "TB4Winner": (75, 0, 25, 0, 20, 9, 9, 9, 0, 9),
    },
)

STATUSES = scen(
    "statuses",
    "C / DQ / NS / P / ERR: score, tier, rank and points. P is no safety or "
    "no run started.",
    teams=[
        team("Competitor", run(1, t=20)),
        team("Disqualified", run(1, t=20), dq="T"),
        bare("NoShow"),
        team("NoSafety", run(1, t=20), safe="F"),
        team("NoRuns"),
        team("Errored", run(1, e=1)),
    ],
    fields=("status", "exp_score", "tier", "exp_tier", "exp_rank", "points"),
    rows={
        "Competitor":   ("C", 125, 1, 1, 1, 1),
        "Disqualified": ("DQ", "DQ", "DQ", "DQ", "DQ", 8),
        "NoShow":       ("NS", "NS", "NS", "NS", "NS", 7),
        "NoSafety":     ("P", "P", "P", "P", "P", 6),
        "NoRuns":       ("P", "P", "P", "P", "P", 6),
        "Errored":      ("ERR", "ERR", "ERR", "ERR", "ERR", "ERR"),
    },
)

INPUT_ERRORS = scen(
    "input_errors",
    "A used run with a missing/unusable time, Elevator count or Loop height, "
    "or missing device dimensions, is ERR; clean rows (incl. a Finish Line "
    "miss with no time) are not; no message for any of these.",
    teams=[
        team("NoTime", run(1, e=1)),
        team("TextTime", run(1, t="x")),
        team("NegTime", run(1, t=-5)),
        team("FinishFalseTextTime", run(1, f="F", t="x")),
        team("Elev3", run(1, t=10, e=3)),
        team("ElevFraction", run(1, t=10, e=1.5)),
        team("ElevText", run(1, t=10, e="x")),
        team("LoopNegative", run(1, t=10, lp=-1)),
        team("LoopText", run(1, t=10, lp="x")),
        team("MissingHeight", run(1, t=10), h=None),
        team("TextWidth", run(1, t=10), w="x"),
        team("NegativeLength", run(1, t=10), l=-5),
        team("Run2Bad", run(1, t=10), run(2, e=1)),
        team("BoxOnly", run(1, c="T")),
        team("Clean", run(1, t=20)),
        team("FinishMiss", run(1, f="F", e=0)),
        # no upper limit on Loop height in the rules: a tall loop just scores
        team("LoopTall", run(1, t=10, lp=75)),
        # dimensions only matter once the team is actually competing
        team("NoDimsNoRun", w=None, l=None, h=None),
        team("NoDimsUnsafe", run(1, t=10), safe="F", h=None),
    ],
    fields=("status", "r1_err", "r2_err", "dims_err", "errors", "exp_score"),
    rows={
        "NoTime":              ("ERR", True, False, False, "", "ERR"),
        "TextTime":            ("ERR", True, False, False, "", "ERR"),
        "NegTime":             ("ERR", True, False, False, "", "ERR"),
        "FinishFalseTextTime": ("ERR", True, False, False, "", "ERR"),
        "Elev3":               ("ERR", True, False, False, "", "ERR"),
        "ElevFraction":        ("ERR", True, False, False, "", "ERR"),
        "ElevText":            ("ERR", True, False, False, "", "ERR"),
        "LoopNegative":        ("ERR", True, False, False, "", "ERR"),
        "LoopText":            ("ERR", True, False, False, "", "ERR"),
        "MissingHeight":       ("ERR", False, False, True, "", "ERR"),
        "TextWidth":           ("ERR", False, False, True, "", "ERR"),
        "NegativeLength":      ("ERR", False, False, True, "", "ERR"),
        "Run2Bad":             ("ERR", False, True, False, "", "ERR"),
        "BoxOnly":             ("ERR", True, False, False, "", "ERR"),
        "Clean":               ("C", False, False, False, "", 125),
        "FinishMiss":          ("C", False, False, False, "", 25),
        "LoopTall":            ("C", False, False, False, "", 300),
        "NoDimsNoRun":         ("P", False, False, False, "", "P"),
        "NoDimsUnsafe":        ("P", False, False, False, "", "P"),
    },
)

ERROR_CONTAINMENT = scen(
    "error_containment",
    "Errored teams are not scored or ranked, take no rank slot, and leave "
    "every other team's score and rank alone -- even with a score that would "
    "have topped the field.",
    teams=[
        team("Top", run(1, t=40)),
        team("Mid", run(1, t=30)),
        team("Low", run(1, t=20)),
        team("BigButBroken", run(1, t="x", e=2, lp=60)),
        team("NoTime", run(1, e=1)),
        team("NoDims", run(1, t=40, e=2), h=None),
    ],
    fields=("status", "exp_score", "score_rank", "rank_ts", "exp_rank", "points"),
    rows={
        "Top":          ("C", 225, 1, 1, 1, 1),
        "Mid":          ("C", 175, 2, 2, 2, 2),
        "Low":          ("C", 125, 3, 3, 3, 3),
        "BigButBroken": ("ERR", "ERR", "ERR", "ERR", "ERR", "ERR"),
        "NoTime":       ("ERR", "ERR", "ERR", "ERR", "ERR", "ERR"),
        "NoDims":       ("ERR", "ERR", "ERR", "ERR", "ERR", "ERR"),
    },
)

TARGET_MISSING = scen(
    "target_missing",
    "No Target Time in D7: a competing team is ERR with the message; teams "
    "that aren't scored anyway get none.",
    teams=[
        team("Competes", run(1, t=20)),
        team("NoSafety", run(1, t=20), safe="F"),
        team("NoRuns"),
        bare("NoShow"),
        team("Disqualified", run(1, t=20), dq="T"),
    ],
    fields=("status", "errors", "exp_score"),
    rows={
        "Competes":     ("ERR", MSG_MISSING, "ERR"),
        "NoSafety":     ("P", "", "P"),
        "NoRuns":       ("P", "", "P"),
        "NoShow":       ("NS", "", "NS"),
        "Disqualified": ("DQ", "", "DQ"),
    },
    cells={"D7": None},
)


def invalid_target(label, value):
    return scen(
        "target_invalid_" + label,
        "A Target Time of %r is not a whole number from 30 to 60." % (value,),
        teams=[team("Competes", run(1, t=20)), team("NoSafety", run(1, t=20), safe="F")],
        fields=("status", "errors", "exp_score"),
        rows={"Competes": ("ERR", MSG_INVALID, "ERR"), "NoSafety": ("P", "", "P")},
        cells={"D7": value},
    )


TARGET_BOUNDS = scen(
    "target_bounds",
    "30 and 60 are valid Target Times: no message, scored normally.",
    teams=[team("Ok", run(1, t=30))],
    fields=("status", "errors", "exp_score"),
    rows={"Ok": ("C", "", 175)},
    cells={"D7": 30},
)

SCENARIOS = [
    RUN_SCORES, TIME_SCORE, TARGET_60, TARGET_30, FINISH_LINE, TIERS, BEST_RUN, TIEBREAKS,
    STATUSES, INPUT_ERRORS, ERROR_CONTAINMENT, TARGET_MISSING,
    invalid_target("text", "abc"), invalid_target("low", 25), invalid_target("high", 61),
    invalid_target("fraction", 45.5), TARGET_BOUNDS,
]

if __name__ == "__main__":
    sys.exit(main(DEFAULT_SHEET, SPEC, SCENARIOS))
