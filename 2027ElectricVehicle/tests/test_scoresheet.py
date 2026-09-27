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
    Vehicle Distance, lower Bottle Pusher measurement -- except two Failed
    Runs with no competition or construction violations stay tied (5.k).
  - A Successful run missing its Run Time, Vehicle or Bottle Distance, or no
    Target Time, makes the team ERR: not scored, ranked or listed. Only the
    Target Time has an Errors-column message; a missing measurement is shown
    by its red input cell.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "test-utils"))

from runner import Scenario, SheetSpec, main, table  # noqa: E402

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
    "scored": "BM",
    "r2_err": "BN", "r2_ds": "BO", "r2_ts": "BP", "r2_bottle": "BQ",
    "r2_pusher": "BR", "r2_pen": "BS", "r2": "BT",
    # each run's own tiebreaks: DS, TS, vehicle distance, pusher measurement
    "r1_tb1": "BI", "r1_tb2": "BJ", "r1_tb3": "BK", "r1_tb4": "BL",
    "r2_tb1": "BU", "r2_tb2": "BV", "r2_tb3": "BW", "r2_tb4": "BX",
    "status": "BZ", "tier": "CA", "run_score": "CB", "final_pen": "CC",
    "etb": "CD", "final": "CE", "rank": "CF",
    "tb1": "CG", "tb2": "CI", "tb3": "CK", "tb4": "CM",
    "rank_tb": "CO", "rank_diff": "CP",
    "errors": "CS",
    # Breakdown of the scored run: run #, DS, TS, bonuses (bottle + pusher +
    # event time), penalties (run + impound). Final = 100 + DS + TS + B + P.
    "bd_run": "CT", "bd_ds": "CU", "bd_ts": "CV", "bd_bonus": "CW", "bd_pen": "CX",
    "exp_score": "CY", "exp_tier": "CZ", "exp_tiebreak": "DA", "exp_rank": "DB",
    "points": "DC",
}

LISTED = "DE8"          # teams on the final rankings list
LIST_SCHOOL = "DQ%d"    # final rankings: school, list starting on row 9

SPEC = SheetSpec(INPUT_COLS, OUT_COLS, first_row=9, last_row=508)

ERR_TT = "ERROR: Set Target Time in cell D7"
ERR_TT_TEXT = "ERROR: Target Time in cell D7 must be a number"

TT14 = {TARGET_TIME: 14}

# A Failed Run against a 14 s Target Time: 100 + 2500 + 0.5 x 14.
FAILED = 2607


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

RUN1 = ("r1_ds", "r1_ts", "r1_bottle", "r1_pusher", "r1_pen", "r1")
RUN2 = ("r2_ds", "r2_ts", "r2_bottle", "r2_pusher", "r2_pen", "r2")

RULES_EXAMPLE = Scenario(
    "rules_example",
    "Section 6 example from the web rules (Final Score 109.79).",
    teams=[
        team("Example",
             run(1, sf=S, time=12.27, vd=17.6, bd=10.6, bonus="F"),
             run(2, sf=S, time=16.37, vd=22.5, bd=12.1, bonus="T"),
             pusher=16),
    ],
    expect=table(RUN1 + RUN2 + ("scored", "run_score", "final", "status"), {
        "Example": (45.8, 0.865, 0, -28.5, 0, 118.165,
                    57.1, 1.185, -20, -28.5, 0, 109.785,
                    "->", 109.785, 109.785, "C"),
    }),
    cells=TT14,
)

FAILED_RUNS = Scenario(
    "failed_runs",
    "Failed Runs score DS 2500 and TS 0.5 x Target; no bonuses; missing runs "
    "fail; S/F inference; double-failed teams stay tied.",
    teams=[
        team("BothFailed", run(1, sf=F), run(2, sf=F), pusher=16),
        # Only check-in boxes: two Failed Runs, but competed.
        team("NoRuns"),
        # An explicit F wins over measurements typed in anyway.
        team("FailedWithData", run(1, sf=F, time=12, vd=5, bd=5, bonus="T")),
        # No S/F but measurements: Successful.
        team("InferredSuccess", run(1, time=14, vd=10, bd=10)),
        team("TiedA", run(1, sf=F), run(2, sf=F), pusher=10),
        team("TiedB", run(1, sf=F), run(2, sf=F), pusher=30),
        # Double-failed with violations: 5.k no longer holds them tied, so the
        # pusher measurement (TB4) decides even though no Pusher Bonus applies.
        team("FailedViolLoser", run(1, sf=F, comp="F"), run(2, sf=F, comp="F"), pusher=30),
        team("FailedViolWinner", run(1, sf=F, comp="F"), run(2, sf=F, comp="F"), pusher=10),
    ],
    expect=table(RUN1 + ("r2", "final", "tb4", "rank", "rank_tb", "errors"), {
        "BothFailed":       (2500, 7, 0, 0, 0, FAILED, FAILED, FAILED, "", 2, 2, ""),
        "NoRuns":           (2500, 7, 0, 0, 0, FAILED, FAILED, FAILED, "", 2, 2, ""),
        "FailedWithData":   (2500, 7, 0, 0, 0, FAILED, FAILED, FAILED, "", 2, 2, ""),
        "InferredSuccess":  (30, 0, 0, 0, 0, 130, FAILED, 130, 35, 1, 1, ""),
        "TiedA":            (2500, 7, 0, 0, 0, FAILED, FAILED, FAILED, "", 2, 2, ""),
        "TiedB":            (2500, 7, 0, 0, 0, FAILED, FAILED, FAILED, "", 2, 2, ""),
        "FailedViolLoser":  (2500, 7, 0, 0, 150, FAILED + 150, FAILED + 150, FAILED + 150,
                             30, 7, 8, ""),
        "FailedViolWinner": (2500, 7, 0, 0, 150, FAILED + 150, FAILED + 150, FAILED + 150,
                             10, 7, 7, ""),
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
    expect=table(("r1_ds", "r1_bottle", "r1", "r2_ds", "r2_bottle", "r2", "final", "status", "errors"), {
        "Past":          (30, 0, 130, 2500, 0, FAILED, 130, "C", ""),
        "PastBlank":     (30, 0, 130, 2500, 0, FAILED, 130, "C", ""),
        "NotPast":       (420, 0, 520, 2500, 0, FAILED, 520, "C", ""),
        "NotPastWithBd": (420, 0, 520, 2500, 0, FAILED, 520, "C", ""),
        "NotPastBonus":  (420, 0, 520, 2500, 0, FAILED, 520, "C", ""),
        "Run2NotPastBonus": (410, 0, 510, 420, 0, 520, 510, "C", ""),
        "Run2NotPast":   (2500, 0, FAILED, 420, 0, 520, 520, "C", ""),
        "PastMissingBd": ("", "", "", "", "", "", "ERR", "ERR", ""),
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
    expect=table(("r1_pen", "r1", "r2_pen", "r2", "scored", "final_pen", "final", "status"), {
        "Clean":      (0, 130, 0, FAILED, "<-", 0, 130, "C"),
        "Comp":       (150, 280, 0, FAILED, "<-", 0, 280, "C"),
        "Const":      (300, 430, 0, FAILED, "<-", 0, 430, "C"),
        "CompConst":  (450, 580, 0, FAILED, "<-", 0, 580, "C"),
        "Kit":        (50, 180, 50, FAILED + 50, "<-", 0, 180, "C"),
        "NoImpound":  (0, 130, 0, FAILED, "<-", 5000, 5130, "C"),
        "FailedComp": (150, FAILED + 150, 0, FAILED, "->", 0, FAILED, "C"),
        "Run2Comp":   (0, 130, 150, 280, "<-", 0, 130, "C"),
        "Lithium":    ("", "", "", "", "", "", "P", "P"),
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
    expect=table(("etb", "final", "rank"), {
        "TwoMinLeft":    (-4, 126, 2),
        "ThirtySecLeft": (-1, 129, 3),
        "NoneLeft":      (0, 130, 4),
        "OverEight":     (-16, 114, 1),
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
    expect=table(("final", "tb1", "tb2", "tb3", "tb4", "rank", "rank_tb", "rank_diff",
                  "exp_rank", "exp_tiebreak", "points"), {
        "TB1Loser":  (130, 30, 0, 10, 35, 1, 2, -1, 2, -1, 2),
        "TB1Winner": (130, 28, 2, 10, 35, 1, 1, 0, 1, 0, 1),
        "TB2Loser":  (186, 36, 50, 10, 35, 9, 10, -1, 10, -1, 10),
        "TB2Winner": (186, 36, 0, 13, 35, 9, 9, 0, 9, 0, 9),
        "TB3Loser":  (140, 40, 0, 15, 35, 3, 4, -1, 4, -1, 4),
        "TB3Winner": (140, 40, 0, 10, 35, 3, 3, 0, 3, 0, 3),
        "TB4Loser":  (150, 50, 0, 10, 35, 5, 6, -1, 6, -1, 6),
        "TB4Winner": (150, 50, 0, 10, 15, 5, 5, 0, 5, 0, 5),
        "TieA":      (160, 50 + 10, 0, 10, 35, 7, 7, 0, 7, 0, 7),
        "TieB":      (160, 50 + 10, 0, 10, 35, 7, 7, 0, 7, 0, 7),
    }),
    cells=TT14,
)

SCORED_RUN = Scenario(
    "scored_run",
    "The better Run Score counts; equal runs fall back to each run's own "
    "tiebreaks: lower DS, then TS, then Vehicle Distance, then pusher.",
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
    expect=table(("r1", "r2", "r1_tb1", "r1_tb2", "r1_tb3", "r1_tb4",
                  "r2_tb1", "r2_tb2", "r2_tb3", "r2_tb4",
                  "scored", "run_score", "tb1", "tb2", "tb3", "tb4"), {
        "Run2Better":    (140, 130, 40, 0, 10, 35, 30, 0, 10, 35, "->", 130, 30, 0, 10, 35),
        "Run1Better":    (130, 140, 30, 0, 10, 35, 40, 0, 10, 35, "<-", 130, 30, 0, 10, 35),
        "EqualLowerDS":  (140, 140, 40, 0, 10, 35, 30, 10, 10, 35, "->", 140, 30, 10, 10, 35),
        "Identical":     (130, 130, 30, 0, 10, 35, 30, 0, 10, 35, "<-", 130, 30, 0, 10, 35),
        "FailThenOK":    (FAILED, 130, 2500, 7, "", 35, 30, 0, 10, 35, "->", 130, 30, 0, 10, 35),
        "EqualLowerVD1": (140, 140, 40, 0, 10, 35, 40, 0, 15, 35, "<-", 140, 40, 0, 10, 35),
        "EqualLowerVD2": (140, 140, 40, 0, 15, 35, 40, 0, 10, 35, "->", 140, 40, 0, 10, 35),
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
    expect=table(("status", "exp_score", "tier", "exp_tier", "exp_rank", "points"), {
        "Competitor":   ("C", 130, 1, 1, 1, 1),
        "Disqualified": ("DQ", "DQ", "DQ", "DQ", "DQ", 7),
        "NoShow":       ("NS", "NS", "NS", "NS", "NS", 6),
        "Lithium":      ("P", "P", "P", "P", "P", 5),
        "Errored":      ("ERR", "ERR", "ERR", "ERR", "ERR", "ERR"),
    }),
    extra=[
        ("blank row status", "BZ{unused}", ""),
        ("blank row score", "CY{unused}", ""),
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
    expect=table(("status", "errors", "exp_score", "exp_rank", "points"), {
        "MissingBottle":   ("ERR", "", "ERR", "ERR", "ERR"),
        "SOnly":           ("ERR", "", "ERR", "ERR", "ERR"),
        "InferredPartial": ("ERR", "", "ERR", "ERR", "ERR"),
        "Negative":        ("ERR", "", "ERR", "ERR", "ERR"),
        "BothRuns":        ("ERR", "", "ERR", "ERR", "ERR"),
        "Run2Only":        ("ERR", "", "ERR", "ERR", "ERR"),
        "FailedEmpty":     ("C", "", FAILED, 3, 3),
        "Clean":           ("C", "", 130, 1, 1),
        "Clean2":          ("C", "", 140, 2, 2),
    }),
    extra=[
        ("errors header (no counter)", "CS8", "Errors"),
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
    expect=table(("status", "errors", "exp_rank", "points"), {
        "Competitor":   ("ERR", ERR_TT, "ERR", "ERR"),
        "Failed":       ("ERR", ERR_TT, "ERR", "ERR"),
        "Disqualified": ("DQ", "", "DQ", 7),
        "NoShow":       ("NS", "", "NS", 6),
        "Lithium":      ("P", "", "P", 5),
    }),
    cells={TARGET_TIME: None},
)

TARGET_TIME_TEXT = Scenario(
    "target_time_text",
    "A non-numeric Target Time has its own message.",
    teams=[team("Competitor", ok(1))],
    expect=table(("status", "errors"), {"Competitor": ("ERR", ERR_TT_TEXT)}),
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
    expect=table(("bd_run", "bd_ds", "bd_ts", "bd_bonus", "bd_pen", "final"), {
        "Example":      (2, 57.1, 1.185, -48.5, 0, 109.785),
        # run 1: 130 + 150 + 50 = 330; run 2: 140 + 50 = 190 -> run 2, + 5000
        "Penalised":    (2, 40, 0, 0, 5050, 5190),
        "TimeLeft":     (1, 30, 0, -2, 0, 128),
        # run 2 (2607) beats run 1 (2907)
        "DoubleFailed": (2, 2500, 7, 0, 0, FAILED),
        "Disqualified": ("", "", "", "", "", "DQ"),
        "Errored":      ("", "", "", "", "", "ERR"),
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
