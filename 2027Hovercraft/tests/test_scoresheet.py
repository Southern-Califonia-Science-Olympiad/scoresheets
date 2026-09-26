#!/usr/bin/env python3
"""Targeted regression tests for the Hovercraft B/C scoresheet.

Each scenario builds a self-contained mock tournament, recalculates the real
workbook with LibreOffice, and asserts on named output columns.

    python3 test_scoresheet.py [path/to/scoresheet.xlsx] [--only SCENARIO]

Expectations encode INTENDED behaviour. A failure means the sheet disagrees
with the rules as specified -- not that the test needs adjusting to match.

Rules under test (2027 B/C rules, section 7):
  - Run Score = DS + TS + MS; Final Score = best Run Score.
  - DS: complete 40; incomplete 40 x (185 - distance from finish) / 185.
    x 0.5 if the ramp was used.
  - TS: complete 40 x (1 - |time - TT| / TT), never below 0; incomplete 0.
  - MS: (4 x rolls + 0.1 x loose nickels) x the DS distance fraction, at most
    20; 0 if the ramp was used. Rolls cap at 4.5, loose nickels at 20.
  - DS, TS and MS are each x 0.7 for a missed impound, x 0.8 per construction
    violation and x 0.9 per competition violation in that run.
  - Touching the vehicle makes the run complete with DS = TS = MS = 0.
  - A team's time ends after two complete or two incomplete runs, so a third
    run only counts after one complete and one incomplete.
  - A blank C/I box is inferred: a time with no (or zero) distance is
    complete, a non-zero distance with no time is incomplete; otherwise the
    run cannot be typed and is not scored.
  - Failing the safety check (box 2 = F) is Participation only.
  - A run time with no Target Time set makes the team ERR.
  - Ties: best TS over all runs, then fewest construction violations in the
    worst run, then best MS, then 2nd best Run Score.
  - The event has no tiering: every team that competed is Tier 1.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "test-utils"))

from runner import Scenario, SheetSpec, main, table  # noqa: E402

HERE = Path(__file__).resolve().parent
DEFAULT_SHEET = HERE.parent / "scoresheet_bc.xlsx"

TARGET_TIME = "D8"

# Input boxes, by the numbering printed in rows 3-4 of the Scoring sheet.
INPUT_COLS = {
    "team_no": "B", "school": "C", "team": "D",
    "impound": "E",      # box 1  Impound (T/F)
    "safety": "F",       # box 2  All safety construction parameters met (T/F)
    "dq": "AH",          # box 12 Disqualify (T/F)
}
# Per run, boxes 3-11: construction violations, competition violations,
# no touching (T/F), ramp not used (T/F), C/I, nickel rolls, loose nickels,
# distance from finish (cm), run time (s).
RUN_FIELDS = ["const", "comp", "notouch", "noramp", "ci", "rolls", "nickels",
              "dist", "time"]
RUN_COLS = {1: "G H I J K L M N O", 2: "P Q R S T U V W X",
            3: "Y Z AA AB AC AD AE AF AG"}
for _n, _cols in RUN_COLS.items():
    for _field, _col in zip(RUN_FIELDS, _cols.split()):
        INPUT_COLS["%s%d" % (_field, _n)] = _col

# Computed columns worth asserting on.
OUT_COLS = {
    "r1_type": "AQ", "r2_type": "AY", "r3_type": "BG",
    "status": "BK", "tier": "BL",
    "r1_done": "BN", "r1_calc": "BO", "r1_mult": "BP",
    "r1_ds": "BR", "r1_ts": "BT", "r1_ms": "BV", "r1": "BW",
    "r2_mult": "BZ", "r2": "CG",
    "run3_on": "CH", "r3": "CR",
    "best": "CV", "score": "CW", "rank": "CX",
    "tb1": "CY", "tb1_rank": "CZ", "tb2": "DA", "tb2_rank": "DB",
    "tb3": "DC", "tb3_rank": "DD", "tb4": "DE", "tb4_rank": "DF",
    "rank_tb": "DG", "rank_diff": "DH",
    "errors": "DK",
    # Breakdown: the scored (best) run and its penalised DS, TS and MS.
    "bd_run": "DL", "bd_ds": "DM", "bd_ts": "DN", "bd_ms": "DO",
    # Export block (Tier is hidden: no tiering), and the final-rankings helper
    # that carries Tier.
    "exp_score": "DP", "exp_tier": "DQ", "exp_tiebreak": "DR", "exp_rank": "DS",
    "points": "DT", "tier_helper": "EC",
}

SPEC = SheetSpec(INPUT_COLS, OUT_COLS, first_row=11, last_row=510)

TT_ERROR = "ERROR: Set Target Time in cell D8"

# The breakdown columns, and their value for a team with no scored run.
BREAKDOWN = ("bd_run", "bd_ds", "bd_ts", "bd_ms")
NO_RUN = ("", "", "", "")


def run(n, const=0, comp=0, ci=None, time=None, dist=None, rolls=None,
        nickels=None, notouch=None, noramp=None):
    """Inputs for run n. Boxes left as None stay blank."""
    values = dict(const=const, comp=comp, ci=ci, time=time, dist=dist,
                  rolls=rolls, nickels=nickels, notouch=notouch, noramp=noramp)
    return {"%s%d" % (k, n): v for k, v in values.items() if v is not None}


def team(school, *runs, **kw):
    """A team that made impound and passed the safety check.

    `runs` are dicts from run(); pass impound="F", safety="F", dq="T" etc. to
    override, or impound=None to leave a box blank.
    """
    row = {"school": school, "impound": "T", "safety": "T"}
    for r in runs:
        row.update(r)
    row.update(kw)
    return {k: v for k, v in row.items() if v is not None}


def bare(school):
    """A team with no box inputs at all -- boxes 1-12 blank means No-Show."""
    return {"school": school}


def on_time(n, **kw):
    """Complete run exactly on a 10 s Target Time: DS 40 + TS 40 before MS."""
    return run(n, ci="C", time=10, **kw)


# --------------------------------------------------------------------------
# Scenarios
#
# Expectations are tables (runner.table): within a scenario every team is
# checked on the same columns, so a team can't pass by leaving one out.
# --------------------------------------------------------------------------

TT10 = {TARGET_TIME: 10}

SCORING = Scenario(
    "scoring",
    "Run score = DS + TS + MS: completion, time, distance, mass, ramp, touching.",
    teams=[
        # 4 rolls + 10 loose nickels: MS = 16 + 1 = 17.
        team("CompleteOnTime", on_time(1, rolls=4, nickels=10)),
        # 12 s against 10 s: TS = 40 x (1 - 0.2) = 32.
        team("CompleteOffTime", run(1, ci="C", time=12)),
        # 25 s is 150 % off: TS would be -20, floored at 0.
        team("TimeFloor", run(1, ci="C", time=25)),
        # Halfway (92.5 cm from the finish): DS 20, TS 0, MS 8 x 0.5 = 4.
        team("Incomplete", run(1, ci="I", dist=92.5, rolls=2)),
        # Ramp: DS x 0.5, MS = 0, TS untouched.
        team("Ramp", on_time(1, rolls=4, noramp="F")),
        # Touched after the start: complete, DS = TS = MS = 0.
        team("Touched", on_time(1, rolls=4, notouch="F")),
        # 4.5 rolls + 20 loose nickels: MS = 20.
        team("MaxScore", on_time(1, rolls=4.5, nickels=20)),
    ],
    expect=table(
        ("status", "r1_type", "r1_done", "r1_mult", "r1_ds", "r1_ts", "r1_ms",
         "r1", "best", "score", "rank", "points") + BREAKDOWN, {
            "CompleteOnTime": ("C", True, True, 1, 40, 40, 17, 97, 97, 97, 2, 2,
                               1, 40, 40, 17),
            "CompleteOffTime": ("C", True, True, 1, 40, 32, 0, 72, 72, 72, 3, 3,
                                1, 40, 32, 0),
            "TimeFloor": ("C", True, True, 1, 40, 0, 0, 40, 40, 40, 5, 5,
                          1, 40, 0, 0),
            "Incomplete": ("C", False, False, 1, 20, 0, 4, 24, 24, 24, 6, 6,
                           1, 20, 0, 4),
            "Ramp": ("C", True, True, 1, 20, 40, 0, 60, 60, 60, 4, 4,
                     1, 20, 40, 0),
            # A touched run is still the scored run, at 0.
            "Touched": ("C", True, True, 1, 0, 0, 0, 0, 0, 0, 7, 7,
                        1, 0, 0, 0),
            "MaxScore": ("C", True, True, 1, 40, 40, 20, 100, 100, 100, 1, 1,
                         1, 40, 40, 20),
        }),
    cells=TT10,
)

MAX_CAPS = Scenario(
    "max_caps",
    "Rolls cap at 4.5 and loose nickels at 20 before MS is computed.",
    teams=[
        # 6 rolls and 50 nickels clamp to 4.5 and 20: MS = 18 + 2 = 20.
        team("MassCap", on_time(1, rolls=6, nickels=50)),
    ],
    expect=table(
        ("status", "r1_type", "r1_mult", "r1_ds", "r1_ts", "r1_ms", "r1",
         "score", "rank", "rank_tb", "points") + BREAKDOWN, {
            "MassCap": ("C", True, 1, 40, 40, 20, 100, 100, 1, 1, 1,
                        1, 40, 40, 20),
        }),
    cells=TT10,
)

# Every team flies the CompleteOnTime run (40 + 40 + 17 = 97 unpenalised)
# unless stated; each component carries the whole multiplier.
PENALTIES = Scenario(
    "penalties",
    "Impound, construction and competition multipliers on DS/TS/MS; ramp on DS only.",
    teams=[
        team("Clean", on_time(1, rolls=4, nickels=10)),
        team("MissedImpound", on_time(1, rolls=4, nickels=10), impound="F"),
        team("OneConstruction", on_time(1, const=1, rolls=4, nickels=10)),
        team("TwoConstruction", on_time(1, const=2, rolls=4, nickels=10)),
        team("OneCompetition", on_time(1, comp=1, rolls=4, nickels=10)),
        team("AllThree", on_time(1, const=1, comp=1, rolls=4, nickels=10),
             impound="F"),
        # Ramp and a construction violation: DS 40 x 0.5 x 0.8, TS 40 x 0.8.
        team("RampAndConstruction",
             on_time(1, const=1, rolls=4, nickels=10, noramp="F")),
        # No cap on the violation count.
        team("FiveConstruction", on_time(1, const=5, rolls=4, nickels=10)),
        # Run 2's penalties come from its own boxes, not Run 1's.
        team("Run2OwnPenalties", on_time(1, const=1), on_time(2, comp=1)),
    ],
    # A blank run has no multiplier or score; tb2 is the most construction
    # violations in any run, tb4 the 2nd best run (0 with only one).
    expect=table(
        ("r1_mult", "r1_ds", "r1_ts", "r1_ms", "r1", "r2_mult", "r2", "tb2",
         "tb4", "score") + BREAKDOWN, {
            "Clean": (1, 40, 40, 17, 97, "", "", 0, 0, 97,
                      1, 40, 40, 17),
            "MissedImpound": (0.7, 28, 28, 11.9, 67.9, "", "", 0, 0, 67.9,
                              1, 28, 28, 11.9),
            "OneConstruction": (0.8, 32, 32, 13.6, 77.6, "", "", 1, 0, 77.6,
                                1, 32, 32, 13.6),
            "TwoConstruction": (0.64, 25.6, 25.6, 10.88, 62.08, "", "", 2, 0,
                                62.08, 1, 25.6, 25.6, 10.88),
            "OneCompetition": (0.9, 36, 36, 15.3, 87.3, "", "", 0, 0, 87.3,
                               1, 36, 36, 15.3),
            "AllThree": (0.504, 20.16, 20.16, 8.568, 48.888, "", "", 1, 0,
                         48.888, 1, 20.16, 20.16, 8.568),
            "RampAndConstruction": (0.8, 16, 32, 0, 48, "", "", 1, 0, 48,
                                    1, 16, 32, 0),
            "FiveConstruction": (0.32768, 13.1072, 13.1072, 5.57056, 31.78496,
                                 "", "", 5, 0, 31.78496,
                                 1, 13.1072, 13.1072, 5.57056),
            # 80 x 0.8 = 64 and 80 x 0.9 = 72: Run 2 is the scored run.
            "Run2OwnPenalties": (0.8, 32, 32, 0, 64, 0.9, 72, 1, 64, 72,
                                 2, 36, 36, 0),
        }),
    cells=TT10,
)

RUN3 = Scenario(
    "run3",
    "A third run counts only after one complete and one incomplete run.",
    teams=[
        # C, C: time is over; Run 3 (would be 100) is ignored.
        team("TwoComplete", run(1, ci="C", time=12), run(2, ci="C", time=11),
             on_time(3, rolls=4.5, nickels=20)),
        # C, I: Run 3 counts and is the best run.
        team("CompleteIncomplete", run(1, ci="C", time=12),
             run(2, ci="I", dist=92.5), on_time(3)),
        # I, C: order doesn't matter.
        team("IncompleteComplete", run(1, ci="I", dist=92.5), on_time(2),
             run(3, ci="C", time=11)),
        # I, I: time is over; Run 3 is ignored.
        team("TwoIncomplete", run(1, ci="I", dist=92.5),
             run(2, ci="I", dist=37), on_time(3)),
    ],
    expect=table(
        ("r1", "r2", "run3_on", "r3", "best", "score", "tb4") + BREAKDOWN, {
            "TwoComplete": (72, 76, False, "", 76, 76, 72,
                            2, 40, 36, 0),
            "CompleteIncomplete": (72, 20, True, 80, 80, 80, 72,
                                   3, 40, 40, 0),
            "IncompleteComplete": (20, 80, True, 76, 80, 80, 76,
                                   2, 40, 40, 0),
            # 37 cm from the finish: DS = 40 x 148 / 185 = 32.
            "TwoIncomplete": (20, 32, False, "", 32, 32, 20,
                              2, 32, 0, 0),
        }),
    cells=TT10,
)

RUN_TYPE = Scenario(
    "run_type",
    "C/I inferred from distance and time when blank; an untypeable run is not scored.",
    teams=[
        team("InferredComplete", run(1, time=10)),
        team("InferredIncomplete", run(1, dist=92.5)),
        # Both a time and a non-zero distance: can't tell, so no run.
        team("Ambiguous", run(1, time=10, dist=50)),
        # An explicit C/I wins over what the cells suggest.
        team("CIOverrides", run(1, ci="I", time=10, dist=92.5)),
        team("CompleteDistZero", run(1, ci="C", time=10, dist=0)),
        # Complete with no time: nothing to score.
        team("CompleteNoTime", run(1, ci="C", dist=50)),
        # Touched with nothing else entered: a complete run of 0.
        team("TouchedOnly", run(1, notouch="F")),
    ],
    expect=table(
        ("status", "r1_type", "r1_done", "r1_calc", "r1", "score")
        + BREAKDOWN, {
            "InferredComplete": ("C", True, True, True, 80, 80,
                                 1, 40, 40, 0),
            "InferredIncomplete": ("C", False, False, True, 20, 20,
                                   1, 20, 0, 0),
            "Ambiguous": ("C", "", "", False, "", 0) + NO_RUN,
            "CIOverrides": ("C", False, False, True, 20, 20,
                            1, 20, 0, 0),
            "CompleteDistZero": ("C", True, True, True, 80, 80,
                                 1, 40, 40, 0),
            "CompleteNoTime": ("C", True, True, False, "", 0) + NO_RUN,
            "TouchedOnly": ("C", "", True, True, 0, 0,
                            1, 0, 0, 0),
        }),
    cells=TT10,
)

# One pair per tiebreak level, each at its own score, loser listed first so
# row order can't be what separates them. 10 teams, so points = final rank.
TIEBREAKS = Scenario(
    "tiebreaks",
    "Score -> best TS -> fewest construction violations -> best MS -> 2nd best run, and an unbroken tie.",
    teams=[
        # 97 each, best TS 40 each; worst run has 2 vs 1 construction violations.
        team("TB2_Lose", on_time(1, rolls=4, nickels=10),
             on_time(2, const=2, rolls=4, nickels=10)),
        team("TB2_Win", on_time(1, rolls=4, nickels=10),
             on_time(2, const=1, rolls=4, nickels=10)),
        # 90 each, same TS, violations and MS; 2nd best run 64 vs 72.
        team("TB4_Lose", on_time(1, rolls=2.5), run(2, ci="C", time=14)),
        team("TB4_Win", on_time(1, rolls=2.5), run(2, ci="C", time=12)),
        # 84 each: TS 32 + MS 12 vs TS 40 + MS 4.
        team("TB1_Lose", run(1, ci="C", time=12, rolls=3)),
        team("TB1_Win", on_time(1, rolls=1)),
        # 80 each, TS 40 each, no violations; best MS (incomplete Run 2) 2 vs 4.
        team("TB3_Lose", on_time(1), run(2, ci="I", dist=92.5, rolls=1)),
        team("TB3_Win", on_time(1), run(2, ci="I", dist=92.5, rolls=2)),
        # 70 each and identical on every tiebreak.
        team("TrueTieA", run(1, ci="C", time=12.5)),
        team("TrueTieB", run(1, ci="C", time=12.5)),
    ],
    # tb4 is 0 for a team with only one run.
    expect=table(
        ("score", "rank", "tb1", "tb2", "tb3", "tb4", "rank_tb", "rank_diff",
         "points") + BREAKDOWN, {
            "TB2_Lose": (97, 1, 40, 2, 17, 62.08, 2, -1, 2, 1, 40, 40, 17),
            "TB2_Win": (97, 1, 40, 1, 17, 77.6, 1, 0, 1, 1, 40, 40, 17),
            "TB4_Lose": (90, 3, 40, 0, 10, 64, 4, -1, 4, 1, 40, 40, 10),
            "TB4_Win": (90, 3, 40, 0, 10, 72, 3, 0, 3, 1, 40, 40, 10),
            "TB1_Lose": (84, 5, 32, 0, 12, 0, 6, -1, 6, 1, 40, 32, 12),
            "TB1_Win": (84, 5, 40, 0, 4, 0, 5, 0, 5, 1, 40, 40, 4),
            "TB3_Lose": (80, 7, 40, 0, 2, 22, 8, -1, 8, 1, 40, 40, 0),
            "TB3_Win": (80, 7, 40, 0, 4, 24, 7, 0, 7, 1, 40, 40, 0),
            # Both keep rank 9 -- the sheet must not invent a split.
            "TrueTieA": (70, 9, 30, 0, 0, 0, 9, 0, 9, 1, 40, 30, 0),
            "TrueTieB": (70, 9, 30, 0, 0, 0, 9, 0, 9, 1, 40, 30, 0),
        }),
    cells=TT10,
)

# The example entered on the workbook itself (rows 11-12, Target Time 20):
# both teams' best run is 64 with TS 24, but the second team's other run had a
# construction violation, so TB2 puts the first team ahead.
SHEET_EXAMPLE = Scenario(
    "sheet_example",
    "The TB2 example entered on the scoresheet (Target Time 20).",
    teams=[
        # 12 s against 20 s: TS = 40 x (1 - 8/20) = 24, so 40 + 24 = 64.
        team("SheetRow11", run(1, ci="C", time=12)),
        # Run 1: 8 s is 12 s off, TS 16; x 0.8 -> 32 + 12.8 = 44.8.
        team("SheetRow12", run(1, const=1, ci="C", time=8),
             run(2, ci="C", time=12)),
    ],
    expect=table(
        ("r1_mult", "r1", "r2", "score", "rank", "tb1", "tb2", "tb4",
         "rank_tb", "rank_diff", "points") + BREAKDOWN, {
            "SheetRow11": (1, 64, "", 64, 1, 24, 0, 0, 1, 0, 1,
                           1, 40, 24, 0),
            "SheetRow12": (0.8, 44.8, 64, 64, 1, 24, 1, 44.8, 2, -1, 2,
                           2, 40, 24, 0),
        }),
    cells={TARGET_TIME: 20},
)

# 6 teams entered, so P = 6 points, NS = 7, DQ = 8.
STATUSES = Scenario(
    "statuses",
    "Status assignment, Tier 1 for competitors, DQ / NS / P points, and the final rankings list.",
    teams=[
        team("Winner", on_time(1, rolls=4, nickels=10)),
        team("Runner", on_time(1)),
        team("Disqualified", on_time(1, rolls=4.5, nickels=20), dq="T"),
        bare("NoShow"),
        # Failed the safety check: may not run, Participation only.
        team("SafetyFail", on_time(1), safety="F"),
        # Checked in but made no runs: competed, with a score of 0.
        team("NoRuns"),
    ],
    # Only competitors have a breakdown.
    expect=table(
        ("status", "tier", "tier_helper", "r1", "score", "exp_score",
         "exp_rank", "points") + BREAKDOWN, {
            "Winner": ("C", 1, 1, 97, 97, 97, 1, 1, 1, 40, 40, 17),
            "Runner": ("C", 1, 1, 80, 80, 80, 2, 2, 1, 40, 40, 0),
            "Disqualified": ("DQ", "DQ", "DQ", "", "", "DQ", "DQ", 8) + NO_RUN,
            "NoShow": ("NS", "NS", "NS", "", "", "NS", "NS", 7) + NO_RUN,
            "SafetyFail": ("P", "P", "P", "", "", "P", "P", 6) + NO_RUN,
            "NoRuns": ("C", 1, 1, "", 0, 0, 3, 3) + NO_RUN,
        }),
    # A row with no team must stay empty; the final rankings list (EF/EH/EJ)
    # fills from row 11 in points order, DQ/NS/P included.
    extra=[("BK{unused}", ""), ("BL{unused}", ""), ("CW{unused}", ""),
           ("DL{unused}", ""), ("DM{unused}", ""), ("DN{unused}", ""),
           ("DO{unused}", ""), ("DP{unused}", ""), ("DQ{unused}", ""),
           ("DT{unused}", ""),
           ("teams in final rankings", "DV10", 6),
           ("1st place in final rankings", "EF11", 1),
           ("1st place school", "EH11", "Winner"),
           ("1st place score", "EJ11", 97),
           ("2nd place in final rankings", "EF12", 2),
           ("3rd place in final rankings", "EF13", 3),
           ("then Participated", "EF14", "P"),
           ("then No-Show", "EF15", "NS"),
           ("Disqualified last", "EF16", "DQ"),
           ("nothing after last place", "EF17", "")],
    cells=TT10,
)

# No Target Time: a team with a run time can't be scored; one without can.
# 4 teams entered, so P = 4 points.
TARGET_TIME_MISSING = Scenario(
    "target_time",
    "A run time with no Target Time is ERR: unscored, unranked, off the final rankings.",
    teams=[
        team("TimedRun", on_time(1, rolls=4)),
        team("DistanceOnly", run(1, ci="I", dist=92.5, rolls=2)),
        # P outranks ERR: the vehicle never ran.
        team("SafetyFail", on_time(1), safety="F"),
        team("TimedRun2", run(1, ci="I", dist=92.5), run(2, ci="C", time=12)),
    ],
    expect=table(
        ("status", "tier", "tier_helper", "r1", "score", "exp_score",
         "exp_rank", "points", "errors") + BREAKDOWN, {
            "TimedRun": ("ERR", "ERR", "ERR", "", "", "ERR", "ERR", "ERR",
                         TT_ERROR) + NO_RUN,
            # Distance-only runs don't need the Target Time.
            "DistanceOnly": ("C", 1, 1, 24, 24, 24, 1, 1, "",
                             1, 20, 0, 4),
            "SafetyFail": ("P", "P", "P", "", "", "P", "P", 4, "") + NO_RUN,
            # One timed run is enough to make the whole team ERR.
            "TimedRun2": ("ERR", "ERR", "ERR", "", "", "ERR", "ERR", "ERR",
                          TT_ERROR) + NO_RUN,
        }),
    # ERR teams take no rank slot and are left off the final rankings.
    extra=[("teams in final rankings", "DV10", 2),
           ("1st place in final rankings", "EF11", 1),
           ("1st place school", "EH11", "DistanceOnly"),
           ("then Participated", "EF12", "P"),
           ("nothing after", "EF13", "")],
    cells={TARGET_TIME: None},
)

SCENARIOS = [SCORING, MAX_CAPS, PENALTIES, RUN3, RUN_TYPE, TIEBREAKS,
             SHEET_EXAMPLE, STATUSES, TARGET_TIME_MISSING]

# The export columns only display working ones: wherever a scenario asserts the
# working value, assert the same on the export, so the two cannot drift. Every
# team in a scenario has the same fields, so they gain the same mirrors.
for _sc in SCENARIOS:
    for _expect in _sc.expect.values():
        for _src, _dst in (("tier", "exp_tier"), ("rank_tb", "exp_rank"),
                           ("rank_diff", "exp_tiebreak")):
            if _src in _expect and _dst not in _expect:
                _expect[_dst] = _expect[_src]
        if "score" in _expect and "exp_score" not in _expect:
            _expect["exp_score"] = _expect["score"]


if __name__ == "__main__":
    sys.exit(main(DEFAULT_SHEET, SPEC, SCENARIOS))
