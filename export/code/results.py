"""The trial's results, from the master log: aggregates only, never a company name.

    python results.py --agent-runs data-exports/agent-runs/batch-v1 data-exports/agent-runs/batch-v2 [--master-log <file>]
    python results.py --agent-runs ... --public --md data-exports/public-export/RESULTS.md --no-json

Reads the Rows sheet of the newest master log (or --master-log) and, for the agent's cost, the run records in the
given batch folders. Writes data-exports/master-log/results-YYYY-MM-DD.json (never overwritten) and RESULTS.md beside
this script (aggregates only, safe to commit), and prints the same numbers.

The measures follow the hypothesis and its five falsification conditions (phase-0 decision record, discovery D5):
  - acceptable rate = the lead's Yes over the rows the lead scored, per arm;
  - SDR minutes per acceptable row = the SDR minutes on the scored rows whose stamps give minutes, over the
    acceptable rows among those same rows (an unreadable row leaves both sums; the arm's mean is put in for it as a check);
  - agent minutes per acceptable row the same way, for the agent's run behind each row (never added to SDR minutes);
  - "the pair" is every scored Pair row; v1 and v2 are shown apart;
  - the flag is scored on the SDR-alone companies against the SDR-alone verdicts the lead accepted; rows the lead did
    not accept are left out and counted; an existing active account never counts as a miss.
Small samples: every rate is given with its count. A paired sign test is printed for the same-company comparisons.
Also shown, each labelled: how the lead's score reads on dead-end rows (with the comments), a second reading on
prospect rows only (not set in advance), how far the unscored Pair rows could move condition 1, and keyword counts over
the lead's comments. Once the lead's dead-end re-check is filed (DECISIONS.md 18), a dead-end row counts as acceptable
when the lead judged the dead end right, and that reading comes first.

--public (DECISIONS.md 18): the same tables for the public export, with the SDR's side shown only over groups of 3 rows
or more (a smaller group shows "fewer than 3 rows"), no SDR minutes listed row by row, and no comment-derived counts on
single rows.
"""

import argparse
import json
import statistics
import sys
from collections import Counter
from datetime import datetime
from math import comb
from pathlib import Path

from openpyxl import load_workbook

from contract import (
    DEAD_END_ACTIVE, DIFFICULTIES, EARLY, LATE, MARKETS, RESEARCH_FIELDS, ROW_AGENT_ALONE, ROW_PAIR, ROW_SDR_ALONE,
    SDR_UNSURE_VALUES, VERDICT_DEAD_END,
)
from xlsx_util import next_version_path

HERE = Path(__file__).resolve().parent
CORRECTNESS = ["None", "One", "More than one"]
COMPLETENESS = ["Full", "Partial", "Thin"]
SDR_MINUTES = "SDR minutes (both passes)"
RECHECK = "Dead end right (lead re-check)"
SMALL = 3                       # the public export shows the SDR's side only over groups of at least this many rows
HIDDEN = "fewer than 3 rows"
UNSURE = {v.lower() for v in SDR_UNSURE_VALUES}
FOUND = {"contacts": "Contact names", "emails": "Emails", "phones": "Phones",
         "revenue": "Turnover / revenue / market share", "priority": "Priority"}


def newest_master_log(folder):
    return max(Path(folder).glob("master-log-*.xlsx"), key=lambda p: p.stat().st_mtime)


def read_rows(path):
    ws = load_workbook(path)["Rows"]
    headers = [c.value for c in ws[1]]
    return [dict(zip(headers, (c.value for c in row))) for row in ws.iter_rows(min_row=2)]


def pct(n, d):
    return f"{n} of {d} ({round(100 * n / d)}%)" if d else "none"


def num(x, digits=1):
    return "n/a" if x is None else f"{x:.{digits}f}"


def median(values):
    values = [v for v in values if v is not None]
    return statistics.median(values) if values else None


def sign_test(a_only, b_only):
    """Two-sided exact sign test on the discordant pairs (McNemar exact)."""
    n, k = a_only + b_only, min(a_only, b_only)
    return 1.0 if n == 0 else min(1.0, 2 * sum(comb(n, i) for i in range(k + 1)) / 2 ** n)


def filled(value):
    text = "" if value is None else str(value).strip()
    return bool(text) and text.lower() not in UNSURE


def arm_stats(rows, minutes_field):
    """Quality, minutes and how full the rows are, for one set of scored rows."""
    scored = [r for r in rows if r["Acceptable"]]
    acc = [r for r in scored if r["Acceptable"] == "Yes"]
    fields = [sum(filled(r.get(name)) for name in RESEARCH_FIELDS) for r in scored]
    out = {"rows": len(scored), "acceptable": len(acc),
           "correctness": {k: sum(r["Correctness"] == k for r in scored) for k in CORRECTNESS},
           "completeness": {k: sum(r["Completeness"] == k for r in scored) for k in COMPLETENESS},
           "lead_minutes_total": sum(r["Lead scoring minutes"] or 0 for r in scored),
           "lead_minutes_median": median([r["Lead scoring minutes"] for r in scored]),
           "fields_filled_median": median(fields), "fields_filled_min": min(fields) if fields else None,
           "fields_filled_max": max(fields) if fields else None,
           "found": {k: sum(filled(r.get(name)) for r in scored) for k, name in FOUND.items()},
           "dead_ends_by_kind": dict(Counter(r["Dead-end category"] or "(no kind)" for r in scored if r["Verdict"] == VERDICT_DEAD_END))}
    if minutes_field:
        timed = [r for r in scored if r[minutes_field] is not None]
        total = sum(r[minutes_field] for r in timed)
        acc_timed = sum(r["Acceptable"] == "Yes" for r in timed)
        out.update({"minutes_field": minutes_field, "timed_rows": len(timed), "minutes_total": round(total, 2),
                    "minutes_median": median([r[minutes_field] for r in timed]),
                    "minutes_min": min((r[minutes_field] for r in timed), default=None),
                    "minutes_max": max((r[minutes_field] for r in timed), default=None),
                    "minutes_mean": round(total / len(timed), 2) if timed else None,
                    "acceptable_timed": acc_timed,
                    "minutes_per_acceptable_row": round(total / acc_timed, 2) if acc_timed else None})
        if len(timed) < len(scored) and timed:
            mean = total / len(timed)
            put_in = total + mean * (len(scored) - len(timed))
            out["minutes_per_acceptable_row_with_mean_put_in"] = round(put_in / len(acc), 2) if acc else None
    return out


def window_use(folder):
    """Usage-window readings in the run records: points of the 5-hour window per company, from consecutive runs that
    sit in one window (the reading comes with each run, so a run's use shows at the next run), and the 7-day window."""
    runs = sorted((json.loads(p.read_text(encoding="utf-8")) for p in Path(folder).glob("C*/run.json")), key=lambda r: r["agent_start"])
    readings = []
    for r in runs:
        windows = ((r["summary"].get("rate_limit") or {}).get("unifiedWindows") or {})
        five, seven = windows.get("five_hour") or {}, windows.get("seven_day") or {}
        readings.append((five.get("resetsAt"), five.get("utilization"), seven.get("utilization")))
    groups = []
    for reset, used, _ in readings:
        if used is None:
            continue
        if groups and groups[-1][0] == reset:
            groups[-1][1].append(used)
        else:
            groups.append((reset, [used]))
    steps = [(g[1][-1] - g[1][0], len(g[1]) - 1) for g in groups if len(g[1]) > 1]
    rise, intervals = sum(s for s, _ in steps), sum(n for _, n in steps)
    sevens = [s for _, _, s in readings if s is not None]
    return {"runs": len(runs), "five_hour_groups": [{"runs": len(g[1]), "from": g[1][0], "to": g[1][-1]} for g in groups],
            "five_hour_points_per_company": round(100 * rise / intervals, 1) if intervals else None,
            "five_hour_intervals": intervals,
            "seven_day_from": sevens[0] if sevens else None, "seven_day_to": sevens[-1] if sevens else None}


def batch_cost(folder):
    runs = [json.loads(p.read_text(encoding="utf-8")) for p in sorted(Path(folder).glob("C*/run.json"))]
    tools = Counter()
    for r in runs:
        tools.update(r["summary"]["tools_called"])
    calls = [len(r["tool_calls"]) for r in runs]
    minutes = [r["agent_minutes"] for r in runs]
    cost = sum((r["summary"]["result"].get("total_cost_usd") or 0) + ((r.get("repair") or {}).get("result") or {}).get("total_cost_usd", 0)
               for r in runs)
    return {"folder": Path(folder).name, "runs": len(runs), "prompt_versions": dict(Counter(r["prompt_version"] for r in runs)),
            "models": dict(Counter(m for r in runs for m in r["summary"]["models_used"])), "effort": dict(Counter(r["effort"] for r in runs)),
            "valid": sum(r["validator_result"] == "valid" for r in runs), "repaired": sum(r["retry_count"] for r in runs),
            "minutes_total": round(sum(minutes), 2), "minutes_median": median(minutes), "minutes_min": min(minutes), "minutes_max": max(minutes),
            "tool_calls_total": sum(calls), "tool_calls_median": median(calls), "tool_calls_min": min(calls), "tool_calls_max": max(calls),
            "tools": dict(tools), "turns_total": sum(r["summary"]["result"].get("num_turns") or 0 for r in runs),
            "apollo_calls": sum(sum(r["apollo_calls_that_ran"].values()) for r in runs),
            "apollo_credit_blocks": sum(len(r["apollo_credit_blocks"]) for r in runs),
            "calls_blocked_by_the_cap": sum(r["calls_blocked_by_the_cap"] for r in runs),
            "api_equivalent_usd": round(cost, 2), "window": window_use(folder),
            "first_start": min(r["agent_start"] for r in runs), "last_end": max(r["agent_end"] for r in runs)}


def judged(rows):
    """The rows as the design reads them once the dead-end re-check is in: a dead-end row with a re-check answer is
    acceptable when the lead judged the dead end right. Every other row keeps the lead's first score."""
    out = []
    for r in rows:
        r = dict(r)
        if r.get(RECHECK) in ("Yes", "No") and r["Verdict"] == VERDICT_DEAD_END and r["Acceptable"]:
            r["Acceptable"] = r[RECHECK]
        out.append(r)
    return out


def compute(rows, folders):
    by_company = {}
    for r in rows:
        by_company.setdefault(r["Company ID"], {})[r["Row type"]] = r
    alone = [r for r in rows if r["Row type"] == ROW_AGENT_ALONE]
    sdr = [r for r in rows if r["Row type"] == ROW_SDR_ALONE]
    pair = [r for r in rows if r["Row type"] == ROW_PAIR]
    pair_scored = [r for r in pair if r["Acceptable"]]
    pair_ids = {r["Company ID"] for r in pair_scored}
    sdr_ids = {r["Company ID"] for r in sdr if r["Acceptable"]}
    sets = {
        "Agent alone": (alone, "Agent minutes"),
        "Agent alone, on the SDR-alone companies": ([r for r in alone if r["Company ID"] in sdr_ids], "Agent minutes"),
        "Agent alone, on the pair companies": ([r for r in alone if r["Company ID"] in pair_ids], "Agent minutes"),
        "SDR alone": (sdr, SDR_MINUTES),
        "SDR alone, early (before the build)": ([r for r in sdr if r["Early or late"] == EARLY], SDR_MINUTES),
        "SDR alone, late (interleaved)": ([r for r in sdr if r["Early or late"] == LATE], SDR_MINUTES),
        "Pair": (pair, SDR_MINUTES),
        "Pair v1": ([r for r in pair if r["Prompt version"] == "v1"], SDR_MINUTES),
        "Pair v2": ([r for r in pair if r["Prompt version"] == "v2"], SDR_MINUTES),
    }
    arms = {name: arm_stats(subset, field) for name, (subset, field) in sets.items()}
    for name in ("Pair", "Pair v1", "Pair v2"):
        arms[name]["agent"] = arm_stats(sets[name][0], "Agent minutes")
        timed = [r for r in sets[name][0] if r["Acceptable"] and r["Check minutes"] is not None]
        arms[name]["check_minutes_total"] = sum(r["Check minutes"] for r in timed)
        arms[name]["completing_minutes_total"] = sum(r["Completing minutes"] or 0 for r in timed)

    early = [r for r in sdr if r["Early or late"] == EARLY and r["Acceptable"]]
    out = {"coverage": {
        "companies": len(by_company), "companies_scored_whole": sum(1 for c in by_company.values() if all(r["Acceptable"] for r in c.values())),
        "scored_rows": dict(Counter(r["Row type"] for r in rows if r["Acceptable"])),
        "pair_rows_not_scored": [{"prompt_version": r["Prompt version"], "sdr_status": r["SDR status"]} for r in pair if not r["Acceptable"]],
        "sdr_rows_without_minutes": sum(1 for r in sdr + pair if r["Acceptable"] and r[SDR_MINUTES] is None),
        "sdr_minutes_basis": dict(Counter(r["Minutes basis"] for r in sdr + pair if r["Acceptable"])),
        "early_rows": len(early), "early_rows_block_average": sum("block average" in (r["Minutes basis"] or "") for r in early),
        "early_rows_one_span": sum((r["Minutes basis"] or "").startswith("one span") and "block average" not in r["Minutes basis"] for r in early),
        "early_rows_unreadable": sum((r["Minutes basis"] or "").startswith("unreadable") for r in early),
        "rechecked_rows": sum(r.get(RECHECK) in ("Yes", "No") for r in rows)},
        "arms": arms}

    # the five falsification conditions
    s, p = arms["SDR alone"], arms
    conditions = {}
    for name in ("Pair", "Pair v1", "Pair v2"):
        a = p[name]
        saving = 1 - a["minutes_per_acceptable_row"] / s["minutes_per_acceptable_row"] if a.get("minutes_per_acceptable_row") and s.get("minutes_per_acceptable_row") else None
        rate, rate_sdr = a["acceptable"] / a["rows"], s["acceptable"] / s["rows"]
        ids = {r["Company ID"] for r in sets[name][0] if r["Acceptable"]}
        same = [(by_company[cid][ROW_PAIR]["Acceptable"] == "Yes", by_company[cid][ROW_AGENT_ALONE]["Acceptable"] == "Yes") for cid in ids
                if by_company[cid][ROW_AGENT_ALONE]["Acceptable"]]
        dead = [r for r in sets[name][0] if r["Acceptable"] and r["Verdict"] == VERDICT_DEAD_END and r[SDR_MINUTES] is not None]
        dead_acc = [r for r in dead if r["Acceptable"] == "Yes"]
        conditions[name] = {
            "1_speed_saving": round(saving, 3) if saving is not None else None,
            "1_holds": saving is not None and saving >= 0.5,
            "2_rate_pair": round(rate, 3), "2_rate_sdr_alone": round(rate_sdr, 3), "2_gap_points": round(100 * (rate - rate_sdr), 1),
            "2_holds": rate - rate_sdr >= -1 / 15 - 1e-9,
            "3_same_companies": len(same), "3_pair_yes": sum(x for x, _ in same), "3_agent_yes": sum(y for _, y in same),
            "3_pair_only": sum(x and not y for x, y in same), "3_agent_only": sum(y and not x for x, y in same),
            "3_holds": sum(x for x, _ in same) > sum(y for _, y in same),
            "3_sign_test_p": round(sign_test(sum(x and not y for x, y in same), sum(y and not x for x, y in same)), 3),
            "4_dead_end_rows": len(dead), "4_mean_sdr_minutes": round(statistics.mean(r[SDR_MINUTES] for r in dead), 2) if dead else None,
            "4_dead_end_rows_accepted": len(dead_acc),
            "4_mean_sdr_minutes_accepted": round(statistics.mean(r[SDR_MINUTES] for r in dead_acc), 2) if dead_acc else None,
            "4_holds": bool(dead) and statistics.mean(r[SDR_MINUTES] for r in dead) < 5,
        }
    # condition 5: the agent's flag on the SDR-alone companies, against the SDR-alone verdicts the lead accepted
    accepted = [(by_company[r["Company ID"]][ROW_AGENT_ALONE], r) for r in sdr if r["Acceptable"] == "Yes" and by_company[r["Company ID"]].get(ROW_AGENT_ALONE)]
    rejected = sum(1 for r in sdr if r["Acceptable"] == "No")
    catches = sum(a["Flag value"] == "Yes" and r["Verdict"] == VERDICT_DEAD_END for a, r in accepted)
    false_alarms = sum(a["Flag value"] == "Yes" and r["Verdict"] != VERDICT_DEAD_END for a, r in accepted)
    misses = [r["Dead-end category"] for a, r in accepted if a["Flag value"] != "Yes" and r["Verdict"] == VERDICT_DEAD_END
              and r["Dead-end category"] != DEAD_END_ACTIVE]
    conditions["flag"] = {"sdr_alone_rows_accepted": len(accepted), "left_out_not_accepted": rejected, "catches": catches,
                          "false_alarms": false_alarms, "misses": len(misses), "misses_by_type": dict(Counter(misses)),
                          "5_holds": not false_alarms > catches,
                          "flags_raised_on_all_agent_rows": sum(r["Flag value"] == "Yes" for r in alone + pair if r["Flag value"])}
    out["conditions"] = conditions

    # verdicts and dead ends between arms, company by company
    def verdicts(kind, ids):
        both = [(by_company[cid][ROW_AGENT_ALONE], by_company[cid][kind]) for cid in sorted(ids)]
        return {"companies": len(both),
                "same_verdict": sum(a["Verdict"] == h["Verdict"] for a, h in both),
                "both_prospect": sum(a["Verdict"] == h["Verdict"] != VERDICT_DEAD_END for a, h in both),
                "both_dead_end": sum(a["Verdict"] == h["Verdict"] == VERDICT_DEAD_END for a, h in both),
                "dead_end_only_human_arm": sum(h["Verdict"] == VERDICT_DEAD_END != a["Verdict"] for a, h in both),
                "dead_end_only_agent": sum(a["Verdict"] == VERDICT_DEAD_END != h["Verdict"] for a, h in both),
                "human_dead_ends_by_type": dict(Counter(h["Dead-end category"] for _, h in both if h["Verdict"] == VERDICT_DEAD_END)),
                "human_dead_ends_accepted_by_the_lead": sum(h["Verdict"] == VERDICT_DEAD_END and h["Acceptable"] == "Yes" for _, h in both),
                "agent_flagged_where_human_found_dead_end": sum(h["Verdict"] == VERDICT_DEAD_END and a["Flag value"] == "Yes" for a, h in both),
                "both_acceptable": sum(a["Acceptable"] == h["Acceptable"] == "Yes" for a, h in both),
                "only_human_arm_acceptable": sum(h["Acceptable"] == "Yes" != a["Acceptable"] for a, h in both),
                "only_agent_acceptable": sum(a["Acceptable"] == "Yes" != h["Acceptable"] for a, h in both),
                "neither_acceptable": sum(a["Acceptable"] == h["Acceptable"] == "No" for a, h in both)}
    agreement = {"Agent alone vs SDR alone": verdicts(ROW_SDR_ALONE, sdr_ids), "Agent alone vs Pair": verdicts(ROW_PAIR, pair_ids)}
    for v in ("v1", "v2"):
        agreement[f"Agent alone vs Pair {v}"] = verdicts(ROW_PAIR, {r["Company ID"] for r in pair_scored if r["Prompt version"] == v})
    for k, a in agreement.items():
        a["sign_test_p"] = round(sign_test(a["only_human_arm_acceptable"], a["only_agent_acceptable"]), 3)
    started = {r["Company ID"]: r for r in pair_scored}
    agreement["Pair: the agent row it started from vs the SDR's final verdict"] = {
        "companies": len(started), "agent_flag_yes": sum(r["Flag value"] == "Yes" for r in started.values()),
        "sdr_closed_as_dead_end": sum(r["Verdict"] == VERDICT_DEAD_END for r in started.values()),
        "by_version": {v: {"rows": sum(r["Prompt version"] == v for r in started.values()),
                           "sdr_closed_as_dead_end": sum(r["Prompt version"] == v and r["Verdict"] == VERDICT_DEAD_END for r in started.values())}
                       for v in ("v1", "v2")}}
    out["verdicts"] = agreement

    # triage time: SDR minutes on dead-end rows, by arm and type
    out["triage"] = {kind: [{"type": r["Dead-end category"], "sdr_minutes": r[SDR_MINUTES], "accepted": r["Acceptable"],
                             "version": r["Prompt version"] if kind == ROW_PAIR else None}
                            for r in rows if r["Row type"] == kind and r["Acceptable"] and r["Verdict"] == VERDICT_DEAD_END]
                     for kind in (ROW_SDR_ALONE, ROW_PAIR)}

    # by market and by difficulty
    cuts = {}
    for label, field, values in (("market", "Market", MARKETS), ("difficulty", "Difficulty", DIFFICULTIES)):
        cuts[label] = {}
        for value in values:
            cuts[label][value] = {name: arm_stats([r for r in sets[name][0] if r[field] == value], sets[name][1])
                                  for name in ("Agent alone", "SDR alone", "Pair", "Pair v1", "Pair v2")}
            for name in ("Pair", "Pair v1", "Pair v2"):
                cuts[label][value][name]["agent"] = arm_stats([r for r in sets[name][0] if r[field] == value], "Agent minutes")
    out["cuts"] = cuts

    # Dead-end rows: the lead's sheet asked "would you let the SDR call from this row?" and, for a dead end, to judge the
    # verdict and its reason. On a dead end those point opposite ways, so the score there is read with the comments.
    dead_rows = [r for r in rows if r["Row type"] != ROW_AGENT_ALONE and r["Acceptable"] and r["Verdict"] == VERDICT_DEAD_END]
    texts = Counter(r["Lead comment"] for r in dead_rows if r["Lead comment"])
    out["dead_end_rows"] = {
        "rows": len(dead_rows), "accepted": sum(r["Acceptable"] == "Yes" for r in dead_rows),
        "correctness_none": sum(r["Correctness"] == "None" for r in dead_rows),
        "rows_sharing_one_comment_word_for_word": max(texts.values()) if texts else 0,
        "comments_saying_no_information_or_research": sum(any(k in (r["Lead comment"] or "").lower() for k in ("no information", "no data", "no research"))
                                                          for r in dead_rows),
        "agent_row_on_the_same_company_not_accepted": sum(by_company[r["Company ID"]][ROW_AGENT_ALONE]["Acceptable"] == "No" for r in dead_rows),
        "rechecked": sum(r.get(RECHECK) in ("Yes", "No") for r in dead_rows),
        "judged_right": sum(r.get(RECHECK) == "Yes" for r in dead_rows),
    }

    # A second reading, not set in advance: prospect rows only, where the lead's question fits the row
    prospects = {name: arm_stats([r for r in sets[name][0] if r["Verdict"] != VERDICT_DEAD_END], sets[name][1])
                 for name in ("Agent alone", "SDR alone", "Pair", "Pair v1", "Pair v2")}
    for name in ("Pair", "Pair v1", "Pair v2"):
        prospects[name]["agent"] = arm_stats([r for r in sets[name][0] if r["Verdict"] != VERDICT_DEAD_END], "Agent minutes")
    pair_prospect_ids = {r["Company ID"] for r in pair_scored if r["Verdict"] != VERDICT_DEAD_END}
    prospects["same companies"] = {"pair_prospect_companies": len(pair_prospect_ids),
                                   "pair_yes": sum(by_company[cid][ROW_PAIR]["Acceptable"] == "Yes" for cid in pair_prospect_ids),
                                   "agent_alone_yes": sum(by_company[cid][ROW_AGENT_ALONE]["Acceptable"] == "Yes" for cid in pair_prospect_ids)}
    sp, pp = prospects["SDR alone"], prospects["Pair"]
    prospects["saving"] = round(1 - pp["minutes_per_acceptable_row"] / sp["minutes_per_acceptable_row"], 3)
    out["prospect_rows_only"] = prospects

    # The unscored Pair rows: how far their SDR minutes could move condition 1, whatever the lead would score
    waiting = [r for r in pair if not r["Acceptable"] and r[SDR_MINUTES] is not None]
    base, sdr_rate = arms["Pair"], arms["SDR alone"]["minutes_per_acceptable_row"]
    extra = sum(r[SDR_MINUTES] for r in waiting)
    span = [(base["minutes_total"] + extra) / (base["acceptable_timed"] + k) for k in range(len(waiting) + 1)]
    pros_extra = sum(r[SDR_MINUTES] for r in waiting if r["Verdict"] != VERDICT_DEAD_END)
    pros_span = [(pp["minutes_total"] + pros_extra) / (pp["acceptable_timed"] + k) for k in range(len(waiting) + 1)]
    out["waiting_pair_rows"] = {
        "rows": len(waiting), "sdr_minutes": [r[SDR_MINUTES] for r in waiting],
        "versions": [r["Prompt version"] for r in waiting], "verdicts": [r["Verdict"] for r in waiting],
        "pair_minutes_per_acceptable_row_if_counted": [round(min(span), 2), round(max(span), 2)],
        "saving_if_counted": [round(1 - max(span) / sdr_rate, 3), round(1 - min(span) / sdr_rate, 3)],
        "prospects_saving_if_counted": [round(1 - max(pros_span) / sp["minutes_per_acceptable_row"], 3),
                                        round(1 - min(pros_span) / sp["minutes_per_acceptable_row"], 3)],
    }

    # What the lead's comments point at: keyword counts (a comment can carry several), by arm and score
    themes = {"priority wrong or missing": ("priority",), "revenue missing": ("revenue",), "contacts missing": ("contact",),
              "researched a different company": ("not matching", "different compan", "not same as", "for which"),
              "not the right target (lab-grown, segment)": ("target account", "lab-grown", "lab grown"),
              "too little information": ("no information", "no data", "no much information", "not much information", "not suffi", "no research"),
              "company closed": ("closed",)}
    out["comment_themes"] = {}
    for kind in (ROW_AGENT_ALONE, ROW_SDR_ALONE, ROW_PAIR):
        scored_kind = [r for r in rows if r["Row type"] == kind and r["Acceptable"]]
        out["comment_themes"][kind] = {
            score: {"rows": sum(r["Acceptable"] == score for r in scored_kind),
                    "with_comment": sum(r["Acceptable"] == score and bool(r["Lead comment"]) for r in scored_kind),
                    **{t: sum(r["Acceptable"] == score and any(k in (r["Lead comment"] or "").lower() for k in keys) for r in scored_kind)
                       for t, keys in themes.items()}}
            for score in ("No", "Yes")}
    out["agent_batches"] = [batch_cost(f) for f in folders]
    return out


# ---------------------------------------------------------------- the report

SDR_SIDE = ("SDR alone", "Pair")


def sdr_side(name):
    return name.startswith(SDR_SIDE)


def arm_table(arms, names, public):
    """One line per arm: rows, the lead's scores, minutes (total, median, range, per acceptable row), how full the rows
    are and the dead ends by kind. In the public export an SDR-side line over fewer than 3 rows is not shown."""
    lines = ["| Arm | Rows | Acceptable | Correctness None / One / More | Completeness Full / Partial / Thin | SDR min total "
             "| SDR min median (range) | SDR min per acceptable row | Agent min median (range) | Agent min per acceptable row "
             "| Research fields filled per row, median (range, of 25) | Rows with contacts / emails / phones / revenue / priority "
             "| Dead ends by kind | Lead min per row (median) |",
             "|---|---|---|---|---|---|---|---|---|---|---|---|---|---|"]
    for name in names:
        a = arms[name]
        if public and sdr_side(name) and a["rows"] < SMALL:
            lines.append(f"| {name} | {a['rows']} | " + " | ".join([HIDDEN] * 12) + " |")
            continue
        human = a.get("minutes_field") == SDR_MINUTES
        agent = a.get("agent") if human else a
        sdr_cells = (f"{num(a['minutes_total'], 0)} on {a['timed_rows']} | {num(a['minutes_median'])} ({num(a['minutes_min'], 0)} to "
                     f"{num(a['minutes_max'], 0)}) | **{num(a['minutes_per_acceptable_row'])}**") if human else "0 | 0 | 0"
        agent_cells = (f"{num(agent['minutes_median'], 2)} ({num(agent['minutes_min'], 2)} to {num(agent['minutes_max'], 2)}) | "
                       f"{num(agent['minutes_per_acceptable_row'], 2)}") if agent else "0 | 0"
        f = a["found"]
        dead = ", ".join(f"{k} {n}" for k, n in sorted(a["dead_ends_by_kind"].items())) or "none"
        lines.append(f"| {name} | {a['rows']} | {pct(a['acceptable'], a['rows'])} | "
                     + " / ".join(str(a["correctness"][k]) for k in CORRECTNESS) + " | "
                     + " / ".join(str(a["completeness"][k]) for k in COMPLETENESS) + f" | {sdr_cells} | {agent_cells} | "
                     f"{num(a['fields_filled_median'], 0)} ({a['fields_filled_min']} to {a['fields_filled_max']}) | "
                     f"{f['contacts']} / {f['emails']} / {f['phones']} / {f['revenue']} / {f['priority']} | {dead} | "
                     f"{num(a['lead_minutes_median'], 0)} |")
    return lines


def conditions_table(c, public, title):
    def holds(x):
        return "holds" if x else "**falsified**"
    lines = ["", title, "", "| Condition (bar) | Pair | Pair v1 | Pair v2 |", "|---|---|---|---|"]
    row1 = "| 1. Speed: the pair saves at least half of the SDR-alone minutes per acceptable row |"
    row2 = "| 2. Quality held: the pair's acceptable rate no more than one row in 15 (6.7 points) below the SDR alone's |"
    row3 = "| 3. Human value: the pair's acceptable rate above the agent alone's on the same companies |"
    row4 = "| 4. Triage: confirmed dead ends in the pair average under 5 SDR minutes |"
    for n in ("Pair", "Pair v1", "Pair v2"):
        x = c[n]
        row1 += f" saves {num(100 * x['1_speed_saving'], 0) if x['1_speed_saving'] is not None else 'n/a'}%: {holds(x['1_holds'])} |"
        row2 += f" {num(100 * x['2_rate_pair'], 0)}% vs {num(100 * x['2_rate_sdr_alone'], 0)}% ({x['2_gap_points']:+} pts): {holds(x['2_holds'])} |"
        row3 += (f" {x['3_pair_yes']} vs {x['3_agent_yes']} of {x['3_same_companies']} (pair only {x['3_pair_only']}, agent only "
                 f"{x['3_agent_only']}; sign test p {x['3_sign_test_p']}): {holds(x['3_holds'])} |")
        if not x["4_dead_end_rows"]:
            row4 += " no dead end: not tested |"
        elif public and x["4_dead_end_rows"] < SMALL:
            row4 += f" {HIDDEN} |"
        else:
            accepted = x["4_dead_end_rows_accepted"]
            judged = ("" if not 0 < accepted < x["4_dead_end_rows"] else
                      f" (the {accepted} the lead judged right: {HIDDEN})" if public and accepted < SMALL else
                      f" (the {accepted} the lead judged right: mean {num(x['4_mean_sdr_minutes_accepted'])} min)")
            row4 += f" {x['4_dead_end_rows']} dead ends, mean {num(x['4_mean_sdr_minutes'])} min{judged}: {holds(x['4_holds'])} |"
    f = c["flag"]
    return lines + [row1, row2, row3, row4, "",
                    f"5. Flag (false alarms must not outnumber catches), on the {f['sdr_alone_rows_accepted']} SDR-alone rows the lead "
                    f"accepted ({f['left_out_not_accepted']} not accepted, left out): catches {f['catches']}, false alarms {f['false_alarms']}, "
                    f"misses {f['misses']}: {'holds' if f['5_holds'] else '**falsified**'} as written. The agent raised its flag on "
                    f"{f['flags_raised_on_all_agent_rows']} of all its rows" + (", so the condition cannot fail: nothing was flagged."
                                                                               if not f["flags_raised_on_all_agent_rows"] else ".")]


def markdown(res, source, public=False, recheck=None):
    a, c, v = res["arms"], res["conditions"], res["verdicts"]
    cov = res["coverage"]
    main_arms = ["Agent alone", "Agent alone, on the SDR-alone companies", "Agent alone, on the pair companies", "SDR alone",
                 "SDR alone, early (before the build)", "SDR alone, late (interleaved)", "Pair", "Pair v1", "Pair v2"]
    lines = [
        "# Results: SDR research agent trial" + (" (public export)" if public else ""),
        "",
        f"Generated by `results.py{' --public' if public else ''}` from `{source['master_log']}` (lead returns: {source['lead_return']}). "
        "Aggregates only." + (" The SDR's side is shown only over groups of 3 rows or more." if public else ""),
        "",
        "## Basis",
        "",
        f"- Scored rows: {', '.join(f'{k} {n}' for k, n in sorted(cov['scored_rows'].items()))}; "
        f"{cov['companies_scored_whole']} of {cov['companies']} companies scored on both rows. Not scored: "
        f"{len(cov['pair_rows_not_scored'])} Pair rows with no Status from the SDR ("
        + ", ".join(f"{x['prompt_version']}" for x in cov["pair_rows_not_scored"]) + "), so their companies wait.",
        "- The lead scored blind: row codes only, no arm, no flag, no CRM, a company's two rows apart. One rater.",
        f"- SDR minutes come from the SDR's hh:mm stamps (minute steps). The first {cov['early_rows']} SDR-alone rows (return 1) were "
        f"stamped in blocks: {cov['early_rows_block_average']} carry a block average (one span shared by two rows, split evenly), "
        f"{cov['early_rows_one_span']} one span for both passes, {cov['early_rows_unreadable']} no readable minutes (it leaves both sums; "
        "the arm's mean put in for it is shown as a check). Every later row is stamped per pass.",
        "- Agent minutes: the run behind each row, unattended; never added to SDR minutes. The agent-alone arm is the v1 run "
        "of each company; Pair v1 rows were built on that same run, Pair v2 rows on a v2 run.",
        "- Small samples: one row moves a 15-row rate by about 7 points. Read every rate with its count.",
        f"- The lead's stamps give {sum(a[n]['lead_minutes_total'] for n in ('Agent alone', 'SDR alone', 'Pair'))} minutes for "
        f"{sum(a[n]['rows'] for n in ('Agent alone', 'SDR alone', 'Pair'))} rows (a median of 1 a row): a quick read of each row, "
        "not a check of its sources, so Correctness counts what the lead noticed.",
        "- Research fields filled: of the 25 research fields, those holding a value (an SDR's \"Not Sure\" counts as empty). "
        "A dead-end row is a short row by design, so it fills few.",
    ]
    d = res["dead_end_rows"]
    if recheck:
        ra, rc = recheck["arms"], recheck["conditions"]
        lines += ["", "## With the re-check: dead ends judged on the verdict (the designed measure)", "",
                  f"The lead judged the {d['rechecked']} dead-end rows again on one question, \"Is this dead end right, on its reason and "
                  f"source?\" (DECISIONS.md 18), and judged {d['judged_right']} right. Here a dead-end row counts as acceptable when its "
                  "dead end was judged right; every other row keeps its first score.", ""]
        lines += arm_table(ra, ["Agent alone", "SDR alone", "Pair", "Pair v1", "Pair v2"], public)
        lines += conditions_table(rc, public, "### The hypothesis, condition by condition (with the re-check)")
        rw = recheck["waiting_pair_rows"]
        lines += ["", f"With all {rw['rows'] + ra['Pair']['rows']} Pair rows counted, whatever the lead would score the {rw['rows']} not yet "
                      f"scored, the pair takes {rw['pair_minutes_per_acceptable_row_if_counted'][0]} to "
                      f"{rw['pair_minutes_per_acceptable_row_if_counted'][1]} SDR minutes per acceptable row: a saving of "
                      f"{round(100 * rw['saving_if_counted'][0])}% to {round(100 * rw['saving_if_counted'][1])}%, so "
                      + ("condition 1 holds on the scored rows only." if rw["saving_if_counted"][1] < 0.5 else
                         "condition 1 holds with them counted too." if rw["saving_if_counted"][0] >= 0.5 else
                         "whether condition 1 holds with them counted depends on how they are scored.")]
    lines += ["", "## By arm (as first scored)", ""] + arm_table(a, main_arms, public)
    s = a["SDR alone"]
    if s.get("minutes_per_acceptable_row_with_mean_put_in") is not None:
        lines += ["", f"Check: with the arm's mean put in for the SDR-alone row without minutes, SDR alone is "
                      f"{num(s['minutes_per_acceptable_row_with_mean_put_in'])} SDR minutes per acceptable row."]
    lines += ["", "Pair SDR minutes split: " + "; ".join(
        f"{n}: check {num(a[n]['check_minutes_total'], 0)}, completing {num(a[n]['completing_minutes_total'], 0)}" for n in ("Pair", "Pair v1", "Pair v2")) + "."]
    lines += conditions_table(c, public, "## The hypothesis, condition by condition (as first scored)")

    w, pr = res["waiting_pair_rows"], res["prospect_rows_only"]
    lines += ["", "### Read with care: the dead-end rows", "",
              f"- The lead accepted {d['accepted']} of the {d['rows']} dead-end rows at first (all from the human arms: the agent never gave a "
              f"dead end). {d['comments_saying_no_information_or_research']} of their comments say there is no information or research to "
              f"call from, and {d['rows_sharing_one_comment_word_for_word']} share one comment word for word; {d['correctness_none']} carry no "
              "wrong fact.",
              "- The sheet asked \"would you let the SDR call from this row?\" and, for a dead end, to judge the verdict and its reason. On a "
              "dead end those point opposite ways, and the comments show the lead answered the first. So a first No on a dead-end row does "
              "not say the verdict was wrong" + (", which is why the dead ends were judged again (above)." if recheck else
                                                  ", and conditions 2, 4 and 5 cannot be read as designed from these scores."),
              f"- On {d['agent_row_on_the_same_company_not_accepted']} of the {d['rows']} companies a human arm closed as a dead end, the lead also "
              "refused the agent's Prospect row for the same company.", "",
              "### A second reading, not set in advance: prospect rows only", "",
              "| Arm | Rows | Acceptable | SDR min per acceptable row | Agent min per acceptable row |", "|---|---|---|---|---|"]
    for name in ("Agent alone", "SDR alone", "Pair", "Pair v1", "Pair v2"):
        x = pr[name]
        if public and sdr_side(name) and x["rows"] < SMALL:
            lines.append(f"| {name} | {x['rows']} | {HIDDEN} | {HIDDEN} | {HIDDEN} |")
            continue
        sdr_cell = num(x.get("minutes_per_acceptable_row")) if x.get("minutes_field") == SDR_MINUTES else "0"
        agent_cell = (num(x.get("minutes_per_acceptable_row"), 2) if x.get("minutes_field") == "Agent minutes"
                      else num(x["agent"]["minutes_per_acceptable_row"], 2) if "agent" in x else "0")
        lines.append(f"| {name} | {x['rows']} | {pct(x['acceptable'], x['rows'])} | {sdr_cell} | {agent_cell} |")
    same = pr["same companies"]
    lines += ["", f"On prospect rows the pair saves {round(100 * pr['saving'])}% of the SDR-alone minutes per acceptable row; on the "
                  f"{same['pair_prospect_companies']} pair companies with a prospect row the pair is acceptable on {same['pair_yes']} and the agent "
                  f"alone on {same['agent_alone_yes']}.", "",
              "### The Pair rows not scored", "",
              f"- {w['rows']} Pair rows ({', '.join(w['versions'])}; verdicts {', '.join(w['verdicts'])}) have no Status, so their companies were "
              "not sent to the lead." + ("" if public else f" Their SDR minutes: {', '.join(num(m, 0) for m in w['sdr_minutes'])}."),
              f"- With all {w['rows'] + a['Pair']['rows']} Pair rows counted, whatever the lead would score these {w['rows']}, the pair would take "
              f"{w['pair_minutes_per_acceptable_row_if_counted'][0]} to {w['pair_minutes_per_acceptable_row_if_counted'][1]} SDR minutes per "
              f"acceptable row: a saving of {round(100 * w['saving_if_counted'][0])}% to {round(100 * w['saving_if_counted'][1])}% as first "
              f"scored, {round(100 * w['prospects_saving_if_counted'][0])}% to {round(100 * w['prospects_saving_if_counted'][1])}% on prospect "
              "rows. Condition 1 holds only while they stay out."]

    lines += ["", "## What the lead's comments point at", "",
              "Keyword counts over the lead's comments; a comment can carry several. Rows marked No / rows marked Yes at first.", "",
              "| Arm | Rows (with a comment) | " + " | ".join(t.capitalize() for t in next(iter(res["comment_themes"].values()))["No"]
                                                    if t not in ("rows", "with_comment")) + " |",
              "|---|---|" + "---|" * (len(next(iter(res["comment_themes"].values()))["No"]) - 2)]
    for kind, by_score in res["comment_themes"].items():
        no, yes = by_score["No"], by_score["Yes"]
        lines.append(f"| {kind} | No {no['rows']} ({no['with_comment']}) / Yes {yes['rows']} ({yes['with_comment']}) | "
                     + " | ".join(f"{no[t]} / {yes[t]}" for t in no if t not in ("rows", "with_comment")) + " |")

    lines += ["", "## Verdicts and dead ends between arms (company by company)", "",
              "| Comparison | Companies | Same verdict | Dead end, human arm only | Dead end, agent only | Human dead ends the lead accepted at first "
              "| Both acceptable | Human arm only | Agent only | Neither | Sign test p |", "|---|---|---|---|---|---|---|---|---|---|---|"]
    for name, x in v.items():
        if "same_verdict" in x:
            dead_n = x["dead_end_only_human_arm"] + x["both_dead_end"]
            accepted = (HIDDEN if public and 0 < dead_n < SMALL else f"{x['human_dead_ends_accepted_by_the_lead']} of {dead_n}")
            lines.append(f"| {name} | {x['companies']} | {x['same_verdict']} | {x['dead_end_only_human_arm']} | {x['dead_end_only_agent']} | "
                         f"{accepted} | {x['both_acceptable']} | {x['only_human_arm_acceptable']} | {x['only_agent_acceptable']} | "
                         f"{x['neither_acceptable']} | {x['sign_test_p']} |")
    st = v["Pair: the agent row it started from vs the SDR's final verdict"]
    lines += ["", f"Pair: the agent rows the SDR started from carried {st['agent_flag_yes']} flags; the SDR closed "
                  f"{st['sdr_closed_as_dead_end']} of {st['companies']} as dead ends ("
                  + ", ".join(f"{k}: {x['sdr_closed_as_dead_end']} of {x['rows']}" for k, x in st["by_version"].items()) + ")."]
    lines += ["", "Dead-end rows, SDR minutes (the triage measure):", ""]
    for kind, items in res["triage"].items():
        kinds = Counter(i["type"] or "(no kind)" for i in items)
        minutes = [i["sdr_minutes"] for i in items if i["sdr_minutes"] is not None]
        counted = ", ".join(f"{t} {n}" for t, n in sorted(kinds.items())) or "none"
        if public:
            stats = (HIDDEN if len(minutes) < SMALL else
                     f"mean {num(statistics.mean(minutes))}, median {num(median(minutes))}, range {num(min(minutes), 0)} to {num(max(minutes), 0)}")
            lines.append(f"- {kind}: {len(items)} dead-end rows ({counted}); SDR minutes {stats}")
        else:
            lines.append(f"- {kind}: " + ("; ".join(
                f"{t}: {n} row(s), SDR minutes " + ", ".join(num(i['sdr_minutes'], 0) for i in items if (i['type'] or '(no kind)') == t)
                + ", accepted at first " + str(sum(i['accepted'] == 'Yes' for i in items if (i['type'] or '(no kind)') == t))
                for t, n in kinds.items()) or "none"))

    for label, title in (("market", "By market"), ("difficulty", "By difficulty")):
        lines += ["", f"## {title} (as first scored: a dead-end row counts as not acceptable)"]
        for value, arms_cut in res["cuts"][label].items():
            lines += ["", f"**{value}**", ""] + arm_table(arms_cut, ["Agent alone", "SDR alone", "Pair", "Pair v1", "Pair v2"], public)

    lines += ["", "## The agent's cost", "", "| Batch | Runs | Valid (repaired) | Minutes total / median (range) | Tool calls total / median (range) "
              "| Searches / fetches | Apollo calls / credits | 5-hour window, points per company | 7-day window | API list-price equivalent |",
              "|---|---|---|---|---|---|---|---|---|---|"]
    for b in res["agent_batches"]:
        wu = b["window"]
        lines.append(f"| {b['folder']} ({', '.join(b['prompt_versions'])}) | {b['runs']} | {b['valid']} ({b['repaired']}) | "
                     f"{num(b['minutes_total'])} / {num(b['minutes_median'], 2)} ({num(b['minutes_min'], 2)} to {num(b['minutes_max'], 2)}) | "
                     f"{b['tool_calls_total']} / {num(b['tool_calls_median'], 1)} ({b['tool_calls_min']} to {b['tool_calls_max']}) | "
                     f"{b['tools'].get('WebSearch', 0)} / {b['tools'].get('WebFetch', 0)} | {b['apollo_calls']} / 0 | "
                     f"{num(wu['five_hour_points_per_company'])} (over {wu['five_hour_intervals']} intervals) | "
                     f"{round(100 * wu['seven_day_from'])}% to {round(100 * wu['seven_day_to'])}% | ${b['api_equivalent_usd']:.2f} |")
    lines += ["", "Window points are read from each run's usage reading, as the rise between consecutive runs in one 5-hour window; "
                  "anything else on the account in those minutes is included, so they are an upper bound. The list-price figure is the "
                  "CLI's own estimate for the main runs and repairs; the plan is a subscription, so it is not a bill.", ""]
    return "\n".join(lines)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("--master-log")
    parser.add_argument("--agent-runs", nargs="+", required=True)
    parser.add_argument("--data-root", default=str(HERE / "data-exports"))
    parser.add_argument("--md", default=str(HERE / "RESULTS.md"))
    parser.add_argument("--public", action="store_true", help="the public export's tables (DECISIONS.md 18)")
    parser.add_argument("--no-json", action="store_true", help="write the markdown only")
    args = parser.parse_args()
    data = Path(args.data_root)
    path = Path(args.master_log) if args.master_log else newest_master_log(data / "master-log")
    rows = read_rows(path)
    res = compute(rows, args.agent_runs)
    recheck = compute(judged(rows), args.agent_runs) if res["coverage"]["rechecked_rows"] else None
    lead_files = sorted({r[k] for r in rows for k in ("Lead file", "Lead re-check file") if r.get(k)})
    source = {"master_log": path.name, "lead_return": ", ".join(f"`{f}`" for f in lead_files),
              "built_at": datetime.now().astimezone().isoformat(timespec="seconds")}
    if not args.no_json:
        out = next_version_path(data / "master-log", f"results-{datetime.now():%Y-%m-%d}", ".json")
        out.write_text(json.dumps({"source": source, **res, "with_recheck": recheck}, indent=2, default=str) + "\n", encoding="utf-8")
        print(f"WROTE {out.name}")
    text = markdown(res, source, public=args.public, recheck=recheck)
    Path(args.md).write_text(text, encoding="utf-8", newline="\n")
    print(text)
    print(f"WROTE {Path(args.md).name}")


if __name__ == "__main__":
    main()
