"""Build the trial's public export: data-exports/public-export/ (DECISIONS.md 18).

    python build_public_export.py --draw data-exports/master-log/draw-2026-10-01.json \
        --agent-runs data-exports/agent-runs/batch-v1 data-exports/agent-runs/batch-v2 \
        --smoke data-exports/smoke-slice/smoke-v1web-2026-10-06 data-exports/smoke-slice/smoke-v2-2026-10-07

What goes out (Mayank's ruling, 9 Oct):
  - the 30 companies by name, each with the agent's v1 row and, for the 7 run again, its v2 row: every research field
    with its marker and source as the agent wrote it, the flag with its reason, the priority; the two smoke-test
    companies alongside, on the frozen v1 and v2;
  - per company: the SDR's final verdict with its dead-end kind, the agent's minutes, the lead's scores and scoring
    minutes on the agent's row;
  - contact names and roles as the agent wrote them; every email and every phone number replaced by "[withheld]",
    wherever it stands (a field, a research note, a source link);
  - the CRM status masked;
  - the SDR's research cells, the SDR's minutes per company and the lead's scores on the SDR's rows only as aggregates by arm,
    market and difficulty (RESULTS.md, by results.py --public);
  - the code (with the prompts, the competitor block left as its placeholder), the decisions log, the build log (per-row
    SDR minutes taken out, each place named in the manifest) and a README.
What stays out: the list file, the reserves, the competitor block, the draw record, every workbook, every stream,
answer and assembled prompt, the lead key, the release and return logs, the master logs.
A rebuild moves the previous export to data-exports/public-export-previous/<built at>/; nothing is overwritten.
check_public_export.py, the contact-and-SDR scrub gate, runs at the end; the export is not to be shared unless it passes.
"""

import argparse
import csv
import hashlib
import json
import re
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

from contract import ARM_AGENT, FIELD_KEYS, RESEARCH_FIELDS, ROW_AGENT_ALONE, ROW_PAIR, ROW_SDR_ALONE, VERDICT_DEAD_END
from results import newest_master_log, read_rows

HERE = Path(__file__).resolve().parent
WITHHELD = "[withheld]"
MASKED = "[masked]"
EMAIL = re.compile(r"(?:mailto:)?[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
PHONE = re.compile(r"(?:tel:)?(?<![\w])(?:\+\d{1,3}[\s.\-]?)?(?:\(\d{1,4}\)[\s.\-]?)?\d{1,4}(?:[\s.\-]\d{1,5}){1,5}(?![\w])")
NOT_A_PHONE = re.compile(r"\d{4}-\d{2}-\d{2}(?:-\d+)?|\d{2}\.\d{2}\.\d{4}|(?:19|20)\d{2}\s?[-–]\s?(?:19|20)\d{2}")
CODE = ["agent_config.json", "excel_open_check.ps1", "prompt/research-agent-v1.md", "prompt/research-agent-v2.md", "prompt/repair-v1.md"]
DOCS = ["DECISIONS.md", "BUILD-LOG.md"]
# Per-row SDR minutes in the build log, replaced in the public copy by the aggregate they come to. Each must match once.
REDACTIONS = {
    "BUILD-LOG.md": [
        ("SDR minutes per row 1, 1, 14, 32, 44, 44 and one unreadable",
         "SDR minutes per row withheld in the public copy (6 readable rows: median 23, range 1 to 44) and one unreadable"),
        ("(19% to 34% at any score; one took 61 minutes)", "(19% to 34% at any score)"),
    ],
}
COMPANY_COLUMNS = [
    "Company", "Market", "Set", "Arm", "Agent rows",
    "v1: agent minutes", "v1: tool calls", "v1: verdict", "v1: dead-end kind", "v1: flag (likely dead end)", "v1: flag reason", "v1: priority",
    "v2: agent minutes", "v2: tool calls", "v2: verdict", "v2: dead-end kind", "v2: flag (likely dead end)", "v2: flag reason", "v2: priority",
    "SDR final verdict", "SDR dead-end kind", "SDR row marked Complete",
    "Lead on the agent's v1 row: Acceptable", "Lead: Correctness", "Lead: Completeness", "Lead: scoring minutes", "CRM status",
]
ROW_COLUMNS = ["Company", "Market", "Set", "Agent version", "Field", "Value", "Marker", "Source"]


def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def scrub(text, counts, where):
    """Every email and every phone number in a text, replaced by [withheld]; dates and year spans are left alone."""
    if not text:
        return text
    text, n = EMAIL.subn(WITHHELD, text)
    if n:
        counts[f"email in {where}"] = counts.get(f"email in {where}", 0) + n

    def phone(m):
        digits = re.sub(r"\D", "", m.group(0))
        if 7 <= len(digits) <= 15 and not NOT_A_PHONE.fullmatch(m.group(0).replace("tel:", "").strip()):
            counts[f"phone in {where}"] = counts.get(f"phone in {where}", 0) + 1
            return WITHHELD
        return m.group(0)
    return PHONE.sub(phone, text)


def agent_run(folder):
    folder = Path(folder)
    row = json.loads((folder / "row.json").read_text(encoding="utf-8"))
    run = json.loads((folder / "run.json").read_text(encoding="utf-8"))
    return row, run


def row_lines(company, market, set_name, version, row, counts):
    """The long-format lines of one agent row: 25 research fields with marker and source, then the flag, the verdict and
    the CRM status. Emails and phones are withheld whole; every other text is scrubbed."""
    lines = []
    for name, key in FIELD_KEYS.items():
        f = row["fields"][key]
        value = f["value"].strip()
        if name in ("Emails", "Phones"):
            if value:
                counts[f"{name} field withheld"] = counts.get(f"{name} field withheld", 0) + 1
            value = WITHHELD if value else ""
        else:
            value = scrub(value, counts, name)
        lines.append([company, market, set_name, version, name, value, f["marker"], scrub(f["source"].strip(), counts, f"{name} source")])
    flag, verdict = row["flag"], row["verdict"]
    dead = verdict["value"] == VERDICT_DEAD_END
    lines += [
        [company, market, set_name, version, "Flag: likely dead end", "Yes" if flag["likely_dead_end"] == "yes" else "No", "",
         scrub(flag["source"].strip(), counts, "flag source")],
        [company, market, set_name, version, "Flag reason", scrub(flag["reason"].strip(), counts, "flag reason"), "", ""],
        [company, market, set_name, version, "Verdict", verdict["value"], "", scrub(verdict["source"].strip(), counts, "verdict source")],
        [company, market, set_name, version, "Dead-end kind", verdict["dead_end_type"] if dead else "", "", ""],
        [company, market, set_name, version, "Verdict reason", scrub(verdict["reason"].strip(), counts, "verdict reason"), "", ""],
        [company, market, set_name, version, "CRM status", MASKED, "", ""],
    ]
    return lines


def run_cells(prefix, row, run, counts):
    if row is None:
        return {f"{prefix}: {k}": "" for k in ("agent minutes", "tool calls", "verdict", "dead-end kind", "flag (likely dead end)",
                                                  "flag reason", "priority")}
    dead = row["verdict"]["value"] == VERDICT_DEAD_END
    return {f"{prefix}: agent minutes": run["agent_minutes"], f"{prefix}: tool calls": len(run["tool_calls"]),
            f"{prefix}: verdict": row["verdict"]["value"], f"{prefix}: dead-end kind": row["verdict"]["dead_end_type"] if dead else "",
            f"{prefix}: flag (likely dead end)": "Yes" if row["flag"]["likely_dead_end"] == "yes" else "No",
            f"{prefix}: flag reason": scrub(row["flag"]["reason"].strip(), counts, "flag reason"),
            f"{prefix}: priority": row["fields"]["priority"]["value"].strip()}


def write_csv(path, columns, rows):
    with open(path, "w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(columns)
        writer.writerows(rows)


def build(args):
    data = Path(args.data_root)
    out = Path(args.out)
    draw = json.loads(Path(args.draw).read_text(encoding="utf-8"))
    master = Path(args.master_log) if args.master_log else newest_master_log(data / "master-log")
    rows = read_rows(master)
    by_row = {(r["Company ID"], r["Row type"]): r for r in rows}
    v1_folder, v2_folder = (Path(f) for f in args.agent_runs)
    if out.exists() and any(out.iterdir()):
        manifest = out / "EXPORT-MANIFEST.json"
        if not manifest.exists():
            raise SystemExit(f"{out} holds files that are not an export: nothing was moved or written")
        stamp = json.loads(manifest.read_text(encoding="utf-8"))["built_at"].replace(":", "").replace("+", "_")
        previous = data / "public-export-previous" / stamp
        previous.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(out), str(previous))
        print(f"MOVED the previous export to {previous.relative_to(data)}")
    out.mkdir(parents=True, exist_ok=True)
    counts = {}

    # per company and per agent row
    company_rows, agent_lines = [], []
    for c in sorted(draw["companies"], key=lambda c: c["company_name"].lower()):
        cid, name, market = c["company_id"], c["company_name"], c["market"]
        row1, run1 = agent_run(v1_folder / cid)
        row2 = run2 = None
        if (v2_folder / cid / "row.json").exists():
            row2, run2 = agent_run(v2_folder / cid)
        human = by_row[(cid, ROW_PAIR if c["arm"] == ARM_AGENT else ROW_SDR_ALONE)]
        alone = by_row[(cid, ROW_AGENT_ALONE)]
        scored = bool(alone["Acceptable"])
        cells = {"Company": name, "Market": market, "Set": "Trial", "Arm": "SDR with agent" if c["arm"] == ARM_AGENT else "SDR alone",
                 "Agent rows": "v1, v2" if row2 else "v1", **run_cells("v1", row1, run1, counts), **run_cells("v2", row2, run2, counts),
                 "SDR final verdict": human["Verdict"] or "", "SDR dead-end kind": (human["Dead-end category"] or "") if human["Verdict"] == VERDICT_DEAD_END else "",
                 "SDR row marked Complete": "Yes" if human["SDR status"] == "Complete" else "No (waiting for the SDR)",
                 "Lead on the agent's v1 row: Acceptable": alone["Acceptable"] if scored else "not scored",
                 "Lead: Correctness": alone["Correctness"] if scored else "not scored",
                 "Lead: Completeness": alone["Completeness"] if scored else "not scored",
                 "Lead: scoring minutes": alone["Lead scoring minutes"] if scored else "not scored", "CRM status": MASKED}
        company_rows.append([cells[k] for k in COMPANY_COLUMNS])
        agent_lines += row_lines(name, market, "Trial", "v1", row1, counts)
        if row2:
            agent_lines += row_lines(name, market, "Trial", "v2", row2, counts)
    smoke = {}
    for folder in args.smoke:
        for sub in sorted(Path(folder).glob("S*/row.json")):
            row, run = agent_run(sub.parent)
            smoke.setdefault(run["company_name"], {"market": run["market"]})[run["prompt_version"]] = (row, run)
    for name, versions in sorted(smoke.items()):
        row1, run1 = versions.get("v1", (None, None))
        row2, run2 = versions.get("v2", (None, None))
        cells = {"Company": name, "Market": versions["market"], "Set": "Smoke test (outside the 30)", "Arm": "",
                 "Agent rows": ", ".join(v for v in ("v1", "v2") if v in versions), **run_cells("v1", row1, run1, counts),
                 **run_cells("v2", row2, run2, counts), "SDR final verdict": "", "SDR dead-end kind": "", "SDR row marked Complete": "",
                 "Lead on the agent's v1 row: Acceptable": "", "Lead: Correctness": "", "Lead: Completeness": "",
                 "Lead: scoring minutes": "", "CRM status": MASKED}
        company_rows.append([cells[k] for k in COMPANY_COLUMNS])
        for version, (row, _) in sorted((v, x) for v, x in versions.items() if v != "market"):
            agent_lines += row_lines(name, versions["market"], "Smoke test", version, row, counts)
    write_csv(out / "companies.csv", COMPANY_COLUMNS, company_rows)
    write_csv(out / "agent-rows.csv", ROW_COLUMNS, agent_lines)

    # the code, the prompts and the logs
    for name in sorted(p.name for p in HERE.glob("*.py")) + CODE:
        target = out / "code" / name
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(HERE / name, target)
    shutil.copyfile(HERE / "README.md", out / "code" / "README-build.md")
    redacted = []
    for name in DOCS:
        text = (HERE / name).read_text(encoding="utf-8")
        for old, new in REDACTIONS.get(name, []):
            assert text.count(old) == 1, f"{name}: a redaction no longer matches once: {old!r}"
            text = text.replace(old, new)
            redacted.append({"file": name, "replaced": old, "with": new})
        (out / name).write_text(text, encoding="utf-8", newline="\n")

    # the results, public cut
    done = subprocess.run([sys.executable, str(HERE / "results.py"), "--agent-runs", str(v1_folder), str(v2_folder), "--master-log", str(master),
                           "--public", "--no-json", "--md", str(out / "RESULTS.md"), "--data-root", str(data)],
                          capture_output=True, text=True, encoding="utf-8")
    if done.returncode:
        raise SystemExit(f"results.py --public failed:\n{done.stdout[-2000:]}{done.stderr[-2000:]}")

    trial = [r for r in company_rows if r[COMPANY_COLUMNS.index("Set")] == "Trial"]
    facts = {
        "companies": len(trial), "smoke_companies": len(smoke), "v2_rows": sum(r[COMPANY_COLUMNS.index("Agent rows")] == "v1, v2" for r in trial),
        "agent_row_lines": len(agent_lines), "scored_agent_rows": sum(r[COMPANY_COLUMNS.index("Lead: Correctness")] != "not scored" for r in trial),
        "waiting": sum(r[COMPANY_COLUMNS.index("SDR row marked Complete")].startswith("No") for r in trial),
        "rechecked": sum(r.get("Dead end right (lead re-check)") in ("Yes", "No") for r in rows),
    }
    (out / "README.md").write_text(readme(facts, counts, redacted, master, rows), encoding="utf-8", newline="\n")
    manifest = {"built_at": datetime.now().astimezone().isoformat(timespec="seconds"), "master_log": master.name,
                "lead_returns": sorted({r[k] for r in rows for k in ("Lead file", "Lead re-check file") if r.get(k)}),
                "agent_runs": [v1_folder.name, v2_folder.name], "smoke": [Path(f).name for f in args.smoke], "facts": facts,
                "withheld": counts, "redacted": redacted,
                "files": {str(p.relative_to(out)).replace("\\", "/"): sha256(p) for p in sorted(out.rglob("*")) if p.is_file()}}
    (out / "EXPORT-MANIFEST.json").write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"PUBLIC EXPORT {out}: {facts['companies']} companies and {facts['smoke_companies']} smoke companies, "
          f"{facts['agent_row_lines']} agent-row lines; withheld {sum(counts.values())} contact items; {len(redacted)} build-log redactions")
    return out


def readme(facts, counts, redacted, master, rows):
    withheld = sum(v for k, v in counts.items())
    lines = [
        "# SDR research agent trial: public export",
        "",
        "An AI agent did the first research pass on a prospect company for a pre-sales SDR in B2B diamonds, and the SDR checked and "
        "completed it. Three arms on 30 companies in the USA and the Gulf: the SDR alone, the agent alone, and the SDR starting from the "
        "agent's row. The SDR lead scored every finished row blind. This folder is what can be shared.",
        "",
        "## What is in it",
        "",
        f"- `companies.csv`: the {facts['companies']} companies of the trial by name and the {facts['smoke_companies']} smoke-test companies, "
        "one line each: the arm; the agent's minutes, tool calls, verdict, flag with its reason and priority on v1 and, for the "
        f"{facts['v2_rows']} companies run again, on v2; the SDR's final verdict with its dead-end kind; the lead's scores and scoring "
        f"minutes on the agent's v1 row ({facts['scored_agent_rows']} scored; the others were never sent to the lead).",
        "- `agent-rows.csv`: every agent row in full, one line per field: the 25 research fields with their marker (sure, guessed, not "
        "found) and source as the agent wrote them, then the flag, its reason, the verdict, its dead-end kind and reason, and the CRM status.",
        "- `RESULTS.md`: the results by arm, by market and by difficulty, the five conditions of the hypothesis, the agent's cost and use "
        "of the usage window. Aggregates only.",
        "- `code/`: the scripts that drew the arms, ran and locked the agent, validated its rows, built and read every exchange file, and "
        "computed the results; `code/prompt/` holds the prompts; `code/README-build.md` is the build's own guide.",
        "- `DECISIONS.md` and `BUILD-LOG.md`: every rule decided during the trial and every build pass, as kept during the work.",
        "- `EXPORT-MANIFEST.json`: when and from what this was built, what was withheld, and a hash of every file.",
        "",
        "## What was withheld, and how",
        "",
        f"- Every email address and every phone number, wherever it stood (a field, a research note, a source link), is replaced by "
        f"`{WITHHELD}`: {withheld} items. Contact names and roles stay as the agent wrote them, each with its public source.",
        f"- The CRM status is masked (`{MASKED}`): whether a company is already a customer stays private.",
        "- The SDR's own research cells, the SDR's minutes per company and the lead's scores on the SDR's rows appear only as aggregates by arm, "
        "market and difficulty in `RESULTS.md`, over groups of 3 rows or more. The SDR's final verdict and its dead-end kind are shown "
        "per company.",
        "- The company list the companies were drawn from, the reserve companies and the competitor list stay out. The prompts keep "
        "`{{COMPETITOR_BLOCK}}` where the competitor list was put in at run time.",
    ]
    if redacted:
        lines.append(f"- In `BUILD-LOG.md`, {len(redacted)} places that gave the SDR's minutes row by row now give the aggregate instead "
                     "(listed in the manifest).")
    lines += [
        "- Not included: the workbooks the SDR and the lead worked in, the master logs, the agent's raw streams and answers, the "
        "assembled prompts, the key that maps the lead's blind row codes to the arms.",
        "",
        "## How to read it",
        "",
        "- Arms: \"SDR alone\" companies were researched by the SDR from scratch; \"SDR with agent\" companies were researched by the "
        "agent first, then checked and completed by the SDR. Every company also has an agent-alone row: the agent's v1 row, which the "
        "lead scored beside the human row without knowing which was which.",
        "- v1 is the agent as frozen before the batch (web search and web fetch only, no paid data). v2 changed the prompt after the "
        "SDR's first checks (DECISIONS.md 14) and was run only on the last 7 \"SDR with agent\" companies; the agent-alone arm stays v1.",
        "- The flag is the agent's own early call that a company is likely a dead end (a competitor, a lab-grown-only seller in the "
        "Gulf, or a non-buyer segment). The verdict is the row's final call: Prospect or Dead end.",
        "- The lead's scores: Acceptable (would you let the SDR call from this row?), Correctness (wrong facts in the key fields: None, "
        "One, More than one) and Completeness (Full, Partial, Thin).",
        f"- {facts['waiting']} \"SDR with agent\" companies have no completed SDR row yet, so neither of their rows was scored.",
    ]
    if facts["rechecked"]:
        lines.append(f"- The {facts['rechecked']} dead-end rows were judged again on the verdict alone (DECISIONS.md 18); RESULTS.md leads "
                     "with that reading.")
    else:
        lines.append("- The dead-end rows were scored on \"would you let the SDR call from this row?\", which a right dead end answers No; "
                     "a re-check on the verdict alone was asked for (DECISIONS.md 18). RESULTS.md shows both readings and why.")
    lines += [
        "",
        "## Limits",
        "",
        "- One SDR, one lead, one model, 30 companies: every rate moves about 7 points per row in an arm of 15. No same-company "
        "comparison is beyond chance at this size.",
        "- The lead spent about a minute per row: the scores are a read of each row, not a check of its sources.",
        "- The v1 and v2 pair rows are different companies, so the gap between them is not a pure version effect.",
        "- The agent's minutes describe this build on these days (model, tools, connection), not the method.",
        "",
        f"Built from `{master.name}` by `code/build_public_export.py`; the contact-and-SDR scrub gate is `code/check_public_export.py`.",
        "",
    ]
    return "\n".join(lines)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("--draw", required=True)
    parser.add_argument("--agent-runs", nargs=2, required=True, help="the v1 batch, then the v2 batch")
    parser.add_argument("--smoke", nargs="*", default=[], help="smoke-test folders on the frozen versions")
    parser.add_argument("--master-log")
    parser.add_argument("--data-root", default=str(HERE / "data-exports"))
    parser.add_argument("--out", default=str(HERE / "data-exports" / "public-export"))
    parser.add_argument("--no-gate", action="store_true", help="tests only: skip the gate")
    args = parser.parse_args()
    out = build(args)
    if not args.no_gate:
        done = subprocess.run([sys.executable, str(HERE / "check_public_export.py"), str(out), "--draw", args.draw,
                               "--master-log", str(Path(args.master_log) if args.master_log else newest_master_log(Path(args.data_root) / "master-log")),
                               "--data-root", args.data_root], text=True, encoding="utf-8")
        sys.exit(done.returncode)


if __name__ == "__main__":
    main()
