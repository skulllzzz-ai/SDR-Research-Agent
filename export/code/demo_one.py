"""Run the agent live on one company, for a demo video: each search and page read as it happens, then the row.

    python demo_one.py --company "<company name>" --market USA
    python demo_one.py --replay data-exports/agent-runs/batch-v1/C01/stream.jsonl     # a recorded run, no agent call

It runs the pinned agent of agent_config.json (the frozen version: web search and web fetch only, no paid data, no
credits) through run_agent.py's research mode, with the same lock as the trial, into a folder of its own,
data-exports/smoke-slice/demo-YYYYMMDD-HHMM/, so no trial file is touched. Emails and phone numbers show as [withheld]
on screen, so the recording can be shared; the full row stays in the folder. One run takes about 2 to 4 minutes and
about 2 to 3 points of the 5-hour usage window. The terminal's own claude login is used (check: claude auth status).
"""

import argparse
import csv
import json
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

from build_public_export import scrub
from contract import FIELD_KEYS, MARKETS

HERE = Path(__file__).resolve().parent
LABEL = {"WebSearch": "search", "WebFetch": "read  "}


def screen(text):
    return scrub(str(text or ""), {}, "screen")


def tool_calls(line):
    """The tool calls in one line of a run's event stream (written as {"t", "event"})."""
    try:
        event = json.loads(line).get("event", {})
    except (json.JSONDecodeError, AttributeError):
        return []
    message = event.get("message")
    if event.get("type") != "assistant" or not isinstance(message, dict):
        return []
    return [(b.get("name"), b.get("input") or {}) for b in message.get("content") or []
            if isinstance(b, dict) and b.get("type") == "tool_use"]


def show(tool, args, started):
    what = args.get("query") or args.get("url") or json.dumps(args)[:120]
    minutes, seconds = divmod(int(time.monotonic() - started), 60)
    print(f"  {minutes:02d}:{seconds:02d}  {LABEL.get(tool, tool)}  {screen(what)}", flush=True)


def follow(stream, stop, started):
    """Prints each tool call from the run's stream as it lands; a half-written line is read again once complete."""
    while not stream.exists():
        if stop.is_set():
            return
        time.sleep(0.5)
    with open(stream, encoding="utf-8") as handle:
        while True:
            where = handle.tell()
            line = handle.readline()
            if not line.endswith("\n"):
                handle.seek(where)
                if stop.is_set():
                    return
                time.sleep(0.3)
                continue
            for tool, args in tool_calls(line):
                show(tool, args, started)


def summary(folder, cid):
    run_path, row_path = folder / cid / "run.json", folder / cid / "row.json"
    if not run_path.exists():
        print("NO RUN RECORD: the run did not finish; see the runner's output above.")
        return
    run = json.loads(run_path.read_text(encoding="utf-8"))
    tools = run["summary"]["tools_called"]
    print(f'\nRESULT  {run["agent_minutes"]} min, {len(run["tool_calls"])} tool calls ({tools.get("WebSearch", 0)} searches, '
          f'{tools.get("WebFetch", 0)} page reads), {sum(run["apollo_calls_that_ran"].values())} paid-data calls, '
          f'validator: {run["validator_result"]}, prompt {run["prompt_version"]}, {run["model_asked"]} at {run["effort"]}')
    if not row_path.exists():
        print(f'        no valid row: {run["validator_problems"]}')
        return
    row = json.loads(row_path.read_text(encoding="utf-8"))
    fields = row["fields"]
    found = [k for k in FIELD_KEYS.values() if fields[k]["value"].strip()]
    markers = {m: sum(fields[k]["marker"] == m for k in found) for m in ("sure", "guessed")}
    flag, verdict = row["flag"], row["verdict"]
    print(f'        verdict: {verdict["value"]}{" (" + verdict["dead_end_type"] + ")" if verdict["value"] == "Dead end" else ""}; '
          f'flag, likely dead end: {"yes" if flag["likely_dead_end"] == "yes" else "no"}: {screen(flag["reason"])[:160]}')
    print(f'        fields found: {len(found)} of {len(FIELD_KEYS)} (sure {markers["sure"]}, guessed {markers["guessed"]}); '
          f'priority {fields["priority"]["value"] or "not found"} ({fields["priority"]["marker"]})')
    print(f'        contacts: {screen(fields["contact_names"]["value"]).replace(chr(10), "; ") or "not found"}')
    print(f'        emails: {"[withheld]" if fields["emails"]["value"].strip() else "not found"}; '
          f'phones: {"[withheld]" if fields["phones"]["value"].strip() else "not found"}')
    print(f"        the full row: {row_path.resolve().relative_to(HERE)}")


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("--company", help="the company's name, as you would give it to the SDR")
    parser.add_argument("--market", choices=MARKETS)
    parser.add_argument("--replay", help="a recorded run's stream.jsonl: print its searches and page reads, no agent call")
    args = parser.parse_args()
    started = time.monotonic()
    if args.replay:
        for line in Path(args.replay).read_text(encoding="utf-8").splitlines():
            for tool, call in tool_calls(line):
                show(tool, call, started)
        return
    if not (args.company and args.market):
        parser.error("give --company and --market (or --replay)")
    folder = HERE / "data-exports" / "smoke-slice" / f"demo-{datetime.now():%Y%m%d-%H%M}"
    if folder.exists():
        raise SystemExit(f"{folder.name} exists: wait a minute and run again (a run is never overwritten)")
    folder.mkdir(parents=True)
    cid = "D01"
    with open(folder / "company.csv", "w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=["company_id", "company_name", "market"])
        writer.writeheader()
        writer.writerow({"company_id": cid, "company_name": args.company, "market": args.market})
    print(f"DEMO  {args.company} ({args.market}): the pinned research agent, live. It checks its lock first, then each "
          "search and page read shows as it happens.\n", flush=True)
    stop = threading.Event()
    watcher = threading.Thread(target=follow, args=(folder / cid / "stream.jsonl", stop, started), daemon=True)
    watcher.start()
    done = subprocess.run([sys.executable, str(HERE / "run_agent.py"), "research", "--companies", str(folder / "company.csv"),
                           "--out", str(folder)], capture_output=True, text=True, encoding="utf-8", errors="replace")
    stop.set()
    watcher.join(timeout=5)
    (folder / "runner-output.txt").write_text(done.stdout + done.stderr, encoding="utf-8")
    if done.returncode:
        print(f"\nTHE RUNNER STOPPED (exit {done.returncode}):\n" + "\n".join((done.stdout + done.stderr).strip().splitlines()[-12:]))
    summary(folder, cid)


if __name__ == "__main__":
    main()
