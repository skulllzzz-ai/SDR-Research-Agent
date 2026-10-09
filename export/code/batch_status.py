"""Where the agent batch stands, in aggregates: pair companies and SDR-alone companies done, valid
rows, agent minutes, usage-window waits, the company in progress. No company names.

    python batch_status.py                  # the batch folder data-exports/agent-runs/batch-v1
    python batch_status.py --wait-pairs     # returns once the 15 pair companies are done, or the batch has ended

--wait-pairs is the watcher: it looks every minute, and also stops when the batch's console log
(<folder>.stdout.txt, if the batch was started with one) shows the batch has ended.
"""

import argparse
import json
import sys
import time
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent


def status(folder):
    draw = json.loads(next((HERE / "data-exports" / "master-log").glob("draw-*.json")).read_text(encoding="utf-8"))
    arms = {c["company_id"]: c["arm"] for c in draw["companies"]}
    runs = {p.parent.name: json.loads(p.read_text(encoding="utf-8")) for p in folder.glob("*/run.json")}
    in_progress = sorted(p.name for p in folder.iterdir() if p.is_dir() and p.name in arms and p.name not in runs)
    waits = (folder / "usage-waits.jsonl").read_text(encoding="utf-8").splitlines() if (folder / "usage-waits.jsonl").exists() else []
    out = {"at": datetime.now().astimezone().isoformat(timespec="seconds"), "waits": len(waits), "in_progress": in_progress}
    for arm, label in (("With agent", "pair"), ("SDR alone", "sdr_alone")):
        done = [r for cid, r in runs.items() if arms.get(cid) == arm]
        out[label] = {"of": sum(a == arm for a in arms.values()), "done": len(done),
                      "valid": sum(r["validator_result"] == "valid" for r in done),
                      "minutes": round(sum(r["agent_minutes"] for r in done), 1),
                      "first_start": min((r["agent_start"] for r in done), default=None),
                      "last_end": max((r["agent_end"] for r in done), default=None),
                      "repaired": sum(r["retry_count"] > 0 for r in done),
                      "timeouts": sum(r["status"] == "timeout" for r in done),
                      "apollo_calls": sum(sum(r["apollo_calls_that_ran"].values()) for r in done)}
    return out


def ended(folder):
    log = folder.parent / f"{folder.name}.stdout.txt"
    return log.exists() and any(line.startswith("exit ") for line in log.read_text(encoding="utf-8", errors="replace").splitlines())


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("--folder", default=str(HERE / "data-exports" / "agent-runs" / "batch-v1"))
    parser.add_argument("--wait-pairs", action="store_true")
    args = parser.parse_args()
    folder = Path(args.folder)
    while args.wait_pairs:
        if folder.exists():
            now = status(folder)
            if now["pair"]["done"] >= now["pair"]["of"] or ended(folder):
                break
        elif ended(folder):
            break
        time.sleep(60)
    print(json.dumps(status(folder), indent=2) if folder.exists() else f"{folder} does not exist yet")


if __name__ == "__main__":
    main()
