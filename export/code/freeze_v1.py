"""Freeze the agent as the version agent_config.json names (spec exit condition 9; v1, then v2 on
DECISIONS.md 14): the template, repair-prompt and competitor-list
hashes, the model, the effort, the timeout and the tool lists, recorded together once, with the
model decision and its evidence (the smoke test, the lock test, the slice and the failure path).

    python freeze_v1.py --smoke <folder> --lock-test <folder> --failure-path <folder>
                        (--slice <folder> [--slice-return <file.xlsx>] | --slice-pending "<why>")
                        [--compare <folder>]... [--reason "<why these sources>"] [--credit-readings '<json list>']
    python freeze_v1.py --self-test

Writes data-exports/master-log/freeze-<version>.json and never overwrites it; --revision '<json>' records
the decision a later version rests on. From then on run_agent.py
runs the 30 drawn companies only while the agent still matches this record (check_freeze), and
refuses them before it exists. A changed agent needs its own record (v2).
"""

import argparse
import json
import shutil
import sys
import tempfile
from collections import Counter
from datetime import datetime
from pathlib import Path

import run_agent as ra

KEEP_RUN = ("company_id", "market", "agent_minutes", "status", "validator_result", "retry_count",
            "apollo_calls_that_ran", "apollo_cap_respected", "calls_blocked_by_the_cap", "canary")


def load(path):
    path = Path(path)
    if not path.exists():
        raise SystemExit(f"REFUSED: {path} is missing")
    return json.loads(path.read_text(encoding="utf-8"))


def markers(folder, company_id):
    row = Path(folder) / company_id / "row.json"
    if not row.exists():
        return {}
    return dict(Counter(v["marker"] for v in json.loads(row.read_text(encoding="utf-8"))["fields"].values()))


def pinned_batch(folder, label, minimum):
    """A batch folder run by the pinned agent: not a probe, the config being frozen, all rows valid."""
    batch = load(Path(folder) / "batch.json")
    problems = []
    if batch.get("probe_not_agent_rows"):
        problems.append("it is a mechanics check, not the pinned agent")
    if batch.get("config") != ra.CONFIG:
        problems.append("it ran with a different agent_config.json")
    if (batch.get("model"), batch.get("effort")) != (ra.CONFIG["model"], ra.CONFIG["effort"]):
        problems.append(f'it ran {batch.get("model")} at {batch.get("effort")}')
    stale = [role for role, value in (batch.get("file_sha256") or {}).items()
             if role not in ra.NON_BLOCKING and ra.file_hashes().get(role) != value]
    if stale:
        problems.append(f"these files changed since it ran: {stale}")
    runs = batch.get("runs", [])
    if len(runs) < minimum or any(r["validator_result"] != "valid" for r in runs):
        problems.append(f'{sum(r["validator_result"] == "valid" for r in runs)} valid rows of {len(runs)}; at least {minimum}, all valid')
    if problems:
        raise SystemExit(f"REFUSED: the {label} evidence in {folder} does not qualify: " + "; ".join(problems))
    return batch


def summary(folder, batch):
    return {"folder": Path(folder).name, "model": batch.get("model"), "effort": batch.get("effort"),
            "probe": batch.get("probe_not_agent_rows"),
            "runs": [{**{k: r.get(k) for k in KEEP_RUN}, "markers": markers(folder, r["company_id"])} for r in batch.get("runs", [])]}


def make_record(smoke, lock_test, slice_, failure_path, compare=None, slice_return=None, slice_pending=None,
                reason=None, credit_readings=None, revision=None):
    """compare: one folder or a list of them. slice_ may be None only with slice_pending, the reason
    the freeze comes before the slice (DECISIONS.md 8)."""
    smoke_batch = pinned_batch(smoke, "smoke-test", 2)
    if slice_ is None and not slice_pending:
        raise SystemExit("REFUSED: no slice evidence and no reason why the freeze comes before the slice")
    slice_batch = pinned_batch(slice_, "slice", 1) if slice_ is not None else None
    report = load(Path(lock_test) / "lock-test.json")
    if report.get("model") != ra.CONFIG["model"] or not report.get("checks") or not all(report["checks"].values()):
        raise SystemExit(f"REFUSED: the lock test in {lock_test} did not pass every check on {ra.CONFIG['model']}")
    lock = report["lock"]
    assembled, hashes = ra.assemble_prompt(lock)
    for label, batch in (("smoke-test", smoke_batch), ("slice", slice_batch)):
        if batch is None:
            continue
        if batch["hashes"]["assembled_prompt_sha256"] != hashes["assembled_prompt_sha256"]:
            raise SystemExit(f"REFUSED: the {label} ran a different assembled prompt from the one this freeze would pin")
        if batch["lock"]["allowed"] != lock["allowed"]:
            raise SystemExit(f"REFUSED: the {label} ran a different allow list from the lock test's")
    failure = load(Path(failure_path) / "batch.json")
    failed = [r for r in failure.get("runs", []) if r["status"] == "timeout" or r["validator_result"] == "rejected"
              or markers(failure_path, r["company_id"]).get("not found", 0) >= 8]
    if not failed:
        raise SystemExit(f"REFUSED: {failure_path} holds no failure path (a timeout, a rejected row, or a thin-data row "
                         "with 8 or more not-found markers)")
    checking = "pending: recorded when the SDR's slice return is filed"
    compares = [compare] if isinstance(compare, (str, Path)) else list(compare or [])
    if slice_return:
        import sdr_return
        rows = sdr_return.read_return(slice_return)["rows"]
        checking = rows[0]["pass1_minutes"] if rows else "unreadable"
    return {
        "version": ra.VERSION,
        "revision": revision,
        "frozen_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "model": ra.CONFIG["model"], "effort": ra.CONFIG["effort"], "timeout_minutes": ra.CONFIG["timeout_minutes"],
        "config": ra.CONFIG,
        "file_sha256": ra.file_hashes(),
        "prompt_sha256": hashes,
        "assembled_prompt_sha256": hashes["assembled_prompt_sha256"],
        "lock": {k: lock[k] for k in ("allowed", "denied_by_name", "denied_servers", "apollo_server", "default_deny_test_tool")},
        "decision": {"model": ra.CONFIG["model"], "effort": ra.CONFIG["effort"], "by": "Mayank", "on": "2026-10-05",
                     "where": "DECISIONS.md, decision 6", "fallback": "never automatic (spec)",
                     "tools": lock["allowed"],
                     "apollo": ("off the allow list, denied as a whole, credit tools capped at 0 in the hook; 0 credits"
                                if not lock["apollo_allowed"] else f'{len(lock["apollo_allowed"])} read tools allowed, capped per company'),
                     "sources_decided": "DECISIONS.md, decision 7 (Mayank, 2026-10-06)" if not lock["apollo_allowed"] else None,
                     "reason": reason, "credit_readings": credit_readings or []},
        "evidence": {
            "smoke_test": summary(smoke, smoke_batch),
            "compare": [summary(folder, load(Path(folder) / "batch.json")) for folder in compares],
            "lock_test": {"folder": Path(lock_test).name, "checks": report["checks"]},
            "slice": summary(slice_, slice_batch) if slice_batch else {"status": "pending", "why": slice_pending},
            "slice_sdr_checking_minutes": checking,
            "failure_path": {**summary(failure_path, failure), "failed_runs": [r["company_id"] for r in failed]},
        },
    }


def write_record(record, path):
    try:
        with open(path, "x", encoding="utf-8") as handle:
            handle.write(json.dumps(record, indent=2, ensure_ascii=False) + "\n")
    except FileExistsError:
        raise SystemExit(f"REFUSED: {Path(path).name} exists. A freeze is written once; a changed agent is frozen as the next version.")


def self_test():
    """On made-up evidence folders in a temporary folder, with the real config, files and lock."""
    results = []

    def check(name, passed, detail=""):
        results.append(bool(passed))
        print(f'{"PASS" if passed else "FAIL"}  {name}{"  [" + str(detail) + "]" if detail != "" else ""}')

    def refused(fn):
        try:
            fn()
            return False
        except SystemExit as stop:
            return str(stop).startswith("REFUSED")

    # the lock the current config builds, on the stored discovery list of the latest preflight
    stored = next(json.loads(q.read_text(encoding="utf-8")) for q in
                  sorted(ra.HERE.glob("data-exports/smoke-slice/*/lock.json"), key=lambda q: q.stat().st_mtime, reverse=True)
                  if "apollo_tools" in json.loads(q.read_text(encoding="utf-8"))["lock"])
    lock = ra.build_lock(stored["lock"]["apollo_tools"], stored["lock"]["apollo_tools"], stored["loaded"]["connectors"])
    assembled, hashes = ra.assemble_prompt(lock)
    root = Path(tempfile.mkdtemp(prefix="freeze-selftest-"))

    def batch(name, runs, probe=False, model=None):
        folder = root / name
        folder.mkdir()
        (folder / "batch.json").write_text(json.dumps({
            "probe_not_agent_rows": probe, "model": model or ra.CONFIG["model"], "effort": ra.CONFIG["effort"],
            "config": ra.CONFIG, "hashes": hashes, "file_sha256": ra.file_hashes(), "lock": lock,
            "runs": [{"company_id": cid, "market": "USA", "agent_minutes": 3.0, "status": status, "validator_result": result,
                      "retry_count": 0, "apollo_calls_that_ran": {}, "apollo_cap_respected": True,
                      "calls_blocked_by_the_cap": 0, "canary": {}} for cid, status, result in runs]}), encoding="utf-8")
        return folder

    smoke = batch("smoke", [("T1", "completed", "valid"), ("T2", "completed", "valid")])
    slice_ = batch("slice", [("T3", "completed", "valid")])
    failure = batch("failure", [("T1", "timeout", "rejected")], probe=True)
    probe = batch("probe", [("T1", "completed", "valid"), ("T2", "completed", "valid")], probe=True)
    other_model = batch("other", [("T1", "completed", "valid"), ("T2", "completed", "valid")], model="claude-fable-5-1")
    one_bad = batch("bad", [("T1", "completed", "valid"), ("T2", "completed", "rejected")])
    no_failure = batch("nofail", [("T1", "completed", "valid")])
    lock_dir = root / "lock"
    lock_dir.mkdir()
    (lock_dir / "lock-test.json").write_text(json.dumps({"model": ra.CONFIG["model"], "lock": lock,
                                                         "checks": {"all": True}}), encoding="utf-8")

    record = make_record(smoke, lock_dir, slice_, failure)
    expected = len(ra.CONFIG["builtin_tools"]) + len(ra.CONFIG["apollo_allowed"])
    check("a record is made from qualifying evidence: model, effort, timeout, hashes, lock, decision",
          record["model"] == ra.CONFIG["model"] and len(record["lock"]["allowed"]) == expected
          and record["assembled_prompt_sha256"] == hashes["assembled_prompt_sha256"])
    readings = [{"at": "2099-01-01T00:00:00+05:30", "lead_used": 2470, "lead_limit": 2500, "lead_left": 30}]
    pending = make_record(smoke, lock_dir, None, failure, compare=[smoke, slice_], slice_pending="made-up order",
                          reason="made-up credit reading", credit_readings=readings)
    check("a record with the slice pending carries the reason, the readings and both comparisons",
          pending["evidence"]["slice"] == {"status": "pending", "why": "made-up order"}
          and pending["decision"]["credit_readings"] == readings and pending["decision"]["reason"] == "made-up credit reading"
          and len(pending["evidence"]["compare"]) == 2)
    check("refused: no slice and no reason for its absence", refused(lambda: make_record(smoke, lock_dir, None, failure)))
    check("refused: a smoke test that was a mechanics check", refused(lambda: make_record(probe, lock_dir, slice_, failure)))
    check("refused: a smoke test on another model", refused(lambda: make_record(other_model, lock_dir, slice_, failure)))
    check("refused: a smoke test with a rejected row", refused(lambda: make_record(one_bad, lock_dir, slice_, failure)))
    check("refused: no failure path", refused(lambda: make_record(smoke, lock_dir, slice_, no_failure)))
    path = root / f"freeze-{ra.VERSION}.json"
    write_record(record, path)
    check("refused: a second freeze over the first", refused(lambda: write_record(record, path)))

    frozen, notes = ra.check_freeze(lock, hashes, path=path)
    check("the guard passes the agent that matches the freeze", frozen["version"] == ra.VERSION and not notes, notes)
    check("refused by the guard: no freeze record", refused(lambda: ra.check_freeze(lock, hashes, path=root / "none.json")))
    ra.RUN["probe"] = True
    check("refused by the guard: a mechanics check", refused(lambda: ra.check_freeze(lock, hashes, path=path)))
    ra.RUN["probe"] = False
    ra.CONFIG["effort"], effort = "low", ra.CONFIG["effort"]
    check("refused by the guard: a changed config (effort)", refused(lambda: ra.check_freeze(lock, hashes, path=path)))
    ra.CONFIG["effort"] = effort
    check("refused by the guard: a changed allow list",
          refused(lambda: ra.check_freeze({**lock, "allowed": lock["allowed"][:-1]}, hashes, path=path)))
    check("refused by the guard: a changed assembled prompt",
          refused(lambda: ra.check_freeze(lock, {**hashes, "assembled_prompt_sha256": "0" * 64}, path=path)))
    real_hashes = ra.file_hashes
    ra.file_hashes = lambda: {**real_hashes(), "prompt_template": "0" * 64}
    check("refused by the guard: a changed prompt template", refused(lambda: ra.check_freeze(lock, hashes, path=path)))
    ra.file_hashes = lambda: {**real_hashes(), "runner": "0" * 64}
    frozen, notes = ra.check_freeze(lock, hashes, path=path)
    check("not refused, only noted: a changed runner", notes == ["runner changed since the freeze"], notes)
    ra.file_hashes = real_hashes
    frozen, notes = ra.check_freeze({**lock, "denied_servers": lock["denied_servers"] + ["mcp__claude_ai_New"]}, hashes, path=path)
    check("not refused, only noted: a new connector on the deny list", len(notes) == 1, notes)

    drawn = ra.drawn_companies()
    order = ra.batch_companies(drawn)
    by_id = {c["company_id"]: c for c in drawn}
    arms = [by_id[c["company_id"]]["arm"] for c in order]
    orders = [by_id[c["company_id"]]["work_order"] for c in order]
    check("batch order: the 30 drawn, pair companies first, each half in work order",
          len(order) == 30 and arms == ["With agent"] * 15 + ["SDR alone"] * 15
          and orders[:15] == sorted(orders[:15]) and orders[15:] == sorted(orders[15:]))
    first = drawn[0]
    check("guard trigger: a drawn company is recognised by id and by name, a smoke company is not",
          ra.is_drawn({"company_id": first["company_id"], "company_name": "x"}, drawn)
          and ra.is_drawn({"company_id": "Z9", "company_name": " " + first["company_name"].upper()}, drawn)
          and not ra.is_drawn({"company_id": "S01", "company_name": "Not on the list"}, drawn))
    shutil.rmtree(root, ignore_errors=True)
    print(f"\n{sum(results)} of {len(results)} self-test checks passed")
    return all(results)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    if "--self-test" in sys.argv:
        sys.exit(0 if self_test() else 1)
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    for name in ("--smoke", "--lock-test", "--failure-path"):
        parser.add_argument(name, required=True)
    parser.add_argument("--slice", help="the slice folder; or give --slice-pending")
    parser.add_argument("--slice-pending", help="why the freeze comes before the slice (DECISIONS.md 8)")
    parser.add_argument("--compare", action="append", help="an earlier smoke test or probe, for the decision's evidence")
    parser.add_argument("--slice-return", help="the SDR's returned slice file, for the checking time")
    parser.add_argument("--reason", help="why the agent has these sources")
    parser.add_argument("--credit-readings", help="JSON list of the Apollo readings behind the reason")
    parser.add_argument("--revision", help="JSON: the decision a version after v1 rests on (who, when, where, what changed)")
    args = parser.parse_args()
    record = make_record(args.smoke, args.lock_test, args.slice, args.failure_path, args.compare, args.slice_return,
                         args.slice_pending, args.reason, json.loads(args.credit_readings) if args.credit_readings else None,
                         json.loads(args.revision) if args.revision else None)
    write_record(record, ra.FREEZE_FILE)
    print(f'FROZEN {ra.VERSION} at {record["frozen_at"]}: {record["model"]} at {record["effort"]}, timeout {record["timeout_minutes"]} min, '
          f'{len(record["lock"]["allowed"])} tools allowed, {len(record["lock"]["denied_by_name"])} denied by name, '
          f'{len(record["lock"]["denied_servers"])} connectors denied')
    for role, value in record["file_sha256"].items():
        print(f"       {role:18} {value[:16]}…")
    print(f'       assembled prompt   {record["assembled_prompt_sha256"][:16]}…')
    print(f"Record: {ra.FREEZE_FILE}")


if __name__ == "__main__":
    main()
