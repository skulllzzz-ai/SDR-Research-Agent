"""Run the research agent: one headless Claude Code run per company, from a new empty folder.

    python run_agent.py preflight
    python run_agent.py lock-test --out data-exports/smoke-slice/<name>
    python run_agent.py research --companies <file.csv> --out data-exports/smoke-slice/<name>
    python run_agent.py batch          # the 30 drawn companies, pair companies first; only after the v1 freeze
    python run_agent.py batch --only C07,C19   # a later version (v2): only the named With-agent companies, after its own freeze
    python run_agent.py self-test      # offline: the lock, the flags, the prompt assembly, the usage-window wait

The companies file has the columns company_id, company_name, market.

How a run is kept clean and locked (agent_config.json holds every setting):
  - launched from a new empty folder outside work/, with every CLAUDE* and ANTHROPIC* variable
    removed from the environment, and with no user or project settings (--setting-sources "");
  - exactly seven tools are allowed: web search, web fetch and five Apollo read tools;
  - the Apollo send, buy and edit tools are denied by name; every other account connector is
    denied as a whole, so the agent cannot even see its tools; the remaining Apollo read tools
    are denied because nothing else is allowed (permission mode dontAsk never prompts);
  - a hook caps the credit-spending Apollo calls per company (credit_cap_hook.py);
  - a per-company timeout; a timeout is logged as a failed row;
  - the answer must pass validate_row.py; one repair pass without tools if it does not.

Tool names are never typed from memory. The connectors finish connecting after the run's first
event, so their tools are not listed there. The preflight therefore has a small harness run search
the tools for "apollo", twice, and this script reads the names out of the search results itself
and compares the two readings. The connector list comes from `claude mcp list`. Tool search is the
harness's own loader: it loads tool definitions, acts on nothing outside, and needs no permission.
Harness runs (discovery, the hidden-connector check) use the small harness model; the model check,
the canary, the research run and the repair pass use the pinned model.

--model, --effort and --web-only are for a mechanics check only; such a run is marked as a probe
in its records and never counts as an agent row.
Outputs go under the --out folder, which must sit inside data-exports/ (row-level, gitignored).
"""

import argparse
import csv
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
from collections import Counter
from datetime import datetime, timedelta, timezone
from pathlib import Path

from validate_row import validate

HERE = Path(__file__).resolve().parent
CONFIG = json.loads((HERE / "agent_config.json").read_text(encoding="utf-8"))
VERSION = CONFIG.get("agent_version", "v1")   # names the freeze record, the batch folder and each run's stamp
RUN = {"model": CONFIG["model"], "effort": CONFIG["effort"], "timeout_minutes": CONFIG["timeout_minutes"]}
APOLLO_NAME = re.compile(r"mcp__[A-Za-z0-9_-]+__apollo_[a-z0-9_]+")
MCP_NAME = re.compile(r"mcp__[A-Za-z0-9_-]+__[A-Za-z0-9_-]+")
CANARY = (
    'Answer with one JSON object and nothing else: {"bullet_point_rule": "yes" or "no", "instruction_files": '
    '"<the names or paths of any CLAUDE.md, memory, rules or lessons files whose content you were given, or none>"}. '
    "The question: apart from this message, do the instructions you were given for this conversation include a rule "
    "that tells you to always reply in bullet points?"
)
DISCOVERY = (
    'Call the ToolSearch tool once with the query "apollo" and max_results set to 100. Do not call any other tool. '
    "Then reply with the single word DONE."
)
OTHERS = (
    'Call the ToolSearch tool three times, with max_results set to 20 each time: first with the query "gmail", then '
    '"slack", then "clickup". Do not call any other tool. Then reply with the single word DONE.'
)


def now():
    return datetime.now().astimezone()


def sha256_text(text):
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def clean_env(extra=None):
    """The environment of every run: no CLAUDE* or ANTHROPIC* variable, and the CLI's
    auto-updater off, so the Claude Code version cannot change in the middle of a batch."""
    env = {k: v for k, v in os.environ.items() if not k.upper().startswith(("CLAUDE", "ANTHROPIC"))}
    env["DISABLE_AUTOUPDATER"] = "1"
    env.update(extra or {})
    return env


def empty_folder():
    folder = Path(tempfile.mkdtemp(prefix="sdr-agent-")).resolve()
    work = next(p for p in HERE.parents if p.name == "work")
    assert work not in folder.parents and not any(folder.iterdir()), f"{folder} must be empty and outside {work}"
    return folder


def remove_folder(folder):
    """Removes the run's folder, and the empty per-project folder Claude Code makes for it under
    ~/.claude/projects (54 of them had piled up by 4 Oct). A project folder holding any file is kept."""
    shutil.rmtree(folder, ignore_errors=True)
    project = Path.home() / ".claude" / "projects" / re.sub(r"[^A-Za-z0-9-]", "-", str(folder))
    if project.is_dir() and not any(p.is_file() for p in project.rglob("*")):
        shutil.rmtree(project, ignore_errors=True)


def command(tools, allowed=(), denied=(), extra=(), harness=False):
    """harness=True runs on the small harness model: plumbing only, never an agent row."""
    cmd = ["claude", "-p", "--model", CONFIG["harness_model"] if harness else RUN["model"]]
    if RUN["effort"] and not harness:
        cmd += ["--effort", RUN["effort"]]
    cmd += [
        "--output-format", "stream-json", "--verbose", "--no-session-persistence",
        "--setting-sources", CONFIG["setting_sources"], "--permission-mode", CONFIG["permission_mode"],
        "--tools", ",".join(tools),
    ]
    if allowed:
        cmd += ["--allowedTools", ",".join(allowed)]
    if denied:
        cmd += ["--disallowedTools", ",".join(denied)]
    return cmd + list(extra)


def launch(cmd, message, env, timeout_s, stream_path=None):
    """Start claude in a new empty folder, send the message on stdin, collect the event stream
    with arrival times. The folder is removed afterwards; anything found in it is reported.
    Returns (events, stderr, status, {"path", "left_behind"})."""
    folder = empty_folder()
    proc = subprocess.Popen(cmd, cwd=folder, env=env, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE, text=True, encoding="utf-8", errors="replace")
    events, errors = [], []

    def read_stream():
        sink = open(stream_path, "w", encoding="utf-8") if stream_path else None
        for line in proc.stdout:
            line = line.strip()
            if not line:
                continue
            stamp = now().isoformat(timespec="seconds")
            try:
                event = json.loads(line)
            except json.JSONDecodeError:
                event = {"type": "unparsed", "text": line}
            events.append((stamp, event))
            if sink:
                sink.write(json.dumps({"t": stamp, "event": event}, ensure_ascii=False) + "\n")
                sink.flush()
        if sink:
            sink.close()

    readers = [threading.Thread(target=read_stream), threading.Thread(target=lambda: errors.append(proc.stderr.read()))]
    for reader in readers:
        reader.start()
    proc.stdin.write(message)
    proc.stdin.close()
    try:
        proc.wait(timeout=timeout_s)
        status = "completed" if proc.returncode == 0 else f"exit {proc.returncode}"
    except subprocess.TimeoutExpired:
        subprocess.run(["taskkill", "/PID", str(proc.pid), "/T", "/F"], capture_output=True)
        status = "timeout"
    for reader in readers:
        reader.join(timeout=30)
    left_behind = sorted(p.name for p in folder.iterdir())
    remove_folder(folder)
    return events, "".join(errors), status, {"path": str(folder), "left_behind": left_behind}


def refresh_clash(events):
    """True when the run died on the CLI's transient OAuth refresh clash (5 Oct 2026: "Failed to
    refresh OAuth token: another Claude Code process is refreshing it or exited mid-refresh"). The
    CLI itself says to retry in a minute, so callers launch once more after 60 seconds."""
    text = " ".join(json.dumps(event) for _, event in events).lower()
    return "refresh oauth token" in text and "another claude code process" in text


def launch_retrying(cmd, message, env, timeout_s, stream_path=None):
    """launch(), repeated once after 60 seconds when the first try hit the refresh clash."""
    result = launch(cmd, message, env, timeout_s, stream_path)
    if refresh_clash(result[0]):
        print("      transient OAuth refresh clash: waiting 60 seconds and launching once more", flush=True)
        time.sleep(60)
        result = launch(cmd, message, env, timeout_s, stream_path)
    return result


def find(events, kind, subtype=None, last=False):
    hits = [e for _, e in events if e.get("type") == kind and (subtype is None or e.get("subtype") == subtype)]
    return (hits[-1] if last else hits[0]) if hits else {}


def flatten(content):
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        # a text part gives its text; any other part (a tool reference, for one) is kept as JSON
        return "\n".join(part["text"] if isinstance(part, dict) and isinstance(part.get("text"), str)
                         else json.dumps(part, ensure_ascii=False) for part in content)
    return "" if content is None else str(content)


def credit_blocks(text):
    """Every "mcp_credits" object found in a tool result, parsed by brace matching."""
    found = []
    for hit in re.finditer(r'"mcp_credits"\s*:\s*\{', text):
        depth, start = 0, hit.end() - 1
        for i in range(start, len(text)):
            depth += {"{": 1, "}": -1}.get(text[i], 0)
            if depth == 0:
                try:
                    found.append(json.loads(text[start:i + 1]))
                except json.JSONDecodeError:
                    found.append({"unparsed": text[start:i + 1][:300]})
                break
    return found


def summarise(events):
    """What loaded, every tool called, what was denied, and how the run ended."""
    init, result = find(events, "system", "init"), find(events, "result", last=True)
    calls, by_id, denied_messages, rate_limit = [], {}, [], None
    for stamp, event in events:
        message = event.get("message")
        if event.get("type") == "rate_limit_event":
            rate_limit = event.get("rate_limit_info")
        # since CLI 2.1.289 a denial is also a system event whose message is plain text
        if event.get("type") == "system" and event.get("subtype") == "permission_denied" and isinstance(message, str):
            denied_messages.append(message[:160])
        content = message.get("content") if isinstance(message, dict) else None
        if not isinstance(content, list):
            continue
        for block in content:
            if not isinstance(block, dict):
                continue
            if event.get("type") == "assistant" and block.get("type") == "tool_use":
                call = {"at": stamp, "tool": block.get("name"), "input": block.get("input"), "is_error": None,
                        "result_chars": None, "credits": [], "result_head": "", "apollo_names": [], "mcp_names": []}
                calls.append(call)
                by_id[block.get("id")] = call
            elif event.get("type") == "user" and block.get("type") == "tool_result" and block.get("tool_use_id") in by_id:
                call, text = by_id[block["tool_use_id"]], flatten(block.get("content"))
                call.update({"is_error": bool(block.get("is_error")), "result_chars": len(text), "credits": credit_blocks(text),
                             "result_head": text[:600], "apollo_names": sorted(set(APOLLO_NAME.findall(text))),
                             "mcp_names": sorted(set(MCP_NAME.findall(text)))})
    denials = result.get("permission_denials") or []
    tools = init.get("tools") or []
    return {
        "loaded": {
            "claude_code_version": init.get("claude_code_version"), "model": init.get("model"),
            "permission_mode": init.get("permissionMode"), "api_key_source": init.get("apiKeySource"),
            "builtin_tools": [t for t in tools if not t.startswith("mcp__")],
            "mcp_tools_offered": [t for t in tools if t.startswith("mcp__")],
            "mcp_servers": init.get("mcp_servers"),
            "plugins": [p.get("name") if isinstance(p, dict) else p for p in init.get("plugins") or []],
            "skills": init.get("skills"), "agents": init.get("agents"),
            "memory_paths": init.get("memory_paths"), "output_style": init.get("output_style"),
        },
        "tool_calls": calls,
        "tools_called": {name: sum(c["tool"] == name for c in calls) for name in sorted({c["tool"] for c in calls})},
        "permission_denials": [{"tool": d.get("tool_name"), "input": d.get("tool_input")} for d in denials],
        "denied_messages": denied_messages,
        "rate_limit": rate_limit,
        "result": {k: result.get(k) for k in ("subtype", "is_error", "api_error_status", "duration_ms", "num_turns", "total_cost_usd")},
        "models_used": sorted((result.get("modelUsage") or {}).keys()),
        "usage": result.get("usage"),
        "answer": result.get("result") if isinstance(result.get("result"), str) else "",
    }


def extract_json(text):
    text = (text or "").strip()
    fenced = re.search(r"```(?:json)?\s*(\{.*\})\s*```", text, re.S)
    for label, candidate in (("clean", text), ("fenced", fenced.group(1) if fenced else None),
                             ("embedded", text[text.find("{"):text.rfind("}") + 1] if "{" in text else None)):
        if candidate:
            try:
                return json.loads(candidate), label
            except json.JSONDecodeError:
                continue
    return None, "none"


def server_rule(name):
    """The permission rule that covers every tool of one MCP server: mcp__<server name, normalised>."""
    return "mcp__" + re.sub(r"[^A-Za-z0-9_-]", "_", name)


def list_servers():
    """The connectors the CLI knows, from `claude mcp list`: the same answer every time, unlike
    the list in a run's first event, which depends on how far the connections have got."""
    folder = empty_folder()
    done = subprocess.run(["claude", "mcp", "list"], cwd=folder, env=clean_env(), capture_output=True, text=True,
                          encoding="utf-8", errors="replace", timeout=300)
    remove_folder(folder)
    servers = []
    for line in done.stdout.splitlines():
        name, found, rest = line.partition(": ")
        if found and " - " in rest:
            servers.append({"name": name.strip(), "status": rest.rsplit(" - ", 1)[-1].strip().lstrip("✔✓!✘ ").strip()})
    if not servers:
        raise SystemExit(f"`claude mcp list` returned no connectors: {(done.stdout + done.stderr).strip()[:300]}")
    return servers


def discover(out_dir, label, attempts=4):
    """A harness run that searches the tools for "apollo". The names are read from the search
    results by this script, not typed out by the model. A run that starts before the connectors
    have connected has no tool search to call, so an empty run is tried again. Returns the names."""
    for attempt in range(1, attempts + 1):
        events, stderr, status, _ = launch(command(["ToolSearch"], harness=True), DISCOVERY, clean_env(), 300,
                                           out_dir / f"{label}-try{attempt}-stream.jsonl" if out_dir else None)
        summary = summarise(events)
        if status != "completed" or summary["result"].get("is_error"):
            raise SystemExit(f"discovery failed ({status}): {summary['answer'] or stderr.strip()[:400]}")
        searches = [c for c in summary["tool_calls"] if c["tool"] == "ToolSearch"]
        names = sorted({name for c in searches for name in c["apollo_names"]})
        called = [c["tool"] for c in summary["tool_calls"] if "apollo" in c["tool"].lower()]
        assert not called, f"the discovery run called an Apollo tool: {called}"
        print(f'DISCOVERY {label}, try {attempt}: {len(names)} Apollo tool names read from {len(searches)} tool search result(s)')
        if names:
            return names
    raise SystemExit(f"discovery found no Apollo tools in {attempts} tries: the Apollo connector may not be connected")


def build_lock(first, second, servers):
    """The allow list (the built-in tools plus the configured Apollo tools: seven to 5 Oct, web search
    and web fetch only from 6 Oct), the deny lists and the default-deny test tool, resolved against
    the names two discovery runs returned and the servers the run reported. When no Apollo tool is
    allowed, Apollo is denied as a whole like every other connector."""
    agreed, either = sorted(set(first) & set(second)), sorted(set(first) | set(second))

    def full(base):
        hits = [t for t in agreed if t.endswith("__" + base)]
        if len(hits) != 1:
            raise SystemExit(f"tool {base!r} resolved to {len(hits)} names among the {len(agreed)} Apollo tools both "
                             f"discovery runs listed. Servers reported: {[s.get('name') for s in servers]}")
        return hits[0]

    apollo_allowed = [full(base) for base in CONFIG["apollo_allowed"]]
    allowed = list(CONFIG["builtin_tools"]) + apollo_allowed
    denied = [t for t in either if t not in allowed
              and any(part in t.rsplit("__", 1)[-1] for part in CONFIG["deny_by_name_if_name_contains"])]
    test_tool = full(CONFIG["default_deny_test_tool"])
    prefixes = {t.rsplit("__", 1)[0] for t in either}
    assert len(prefixes) == 1, f"the Apollo tools come from more than one server: {sorted(prefixes)}"
    apollo_server = prefixes.pop()
    rules = {server_rule(s["name"]): s["name"] for s in servers}
    assert apollo_server in rules, f"the Apollo tools' server {apollo_server} is not among the listed connectors {sorted(rules)}"
    other_servers = sorted(rule for rule in rules if rule != apollo_server or not apollo_allowed)
    expected = len(CONFIG["builtin_tools"]) + len(CONFIG["apollo_allowed"])
    assert len(allowed) == expected and len(set(allowed)) == expected, allowed
    assert not set(allowed) & set(denied), "an allowed tool is also on the deny list"
    assert test_tool not in allowed and test_tool not in denied, "the default-deny test tool must be on neither list"
    return {
        "allowed": allowed, "apollo_allowed": apollo_allowed, "denied_by_name": denied, "denied_servers": other_servers,
        "apollo_server": apollo_server, "apollo_denied_as_a_whole": apollo_server in other_servers,
        "default_deny_test_tool": test_tool,
        "discovery": {"first": len(first), "second": len(second), "in_both": len(agreed), "in_either": len(either),
                      "same_list": sorted(first) == sorted(second)},
        "apollo_read_tools_left_to_default_deny": len(either) - len(apollo_allowed) - len(denied),
        "apollo_tools": either,
    }


def preflight(out_dir=None):
    """Proves the login and the pinned model, shows what loads, and builds the lock."""
    events, stderr, status, _ = launch_retrying(command([""]), "Reply with the single word OK.", clean_env(), 300,
                                                out_dir / "model-check-stream.jsonl" if out_dir else None)
    summary = summarise(events)
    if status != "completed" or summary["result"].get("is_error"):
        raise SystemExit(f"preflight failed ({status}): {summary['answer'] or stderr.strip()[:400]}")
    servers = list_servers()
    # the run's first event can name a connector the CLI's list does not, and the other way round
    # (4 Oct: "claude.ai PostHog" against "plugin:posthog:posthog"); every server either side
    # reports is denied as a whole, except Apollo
    known = {server_rule(s["name"]) for s in servers}
    for s in summary["loaded"]["mcp_servers"] or []:
        if server_rule(s["name"]) not in known:
            servers.append({"name": s["name"], "status": f'{s.get("status", "")} (named by the run only)'})
            known.add(server_rule(s["name"]))
    # every connector ever seen stays denied, so the deny list does not shrink when one is slow to
    # connect (6 Oct: the run-only PostHog name came and went between runs)
    seen_file = MASTER_LOG / "connectors-seen.json"
    seen = json.loads(seen_file.read_text(encoding="utf-8")) if seen_file.exists() else []
    for name in seen:
        if server_rule(name) not in known:
            servers.append({"name": name, "status": "not reported in this run (seen in an earlier one)"})
            known.add(server_rule(name))
    MASTER_LOG.mkdir(parents=True, exist_ok=True)
    seen_file.write_text(json.dumps(sorted({s["name"] for s in servers}), indent=2) + "\n", encoding="utf-8")
    lock = build_lock(discover(out_dir, "discovery-1"), discover(out_dir, "discovery-2"), servers)
    summary["loaded"]["connectors"] = servers
    return summary, lock


def lock_flags(lock, extra=()):
    """Tools and permission flags of a locked run. Tool search is offered only as the loader of an
    allowed connector tool, so a web-only agent is offered exactly web search and web fetch."""
    tools = list(CONFIG["builtin_tools"]) + (["ToolSearch"] if lock["apollo_allowed"] else [])
    return command(tools, lock["allowed"], lock["denied_by_name"] + lock["denied_servers"], extra)


def bare_flags(lock, extra=()):
    """A run with no tools at all: the canary and the repair pass."""
    return command([""], denied=sorted(set(lock["denied_servers"]) | {lock["apollo_server"]}), extra=extra)


def canary(lock, out_dir):
    events, _, status, _ = launch_retrying(bare_flags(lock), CANARY, clean_env(), 180, out_dir / "canary-stream.jsonl")
    summary = summarise(events)
    stop_if_rate_limited(summary)
    answer, _ = extract_json(summary["answer"])
    return {"status": status, "answer": answer, "raw": summary["answer"][:500]}


def tool_lines(lock):
    names = dict(zip(CONFIG["apollo_allowed"], lock["apollo_allowed"]))
    caps = CONFIG["apollo_caps_per_company"]
    labels = {
        "apollo_mixed_companies_search": "Apollo company search",
        "apollo_organizations_lookup": "Apollo organization lookup",
        "apollo_organizations_enrich": "Apollo organization enrich",
        "apollo_mixed_people_api_search": "Apollo people search",
        "apollo_people_match": "Apollo people match",
    }
    lines = ["- Web search: WebSearch", "- Web fetch: WebFetch"]
    for base, label in labels.items():
        if base not in names:
            continue
        cap = f' (at most {caps[base]} call{"s" if caps[base] > 1 else ""} per company)' if base in caps else ""
        lines.append(f"- {label}: {names[base]}{cap}")
    return "\n".join(lines)


def assemble_prompt(lock):
    template = (HERE / CONFIG["prompt_template"]).read_text(encoding="utf-8")
    block = (HERE / CONFIG["competitor_block"]).read_text(encoding="utf-8")
    for placeholder in ("{{COMPETITOR_BLOCK}}", "{{TOOL_NAMES}}"):
        assert template.count(placeholder) == 1, f"the template must hold {placeholder} exactly once"
    assembled = template.replace("{{COMPETITOR_BLOCK}}", block.strip()).replace("{{TOOL_NAMES}}", tool_lines(lock))
    return assembled, {
        "prompt_template": CONFIG["prompt_template"], "prompt_template_sha256": sha256_text(template),
        "competitor_block": Path(CONFIG["competitor_block"]).name, "competitor_block_sha256": sha256_text(block),
        "repair_prompt_sha256": sha256_text((HERE / CONFIG["repair_prompt"]).read_text(encoding="utf-8")),
        "assembled_prompt_sha256": sha256_text(assembled),
    }


def hook_settings():
    hook = f'"{Path(sys.executable).as_posix()}" "{(HERE / "credit_cap_hook.py").as_posix()}"'
    return json.dumps({"hooks": {"PreToolUse": [{"matcher": "mcp__.*", "hooks": [{"type": "command", "command": hook}]}]}})


def stop_if_login_failed(answer):
    """A dead login fails every run that follows, so the batch stops at the first one."""
    text = (answer or "").lower()
    if "failed to authenticate" in text or ("oauth" in text and "expired" in text):
        raise SystemExit(f"STOPPED: the terminal claude is not logged in ({answer.strip()[:120]}). Run `claude auth login`, "
                         "then start the same command again: finished companies are kept and the rest are run.")


IST = timezone(timedelta(hours=5, minutes=30))
MAX_WAIT_HOURS = 6


class UsageLimit(Exception):
    """The account's usage window rejected a run. resets_at is the reset as an epoch second, or None."""

    def __init__(self, resets_at, detail):
        super().__init__(detail)
        self.resets_at = resets_at


def stop_if_rate_limited(summary):
    """Raises UsageLimit when the account's usage window (the Max plan's 5-hour and 7-day windows)
    rejected the run. The batch loop waits for the reset and runs the same company again
    (DECISIONS.md 10)."""
    result, info, text = summary["result"], summary.get("rate_limit") or {}, summary.get("answer") or ""
    low = text.lower()
    status = str(info.get("status") or "allowed").lower()
    rejected = (status == "rejected" or result.get("api_error_status") == 429
                or (result.get("is_error") and ("rate limit" in low or "usage limit" in low or "limit reached" in low)))
    if not rejected:
        return
    resets = info.get("resetsAt")
    stamped = re.search(r"\|(\d{9,11})\b", text)          # "Claude AI usage limit reached|<epoch>"
    if stamped:
        resets = int(stamped.group(1))
    if resets is None:
        spent = [w["resetsAt"] for w in (info.get("unifiedWindows") or {}).values()
                 if isinstance(w, dict) and "resetsAt" in w and w.get("utilization", 0) >= 1]
        resets = min(spent, default=None)
    raise UsageLimit(resets, f"{status}; {text.strip()[:120]}")


def wait_seconds(limit, clock=time.time):
    """Until two minutes after the reset; a quarter of an hour when the reset is unknown."""
    if limit.resets_at is None:
        return 15 * 60
    return max(60, int(limit.resets_at - clock()) + 120)


def run_companies(companies, run_one, out, sleep=time.sleep, clock=time.time, say=print):
    """Runs the companies in order and yields (company, run). At a usage-limit rejection it waits for
    the window to reset and runs the same company again; it stops when the reset is more than
    MAX_WAIT_HOURS away (the 7-day window). Every wait is logged in usage-waits.jsonl."""
    for company in companies:
        say(f'\nRUN  {company["company_id"]}  {company["company_name"]}  ({company["market"]})')
        while True:
            try:
                run = run_one(company)
                break
            except UsageLimit as limit:
                wait = wait_seconds(limit, clock)
                resume = datetime.fromtimestamp(clock() + wait, IST)
                if wait > MAX_WAIT_HOURS * 3600:
                    raise SystemExit(f"STOPPED: the usage window resets at {resume:%a %d %b %H:%M} IST, more than "
                                     f"{MAX_WAIT_HOURS} hours away ({limit}). Start the same command again then: "
                                     "finished companies are kept.")
                with open(Path(out) / "usage-waits.jsonl", "a", encoding="utf-8") as handle:
                    handle.write(json.dumps({"at": datetime.fromtimestamp(clock(), IST).isoformat(timespec="seconds"),
                                             "company_id": company["company_id"], "reason": str(limit),
                                             "resets_at": limit.resets_at, "wait_seconds": wait,
                                             "resume_at": resume.isoformat(timespec="seconds")}) + "\n")
                say(f"     USAGE LIMIT ({limit}). Waiting until {resume:%a %d %b %H:%M} IST, then this company again")
                sleep(wait)
        yield company, run


def research_one(company, lock, assembled, prompt_file, out_dir):
    """One company. Returns None if this batch folder already holds its finished run: a run is
    never overwritten. A run that was cut short is set aside and done again."""
    company_dir = out_dir / company["company_id"]
    if (company_dir / "run.json").exists():
        return None
    if company_dir.exists():
        n = 1
        while company_dir.with_name(f"{company_dir.name}.cut-short-{n}").exists():
            n += 1
        company_dir.rename(company_dir.with_name(f"{company_dir.name}.cut-short-{n}"))
    company_dir.mkdir(parents=True)
    expect = {k: company[k] for k in ("company_id", "company_name", "market")}
    inputs = f'Company ID: {company["company_id"]}\nCompany name: {company["company_name"]}\nMarket: {company["market"]}\n'
    if lock.get("web_only"):
        inputs += "Apollo is not available in this run: use web search and web fetch only.\n"
    appended = CONFIG["prompt_mode"] == "append"
    prompt_flag = ["--append-system-prompt-file", str(prompt_file)] if appended else []
    cmd = lock_flags(lock, ["--settings", hook_settings()] + prompt_flag)
    env = clean_env({"SDR_AGENT_CAPS": json.dumps(CONFIG["apollo_caps_per_company"]),
                     "SDR_AGENT_CAP_FILE": str(company_dir / "credit-cap-counter.json")})

    canary_result = canary(lock, company_dir)
    stop_if_login_failed(canary_result["raw"])
    started = now()
    events, stderr, status, folder = launch_retrying(cmd, inputs if appended else assembled + "\n\n" + inputs, env,
                                                     RUN["timeout_minutes"] * 60, company_dir / "stream.jsonl")
    summary = summarise(events)
    stop_if_login_failed(summary["answer"])
    stop_if_rate_limited(summary)
    (company_dir / "answer.txt").write_text(summary["answer"], encoding="utf-8")
    (company_dir / "stderr.txt").write_text(stderr, encoding="utf-8")
    row, shape = extract_json(summary["answer"])
    problems = validate(row, expect) if row is not None else ["no JSON object in the answer"]
    if status == "timeout":
        problems = [f'timeout after {RUN["timeout_minutes"]} minutes']
    first_problems, retries, repair = list(problems), 0, None

    if problems and status == "completed":
        retries = 1
        repair_text = (HERE / CONFIG["repair_prompt"]).read_text(encoding="utf-8")
        repair_message = (repair_text.replace("{{INPUT}}", inputs).replace("{{PROBLEMS}}", "\n".join(f"- {p}" for p in problems))
                          .replace("{{PREVIOUS}}", summary["answer"]))
        if not appended:
            repair_message = assembled + "\n\n" + repair_message
        r_events, _, _, _ = launch(bare_flags(lock, prompt_flag), repair_message, clean_env(), 600,
                                   company_dir / "repair-stream.jsonl")
        repair = summarise(r_events)
        stop_if_rate_limited(repair)
        row, shape = extract_json(repair["answer"])
        problems = validate(row, expect) if row is not None else ["no JSON object in the repaired answer"]
        (company_dir / "repair-answer.txt").write_text(repair["answer"], encoding="utf-8")
    ended = now()

    if not problems:
        (company_dir / "row.json").write_text(json.dumps(row, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    else:
        (company_dir / "rejected.json").write_text(
            json.dumps({"problems": problems, "row": row}, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

    caps = CONFIG["apollo_caps_per_company"]
    ran = [c for c in summary["tool_calls"] if c["is_error"] is False]
    apollo_calls = dict(Counter(c["tool"].rsplit("__", 1)[-1] for c in ran if "apollo" in (c["tool"] or "").lower()))
    blocked = sum("Credit cap reached" in c["result_head"] for c in summary["tool_calls"])
    run = {
        **expect,
        "prompt_version": VERSION, "model_asked": RUN["model"], "effort": RUN["effort"], "timeout_minutes": RUN["timeout_minutes"],
        "probe_not_an_agent_row": RUN.get("probe", False),
        "prompt_mode": CONFIG["prompt_mode"], "command": cmd,
        "agent_start": started.isoformat(timespec="seconds"), "agent_end": ended.isoformat(timespec="seconds"),
        "agent_minutes": round((ended - started).total_seconds() / 60, 2),
        "status": status, "answer_shape": shape,
        "validator_result": "valid" if not problems else "rejected", "validator_problems": problems,
        "retry_count": retries, "problems_before_retry": first_problems if retries else [],
        "apollo_calls_that_ran": apollo_calls,
        "apollo_cap_respected": all(apollo_calls.get(base, 0) <= cap for base, cap in caps.items()),
        "calls_blocked_by_the_cap": blocked,
        "apollo_credit_blocks": [block for c in summary["tool_calls"] for block in c["credits"]],
        "canary": canary_result,
        "empty_folder": folder,
        "summary": {k: summary[k] for k in ("loaded", "tools_called", "permission_denials", "result", "models_used", "usage",
                                              "rate_limit")},
        "tool_calls": [{k: c[k] for k in ("at", "tool", "input", "is_error", "result_chars")} for c in summary["tool_calls"]],
        "repair": {k: repair[k] for k in ("result", "models_used", "usage")} if repair else None,
    }
    (company_dir / "run.json").write_text(json.dumps(run, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    return run


VISIBILITY = (
    "List, by exact name, every tool you can call right now. If one of them is a tool search tool, call it once with "
    'the query "apollo" and once with the query "gmail". Call no other tool. Then reply with one JSON object and '
    'nothing else: {"tools": ["<name>", ...]}'
)


def lock_test(lock, out):
    """Exit condition 5, for the allow list the config sets (seven tools to 5 Oct; web search and web
    fetch only from 6 Oct). Default deny: a harmless Apollo read tool on neither list is made
    reachable and attempted, and must be denied. Visibility: under the real lock no connector tool
    is found or called. The credit-cap hook must block a capped call before it runs. No send, buy or
    edit tool is ever tried."""
    test_tool = lock["default_deny_test_tool"]
    search_offered = bool(lock["apollo_allowed"])
    # the default-deny probe makes Apollo reachable (not denied as a whole, tool search offered)
    # while the test tool stays off the allow list
    reach_denied = lock["denied_by_name"] + [rule for rule in lock["denied_servers"] if rule != lock["apollo_server"]]

    def reach(extra=()):
        return command(list(CONFIG["builtin_tools"]) + ["ToolSearch"], lock["allowed"], reach_denied, extra)

    message = (f"Use the tool search tool to load the tool named {test_tool}, then call that tool exactly once with no "
               "arguments. Call no other tool. Then reply with one JSON object and nothing else: "
               '{"called": "yes" or "no", "outcome": "<what the tool returned, or why it could not run>"}')
    def attempt(cmd, env, label):
        """Send the test message until the test tool is actually attempted: a run that starts
        before the connectors have connected has nothing to load yet."""
        for n in range(1, 5):
            events, _, status, _ = launch(cmd, message, env, 300, out / f"lock-test-{label}-try{n}-stream.jsonl")
            run = summarise(events)
            if any(c["tool"] == test_tool for c in run["tool_calls"]):
                break
        return run, status

    result, status = attempt(reach(), clean_env(), "deny")
    attempted = [c for c in result["tool_calls"] if c["tool"] == test_tool]
    denied = [d for d in result["permission_denials"] if d["tool"] == test_tool]
    ran = [c for c in attempted if c["is_error"] is False]
    tried = [c["tool"] for c in result["tool_calls"]]

    def other_tools_found(cmd, label):
        """Tool names of other connectors that tool search returns: read from the search results."""
        for n in range(1, 5):
            o_events, _, _, _ = launch(cmd, OTHERS, clean_env(), 300, out / f"lock-test-{label}-try{n}-stream.jsonl")
            run = summarise(o_events)
            searches = [c for c in run["tool_calls"] if c["tool"] == "ToolSearch"]
            if len(searches) >= 3:
                break
        found = sorted({n for c in searches for n in c["mcp_names"] if not n.startswith(lock["apollo_server"] + "__")})
        return found, len(searches), run

    if search_offered:
        seen, searches_made, others = other_tools_found(
            command(["ToolSearch"], denied=lock["denied_by_name"] + lock["denied_servers"], harness=True), "others")
        visibility = [("the other connectors are invisible under the lock: tool search finds none of their tools",
                       searches_made >= 3 and seen == [])]
    else:
        v_events, _, _, _ = launch(command(list(CONFIG["builtin_tools"]), lock["allowed"],
                                           lock["denied_by_name"] + lock["denied_servers"], harness=True),
                                   VISIBILITY, clean_env(), 300, out / "lock-test-visibility-stream.jsonl")
        others = summarise(v_events)
        offered, mcp_offered = others["loaded"]["builtin_tools"], others["loaded"]["mcp_tools_offered"]
        reached = [c["tool"] for c in others["tool_calls"] if (c["tool"] or "").startswith("mcp__") or c["tool"] == "ToolSearch"]
        seen, searches_made = reached, 0
        visibility = [
            ("under the lock the run is offered only the allowed tools: no tool search, no connector tool",
             sorted(offered) == sorted(lock["allowed"]) and not mcp_offered),
            ("under the lock no connector tool and no tool search was called, Apollo included", not reached),
            ("Apollo is denied as a whole under the lock", lock["apollo_server"] in lock["denied_servers"]),
        ]
        print(f"      offered under the lock: {offered} + {len(mcp_offered)} connector tools; called: {reached or 'nothing'}")
    control, control_searches, _ = other_tools_found(command(["ToolSearch"], harness=True), "others-control")

    base = test_tool.rsplit("__", 1)[-1]
    env = clean_env({"SDR_AGENT_CAPS": json.dumps({base: 0}), "SDR_AGENT_CAP_FILE": str(out / "lock-test-cap-counter.json")})
    hooked, _ = attempt(reach(["--settings", hook_settings()]), env, "hook")
    hook_blocked = [c for c in hooked["tool_calls"] if c["tool"] == test_tool and "Credit cap reached" in c["result_head"]]
    hook_ran = [c for c in hooked["tool_calls"] if c["tool"] == test_tool and c["is_error"] is False]

    names = ", ".join(t.rsplit("__", 1)[-1] for t in lock["allowed"])
    checks = [
        (f"the allow list holds exactly the {len(lock['allowed'])} configured tools: {names}",
         lock["allowed"] == list(CONFIG["builtin_tools"]) + lock["apollo_allowed"]
         and len(lock["apollo_allowed"]) == len(CONFIG["apollo_allowed"])),
        ("the deny list names the Apollo send, buy and edit tools", len(lock["denied_by_name"]) > 0),
        ("the test tool is on neither the allow list nor the deny-by-name list",
         test_tool not in lock["allowed"] and test_tool not in lock["denied_by_name"]),
        ("default deny: the test tool, reachable but not allowed, was attempted", bool(attempted)),
        ("the attempt was denied and the tool never ran", bool(denied) and not ran),
        ("no send, buy or edit tool was tried in any test run",
         not [t for t in tried + [c["tool"] for c in others["tool_calls"] + hooked["tool_calls"]] if t in lock["denied_by_name"]]),
        *visibility,
        ("control: without the lock the same searches do find other connectors' tools", len(control) > 0),
        ("the credit cap hook blocks a call over the cap before it runs", bool(hook_blocked) and not hook_ran),
    ]
    for name, passed in checks:
        print(f'{"PASS" if passed else "FAIL"}  {name}')
    print(f'      attempts: {[(c["tool"], "ran" if c["is_error"] is False else "did not run", c["result_head"][:160]) for c in attempted]}')
    print(f'      denials logged by the run: {result["permission_denials"]}')
    print(f"      other connectors' tools found under the lock: {len(seen)} in {searches_made} searches; "
          f"without the lock: {len(control)} in {control_searches} searches")
    print(f'      hook run: {[(c["tool"], c["result_head"][:120]) for c in hooked["tool_calls"] if c["tool"] == test_tool]}')
    report = {"status": status, "model": RUN["model"], "effort": RUN["effort"], "test_tool": test_tool, "attempts": attempted,
              "permission_denials": result["permission_denials"], "answer": result["answer"], "other_connector_tools_seen": seen, "other_connector_tools_seen_without_the_lock": control,
              "hook_run_calls": hooked["tool_calls"], "lock": lock, "checks": {name: passed for name, passed in checks}}
    (out / "lock-test.json").write_text(json.dumps(report, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    return all(passed for _, passed in checks)


MASTER_LOG = HERE / "data-exports" / "master-log"
FREEZE_FILE = MASTER_LOG / f"freeze-{VERSION}.json"
BATCH_FOLDER = f"data-exports/agent-runs/batch-{VERSION}"
# The files that make the agent. All are recorded at the freeze; a change to a blocking one after
# the freeze makes a different agent, so the drawn companies are refused until a new freeze.
# run_agent.py itself is recorded but does not block: each run records the exact command it launched.
NON_BLOCKING = ("runner",)


def agent_files():
    return {
        "prompt_template": HERE / CONFIG["prompt_template"],
        "repair_prompt": HERE / CONFIG["repair_prompt"],
        "competitor_block": HERE / CONFIG["competitor_block"],
        "validator": HERE / "validate_row.py",
        "contract": HERE / "contract.py",
        "credit_cap_hook": HERE / "credit_cap_hook.py",
        "runner": HERE / "run_agent.py",
    }


def file_hashes():
    """SHA-256 of each agent file's text (line endings normalised, so a git checkout that rewrites
    them does not count as a change)."""
    return {role: sha256_text(path.read_text(encoding="utf-8")) for role, path in agent_files().items()}


def drawn_companies():
    """The 30 drawn companies, from the one draw record in the master-log folder (none if it is absent)."""
    records = sorted(MASTER_LOG.glob("draw-*.json"))
    assert len(records) <= 1, f"more than one draw record: {[r.name for r in records]}"
    return json.loads(records[0].read_text(encoding="utf-8"))["companies"] if records else []


def is_drawn(company, drawn):
    names = {c["company_name"].strip().lower() for c in drawn}
    return company["company_id"] in {c["company_id"] for c in drawn} or company["company_name"].strip().lower() in names


def batch_companies(drawn):
    """Spec exit condition 10: pair companies first, the earliest in the work order first of all,
    then the SDR-alone companies, whose agent rows are scored as agent-alone rows."""
    ordered = (sorted((c for c in drawn if c["arm"] == "With agent"), key=lambda c: c["work_order"])
               + sorted((c for c in drawn if c["arm"] == "SDR alone"), key=lambda c: c["work_order"]))
    return [{k: c[k] for k in ("company_id", "company_name", "market")} for c in ordered]


def check_freeze(lock=None, hashes=None, path=None):
    """Stops unless the freeze record of the configured version exists and the agent about to run still matches it.
    Without lock and hashes it checks what can be checked before the preflight (config and files).
    Returns the record and the non-blocking notes."""
    path = path or FREEZE_FILE
    if not path.exists():
        raise SystemExit(f"REFUSED: the drawn companies never run before {VERSION} is frozen (spec). Freeze first with freeze_v1.py; "
                         f"the record goes to {path.name}.")
    frozen = json.loads(path.read_text(encoding="utf-8"))
    problems, notes = [], []
    if RUN.get("probe"):
        problems.append("a mechanics check (a model, effort, timeout or web-only override) never runs the drawn companies")
    if frozen["config"] != CONFIG:
        changed = sorted(k for k in set(frozen["config"]) | set(CONFIG) if frozen["config"].get(k) != CONFIG.get(k))
        problems.append(f"agent_config.json differs from the frozen config in {changed}")
    current = file_hashes()
    for role, value in frozen["file_sha256"].items():
        if current.get(role) != value:
            (notes if role in NON_BLOCKING else problems).append(f"{role} changed since the freeze")
    if lock is not None:
        if lock["allowed"] != frozen["lock"]["allowed"]:
            problems.append("the allow list differs from the frozen one")
        for key in ("denied_by_name", "denied_servers"):
            if sorted(lock[key]) != sorted(frozen["lock"][key]):
                notes.append(f"{key} differs from the freeze (everything not allowed is denied by default, so not blocking)")
    if hashes is not None and hashes["assembled_prompt_sha256"] != frozen["assembled_prompt_sha256"]:
        problems.append("the assembled prompt differs from the frozen one")
    if problems:
        raise SystemExit(f"REFUSED: the agent no longer matches the {VERSION} freeze: " + "; ".join(problems)
                         + ". A changed agent needs its own freeze (v2) before it runs the drawn companies.")
    return frozen, notes


def self_test():
    """Offline, no model call: the lock, the flags and the prompt for the configured allow list, and
    the usage-window wait-and-resume, on the stored discovery list of the latest preflight."""
    results = []

    def check(name, passed, detail=""):
        results.append(bool(passed))
        print(f'{"PASS" if passed else "FAIL"}  {name}{"  [" + str(detail) + "]" if detail != "" else ""}')

    stored = next(json.loads(path.read_text(encoding="utf-8")) for path in
                  sorted(HERE.glob("data-exports/smoke-slice/*/lock.json"), key=lambda q: q.stat().st_mtime, reverse=True)
                  if "apollo_tools" in json.loads(path.read_text(encoding="utf-8"))["lock"])
    names, servers = stored["lock"]["apollo_tools"], stored["loaded"]["connectors"]
    lock = build_lock(names, names, servers)
    web_only = not CONFIG["apollo_allowed"]
    check(f"the allow list is the configured one: {lock['allowed']}",
          lock["allowed"] == list(CONFIG["builtin_tools"]) + lock["apollo_allowed"])
    if web_only:
        check("web-only: Apollo is denied as a whole, with every other connector", lock["apollo_server"] in lock["denied_servers"]
              and len(lock["denied_servers"]) == len({server_rule(s["name"]) for s in servers}))
        check("web-only: the credit tools are denied by name as well",
              all(any(t.endswith(base) for t in lock["denied_by_name"]) for base in ("apollo_people_match", "apollo_organizations_enrich")))
        flags = lock_flags(lock)
        check("web-only: a locked run is offered web search and web fetch only, no tool search",
              flags[flags.index("--tools") + 1] == "WebSearch,WebFetch" and flags[flags.index("--allowedTools") + 1] == "WebSearch,WebFetch")
        check("web-only: the prompt lists the two web tools and nothing else", tool_lines(lock) == "- Web search: WebSearch\n- Web fetch: WebFetch")
        assembled, hashes = assemble_prompt(lock)
        check("web-only: the assembled prompt holds no placeholder and names Apollo only as out of reach",
              "{{" not in assembled and assembled.count("Apollo") == 1 and "cannot reach Apollo" in assembled)
    bare = bare_flags(lock)
    denied = bare[bare.index("--disallowedTools") + 1].split(",")
    check("the canary and repair runs get no tool and deny every connector once", bare[bare.index("--tools") + 1] == ""
          and len(denied) == len(set(denied)) and lock["apollo_server"] in denied)

    window = {"five_hour": {"utilization": 1.0, "resetsAt": 2_000_000_000}}
    def limited(summary):
        try:
            stop_if_rate_limited(summary)
            return None
        except UsageLimit as limit:
            return limit
    ok = {"result": {"is_error": False}, "rate_limit": {"status": "allowed"}, "answer": "{}"}
    check("usage window: an allowed run passes", limited(ok) is None)
    hit = limited({"result": {"is_error": True}, "rate_limit": {"status": "rejected", "resetsAt": 2_000_000_100, "unifiedWindows": window},
                   "answer": "You have hit your limit"})
    check("usage window: a rejected run raises, with the reset it names", hit is not None and hit.resets_at == 2_000_000_100)
    stamped = limited({"result": {"is_error": True}, "rate_limit": None, "answer": "Claude AI usage limit reached|2000000200"})
    check("usage window: the reset is read from the CLI's limit message", stamped is not None and stamped.resets_at == 2_000_000_200)
    check("usage window: an HTTP 429 raises", limited({"result": {"is_error": True, "api_error_status": 429}, "answer": ""}) is not None)
    check("usage window: another API error does not", limited({"result": {"is_error": True, "api_error_status": 500},
                                                              "rate_limit": {"status": "allowed"}, "answer": "API Error: internal"}) is None)

    folder = Path(tempfile.mkdtemp(prefix="runner-selftest-"))
    clock, slept, calls = [1_000_000.0], [], Counter()
    companies = [{"company_id": f"T{i}", "company_name": f"Madeup {i}", "market": "USA"} for i in (1, 2, 3)]

    def run_one(company):
        calls[company["company_id"]] += 1
        if company["company_id"] == "T2" and calls["T2"] == 1:
            raise UsageLimit(clock[0] + 1000, "rejected; made-up limit")
        return {"company_id": company["company_id"]}

    def sleep(seconds):
        slept.append(seconds)
        clock[0] += seconds
    done = [c["company_id"] for c, _ in run_companies(companies, run_one, folder, sleep=sleep, clock=lambda: clock[0], say=lambda t: None)]
    waits = (folder / "usage-waits.jsonl").read_text(encoding="utf-8").splitlines()
    check("wait and resume: the batch waits until two minutes after the reset, then runs the same company again",
          done == ["T1", "T2", "T3"] and calls["T2"] == 2 and slept == [1120] and len(waits) == 1, (done, dict(calls), slept))

    def far(company):
        raise UsageLimit(clock[0] + 8 * 3600, "rejected; the 7-day window")
    try:
        list(run_companies(companies[:1], far, folder, sleep=sleep, clock=lambda: clock[0], say=lambda t: None))
        stopped = False
    except SystemExit:
        stopped = True
    check("wait and resume: a reset more than 6 hours away stops the batch instead", stopped)
    check("wait and resume: an unknown reset is looked at again after 15 minutes", wait_seconds(UsageLimit(None, "x")) == 900)
    shutil.rmtree(folder, ignore_errors=True)
    print(f"\n{sum(results)} of {len(results)} self-test checks passed")
    return all(results)


def out_folder(path):
    out = Path(path).resolve()
    assert "data-exports" in out.parts, "--out must sit inside data-exports/ (row-level output)"
    out.mkdir(parents=True, exist_ok=True)
    return out


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("mode", choices=["preflight", "lock-test", "research", "batch", "self-test"])
    parser.add_argument("--companies")
    parser.add_argument("--only", help="batch: only these drawn companies (comma-separated ids); a version after v1 needs it, "
                                       "With-agent companies only")
    parser.add_argument("--out")
    parser.add_argument("--model", help="mechanics check only: override the pinned model")
    parser.add_argument("--effort", help='mechanics check only: override the pinned effort ("" leaves the flag out)')
    parser.add_argument("--timeout-minutes", type=float,
                        help="mechanics check only: a shorter timeout, to exercise the timeout path (the failure path of spec exit condition 8)")
    parser.add_argument("--web-only", action="store_true",
                        help="mechanics check only: no Apollo tool is allowed, so no credit can be spent")
    args = parser.parse_args()
    if args.mode == "self-test":
        sys.exit(0 if self_test() else 1)
    if args.model is not None:
        RUN["model"] = args.model
    if args.effort is not None:
        RUN["effort"] = args.effort
    if args.timeout_minutes is not None:
        RUN["timeout_minutes"] = args.timeout_minutes
    probe = (RUN["model"] != CONFIG["model"] or RUN["effort"] != CONFIG["effort"] or args.web_only
             or RUN["timeout_minutes"] != CONFIG["timeout_minutes"])
    RUN["probe"] = probe
    if probe:
        print(f'MECHANICS CHECK ONLY: model {RUN["model"]}, effort {RUN["effort"] or "(not set)"}; '
              f'the pinned agent is {CONFIG["model"]} at {CONFIG["effort"]}\n')

    drawn = drawn_companies()
    companies = None
    if args.mode == "batch":
        if probe:
            raise SystemExit("REFUSED: the batch runs the pinned agent only: no --model, --effort, --timeout-minutes or --web-only")
        companies = batch_companies(drawn)
        if args.only:
            wanted = [x.strip() for x in args.only.split(",") if x.strip()]
            unknown = sorted(set(wanted) - {c["company_id"] for c in companies})
            if unknown:
                raise SystemExit(f"REFUSED: not among the drawn companies: {unknown}")
            companies = [c for c in companies if c["company_id"] in wanted]
        if VERSION != "v1":
            arm_of = {c["company_id"]: c["arm"] for c in drawn}
            if not args.only or any(arm_of[c["company_id"]] != "With agent" for c in companies):
                raise SystemExit(f"REFUSED: {VERSION} runs only the With-agent companies named with --only (DECISIONS.md 9 and 14); "
                                 "the agent-alone arm stays v1")
        args.out = args.out or BATCH_FOLDER
    elif args.mode == "research":
        assert args.companies, "research needs --companies <file.csv>"
        with open(args.companies, newline="", encoding="utf-8") as handle:
            companies = list(csv.DictReader(handle))
    guarded = bool(companies) and any(is_drawn(c, drawn) for c in companies)
    if guarded:
        check_freeze()          # fail fast, before any model call; the lock is checked after the preflight
    out = out_folder(args.out) if args.out else None
    summary, lock = preflight(out)
    loaded = summary["loaded"]
    print(f'LOADED  Claude Code {loaded["claude_code_version"]}, model {loaded["model"]}, permission mode {loaded["permission_mode"]}, '
          f'login {loaded["api_key_source"]}; plugins {len(loaded["plugins"])}, skills {len(loaded["skills"] or [])}, '
          f'agents {len(loaded["agents"] or [])}')
    print(f'        connectors: {[s["name"] + " (" + s["status"] + ")" for s in loaded["connectors"]]}')
    print(f'LOCK    allowed {len(lock["allowed"])}: {lock["allowed"]}')
    print(f'        denied by name: {len(lock["denied_by_name"])} Apollo send, buy and edit tools')
    print(f'        denied as a whole: {len(lock["denied_servers"])} other connectors')
    print(f'        left to default deny: {lock["apollo_read_tools_left_to_default_deny"]} other Apollo read tools; '
          f'Apollo denied as a whole: {"yes" if lock["apollo_server"] in lock["denied_servers"] else "no"}')
    print(f'        discovery: {lock["discovery"]}')
    if out:
        if args.web_only:
            lock = {**lock, "allowed": list(CONFIG["builtin_tools"]), "apollo_allowed": [], "web_only": True,
                    "denied_servers": sorted(set(lock["denied_servers"]) | {lock["apollo_server"]})}
            print("        WEB ONLY: every Apollo tool is denied in this run")
        (out / "lock.json").write_text(json.dumps({"probe": probe, "model": RUN["model"], "effort": RUN["effort"], "lock": lock,
                                                   "loaded": loaded}, indent=2) + "\n", encoding="utf-8")
    if args.mode == "preflight":
        return
    if args.mode == "lock-test":
        sys.exit(0 if lock_test(lock, out) else 1)

    assembled, hashes = assemble_prompt(lock)
    prompt_file = out / "assembled-prompt.md"
    if prompt_file.exists() and prompt_file.read_text(encoding="utf-8") != assembled:
        raise SystemExit(f"{prompt_file.name} in this folder differs from the prompt assembled now: the template, the competitor "
                         "list or the tool names changed since this batch started. Use a new --out folder.")
    prompt_file.write_text(assembled, encoding="utf-8")
    freeze = None
    if guarded:
        frozen, notes = check_freeze(lock, hashes)
        freeze = {"file": FREEZE_FILE.name, "sha256": sha256_text(FREEZE_FILE.read_text(encoding="utf-8")),
                  "frozen_at": frozen["frozen_at"], "notes": notes}
        print(f'FROZEN  {VERSION} of {frozen["frozen_at"]}: the agent matches it{" (notes: " + "; ".join(notes) + ")" if notes else ""}')
    if args.mode == "batch":
        with open(out / "companies.csv", "w", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, fieldnames=["company_id", "company_name", "market"])
            writer.writeheader()
            writer.writerows(companies)
        print(f"BATCH   {len(companies)} companies, pair companies first, in work order: {out / 'companies.csv'}")
    started = now()
    print(f"START   {started:%a %d %b %H:%M:%S} IST", flush=True)
    for company, run in run_companies(companies, lambda c: research_one(c, lock, assembled, prompt_file, out), out,
                                      say=lambda text: print(text, flush=True)):
        if run is None:
            print("     already run in this batch folder: kept, not run again")
            continue
        print(f'     {run["status"]}, {run["agent_minutes"]} min, validator: {run["validator_result"]} '
              f'(retries {run["retry_count"]}), tools: {run["summary"]["tools_called"]}')
        print(f'     Apollo calls that ran: {run["apollo_calls_that_ran"]}, cap respected: {run["apollo_cap_respected"]}, '
              f'blocked by the cap: {run["calls_blocked_by_the_cap"]}, credit blocks: {run["apollo_credit_blocks"]}')
        print(f'     canary: {run["canary"]["answer"]}, denials: {run["summary"]["permission_denials"]}, '
              f'left in the folder: {run["empty_folder"]["left_behind"]}')
        if run["validator_problems"]:
            print(f'     problems: {run["validator_problems"]}')
        print(f'     done {now():%H:%M:%S} IST', flush=True)
    # the batch record is rebuilt from every finished run in the folder, so a resumed batch is whole
    runs = [json.loads(path.read_text(encoding="utf-8")) for path in sorted(out.glob("*/run.json"))]
    keep = ("company_id", "company_name", "market", "agent_start", "agent_end", "agent_minutes", "status", "validator_result",
            "retry_count", "apollo_calls_that_ran", "apollo_cap_respected", "calls_blocked_by_the_cap", "apollo_credit_blocks", "canary")
    batch = {"built_at": now().isoformat(timespec="seconds"), "probe_not_agent_rows": probe, "model": RUN["model"],
             "effort": RUN["effort"], "timeout_minutes": RUN["timeout_minutes"], "config": CONFIG, "hashes": hashes,
             "file_sha256": file_hashes(), "freeze": freeze, "order": [c["company_id"] for c in companies], "lock": lock,
             "runs": [{k: r[k] for k in keep} for r in runs]}
    (out / "batch.json").write_text(json.dumps(batch, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"\n{sum(r['validator_result'] == 'valid' for r in runs)} of {len(runs)} rows valid. Batch record: {out / 'batch.json'}")


if __name__ == "__main__":
    main()
