"""PreToolUse hook: a hard cap on credit-spending Apollo calls for one company's run.

The runner passes two environment variables to the agent's process:
  SDR_AGENT_CAPS      JSON, tool base name -> calls allowed, e.g. {"apollo_people_match": 3}
  SDR_AGENT_CAP_FILE  a counter file, one per company run
Claude Code sends the pending tool call as JSON on stdin. Exit code 2 blocks the call and shows
the message to the agent; exit code 0 lets the permission rules decide as usual.
"""

import json
import os
import sys


def main():
    call = json.load(sys.stdin)
    caps = json.loads(os.environ.get("SDR_AGENT_CAPS", "{}"))
    counter_file = os.environ.get("SDR_AGENT_CAP_FILE")
    base = call.get("tool_name", "").rsplit("__", 1)[-1]
    if base not in caps or not counter_file:
        return 0
    counts = {}
    if os.path.exists(counter_file):
        with open(counter_file, encoding="utf-8") as handle:
            counts = json.load(handle)
    counts[base] = counts.get(base, 0) + 1
    with open(counter_file, "w", encoding="utf-8") as handle:
        json.dump(counts, handle)
    if counts[base] > caps[base]:
        print(f"Credit cap reached: {base} is limited to {caps[base]} call(s) per company. "
              "Do not retry it. Finish the row with what you have.", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
