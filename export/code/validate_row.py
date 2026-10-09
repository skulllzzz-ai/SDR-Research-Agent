"""Validate one agent row (JSON) against the row contract. Exit 0 means valid.

    python validate_row.py <row.json> [--company-id C07 --company-name "Name" --market USA]
    python validate_row.py --self-test

The rules mirror prompt/research-agent-v1.md. --self-test is the negative control (spec exit
condition 4): a valid row must pass, and planted bad rows must each be rejected for the planted
reason. The first three are the spec's: a missing field, an invalid marker, a flag without its
source. The test rows are made up; no company data lives in this file.
"""

import argparse
import copy
import json
import re
import sys
from pathlib import Path

from contract import (
    AGENT_DEAD_END_TYPES, CRM_OUT_OF_SCOPE, DEAD_END_WRONG_TARGET, FIELD_KEYS, MARKERS, MARKETS,
    NOT_FOUND, PRIORITIES, VERDICT_DEAD_END, VERDICT_PROSPECT, YN_KEYS,
)

TOP_KEYS = ["company_id", "company_name", "market", "flag", "fields", "verdict", "crm_status"]
FLAG_KEYS = ["likely_dead_end", "reason", "source"]
ENTRY_KEYS = ["value", "marker", "source"]
VERDICT_KEYS = ["value", "dead_end_type", "reason", "source"]
A_SOURCE = re.compile(r"https?://\S+|Apollo: \S+")
# A whole address in plain characters: a masked one (k***@firm.com) is not an email.
AN_EMAIL = re.compile(r"^[A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")


def is_text(value):
    return isinstance(value, str)


def lines_of(value):
    return [line.strip() for line in value.split("\n") if line.strip()]


def validate(row, expect=None):
    """Returns the list of problems. Empty list = the row is valid."""
    problems = []
    if not isinstance(row, dict):
        return ["the row is not a JSON object"]
    for key in TOP_KEYS:
        if key not in row:
            problems.append(f"missing key: {key}")
    for key in row:
        if key not in TOP_KEYS:
            problems.append(f"unknown key: {key}")
    if problems:
        return problems

    for key in ("company_id", "company_name"):
        if not is_text(row[key]) or not row[key].strip():
            problems.append(f"{key} is empty")
    if row["market"] not in MARKETS:
        problems.append(f'market must be one of {MARKETS}, got {row["market"]!r}')
    if expect:
        for key, wanted in expect.items():
            if row.get(key) != wanted:
                problems.append(f"{key} does not match the input: {row.get(key)!r} instead of {wanted!r}")
    if row["crm_status"] != CRM_OUT_OF_SCOPE:
        problems.append(f"crm_status must be exactly {CRM_OUT_OF_SCOPE!r}")

    flag = row["flag"]
    flag_yes = None
    if not isinstance(flag, dict) or sorted(flag) != sorted(FLAG_KEYS) or not all(is_text(flag[k]) for k in FLAG_KEYS):
        problems.append(f"flag must hold exactly the text keys {FLAG_KEYS}")
    elif flag["likely_dead_end"] not in ("yes", "no"):
        problems.append(f'flag.likely_dead_end must be "yes" or "no", got {flag["likely_dead_end"]!r}')
    else:
        flag_yes = flag["likely_dead_end"] == "yes"
        if flag_yes and not flag["reason"].strip():
            problems.append("flag is yes without a reason")
        if flag_yes and not A_SOURCE.search(flag["source"]):
            problems.append("flag is yes without its source (a link or an Apollo reference)")

    fields = row["fields"]
    fields_ok = isinstance(fields, dict)
    if not fields_ok:
        problems.append("fields is not an object")
    else:
        for key in FIELD_KEYS.values():
            if key not in fields:
                problems.append(f"missing field: {key}")
                fields_ok = False
        for key in fields:
            if key not in FIELD_KEYS.values():
                problems.append(f"unknown field: {key}")
        for key in FIELD_KEYS.values():
            entry = fields.get(key)
            if entry is None:
                continue
            if not isinstance(entry, dict) or sorted(entry) != sorted(ENTRY_KEYS) or not all(is_text(entry[k]) for k in ENTRY_KEYS):
                problems.append(f"{key}: must hold exactly the text keys {ENTRY_KEYS}")
                fields_ok = False
                continue
            value, marker, source = entry["value"].strip(), entry["marker"], entry["source"]
            if marker not in MARKERS:
                problems.append(f"{key}: invalid marker {marker!r} (allowed: {MARKERS})")
                fields_ok = False
                continue
            if marker == NOT_FOUND and value:
                problems.append(f'{key}: marked not found but carries a value; the value must be ""')
            if marker != NOT_FOUND and not value:
                problems.append(f"{key}: marked {marker} but the value is empty")
            if marker == "sure" and not A_SOURCE.search(source):
                problems.append(f"{key}: marked sure without a source (a link or an Apollo reference)")
            if key in YN_KEYS and marker != NOT_FOUND and value not in ("Y", "N"):
                problems.append(f'{key}: must be "Y" or "N", got {value!r}')
            if key == "priority" and marker != NOT_FOUND and value not in PRIORITIES:
                problems.append(f"priority: must be one of {PRIORITIES}, got {value!r}")

    if fields_ok:
        names = lines_of(fields["contact_names"]["value"])
        for key, looks_right in (("emails", lambda t: bool(AN_EMAIL.match(t))),
                                 ("phones", lambda t: sum(ch.isdigit() for ch in t) >= 7)):
            entry = fields[key]
            if entry["marker"] == NOT_FOUND:
                continue
            given = lines_of(entry["value"])
            # v2 (DECISIONS.md 14): with no contact named, phones may hold the company's main number on one line
            main_number_only = key == "phones" and not names and len(given) == 1
            if len(given) != len(names) and not main_number_only:
                problems.append(f"{key}: {len(given)} lines for {len(names)} contact names; one line per name, same order"
                                + ("; with no contact named, one line for the company's main number" if key == "phones" and not names else ""))
            for text in given:
                if text != NOT_FOUND and not looks_right(text):
                    problems.append(f'{key}: {text!r} is neither a valid entry nor "{NOT_FOUND}"')

    verdict = row["verdict"]
    if not isinstance(verdict, dict) or sorted(verdict) != sorted(VERDICT_KEYS) or not all(is_text(verdict[k]) for k in VERDICT_KEYS):
        problems.append(f"verdict must hold exactly the text keys {VERDICT_KEYS}")
    elif verdict["value"] not in (VERDICT_PROSPECT, VERDICT_DEAD_END):
        problems.append(f'verdict.value must be "{VERDICT_PROSPECT}" or "{VERDICT_DEAD_END}", got {verdict["value"]!r}')
    elif verdict["value"] == VERDICT_DEAD_END:
        if verdict["dead_end_type"] not in AGENT_DEAD_END_TYPES:
            problems.append(f'verdict.dead_end_type must be one of {AGENT_DEAD_END_TYPES}, got {verdict["dead_end_type"]!r}')
        if verdict["dead_end_type"] == DEAD_END_WRONG_TARGET and row["market"] != "Gulf":
            problems.append("lab-grown only (Wrong target account) is a dead end in the Gulf only, not in this market")
        if not verdict["reason"].strip() or not A_SOURCE.search(verdict["source"]):
            problems.append("a Dead end verdict needs its reason and its source")
        if flag_yes is False:
            problems.append("verdict is Dead end but the flag is no: a dead end needs the flag's positive evidence")
    else:
        if verdict["dead_end_type"] or verdict["reason"].strip() or verdict["source"].strip():
            problems.append('a Prospect verdict carries no dead_end_type, reason or source: all three must be ""')
        if flag_yes:
            problems.append("flag is yes but the verdict is Prospect")
        if fields_ok and fields["priority"]["marker"] == NOT_FOUND:
            problems.append("a Prospect verdict needs a priority")
    return problems


def example_row():
    """A made-up valid row, used by the self-test only."""
    def entry(value, marker="sure", source="https://example.com/about"):
        return {"value": value, "marker": marker, "source": source if marker == "sure" else ""}

    fields = {key: entry("", NOT_FOUND) for key in FIELD_KEYS.values()}
    fields.update({
        "parent_company": entry("Example Holdings"),
        "turnover_revenue_market_share": entry("USD 12m (2025)", "guessed"),
        "detailed_research": entry("- Founded 1998, example city\n- Retail chain, 6 stores"),
        "gold": entry("Y"), "natural_diamond": entry("Y"), "lgd": entry("N", "guessed"),
        "gemstone": entry("N", "guessed"), "pearl": entry("N", "guessed"), "watch": entry("N", "guessed"),
        "jewellery": entry("Y"),
        "product_segments": entry("Bridal, fashion"), "type_of_customer": entry("Retailer"),
        "main_country": entry("USA"), "primary_state": entry("Example State"), "primary_city": entry("Example City"),
        "priority": entry("A", "guessed"),
        "contact_names": entry("A. Example (Owner)\nB. Example (Buyer)", "sure", "Apollo: people search"),
        "emails": entry("a@example.com\nnot found", "sure", "Apollo: people match"),
        "phones": entry("+1 000 000 0000\n+1 000 000 0001", "guessed"),
    })
    return {
        "company_id": "EX00", "company_name": "Example Jewels LLC", "market": "USA",
        "flag": {"likely_dead_end": "no", "reason": "No competitor ownership found; sells natural diamonds", "source": ""},
        "fields": fields,
        "verdict": {"value": VERDICT_PROSPECT, "dead_end_type": "", "reason": "", "source": ""},
        "crm_status": CRM_OUT_OF_SCOPE,
    }


def example_main_number_only():
    """No contact named; the company's main number on one line (allowed since v2, DECISIONS.md 14)."""
    row = example_row()
    for key in ("contact_names", "emails"):
        row["fields"][key] = {"value": "", "marker": NOT_FOUND, "source": ""}
    row["fields"]["phones"] = {"value": "+1 000 000 0000", "marker": "sure", "source": "https://example.com/contact"}
    return row


def example_dead_end():
    row = example_row()
    row["market"] = "Gulf"
    row["flag"] = {"likely_dead_end": "yes", "reason": "Lab-grown only: the company states it sells only lab-grown diamonds",
                   "source": "https://example.com/about"}
    row["verdict"] = {"value": VERDICT_DEAD_END, "dead_end_type": DEAD_END_WRONG_TARGET,
                      "reason": "States it sells only lab-grown diamonds", "source": "https://example.com/about"}
    row["fields"]["priority"] = {"value": "", "marker": NOT_FOUND, "source": ""}
    return row


def planted():
    """(label, bad row, text the rejection must contain). The first three are the spec's."""
    cases = []
    row = example_row()
    del row["fields"]["phones"]
    cases.append(("a missing field", row, "missing field: phones"))
    row = example_row()
    row["fields"]["primary_city"]["marker"] = "confident"
    cases.append(("an invalid marker", row, "invalid marker"))
    row = example_dead_end()
    row["flag"]["source"] = ""
    cases.append(("a flag without its source", row, "flag is yes without its source"))
    row = example_row()
    row["fields"]["parent_company"]["source"] = ""
    cases.append(("a sure field without a source", row, "marked sure without a source"))
    row = example_row()
    row["fields"]["main_country"]["marker"] = NOT_FOUND
    cases.append(("a not-found field that still carries a value", row, "marked not found but carries a value"))
    row = example_dead_end()
    row["market"] = "USA"
    cases.append(("lab-grown only called a dead end in the USA", row, "Gulf only"))
    row = example_row()
    row["crm_status"] = "New"
    cases.append(("a CRM status the agent has no way to know", row, "crm_status must be exactly"))
    row = example_row()
    row["verdict"]["value"] = VERDICT_DEAD_END
    row["verdict"].update({"dead_end_type": "Competitor", "reason": "Looks like a competitor", "source": "https://example.com"})
    cases.append(("a dead-end verdict with the flag left at no", row, "flag is no"))
    row = example_row()
    row["fields"]["priority"] = {"value": "", "marker": NOT_FOUND, "source": ""}
    cases.append(("a prospect without a priority", row, "needs a priority"))
    row = example_row()
    row["fields"]["emails"]["value"] = "a@example.com"
    cases.append(("emails not lined up with the contact names", row, "one line per name"))
    row = example_dead_end()
    row["verdict"]["dead_end_type"] = "Existing active account"
    cases.append(("a dead-end type only the CRM check can give", row, "dead_end_type must be one of"))
    row = example_row()
    row["fields"]["emails"]["value"] = "a***@example.com\nnot found"
    cases.append(("a masked email passed off as an email", row, "is neither a valid entry"))
    row = example_row()
    row["verdict"]["reason"] = "Looks like a good buyer"
    cases.append(("a prospect verdict carrying a dead-end reason", row, "carries no dead_end_type, reason or source"))
    row = example_main_number_only()
    row["fields"]["phones"]["value"] = "+1 000 000 0000\n+1 000 000 0001"
    cases.append(("two phone lines with no contact named", row, "one line per name"))
    return cases


def self_test():
    failures = 0
    for label, row in (("a valid prospect row", example_row()), ("a valid Gulf dead-end row", example_dead_end()),
                       ("a valid row with no contact and the company's main number", example_main_number_only())):
        problems = validate(row)
        failures += bool(problems)
        print(f'{"PASS" if not problems else "FAIL"}  {label} is accepted{"" if not problems else "  " + str(problems)}')
    for label, row, needle in planted():
        problems = validate(copy.deepcopy(row))
        caught = any(needle in problem for problem in problems)
        failures += not caught
        print(f'{"PASS" if caught else "FAIL"}  planted: {label} -> rejected: {problems[:2]}')
    print(f"\n{'ALL PASS' if not failures else str(failures) + ' FAILED'}: 3 valid rows and {len(planted())} planted bad rows")
    return failures


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("row", nargs="?")
    parser.add_argument("--company-id")
    parser.add_argument("--company-name")
    parser.add_argument("--market")
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()
    if args.self_test:
        sys.exit(1 if self_test() else 0)
    expect = {k: v for k, v in (("company_id", args.company_id), ("company_name", args.company_name), ("market", args.market)) if v}
    problems = validate(json.loads(Path(args.row).read_text(encoding="utf-8")), expect or None)
    print("VALID" if not problems else "REJECTED")
    for problem in problems:
        print(f"  {problem}")
    sys.exit(1 if problems else 0)


if __name__ == "__main__":
    main()
