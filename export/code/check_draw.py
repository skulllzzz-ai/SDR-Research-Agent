"""Audit a saved draw record against spec exit condition 1. Reads the FILE, not the draw code.

    python check_draw.py data-exports/master-log/draw-2026-10-01.json
    python check_draw.py data-exports/master-log/draw-2026-10-01.json --self-test

--self-test plants five errors in copies of the record (in memory) and confirms each one is
caught: a check that cannot fail proves nothing.
"""

import copy
import csv
import hashlib
import json
import math
import sys
from collections import Counter
from pathlib import Path

SDR, AGENT = "SDR alone", "With agent"


def largest_remainder(total, sizes):
    pool = sum(sizes.values())
    exact = {k: total * v / pool for k, v in sizes.items()}
    places = {k: math.floor(exact[k]) for k in sizes}
    for k in sorted(sizes, key=lambda k: (-(exact[k] - places[k]), -sizes[k], k))[: total - sum(places.values())]:
        places[k] += 1
    return places


def longest_run(labels, label):
    best = run = 0
    for item in labels:
        run = run + 1 if item == label else 0
        best = max(best, run)
    return best


def audit(record, list_rows):
    """Returns a list of (check, passed, detail)."""
    companies, not_drawn = record["companies"], record["not_drawn"]
    results = []

    def check(name, passed, detail=""):
        results.append((name, bool(passed), detail))

    n = len(companies)
    check("draw size is 30", n == record["draw_size"] == 30, f"{n} companies")
    ids = [c["company_id"] for c in sorted(companies, key=lambda c: c["list_no"])]
    check("company IDs run C01..C30 in list order", ids == [f"C{i:02d}" for i in range(1, n + 1)])

    by_name = {r["company_name"]: r for r in list_rows}
    drawn_names = [c["company_name"] for c in companies]
    left_names = [r["company_name"] for r in not_drawn]
    check("drawn + not drawn = the whole list, no overlap",
          sorted(drawn_names + left_names) == sorted(by_name) and not set(drawn_names) & set(left_names),
          f"{len(drawn_names)} + {len(left_names)} of {len(by_name)}")
    check("market, difficulty and CRM match the list on every company",
          all(c["company_name"] in by_name and (c["market"], c["difficulty"], c["crm"]) ==
              (by_name[c["company_name"]]["market"], by_name[c["company_name"]]["difficulty"], by_name[c["company_name"]]["crm"])
              for c in companies + not_drawn))

    pool_market = Counter(r["market"] for r in list_rows)
    want_market = largest_remainder(n, pool_market)
    got_market = Counter(c["market"] for c in companies)
    check("markets drawn in pool proportions", dict(got_market) == want_market, f"{dict(got_market)}")
    ok = True
    for market in pool_market:
        pool_diff = Counter(r["difficulty"] for r in list_rows if r["market"] == market)
        want = {k: v for k, v in largest_remainder(want_market[market], pool_diff).items() if v}
        got = dict(Counter(c["difficulty"] for c in companies if c["market"] == market))
        ok = ok and got == want
    check("difficulty drawn in pool proportions within each market", ok,
          f"{dict(Counter(c['difficulty'] for c in companies))}")

    arms = Counter(c["arm"] for c in companies)
    check("arms are exactly 15 and 15", arms[SDR] == 15 and arms[AGENT] == 15, f"{dict(arms)}")
    for attr in ("market", "difficulty"):
        worst = 0
        for label in {c[attr] for c in companies}:
            split = Counter(c["arm"] for c in companies if c[attr] == label)
            worst = max(worst, abs(split[SDR] - split[AGENT]))
        check(f"arms differ by 1 at most within every {attr}", worst <= 1, f"largest gap {worst}")

    sdr = [c for c in companies if c["arm"] == SDR]
    early = [c for c in companies if c["early_or_late"] == "Early"]
    check("the early block is 7 companies, all SDR alone",
          len(early) == 7 and all(c["arm"] == SDR for c in early), f"{len(early)} early")
    check("every With-agent company is Late",
          all(c["early_or_late"] == "Late" for c in companies if c["arm"] == AGENT))
    for attr in ("market", "difficulty"):
        ok, detail = True, {}
        total = Counter(c[attr] for c in sdr)
        got = Counter(c[attr] for c in early)
        for label, count in total.items():
            share = 7 * count / len(sdr)
            detail[label] = f"{got[label]} of {count}"
            ok = ok and math.floor(share) <= got[label] <= math.ceil(share)
        check(f"the early block is spread across {attr}", ok, str(detail))

    ordered = sorted(companies, key=lambda c: c["work_order"])
    check("work order runs 1..30 with no gap or repeat", [c["work_order"] for c in ordered] == list(range(1, n + 1)))
    check("orders 1 to 7 are the early block", all(c["early_or_late"] == "Early" for c in ordered[:7]))
    late = ordered[7:]
    ok, detail = True, {}
    for market in pool_market:
        seq = [c["arm"] for c in late if c["market"] == market]
        counts = Counter(seq)
        major, minor = max(counts.values()), min(counts[SDR], counts[AGENT])
        major_arm = SDR if counts[SDR] > counts[AGENT] else AGENT
        minor_arm = AGENT if major_arm == SDR else SDR
        allowed = math.ceil(major / (minor + 1))
        runs = (longest_run(seq, major_arm), longest_run(seq, minor_arm))
        detail[market] = f"longest runs {runs}, allowed ({allowed}, 1)"
        ok = ok and runs[0] <= allowed and runs[1] <= 1
    check("late block: arms alternate within market as evenly as the counts allow", ok, str(detail))
    markets = [c["market"] for c in late]
    counts = Counter(markets)
    major_market = max(counts, key=counts.get)
    allowed = math.ceil(counts[major_market] / (min(counts.values()) + 1))
    check("late block: markets are mixed, not blocked",
          longest_run(markets, major_market) <= allowed, f"longest {major_market} run {longest_run(markets, major_market)}")

    reserve_names = [r["company_name"] for r in record["pasted_reserves"]]
    check("pasted reserves do not overlap the list", not set(reserve_names) & set(by_name), f"{len(reserve_names)} reserves")
    fronts = Counter(r["market"] for r in not_drawn)
    check("not-drawn companies are numbered at the front of their market's reserve queue",
          all(sorted(r["reserve_front_no"] for r in not_drawn if r["market"] == m) == list(range(1, k + 1))
              for m, k in fronts.items()), f"{dict(fronts)}")
    check("a seed is recorded", isinstance(record.get("seed"), int), str(record.get("seed")))
    return results


def planted_errors(record):
    """Five tampered copies. Each must fail at least one check."""
    cases = {}
    r = copy.deepcopy(record)
    r["companies"][0]["arm"] = AGENT if r["companies"][0]["arm"] == SDR else SDR
    cases["one arm flipped"] = r
    r = copy.deepcopy(record)
    victim = next(c for c in r["companies"] if c["early_or_late"] == "Early")
    victim["early_or_late"] = "Late"
    cases["one early company moved to late"] = r
    r = copy.deepcopy(record)
    r["companies"][1]["work_order"] = r["companies"][0]["work_order"]
    cases["a work-order number repeated"] = r
    r = copy.deepcopy(record)
    r["companies"][2]["difficulty"] = "easy" if r["companies"][2]["difficulty"] != "easy" else "hard"
    cases["a difficulty label changed"] = r
    r = copy.deepcopy(record)
    r["companies"][3]["company_name"] = r["not_drawn"][0]["company_name"]
    cases["a not-drawn company swapped in"] = r
    return cases


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    record_path = Path(sys.argv[1])
    record = json.loads(record_path.read_text(encoding="utf-8"))
    list_path = record_path.parent / record["list_file"]
    with open(list_path, newline="", encoding="utf-8") as handle:
        list_rows = list(csv.DictReader(handle))
    list_hash_ok = hashlib.sha256(list_path.read_bytes()).hexdigest() == record["list_sha256"]

    results = audit(record, list_rows) + [("the list file is the one that was drawn from (hash)", list_hash_ok, "")]
    for name, passed, detail in results:
        print(f'{"PASS" if passed else "FAIL"}  {name}{"  [" + detail + "]" if detail else ""}')
    failed = [name for name, passed, _ in results if not passed]
    print(f"\n{len(results) - len(failed)} of {len(results)} checks passed")

    if "--self-test" in sys.argv:
        print("\nSELF-TEST: planted errors must be caught")
        missed = 0
        for label, tampered in planted_errors(record).items():
            caught = [name for name, passed, _ in audit(tampered, list_rows) if not passed]
            missed += not caught
            print(f'{"CAUGHT" if caught else "MISSED"}  {label}: {len(caught)} check(s) failed')
        failed += ["self-test"] * missed
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
