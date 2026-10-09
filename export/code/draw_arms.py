"""Draw the capstone companies, arms, early block and work order from the labelled list.

One seeded random generator, used in a fixed order:
  1. pick DRAW_SIZE companies in pool proportions (by market, then by difficulty within market);
     the ones not drawn go to the front of their market's reserve queue,
  2. assign arms: exactly half each, per market and per difficulty differing by 1 at most,
  3. pick the early SDR-alone block, spread across markets and difficulty,
  4. fix the work order: the early block first, then the rest with arms alternating within market.
Prints the balance table and the order. --write saves the draw record (JSON) and never overwrites
one. --verify recomputes from a record's own seed and list and compares the two.

    python draw_arms.py --list data-exports/list/companies-2026-10-01.csv \
        --reserves data-exports/list/reserves-2026-10-01.csv \
        --paste data-exports/list/pasted-list-2026-10-01.md --seed 20261001 --write
    python draw_arms.py --verify data-exports/master-log/draw-2026-10-01.json

Check the saved record with check_draw.py (an independent audit of the file, not of this code).
"""

import argparse
import csv
import hashlib
import json
import math
import os
import platform
import random
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

from contract import ARM_AGENT, ARM_SDR, DIFFICULTIES, EARLY, LATE, MARKETS

HERE = Path(__file__).resolve().parent
DRAW_SIZE = 30
EARLY_SDR = 7
MAX_TRIES = 100_000


def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def load_list(path):
    with open(path, newline="", encoding="utf-8") as handle:
        rows = list(csv.DictReader(handle))
    names = [r["company_name"].strip().lower() for r in rows]
    assert len(set(names)) == len(names), "duplicate company name in the list"
    for r in rows:
        r["list_no"] = int(r["list_no"])
        assert r["market"] in MARKETS, f"unknown market: {r}"
        assert r["difficulty"] in DIFFICULTIES, f"unknown difficulty: {r}"
    assert [r["list_no"] for r in rows] == list(range(1, len(rows) + 1)), "list_no must run 1..n"
    return rows


def load_reserves(path):
    with open(path, newline="", encoding="utf-8") as handle:
        rows = list(csv.DictReader(handle))
    for r in rows:
        r["reserve_no"] = int(r["reserve_no"])
        assert r["market"] in MARKETS, f"unknown market: {r}"
    return rows


def check_against_paste(rows, reserves, paste_path):
    """Every structured row must be a verbatim line of the paste, and nothing in the paste is missed."""
    lines = [line.strip() for line in Path(paste_path).read_text(encoding="utf-8").splitlines()]
    table = [l for l in lines if l.count(" | ") == 3 and l.split(" | ")[1] in MARKETS]
    ours = [f'{r["company_name"]} | {r["market"]} | {r["difficulty"]} | {r["crm"]}' for r in rows]
    assert table == ours, "structured list differs from the pasted table"
    sentence = next(l for l in lines if l.startswith("Reserves, in order"))
    usa_part, gulf_part = sentence.split(" USA: ", 1)[1].split(" Gulf: ", 1)
    pasted = {
        "USA": [n.strip() for n in usa_part.rstrip(". ").split(";")],
        "Gulf": [n.strip() for n in gulf_part.rstrip(". ").split(";")],
    }
    for market in MARKETS:
        mine = [r["company_name"] for r in sorted(reserves, key=lambda r: r["reserve_no"]) if r["market"] == market]
        assert mine == pasted[market], f"reserves differ from the paste for {market}"
    return len(table), {m: len(pasted[m]) for m in MARKETS}


def apportion(total, sizes):
    """Largest-remainder split of `total` places across groups in proportion to their sizes."""
    pool = sum(sizes.values())
    exact = {k: total * v / pool for k, v in sizes.items()}
    places = {k: int(exact[k]) for k in sizes}
    left = total - sum(places.values())
    by_remainder = sorted(sizes, key=lambda k: (-(exact[k] - places[k]), -sizes[k], k))
    for k in by_remainder[:left]:
        places[k] += 1
    return places


def pick_companies(rows, rng):
    market_places = apportion(DRAW_SIZE, Counter(r["market"] for r in rows))
    picked, not_drawn = [], []
    for market in MARKETS:
        in_market = [r for r in rows if r["market"] == market]
        places = apportion(market_places[market], Counter(r["difficulty"] for r in in_market))
        for difficulty in DIFFICULTIES:
            cell = [r for r in in_market if r["difficulty"] == difficulty]
            rng.shuffle(cell)
            keep = places.get(difficulty, 0)
            picked += cell[:keep]
            not_drawn += cell[keep:]
    return picked, not_drawn


def within_one(counts):
    return abs(counts[ARM_SDR] - counts[ARM_AGENT]) <= 1


def arms_balanced(companies, arm):
    total = Counter(arm.values())
    if total[ARM_SDR] != total[ARM_AGENT]:
        return False
    for attr in ("market", "difficulty"):
        split = defaultdict(Counter)
        for c in companies:
            split[c[attr]][arm[c["company_name"]]] += 1
        if not all(within_one(counts) for counts in split.values()):
            return False
    return True


def draw_arms(companies, rng):
    cells = defaultdict(list)
    for c in companies:
        cells[(c["market"], c["difficulty"])].append(c)
    for tries in range(1, MAX_TRIES + 1):
        arm = {}
        for key in sorted(cells):
            members = cells[key][:]
            rng.shuffle(members)
            n_agent = len(members) // 2
            if len(members) % 2 and rng.random() < 0.5:
                n_agent += 1
            for i, c in enumerate(members):
                arm[c["company_name"]] = ARM_AGENT if i < n_agent else ARM_SDR
        if arms_balanced(companies, arm):
            return arm, tries
    raise SystemExit("no balanced arm draw found")


def spread_ok(sample, everyone, size):
    for attr in ("market", "difficulty"):
        total = Counter(c[attr] for c in everyone)
        got = Counter(c[attr] for c in sample)
        for label, count in total.items():
            share = size * count / len(everyone)
            if not math.floor(share) <= got[label] <= math.ceil(share):
                return False
    return True


def pick_early(sdr_alone, rng):
    for tries in range(1, MAX_TRIES + 1):
        sample = rng.sample(sdr_alone, EARLY_SDR)
        if spread_ok(sample, sdr_alone, EARLY_SDR):
            return {c["company_name"] for c in sample}, tries
    raise SystemExit("no spread early block found")


def spread(groups):
    """Merge lists so each is spread evenly: always take from the list with the largest share left."""
    queues = {k: list(v) for k, v in groups.items() if v}
    sizes = {k: len(v) for k, v in queues.items()}
    merged = []
    while any(queues.values()):
        label = max((k for k in queues if queues[k]), key=lambda k: (len(queues[k]) / sizes[k], sizes[k], k))
        merged.append(queues[label].pop(0))
    return merged


def shuffled(items, rng):
    items = list(items)
    rng.shuffle(items)
    return items


def work_order(companies, rng):
    early = [c for c in companies if c["early_or_late"] == EARLY]
    late = [c for c in companies if c["early_or_late"] == LATE]
    early_seq = spread({m: shuffled([c for c in early if c["market"] == m], rng) for m in MARKETS})
    late_by_market = {}
    for market in MARKETS:
        in_market = [c for c in late if c["market"] == market]
        late_by_market[market] = spread(
            {arm: shuffled([c for c in in_market if c["arm"] == arm], rng) for arm in (ARM_AGENT, ARM_SDR)}
        )
    return early_seq + spread(late_by_market)


def run_draw(rows, seed):
    rng = random.Random(seed)
    picked, not_drawn = pick_companies(rows, rng)
    picked.sort(key=lambda r: r["list_no"])
    arm, arm_tries = draw_arms(picked, rng)
    companies = []
    for i, r in enumerate(picked, start=1):
        companies.append({
            "company_id": f"C{i:02d}",
            "list_no": r["list_no"],
            "company_name": r["company_name"],
            "market": r["market"],
            "difficulty": r["difficulty"],
            "crm": r["crm"],
            "arm": arm[r["company_name"]],
        })
    early_names, early_tries = pick_early([c for c in companies if c["arm"] == ARM_SDR], rng)
    for c in companies:
        c["early_or_late"] = EARLY if c["company_name"] in early_names else LATE
    for order, c in enumerate(work_order(companies, rng), start=1):
        c["work_order"] = order
    front = Counter()
    reserve_front = []
    for r in not_drawn:
        front[r["market"]] += 1
        reserve_front.append({
            "list_no": r["list_no"], "company_name": r["company_name"], "market": r["market"],
            "difficulty": r["difficulty"], "crm": r["crm"], "reserve_front_no": front[r["market"]],
        })
    return companies, reserve_front, {"arms": arm_tries, "early": early_tries}


def table(title, companies, row_attr, row_labels, col_attr, col_labels):
    lines = [title, f'{"":<10}' + "".join(f"{c:>12}" for c in col_labels) + f'{"Total":>8}']
    for label in row_labels:
        counts = [sum(1 for c in companies if c[row_attr] == label and c[col_attr] == col) for col in col_labels]
        if sum(counts):
            lines.append(f"{label:<10}" + "".join(f"{n:>12}" for n in counts) + f"{sum(counts):>8}")
    totals = [sum(1 for c in companies if c[col_attr] == col) for col in col_labels]
    lines.append(f'{"Total":<10}' + "".join(f"{n:>12}" for n in totals) + f"{sum(totals):>8}")
    return "\n".join(lines)


def report(rows, companies, reserve_front):
    out = []
    pool = Counter((r["market"], r["difficulty"]) for r in rows)
    drawn = Counter((c["market"], c["difficulty"]) for c in companies)
    out.append(f"POOL {len(rows)} -> DRAWN {len(companies)} (pool proportions), NOT DRAWN {len(reserve_front)}")
    out.append(f'{"":<14}{"pool":>4}{"drawn":>8}')
    for market in MARKETS:
        for difficulty in DIFFICULTIES:
            if pool[(market, difficulty)]:
                out.append(f"{market + ' ' + difficulty:<14}{pool[(market, difficulty)]:>4}{drawn[(market, difficulty)]:>8}")
    out.append("")
    out.append(table("ARMS BY MARKET", companies, "market", MARKETS, "arm", [ARM_SDR, ARM_AGENT]))
    out.append("")
    out.append(table("ARMS BY DIFFICULTY", companies, "difficulty", DIFFICULTIES, "arm", [ARM_SDR, ARM_AGENT]))
    out.append("")
    sdr = [c for c in companies if c["arm"] == ARM_SDR]
    out.append(table("SDR-ALONE COMPANIES: EARLY OR LATE, BY MARKET", sdr, "market", MARKETS, "early_or_late", [EARLY, LATE]))
    out.append("")
    out.append(table("SDR-ALONE COMPANIES: EARLY OR LATE, BY DIFFICULTY", sdr, "difficulty", DIFFICULTIES, "early_or_late", [EARLY, LATE]))
    out.append("")
    out.append("WORK ORDER")
    out.append(f'{"Order":<6}{"ID":<5}{"Block":<7}{"Arm":<12}{"Market":<8}{"Difficulty":<11}Company')
    for c in sorted(companies, key=lambda c: c["work_order"]):
        out.append(
            f'{c["work_order"]:<6}{c["company_id"]:<5}{c["early_or_late"]:<7}{c["arm"]:<12}'
            f'{c["market"]:<8}{c["difficulty"]:<11}{c["company_name"]}'
        )
    out.append("")
    out.append("NOT DRAWN: front of the reserve queue for their market")
    for r in reserve_front:
        out.append(f'{r["market"]:<6}front {r["reserve_front_no"]}  {r["difficulty"]:<8}{r["company_name"]}')
    return "\n".join(out)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("--list")
    parser.add_argument("--reserves")
    parser.add_argument("--paste")
    parser.add_argument("--seed", type=int)
    parser.add_argument("--note", action="append", default=[], help="provenance line for the record (repeatable)")
    parser.add_argument("--write", action="store_true")
    parser.add_argument("--out")
    parser.add_argument("--verify")
    args = parser.parse_args()

    if args.verify:
        record_path = Path(args.verify)
        record = json.loads(record_path.read_text(encoding="utf-8"))
        list_path = record_path.parent / record["list_file"]
        assert sha256(list_path) == record["list_sha256"], "the list file changed since the draw"
        companies, reserve_front, _ = run_draw(load_list(list_path), record["seed"])
        same = companies == record["companies"] and reserve_front == record["not_drawn"]
        print(f'VERIFY {"OK" if same else "FAILED"}: seed {record["seed"]}, {len(companies)} companies, '
              f'{len(reserve_front)} not drawn, recomputed and compared with {record_path.name}')
        sys.exit(0 if same else 1)

    rows = load_list(args.list)
    reserves = load_reserves(args.reserves) if args.reserves else []
    if args.paste:
        n_table, n_reserves = check_against_paste(rows, reserves, args.paste)
        print(f"PASTE CHECK OK: {n_table} list rows and reserves {n_reserves} match the paste line for line\n")
    companies, reserve_front, tries = run_draw(rows, args.seed)
    print(f"SEED {args.seed} | arm draw accepted on try {tries['arms']} | early block accepted on try {tries['early']}\n")
    print(report(rows, companies, reserve_front))

    if args.write:
        now = datetime.now().astimezone()
        out = Path(args.out) if args.out else HERE / "data-exports" / "master-log" / f"draw-{now:%Y-%m-%d}.json"
        if out.exists():
            raise SystemExit(f"{out} exists. A draw record is never overwritten.")
        record = {
            "what": "Capstone draw: companies, arms, early block, work order",
            "drawn_at": now.isoformat(timespec="seconds"),
            "seed": args.seed,
            "python": platform.python_version(),
            "script_sha256": sha256(__file__),
            "list_file": os.path.relpath(args.list, out.parent).replace("\\", "/"),
            "list_sha256": sha256(args.list),
            "pool_size": len(rows),
            "draw_size": DRAW_SIZE,
            "early_sdr_alone": EARLY_SDR,
            "rule": "pool proportions by market, then by difficulty within market (largest remainder); "
                    "arms exactly half each, per market and per difficulty differing by 1 at most",
            "accepted_on_try": tries,
            "notes": args.note,
            "companies": companies,
            "not_drawn": reserve_front,
            "pasted_reserves": reserves,
        }
        out.write_text(json.dumps(record, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
        print(f"\nWROTE {out}")


if __name__ == "__main__":
    main()
