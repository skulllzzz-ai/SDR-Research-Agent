# SDR research agent: a measured three-arm trial

An AI agent does the first research pass on a prospect company for a pre-sales SDR in B2B diamonds, and the SDR checks and completes it. Thirty real companies in the USA and the Gulf, three arms (SDR alone, agent alone, SDR with agent), one blind judge (the SDR lead), one working week. Built for the 100xEngineers Cohort 7 capstone, "Collaborative Intelligence" brief, by Mayank Panchal.

## The deliverables

| Deliverable | Where |
|---|---|
| Case study (the live page) | https://skulllzzz-ai.github.io/SDR-Research-Agent/ (source: [docs/index.md](docs/index.md)) |
| Code, prompts, logs, rows, results | [export/](export/): the build's public export, verbatim; start with [export/README.md](export/README.md) |
| Demo video | Recorded on 2026-10-09; the link lands here |

## Headline results

The re-check reading, where a dead-end row counts as acceptable when the lead confirmed the dead end (the designed measure); the first scoring in brackets.

| Arm | Rows scored | Acceptable | SDR minutes per acceptable row |
|---|---|---|---|
| Agent alone | 28 | 19 (68%) | 0 (the agent's own run: 1.9 min median, unattended) |
| SDR alone | 15 | 12 (80%) | 24.7 |
| SDR with agent | 13 | 12 (92%) [9 (69%)] | 9.1 [12.1] |

- Speed: the pair saves 63% of the SDR's minutes per acceptable row [51%]. With the 2 unscored pair rows counted at any score, 39% to 48%, under the "half" bar.
- Quality: 92% against 80% [69%: falsified as first scored].
- Human value: 12 against 8 of 13 on the same companies [9 against 8]; sign test p 0.22.
- Triage: the pair's 4 dead ends took a mean of 6.0 SDR minutes against a bar of 5: falsified.
- The flag: fired on 0 of 37 runs, so the condition holds only because nothing was flagged.

The full tables, the cuts by market and difficulty, and the agent's cost: [export/RESULTS.md](export/RESULTS.md).

## What is in this repository

- `docs/`: the case study and its three figures (GitHub Pages serves this folder).
- `export/`: the public export of the build, byte for byte as the build's export script wrote it, with `EXPORT-MANIFEST.json` carrying a SHA-256 for every file. `export/README.md` explains each file, what was withheld and how to read the rows. `export/code/` is the code, with `export/code/README-build.md` as its guide and `export/code/prompt/` the prompts. `export/DECISIONS.md` and `export/BUILD-LOG.md` are the trial's decisions and build passes as kept during the work.
- `verify_export.py`: recomputes the manifest's hashes.

```bash
python verify_export.py
```

## What was withheld, and what was not

- Shown: the 30 companies by name, every agent row in full with its markers and sources, the agent's verdicts and flags, the SDR's final verdict and dead-end kind per company, the lead's scores on the agent's rows.
- Withheld: every email and phone number (`[withheld]`, 58 places), the CRM status (`[masked]`), the SDR's own research cells and minutes and the lead's scores on the SDR's rows (aggregates only, over groups of 3 rows or more), the lead's comments, the names of the SDR and the lead, the source company list, the reserve companies and the competitor list.
- The scrub is a gated script with a negative control: `export/code/check_public_export.py`.

## Run it

Needs Python 3 with `openpyxl`. The checks that run offline, from `export/code/`:

```bash
python validate_row.py --self-test
python sdr_return.py --self-test
```

The agent itself (`run_agent.py`, `demo_one.py`) needs Claude Code logged in, plus the company list and the competitor block, which are not in this repository; `export/code/README-build.md` lists every script and what it needs.

## Author

Mayank Panchal, sole author. Built with Claude Code as the pair programmer, in its own session with an append-only build log and a decisions log. The SDR and the SDR lead gave their time in a working week and stay unnamed by design.
