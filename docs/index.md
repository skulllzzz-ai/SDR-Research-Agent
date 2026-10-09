---
title: I put an AI research agent next to our SDR for a week
description: Here is what we measured on 30 real companies, three arms, one blind judge. A 100xEngineers Cohort 7 capstone by Mayank Panchal.
---

# I put an AI research agent next to our SDR for a week. Here is what we measured.

## Summary

Our SDR researches every prospect by hand before he reaches out: 25 fields, a CRM check, a
verdict. I built an agent on Claude Code that drafts that row from public web pages and
tested it on 30 real companies in the USA and the Gulf, in three arms: the SDR alone, the
agent alone, and the SDR working from the agent's draft. Our SDR lead scored 56 rows blind.
The agent drafts a row in about two minutes at zero data-vendor cost. With the draft, the
SDR's time per accepted row dropped from 24.7 to 9.1 minutes, and the lead accepted 92
percent of the pair's rows against 80 percent of the SDR's own and 68 percent of the agent's.
The "likely dead end" flag, the feature I expected the most from, never fired once in 37
runs.

## Problem and insight

The SDR is on my team, and research is his biggest block of desk time. When I sat with him
and his lead, most of that time was finding and reading pages, not judging them. The
expensive case is a dead end found after ten or fifteen minutes: a competitor's arm, the
wrong segment, or an account we already have under another name. Why now? A headless agent
with two web tools now drafts a full row in two minutes for about 1.4 dollars, so the real
question was whether a working SDR gets faster without getting worse. I set the test up
before writing any code: three arms, a blind judge, five pass-or-fail checks, and we measured
the system, never the person.

## Solution overview

A headless Claude Code agent (Opus 5.5, high effort, 20-minute timeout) researches one
company per run with exactly two tools, web search and web fetch; every other connector is
locked out. Every cell carries a marker (sure, guessed, not found) and a source, and a
validator rejects any row with a missing field, a bad marker or a flag without a source. The
SDR sees the agent's row inside his own Excel sheet: markers as colours, sources as comments,
the flag in a grey band. The lead scores rows by code only.

![Figure 1: the trial design](figures/fig-1-trial-design.png)

*Figure 1. Three arms, one blind judge.*

## Build journey

- Four days went to tooling before the agent did anything: a login that expired every day and
  a CLI version the pinned model refused.
- The first model I tried burned a tenth of my five-hour usage window per company; Opus 5.5
  at high effort used two points, so I switched before the freeze.
- On batch night our shared enrichment account had 30 credits left of 2,500, so version 1
  ran on the web alone: 30 of 30 rows valid in 67 minutes, zero credits. Contacts paid the
  price: an email on 5 of 28 agent rows against 10 of 15 SDR rows.
- The SDR's first checks moved 32 cells on 7 of the 8 agent rows he saw; 84 percent stood.
  The causes: a field we never defined (which country is "main"), pages read but not used,
  pages not opened, facts on no public page. Version 2 fixed those in four prompt lines and
  ran on the last 7 pair companies.
- My scoring sheet asked the wrong question. A correct dead end can only fail "would you let
  the SDR call from this row?", and the lead marked all six dead-end rows No. I sent the six
  back with the question I meant, "is this dead end right?", and he confirmed three. Both
  readings are below.

## Results and metrics

| Arm | Rows scored | Accepted by the lead | SDR minutes per accepted row | Agent minutes per row, median |
|---|---|---|---|---|
| Agent alone | 28 | 19 (68%) | 0 | 1.9 |
| SDR alone | 15 | 12 (80%) | 24.7 | 0 |
| Pair, SDR checks the agent's row | 13 | 12 (92%); first scored 9 (69%) | 9.1; first scored 12.1 | 2.2 |
| Pair, v1 rows / v2 rows | 7 / 6 | 7 (100%) / 5 (83%); first scored 2 (33%) | 10.0 / 7.8; first scored 19.5 | |

![Figure 2: acceptance and SDR minutes by arm](figures/fig-2-results.png)

*Figure 2. Solid: the re-check reading; outline: the first scoring.*

- **Speed held.** The pair used 63 percent less SDR time per accepted row (51 as first
  scored). Two pair rows the SDR never marked complete were not scored; count them in at any
  score and the saving is 39 to 48 percent, under my "half" bar.
- **Quality held after the re-check.** 92 percent against 80, where I had allowed one row in
  fifteen below. On the first scoring it failed, 69 against 80.
- **The human still adds value.** On the same 13 companies the pair had 12 accepted rows
  against the agent's 8. At 13 companies that gap could still be chance.
- **Triage failed.** The pair's four dead ends took 6.0 SDR minutes on average against my bar
  of 5.
- **The flag never fired.** It fires on positive evidence of three rules: a competitor group,
  a lab-grown-only seller in the Gulf, a non-buyer segment. The six dead ends our people
  called were mostly companies nobody could identify, and "could not identify" was never a
  flag condition. On those six the lead sided with the agent three times and with the human
  three times.
- **Where the agent falls down.** USA 13 of 15 rows accepted, Gulf 6 of 13; medium companies
  11 of 11, hard 7 of 16. The lead's reasons on the nine rejected rows: a similar-named
  company 4 times, wrong target 3, too little information 3.
- **Cost.** 65 agent minutes, 409 web calls, about 42 dollars at list price for 30
  companies.
- **Limits.** One rater at a minute a row; one row moves a rate by 7 points.

![Figure 3: where the agent fails](figures/fig-3-where-the-agent-fails.png)

*Figure 3. The agent alone by market and difficulty, and why the lead rejected nine rows.*

## Learnings and next steps

- Define every field before two people fill it: eight of the first 32 edits were two honest
  readings of one undefined field.
- Ask the judge the question you mean. One line on the scoring sheet flipped the quality
  result and meant a re-check in the middle of the night.
- "Could not identify" needs its own flag. Four of the six dead ends were companies the agent
  could not pin down; it said so in prose and the flag stayed off. Next: flag it, and check
  our CRM, where "existing account under another name" lives.
- Telling similar-named companies apart is the agent's weakest step; the version 2 fix came
  too late for the agent-alone arm.
- What surprised me most was not the time saving, which I expected. It was how much the
  result hung on one line of a scoring sheet, and that the feature I expected the most from
  never fired once.
- Next: licensed contacts, the identity flag with the CRM lookup, a second SDR and rater to
  see if the 63 percent holds, and a separate research path for the Gulf, where both the
  agent and the SDR struggled.

## Team and roles

Solo capstone: discovery, measurement design, the build with Claude Code as my pair
programmer, and this write-up. Our SDR was the worker and our SDR lead the judge; we measured
the system, never the people, and both stay unnamed on purpose.

## Resources and inspirations

100xEngineers and the Collaborative Intelligence brief; Claude Code and Opus 5.5; our SDR and
SDR lead, who gave me their time in a working week. If you run a sales team and want to try
the same test, the repository has everything except our company list. Happy to compare notes.

- Repository (code, prompts, logs, rows, results): https://github.com/skulllzzz-ai/SDR-Research-Agent
- Live page: https://skulllzzz-ai.github.io/SDR-Research-Agent/
- Demo video: linked from the repository README.

@100xEngineers #0to100xEngineer
