---
title: One SDR, one research agent, 30 real companies
description: What a three-arm trial measured. A 100xEngineers Cohort 7 capstone by Mayank Panchal.
---

# One SDR, one research agent, 30 real companies: what a three-arm trial measured

## TL;DR

A sales development rep (SDR) at a B2B luxury diamond business researches every prospect by
hand before outreach: 25 fields, a CRM check, a verdict. I built a research agent on Claude
Code that drafts that row from public web pages and ran a three-arm trial on 30 real
companies in the USA and the Gulf: SDR alone, agent alone, and the SDR completing the agent's
row. The SDR lead scored 56 rows blind. The agent drafts a row in a median of two minutes at
zero data-vendor cost. With the draft, the SDR's minutes per accepted row fell from 24.7 to
9.1, and the lead accepted 92 percent of the pair's rows against 80 percent of the SDR's own
and 68 percent of the agent's. The "likely dead end" flag never fired in 37 runs.

## Problem and insight

Pre-sales research is the SDR's largest block of desk time, most of it finding and reading
pages rather than judging them. The costliest outcome is a dead end found after ten or fifteen
minutes: a competitor's arm, the wrong segment, an existing account under another name. Why now: a headless agent with two web tools drafts a full row
in two minutes for about 1.4 dollars, so the question became whether a working SDR gets
faster without getting worse. The trial was designed before any code: three arms, a blind
judge, five pass-or-fail conditions, the system measured and never the person.

## Solution overview

A headless Claude Code agent (Opus 5.5, high effort, 20-minute timeout) researches one
company per run with exactly two tools, web search and web fetch, under a tool lock denying
every other connector. Every cell carries a marker (sure, guessed, not found)
and a source; a validator rejects a row with a missing field, an invalid marker or a flag
without its source. The SDR sees the agent's row inside his own Excel sheet, markers as
colours, sources as comments, the flag in a grey band. The lead scores rows that carry codes,
never arms.

![Figure 1: the trial design](figures/fig-1-trial-design.png)

*Figure 1. Three arms, one blind judge.*

## Build journey

- **Four days lost to tooling:** a login that expired daily and a CLI version the pinned
  model refused.
- **Pivot 1, the model.** The first model burned a tenth of the five-hour usage window per
  company; Opus 5.5 at high effort used two points. Changed before the freeze.
- **Pivot 2, the data vendor.** On batch night the shared enrichment account had 30 credits
  left of 2,500. Version 1 ran web-only: 30 of 30 rows valid in 67 minutes, zero credits.
  Contacts paid: an email on 5 of 28 agent rows against 10 of 15 SDR rows.
- **The first human checks** moved 32 cells on 7 of the 8 agent rows the SDR saw; 84 percent
  of cells stood. The causes: an undefined field (which country is "main"), pages read but
  not used, pages not opened, facts on no public page.
  Version 2 fixed those in four prompt lines, applied to the last 7 pair companies.
- **Blinding bit back.** A scoring file built before all rows existed would reveal the arm
  by row count alone; the lead now gets whole companies only.
- **The scoring sheet asked the wrong question.** "Would you let the SDR call from this row?"
  is a question a correct dead end can only fail; the lead marked all six dead-end rows No.
  Asked the question the design meant, "is this dead end right?", he confirmed three. Both
  readings are reported.

## Results and metrics

| Arm | Rows scored | Accepted by the lead | SDR minutes per accepted row | Agent minutes per row, median |
|---|---|---|---|---|
| Agent alone | 28 | 19 (68%) | 0 | 1.9 |
| SDR alone | 15 | 12 (80%) | 24.7 | 0 |
| Pair, SDR checks the agent's row | 13 | 12 (92%); first scored 9 (69%) | 9.1; first scored 12.1 | 2.2 |
| Pair, v1 rows / v2 rows | 7 / 6 | 7 (100%) / 5 (83%); first scored 2 (33%) | 10.0 / 7.8; first scored 19.5 | |

![Figure 2: acceptance and SDR minutes by arm](figures/fig-2-results.png)

*Figure 2. Solid: the re-check reading; outline: the first scoring.*

- **Speed held:** 63 percent less SDR time per accepted row (51 as first scored). Two pair
  rows the SDR never marked complete stayed unscored; counted in at any score, the saving is
  39 to 48 percent, under the bar.
- **Quality held on the re-check:** 92 against 80 percent, with one row in fifteen allowed
  below; as first scored it failed, 69 against 80.
- **The human still adds value:** on the same 13 companies the pair had 12 accepted rows
  against the agent's 8 (sign test p 0.22, within chance at this size).
- **Triage failed:** the pair's four dead ends took a mean of 6.0 SDR minutes against a bar
  of 5.
- **The flag never fired.** It fires on positive evidence of three rules: competitor group,
  lab-grown-only seller in the Gulf, non-buyer segment. The six human dead ends were mostly a
  company that could not be identified, and "could not identify" was never a flag condition.
  On the verdict, the lead sided with the agent on three and the human on three.
- **Where the agent fails:** USA 13 of 15 rows accepted, Gulf 6 of 13; medium companies 11
  of 11, hard 7 of 16. On the nine rejected rows: "a different company with a similar
  name" 4 times, "wrong target" 3, "too little information" 3.
- **Cost:** 65 agent minutes and 409 web calls for 30 companies, about 42 dollars at list
  price, zero data-vendor credits.
- **Limits:** one rater, a minute a row; one row moves a rate by 7 points; the first 7 SDR
  rows stamped in blocks.

![Figure 3: where the agent fails](figures/fig-3-where-the-agent-fails.png)

*Figure 3. The agent alone by market and difficulty, and why the lead rejected nine rows.*

## Learnings and next steps

- **Define every field before two arms fill it.** Eight of the first 32 edits were two
  honest readings of an undefined field.
- **Ask the judge the question you mean.** One line on the scoring sheet flipped the quality
  result and cost a day.
- **"Could not identify" needs its own flag.** Four of the six dead ends were companies the
  agent could not pin down; it said so in prose while the flag stayed off. Next: flag it,
  and check the CRM, where "existing account under another name" lives.
- **Disambiguation is the weakest step.** Four of nine rejections were a similar-named
  company; the v2 fix came too late for the agent-alone arm.
- **Open questions:** would a second SDR and rater reproduce the 63 percent, and does the
  Gulf need its own research path, where both the agent and the SDR struggled.
- **Next:** licensed contacts, the identity flag with the CRM lookup, a second SDR and rater,
  a longer run.

## Team and roles

Solo capstone: discovery, measurement design, the build with Claude Code as pair programmer,
the file exchange, the write-up. The SDR was the worker and the SDR lead
the judge; the system was measured, never the people, who stay unnamed by design.

## Resources and inspirations

100xEngineers and the Collaborative Intelligence brief; Claude Code and Opus 5.5; the SDR and
the lead, who gave their time in a working week.

- Repository (code, prompts, logs, rows, results): https://github.com/skulllzzz-ai/SDR-Research-Agent
- Live page: https://skulllzzz-ai.github.io/SDR-Research-Agent/
- Demo video: linked from the repository README.

@100xEngineers #0to100xEngineer
