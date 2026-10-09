# SDR research agent trial: public export

An AI agent did the first research pass on a prospect company for a pre-sales SDR in B2B diamonds, and the SDR checked and completed it. Three arms on 30 companies in the USA and the Gulf: the SDR alone, the agent alone, and the SDR starting from the agent's row. The SDR lead scored every finished row blind. This folder is what can be shared.

## What is in it

- `companies.csv`: the 30 companies of the trial by name and the 2 smoke-test companies, one line each: the arm; the agent's minutes, tool calls, verdict, flag with its reason and priority on v1 and, for the 7 companies run again, on v2; the SDR's final verdict with its dead-end kind; the lead's scores and scoring minutes on the agent's v1 row (28 scored; the others were never sent to the lead).
- `agent-rows.csv`: every agent row in full, one line per field: the 25 research fields with their marker (sure, guessed, not found) and source as the agent wrote them, then the flag, its reason, the verdict, its dead-end kind and reason, and the CRM status.
- `RESULTS.md`: the results by arm, by market and by difficulty, the five conditions of the hypothesis, the agent's cost and use of the usage window. Aggregates only.
- `code/`: the scripts that drew the arms, ran and locked the agent, validated its rows, built and read every exchange file, and computed the results; `code/prompt/` holds the prompts; `code/README-build.md` is the build's own guide.
- `DECISIONS.md` and `BUILD-LOG.md`: every rule decided during the trial and every build pass, as kept during the work.
- `EXPORT-MANIFEST.json`: when and from what this was built, what was withheld, and a hash of every file.

## What was withheld, and how

- Every email address and every phone number, wherever it stood (a field, a research note, a source link), is replaced by `[withheld]`: 58 items. Contact names and roles stay as the agent wrote them, each with its public source.
- The CRM status is masked (`[masked]`): whether a company is already a customer stays private.
- The SDR's own research cells, the SDR's minutes per company and the lead's scores on the SDR's rows appear only as aggregates by arm, market and difficulty in `RESULTS.md`, over groups of 3 rows or more. The SDR's final verdict and its dead-end kind are shown per company.
- The company list the companies were drawn from, the reserve companies and the competitor list stay out. The prompts keep `{{COMPETITOR_BLOCK}}` where the competitor list was put in at run time.
- In `BUILD-LOG.md`, 2 places that gave the SDR's minutes row by row now give the aggregate instead (listed in the manifest).
- Not included: the workbooks the SDR and the lead worked in, the master logs, the agent's raw streams and answers, the assembled prompts, the key that maps the lead's blind row codes to the arms.

## How to read it

- Arms: "SDR alone" companies were researched by the SDR from scratch; "SDR with agent" companies were researched by the agent first, then checked and completed by the SDR. Every company also has an agent-alone row: the agent's v1 row, which the lead scored beside the human row without knowing which was which.
- v1 is the agent as frozen before the batch (web search and web fetch only, no paid data). v2 changed the prompt after the SDR's first checks (DECISIONS.md 14) and was run only on the last 7 "SDR with agent" companies; the agent-alone arm stays v1.
- The flag is the agent's own early call that a company is likely a dead end (a competitor, a lab-grown-only seller in the Gulf, or a non-buyer segment). The verdict is the row's final call: Prospect or Dead end.
- The lead's scores: Acceptable (would you let the SDR call from this row?), Correctness (wrong facts in the key fields: None, One, More than one) and Completeness (Full, Partial, Thin).
- 2 "SDR with agent" companies have no completed SDR row yet, so neither of their rows was scored.
- The 6 dead-end rows were judged again on the verdict alone (DECISIONS.md 18); RESULTS.md leads with that reading.

## Limits

- One SDR, one lead, one model, 30 companies: every rate moves about 7 points per row in an arm of 15. No same-company comparison is beyond chance at this size.
- The lead spent about a minute per row: the scores are a read of each row, not a check of its sources.
- The v1 and v2 pair rows are different companies, so the gap between them is not a pure version effect.
- The agent's minutes describe this build on these days (model, tools, connection), not the method.

Built from `master-log-2026-10-09-2.xlsx` by `code/build_public_export.py`; the contact-and-SDR scrub gate is `code/check_public_export.py`.
