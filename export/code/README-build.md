# SDR research agent: measured trial (pre-sales lead research)

An AI agent does the first research pass on a company before outreach: a full row with a
confidence marker on every field and a dead-end flag. The SDR checks and completes it. The trial
measures three arms on 30 companies (USA and Gulf): SDR alone, agent alone, SDR with agent.

- **Owner:** Mayank. Pre-sales sits with marketing, so the build lives here.
- **Spec and design (read-only from this repo):** `work/ai-mentor/tasks/specs/2026-09-28-capstone-agent-arm-and-harness.md`,
  plus the decision record, discovery document and row format beside it in
  `work/ai-mentor/courses/100x-engineers/final-capstone/`.
- **Build log:** [BUILD-LOG.md](BUILD-LOG.md): one line per pass, with the check that proved it.
- **Decisions log:** [DECISIONS.md](DECISIONS.md): rules decided during the trial (stamps, unsure cells,
  dead-end kinds, grey cells, return versions), who decided, and where each rule is applied.
- **AI register row:** `ai-transformation/use-cases/register.md`, row 5.

## Two zones

| Lives here, tracked in git | Lives in `data-exports/`, never in git |
|---|---|
| Code, the prompt template with a placeholder, the blank workbook template, this README | The company list, the competitor block, the draw, the master log, every SDR and lead file, every agent run |

The file exchange with the SDR and the lead is offline. `data-exports/INDEX.md` lists the folders
and who each file goes to.

## What each file does

| File | Job | Check it carries |
|---|---|---|
| `contract.py` | The 46 workbook columns, the agent's JSON keys, the master-log fields. One place | Asserted against the template on every release |
| `draw_arms.py` | Picks 30 of the listed companies in pool proportions, draws the arms, the early block and the work order from one recorded seed. Prints the balance table and the order | `--verify` re-computes the draw from the record |
| `check_draw.py` | Audits the saved draw record against the spec (21 checks) | `--self-test` plants 5 errors; each must be caught |
| `build_master_log.py` | Builds the master log workbook from the draw record, the release log and the filed SDR returns: the SDR's Step 1 and Step 2 stamps become pass 1 and pass 2 (SDR-alone rows) or check and completing (Pair rows), values as typed, the reader's issues in "Ingest notes", a Returns sheet. With `--agent-runs <v1 batch>` it also files the agent runs (Agent alone rows from that folder; each Pair row from the run the release log names, v1 or v2, its row file's hash checked) and the lead's filed returns (scores through the lead key; a Lead returns sheet). A new file each time | Re-reads the file: 16 checks (every filed row matches its return; every agent-alone row carries its run; every pair row the run it was built on; the lead's scores equal the filed return row for row), then the spreadsheet gate |
| `sdr_return.py` | Reads and checks a returned SDR workbook (stamps in order and not copied, listed values, grey cells untouched, a Pass 1 snapshot per SDR-alone row), reads the SDR's minutes per row with a stated basis (DECISIONS.md, decision 1), and records its filing in `master-log/returns.jsonl`, with the release it answers (from the release builder's list, never a `.FAILED` copy); a correction is filed as the next version with a note. Changes nothing in the return | `--self-test`: a clean made-up return gives 0 issues; seven planted faults are all found; the filing guards and the correction filing hold (14 checks) |
| `check_master_log.py` | Asserts the master log holds every field the analysis needs | Removes each field in turn; each must be reported |
| `lead_return.py` | Reads, checks and files the lead's scored return: the byte-identical copy, the release it answers, every row with its code in the same order, the row code, company, market and research cells unchanged against the release as sent (or the copy a re-issue stands on, when the re-issued file was saved over), scores from their lists, the stamps, the lead's minutes per row; records the filing in `master-log/lead-returns.jsonl`. Changes nothing | Negative controls on scratch copies of the real return (an edited research cell, two swapped rows, a score off its list, an end before its start, a deleted row, a note below the table) each fail their check; the self-test files a made-up return |
| `build_public_export.py` | The public export (DECISIONS.md 18): the companies by name with the agent's rows (emails and phones withheld, CRM masked), per company the SDR's verdict, the agent's minutes and the lead's scores on the agent's row; the SDR's side only in aggregates; the code, prompts, logs and a README. A rebuild moves the previous export aside; then the gate runs | `check_public_export.py` |
| `check_public_export.py` | The contact-and-SDR scrub gate over the export: no email, phone, CRM value, SDR stamp, SDR research cell, lead comment, reserve company or competitor-block line | `--self-test`: an email, a phone, a CRM value, an SDR stamp, an SDR research cell and a lead comment planted in scratch inputs must each appear 0 times in the export built from them, and a leaky copy must fail |
| `demo_one.py` | Runs the pinned agent live on one company for a demo: each search and page read as it happens, then the row; its own folder under `smoke-slice/`; emails and phones withheld on screen. `--replay <stream.jsonl>` shows a recorded run | Replay checked against the run record's tool count |
| `results.py` | The trial's results, aggregates only: by arm, the five conditions of the hypothesis, the dead-end rows read with the lead's comments, a second reading on prospect rows only (labelled as not set in advance), how far the unscored pair rows could move condition 1, keyword counts over the comments, verdict agreement between arms, cuts by market and difficulty, the agent's minutes, tool calls, credits and window use. Writes `RESULTS.md` and `master-log/results-YYYY-MM-DD.json` | On 9 Oct a second route (both returns read straight from their XML, its own minute rules, the run records) agreed on every headline number |
| `RESULTS.md` | The generated results, aggregates only, safe to commit | Regenerate with `results.py`; never edited by hand |
| `build_sdr_release.py` | Builds the next SDR file on the latest return (or the blank template). Adds a With-agent row only once its agent row exists: fields coloured by marker, sources as comments, the flag in the dark grey band. Prints the rules added on 5 Oct once on the Instructions sheet; restores a grey cell the SDR cleared, counted in the release log. Refused while the last release has no filed return answering it (a correction is not a second return). The release log names each agent row's run folder, hash, prompt version, model and effort, so pair v1 and pair v2 can be told apart. Every cell the release writes is Arial; the SDR's own cells keep their font as returned, counted. Never overwrites | 18 to 21 checks incl. the contamination count and the rules, then the gate |
| `build_lead_view.py` | Builds the lead's blind scoring file: whole companies only (a company goes in when both its rows are finished, its v1 agent row and its Complete SDR or pair row; DECISIONS.md 12), values only, "Not Sure" in a Y/N cell shown blank (DECISIONS.md 13), dead-end rows cut to the short-row fields, a company's two rows kept apart, the scoring steps, questions and Comment rule printed on the sheet (DECISIONS.md 17). Pass only the v1 batch. `--reissue <file> --reason "..."` issues the latest unworked lead release again under the next name (-2), the same rows, codes and order; refused for a file carrying scores The key stays in `master-log/lead-key.json` | 15 checks incl. whole companies, Y/N cells Y, N or blank, no marker colour, no comment, no arm or CRM wording, then the gate |
| `batch_status.py` | Where the agent batch stands, in aggregates: pair and SDR-alone companies done, valid rows, minutes, usage waits, the company in progress. `--wait-pairs` returns once the 15 pair companies are done | Reads the run records only |
| `selftest_exchange.py` | Runs the whole exchange on made-up data in a temporary folder: releases, Excel-saved returns, the filing and the master-log ingest of a return, a blanked grey cell restored by the next release, With-agent rows, the lead file, and every guard | 46 checks, incl. planted marker, flag and CRM values that must appear 0 times in the lead's file, a lead file issued again with its codes in order, a scored one refused, the lead's return filed (and refused a second time) with its scores and the agent runs on the right master-log rows, (and a check that the rows carrying them are in it), a company with one finished row held back, an SDR "Not Sure" shown blank, a filed correction that must not open the release guard, and a Calibri cell the release must catch only where it wrote |
| `validate_row.py` | Accepts or rejects one agent row. Since v2 a row with no contact named may give the company's main number on one phone line (DECISIONS.md 14) | `--self-test`: 3 valid rows pass, 14 planted bad rows are rejected |
| `run_agent.py` | Runs the agent, one company per headless Claude Code run, from an empty folder, with the tool lock. `batch` runs the 30 drawn companies, pair companies first; the drawn companies are refused before the freeze of the configured version (`agent_version` in the config: v1, then v2) and whenever the agent no longer matches it, checked before any model call. A later version runs only the With-agent companies named with `--only`, into its own `batch-<version>` folder; the agent-alone arm stays v1. With the web-only v1 a run is offered exactly web search and web fetch. A usage-limit rejection is waited out and the company run again. `self-test` checks the lock, the flags, the prompt assembly and the wait offline. `preflight` proves the login and the pinned model, lists the connectors, and reads the Apollo tool names out of two tool searches. A batch stops at the first login failure and resumes without overwriting a finished run | `lock-test`: a harmless read tool is attempted and denied, the other connectors are invisible, the credit cap hook blocks a call |
| `build_slice_file.py` | The slice rehearsal: one With-agent row for one company outside the 30, built like a real release, for the SDR to check (spec exit condition 8) | 13 checks (values, marker colours, comments, flag band, rules, empty SDR cells), then the gate |
| `freeze_v1.py` | Writes the freeze record of the configured version once (`freeze-v1.json`, then `freeze-v2.json`): config, model, effort, timeout, file hashes, assembled prompt, tool lists, the model decision and its evidence (smoke test, lock test, slice, failure path); `--revision` records the decision a later version rests on | `--self-test`: 19 checks, incl. every refusal of the record and of the runner's guard |
| `credit_cap_hook.py` | Hard cap on credit-spending Apollo calls per company | Unit-checked: the call over the cap is blocked |
| `xlsx_util.py`, `excel_open_check.ps1` | Styles, never-overwrite file names, the spreadsheet gate (XML parse + Excel open) | Two deliberately broken copies must fail |
| `agent_config.json` | The agent's version, model, effort, timeout, prompt files, the allowed tools (web search and web fetch since 6 Oct), the deny patterns | Frozen with hashes at each version's freeze |
| `prompt/research-agent-v1.md`, `prompt/research-agent-v2.md` | The agent's instructions, v1 (the agent-alone arm and the first 8 pair rows) and v2 (the last 7 pair rows: phones without a named contact, location as the head office of the company or its group, the company's About, Contact and product pages, a directory profile when there is no website). The competitor list is injected at run time | Hash recorded per batch and at each freeze |
| `prompt/repair-v1.md` | The one repair pass when a row fails validation (no tools, no new facts) | Hash recorded per batch |
| `templates/sdr-workbook-template-v3.xlsx` | The blank SDR workbook, copied unchanged from the design folder | SHA-256 `2b64752f…2aec` |

## Run it

Needs Python with `openpyxl` (`python` on this machine has it). Run from this folder.

```bash
# the draw: run ONCE; the record is never overwritten
python draw_arms.py --list data-exports/list/companies-2026-10-01.csv \
    --reserves data-exports/list/reserves-2026-10-01.csv \
    --paste data-exports/list/pasted-list-2026-10-01.md --seed 20261001 --write
python check_draw.py data-exports/master-log/draw-2026-10-01.json --self-test
python draw_arms.py --verify data-exports/master-log/draw-2026-10-01.json

# an SDR return: keep the file as received, copy it (never move) to the contract name, record the
# filing, then rebuild the master log; the reader prints aggregates only
cp -n "data-exports/sdr/in/<name as received>.xlsx" data-exports/sdr/in/sdr-YYYY-MM-DD-in.xlsx
python sdr_return.py file --received "<name as received>.xlsx" --filed sdr-YYYY-MM-DD-in.xlsx --received-on YYYY-MM-DD
python sdr_return.py check data-exports/sdr/in/sdr-YYYY-MM-DD-in.xlsx   # read-only, any time
# a correction to a filed return: never edit the filed file; copy it, make the change in the copy, file it
cp -n data-exports/sdr/in/sdr-YYYY-MM-DD-in.xlsx data-exports/sdr/in/sdr-YYYY-MM-DD-in-2.xlsx
python sdr_return.py file --filed sdr-YYYY-MM-DD-in-2.xlsx --correction-of sdr-YYYY-MM-DD-in.xlsx --note "what, who, when"

# the master log (a new version each time) and its column check; --agent-runs files the agent runs (always the v1 batch for the
# agent-alone arm; pair rows follow the release log) and the lead's filed returns
python build_master_log.py --draw data-exports/master-log/draw-2026-10-01.json --agent-runs data-exports/agent-runs/batch-v1
python check_master_log.py

# an SDR release: every company up to this work order, built on the latest return
python build_sdr_release.py --draw data-exports/master-log/draw-2026-10-01.json --through-order 7
# later releases add With-agent rows from an agent batch (release 2, 6 Oct: --through-order 20, batch-v1)
python build_sdr_release.py --draw data-exports/master-log/draw-2026-10-01.json --through-order 12 \
    --agent-runs data-exports/agent-runs/<batch>

# the lead's blind scoring file, built on the lead's latest return: whole companies only
# (a lead file that comes back unworked is filed as its own return, with a note in master-log/lead-returns.jsonl; DECISIONS.md 16)
python build_lead_view.py --draw data-exports/master-log/draw-2026-10-01.json \
    --agent-runs data-exports/agent-runs/batch-v1     # always the v1 batch: it is the agent-alone arm
# the lead's return: keep the file as received, copy it (never move) to the contract name, record the filing (aggregates only)
cp -n "data-exports/lead/in/<name as received>.xlsx" data-exports/lead/in/lead-YYYY-MM-DD-in.xlsx
python lead_return.py file --received "<name as received>.xlsx" --filed lead-YYYY-MM-DD-in.xlsx --received-on YYYY-MM-DD --note "..."
# the dead-end re-check: the dead-end rows of the lead's latest return, judged on the verdict alone (DECISIONS.md 18)
python build_lead_view.py --draw data-exports/master-log/draw-2026-10-01.json --recheck-dead-ends
# then the master log (above) and the results: RESULTS.md and master-log/results-YYYY-MM-DD.json, aggregates only
python results.py --agent-runs data-exports/agent-runs/batch-v1 data-exports/agent-runs/batch-v2
# the public export and its gate; the gate's planted negative control
python build_public_export.py --draw data-exports/master-log/draw-2026-10-01.json \
    --agent-runs data-exports/agent-runs/batch-v1 data-exports/agent-runs/batch-v2 \
    --smoke data-exports/smoke-slice/smoke-v1web-2026-10-06 data-exports/smoke-slice/smoke-v2-2026-10-07
python check_public_export.py --self-test --draw data-exports/master-log/draw-2026-10-01.json \
    --agent-runs data-exports/agent-runs/batch-v1 data-exports/agent-runs/batch-v2 \
    --smoke data-exports/smoke-slice/smoke-v1web-2026-10-06 data-exports/smoke-slice/smoke-v2-2026-10-07
# the demo: the pinned agent live on one company (uses the usage window, no credits)
python demo_one.py --company "<company name>" --market USA

# the whole exchange on made-up data (touches nothing real)
python selftest_exchange.py

# the agent
python validate_row.py --self-test
python run_agent.py preflight
python run_agent.py lock-test --out data-exports/smoke-slice/lock-test-YYYY-MM-DD
python run_agent.py research --companies <file.csv> --out data-exports/smoke-slice/<name>

# the slice: run the company the SDR or lead picked, build the SDR's one-row file, read it back
python run_agent.py research --companies <slice.csv> --out data-exports/smoke-slice/slice-YYYY-MM-DD
python build_slice_file.py --run data-exports/smoke-slice/slice-YYYY-MM-DD/<company id>
python sdr_return.py check data-exports/smoke-slice/slice-YYYY-MM-DD-in.xlsx --draw none

# the v1 freeze (once), then the batch (Mayank starts it); a stopped batch resumes with the same command
python freeze_v1.py --smoke <smoke folder> --lock-test <lock-test folder> --slice <slice folder>     --failure-path <failure folder> [--compare <earlier smoke folder>]
python run_agent.py batch > data-exports/agent-runs/batch-v1.stdout.txt 2>&1; echo "exit $?" >> data-exports/agent-runs/batch-v1.stdout.txt
# v2 (7 Oct): lock test and smoke test on v2, its freeze with --revision, then the last 7 With-agent companies
python run_agent.py batch --only C07,C19,C03,C09,C22,C20,C14 > data-exports/agent-runs/batch-v2.stdout.txt 2>&1
python batch_status.py                     # progress at any time; --wait-pairs returns when the 15 pair companies are done
python run_agent.py self-test              # offline checks of the lock, flags, prompt and the usage-window wait

# mechanics check only (small model, no Apollo tool, no credits): never an agent row
python run_agent.py research --model claude-haiku-4-5-20251001 --effort "" --web-only \
    --companies <file.csv> --out data-exports/smoke-slice/<name>
# the timeout path, without credits: a probe on the pinned model with a one-minute timeout
python run_agent.py research --web-only --timeout-minutes 1 --companies <file.csv> --out data-exports/smoke-slice/<name>
```

## What the agent runs need on this machine

- The terminal `claude` at version 2.1.251 or newer: Fable 5.1, the pinned model until 5 Oct, was
  refused by older versions. Update with `claude update`. The pinned model is `claude-opus-5-5` since
  5 Oct (DECISIONS.md, decision 6), effort `high` passed explicitly on every run.
- The terminal `claude` logged in (`claude auth login`, then `claude auth status` must say
  `"loggedIn": true`). This login is separate from the desktop app's. On 1 to 2 Oct 2026 it did
  not survive a day, so log in right before a batch. The runner checks it first, stops at the
  first login failure, and picks up where it stopped when the same command is started again.
- The Apollo connector connected on the account (`claude mcp list` shows it): the preflight still reads
  its tool names, to deny the send, buy and edit tools by name, although v1 allows none of them.
- Known to work: 2.1.289 (4 Oct 2026). That version loads 2 bundled plugins and the CLI's own
  skills and agents even with no user or project settings; none of them are Mayank's files, and the
  canary run checks that his instructions did not load. It also reports a denied tool call as a
  text-only system event, which the summariser skips. The run's first event and `claude mcp list`
  can name the same connector differently: the lock denies every server either of them names.
- Every headless run makes an empty per-project folder under `~/.claude/projects`; the runner removes
  it after the run when it holds no file.

## Rules that must hold

- The list of 30 is never run through the agent before v1 is frozen. Smoke tests and the slice use
  other companies.
- v1 is web-only (DECISIONS.md 7, 6 Oct 2026): two tools allowed, web search and web fetch; every
  connector on the account, Apollo included, denied as a whole; the Apollo send, buy and edit tools
  also denied by name; Apollo's credit tools capped at 0 in the hook. The last 30 lead credits are
  not spent. Apollo returns only in a v2, for the pair companies, as a logged revision (decision 9).
  Any credit use needs Mayank's yes first.
- The batch waits out a usage-limit rejection and resumes by itself; it stops when the reset is more
  than 6 hours away (decision 10). Every wait is logged in the batch folder's `usage-waits.jsonl`.
- The SDR file never holds an agent row for an SDR-alone company. The count is checked at every
  release and logged in `data-exports/master-log/releases.jsonl`.
- The difficulty label lives in the master log only.
- The batch is started by Mayank by hand. Nothing is scheduled.
