"""The column contract shared by every script in this folder.

Source: the capstone's row format v3 (2026-09-29): one layout for the SDR workbook, the agent's
JSON and the master log. The 46 RESEARCH columns below must match the blank template in
templates/ header for header; build_sdr_release.py asserts that on every run.
No company data lives in this file.
"""

MARKETS = ["USA", "Gulf"]
DIFFICULTIES = ["easy", "medium", "hard"]
ARM_SDR = "SDR alone"
ARM_AGENT = "With agent"
ARMS = [ARM_SDR, ARM_AGENT]
EARLY = "Early"
LATE = "Late"

SHEET_INSTRUCTIONS = "Instructions"
SHEET_RESEARCH = "RESEARCH (fill this)"
SHEET_SNAPSHOT = "Pass 1 snapshot (paste only)"
SHEET_LISTS = "Lists (read only)"
HEADER_ROW = 2
FIRST_DATA_ROW = 3
LAST_DATA_ROW = 402

# The 46 columns of the RESEARCH sheet, in order.
SET_BY_MAYANK = ["Order", "Company ID", "Company name", "Market", "Arm"]
SDR_STATUS = ["Status", "Date started"]
TIME_STAMPS = [
    "Step 1 start", "Step 1 end", "Step 2 start", "Step 2 end",
    "Management asked at", "Management answer added at",
]
AGENT_FLAG = ["Agent flag", "Agent flag reason", "Agent flag source"]
PRODUCT_CATEGORIES = ["Gold", "Natural diamond", "LGD", "Gemstone", "Pearl", "Watch", "Jewellery"]
RESEARCH_FIELDS = (
    ["Parent company", "Turnover / revenue / market share", "Detailed research"]
    + PRODUCT_CATEGORIES
    + [
        "Product segments", "Type of customer", "Forevermark type", "Forevermark segment",
        "JBT segment", "JBT name", "Main country", "Second country", "Other countries",
        "Primary state", "Primary city", "Priority", "Contact names", "Emails", "Phones",
    ]
)
VERDICT_COLUMNS = ["Verdict", "Dead-end type", "Dead-end reason with source", "CRM status"]
NOTES = ["Notes"]
RESEARCH_COLUMNS = (
    SET_BY_MAYANK + SDR_STATUS + TIME_STAMPS + AGENT_FLAG + RESEARCH_FIELDS + VERDICT_COLUMNS + NOTES
)
assert len(RESEARCH_COLUMNS) == 46 and len(RESEARCH_FIELDS) == 25

# The agent's JSON: one entry per research field, keyed in snake_case.
FIELD_KEYS = {
    "Parent company": "parent_company",
    "Turnover / revenue / market share": "turnover_revenue_market_share",
    "Detailed research": "detailed_research",
    "Gold": "gold",
    "Natural diamond": "natural_diamond",
    "LGD": "lgd",
    "Gemstone": "gemstone",
    "Pearl": "pearl",
    "Watch": "watch",
    "Jewellery": "jewellery",
    "Product segments": "product_segments",
    "Type of customer": "type_of_customer",
    "Forevermark type": "forevermark_type",
    "Forevermark segment": "forevermark_segment",
    "JBT segment": "jbt_segment",
    "JBT name": "jbt_name",
    "Main country": "main_country",
    "Second country": "second_country",
    "Other countries": "other_countries",
    "Primary state": "primary_state",
    "Primary city": "primary_city",
    "Priority": "priority",
    "Contact names": "contact_names",
    "Emails": "emails",
    "Phones": "phones",
}
assert list(FIELD_KEYS) == RESEARCH_FIELDS
YN_KEYS = [FIELD_KEYS[name] for name in PRODUCT_CATEGORIES]

MARKERS = ["sure", "guessed", "not found"]
# What the SDR typed when unsure in a Y/N product cell on the first return: left as typed and counted
# as "not found" in the scoring (Mayank, 2026-10-05; DECISIONS.md, decision 2). From then on the SDR
# leaves the cell blank.
SDR_UNSURE_VALUES = ["Not Sure", "Not sure", "not sure", "NOT SURE", "Unsure", "unsure", "?"]
NOT_FOUND = "not found"
PRIORITIES = ["AAA", "AA", "A", "B", "C", "D", "E"]
VERDICT_PROSPECT = "Prospect"
VERDICT_DEAD_END = "Dead end"
# The four dead-end types on the sheet. The agent may use only the first three: an existing
# active account is found by the SDR's CRM check, which is outside the agent's declared scope.
DEAD_END_COMPETITOR = "Competitor"
DEAD_END_WRONG_TARGET = "Wrong target account"   # lab-grown only, Gulf only
DEAD_END_WRONG_SEGMENT = "Wrong segment"         # miner, bank, tool manufacturer, security firm
DEAD_END_ACTIVE = "Existing active account"
AGENT_DEAD_END_TYPES = [DEAD_END_COMPETITOR, DEAD_END_WRONG_TARGET, DEAD_END_WRONG_SEGMENT]
CRM_OUT_OF_SCOPE = "Not attempted, out of declared scope"
CRM_ON_SHEET = "Not attempted"   # the dropdown value the SDR sees on an agent row

# Master log: every field the analysis needs (spec exit condition 13). The Rows sheet may hold
# more columns than these; check_master_log.py asserts none of these is missing.
MASTER_LOG_CONTRACT = [
    "Company ID", "Company name", "Market", "Difficulty", "Arm", "Early or late", "Work order",
    "Prompt version", "Model", "Effort", "Agent start", "Agent end", "Agent minutes",
    "Apollo credits", "Flag value", "Flag reason", "Flag source", "Verdict", "Dead-end category",
    "Pass-1 start", "Pass-1 end", "Pass-2 start", "Pass-2 end",
    "Check start", "Check end", "Completing start", "Completing end",
    "Management asked at", "Management answer added at",
    "Acceptable", "Correctness", "Completeness", "Lead scoring minutes",
    "Validator result", "Retry count", "Notes",
]
ROW_AGENT_ALONE = "Agent alone"
ROW_SDR_ALONE = "SDR alone"
ROW_PAIR = "Pair"
