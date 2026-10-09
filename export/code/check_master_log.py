"""Assert the master log's Rows header holds every field the analysis needs (spec exit condition 13).

    python check_master_log.py                 # checks the newest master log
    python check_master_log.py <file.xlsx>

The fields are listed once, in contract.MASTER_LOG_CONTRACT. The check ends with a negative
control: each field is removed in turn from a copy of the header, and the check must then fail.
"""

import sys
from pathlib import Path

from openpyxl import load_workbook

from contract import MASTER_LOG_CONTRACT

HERE = Path(__file__).resolve().parent


def missing_fields(headers):
    return [name for name in MASTER_LOG_CONTRACT if name not in headers]


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) > 1:
        path = Path(sys.argv[1])
    else:
        logs = sorted((HERE / "data-exports" / "master-log").glob("master-log-*.xlsx"), key=lambda p: p.stat().st_mtime)
        path = logs[-1]
    headers = [cell.value for cell in load_workbook(path, read_only=True)["Rows"][1]]
    missing = missing_fields(headers)
    print(f"{path.name}: Rows header has {len(headers)} columns; contract needs {len(MASTER_LOG_CONTRACT)} fields")
    print(f'{"PASS" if not missing else "FAIL"}  every contract field is present{"" if not missing else "  missing: " + str(missing)}')
    caught = sum(1 for name in MASTER_LOG_CONTRACT if missing_fields([h for h in headers if h != name]) == [name])
    print(f'{"PASS" if caught == len(MASTER_LOG_CONTRACT) else "FAIL"}  negative control: '
          f"{caught} of {len(MASTER_LOG_CONTRACT)} removed fields were reported missing")
    sys.exit(0 if not missing and caught == len(MASTER_LOG_CONTRACT) else 1)


if __name__ == "__main__":
    main()
