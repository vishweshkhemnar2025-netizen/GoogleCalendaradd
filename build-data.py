#!/usr/bin/env python3
"""
Build calendars.json and enrollments.json for the Term IV calendar page.

Inputs
------
1. calendars.template.csv   (next to this script)
       course_code,display_name,calendar_id
       The calendar_id column is IGNORED here - IDs live in CALENDAR_IDS below,
       which is the single source of truth the ACAD committee edits each term.

2. One .xlsx per course-section inside ENROLL_DIR, named exactly as the
       course_code in the CSV, e.g. "B2B M (A).xlsx", "CV.xlsx", "HRM(IR).xlsx".
       Each sheet has a header row containing "Roll No." and "Name".

Outputs (written next to this script)
-------------------------------------
   calendars.json     38 entries:  key -> { name, calId }
   enrollments.json    one entry per roll number: roll -> { name, sections[] }

The section KEY is derived deterministically from course_code so the two files
always stay consistent. Run this whenever the .xlsx files or CALENDAR_IDS change.
"""

import csv
import json
import re
import sys
from collections import OrderedDict
from pathlib import Path

import openpyxl

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
HERE = Path(__file__).parent
CALENDARS_CSV = HERE / "calendars.template.csv"

# Folder containing the 38 per-course .xlsx enrollment sheets.
# Override with:  python build-data.py "/path/to/Enrollment details"
ENROLL_DIR = Path(sys.argv[1]) if len(sys.argv) > 1 else (
    HERE.parent / "Enrollment details"
)

OUT_CALENDARS = HERE / "calendars.json"
OUT_ENROLLMENTS = HERE / "enrollments.json"

# ---------------------------------------------------------------------------
# The 38 Google Calendar IDs. THIS IS WHAT ACAD EDITS EACH TERM.
# Key = slug(course_code).  Get the ID from Google Calendar:
#   Calendar settings -> "Integrate calendar" -> "Calendar ID".
# ---------------------------------------------------------------------------
CALENDAR_IDS = {
    'B2B_M_SecA': 'c_5628375a0e79765e81c15da372de8b9186cbfc8c73412d120078aae0b599c36b@group.calendar.google.com',
    'B2B_M_SecB': 'c_d96a6ac8539d1d484fac51a68bece9e2261a0296512ebbdaf3605dbd12417990@group.calendar.google.com',
    'B2B_M_SecC': 'c_e86aaccd94a101fdf0c4bd1c1e9a06697453fe211e12a9304d50718fd3c8982a@group.calendar.google.com',
    'BM_SecA':    'c_01f1ef38f3d8da76339cbcd6505c9758e2f8b340280d686f7bd7bb648dc0ffb0@group.calendar.google.com',
    'BM_SecB':    'c_6288cb3281a8fb44873971abc787f70d2f0e9e1e9561a02bcea908e13c3a1e68@group.calendar.google.com',
    'CB_SecA':    'c_5a8afd8f03f3017eed0f36d813cb2f90e604e50f970c6a87cd611f42224e9c63@group.calendar.google.com',
    'CB_SecB':    'c_3c67e353459118f6bff4a4dc3c5dbe13d1e5e93c479aaac30bc050ba95a8e199@group.calendar.google.com',
    'CV':         'c_37065ae030f7b74036b46184db3712d2dfcdcb389ecfde6c255e420d409a5e61@group.calendar.google.com',
    'DSDT_SecA':  'c_4e429deea3e49efeed3d8dc9d667b918555348e99564f7663e95a3a001409bb8@group.calendar.google.com',
    'DSDT_SecB':  'c_0ca3d3e0399cf33f2ce6bf48fcdeb452637a76b22c9925b01b822f6f6fb869bd@group.calendar.google.com',
    'DSDT_SecC':  'c_f51547fd0c78aee3aed11ee20d2cca6490743d95a69f5e7d6c081d93b67ac0eb@group.calendar.google.com',
    'FD_SecA':    'c_c9f228bb1a16ec0f06c64b6b5ec336a58fc58f4e3b2150e44841dd7fefd8cc36@group.calendar.google.com',
    'FD_SecB':    'c_5a8c3f42cd6add3691460179bfd9a9f5e003e6e485eb6906b67d9246ec690d09@group.calendar.google.com',
    'FSA_SecA':   'c_85e2de9cd5600c2226122b120824d2e53130da69ce05cbc9cc8167451355f4e9@group.calendar.google.com',
    'FSA_SecB':   'c_473ed0a5e107dd892122bbb75cc4bba4744438bbe30a7a27410d476f209fd92b@group.calendar.google.com',
    'HRM_IR':     'c_f85fc819f6e26abdbea3f0aac90e5796d6fdb8ef1ab77d1e0cec49a4d0abcf78@group.calendar.google.com',
    'INV_SecA':   'c_d4d30791a48d84133d90fb67c7f70c9d76208a12555e827921f609ad3a93434a@group.calendar.google.com',
    'INV_SecB':   'c_5f08704f2e46d13f013941395f10c04cd9f01317c119d453358073074887b107@group.calendar.google.com',
    'MG_SecA':    'c_8df3ab99be903dfda41e14aad4c192c1df647f7c9c86f07030346c6374b5ec73@group.calendar.google.com',
    'MG_SecB':    'c_15135148b62b579e1d79b20d1c158dfddcdcd310b8c9f1c0331e890ae2e30c3c@group.calendar.google.com',
    'MSAIC_SecA': 'c_e8392a03d7baf699e37377899aefafe60aa32d48427699877aa8791a05bab893@group.calendar.google.com',
    'MSAIC_SecB': 'c_c01ba4eab309387d9a944bf6113b1a0e8376af277c2c8dc3c9a104ca412338d0@group.calendar.google.com',
    'MoB':        'c_f8251bb6f8fffbfa77d3c6816ac86075e51cea95f0b2735fb450acc8b03e41d9@group.calendar.google.com',
    'PA_SecA':    'c_c2814b31f6e703dfdad71bf7869121690d217ff3b9d92502dc1825153efe0a9c@group.calendar.google.com',
    'PA_SecB':    'c_59053100bac60669e2e98218576138dccd36a2a3bfbcb83990c4fd92894ac5ec@group.calendar.google.com',
    'PCG_SecA':   'c_9f5fe19b100d107511bef20eff6ad0de5741a6e167c027ddc605034ef73529d7@group.calendar.google.com',
    'PCG_SecB':   'c_c8a8aedad08f3c0a2b412645d7b2cdf6f0a06dc502e34a6653e32bf00932662d@group.calendar.google.com',
    'PEVC_SecA':  'c_9abec2f6e5e96316ec96fe732cde597ead16cebe3f42452ccb44fe6ef1540082@group.calendar.google.com',
    'PEVC_SecB':  'c_afbfe0e6b20e80ab43b3f39f82edc541978cf94f719d5df3c3b715d27bd16b2a@group.calendar.google.com',
    'PM':         'c_f39ff0d8b5daacd5e1226b26f9ffeefb6f3e1c92e3eef5fa1c916d47becf39e4@group.calendar.google.com',
    'PSM_SecA':   'c_6517d9726bf27f11badb2f472a7fa9ec4418c2f125a01a3c8a320fd1f9372d50@group.calendar.google.com',
    'PSM_SecB':   'c_6b549d0f97df88fe7f6b99f0d355dfe5d2874c12e9bd6965db500054534e9cfd@group.calendar.google.com',
    'Rev_Mgmt':   'c_353b39ea152788830d4d55a2fa963913f24cd021fbd9486c1a573e94b76fd521@group.calendar.google.com',
    'SCM_SecA':   'c_f4b0e34879e5e8332901e835612d43534c08bd0b6fd502e02792173e46069b6e@group.calendar.google.com',
    'SCM_SecB':   'c_9463410989735c9bf7ce10420ac3afbadee693b72d1f15808941c4fdf6025e5a@group.calendar.google.com',
    'SDM':        'c_ba77c76f5248eb9a02ff84130a6322c8feec297f61ceb80a3621e0dd526e7f2e@group.calendar.google.com',
    'SRC':        'c_5df13e3685dcc307447ca7cd9c8069afde9038d9b3498bf7aafbb55e1d5e85eb@group.calendar.google.com',
    'TS_ADR':     'c_05ff81e3c091d91e777ddd3243433016add2be587a4b52ee91d7106e724b09e2@group.calendar.google.com',
}

SECTION_SUFFIX = {" (A)": "_SecA", " (B)": "_SecB", " (C)": "_SecC"}


def slug(course_code: str) -> str:
    """'B2B M (A)' -> 'B2B_M_SecA', 'HRM(IR)' -> 'HRM_IR', 'CV' -> 'CV'."""
    base = course_code
    suffix = ""
    for raw, key_suffix in SECTION_SUFFIX.items():
        if base.endswith(raw):
            base = base[: -len(raw)]
            suffix = key_suffix
            break
    base = re.sub(r"[^A-Za-z0-9]+", "_", base).strip("_")
    return base + suffix


def read_calendars_csv():
    rows = []
    with open(CALENDARS_CSV, newline="", encoding="utf-8") as fh:
        for r in csv.DictReader(fh):
            code = (r["course_code"] or "").strip()
            if not code:
                continue
            rows.append((code, (r["display_name"] or "").strip()))
    return rows


def col_index(header_row, *names):
    for i, cell in enumerate(header_row):
        label = str(cell).strip().lower() if cell is not None else ""
        for n in names:
            if label == n.lower():
                return i
    return None


def read_enrollment_xlsx(path: Path):
    """Yield (roll_str, name_str) for every student row in the sheet."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    ri = col_index(header, "Roll No.", "Roll No", "Roll Number")
    ni = col_index(header, "Name")
    if ri is None or ni is None:
        raise SystemExit(f"{path.name}: could not find 'Roll No.'/'Name' header")
    for row in rows:
        if row is None or ri >= len(row):
            continue
        roll, name = row[ri], row[ni] if ni < len(row) else None
        if roll is None or str(roll).strip() == "":
            continue
        roll = str(roll).strip()
        if roll.endswith(".0"):
            roll = roll[:-2]
        name = (str(name).strip() if name is not None else "")
        yield roll, name
    wb.close()


def main():
    cal_rows = read_calendars_csv()
    if len(cal_rows) != 38:
        print(f"WARNING: expected 38 calendars, CSV has {len(cal_rows)}")

    calendars = OrderedDict()
    code_to_key = OrderedDict()
    for code, display in cal_rows:
        key = slug(code)
        code_to_key[code] = key
        if key not in CALENDAR_IDS:
            raise SystemExit(
                f"No calendar ID for key '{key}' (course_code '{code}'). "
                f"Add it to CALENDAR_IDS in build-data.py."
            )
        calendars[key] = {"name": display, "calId": CALENDAR_IDS[key]}

    # Build enrollments by reading each per-course .xlsx
    enroll = {}            # roll -> {"name":..., "sections": Orderedset-ish list}
    cal_order = list(calendars.keys())
    missing_files = []
    for code, display in cal_rows:
        key = code_to_key[code]
        xlsx = ENROLL_DIR / f"{code}.xlsx"
        if not xlsx.exists():
            missing_files.append(xlsx.name)
            continue
        for roll, name in read_enrollment_xlsx(xlsx):
            rec = enroll.setdefault(roll, {"name": name, "sections": set()})
            # Prefer a non-empty name; keep the first good one we see
            if not rec["name"] and name:
                rec["name"] = name
            rec["sections"].add(key)

    if missing_files:
        raise SystemExit("Missing enrollment files: " + ", ".join(missing_files))

    enrollments = OrderedDict()
    for roll in sorted(enroll.keys()):
        rec = enroll[roll]
        ordered = [k for k in cal_order if k in rec["sections"]]
        enrollments[roll] = {"name": rec["name"], "sections": ordered}

    OUT_CALENDARS.write_text(
        json.dumps(calendars, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )
    OUT_ENROLLMENTS.write_text(
        json.dumps(enrollments, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )

    # ---- stats ----
    counts = [len(v["sections"]) for v in enrollments.values()]
    print(f"calendars.json   : {len(calendars)} calendars")
    print(f"enrollments.json : {len(enrollments)} students")
    if counts:
        print(f"courses/student  : min {min(counts)}, max {max(counts)}, "
              f"avg {sum(counts)/len(counts):.1f}")
    no_cal = [k for k, v in calendars.items() if not v["calId"]]
    if no_cal:
        print("WARNING: calendars with blank calId: " + ", ".join(no_cal))


if __name__ == "__main__":
    main()
