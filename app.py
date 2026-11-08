"""
Oncology Consultant Rota Tool - V2.0 (Phase 2 / Option 1)

The Master Excel file is the database, configuration and output:
  Config_Shifts, Config_Staff  -> restore the Settings tab
  Rota, Previous Period Rota   -> assignment memory + future leave markers
  State                        -> master A order, current A/B, period end
  Summary, Legend              -> reporting

Engines (ported unchanged in logic from the V1 final baseline):
  * First "7-Day Continuous" shift -> Engine A (global DFS backtracking)
  * First "5-Day Weekday"   shift -> Engine B (daily/weekly greedy selector)
  * Nights                         -> inherited from A/B by weekday + BH week
Any additional shifts are shown as columns but are not auto-filled yet.
"""
import io
import math
import random
from collections import defaultdict
from datetime import datetime, timedelta

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# --- ROBUST BANK HOLIDAY LOADING ---
try:
    import holidays
    uk_holidays = holidays.country_holidays('GB', subdiv='ENG')
except ImportError:
    uk_holidays = None

# --- Configuration ---
B_WORK_DAYS = [0, 1, 2, 3, 4]
DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
SHIFT_TYPES = ["7-Day Continuous", "5-Day Weekday", "2-Day Weekend", "Nights Only"]
SHIFT_TYPE_ALIASES = {
    "7 DAYS (INCLUDES WEEKEND)": "7-Day Continuous",
    "7-DAY CONTINUOUS": "7-Day Continuous",
    "WEEKDAY SEPARATE FROM WEEKEND": "5-Day Weekday",
    "5-DAY WEEKDAY": "5-Day Weekday",
    "WEEKEND ONLY": "2-Day Weekend",
    "2-DAY WEEKEND": "2-Day Weekend",
    "NIGHTS ONLY": "Nights Only",
}
NIGHTS_MODES = ["Alternating Consultants", "With Primary / Weekday", "Split / Separate Pool", "None"]
B_ROLES = ["Second Parallel Rota + Nights", "Just Nights", "Independent Parallel Rota"]
GROUP_PAIRING_OPTIONS = ["Different Group Ideally", "Same Group Allowed"]
STAFF_NIGHTS_OPTIONS = ["Yes", "Separate Nights Only", "No"]
SHIFT_COLS = [
    "Parallel Consultants", "Shift Name", "Type", "Start Day", "Min Rest Weeks",
    "Nights Covered", "Flip Nights on BH", "Group Pairing",
    "Monday Off After Weekend", "Allow Dual Cover",
    "Nights Mode", "B Consultant Role",
]
EXPORT_SHIFT_COLS = [
    "Parallel Consultants", "Shift Name", "Type", "Start Day", "Min Rest Weeks",
    "Nights Covered", "Flip Nights on BH", "Group Pairing",
    "Monday Off After Weekend", "Allow Dual Cover",
]
STAFF_DATE_COLS = ["On Service From", "FTE Changed From"]
BASE_STAFF_COLS = [
    "Initials", "Full Name", "Active", "DIL Day", "Group",
    "Weekdays", "Nights", "Preferred Weeks in a Row",
    "Rest Days on Call", "Rest Days (Mon-Fri)",
] + STAFF_DATE_COLS
FTE_ORIGIN = datetime(2000, 1, 1).date()  # "since forever": FTE assumed constant before the first logged change
FTE_HISTORY_COLS = ["Initials", "Shift", "Effective From", "FTE"]
# SICK = sick leave: cannot be rostered AND creates no catch-up. AL/SL/NO/OFF cannot be rostered but must be caught up.
LEAVE_BLOCK = ["AL", "SL", "NO", "OFF", "SICK"]
CATCHUP_LEAVE = ["AL", "SL", "NO", "OFF"]
B_DAY_BLOCK = ["AL", "SL", "NO", "OFF", "SICK", "DIL", "REQ", "REQB"]
# Days that do NOT count as availability when working out each person's fair share:
NO_EXPOSURE_A = ["SICK"]
NO_EXPOSURE_B = ["SICK", "DIL", "REQ", "REQB"]  # DIL / REQ / REQB are structural; AL, SL, NO, OFF still count (catch-up)
NON_ROTA_SHEETS = {"CONFIG_SHIFTS", "CONFIG_STAFF", "CONFIG_FTE_HISTORY", "BH_DATES", "STATE", "SUMMARY", "LEGEND"}

DEFAULT_SHIFTS = [
    {
        "Shift Name": "Consultant A", "Type": "7-Day Continuous", "Start Day": "Friday", "Min Rest Weeks": 3,
        "Parallel Consultants": 2, "Nights Mode": "Alternating Consultants", "Nights Covered": "Tue,Thu,Fri,Sat,Sun",
        "Flip Nights on BH": True, "B Consultant Role": "Second Parallel Rota + Nights",
        "Group Pairing": "Different Group Ideally", "Allow Dual Cover": False, "Monday Off After Weekend": False,
    },
    {
        "Shift Name": "Consultant B", "Type": "5-Day Weekday", "Start Day": "Monday", "Min Rest Weeks": 3,
        "Parallel Consultants": 2, "Nights Mode": "Alternating Consultants", "Nights Covered": "Mon,Wed",
        "Flip Nights on BH": True, "B Consultant Role": "Second Parallel Rota + Nights",
        "Group Pairing": "Different Group Ideally", "Allow Dual Cover": False, "Monday Off After Weekend": False,
    },
]

DEFAULT_STAFF = [
    {"Initials": "GB", "Full Name": "Giuseppe Barone", "Active": True, "DIL Day": "Friday", "Group": "Solid", "Weekdays": True, "Nights": "Yes", "Preferred Weeks in a Row": 1, "Rest Days on Call": False, "Rest Days (Mon-Fri)": "", "FTE Consultant A": 1.0, "FTE Consultant B": 1.0},
    {"Initials": "SD", "Full Name": "Sarita Depani", "Active": True, "DIL Day": "Monday", "Group": "Neuro", "Weekdays": True, "Nights": "Yes", "Preferred Weeks in a Row": 1, "Rest Days on Call": False, "Rest Days (Mon-Fri)": "", "FTE Consultant A": 1.0, "FTE Consultant B": 1.0},
    {"Initials": "OS", "Full Name": "Olga Slater", "Active": True, "DIL Day": "Tuesday", "Group": "Solid", "Weekdays": True, "Nights": "Yes", "Preferred Weeks in a Row": 1, "Rest Days on Call": False, "Rest Days (Mon-Fri)": "", "FTE Consultant A": 1.0, "FTE Consultant B": 1.0},
    {"Initials": "MJ", "Full Name": "Mette Jorgensen", "Active": True, "DIL Day": "Friday", "Group": "Neuro", "Weekdays": True, "Nights": "Yes", "Preferred Weeks in a Row": 1, "Rest Days on Call": False, "Rest Days (Mon-Fri)": "", "FTE Consultant A": 1.0, "FTE Consultant B": 1.0},
    {"Initials": "CD", "Full Name": "Cat Duncan", "Active": True, "DIL Day": "Tuesday", "Group": "Solid", "Weekdays": True, "Nights": "Yes", "Preferred Weeks in a Row": 1, "Rest Days on Call": False, "Rest Days (Mon-Fri)": "", "FTE Consultant A": 1.0, "FTE Consultant B": 1.0},
    {"Initials": "KS", "Full Name": "Kavitha Srivatsa", "Active": True, "DIL Day": "Monday", "Group": "Neuro", "Weekdays": True, "Nights": "Yes", "Preferred Weeks in a Row": 1, "Rest Days on Call": False, "Rest Days (Mon-Fri)": "", "FTE Consultant A": 1.0, "FTE Consultant B": 1.0},
    {"Initials": "TC", "Full Name": "Tanzina Chowdhury", "Active": True, "DIL Day": "Tuesday", "Group": "Solid", "Weekdays": True, "Nights": "Yes", "Preferred Weeks in a Row": 1, "Rest Days on Call": False, "Rest Days (Mon-Fri)": "", "FTE Consultant A": 1.0, "FTE Consultant B": 1.0},
    {"Initials": "ES", "Full Name": "Elwira Szychot", "Active": True, "DIL Day": "Friday", "Group": "Neuro", "Weekdays": True, "Nights": "Yes", "Preferred Weeks in a Row": 1, "Rest Days on Call": False, "Rest Days (Mon-Fri)": "", "FTE Consultant A": 1.0, "FTE Consultant B": 1.0},
    {"Initials": "TN", "Full Name": "Trung Nguyen", "Active": True, "DIL Day": "Wednesday", "Group": "Solid", "Weekdays": True, "Nights": "Yes", "Preferred Weeks in a Row": 1, "Rest Days on Call": False, "Rest Days (Mon-Fri)": "", "FTE Consultant A": 1.0, "FTE Consultant B": 1.0},
]


# =============================================================================
# GENERIC HELPERS
# =============================================================================
def to_date(d):
    return d.date() if hasattr(d, 'date') else d


def gosh_fy_bounds(start_d, end_d=None):
    """Financial GOSH year (1 April to 31 March) for the period [start_d, end_d]."""
    s = to_date(start_d)
    e = to_date(end_d) if end_d is not None else s
    mid = s + (e - s) // 2
    y = mid.year if mid.month >= 4 else mid.year - 1
    return datetime(y, 4, 1).date(), datetime(y + 1, 3, 31).date()


def safe_ini(val):
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    s = str(val).strip().upper()
    if s in ["", "NAN", "NONE", "NAT", "NULL", "<NA>"]:
        return ""
    return s


def safe_raw(val):
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    s = str(val).strip()
    return "" if s.upper() in ["NAN", "NONE", "NAT", ""] else s


def match_name(val, consultants):
    val = safe_ini(val)
    if not val:
        return None
    for prefix in ["DR ", "DR. ", "MR ", "MR. ", "MS ", "MS. ", "MISS ", "PROF ", "PROF. "]:
        if val.startswith(prefix):
            val = val[len(prefix):].strip()
    v_words = [w for w in val.replace(".", " ").split() if len(w) > 1]
    for c in consultants:
        ini = safe_ini(c.get('Initials'))
        if not ini:
            continue
        fname = str(c.get('Full Name', '')).strip().upper()
        f_words = [w for w in fname.replace(".", " ").split() if len(w) > 1]
        if val == ini or val == fname:
            return ini
        if any(vw in f_words for vw in v_words):
            return ini
    return None


def get_bh_group(d):
    if uk_holidays is None or d not in uk_holidays:
        return None
    name = uk_holidays.get(d).upper()
    if "NEW YEAR" in name: return "NEWYEAR"
    if "EASTER" in name or "GOOD FRIDAY" in name: return "EASTER"
    if "MAY" in name or "SPRING" in name or "JUBILEE" in name or "CORONATION" in name:
        return "EARLY_MAY" if d.day <= 15 else "LATE_MAY"
    if "SUMMER" in name or "AUGUST" in name: return "AUGUST"
    if "CHRISTMAS" in name or "BOXING" in name: return "CHRISTMAS"
    return "OTHER"


def parse_any_date(d_val):
    if d_val is None:
        return None
    try:
        if pd.isna(d_val):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(d_val, str):
        ts = pd.to_datetime(d_val, dayfirst=True, errors='coerce')
    elif isinstance(d_val, datetime):
        return d_val.date()
    else:
        ts = pd.to_datetime(d_val, errors='coerce')
    return None if pd.isna(ts) else ts.date()


# =============================================================================
# CONFIG TABLE HELPERS (Settings tab)
# =============================================================================
DAY_SHORT_MAP = {
    "MON": 0, "MONDAY": 0,
    "TUE": 1, "TUES": 1, "TUESDAY": 1,
    "WED": 2, "WEDS": 2, "WEDNESDAY": 2,
    "THU": 3, "THUR": 3, "THURS": 3, "THURSDAY": 3,
    "FRI": 4, "FRIDAY": 4,
    "SAT": 5, "SATURDAY": 5,
    "SUN": 6, "SUNDAY": 6,
}


def safe_bool(val, default=False):
    if val is None:
        return default
    try:
        if pd.isna(val):
            return default
    except (TypeError, ValueError):
        pass
    if isinstance(val, bool):
        return val
    s = str(val).strip().upper()
    if s in ("TRUE", "YES", "1", "T", "Y"):
        return True
    if s in ("FALSE", "NO", "0", "F", "N"):
        return False
    return default


SHORT_DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
WEEKDAY_SHORT_DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri"]


def parse_weekdays(val, default=None):
    if isinstance(val, (list, tuple, set)):
        out = []
        for item in val:
            p = str(item).strip().upper()
            if p in DAY_SHORT_MAP and DAY_SHORT_MAP[p] not in out:
                out.append(DAY_SHORT_MAP[p])
        return sorted(out)
    s = safe_raw(val).upper()
    if not s:
        return list(default) if default is not None else []
    if s in ("NONE", "NO", "FALSE", "-", "[]"):
        return []
    if s in ("ALL", "7 DAYS", "EVERY DAY", "MON-SUN"):
        return [0, 1, 2, 3, 4, 5, 6]
    if s in ("WEEKDAYS", "MON-FRI"):
        return [0, 1, 2, 3, 4]
    if s in ("WEEKEND", "WEEKENDS", "SAT-SUN"):
        return [5, 6]
    out = []
    for part in s.strip("[]").replace("'", "").replace('"', "").replace(";", ",").replace("/", ",").split(","):
        p = part.strip().upper()
        if p in DAY_SHORT_MAP and DAY_SHORT_MAP[p] not in out:
            out.append(DAY_SHORT_MAP[p])
    return sorted(out) if out else (list(default) if default is not None else [])


def to_day_tags(val, default=None):
    return [SHORT_DAYS[i] for i in parse_weekdays(val, default)]


def from_day_tags(val, empty_str="None"):
    if isinstance(val, (list, tuple, set)):
        days = [SHORT_DAYS[i] for i in parse_weekdays(val, [])]
        return ",".join(days) if days else empty_str
    s = safe_raw(val)
    if not s:
        return ""
    if s.upper() in ("NONE", "NO", "FALSE", "-", "[]"):
        return empty_str
    days = [SHORT_DAYS[i] for i in parse_weekdays(s, [])]
    return ",".join(days) if days else empty_str


def _default_nights_for_type(t_str):
    if t_str == "7-Day Continuous":
        return "Tue,Thu,Fri,Sat,Sun"
    if t_str == "5-Day Weekday":
        return "Mon,Wed"
    if t_str == "2-Day Weekend":
        return "Sat,Sun"
    if t_str == "Nights Only":
        return "Mon,Wed"
    return "None"


def clean_shifts(df_shifts):
    df = df_shifts.copy() if df_shifts is not None else pd.DataFrame()
    has_explicit_b_role = "B Consultant Role" in df.columns
    defaults = [
        ("Parallel Consultants", 2),
        ("Shift Name", ""),
        ("Type", "5-Day Weekday"),
        ("Start Day", "Monday"),
        ("Min Rest Weeks", 3),
        ("Nights Covered", None),
        ("Flip Nights on BH", True),
        ("Group Pairing", "Different Group Ideally"),
        ("Monday Off After Weekend", False),
        ("Allow Dual Cover", False),
        ("Nights Mode", "Alternating Consultants"),
        ("B Consultant Role", ""),
    ]
    for col, default in defaults:
        if col not in df.columns:
            df[col] = default
    df = df[SHIFT_COLS].copy()
    df["Shift Name"] = df["Shift Name"].apply(safe_raw)
    df = df[df["Shift Name"] != ""].reset_index(drop=True)
    df["Type"] = df["Type"].apply(lambda v: SHIFT_TYPE_ALIASES.get(safe_raw(v).upper(), safe_raw(v) if safe_raw(v) in SHIFT_TYPES else "5-Day Weekday"))
    df["Start Day"] = df["Start Day"].apply(lambda v: safe_raw(v).capitalize() if safe_raw(v).capitalize() in DAY_NAMES else "Monday")
    df["Min Rest Weeks"] = pd.to_numeric(df["Min Rest Weeks"], errors="coerce").fillna(3).astype(int)
    df["Parallel Consultants"] = pd.to_numeric(df["Parallel Consultants"], errors="coerce").fillna(2).clip(1, 3).astype(int)
    df["Nights Mode"] = df["Nights Mode"].apply(lambda v: safe_raw(v) if safe_raw(v) in NIGHTS_MODES else "Alternating Consultants")
    nights_cov = []
    for idx, row in df.iterrows():
        raw_nc = row.get("Nights Covered")
        if isinstance(raw_nc, (list, tuple, set)):
            s_nc = from_day_tags(raw_nc, empty_str="None")
        else:
            s_nc = from_day_tags(raw_nc, empty_str="None")
            if not s_nc:
                s_nc = _default_nights_for_type(row["Type"])
        nights_cov.append(s_nc)
    df["Nights Covered"] = nights_cov
    df["Flip Nights on BH"] = [safe_bool(v, True) for v in df["Flip Nights on BH"]]
    b_roles = []
    for idx, row in df.iterrows():
        raw_br = safe_raw(row.get("B Consultant Role")) if has_explicit_b_role else ""
        if row["Type"] == "Nights Only":
            b_roles.append("Just Nights")
        elif raw_br in B_ROLES:
            b_roles.append(raw_br)
        elif row["Nights Covered"] == "None":
            b_roles.append("Independent Parallel Rota")
        else:
            b_roles.append("Second Parallel Rota + Nights")
    df["B Consultant Role"] = b_roles
    df["Group Pairing"] = df["Group Pairing"].apply(lambda v: safe_raw(v) if safe_raw(v) in GROUP_PAIRING_OPTIONS else "Different Group Ideally")
    df["Allow Dual Cover"] = [safe_bool(v, False) for v in df["Allow Dual Cover"]]
    df["Monday Off After Weekend"] = [safe_bool(v, False) for v in df["Monday Off After Weekend"]]
    return df.reset_index(drop=True)


def shift_names_of(df_shifts):
    if df_shifts is None or df_shifts.empty or "Shift Name" not in df_shifts.columns:
        return []
    names = []
    for s in df_shifts["Shift Name"]:
        s = safe_raw(s)
        if s and s not in names:
            names.append(s)
    return names


def sync_staff_fte_columns(df_staff, df_shifts):
    """Guarantees an 'FTE <Shift Name>' column for every shift; drops FTE columns of deleted shifts."""
    df_staff = df_staff.copy() if df_staff is not None else pd.DataFrame()
    col_defaults = {
        "Active": True,
        "Group": "Others",
        "Weekdays": True,
        "Nights": "Yes",
        "Preferred Weeks in a Row": 1,
        "Rest Days on Call": False,
        "Rest Days (Mon-Fri)": "",
    }
    for col in BASE_STAFF_COLS:
        if col not in df_staff.columns:
            df_staff[col] = col_defaults.get(col, pd.NaT if col in STAFF_DATE_COLS else "")
    for col in STAFF_DATE_COLS:
        df_staff[col] = pd.to_datetime(df_staff[col], errors="coerce")

    df_staff["Weekdays"] = [safe_bool(v, True) for v in df_staff["Weekdays"]]
    df_staff["Nights"] = [
        safe_raw(v) if safe_raw(v) in STAFF_NIGHTS_OPTIONS else ("No" if str(v).strip().upper() in ("FALSE", "NO", "0") else "Yes")
        for v in df_staff["Nights"]
    ]
    df_staff["Preferred Weeks in a Row"] = pd.to_numeric(df_staff["Preferred Weeks in a Row"], errors="coerce").fillna(1).clip(1, 4).astype(int)
    df_staff["Rest Days on Call"] = [safe_bool(v, False) for v in df_staff["Rest Days on Call"]]
    df_staff["Rest Days (Mon-Fri)"] = [from_day_tags(v, empty_str="") for v in df_staff["Rest Days (Mon-Fri)"]]

    fte_cols = [f"FTE {s}" for s in shift_names_of(df_shifts)]
    for fte_col in fte_cols:
        if fte_col not in df_staff.columns:
            # Legacy V1 files used "FTE A"/"FTE B"
            legacy = None
            if fte_col == "FTE Consultant A" and "FTE A" in df_staff.columns: legacy = "FTE A"
            elif fte_col == "FTE Consultant B" and "FTE B" in df_staff.columns: legacy = "FTE B"
            df_staff[fte_col] = pd.to_numeric(df_staff[legacy], errors="coerce").fillna(1.0).clip(0.0, 1.0) if legacy else 1.0
        else:
            df_staff[fte_col] = pd.to_numeric(df_staff[fte_col], errors="coerce").fillna(1.0).clip(0.0, 1.0)

    stale = [c for c in df_staff.columns if str(c).startswith("FTE ") and c not in fte_cols and c not in STAFF_DATE_COLS]
    df_staff = df_staff.drop(columns=stale)
    others = [c for c in df_staff.columns if c not in BASE_STAFF_COLS and c not in fte_cols]
    return df_staff[BASE_STAFF_COLS + fte_cols + others].reset_index(drop=True)


def clean_staff_for_engine(df_staff):
    df = df_staff.copy()
    if df.empty:
        return df
    df["Initials"] = df["Initials"].apply(safe_ini)
    df = df[df["Initials"] != ""]
    df["Active"] = [safe_bool(v, True) for v in df["Active"]]
    df["Group"] = df["Group"].apply(lambda v: safe_raw(v) or "Others")
    df["DIL Day"] = df["DIL Day"].apply(safe_raw)
    if "Weekdays" in df.columns:
        df["Weekdays"] = [safe_bool(v, True) for v in df["Weekdays"]]
    if "Nights" in df.columns:
        df["Nights"] = [
            safe_raw(v) if safe_raw(v) in STAFF_NIGHTS_OPTIONS else ("No" if str(v).strip().upper() in ("FALSE", "NO", "0") else "Yes")
            for v in df["Nights"]
        ]
    if "Preferred Weeks in a Row" in df.columns:
        df["Preferred Weeks in a Row"] = pd.to_numeric(df["Preferred Weeks in a Row"], errors="coerce").fillna(1).clip(1, 4).astype(int)
    if "Rest Days on Call" in df.columns:
        df["Rest Days on Call"] = [safe_bool(v, False) for v in df["Rest Days on Call"]]
    if "Rest Days (Mon-Fri)" in df.columns:
        df["Rest Days (Mon-Fri)"] = [from_day_tags(v, empty_str="") for v in df["Rest Days (Mon-Fri)"]]
    return df.reset_index(drop=True)


# --- Effective-dated FTE log (stored in the Master file as 'Config_FTE_History') ---
def normalize_fte_history(df):
    if df is None or df.empty or not set(FTE_HISTORY_COLS) <= set(df.columns):
        return pd.DataFrame(columns=FTE_HISTORY_COLS)
    out = df[FTE_HISTORY_COLS].copy()
    out["Initials"] = out["Initials"].apply(safe_ini)
    out["Shift"] = out["Shift"].apply(safe_raw)
    out["Effective From"] = [parse_any_date(x) for x in out["Effective From"]]
    out["FTE"] = pd.to_numeric(out["FTE"], errors="coerce")
    out = out.dropna(subset=["Effective From", "FTE"])
    out = out[(out["Initials"] != "") & (out["Shift"] != "")]
    return out.sort_values(["Initials", "Shift", "Effective From"]).reset_index(drop=True)


def history_from_staff(df_staff, df_shifts):
    """Seed the log from a staff table: each person's FTE is assumed to have applied 'since forever'."""
    rows = []
    for rec in (df_staff.to_dict("records") if df_staff is not None else []):
        ini = safe_ini(rec.get("Initials"))
        if not ini:
            continue
        for shift in shift_names_of(df_shifts):
            f = pd.to_numeric(rec.get(f"FTE {shift}"), errors="coerce")
            rows.append({"Initials": ini, "Shift": shift, "Effective From": FTE_ORIGIN, "FTE": 1.0 if pd.isna(f) else float(f)})
    return normalize_fte_history(pd.DataFrame(rows, columns=FTE_HISTORY_COLS))


def update_fte_history(hist_df, staff_df, shifts_df, default_effective):
    """Compare today's FTE table with the log. A changed FTE becomes a new log row effective from the
    'FTE Changed From' date (or from the start of the new period if none is typed). Pure function: the
    same inputs always give the same log, so regenerating a period is repeatable."""
    rows = normalize_fte_history(hist_df).to_dict("records")
    default_effective = to_date(default_effective)
    for rec in staff_df.to_dict("records"):
        ini = safe_ini(rec.get("Initials"))
        if not ini:
            continue
        changed_from = parse_any_date(rec.get("FTE Changed From"))
        for shift in shift_names_of(shifts_df):
            cur = pd.to_numeric(rec.get(f"FTE {shift}"), errors="coerce")
            cur = 1.0 if pd.isna(cur) else float(cur)
            entries = sorted([r for r in rows if r["Initials"] == ini and r["Shift"] == shift], key=lambda r: r["Effective From"])
            if not entries:
                rows.append({"Initials": ini, "Shift": shift, "Effective From": FTE_ORIGIN, "FTE": cur})
                continue
            last = entries[-1]
            if abs(last["FTE"] - cur) > 1e-9:
                eff = max(changed_from or default_effective, last["Effective From"] + timedelta(days=1))
                rows.append({"Initials": ini, "Shift": shift, "Effective From": eff, "FTE": cur})
            elif changed_from and len(entries) > 1 and last["Effective From"] != changed_from:
                last["Effective From"] = max(changed_from, entries[-2]["Effective From"] + timedelta(days=1))
    return normalize_fte_history(pd.DataFrame(rows, columns=FTE_HISTORY_COLS))


def build_fte_timeline(hist_df, shift):
    tl = {}
    for r in normalize_fte_history(hist_df).to_dict("records"):
        if r["Shift"] == shift:
            tl.setdefault(r["Initials"], []).append((r["Effective From"], float(r["FTE"])))
    return tl


def pick_engine_shifts(df_shifts):
    """Maps shifts to Engine A (Primary), Engine B (Secondary / Nights / Weekend), and any 3rd/extra shifts."""
    df = clean_shifts(df_shifts)
    recs = df.to_dict("records")
    a_cfg = b_cfg = None
    for rec in recs:
        if a_cfg is None and rec["Type"] == "7-Day Continuous":
            a_cfg = rec
        elif b_cfg is None and rec["Type"] == "5-Day Weekday":
            b_cfg = rec
    if a_cfg is None and recs:
        a_cfg = recs[0]
        if b_cfg is not None and b_cfg["Shift Name"] == a_cfg["Shift Name"]:
            b_cfg = None
    if b_cfg is None:
        for rec in recs:
            if a_cfg is None or rec["Shift Name"] != a_cfg["Shift Name"]:
                b_cfg = rec
                break
    used = {c["Shift Name"] for c in (a_cfg, b_cfg) if c}
    extras = [n for n in df["Shift Name"] if n not in used]
    return a_cfg, b_cfg, extras


# =============================================================================
# MEMORY PARSER (Previous rota / Master Excel)
# =============================================================================
def empty_memory():
    return {"counts": {}, "master_a_seq": [], "cycle_a": {}, "curr_a": None, "curr_b": None,
            "last_a_dates": {}, "hist_b_dates": {}, "prev_rota_df": None,
            "leave_df": pd.DataFrame(), "last_b_week": {}, "period_end": None, "warnings": []}


def _detect_rota_cols(df, a_name, b_name):
    cols = list(df.columns)
    up = {c: str(c).strip().upper() for c in cols}
    date_col = next((c for c in cols if "DATE" in up[c]), cols[0])

    def find(exact, tests):
        for c in cols:
            if exact and up[c] == exact:
                return c
        for c in cols:
            if any(t(up[c]) for t in tests):
                return c
        return None

    col_a = find(a_name.upper(), [lambda s: "CONSULTANT A" in s, lambda s: "ATTENDING" in s, lambda s: s == "A"])
    col_b = find(b_name.upper(), [lambda s: "CONSULTANT B" in s, lambda s: "B CONSULTANT" in s, lambda s: s == "B"])
    col_n = find("NIGHTS", [lambda s: "NIGHT" in s])
    if col_a is None and len(cols) > 2: col_a = cols[2]
    if col_b is None and len(cols) > 3: col_b = cols[3]
    if col_n is None and len(cols) > 4: col_n = cols[4]
    return date_col, col_a, col_b, col_n


def parse_previous_rota(uploaded_file, consultants, sel_s, ytd_s, ytd_e, a_name="Consultant A", b_name="Consultant B"):
    """Reads 'Previous Period Rota' + 'Rota' (Rota wins on overlapping dates) to build cumulative memory."""
    mem = empty_memory()
    counts = defaultdict(lambda: {"A_days": 0, "B_days": 0, "Nights": 0, "BH_days": 0, "BH_groups": set()})
    sel_s, ytd_s, ytd_e = to_date(sel_s), to_date(ytd_s), to_date(ytd_e)

    active_initials = [safe_ini(c.get('Initials')) for c in consultants if c.get('Active', True) and safe_ini(c.get('Initials'))]
    name_map = {safe_ini(c.get('Initials')): str(c.get('Full Name', '')).strip() for c in consultants if safe_ini(c.get('Initials'))}
    osf_map = {safe_ini(c.get('Initials')): parse_any_date(c.get('On Service From')) for c in consultants if safe_ini(c.get('Initials'))}

    def served(ini, d):  # shifts before someone's "On Service From" date never count towards fairness
        o = osf_map.get(ini)
        return o is None or d >= o

    try:
        xls = pd.ExcelFile(uploaded_file)
        sheets = [s for s in ["Previous Period Rota", "Rota"] if s in xls.sheet_names]
        if not sheets:
            for sn in xls.sheet_names:
                if sn.strip().upper() in NON_ROTA_SHEETS:
                    continue
                df_test = pd.read_excel(xls, sheet_name=sn, nrows=5)
                if any("DATE" in str(c).strip().upper() for c in df_test.columns):
                    sheets = [sn]
                    break
            if not sheets:
                sheets = [xls.sheet_names[0]]

        records = {}
        for sn in sheets:
            df_rota = pd.read_excel(xls, sheet_name=sn)
            if df_rota.empty:
                continue
            date_col, col_a, col_b, col_n = _detect_rota_cols(df_rota, a_name, b_name)
            for _, row in df_rota.iterrows():
                try:
                    r_date = parse_any_date(row[date_col])
                except Exception:
                    r_date = None
                if r_date is None:
                    continue
                g = lambda col: row.get(col) if col is not None else None
                records[r_date] = {
                    "a": match_name(g(col_a), consultants), "a_raw": safe_raw(g(col_a)),
                    "b": match_name(g(col_b), consultants), "b_raw": safe_raw(g(col_b)),
                    "n": match_name(g(col_n), consultants), "n_raw": safe_raw(g(col_n)),
                    "leave": {ini: safe_raw(row.get(ini)) for ini in active_initials},
                }

        master_a_seq_temp, memory_leave, clean_rows = [], {}, []
        for r_date in sorted(records):
            rec = records[r_date]
            a_ini, b_ini, n_ini = rec["a"], rec["b"], rec["n"]
            is_bh = uk_holidays is not None and r_date in uk_holidays and r_date.weekday() < 5
            bh_grp = get_bh_group(r_date) if is_bh else None

            if r_date < sel_s:
                if r_date.weekday() == 0 and b_ini:
                    mem["last_b_week"][b_ini] = r_date
                if ytd_s <= r_date <= ytd_e:
                    if a_ini and served(a_ini, r_date): counts[a_ini]["A_days"] += 1
                    if b_ini and served(b_ini, r_date): counts[b_ini]["B_days"] += 1
                    if n_ini and served(n_ini, r_date): counts[n_ini]["Nights"] += 1
                    if is_bh:
                        for ini in (a_ini, b_ini):
                            if ini and served(ini, r_date):
                                counts[ini]["BH_days"] += 1
                                if bh_grp: counts[ini]["BH_groups"].add(bh_grp)
                if a_ini:
                    mem["cycle_a"][r_date] = a_ini
                    mem["last_a_dates"][a_ini] = max(mem["last_a_dates"].get(a_ini, datetime.min.date()), r_date)
                    if not master_a_seq_temp or master_a_seq_temp[-1] != a_ini:
                        master_a_seq_temp.append(a_ini)
                if b_ini: mem["hist_b_dates"][b_ini] = max(mem["hist_b_dates"].get(b_ini, datetime.min.date()), r_date)
                if n_ini: mem["hist_b_dates"][n_ini] = max(mem["hist_b_dates"].get(n_ini, datetime.min.date()), r_date)
                clean_rows.append({
                    "Date": r_date.strftime("%d/%m/%Y"), "Day": DAY_NAMES[r_date.weekday()],
                    a_name: name_map.get(a_ini, rec["a_raw"]), b_name: name_map.get(b_ini, rec["b_raw"]),
                    "Nights": name_map.get(n_ini, rec["n_raw"]), **rec["leave"]})
            memory_leave[pd.Timestamp(r_date)] = rec["leave"]

        # Who was on A / B immediately before the new period (robust to re-running a period)
        day_before = sel_s - timedelta(days=1)
        if day_before in records and records[day_before]["a"]:
            mem["curr_a"] = records[day_before]["a"]
        prior_wd = [d for d in records if d < sel_s and d.weekday() < 5 and records[d]["b"] and (sel_s - d).days <= 7]
        if prior_wd:
            mem["curr_b"] = records[max(prior_wd)]["b"]

        if clean_rows:
            df_temp = pd.DataFrame(clean_rows)
            mask = (df_temp[a_name] != '') | (df_temp[b_name] != '') | (df_temp['Nights'] != '')
            if mask.any():
                mem["prev_rota_df"] = df_temp.loc[:mask[::-1].idxmax()].copy()

        if "State" in xls.sheet_names:
            df_state = pd.read_excel(xls, sheet_name="State")
            if df_state.shape[1] >= 2:
                sd = dict(zip(df_state.iloc[:, 0].astype(str).str.strip(), df_state.iloc[:, 1]))
                mem["master_a_seq"] = [x.strip().upper() for x in str(sd.get("A_MASTER_SEQ", "")).split(",") if safe_ini(x)]
                if mem["curr_a"] is None and safe_ini(sd.get("CURRENT_A")): mem["curr_a"] = safe_ini(sd.get("CURRENT_A"))
                if mem["curr_b"] is None and safe_ini(sd.get("CURRENT_B")): mem["curr_b"] = safe_ini(sd.get("CURRENT_B"))
                mem["period_end"] = parse_any_date(sd.get("PERIOD_END"))

        if not mem["master_a_seq"] and master_a_seq_temp:
            seen = set()
            mem["master_a_seq"] = [x for x in reversed(master_a_seq_temp) if not (x in seen or seen.add(x))][::-1]

        mem["counts"] = dict(counts)
        mem["leave_df"] = pd.DataFrame.from_dict(memory_leave, orient='index') if memory_leave else pd.DataFrame()
    except Exception as e:
        mem["warnings"].append(f"Memory parsing warning: {e}")
    return mem


# =============================================================================
# CORE ENGINE (V1 final baseline, with multi-specialty & per-consultant settings)
# =============================================================================
def generate_rota(leave_df, consultants, start_date, end_date, hist_counts, hist_master_a, hist_cycle_a, hist_b_q,
                  hist_curr_a, hist_curr_b, hist_last_a_dates, locum_dict, hist_b_dates, ytd_s, ytd_e, hist_last_b_week,
                  a_anchor_wd=4, a_min_gap_days=21, b_bb_gap_days=21,
                  fte_tl_a=None, fte_tl_b=None, on_service=None, level_start=None,
                  scheduler_cfg=None):
    active_initials = [safe_ini(c.get('Initials')) for c in consultants if c.get('Active', True) and safe_ini(c.get('Initials'))]
    name_map = {safe_ini(c.get('Initials')): c.get('Full Name', '') for c in consultants if safe_ini(c.get('Initials'))}
    ini_by_name = {v: k for k, v in name_map.items()}

    scfg = scheduler_cfg or {}
    parallel_consultants = int(scfg.get("parallel_consultants", 2))
    a_type = scfg.get("a_type", "7-Day Continuous")
    b_type = scfg.get("b_type", "5-Day Weekday")
    b_start_wd = int(scfg.get("b_start_wd", 0))
    nights_mode = scfg.get("nights_mode", "Alternating Consultants")
    a_nights_covered = list(scfg.get("a_nights_covered", [1, 3, 4, 5, 6]))
    b_nights_covered = list(scfg.get("b_nights_covered", [0, 2]))
    flip_nights_on_bh = bool(scfg.get("flip_nights_on_bh", True))
    b_role = scfg.get("b_role", "Second Parallel Rota + Nights")
    group_pairing = scfg.get("group_pairing", "Different Group Ideally")
    allow_dual_cover = bool(scfg.get("allow_dual_cover", False))
    monday_off_after_weekend = bool(scfg.get("monday_off_after_weekend", False))
    b_active = bool(scfg.get("b_active", True))
    weekend_workers_by_date = scfg.get("weekend_workers_by_date") or {}
    a_weekdays_only = (a_type == "5-Day Weekday")
    b_weekend_only = (b_type == "2-Day Weekend")

    def safe_float(val, default=1.0):
        try:
            f = float(val)
            return default if pd.isna(f) else f
        except Exception:
            return default

    curr_d_start = pd.to_datetime(start_date).date()
    fin_d = pd.to_datetime(end_date).date()
    ytd_s_d = pd.to_datetime(ytd_s).date()
    ytd_e_d = pd.to_datetime(ytd_e).date()
    ideal_gap = max(28, a_min_gap_days)

    d_map, fte_a_map, fte_b_map, group_map = {}, {}, {}, {}
    weekdays_map, nights_pref_map, pref_weeks_map, rest_on_call_map, rest_days_map = {}, {}, {}, {}, {}
    for c in consultants:
        ini = safe_ini(c.get('Initials'))
        if c.get('Active', True) and ini:
            raw_dil = safe_raw(c.get('DIL Day'))
            d_map[ini] = raw_dil.capitalize()
            raw_fa = safe_float(c.get('FTE A', 1.0))
            raw_fb = safe_float(c.get('FTE B', 1.0))
            fte_a_map[ini] = 0.0 if raw_fa <= 0.0 else max(0.1, raw_fa)
            fte_b_map[ini] = 0.0 if raw_fb <= 0.0 else max(0.1, raw_fb)
            group_map[ini] = (safe_raw(c.get('Group')) or 'Others').capitalize()
            weekdays_map[ini] = safe_bool(c.get('Weekdays', True), True)
            npref = safe_raw(c.get('Nights'))
            nights_pref_map[ini] = npref if npref in STAFF_NIGHTS_OPTIONS else 'Yes'
            pref_weeks_map[ini] = max(1, min(4, int(safe_float(c.get('Preferred Weeks in a Row', 1), 1))))
            rest_on_call_map[ini] = safe_bool(c.get('Rest Days on Call', False), False)
            rest_days_map[ini] = set(parse_weekdays(c.get('Rest Days (Mon-Fri)', ''), [])) if rest_on_call_map[ini] else set()

    has_multi_week_pref = any(pref_weeks_map.get(i, 1) > 1 for i in active_initials)
    has_custom_nights_pref = any(nights_pref_map.get(i, "Yes") != "Yes" for i in active_initials)

    on_service = on_service or {}
    fte_tl_a, fte_tl_b = fte_tl_a or {}, fte_tl_b or {}

    def on_serv(ini, d_date):
        o = on_service.get(ini)
        return o is None or d_date >= o

    def fte_at(tl, ini, d_date, default):
        lst = tl.get(ini)
        if not lst:
            return default
        val = lst[0][1]
        for eff, f in lst:
            if eff <= d_date:
                val = f
            else:
                break
        raw_v = safe_float(val, default)
        return 0.0 if raw_v <= 0.0 else max(0.1, raw_v)

    fte_a_at = lambda ini, d_date: fte_at(fte_tl_a, ini, d_date, fte_a_map.get(ini, 1.0))
    fte_b_at = lambda ini, d_date: fte_at(fte_tl_b, ini, d_date, fte_b_map.get(ini, 1.0))

    def can_a(ini, d_date):
        return on_serv(ini, d_date) and fte_a_at(ini, d_date) > 0.0 and (not a_weekdays_only or weekdays_map.get(ini, True))

    def can_b(ini, d_date):
        if b_weekend_only:
            return on_serv(ini, d_date) and fte_b_at(ini, d_date) > 0.0
        if b_role == "Just Nights":
            return on_serv(ini, d_date) and fte_b_at(ini, d_date) > 0.0 and nights_pref_map.get(ini, "Yes") != "No"
        return on_serv(ini, d_date) and fte_b_at(ini, d_date) > 0.0 and weekdays_map.get(ini, True)

    def get_eff_marker(d_date, ini):
        ts = pd.Timestamp(d_date).normalize()
        if not leave_df.empty and ts in leave_df.index and ini in leave_df.columns:
            v = leave_df.loc[ts, ini]
            if isinstance(v, pd.Series):
                v = v.iloc[0]
            val = safe_ini(v)
            if val:
                return val
        if d_map.get(ini) == DAY_NAMES[d_date.weekday()]:
            return "DIL"
        if rest_on_call_map.get(ini) and d_date.weekday() in rest_days_map.get(ini, set()):
            return "DIL"
        return ""

    if hist_master_a:
        master_a_seq = [i for i in hist_master_a if i in active_initials] + [i for i in active_initials if i not in hist_master_a]
    else:
        solids = [i for i in active_initials if group_map.get(i) == "Solid"]
        neuros = [i for i in active_initials if group_map.get(i) == "Neuro"]
        others = [i for i in active_initials if group_map.get(i) not in ["Solid", "Neuro"]]
        random.seed(); random.shuffle(solids); random.shuffle(neuros); random.shuffle(others)
        master_a_seq = []
        while solids or neuros or others:
            if solids: master_a_seq.append(solids.pop(0))
            if neuros: master_a_seq.append(neuros.pop(0))
            if others: master_a_seq.append(others.pop(0))

    start_anc = curr_d_start - timedelta(days=(curr_d_start.weekday() - a_anchor_wd) % 7)
    cycle_a_schedule = {k: v for k, v in hist_cycle_a.items() if k < start_anc}
    last_a_end = {k: v for k, v in hist_last_a_dates.items() if v < start_anc}

    # Mid-block start: the A already on duty keeps the block that is in progress
    carry_a = hist_curr_a if (hist_curr_a in active_initials and start_anc < curr_d_start) else None

    c_fs = []
    chk = start_anc
    while chk <= fin_d:
        if not any(locum_dict.get(chk + timedelta(days=o)) for o in range(7)) and not (carry_a and chk == start_anc):
            c_fs.append(chk)
        chk += timedelta(days=7)
    if carry_a:
        cycle_a_schedule[start_anc] = carry_a

    # REQ handling: a request covers EXACTLY the days marked REQ (the day-level override in the daily loop).
    # If the person holds REQ on the anchor day, they are booked into that block for rest-gap/fairness purposes,
    # but only the requested run of days is theirs. Unless every remaining day of the block is covered by other
    # REQs, the block stays in the solver so someone else is added in for the rest of it (until the last day).
    temp = {f: {'doc': None, 'locked': False, 'run': 0} for f in c_fs}

    def req_on(d_date, doc_ini):
        return get_eff_marker(d_date, doc_ini) == "REQ"

    for f in c_fs:
        for doc in active_initials:
            if req_on(f, doc):
                run = 0
                while run < 7 and req_on(f + timedelta(days=run), doc):
                    run += 1
                rest_covered = all(any(req_on(f + timedelta(days=o), d) for d in active_initials) for o in range(run, 7))
                temp[f]['doc'], temp[f]['run'] = doc, run
                temp[f]['locked'] = run >= 7 or rest_covered
                break

    # --- FAIRNESS BASELINE ------------------------------------------------------------------------
    # Standing = shifts worked minus the fair share for the time actually on service, at the FTE of the time.
    # Time before "On Service From" (e.g. long sick leave), and days marked SICK, have zero exposure, so a returner starts level with
    # colleagues: nothing to catch up, no head start. FTE changes count only from their effective date.
    level_start = set(level_start or [])
    hist_end = min(curr_d_start - timedelta(days=1), ytd_e_d)

    def exposure(ini, fte_fn, blocked_markers, weekdays_only, require_weekdays_toggle=False):
        if ini in level_start:
            return 0.0
        if require_weekdays_toggle and not weekdays_map.get(ini, True):
            return 0.0
        total, d = 0.0, max(ytd_s_d, on_service.get(ini) or ytd_s_d)
        while d <= hist_end:
            counts_today = (not weekdays_only) or (d.weekday() in B_WORK_DAYS and not (uk_holidays is not None and d in uk_holidays))
            if counts_today and get_eff_marker(d, ini) not in blocked_markers:
                total += fte_fn(ini, d)
            d += timedelta(days=1)
        return total

    exp_a = {i: exposure(i, fte_a_at, NO_EXPOSURE_A, a_weekdays_only, require_weekdays_toggle=a_weekdays_only) for i in active_initials}
    exp_b_hist = {i: exposure(i, fte_b_at, NO_EXPOSURE_B, not b_weekend_only, require_weekdays_toggle=(not b_weekend_only and b_role != "Just Nights")) for i in active_initials}
    hist_a = {i: hist_counts.get(i, {}).get("A_days", 0) for i in active_initials}
    hist_b = {i: hist_counts.get(i, {}).get("B_days", 0) for i in active_initials}
    rate_a = sum(hist_a.values()) / sum(exp_a.values()) if sum(exp_a.values()) > 0 else 0.0
    rate_b = sum(hist_b.values()) / sum(exp_b_hist.values()) if sum(exp_b_hist.values()) > 0 else 0.0
    surplus_a = {i: 0.0 if i in level_start else hist_a[i] - rate_a * exp_a[i] for i in active_initials}
    surplus_b = {i: 0.0 if i in level_start else hist_b[i] - rate_b * exp_b_hist[i] for i in active_initials}
    # Within the new period everyone is expected to do a share proportional to the FTE-days they are on service
    # (sick days excluded). Ordinary leave (AL/SL/NO/OFF) still counts, so it has to be caught up; a late
    # starter is not "owed" the weeks before they arrived.
    exp_a_new = {i: 0.0 for i in active_initials}
    d = curr_d_start
    while d <= fin_d:
        for i in active_initials:
            if can_a(i, d) and get_eff_marker(d, i) not in NO_EXPOSURE_A:
                exp_a_new[i] += fte_a_at(i, d)
        d += timedelta(days=1)
    rate_a_new = ((fin_d - curr_d_start).days + 1) / sum(exp_a_new.values()) if sum(exp_a_new.values()) > 0 else 0.0
    # A standing is kept "per FTE": a 7-day block costs 7 / (FTE at that time)
    sim_a_units = {
        i: 0.0 if fte_a_at(i, curr_d_start) <= 0.0 else (surplus_a[i] - rate_a_new * exp_a_new[i]) / max(0.1, fte_a_at(i, curr_d_start))
        for i in active_initials
    }

    sim_bh_groups = {ini: set(hist_counts.get(ini, {}).get("BH_groups", set())) for ini in active_initials}

    def week_has_bh(slot_f):
        return any(uk_holidays is not None and (slot_f + timedelta(days=i)) in uk_holidays and (slot_f + timedelta(days=i)).weekday() < 5 for i in range(7))

    def get_week_bh_group(slot_f):
        for i in range(7):
            d = slot_f + timedelta(days=i)
            if uk_holidays is not None and d in uk_holidays and d.weekday() < 5:
                return get_bh_group(d)
        return None

    # --- A-WEEK: DFS LEAVE-SOFTENED ENGINE ---
    def week_start_for_a_gap(d_date):
        return d_date - timedelta(days=(d_date.weekday() - a_anchor_wd) % 7)

    def count_leave_in_week(anchor_f, cand):
        return sum(1 for o in range(7) if get_eff_marker(anchor_f + timedelta(days=o), cand) in LEAVE_BLOCK)

    def count_catchup_leave_in_week(anchor_f, cand):
        return sum(1 for o in range(7) if get_eff_marker(anchor_f + timedelta(days=o), cand) in CATCHUP_LEAVE)

    leave_blocks = {ini: [] for ini in active_initials}
    for ini in active_initials:
        current_block_start, block_length = None, 0
        for f in c_fs:
            if count_catchup_leave_in_week(f, ini) >= 3:
                if current_block_start is None:
                    current_block_start = f
                block_length += 1
            else:
                if block_length >= 2:
                    leave_blocks[ini].append((current_block_start, f))
                current_block_start, block_length = None, 0

    week_base_candidates, week_block_counts = {}, {}
    for f in c_fs:
        blocked = sum(1 for cand in active_initials if count_leave_in_week(f, cand) >= 3)
        week_base_candidates[f] = [c for c in active_initials if can_a(c, f)]  # leave is soft-filtered via leave_in_week_pen
        week_block_counts[f] = blocked

    prior_a_week = {ini: week_start_for_a_gap(last_date) for ini, last_date in hist_last_a_dates.items() if ini in active_initials and last_date}
    master_order = {ini: idx for idx, ini in enumerate(master_a_seq)}
    last_b_anchor = {ini: week_start_for_a_gap(dt) for ini, dt in hist_last_b_week.items() if ini in active_initials and dt}

    def _consec_run_len_before(cand, f, weeks_list):
        wset = set(weeks_list)
        cnt = 0
        cur_w = f - timedelta(days=7)
        while cur_w in wset:
            cnt += 1
            cur_w -= timedelta(days=7)
        return cnt

    def _gap_ok_for_cand(cand, f, weeks_list, min_gap_days):
        if not has_multi_week_pref or pref_weeks_map.get(cand, 1) <= 1:
            return all(abs((f - prev_f).days) >= min_gap_days for prev_f in weeks_list)
        pw = pref_weeks_map.get(cand, 1)
        wset = set(weeks_list)
        if (f - timedelta(days=7)) in wset and _consec_run_len_before(cand, f, weeks_list) < pw and count_leave_in_week(f, cand) < 3:
            # Continuing an in-progress multi-week block: only check min_gap_days against weeks outside this run
            run_len = _consec_run_len_before(cand, f, weeks_list)
            run_weeks = {f - timedelta(days=7 * k) for k in range(1, run_len + 1)}
            return all(abs((f - prev_f).days) >= min_gap_days for prev_f in weeks_list if prev_f not in run_weeks)
        return all(abs((f - prev_f).days) >= min_gap_days for prev_f in weeks_list)

    def solve_a_weeks(min_gap_days):
        assigned = {}
        assigned_by_doc = {ini: [] for ini in active_initials}
        local_a_units = sim_a_units.copy()
        local_bh_groups = {ini: set(groups) for ini, groups in sim_bh_groups.items()}
        for ini, wk in prior_a_week.items():
            assigned_by_doc.setdefault(ini, []).append(wk)

        for f in c_fs:
            req_doc = temp[f]['doc']
            if req_doc is None:
                continue
            if req_doc not in active_initials:
                if temp[f]['locked']:
                    assigned[f] = req_doc
                continue
            # Requests are honoured unconditionally (a clash must never wipe the whole solve);
            # they are booked so rest gaps and fairness around them stay correct.
            assigned_by_doc.setdefault(req_doc, []).append(f)
            local_a_units[req_doc] += max(1, temp[f]['run']) / max(0.1, fte_a_at(req_doc, f))
            if week_has_bh(f):
                bg = get_week_bh_group(f)
                if bg: local_bh_groups[req_doc].add(bg)
            if temp[f]['locked']:
                assigned[f] = req_doc

        unassigned = {f for f in c_fs if f not in assigned}

        def current_candidates(f):
            return [cand for cand in week_base_candidates.get(f, []) if _gap_ok_for_cand(cand, f, assigned_by_doc.get(cand, []), min_gap_days)]

        def choose_next_week():
            if has_multi_week_pref:
                # When multi-week blocks are active, assign weeks in chronological order so consecutive runs build naturally
                for f in sorted(unassigned):
                    return f, current_candidates(f)
            best_week, best_key, best_cands = None, None, None
            for f in unassigned:
                cands = current_candidates(f)
                key = (len(cands), -week_block_counts.get(f, 0), 0 if week_has_bh(f) else 1, f)
                if best_key is None or key < best_key:
                    best_week, best_key, best_cands = f, key, cands
            return best_week, best_cands

        def candidate_score(cand, f):
            prev_weeks = assigned_by_doc.get(cand, [])
            closest_gap = min([abs((f - prev_f).days) for prev_f in prev_weeks], default=9999)
            if has_multi_week_pref and pref_weeks_map.get(cand, 1) > 1 and (f - timedelta(days=7)) in prev_weeks and _consec_run_len_before(cand, f, prev_weeks) < pref_weeks_map.get(cand, 1) and count_leave_in_week(f, cand) < 3:
                gap_tier = -1
            else:
                gap_tier = 0 if closest_gap >= ideal_gap else (1 if closest_gap >= a_min_gap_days else 2)

            leave_in_week_pen = count_leave_in_week(f, cand)

            future_pressure = 0
            for other_f in unassigned:
                if other_f == f:
                    continue
                remaining = sum(1 for other_cand in week_base_candidates.get(other_f, []) if other_cand != cand and _gap_ok_for_cand(other_cand, other_f, assigned_by_doc.get(other_cand, []), min_gap_days))
                if remaining == 0: future_pressure += 1000
                elif remaining == 1: future_pressure += 10

            return_skip_pen, leave_return_tier = 0, 1
            for start_f, return_f in leave_blocks[cand]:
                if return_f <= f:
                    last_a = max([w for w in assigned_by_doc.get(cand, []) if w < start_f], default=datetime.min.date())
                    if (start_f - last_a).days >= 21:
                        if f == return_f:
                            leave_return_tier = 0
                        elif f > return_f and not any(return_f <= w < f for w in assigned_by_doc.get(cand, [])) and return_f in unassigned:
                            return_skip_pen = 1

            last_b_gap = (f - last_b_anchor[cand]).days if cand in last_b_anchor else 9999
            sacrifice_bonus = 0 if gap_tier > 0 and 0 <= last_b_gap <= 14 else 1

            bh_pen = len(local_bh_groups[cand]) if week_has_bh(f) else 0
            fairness_over_target_a = local_a_units[cand]

            last_a_overall = max([w for w in prev_weeks if w < f], default=datetime.min.date()) if prev_weeks else datetime.min.date()
            absolute_rest_days = -(f - last_a_overall).days if last_a_overall != datetime.min.date() else -9999

            # MASTER A-TUPLE
            return (gap_tier, leave_in_week_pen, future_pressure, return_skip_pen, leave_return_tier, sacrifice_bonus,
                    bh_pen, math.floor(fairness_over_target_a), absolute_rest_days, master_order.get(cand, 9999))

        def dfs():
            if not unassigned:
                return True
            f, cands = choose_next_week()
            if f is None:
                return True
            if not cands:
                return False
            ordered = sorted(cands, key=lambda cand: candidate_score(cand, f))
            unassigned.remove(f)
            for cand in ordered:
                assigned[f] = cand
                assigned_by_doc.setdefault(cand, []).append(f)
                local_a_units[cand] += 7 / max(0.1, fte_a_at(cand, f))
                added_bh = get_week_bh_group(f) if week_has_bh(f) else None
                if added_bh:
                    local_bh_groups[cand].add(added_bh)

                if dfs():
                    return True

                if added_bh and not any(other_f != f and assigned.get(other_f) == cand and get_week_bh_group(other_f) == added_bh for other_f in assigned) and added_bh not in hist_counts.get(cand, {}).get("BH_groups", set()):
                    local_bh_groups[cand].discard(added_bh)
                local_a_units[cand] -= 7 / max(0.1, fte_a_at(cand, f))
                assigned_by_doc[cand].pop()
                assigned.pop(f, None)
            unassigned.add(f)
            return False

        if dfs():
            return assigned
        return None

    solved_a = solve_a_weeks(a_min_gap_days)
    if solved_a is None and a_min_gap_days > 14:
        solved_a = solve_a_weeks(14)
    if solved_a is None:
        solved_a = solve_a_weeks(7)

    for f in c_fs:
        if temp[f]['locked']:
            cycle_a_schedule[f] = temp[f]['doc'] if temp[f]['doc'] in active_initials else "UNCOVERED"
        elif solved_a and f in solved_a:
            cycle_a_schedule[f] = solved_a[f]
        else:
            cycle_a_schedule[f] = "UNCOVERED"

    rota_data = []
    curr_d = curr_d_start
    cur_a = carry_a if carry_a else cycle_a_schedule.get(start_anc, "UNCOVERED")
    cur_b = hist_curr_b if hist_curr_b in active_initials else None
    final_cur_a, final_cur_b = cur_a, cur_b
    b_consec_weeks = 1

    counts = {ini: {"A_days": hist_counts.get(ini, {}).get("A_days", 0), "B_days": hist_counts.get(ini, {}).get("B_days", 0),
                    "Nights": hist_counts.get(ini, {}).get("Nights", 0), "BH_days": hist_counts.get(ini, {}).get("BH_days", 0),
                    "BH_groups": set(hist_counts.get(ini, {}).get("BH_groups", set()))} for ini in active_initials}
    period_counts = {ini: {"A_days": 0, "B_days": 0, "Nights": 0, "BH_days": 0} for ini in active_initials}

    recent_b_shifts, last_b_week_start = dict(hist_b_dates), dict(hist_last_b_week)
    recent_n_shifts = {}
    sim_b_counts = {ini: 0 for ini in active_initials}  # B shifts given in THIS run (history is in surplus_b)
    sim_n_counts = {ini: 0 for ini in active_initials}

    def add_score(doc_ini, s_type, d_date):
        if doc_ini and doc_ini in period_counts:
            if curr_d_start <= d_date <= fin_d: period_counts[doc_ini][s_type] += 1
            if ytd_s_d <= d_date <= ytd_e_d: counts[doc_ini][s_type] += 1
            if s_type == "B_days": sim_b_counts[doc_ini] += 1
            if s_type == "Nights": sim_n_counts[doc_ini] += 1

    # --- B-SHIFT CAPACITY AND FAIRNESS TARGET SETUP ---
    weighted_cap = {ini: 0.0 for ini in active_initials}  # available days on service, weighted by FTE-at-the-time
    test_d = curr_d_start
    while test_d <= fin_d:
        wd = test_d.weekday()
        b_day_ok = (wd >= 5) if b_weekend_only else (wd in B_WORK_DAYS and not (uk_holidays is not None and test_d in uk_holidays))
        if b_day_ok:
            for ini in active_initials:
                if can_b(ini, test_d) and get_eff_marker(test_d, ini) not in NO_EXPOSURE_B:
                    weighted_cap[ini] += fte_b_at(ini, test_d)
        test_d += timedelta(days=1)

    total_weighted_cap = sum(weighted_cap.values())
    b_target_share = {ini: (weighted_cap[ini] / total_weighted_cap if total_weighted_cap > 0 else 1.0 / max(1, len(active_initials))) for ini in active_initials}

    exp_n_hist = {i: (0.0 if nights_pref_map.get(i, "Yes") == "No" else exp_a[i]) for i in active_initials}
    exp_n_new = {i: (0.0 if nights_pref_map.get(i, "Yes") == "No" else exp_a_new[i]) for i in active_initials}
    hist_n = {i: hist_counts.get(i, {}).get("Nights", 0) for i in active_initials}
    rate_n = sum(hist_n.values()) / sum(exp_n_hist.values()) if sum(exp_n_hist.values()) > 0 else 0.0
    surplus_n = {i: 0.0 if i in level_start or nights_pref_map.get(i, "Yes") == "No" else hist_n[i] - rate_n * exp_n_hist[i] for i in active_initials}
    n_target_share = {i: (exp_n_new[i] / sum(exp_n_new.values()) if sum(exp_n_new.values()) > 0 else 0.0) for i in active_initials}

    def get_fairness_over_target(c):
        return surplus_b[c] + sim_b_counts[c] - sum(sim_b_counts.values()) * b_target_share[c]

    def get_night_over_target(c):
        return surplus_n[c] + sim_n_counts[c] - sum(sim_n_counts.values()) * n_target_share[c]

    def group_penalty(a_ini, b_ini):
        if group_pairing == "Same Group Allowed" or not a_ini or not b_ini:
            return 0
        g_a, g_b = group_map.get(a_ini, "Others"), group_map.get(b_ini, "Others")
        return 1 if g_a != "Others" and g_a == g_b else 0

    def is_monday_off_after_wknd(ini, d_date):
        if not monday_off_after_weekend or d_date.weekday() != 0:
            return False
        sat_d, sun_d = d_date - timedelta(days=2), d_date - timedelta(days=1)
        return ini in weekend_workers_by_date.get(sat_d, set()) or ini in weekend_workers_by_date.get(sun_d, set())

    while curr_d <= fin_d + timedelta(days=365):
        wd = curr_d.weekday()
        row = {"Date": curr_d.strftime("%d/%m/%Y"), "Day": DAY_NAMES[wd], "Consultant A": "", "Consultant B": "", "Nights": ""}

        if curr_d <= fin_d:
            is_bh_today = uk_holidays is not None and curr_d in uk_holidays and wd < 5

            w_start = curr_d - timedelta(days=wd)
            is_bh_week = uk_holidays is not None and any((w_start + timedelta(days=d)) in uk_holidays for d in range(5))
            if nights_mode == "None":
                active_a_nights, active_b_nights = [], []
            elif nights_mode == "With Primary / Weekday" or (parallel_consultants == 1 and b_role != "Just Nights"):
                active_b_nights = list(b_nights_covered) if (b_active and b_role != "Just Nights" and b_weekend_only) else []
                active_a_nights = [d for d in a_nights_covered if d not in active_b_nights]
            else:
                b_clean = list(b_nights_covered) if b_active else []
                a_clean = [d for d in a_nights_covered if d not in b_clean]
                if flip_nights_on_bh and is_bh_week and a_clean and b_clean:
                    if a_clean == [1, 3, 4, 5, 6] and b_clean == [0, 2]:
                        active_a_nights, active_b_nights = [0, 2, 4, 5, 6], [1, 3]
                    else:
                        active_a_nights = sorted([d for d in b_clean if d < 4] + [d for d in a_clean if d >= 4])
                        active_b_nights = sorted([d for d in a_clean if d < 4] + [d for d in b_clean if d >= 4])
                else:
                    active_a_nights, active_b_nights = a_clean, b_clean

            if wd == a_anchor_wd:
                anc = curr_d
                cur_a = "LOCUM_ACADEMIC" if any(locum_dict.get(anc + timedelta(days=o)) for o in range(7)) else cycle_a_schedule.get(anc, "UNCOVERED")

            blocked_count = sum(1 for c in active_initials if sum(1 for i in range(5) if get_eff_marker(w_start + timedelta(days=i), c) in LEAVE_BLOCK) >= 3)
            high_leave_mode = blocked_count >= 4

            if b_active and not b_weekend_only and b_role != "Just Nights" and (wd == b_start_wd or cur_b is None or cur_b == "UNCOVERED"):
                # Multi-week continuation check when Preferred Weeks in a Row > 1
                if (wd == b_start_wd and cur_b in active_initials and b_consec_weeks < pref_weeks_map.get(cur_b, 1)
                        and (allow_dual_cover or cur_b != cur_a) and can_b(cur_b, curr_d)
                        and sum(1 for i in range(5) if get_eff_marker(w_start + timedelta(days=i), cur_b) in LEAVE_BLOCK) < 3):
                    b_consec_weeks += 1
                    last_b_week_start[cur_b] = curr_d
                else:
                    avail_week = [c for c in active_initials if (allow_dual_cover or c != cur_a) and can_b(c, curr_d)
                                  and sum(1 for i in range(5) if get_eff_marker(w_start + timedelta(days=i), c) in LEAVE_BLOCK) < 3]
                    if not avail_week:
                        avail_week = [c for c in active_initials if (allow_dual_cover or c != cur_a) and can_b(c, curr_d) and get_eff_marker(curr_d, c) not in B_DAY_BLOCK]
                    if avail_week:
                        def week_score(x):
                            fot = get_fairness_over_target(x)
                            is_runaway = 1 if fot >= 3.0 else 0
                            monday_guard = 1 if fot >= 1.5 else 0
                            bb_pen = 1 if x in last_b_week_start and (curr_d - last_b_week_start[x]).days < b_bb_gap_days else 0
                            leave_in_week = sum(1 for i in range(5) if get_eff_marker(curr_d + timedelta(days=i), x) in LEAVE_BLOCK)
                            g_pen = group_penalty(cur_a, x)
                            past_a = [f for f, cand in cycle_a_schedule.items() if cand == x and f <= curr_d]
                            b_after_a_pen = 1 if past_a and (curr_d - max(past_a)).days <= 14 else 0
                            if high_leave_mode: b_after_a_pen = 0
                            dual_pen = 1 if x == cur_a else 0
                            # MASTER B-WEEK TUPLE
                            if allow_dual_cover:
                                return (dual_pen, is_runaway, monday_guard, bb_pen, leave_in_week, int(round(fot)), g_pen, b_after_a_pen, fot)
                            return (is_runaway, monday_guard, bb_pen, leave_in_week, int(round(fot)), g_pen, b_after_a_pen, fot)
                        cur_b = min(avail_week, key=week_score)
                        last_b_week_start[cur_b] = curr_d
                        b_consec_weeks = 1
                    else:
                        cur_b = "UNCOVERED"
                        b_consec_weeks = 1

            rq_d = next((i for i in active_initials if get_eff_marker(curr_d, i) == "REQ"), None)
            rqb_d = next((i for i in active_initials if get_eff_marker(curr_d, i) == "REQB"), None)
            lc_d = locum_dict.get(curr_d)

            if a_weekdays_only and (wd >= 5 or is_bh_today):
                da_n, da_i = "", None
            else:
                da_n = lc_d if lc_d else (name_map.get(rq_d, rq_d) if rq_d else name_map.get(cur_a, cur_a))
                da_i = (cur_a if cur_a not in ["UNCOVERED", "LOCUM_ACADEMIC"] else None) if not rq_d and not lc_d else rq_d
                if da_i and is_monday_off_after_wknd(da_i, curr_d):
                    rep_a = [i for i in active_initials if i != da_i and can_a(i, curr_d) and not is_monday_off_after_wknd(i, curr_d) and get_eff_marker(curr_d, i) not in B_DAY_BLOCK]
                    if rep_a:
                        da_i = min(rep_a, key=lambda x: (sim_a_units.get(x, 0.0), master_order.get(x, 9999)))
                        da_n = name_map.get(da_i, da_i)

            row["Consultant A"] = da_n
            if da_i:
                add_score(da_i, "A_days", curr_d)
                if wd >= 5:
                    weekend_workers_by_date.setdefault(curr_d, set()).add(da_i)
            if wd in active_a_nights and da_n:
                row["Nights"] = da_n
                if da_i: add_score(da_i, "Nights", curr_d)

            if b_active and b_role == "Just Nights":
                # Renal / Nights-only mode: daytime Consultant B stays blank; B engine only fills B's night shifts
                row["Consultant B"] = ""
                if wd in active_b_nights and not is_bh_today:
                    if rqb_d and (allow_dual_cover or rqb_d != da_i):
                        row["Nights"] = name_map.get(rqb_d, rqb_d)
                        add_score(rqb_d, "Nights", curr_d)
                        recent_b_shifts[rqb_d] = curr_d
                    else:
                        avail_n = [i for i in active_initials if (allow_dual_cover or i != da_i) and can_b(i, curr_d) and get_eff_marker(curr_d, i) not in B_DAY_BLOCK]
                        if avail_n:
                            def just_nights_score(x):
                                not_over = get_night_over_target(x)
                                bb_pen = 1 if x in recent_b_shifts and (curr_d - recent_b_shifts[x]).days < 7 else 0
                                g_pen = group_penalty(da_i, x)
                                past_a = [f for f, cand in cycle_a_schedule.items() if cand == x and f <= curr_d]
                                b_after_a_pen = 1 if past_a and (curr_d - max(past_a)).days <= 14 else 0
                                dual_pen = 1 if x == da_i else 0
                                return (dual_pen, bb_pen, int(round(not_over)), g_pen, b_after_a_pen, not_over)
                            best_n = min(avail_n, key=just_nights_score)
                            row["Nights"] = name_map.get(best_n, "UNCOVERED")
                            add_score(best_n, "Nights", curr_d)
                            recent_b_shifts[best_n] = curr_d
                        else:
                            row["Nights"] = "UNCOVERED"
            elif b_active and b_weekend_only:
                # 2-Day Weekend shift in Slot B (Saturday & Sunday)
                if wd >= 5:
                    if wd == 5 or not cur_b or cur_b == "UNCOVERED" or (not allow_dual_cover and cur_b == da_i) or get_eff_marker(curr_d, cur_b) in B_DAY_BLOCK:
                        span = 2 if wd == 5 else 1
                        avail_wknd = [
                            c for c in active_initials
                            if (allow_dual_cover or c != da_i) and can_b(c, curr_d)
                            and all(get_eff_marker(curr_d + timedelta(days=o), c) not in B_DAY_BLOCK for o in range(span))
                        ]
                        if not avail_wknd:
                            avail_wknd = [c for c in active_initials if (allow_dual_cover or c != da_i) and can_b(c, curr_d) and get_eff_marker(curr_d, c) not in B_DAY_BLOCK]
                        if avail_wknd:
                            def wknd_score(x):
                                fot = get_fairness_over_target(x)
                                bb_pen = 1 if x in last_b_week_start and (curr_d - last_b_week_start[x]).days < b_bb_gap_days else 0
                                g_pen = group_penalty(da_i, x)
                                dual_pen = 1 if x == da_i else 0
                                return (dual_pen, bb_pen, int(round(fot)), g_pen, fot)
                            cur_b = min(avail_wknd, key=wknd_score)
                            if wd == 5:
                                last_b_week_start[cur_b] = curr_d
                        else:
                            cur_b = "UNCOVERED"
                    row["Consultant B"] = name_map.get(cur_b, cur_b) if cur_b != "UNCOVERED" else "UNCOVERED"
                    if cur_b and cur_b != "UNCOVERED":
                        add_score(cur_b, "B_days", curr_d)
                        recent_b_shifts[cur_b] = curr_d
                        weekend_workers_by_date.setdefault(curr_d, set()).add(cur_b)
                        if wd in active_b_nights:
                            row["Nights"] = row["Consultant B"]
                            add_score(cur_b, "Nights", curr_d)
            elif b_active and wd in B_WORK_DAYS:
                if is_bh_today:
                    row["Consultant B"] = ""
                elif rqb_d and (allow_dual_cover or rqb_d != da_i):  # explicit request wins over the automatic pick
                    row["Consultant B"] = name_map.get(rqb_d, rqb_d)
                    add_score(rqb_d, "B_days", curr_d)
                    recent_b_shifts[rqb_d] = curr_d
                    if wd in active_b_nights:
                        row["Nights"] = row["Consultant B"]
                        add_score(rqb_d, "Nights", curr_d)
                else:
                    if not cur_b or get_eff_marker(curr_d, cur_b) in B_DAY_BLOCK or (not allow_dual_cover and cur_b == da_i) or is_monday_off_after_wknd(cur_b, curr_d):
                        # HARD FILTER: Cannot be B if A (unless dual cover allowed), or if marked with leave/Monday-post-weekend rest today
                        avail = [i for i in active_initials if (allow_dual_cover or i != da_i) and can_b(i, curr_d) and not is_monday_off_after_wknd(i, curr_d) and get_eff_marker(curr_d, i) not in B_DAY_BLOCK]
                        if avail:
                            def daily_score(x):
                                fot = get_fairness_over_target(x)
                                not_over = get_night_over_target(x)
                                is_runaway = 1 if fot >= 3.0 else 0
                                bb_pen = 1 if x in recent_b_shifts and (curr_d - recent_b_shifts[x]).days < 7 else 0
                                g_pen = group_penalty(da_i, x)
                                past_a = [f for f, cand in cycle_a_schedule.items() if cand == x and f <= curr_d]
                                b_after_a_pen = 1 if past_a and (curr_d - max(past_a)).days <= 14 else 0
                                if high_leave_mode: b_after_a_pen = 0
                                dual_pen = 1 if x == da_i else 0
                                x_dil = d_map.get(x)
                                misses_other_b_night = x_dil in DAY_NAMES and DAY_NAMES.index(x_dil) in [n for n in active_b_nights if n != wd]
                                if wd in active_b_nights:
                                    # On Wednesday (and other B-night days), prioritize consultants whose DIL is Monday (or the other B-night day)
                                    # and those behind on nights so they catch up on nights
                                    dil_night_catchup = -1 if misses_other_b_night and not_over < 1.0 else 0
                                    needs_nights = 0 if not_over < 0 else 1
                                    if allow_dual_cover:
                                        return (dual_pen, is_runaway, bb_pen, dil_night_catchup, needs_nights, int(round(not_over)), int(round(fot)), g_pen, b_after_a_pen, fot)
                                    return (is_runaway, bb_pen, dil_night_catchup, needs_nights, int(round(not_over)), int(round(fot)), g_pen, b_after_a_pen, fot)
                                else:
                                    # On non-night B days (Tue/Thu/Fri), keep Monday-DIL / night-deficit consultants free for Wednesday and B-weeks
                                    save_for_b_night = 1 if (x_dil in DAY_NAMES and DAY_NAMES.index(x_dil) in active_b_nights and not_over < 1.0) or not_over < -0.5 else 0
                                    if allow_dual_cover:
                                        return (dual_pen, is_runaway, bb_pen, save_for_b_night, int(round(fot)), -int(round(not_over)), g_pen, b_after_a_pen, fot)
                                    return (is_runaway, bb_pen, save_for_b_night, int(round(fot)), -int(round(not_over)), g_pen, b_after_a_pen, fot)
                            best = min(avail, key=daily_score)
                            row["Consultant B"] = name_map.get(best)
                            add_score(best, "B_days", curr_d)
                            recent_b_shifts[best] = curr_d
                        else:
                            row["Consultant B"] = "UNCOVERED"
                    else:
                        row["Consultant B"] = name_map.get(cur_b, "UNCOVERED")
                        add_score(cur_b, "B_days", curr_d)
                        recent_b_shifts[cur_b] = curr_d

                    if wd in active_b_nights:
                        row["Nights"] = row["Consultant B"]
                        b_ini = ini_by_name.get(row["Consultant B"])
                        if b_ini: add_score(b_ini, "Nights", curr_d)

            # Per-consultant Nights preference override ("Separate Nights Only" / "No"):
            # If the daytime consultant assigned to tonight's night does not cover nights with their daytime block,
            # replace tonight's night with an eligible consultant from the night pool.
            if has_custom_nights_pref and row["Nights"] and row["Nights"] != "UNCOVERED":
                cur_n_ini = ini_by_name.get(row["Nights"])
                b_day_ini = ini_by_name.get(row["Consultant B"])
                if cur_n_ini and (
                    nights_pref_map.get(cur_n_ini, "Yes") == "No"
                    or (nights_pref_map.get(cur_n_ini, "Yes") == "Separate Nights Only" and cur_n_ini in (da_i, b_day_ini))
                ):
                    if curr_d_start <= curr_d <= fin_d: period_counts[cur_n_ini]["Nights"] -= 1
                    if ytd_s_d <= curr_d <= ytd_e_d: counts[cur_n_ini]["Nights"] -= 1
                    sim_n_counts[cur_n_ini] -= 1
                    pool_n = [
                        c for c in active_initials
                        if nights_pref_map.get(c, "Yes") in ("Yes", "Separate Nights Only")
                        and (nights_pref_map.get(c, "Yes") == "Yes" or c not in (da_i, b_day_ini))
                        and on_serv(c, curr_d)
                        and get_eff_marker(curr_d, c) not in B_DAY_BLOCK
                    ]
                    if pool_n:
                        best_pool_n = min(
                            pool_n,
                            key=lambda x: (
                                1 if x in (da_i, b_day_ini) else 0,
                                1 if x in recent_n_shifts and (curr_d - recent_n_shifts[x]).days < 2 else 0,
                                int(round(get_night_over_target(x))),
                                get_night_over_target(x),
                            ),
                        )
                        row["Nights"] = name_map.get(best_pool_n, "UNCOVERED")
                        add_score(best_pool_n, "Nights", curr_d)
                        recent_n_shifts[best_pool_n] = curr_d
                    else:
                        row["Nights"] = "UNCOVERED"

            if is_bh_today:
                b_ini = ini_by_name.get(row["Consultant B"])
                bh_grp = get_bh_group(curr_d)
                for ini in [da_i, b_ini]:
                    if ini and ini in period_counts:
                        if curr_d_start <= curr_d <= fin_d: period_counts[ini]["BH_days"] += 1
                        if ytd_s_d <= curr_d <= ytd_e_d:
                            counts[ini]["BH_days"] += 1
                            if bh_grp: counts[ini]["BH_groups"].add(bh_grp)

            if da_i: last_a_end[da_i] = curr_d
            if curr_d == fin_d: final_cur_a, final_cur_b = cur_a, cur_b

        for i in active_initials:
            row[i] = get_eff_marker(curr_d, i)
        rota_data.append(row)
        curr_d += timedelta(days=1)

    # AL / SL days requested in the Period and in the Financial GOSH Year (1 Apr to 31 Mar)
    fy_s, fy_e = gosh_fy_bounds(curr_d_start, fin_d)
    leave_ct = {i: {"AL_p": 0, "SL_p": 0, "AL_fy": 0, "SL_fy": 0} for i in active_initials}
    for r_row in rota_data:
        r_d = datetime.strptime(r_row["Date"], "%d/%m/%Y").date()
        for i in active_initials:
            mk = r_row.get(i)
            if mk in ("AL", "SL"):
                if curr_d_start <= r_d <= fin_d: leave_ct[i][mk + "_p"] += 1
                if fy_s <= r_d <= fy_e: leave_ct[i][mk + "_fy"] += 1

    state = {"A_MASTER_SEQ": ",".join(master_a_seq), "B_QUEUE": ",".join(active_initials),
             "CURRENT_A": final_cur_a or "", "CURRENT_B": final_cur_b or ""}
    summ = [{"Initials": k, "FTE A": fte_a_at(k, fin_d), "FTE B": fte_b_at(k, fin_d),
             "Period A": period_counts[k]["A_days"], "Period B": period_counts[k]["B_days"], "Period Nights": period_counts[k]["Nights"],
             "Period Total (A+N)": period_counts[k]["A_days"] + period_counts[k]["Nights"],
             "YTD A": counts[k]["A_days"], "YTD B": counts[k]["B_days"], "YTD Nights": counts[k]["Nights"],
             "YTD Total (A+N)": counts[k]["A_days"] + counts[k]["Nights"], "YTD BH": counts[k]["BH_days"],
             "Period AL (days)": leave_ct[k]["AL_p"], "Period SL (days)": leave_ct[k]["SL_p"],
             "GOSH Year AL (days)": leave_ct[k]["AL_fy"], "GOSH Year SL (days)": leave_ct[k]["SL_fy"]} for k in active_initials]
    return pd.DataFrame(rota_data), pd.DataFrame(summ), state


def _schedule_extra_shift(ex_cfg, consultants, leave_df, sel_s, sel_e, fte_hist_out, on_service, rota_df=None, a_name=None, b_name=None):
    """Auto-schedules a 3rd parallel shift (2-Day Weekend, 5-Day Weekday, or 7-Day Continuous) over [sel_s, sel_e]."""
    ex_name = ex_cfg["Shift Name"]
    ex_type = ex_cfg.get("Type", "2-Day Weekend")
    start_wd = DAY_NAMES.index(ex_cfg.get("Start Day", "Saturday" if ex_type == "2-Day Weekend" else "Monday"))
    min_gap_days = max(0, int(ex_cfg.get("Min Rest Weeks", 2)) * 7)
    allow_dual = safe_bool(ex_cfg.get("Allow Dual Cover", False), False)
    ex_nights = parse_weekdays(ex_cfg.get("Nights Covered", "None"), []) if ex_cfg.get("Nights Mode") != "None" else []
    tl_ex = build_fte_timeline(fte_hist_out, ex_name)

    active_initials = [safe_ini(c.get("Initials")) for c in consultants if c.get("Active", True) and safe_ini(c.get("Initials"))]
    name_map = {safe_ini(c.get("Initials")): c.get("Full Name", "") for c in consultants if safe_ini(c.get("Initials"))}
    ini_by_name = {v: k for k, v in name_map.items()}
    d_map = {safe_ini(c.get("Initials")): safe_raw(c.get("DIL Day")).capitalize() for c in consultants if safe_ini(c.get("Initials"))}
    weekdays_map = {safe_ini(c.get("Initials")): safe_bool(c.get("Weekdays", True), True) for c in consultants if safe_ini(c.get("Initials"))}
    rest_on_call_map = {safe_ini(c.get("Initials")): safe_bool(c.get("Rest Days on Call", False), False) for c in consultants if safe_ini(c.get("Initials"))}
    rest_days_map = {
        safe_ini(c.get("Initials")): set(parse_weekdays(c.get("Rest Days (Mon-Fri)", ""), [])) if safe_bool(c.get("Rest Days on Call", False), False) else set()
        for c in consultants if safe_ini(c.get("Initials"))
    }
    base_fte = {}
    for c in consultants:
        ini = safe_ini(c.get("Initials"))
        if ini:
            f = pd.to_numeric(c.get(f"FTE {ex_name}", 1.0), errors="coerce")
            base_fte[ini] = 0.0 if (pd.notna(f) and float(f) <= 0.0) else (1.0 if pd.isna(f) else max(0.1, float(f)))

    def fte_ex(ini, d_date):
        lst = tl_ex.get(ini)
        val = base_fte.get(ini, 1.0)
        if lst:
            val = lst[0][1]
            for eff, f in lst:
                if eff <= d_date:
                    val = f
                else:
                    break
        return 0.0 if val <= 0.0 else max(0.1, float(val))

    def get_mk(d_date, ini):
        ts = pd.Timestamp(d_date).normalize()
        if leave_df is not None and not leave_df.empty and ts in leave_df.index and ini in leave_df.columns:
            v = leave_df.loc[ts, ini]
            if isinstance(v, pd.Series):
                v = v.iloc[0]
            val = safe_ini(v)
            if val:
                return val
        if d_map.get(ini) == DAY_NAMES[d_date.weekday()]:
            return "DIL"
        if rest_on_call_map.get(ini) and d_date.weekday() in rest_days_map.get(ini, set()):
            return "DIL"
        return ""

    busy_by_date = defaultdict(set)
    if rota_df is not None:
        for _, r in rota_df.iterrows():
            rd = parse_any_date(r.get("Date"))
            if rd:
                for col in (a_name, b_name):
                    if col and col in rota_df.columns:
                        ini = ini_by_name.get(r.get(col))
                        if ini:
                            busy_by_date[rd].add(ini)

    s_d, e_d = to_date(sel_s), to_date(sel_e)
    sim_counts = {i: 0.0 for i in active_initials}
    last_start = {}
    assignments = {}
    weekend_workers = defaultdict(set)
    cur_c = None
    d = s_d
    while d <= e_d:
        wd = d.weekday()
        active_today = (wd >= 5) if ex_type == "2-Day Weekend" else ((wd < 5 and not (uk_holidays is not None and d in uk_holidays)) if ex_type == "5-Day Weekday" else True)
        if active_today:
            is_block_start = (wd == 5) if ex_type == "2-Day Weekend" else (wd == start_wd)
            if is_block_start or not cur_c or get_mk(d, cur_c) in B_DAY_BLOCK or (not allow_dual and cur_c in busy_by_date.get(d, set())):
                cands = [
                    i for i in active_initials
                    if (on_service.get(i) is None or d >= on_service.get(i))
                    and fte_ex(i, d) > 0.0
                    and (ex_type != "5-Day Weekday" or weekdays_map.get(i, True))
                    and get_mk(d, i) not in B_DAY_BLOCK
                    and (allow_dual or i not in busy_by_date.get(d, set()))
                ]
                if cands:
                    cur_c = min(
                        cands,
                        key=lambda x: (
                            1 if x in busy_by_date.get(d, set()) else 0,
                            1 if x in last_start and (d - last_start[x]).days < min_gap_days else 0,
                            sim_counts[x] / max(0.1, fte_ex(x, d)),
                        ),
                    )
                    if is_block_start:
                        last_start[cur_c] = d
                else:
                    cur_c = None
            if cur_c:
                assignments[d] = name_map.get(cur_c, cur_c)
                sim_counts[cur_c] += 1.0
                if wd >= 5:
                    weekend_workers[d].add(cur_c)
            else:
                assignments[d] = "UNCOVERED"
        d += timedelta(days=1)
    return assignments, weekend_workers, ex_nights


# =============================================================================
# V2 ADAPTER: Settings tables -> engine -> named output columns
# =============================================================================
def run_rota_v2(leave_df, staff_df, shifts_df, sel_s, sel_e, ytd_s, ytd_e, locum_dict, memory_file=None, fte_history=None):
    clean_sh = clean_shifts(shifts_df)
    a_cfg, b_cfg, extras = pick_engine_shifts(clean_sh)
    if a_cfg is None:
        raise ValueError("Add at least one shift in Settings.")
    parallel_n = int(a_cfg.get("Parallel Consultants", 2))
    if b_cfg is None:
        b_cfg = {
            "Shift Name": "Consultant B", "Type": "5-Day Weekday", "Start Day": "Monday", "Min Rest Weeks": 3,
            "Parallel Consultants": parallel_n, "Nights Mode": a_cfg.get("Nights Mode", "With Primary / Weekday"),
            "Nights Covered": "None", "Flip Nights on BH": False, "B Consultant Role": "Independent Parallel Rota",
            "Group Pairing": "Same Group Allowed", "Allow Dual Cover": False, "Monday Off After Weekend": False,
        }
        b_is_synthetic = True
    else:
        b_is_synthetic = False
    a_name, b_name = a_cfg["Shift Name"], b_cfg["Shift Name"]
    b_role = b_cfg.get("B Consultant Role", a_cfg.get("B Consultant Role", "Second Parallel Rota + Nights"))
    if a_cfg.get("B Consultant Role") == "Just Nights" or b_cfg.get("Type") == "Nights Only":
        b_role = "Just Nights"
    b_active = (not b_is_synthetic) and (parallel_n >= 2 or b_role == "Just Nights")

    staff = clean_staff_for_engine(sync_staff_fte_columns(staff_df, clean_sh))
    consultants = []
    for rec in staff.to_dict("records"):
        r = dict(rec)
        r["FTE A"] = rec.get(f"FTE {a_name}", 1.0)
        r["FTE B"] = rec.get(f"FTE {b_name}", 1.0)
        consultants.append(r)
    if not any(c.get("Active", True) for c in consultants):
        raise ValueError("No active consultants in the Staff table.")

    mem = parse_previous_rota(memory_file, consultants, sel_s, ytd_s, ytd_e, a_name, b_name) if memory_file is not None else empty_memory()
    warnings = list(mem["warnings"])

    # Effective-dated FTE + on-service dates
    fte_hist_out = update_fte_history(fte_history, staff, clean_sh, sel_s)
    on_service = {safe_ini(c["Initials"]): parse_any_date(c.get("On Service From")) for c in consultants}

    # Someone with no history at all while colleagues have plenty (new joiner / returner whose date was not set)
    # starts level instead of being treated as months behind.
    tally = lambda i: sum(mem["counts"].get(i, {}).get(k, 0) for k in ("A_days", "B_days", "Nights"))
    level_start = set()
    if any(tally(i) > 0 for i in mem["counts"]):
        for c in consultants:
            ini = safe_ini(c["Initials"])
            if c.get("Active", True) and not on_service.get(ini) and tally(ini) == 0:
                level_start.add(ini)
                warnings.append(f"{ini} has no shifts in the history and no 'On Service From' date, so they start level with colleagues (no catch-up). "
                                f"Set 'On Service From' to control this.")

    leave = leave_df if leave_df is not None else pd.DataFrame()
    if not mem["leave_df"].empty:
        leave = leave.combine_first(mem["leave_df"]) if not leave.empty else mem["leave_df"]

    extra_cfgs = [r for r in clean_sh.to_dict("records") if r["Shift Name"] in extras]
    pre_weekend_workers = defaultdict(set)
    if parallel_n >= 3 and extra_cfgs:
        for ex_c in extra_cfgs:
            if ex_c.get("Type") == "2-Day Weekend" and (safe_bool(ex_c.get("Monday Off After Weekend"), False) or safe_bool(a_cfg.get("Monday Off After Weekend"), False)):
                _, wknd_w, _ = _schedule_extra_shift(ex_c, consultants, leave, sel_s, sel_e, fte_hist_out, on_service)
                for kd, vset in wknd_w.items():
                    pre_weekend_workers[kd].update(vset)

    scheduler_cfg = {
        "parallel_consultants": parallel_n,
        "a_type": a_cfg.get("Type", "7-Day Continuous"),
        "b_type": b_cfg.get("Type", "5-Day Weekday"),
        "b_start_wd": DAY_NAMES.index(b_cfg.get("Start Day", "Monday")),
        "nights_mode": a_cfg.get("Nights Mode", "Alternating Consultants"),
        "a_nights_covered": parse_weekdays(a_cfg.get("Nights Covered"), [1, 3, 4, 5, 6]),
        "b_nights_covered": parse_weekdays(b_cfg.get("Nights Covered"), [0, 2]),
        "flip_nights_on_bh": safe_bool(a_cfg.get("Flip Nights on BH", True), True) and safe_bool(b_cfg.get("Flip Nights on BH", True), True),
        "b_role": b_role,
        "group_pairing": b_cfg.get("Group Pairing", a_cfg.get("Group Pairing", "Different Group Ideally")),
        "allow_dual_cover": safe_bool(a_cfg.get("Allow Dual Cover", False), False) or safe_bool(b_cfg.get("Allow Dual Cover", False), False),
        "monday_off_after_weekend": safe_bool(a_cfg.get("Monday Off After Weekend", False), False) or safe_bool(b_cfg.get("Monday Off After Weekend", False), False) or any(safe_bool(ec.get("Monday Off After Weekend", False), False) for ec in extra_cfgs),
        "b_active": b_active,
        "weekend_workers_by_date": pre_weekend_workers,
    }

    rota, summ, state = generate_rota(
        leave, consultants, sel_s, sel_e, mem["counts"], mem["master_a_seq"], mem["cycle_a"], [],
        mem["curr_a"], mem["curr_b"], mem["last_a_dates"], locum_dict, mem["hist_b_dates"], ytd_s, ytd_e, mem["last_b_week"],
        a_anchor_wd=DAY_NAMES.index(a_cfg["Start Day"]),
        a_min_gap_days=max(7, int(a_cfg["Min Rest Weeks"]) * 7),
        b_bb_gap_days=max(0, int(b_cfg["Min Rest Weeks"]) * 7),
        fte_tl_a=build_fte_timeline(fte_hist_out, a_name), fte_tl_b=build_fte_timeline(fte_hist_out, b_name),
        on_service=on_service, level_start=level_start,
        scheduler_cfg=scheduler_cfg,
    )

    rota = rota.rename(columns={"Consultant A": a_name, "Consultant B": b_name})
    insert_at = list(rota.columns).index(b_name) + 1
    unscheduled_extras = []
    for i, ex in enumerate(extras):
        rota.insert(insert_at + i, ex, "")
        if parallel_n >= 3:
            ex_cfg = next((r for r in extra_cfgs if r["Shift Name"] == ex), None)
            if ex_cfg is not None:
                assigns, _, ex_nights = _schedule_extra_shift(ex_cfg, consultants, leave, sel_s, sel_e, fte_hist_out, on_service, rota, a_name, b_name)
                rota[ex] = [assigns.get(parse_any_date(d_s), "") for d_s in rota["Date"]]
                if ex_nights:
                    name_to_ini = {c.get("Full Name", ""): safe_ini(c.get("Initials")) for c in consultants}
                    ini_to_name = {safe_ini(c.get("Initials")): c.get("Full Name", "") for c in consultants if safe_ini(c.get("Initials"))}
                    npref_map = {safe_ini(c.get("Initials")): safe_raw(c.get("Nights")) or "Yes" for c in consultants if safe_ini(c.get("Initials"))}
                    act_inis = [safe_ini(c.get("Initials")) for c in consultants if c.get("Active", True) and safe_ini(c.get("Initials"))]
                    for r_idx, r_row in rota.iterrows():
                        rd = parse_any_date(r_row["Date"])
                        if rd and sel_s <= rd <= sel_e and rd.weekday() in ex_nights:
                            cand_name = r_row[ex]
                            if not cand_name or cand_name == "UNCOVERED":
                                cand_name = r_row.get(a_name) or ""
                            cand_ini = name_to_ini.get(cand_name)
                            if cand_ini and npref_map.get(cand_ini, "Yes") in ("No", "Separate Nights Only"):
                                day_inis = {name_to_ini.get(r_row.get(col)) for col in (a_name, b_name, ex) if col in rota.columns}
                                pool = [
                                    i for i in act_inis
                                    if npref_map.get(i, "Yes") in ("Yes", "Separate Nights Only")
                                    and (npref_map.get(i, "Yes") == "Yes" or i not in day_inis)
                                    and (on_service.get(i) is None or rd >= on_service.get(i))
                                    and r_row.get(i, "") not in B_DAY_BLOCK
                                ]
                                cand_name = ini_to_name.get(pool[0], "UNCOVERED") if pool else "UNCOVERED"
                            if cand_name:
                                rota.at[r_idx, "Nights"] = cand_name
        else:
            unscheduled_extras.append(ex)

    # Recompute Period Nights / YTD Nights on summ if a 3rd shift updated Nights
    if parallel_n >= 3 and any(parse_weekdays(ec.get("Nights Covered", "None"), []) for ec in extra_cfgs):
        name_to_ini = {c.get("Full Name", ""): safe_ini(c.get("Initials")) for c in consultants}
        p_n = {ini: 0 for ini in summ["Initials"]}
        y_n = {ini: mem["counts"].get(ini, {}).get("Nights", 0) for ini in summ["Initials"]}
        s_d, e_d, ys_d, ye_d = to_date(sel_s), to_date(sel_e), to_date(ytd_s), to_date(ytd_e)
        for _, r_row in rota.iterrows():
            rd = parse_any_date(r_row["Date"])
            n_ini = name_to_ini.get(r_row.get("Nights"))
            if rd and n_ini in p_n:
                if s_d <= rd <= e_d: p_n[n_ini] += 1
                if ys_d <= rd <= ye_d: y_n[n_ini] += 1
        summ["Period Nights"] = summ["Initials"].map(p_n)
        summ["Period Total (A+N)"] = summ["Period A"] + summ["Period Nights"]
        summ["YTD Nights"] = summ["Initials"].map(y_n)
        summ["YTD Total (A+N)"] = summ["YTD A"] + summ["YTD Nights"]

    summ = summ.rename(columns={
        "FTE A": f"FTE {a_name}", "FTE B": f"FTE {b_name}",
        "Period A": f"Period {a_name}", "Period B": f"Period {b_name}", "Period Total (A+N)": f"Period Total ({a_name}+Nights)",
        "YTD A": f"YTD {a_name}", "YTD B": f"YTD {b_name}", "YTD Total (A+N)": f"YTD Total ({a_name}+Nights)"})
    name_lookup = {safe_ini(c["Initials"]): c.get("Full Name", "") for c in consultants}
    summ.insert(1, "Full Name", summ["Initials"].map(name_lookup))

    # Add AL / SL days from Previous Period Rota that fall inside the Financial GOSH Year (1 Apr to 31 Mar)
    fy_s, fy_e = gosh_fy_bounds(sel_s, sel_e)
    prev_df = mem["prev_rota_df"]
    if prev_df is not None and not prev_df.empty and "Date" in prev_df.columns:
        prev_al = {i: 0 for i in summ["Initials"]}
        prev_sl = {i: 0 for i in summ["Initials"]}
        for _, prow in prev_df.iterrows():
            pd_d = parse_any_date(prow.get("Date"))
            if pd_d and fy_s <= pd_d <= fy_e:
                for ini in prev_al:
                    mk = safe_ini(prow.get(ini))
                    if mk == "AL": prev_al[ini] += 1
                    elif mk == "SL": prev_sl[ini] += 1
        summ["GOSH Year AL (days)"] = summ["GOSH Year AL (days)"] + summ["Initials"].map(prev_al)
        summ["GOSH Year SL (days)"] = summ["GOSH Year SL (days)"] + summ["Initials"].map(prev_sl)

    # Worked BEFORE this period (within the YTD window): fixed starting point for the live Excel shift formulas
    prior = mem["counts"]
    summ[f"Prior YTD {a_name}"] = summ["Initials"].map(lambda i: prior.get(i, {}).get("A_days", 0))
    summ[f"Prior YTD {b_name}"] = summ["Initials"].map(lambda i: prior.get(i, {}).get("B_days", 0))
    summ["Prior YTD Nights"] = summ["Initials"].map(lambda i: prior.get(i, {}).get("Nights", 0))
    summ["Prior YTD BH"] = summ["Initials"].map(lambda i: prior.get(i, {}).get("BH_days", 0))

    state["PERIOD_START"] = to_date(sel_s).strftime("%d/%m/%Y")
    state["PERIOD_END"] = to_date(sel_e).strftime("%d/%m/%Y")
    state["YTD_START"] = to_date(ytd_s).strftime("%d/%m/%Y")
    state["YTD_END"] = to_date(ytd_e).strftime("%d/%m/%Y")
    state["GOSH_FY_START"] = fy_s.strftime("%d/%m/%Y")
    state["GOSH_FY_END"] = fy_e.strftime("%d/%m/%Y")
    state["ENGINE_A_SHIFT"], state["ENGINE_B_SHIFT"] = a_name, b_name

    if unscheduled_extras:
        warnings.append(f"Not auto-filled yet (Phase 2 Option 2): {', '.join(unscheduled_extras)}. These columns are left blank for manual entry.")
    return rota, summ, state, mem["prev_rota_df"], warnings, fte_hist_out


def recompute_summary_from_rota(rota_df, summ_df, state_dict, prev_rota_df=None):
    """Recalculates the Summary (Fairness Calculator) DataFrame live from rota_df so any manual edits
    made in the interactive Rota preview immediately update Period, YTD, BH, and AL/SL counts."""
    if rota_df is None or summ_df is None or summ_df.empty:
        return summ_df
    out = summ_df.copy()
    a_name = state_dict.get("ENGINE_A_SHIFT", "Consultant A")
    b_name = state_dict.get("ENGINE_B_SHIFT", "Consultant B")
    ps = parse_any_date(state_dict.get("PERIOD_START"))
    pe = parse_any_date(state_dict.get("PERIOD_END"))
    ys = parse_any_date(state_dict.get("YTD_START"))
    ye = parse_any_date(state_dict.get("YTD_END"))
    fs = parse_any_date(state_dict.get("GOSH_FY_START")) or (gosh_fy_bounds(ps, pe)[0] if ps and pe else None)
    fe = parse_any_date(state_dict.get("GOSH_FY_END")) or (gosh_fy_bounds(ps, pe)[1] if ps and pe else None)
    if not all([ps, pe, ys, ye, fs, fe]):
        return out

    inis = [safe_ini(i) for i in out["Initials"]]
    token_to_ini = {}
    for _, srow in out.iterrows():
        ini = safe_ini(srow.get("Initials"))
        fname = safe_raw(srow.get("Full Name"))
        if ini:
            token_to_ini[ini.upper()] = ini
        if fname:
            token_to_ini[fname.upper()] = ini

    def resolve_ini(val):
        s = safe_raw(val).upper()
        if not s or s == "UNCOVERED":
            return None
        return token_to_ini.get(s)

    per_a = {i: 0 for i in inis}
    per_b = {i: 0 for i in inis}
    per_n = {i: 0 for i in inis}
    ytd_a = {i: 0 for i in inis}
    ytd_b = {i: 0 for i in inis}
    ytd_n = {i: 0 for i in inis}
    ytd_bh = {i: 0 for i in inis}
    per_al = {i: 0 for i in inis}
    per_sl = {i: 0 for i in inis}
    fy_al = {i: 0 for i in inis}
    fy_sl = {i: 0 for i in inis}

    for _, r_row in rota_df.iterrows():
        dt = parse_any_date(r_row.get("Date"))
        if dt is None:
            continue
        in_per = ps <= dt <= pe
        in_ytd = ys <= dt <= ye
        in_fy = fs <= dt <= fe
        is_bh = dt.weekday() < 5 and uk_holidays is not None and dt in uk_holidays

        ai = resolve_ini(r_row.get(a_name)) if a_name in rota_df.columns else None
        bi = resolve_ini(r_row.get(b_name)) if b_name in rota_df.columns else None
        ni = resolve_ini(r_row.get("Nights")) if "Nights" in rota_df.columns else None

        if ai in per_a:
            if in_per:
                per_a[ai] += 1
            if in_ytd:
                ytd_a[ai] += 1
                if is_bh:
                    ytd_bh[ai] += 1
        if bi in per_b:
            if in_per:
                per_b[bi] += 1
            if in_ytd:
                ytd_b[bi] += 1
                if is_bh:
                    ytd_bh[bi] += 1
        if ni in per_n:
            if in_per:
                per_n[ni] += 1
            if in_ytd:
                ytd_n[ni] += 1

        for ini in inis:
            if ini in rota_df.columns:
                mk = safe_ini(r_row.get(ini))
                if mk == "AL":
                    if in_per:
                        per_al[ini] += 1
                    if in_fy:
                        fy_al[ini] += 1
                elif mk == "SL":
                    if in_per:
                        per_sl[ini] += 1
                    if in_fy:
                        fy_sl[ini] += 1

    if prev_rota_df is not None and not prev_rota_df.empty and "Date" in prev_rota_df.columns:
        for _, prow in prev_rota_df.iterrows():
            pd_d = parse_any_date(prow.get("Date"))
            if pd_d and fs <= pd_d <= fe:
                for ini in inis:
                    mk = safe_ini(prow.get(ini))
                    if mk == "AL":
                        fy_al[ini] += 1
                    elif mk == "SL":
                        fy_sl[ini] += 1

    prior_a = out[f"Prior YTD {a_name}"] if f"Prior YTD {a_name}" in out.columns else 0
    prior_b = out[f"Prior YTD {b_name}"] if f"Prior YTD {b_name}" in out.columns else 0
    prior_n = out["Prior YTD Nights"] if "Prior YTD Nights" in out.columns else 0
    prior_bh = out["Prior YTD BH"] if "Prior YTD BH" in out.columns else 0

    out[f"Period {a_name}"] = out["Initials"].map(per_a)
    out[f"Period {b_name}"] = out["Initials"].map(per_b)
    out["Period Nights"] = out["Initials"].map(per_n)
    out[f"Period Total ({a_name}+Nights)"] = out[f"Period {a_name}"] + out["Period Nights"]
    out[f"YTD {a_name}"] = prior_a + out["Initials"].map(ytd_a)
    out[f"YTD {b_name}"] = prior_b + out["Initials"].map(ytd_b)
    out["YTD Nights"] = prior_n + out["Initials"].map(ytd_n)
    out[f"YTD Total ({a_name}+Nights)"] = out[f"YTD {a_name}"] + out["YTD Nights"]
    out["YTD BH"] = prior_bh + out["Initials"].map(ytd_bh)
    out["Period AL (days)"] = out["Initials"].map(per_al)
    out["Period SL (days)"] = out["Initials"].map(per_sl)
    out["GOSH Year AL (days)"] = out["Initials"].map(fy_al)
    out["GOSH Year SL (days)"] = out["Initials"].map(fy_sl)
    return out


# =============================================================================
# MASTER EXCEL EXPORT
# =============================================================================
def _col(df, name):
    """1-based Excel column letter of a DataFrame column, or None."""
    return get_column_letter(list(df.columns).index(name) + 1) if name in df.columns else None


def _with_real_dates(df):
    """Copy with the text 'Date' column converted to real Excel dates (needed for date-bounded formulas)."""
    out = df.copy()
    if "Date" in out.columns:
        out["Date"] = [parse_any_date(d) or d for d in out["Date"]]
    return out


def _write_live_summary(ws, summary_df, rota_df, state_dict, prev_rota_df=None):
    """Overwrite the Summary sheet's count cells with formulas that read the Rota sheet, so manual edits
    to the rota (whether typed as Full Name or Initials) update the fairness table instantly."""
    a_name, b_name = state_dict.get("ENGINE_A_SHIFT"), state_dict.get("ENGINE_B_SHIFT")
    ca, cb, cn = _col(rota_df, a_name), _col(rota_df, b_name), _col(rota_df, "Nights")
    ps, pe = parse_any_date(state_dict.get("PERIOD_START")), parse_any_date(state_dict.get("PERIOD_END"))
    ys, ye = parse_any_date(state_dict.get("YTD_START")), parse_any_date(state_dict.get("YTD_END"))
    fs = parse_any_date(state_dict.get("GOSH_FY_START")) or gosh_fy_bounds(ps, pe)[0]
    fe = parse_any_date(state_dict.get("GOSH_FY_END")) or gosh_fy_bounds(ps, pe)[1]
    if not all([ca, cb, cn, ps, pe, ys, ye, fs, fe]):
        return False

    sc = {c: get_column_letter(i + 1) for i, c in enumerate(summary_df.columns)}
    needed = [f"Period {a_name}", f"Period {b_name}", "Period Nights", f"Period Total ({a_name}+Nights)",
              f"YTD {a_name}", f"YTD {b_name}", "YTD Nights", f"YTD Total ({a_name}+Nights)", "YTD BH",
              f"Prior YTD {a_name}", f"Prior YTD {b_name}", "Prior YTD Nights", "Prior YTD BH",
              "Period AL (days)", "Period SL (days)", "GOSH Year AL (days)", "GOSH Year SL (days)", "Full Name"]
    if any(n not in sc for n in needed):
        return False

    n_rows = len(summary_df)
    last = len(rota_df) + 1
    rng = lambda col: f"Rota!${col}$2:${col}${last}"
    dr = rng("A")
    has_prev = prev_rota_df is not None and not prev_rota_df.empty
    plast = (len(prev_rota_df) + 1) if has_prev else 0
    prng = lambda col: f"'Previous Period Rota'!${col}$2:${col}${plast}"

    # Editable date parameters, to the right of the table
    lab_c, val_c = len(summary_df.columns) + 2, len(summary_df.columns) + 3
    vl = get_column_letter(val_c)
    for r, (label, val) in enumerate([("Period start", ps), ("Period end", pe), ("YTD start", ys), ("YTD end", ye),
                                      ("GOSH Year start", fs), ("GOSH Year end", fe)], start=1):
        ws.cell(row=r, column=lab_c, value=label)
        cell = ws.cell(row=r, column=val_c, value=val)
        cell.number_format = "DD/MM/YYYY"
    ws.cell(row=8, column=lab_c, value="Counts are live formulas on the Rota sheet. Edit names or initials there (or the dates above) and this table updates.")
    ws.cell(row=9, column=lab_c, value="'Prior YTD' = shifts worked before this period (from the uploaded history).")

    def cnt(col, r, lo, hi, match_ini=True):
        base = f'COUNTIFS({rng(col)},$B{r},{dr},">="&${vl}${lo},{dr},"<="&${vl}${hi})'
        if match_ini:
            base += f'+COUNTIFS({rng(col)},$A{r},{dr},">="&${vl}${lo},{dr},"<="&${vl}${hi})'
        return base

    bh_dates = []
    if uk_holidays is not None:
        lo_d, hi_d = min(ps, ys), max(pe, ye)
        d = lo_d
        while d <= hi_d:
            if d.weekday() < 5 and d in uk_holidays:
                bh_dates.append(d)
            d += timedelta(days=1)
    if bh_dates:
        bh_ws = ws.parent.create_sheet("BH_Dates")
        bh_ws.append(["Date", "Bank Holiday", a_name, b_name])
        for idx_bh, d in enumerate(bh_dates, start=2):
            f_a = f'=IF(AND($A{idx_bh}>=Summary!${vl}$3,$A{idx_bh}<=Summary!${vl}$4),IFERROR(INDEX({rng(ca)},MATCH($A{idx_bh},{dr},0)),""),"")'
            f_b = f'=IF(AND($A{idx_bh}>=Summary!${vl}$3,$A{idx_bh}<=Summary!${vl}$4),IFERROR(INDEX({rng(cb)},MATCH($A{idx_bh},{dr},0)),""),"")'
            bh_ws.append([d, uk_holidays.get(d), f_a, f_b])
        for row in bh_ws.iter_rows(min_row=2, max_col=1):
            row[0].number_format = "DD/MM/YYYY"
        bh_ws.column_dimensions["A"].width = 14
        bh_ws.column_dimensions["B"].width = 28
        bh_ws.column_dimensions["C"].width = 22
        bh_ws.column_dimensions["D"].width = 22
        bh_last = len(bh_dates) + 1

    for r in range(2, n_rows + 2):
        ini = safe_ini(summary_df.iloc[r - 2]["Initials"])
        fname = safe_raw(summary_df.iloc[r - 2].get("Full Name", ""))
        match_both = bool(ini and fname and ini.upper() != fname.upper())

        ws[f"{sc[f'Period {a_name}']}{r}"] = f"={cnt(ca, r, 1, 2, match_both)}"
        ws[f"{sc[f'Period {b_name}']}{r}"] = f"={cnt(cb, r, 1, 2, match_both)}"
        ws[f"{sc['Period Nights']}{r}"] = f"={cnt(cn, r, 1, 2, match_both)}"
        ws[f"{sc[f'Period Total ({a_name}+Nights)']}{r}"] = f"={sc[f'Period {a_name}']}{r}+{sc['Period Nights']}{r}"
        ws[f"{sc[f'YTD {a_name}']}{r}"] = f"={sc[f'Prior YTD {a_name}']}{r}+{cnt(ca, r, 3, 4, match_both)}"
        ws[f"{sc[f'YTD {b_name}']}{r}"] = f"={sc[f'Prior YTD {b_name}']}{r}+{cnt(cb, r, 3, 4, match_both)}"
        ws[f"{sc['YTD Nights']}{r}"] = f"={sc['Prior YTD Nights']}{r}+{cnt(cn, r, 3, 4, match_both)}"
        ws[f"{sc[f'YTD Total ({a_name}+Nights)']}{r}"] = f"={sc[f'YTD {a_name}']}{r}+{sc['YTD Nights']}{r}"
        if bh_dates:
            bh_f = f"COUNTIF(BH_Dates!$C$2:$C${bh_last},$B{r})+COUNTIF(BH_Dates!$D$2:$D${bh_last},$B{r})"
            if match_both:
                bh_f += f"+COUNTIF(BH_Dates!$C$2:$C${bh_last},$A{r})+COUNTIF(BH_Dates!$D$2:$D${bh_last},$A{r})"
            ws[f"{sc['YTD BH']}{r}"] = f"={sc['Prior YTD BH']}{r}+{bh_f}"
        else:
            ws[f"{sc['YTD BH']}{r}"] = f"={sc['Prior YTD BH']}{r}"

        l_col = _col(rota_df, ini)  # this person's leave-marker column on the Rota sheet
        p_col = _col(prev_rota_df, ini) if has_prev else None
        if l_col:
            for kind in ("AL", "SL"):
                lr = rng(l_col)
                ws[f"{sc[f'Period {kind} (days)']}{r}"] = f'=COUNTIFS({lr},"{kind}",{dr},">="&${vl}$1,{dr},"<="&${vl}$2)'
                fy_formula = f'COUNTIFS({lr},"{kind}",{dr},">="&${vl}$5,{dr},"<="&${vl}$6)'
                if p_col:
                    fy_formula += f'+COUNTIFS({prng(p_col)},"{kind}",{prng("A")},">="&${vl}$5,{prng("A")},"<="&${vl}$6)'
                ws[f"{sc[f'GOSH Year {kind} (days)']}{r}"] = f"={fy_formula}"

    tot = n_rows + 2
    ws[f"A{tot}"] = "TOTAL"
    for name in needed[:-1]:
        ws[f"{sc[name]}{tot}"] = f"=SUM({sc[name]}2:{sc[name]}{n_rows + 1})"
    return True


def export_master_excel(config_shifts, config_staff, rota_df, summary_df, state_dict, prev_rota_df=None, fte_history=None):
    out = io.BytesIO()
    has_prev = prev_rota_df is not None and not prev_rota_df.empty
    rota_x = _with_real_dates(rota_df)
    prev_x = _with_real_dates(prev_rota_df) if has_prev else None
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        rota_x.to_excel(writer, index=False, sheet_name="Rota")
        if has_prev:
            prev_x.to_excel(writer, index=False, sheet_name="Previous Period Rota")
        summary_df.to_excel(writer, index=False, sheet_name="Summary")
        pd.DataFrame(list(state_dict.items()), columns=["Key", "Value"]).to_excel(writer, index=False, sheet_name="State")
        pd.DataFrame([["AL", "Annual Leave"], ["SL", "Study Leave"], ["OFF", "Off for other reasons"],
                      ["NO", "Please do not put me on call"], ["DIL", "Day in Lieu"],
                      ["REQ", "Please add me as A on this day. Mark every day you want: exactly those days are honoured, "
                              "and someone else is added in for the rest of that 7-day block."],
                      ["REQB", "Please add me as B Consultant on this day"],
                      ["SICK", "Sick leave: cannot be rostered, and the days do NOT have to be caught up (AL, SL, NO and OFF do)"]],
                     columns=["Marker", "Description"]).to_excel(writer, index=False, sheet_name="Legend")
        clean_shifts(config_shifts)[EXPORT_SHIFT_COLS].to_excel(writer, index=False, sheet_name="Config_Shifts")
        staff_out = sync_staff_fte_columns(config_staff.copy(), config_shifts)
        if "FTE Changed From" in staff_out.columns:
            staff_out["FTE Changed From"] = pd.NaT  # a typed change is now recorded in Config_FTE_History
        staff_out.to_excel(writer, index=False, sheet_name="Config_Staff")
        hist_out = normalize_fte_history(fte_history)
        if hist_out.empty:
            hist_out = history_from_staff(config_staff, config_shifts)
        hist_out.to_excel(writer, index=False, sheet_name="Config_FTE_History")

        live = _write_live_summary(writer.sheets["Summary"], summary_df, rota_x, state_dict, prev_x)
        writer.book.calculation.calcMode = "auto"
        writer.book.calculation.fullCalcOnLoad = True  # make Excel compute the formulas on open

        align_left = Alignment(horizontal="left", vertical="center")
        for s_name in writer.sheets:
            ws = writer.sheets[s_name]
            if s_name in ["Rota", "Previous Period Rota"]:
                ws.freeze_panes = "C2"
                for cell in ws["A"][1:]:
                    cell.number_format = "DD/MM/YYYY"
            if s_name in ("Config_Staff", "Config_FTE_History"):
                for ci in range(1, ws.max_column + 1):
                    if ws.cell(row=1, column=ci).value in ("On Service From", "FTE Changed From", "Effective From"):
                        for r_cells in ws.iter_rows(min_row=2, min_col=ci, max_col=ci):
                            r_cells[0].number_format = "DD/MM/YYYY"
            for i, col_cells in enumerate(ws.columns, 1):
                head = ws.cell(row=1, column=i).value
                vals = [len(str(cell.value)) if not (isinstance(cell.value, str) and cell.value.startswith("=")) else 8
                        for cell in col_cells if cell.value is not None and not (s_name == "Summary" and cell.row in (8, 9))]
                ws.column_dimensions[get_column_letter(i)].width = min(60, max(vals + [len(str(head)) if head else 5]) + 2)
                for cell in col_cells:
                    cell.alignment = align_left

        # Name drop-downs on the A / B / Nights columns (free text still allowed for locums etc.)
        n_staff = len(config_staff)
        if n_staff:
            dv = DataValidation(type="list", formula1=f"=Config_Staff!$B$2:$B${n_staff + 1}", allow_blank=True, showErrorMessage=False)
            writer.sheets["Rota"].add_data_validation(dv)
            for col in {_col(rota_x, c) for c in (state_dict.get("ENGINE_A_SHIFT"), state_dict.get("ENGINE_B_SHIFT"), "Nights")} - {None}:
                dv.add(f"{col}2:{col}{len(rota_x) + 1}")

        fill = PatternFill(start_color="E5E5EA", end_color="E5E5EA", fill_type="solid")

        def shade(ws, df, hide_before=None):
            for i, d_str in enumerate(df["Date"], start=2):
                dt = parse_any_date(d_str)
                if dt is None:
                    continue
                if dt.weekday() >= 5 or (uk_holidays is not None and dt in uk_holidays):
                    for c in range(1, len(df.columns) + 1):
                        ws.cell(row=i, column=c).fill = fill
                if hide_before and dt < hide_before:
                    ws.row_dimensions[i].hidden = True

        shade(writer.sheets["Rota"], rota_x)
        if has_prev and "Date" in prev_x.columns:
            dates = [parse_any_date(x) for x in prev_x["Date"]]
            dates = [d for d in dates if d]
            shade(writer.sheets["Previous Period Rota"], prev_x, (max(dates) - timedelta(days=365)) if dates else None)
    out.seek(0)
    return out


# =============================================================================
# STREAMLIT UI
# =============================================================================
def load_master_into_state(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    restored = []
    if "Config_Shifts" in xls.sheet_names:
        st.session_state.shifts_base = clean_shifts(pd.read_excel(xls, sheet_name="Config_Shifts"))
        restored.append("Config_Shifts")
    if "Config_Staff" in xls.sheet_names:
        df_staff = pd.read_excel(xls, sheet_name="Config_Staff")
        df_staff["DIL Day"] = df_staff.get("DIL Day", "").apply(safe_raw) if "DIL Day" in df_staff.columns else ""
        st.session_state.staff_base = sync_staff_fte_columns(df_staff, st.session_state.shifts_base)
        restored.append("Config_Staff")
        hist = normalize_fte_history(pd.read_excel(xls, sheet_name="Config_FTE_History")) if "Config_FTE_History" in xls.sheet_names else None
        if hist is None or hist.empty:  # older Master file: its saved FTEs are taken as the FTEs in force so far
            hist = history_from_staff(st.session_state.staff_base, st.session_state.shifts_base)
        st.session_state.fte_history_base = hist
        restored.append("FTE history")
    st.session_state.suggested_start = None
    if "State" in xls.sheet_names:
        df_state = pd.read_excel(xls, sheet_name="State")
        if df_state.shape[1] >= 2:
            sd = dict(zip(df_state.iloc[:, 0].astype(str).str.strip(), df_state.iloc[:, 1]))
            pe = parse_any_date(sd.get("PERIOD_END"))
            if pe:
                st.session_state.suggested_start = pe + timedelta(days=1)
    st.session_state.master_bytes = file_bytes
    st.session_state.editor_ver += 1
    return restored


def main():
    st.set_page_config(layout="wide", page_title="Consultant Rota Tool (V2.0)")
    st.markdown(
        "<style>"
        ".stApp {background-color: #fbfbfd !important;} "
        "p, h1, h2, h3, span, label, div {color: #1d1d1f; font-family: -apple-system, sans-serif;} "
        "div.stButton > button, div.stDownloadButton > button {background-color: #0071e3 !important; color: #fff !important; border-radius: 20px !important;} "
        "div[data-baseweb='select'] > div {border: 1.5px solid #0071e3 !important; border-radius: 8px !important; background-color: #ffffff !important; box-shadow: 0 1px 3px rgba(0,113,227,0.12) !important;} "
        "div[data-baseweb='select']:hover > div {border-color: #005bb5 !important; background-color: #f5f9ff !important;} "
        "</style>",
        unsafe_allow_html=True)
    st.title("Consultant Rota Dashboard (V2.0)")

    ss = st.session_state
    ss.setdefault("editor_ver", 0)
    ss.setdefault("shifts_base", pd.DataFrame(DEFAULT_SHIFTS))
    ss.setdefault("staff_base", sync_staff_fte_columns(pd.DataFrame(DEFAULT_STAFF), ss.shifts_base))
    ss.setdefault("locums_base", pd.DataFrame({"Start Date": pd.Series(dtype="datetime64[ns]"), "End Date": pd.Series(dtype="datetime64[ns]"), "Locum Name": pd.Series(dtype="str")}))
    ss.setdefault("master_bytes", None)
    ss.setdefault("loaded_master_id", None)
    ss.setdefault("suggested_start", None)
    ss.setdefault("rota_result", None)
    ss.setdefault("fte_history_base", pd.DataFrame(columns=FTE_HISTORY_COLS))

    tab_settings, tab_generate = st.tabs(["⚙️ Settings", "📅 Generate Rota"])

    # ------------------------------------------------------------------ SETTINGS
    with tab_settings:
        st.subheader("1. Master Excel Rota File")
        st.markdown("<p style='color:#86868b; font-size:14px;'>Returning user? Upload the Master Excel you downloaded last time. "
                    "Shifts, staff, FTEs, assignment history and any leave typed into the Rota sheet are restored.</p>", unsafe_allow_html=True)
        master_file = st.file_uploader("Upload Master Excel Rota File (.xlsx)", type=["xlsx"], key="master_uploader")
        if master_file is not None:
            file_id = (master_file.name, master_file.size)
            if ss.loaded_master_id != file_id:  # load once, so later edits are not overwritten on every rerun
                try:
                    restored = load_master_into_state(master_file.getvalue())
                    ss.loaded_master_id = file_id
                    st.success(f"Restored: {', '.join(restored) if restored else 'history only (no Config sheets found)'}")
                except Exception as e:
                    st.error(f"Could not read Master Excel: {e}")
        elif ss.loaded_master_id is not None:
            ss.loaded_master_id, ss.master_bytes = None, None

        st.markdown("---")
        st.subheader("2. Shift Scheduler")
        st.markdown(
            "<p style='color:#86868b; font-size:14px;'>Choose how many parallel shifts the rota produces (1 = Consultant A only, "
            "2 = Consultant A & B, 3 = Consultant A, B & C), then configure each shift below. "
            "Columns marked with <b>▾</b> have a dropdown or tick-box menu when clicked. "
            "All settings are saved automatically in the <b>Config_Shifts</b> sheet of the Master Excel file.</p>",
            unsafe_allow_html=True,
        )

        base_sh = clean_shifts(ss.shifts_base)
        init_p_n = int(base_sh["Parallel Consultants"].iloc[0]) if not base_sh.empty else 2
        init_grp = str(base_sh["Group Pairing"].iloc[0]) if not base_sh.empty else "Different Group Ideally"
        init_flip_bh = bool(base_sh["Flip Nights on BH"].iloc[0]) if not base_sh.empty else True
        init_mon_off = bool(base_sh["Monday Off After Weekend"].iloc[0]) if not base_sh.empty else False
        init_dual = bool(base_sh["Allow Dual Cover"].iloc[0]) if not base_sh.empty else False

        sc_col1, sc_col2 = st.columns(2)
        with sc_col1:
            sel_parallel_n = st.selectbox(
                "Number of Parallel Shifts to Produce (e.g. Consultant A, B, C) ▾",
                options=[1, 2, 3],
                index=[1, 2, 3].index(init_p_n if init_p_n in (1, 2, 3) else 2),
                key=f"parallel_n_sel_{ss.editor_ver}",
                help="Controls how many parallel shift columns the scheduler produces (1 = single shift, 2 = two shifts e.g. A & B, 3 = three shifts e.g. A, B & C).",
            )
        with sc_col2:
            sel_group_pair = st.selectbox(
                "Group Pairing Between Parallel Shifts ▾",
                options=GROUP_PAIRING_OPTIONS,
                index=GROUP_PAIRING_OPTIONS.index(init_grp if init_grp in GROUP_PAIRING_OPTIONS else "Different Group Ideally"),
                key=f"group_pair_sel_{ss.editor_ver}",
                help="'Different Group Ideally' = soft preference for pairing consultants from different subspecialty groups; 'Same Group Allowed' = no group preference.",
            )

        chk_c1, chk_c2, chk_c3 = st.columns(3)
        with chk_c1:
            sel_flip_bh = st.checkbox(
                "Flip Weekday Nights on Bank Holiday Weeks",
                value=init_flip_bh,
                key=f"flip_bh_chk_{ss.editor_ver}",
                help="Oncology default: flips Mon/Wed and Tue/Thu nights between A and B on weeks containing a Bank Holiday. Untick to keep night days fixed every week.",
            )
        with chk_c2:
            sel_mon_off = st.checkbox(
                "Monday Off After Weekend on Call",
                value=init_mon_off,
                key=f"mon_off_chk_{ss.editor_ver}",
                help="Gives compensatory Monday off from weekday daytime shifts to a consultant who worked the preceding weekend.",
            )
        with chk_c3:
            sel_dual = st.checkbox(
                "Allow Same Consultant on Parallel Shifts",
                value=init_dual,
                key=f"dual_chk_{ss.editor_ver}",
                help="Allows the same consultant to cover two parallel shifts on the same day if needed.",
            )

        # Keep the number of visible shift rows synced with sel_parallel_n (1, 2, or 3)
        default_third_shift = {
            "Shift Name": "Consultant C", "Type": "2-Day Weekend", "Start Day": "Saturday", "Min Rest Weeks": 2,
            "Parallel Consultants": 3, "Nights Mode": "Alternating Consultants", "Nights Covered": "Sat,Sun",
            "Flip Nights on BH": sel_flip_bh, "B Consultant Role": "Independent Parallel Rota",
            "Group Pairing": sel_group_pair, "Allow Dual Cover": sel_dual, "Monday Off After Weekend": sel_mon_off,
        }
        if ss.get("shift_rows_cache_ver") != ss.editor_ver:
            cache_rows = base_sh.to_dict("records")
            while len(cache_rows) < 3:
                if len(cache_rows) == 0:
                    cache_rows.append(dict(DEFAULT_SHIFTS[0]))
                elif len(cache_rows) == 1:
                    cache_rows.append(dict(DEFAULT_SHIFTS[1]))
                else:
                    cache_rows.append(dict(default_third_shift))
            ss.shift_rows_cache = cache_rows[:3]
            ss.shift_rows_cache_ver = ss.editor_ver
            ss.last_sel_parallel_n = sel_parallel_n
        elif ss.get("last_sel_parallel_n") != sel_parallel_n and "latest_config_shifts" in ss:
            prev_edited = ss.latest_config_shifts.to_dict("records")
            for idx_r, r_ed in enumerate(prev_edited):
                if idx_r < len(ss.shift_rows_cache):
                    ss.shift_rows_cache[idx_r] = dict(r_ed)
            ss.last_sel_parallel_n = sel_parallel_n

        rows_sh = [dict(r) for r in ss.shift_rows_cache[:sel_parallel_n]]

        ui_shifts_df = pd.DataFrame(rows_sh)[["Shift Name", "Type", "Start Day", "Min Rest Weeks", "Nights Covered"]].copy()
        ui_shifts_df["Nights Covered"] = [to_day_tags(v) for v in ui_shifts_df["Nights Covered"]]

        edited_shifts_ui = st.data_editor(
            ui_shifts_df, num_rows="fixed", use_container_width=True, key=f"shift_editor_{ss.editor_ver}_{sel_parallel_n}",
            column_config={
                "Shift Name": st.column_config.TextColumn("Shift Name", required=True),
                "Type": st.column_config.SelectboxColumn(
                    "Type (Cover Schedule) ▾", options=SHIFT_TYPES, required=True,
                    help="7-Day Continuous (all 7 days) · 5-Day Weekday (Mon–Fri daytime) · 2-Day Weekend (Sat–Sun) · Nights Only (no daytime column, only covers its ticked nights)."),
                "Start Day": st.column_config.SelectboxColumn("Start Day ▾", options=DAY_NAMES, required=True),
                "Min Rest Weeks": st.column_config.NumberColumn("Min Rest Weeks", min_value=0, max_value=20, step=1, default=3),
                "Nights Covered": st.column_config.MultiselectColumn(
                    "Nights Covered (Tick Days) ▾",
                    options=SHORT_DAYS,
                    help="Click to tick or untick which nights of the week (Mon–Sun) this shift covers.",
                ),
            })

        full_shifts = edited_shifts_ui.copy()
        full_shifts["Parallel Consultants"] = int(sel_parallel_n)
        full_shifts["Flip Nights on BH"] = bool(sel_flip_bh)
        full_shifts["Group Pairing"] = sel_group_pair
        full_shifts["Allow Dual Cover"] = bool(sel_dual)
        full_shifts["Monday Off After Weekend"] = bool(sel_mon_off)
        full_shifts["Nights Covered"] = [from_day_tags(v, empty_str="None") for v in full_shifts["Nights Covered"]]
        full_shifts["B Consultant Role"] = [
            "Just Nights" if t == "Nights Only" else ("Second Parallel Rota + Nights" if nc != "None" else "Independent Parallel Rota")
            for t, nc in zip(full_shifts["Type"], full_shifts["Nights Covered"])
        ]
        full_shifts["Nights Mode"] = "Alternating Consultants"
        config_shifts = clean_shifts(full_shifts)
        ss.latest_config_shifts = config_shifts.copy()

        a_cfg, b_cfg, extras = pick_engine_shifts(config_shifts)
        if a_cfg is None:
            st.warning("Add at least one shift so the engine can fill the rota.")
        else:
            p_n = int(a_cfg.get("Parallel Consultants", 2))
            msg = f"Parallel Shifts: **{p_n}** · Shift 1 → **{a_cfg['Shift Name']}** ({a_cfg['Type']}, starts {a_cfg['Start Day']}, nights: {a_cfg['Nights Covered']})"
            if b_cfg is not None and (p_n >= 2 or b_cfg.get("B Consultant Role") == "Just Nights"):
                msg += f" · Shift 2 → **{b_cfg['Shift Name']}** ({b_cfg['Type']}, nights: {b_cfg['Nights Covered']})"
            if extras and p_n >= 3:
                msg += f" · Shift 3 → **{', '.join(extras)}**"
            st.caption(msg)

        st.markdown("---")
        st.subheader("3. Consultant Roster & FTE Matrix")
        st.markdown(
            "<p style='color:#86868b; font-size:14px;'>Columns marked with <b>▾</b> have a dropdown or tick-box menu when clicked.</p>",
            unsafe_allow_html=True,
        )
        staff_input = sync_staff_fte_columns(ss.staff_base, config_shifts)
        staff_ui = staff_input.copy()
        staff_ui["Rest Days (Mon-Fri)"] = [to_day_tags(v, []) for v in staff_ui["Rest Days (Mon-Fri)"]]
        existing_groups = [safe_raw(g) for g in staff_ui.get("Group", []) if safe_raw(g)]
        group_opts = list(dict.fromkeys(["Solid", "Neuro", "Others"] + existing_groups))
        staff_cfg = {
            "Initials": st.column_config.TextColumn("Initials", required=True),
            "Full Name": st.column_config.TextColumn("Full Name", required=True),
            "Active": st.column_config.CheckboxColumn("Active", default=True),
            "DIL Day": st.column_config.SelectboxColumn("DIL Day ▾", options=["", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]),
            "Group": st.column_config.SelectboxColumn("Group ▾", options=group_opts, default="Others"),
            "Weekdays": st.column_config.CheckboxColumn("Weekdays", default=True, help="Whether this consultant covers weekday shifts."),
            "Nights": st.column_config.SelectboxColumn(
                "Nights ▾", options=STAFF_NIGHTS_OPTIONS, default="Yes",
                help="'Yes' = covers nights with own block and in night pool · 'Separate Nights Only' = daytime only during own block, covers separate nights in other weeks · 'No' = exempt from nights."),
            "Preferred Weeks in a Row": st.column_config.SelectboxColumn(
                "Preferred Weeks in a Row ▾", options=[1, 2, 3, 4], default=1,
                help="Number of consecutive weeks preferred in a row for block shifts (1 to 4)."),
            "Rest Days on Call": st.column_config.CheckboxColumn(
                "Rest Days on Call", default=False,
                help="Enable if this consultant has fixed rest/off day(s) Mon–Fri while on call."),
            "Rest Days (Mon-Fri)": st.column_config.MultiselectColumn(
                "Rest Days (Mon-Fri) ▾",
                options=WEEKDAY_SHORT_DAYS,
                help="Click to tick weekday(s) off while on call (Mon–Fri) when 'Rest Days on Call' is checked.",
            ),
            "On Service From": st.column_config.DateColumn(
                "On Service From", format="DD/MM/YYYY",
                help="First day this consultant is available (e.g. return from long sick leave or a new starter). "
                     "Nothing before this date counts towards fairness, so there is no catch-up. Leave blank if on service all along."),
            "FTE Changed From": st.column_config.DateColumn(
                "FTE Changed From", format="DD/MM/YYYY",
                help="Only fill in after changing an FTE: the date the new FTE starts. Blank = the start of the new rota period. "
                     "Shifts before this date stay measured against the old FTE."),
        }
        for s in shift_names_of(config_shifts):
            staff_cfg[f"FTE {s}"] = st.column_config.NumberColumn(
                f"FTE {s}", min_value=0.0, max_value=1.0, step=0.01, format="%.2f", default=1.0,
                help="0.00 = exempt from this shift (zero exposure); 1.00 = full share.")
        edited_staff = st.data_editor(staff_ui, num_rows="dynamic", use_container_width=True, key=f"staff_editor_{ss.editor_ver}", column_config=staff_cfg)
        config_staff = sync_staff_fte_columns(edited_staff, config_shifts)
        ss.latest_config_staff = config_staff.copy()
        st.caption("**Returning after sick leave / new starter:** set *On Service From* to their first day. They start level with colleagues, "
                   "with no catch-up. **Changing someone's FTE:** edit the FTE column and put the start date in *FTE Changed From*. "
                   "Earlier shifts keep counting at the old FTE; the change is logged in the Master file (sheet Config_FTE_History).")

        st.markdown("---")
        if st.button("Reset to Default Configuration"):
            ss.shifts_base = pd.DataFrame(DEFAULT_SHIFTS)
            ss.staff_base = sync_staff_fte_columns(pd.DataFrame(DEFAULT_STAFF), ss.shifts_base)
            ss.fte_history_base = pd.DataFrame(columns=FTE_HISTORY_COLS)
            ss.editor_ver += 1
            st.rerun()

    # ------------------------------------------------------------------ GENERATE
    with tab_generate:
        st.subheader("🗓️ Calendar & Fairness Settings")
        current_year = datetime.today().year
        def_s = ss.suggested_start or datetime(current_year, 1, 1).date()
        def_e = datetime(def_s.year, 12, 31).date()
        if def_e <= def_s:
            def_e = def_s + timedelta(days=182)
        c_p1, c_p2, c_y1, c_y2 = st.columns(4)
        with c_p1: sel_s = st.date_input("Start New Rota:", value=def_s, format="DD/MM/YYYY")
        with c_p2: sel_e = st.date_input("End New Rota:", value=max(def_e, sel_s), min_value=sel_s, format="DD/MM/YYYY")
        with c_y1: ytd_s = st.date_input("YTD Start Date:", value=datetime(sel_s.year, 1, 1).date(), format="DD/MM/YYYY")
        with c_y2: ytd_e = st.date_input("YTD End Date:", value=datetime(sel_s.year, 12, 31).date(), format="DD/MM/YYYY")

        st.markdown("---")
        st.subheader("Inputs & Overrides")
        st.caption("Leave markers: **AL, SL, NO, OFF** = cannot be rostered and must be caught up later · **SICK** = cannot be rostered, "
                   "no catch-up · **DIL** day in lieu · **REQ / REQB** requests.")
        c1, c2 = st.columns(2)
        with c1: leave_f = st.file_uploader("Upload Separate Leave Requests (Optional)", type=["xlsx"], key="leave_uploader")
        with c2: legacy_f = st.file_uploader("Upload Legacy Previous Rota (only if not using a Master file)", type=["xlsx"], key="legacy_uploader")
        if legacy_f is not None:
            st.caption("Memory source: legacy previous rota upload.")
        elif ss.master_bytes is not None:
            st.caption("Memory source: Master Excel uploaded in Settings.")
        else:
            st.caption("Memory source: none. This rota starts with no history (fairness counts start at zero).")

        st.markdown("##### Locum / Academic Overrides")
        st.markdown("<p style='color:#86868b; font-size:13px;'>Leave 'End Date' blank for a 1-day assignment.</p>", unsafe_allow_html=True)
        edited_locums = st.data_editor(ss.locums_base, num_rows="dynamic", use_container_width=True, key="locum_editor",
                                       column_config={"Start Date": st.column_config.DateColumn("Start Date", format="DD/MM/YYYY"),
                                                      "End Date": st.column_config.DateColumn("End Date", format="DD/MM/YYYY")})
        locum_dict = {}
        for _, row in edited_locums.iterrows():
            name = safe_raw(row.get("Locum Name"))
            if pd.notna(row.get("Start Date")) and name:
                s = pd.to_datetime(row["Start Date"]).date()
                e = pd.to_datetime(row["End Date"]).date() if pd.notna(row.get("End Date")) else s
                while s <= e:
                    locum_dict[s] = name
                    s += timedelta(days=1)

        df_l = pd.DataFrame()
        if leave_f:
            try:
                df_l = pd.read_excel(leave_f)
                df_l.columns = [str(c).strip().upper() for c in df_l.columns]
                d_col = next((c for c in df_l.columns if "DATE" in c), df_l.columns[1] if len(df_l.columns) > 1 else df_l.columns[0])
                df_l[d_col] = pd.to_datetime(df_l[d_col], dayfirst=True, errors="coerce").dt.normalize()
                df_l = df_l.dropna(subset=[d_col]).set_index(d_col)
                df_l = df_l[~df_l.index.duplicated(keep="last")]
            except Exception as e:
                st.error(f"Error loading leave file: {e}")
                df_l = pd.DataFrame()

        st.markdown("---")
        if st.button("Generate Final Rota", type="primary"):
            memory_file = legacy_f if legacy_f is not None else (io.BytesIO(ss.master_bytes) if ss.master_bytes else None)
            try:
                with st.spinner("Solving A-weeks and B-shifts..."):
                    rota, summ, state, prev_rota_df, warns, fte_hist_out = run_rota_v2(
                        df_l, config_staff, config_shifts, sel_s, sel_e, ytd_s, ytd_e, locum_dict, memory_file, ss.fte_history_base)
                ss.rota_result = (rota, summ, state, sel_e, prev_rota_df, warns, config_shifts.copy(), config_staff.copy(), fte_hist_out)
                ss.rota_preview_month_idx = 0
                ss.rota_editor_ver = ss.get("rota_editor_ver", 0) + 1
            except ValueError as e:
                st.error(str(e))
                ss.rota_result = None

        if ss.rota_result:
            @st.fragment
            def render_rota_results_fragment():
                base_rota, base_summ, state, res_e, prev_rota_df, warns, res_shifts, res_staff, res_fte_hist = ss.rota_result
                for w in warns:
                    st.warning(w)

                def highlight(row):
                    dt = parse_any_date(row["Date"])
                    if dt and (dt.weekday() >= 5 or (uk_holidays is not None and dt in uk_holidays)):
                        return ["background-color: #e5e5ea;"] * len(row)
                    return [""] * len(row)

                st.subheader("Rota for the Selected Period (Interactive Editor)")
                st.markdown(
                    "<p style='color:#86868b; font-size:14px;'>Sized to show a full month at a time and continuously scrollable across the entire period without interruption. "
                    "Click any shift or leave cell (marked <b>▾</b>) to modify it directly — the <b>Fairness Calculator</b> "
                    "and the <b>Download Master Excel</b> file update automatically.</p>",
                    unsafe_allow_html=True,
                )

                target = pd.to_datetime(res_e).strftime("%d/%m/%Y")
                idx = base_rota.index[base_rota["Date"] == target]
                base_view = base_rota.iloc[: idx[-1] + 1].copy() if len(idx) else base_rota.copy()
                orig_indices = list(base_view.index)
                base_view_df = base_view.reset_index(drop=True)

                ss.setdefault("rota_editor_ver", 0)

                staff_inis = [safe_ini(x) for x in res_staff.get("Initials", []) if safe_ini(x)]
                staff_names = [safe_raw(x) for x in res_staff.get("Full Name", []) if safe_raw(x)]
                ini_to_fname = {
                    safe_ini(r.get("Initials")).upper(): safe_raw(r.get("Full Name"))
                    for _, r in res_staff.iterrows()
                    if safe_ini(r.get("Initials")) and safe_raw(r.get("Full Name"))
                }
                leave_cols = [c for c in base_view_df.columns if c in set(staff_inis)]
                shift_cols = [c for c in base_view_df.columns if c not in ("Date", "Day") and c not in set(leave_cols)]

                rota_col_cfg = {
                    "Date": st.column_config.TextColumn("Date", disabled=True),
                    "Day": st.column_config.TextColumn("Day", disabled=True),
                }
                for sc_name in shift_cols:
                    existing_vals = [safe_raw(v) for v in base_rota[sc_name].unique() if safe_raw(v)]
                    opts = list(dict.fromkeys([""] + staff_names + existing_vals + staff_inis + ["UNCOVERED"]))
                    rota_col_cfg[sc_name] = st.column_config.SelectboxColumn(
                        f"{sc_name} ▾",
                        options=opts,
                        help="Select a consultant Full Name or Initials (or leave blank / UNCOVERED).",
                    )
                base_leave_opts = ["", "AL", "SL", "DIL", "OFF", "NO", "SICK", "REQ", "REQB"]
                for lc_name in leave_cols:
                    existing_l = [safe_ini(v) for v in base_rota[lc_name].unique() if safe_ini(v)]
                    l_opts = list(dict.fromkeys(base_leave_opts + existing_l))
                    rota_col_cfg[lc_name] = st.column_config.SelectboxColumn(
                        f"{lc_name} ▾",
                        options=l_opts,
                        help=f"Leave / request marker for {ini_to_fname.get(lc_name, lc_name)}.",
                    )

                # Keep base_view_df immutable so st.data_editor never resets scroll position or drops consecutive edits
                editor_height = 1130
                edited_view = st.data_editor(
                    base_view_df.style.apply(highlight, axis=1),
                    use_container_width=True,
                    height=editor_height,
                    num_rows="fixed",
                    disabled=["Date", "Day"],
                    key=f"rota_preview_editor_{ss.rota_editor_ver}",
                    column_config=rota_col_cfg,
                )

                # Build live_rota and live_summ from edited_view without mutating base_rota in ss.rota_result
                live_rota = base_rota.copy()
                has_changes = False
                for sc_name in shift_cols:
                    norm_vals = [ini_to_fname.get(safe_raw(v).upper(), safe_raw(v)) for v in edited_view[sc_name]]
                    old_vals = [safe_raw(v) for v in base_view_df[sc_name]]
                    if norm_vals != old_vals:
                        live_rota.loc[orig_indices, sc_name] = norm_vals
                        has_changes = True
                for lc_name in leave_cols:
                    norm_l = [safe_ini(v) for v in edited_view[lc_name]]
                    old_l = [safe_ini(v) for v in base_view_df[lc_name]]
                    if norm_l != old_l:
                        live_rota.loc[orig_indices, lc_name] = norm_l
                        has_changes = True

                live_summ = recompute_summary_from_rota(live_rota, base_summ, state, prev_rota_df) if has_changes else base_summ
                live_view = live_rota.iloc[: idx[-1] + 1] if len(idx) else live_rota

                uncovered = int(live_view.drop(columns=["Date", "Day"]).isin(["UNCOVERED"]).sum().sum())
                if uncovered:
                    st.error(f"{uncovered} cell(s) UNCOVERED in the selected period. Click the cell(s) above to assign a consultant or relax Min Rest Weeks.")

                st.subheader("Fairness Calculator (Live Summary)")
                st.caption("Updates automatically when you edit shifts or leave in the Rota above, and uses live Excel formulas in the downloaded Master file.")
                st.dataframe(live_summ, use_container_width=True)
                with st.expander("FTE history recorded in the Master file"):
                    st.dataframe(res_fte_hist, use_container_width=True)

                latest_shifts = ss.get("latest_config_shifts", res_shifts)
                latest_staff = ss.get("latest_config_staff", res_staff)
                xbytes = export_master_excel(latest_shifts, latest_staff, live_rota, live_summ, state, prev_rota_df, res_fte_hist)
                st.download_button("Download Master Excel", xbytes.getvalue(), f"Master_Rota_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

            render_rota_results_fragment()


if __name__ == "__main__":
    main()
