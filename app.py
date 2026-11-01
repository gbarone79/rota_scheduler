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
SHIFT_TYPES = ["5-Day Weekday", "2-Day Weekend", "7-Day Continuous"]
STAFF_DATE_COLS = ["On Service From", "FTE Changed From"]
BASE_STAFF_COLS = ["Initials", "Full Name", "Active", "DIL Day", "Group"] + STAFF_DATE_COLS
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
    {"Shift Name": "Consultant A", "Type": "7-Day Continuous", "Start Day": "Friday", "Min Rest Weeks": 3},
    {"Shift Name": "Consultant B", "Type": "5-Day Weekday", "Start Day": "Monday", "Min Rest Weeks": 3},
]

DEFAULT_STAFF = [
    {"Initials": "GB", "Full Name": "Giuseppe Barone", "Active": True, "DIL Day": "Friday", "Group": "Solid", "FTE Consultant A": 1.0, "FTE Consultant B": 1.0},
    {"Initials": "SD", "Full Name": "Sarita Depani", "Active": True, "DIL Day": "Monday", "Group": "Neuro", "FTE Consultant A": 1.0, "FTE Consultant B": 1.0},
    {"Initials": "OS", "Full Name": "Olga Slater", "Active": True, "DIL Day": "Tuesday", "Group": "Solid", "FTE Consultant A": 1.0, "FTE Consultant B": 1.0},
    {"Initials": "MJ", "Full Name": "Mette Jorgensen", "Active": True, "DIL Day": "Friday", "Group": "Neuro", "FTE Consultant A": 1.0, "FTE Consultant B": 1.0},
    {"Initials": "CD", "Full Name": "Cat Duncan", "Active": True, "DIL Day": "Tuesday", "Group": "Solid", "FTE Consultant A": 1.0, "FTE Consultant B": 1.0},
    {"Initials": "KS", "Full Name": "Kavitha Srivatsa", "Active": True, "DIL Day": "Monday", "Group": "Neuro", "FTE Consultant A": 1.0, "FTE Consultant B": 1.0},
    {"Initials": "TC", "Full Name": "Tanzina Chowdhury", "Active": True, "DIL Day": "Tuesday", "Group": "Solid", "FTE Consultant A": 1.0, "FTE Consultant B": 1.0},
    {"Initials": "ES", "Full Name": "Elwira Szychot", "Active": True, "DIL Day": "Friday", "Group": "Neuro", "FTE Consultant A": 1.0, "FTE Consultant B": 1.0},
    {"Initials": "TN", "Full Name": "Trung Nguyen", "Active": True, "DIL Day": "Wednesday", "Group": "Solid", "FTE Consultant A": 1.0, "FTE Consultant B": 1.0},
]


# =============================================================================
# GENERIC HELPERS
# =============================================================================
def to_date(d):
    return d.date() if hasattr(d, 'date') else d


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
def clean_shifts(df_shifts):
    df = df_shifts.copy() if df_shifts is not None else pd.DataFrame()
    for col, default in [("Shift Name", ""), ("Type", "5-Day Weekday"), ("Start Day", "Monday"), ("Min Rest Weeks", 3)]:
        if col not in df.columns:
            df[col] = default
    df = df[["Shift Name", "Type", "Start Day", "Min Rest Weeks"]]
    df["Shift Name"] = df["Shift Name"].apply(safe_raw)
    df = df[df["Shift Name"] != ""]
    df["Type"] = df["Type"].apply(lambda v: safe_raw(v) if safe_raw(v) in SHIFT_TYPES else "5-Day Weekday")
    df["Start Day"] = df["Start Day"].apply(lambda v: safe_raw(v).capitalize() if safe_raw(v).capitalize() in DAY_NAMES else "Monday")
    df["Min Rest Weeks"] = pd.to_numeric(df["Min Rest Weeks"], errors="coerce").fillna(3).astype(int)
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
    for col in BASE_STAFF_COLS:
        if col not in df_staff.columns:
            df_staff[col] = True if col == "Active" else ("Others" if col == "Group" else (pd.NaT if col in STAFF_DATE_COLS else ""))
    for col in STAFF_DATE_COLS:
        df_staff[col] = pd.to_datetime(df_staff[col], errors="coerce")

    fte_cols = [f"FTE {s}" for s in shift_names_of(df_shifts)]
    for fte_col in fte_cols:
        if fte_col not in df_staff.columns:
            # Legacy V1 files used "FTE A"/"FTE B"
            legacy = None
            if fte_col == "FTE Consultant A" and "FTE A" in df_staff.columns: legacy = "FTE A"
            elif fte_col == "FTE Consultant B" and "FTE B" in df_staff.columns: legacy = "FTE B"
            df_staff[fte_col] = pd.to_numeric(df_staff[legacy], errors="coerce").fillna(1.0) if legacy else 1.0
        else:
            df_staff[fte_col] = pd.to_numeric(df_staff[fte_col], errors="coerce").fillna(1.0)

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
    df["Active"] = df["Active"].fillna(True).astype(bool)
    df["Group"] = df["Group"].apply(lambda v: safe_raw(v) or "Others")
    df["DIL Day"] = df["DIL Day"].apply(safe_raw)
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
    """Option 1 mapping: first 7-Day Continuous -> Engine A, first 5-Day Weekday -> Engine B."""
    df = clean_shifts(df_shifts)
    a_cfg = b_cfg = None
    for rec in df.to_dict("records"):
        if a_cfg is None and rec["Type"] == "7-Day Continuous":
            a_cfg = rec
        elif b_cfg is None and rec["Type"] == "5-Day Weekday":
            b_cfg = rec
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
    counts = defaultdict(lambda: {"A_days": 0, "B_days": 0, "Nights": 0, "BH_days": 0, "BH_groups": set(), "AL_days": 0, "SL_days": 0})
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
                    for l_ini, l_mark in rec["leave"].items():  # AL / SL days requested before this period
                        l_mark = safe_ini(l_mark)
                        if l_mark in ("AL", "SL") and served(l_ini, r_date):
                            counts[l_ini][l_mark + "_days"] += 1
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
# CORE ENGINE (V1 final baseline, with A anchor / rest gaps parameterised)
# =============================================================================
def generate_rota(leave_df, consultants, start_date, end_date, hist_counts, hist_master_a, hist_cycle_a, hist_b_q,
                  hist_curr_a, hist_curr_b, hist_last_a_dates, locum_dict, hist_b_dates, ytd_s, ytd_e, hist_last_b_week,
                  a_anchor_wd=4, a_min_gap_days=21, b_bb_gap_days=21,
                  fte_tl_a=None, fte_tl_b=None, on_service=None, level_start=None):
    active_initials = [safe_ini(c.get('Initials')) for c in consultants if c.get('Active', True) and safe_ini(c.get('Initials'))]
    name_map = {safe_ini(c.get('Initials')): c.get('Full Name', '') for c in consultants if safe_ini(c.get('Initials'))}
    ini_by_name = {v: k for k, v in name_map.items()}

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
    for c in consultants:
        ini = safe_ini(c.get('Initials'))
        if c.get('Active', True) and ini:
            raw_dil = safe_raw(c.get('DIL Day'))
            d_map[ini] = raw_dil.capitalize()
            fte_a_map[ini] = max(0.1, safe_float(c.get('FTE A', 1.0)))
            fte_b_map[ini] = max(0.1, safe_float(c.get('FTE B', 1.0)))
            group_map[ini] = (safe_raw(c.get('Group')) or 'Others').capitalize()

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
        return max(0.1, safe_float(val, default))

    fte_a_at = lambda ini, d_date: fte_at(fte_tl_a, ini, d_date, fte_a_map.get(ini, 1.0))
    fte_b_at = lambda ini, d_date: fte_at(fte_tl_b, ini, d_date, fte_b_map.get(ini, 1.0))

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

    def exposure(ini, fte_fn, blocked_markers, weekdays_only):
        if ini in level_start:
            return 0.0
        total, d = 0.0, max(ytd_s_d, on_service.get(ini) or ytd_s_d)
        while d <= hist_end:
            counts_today = (not weekdays_only) or (d.weekday() in B_WORK_DAYS and not (uk_holidays is not None and d in uk_holidays))
            if counts_today and get_eff_marker(d, ini) not in blocked_markers:
                total += fte_fn(ini, d)
            d += timedelta(days=1)
        return total

    exp_a = {i: exposure(i, fte_a_at, NO_EXPOSURE_A, False) for i in active_initials}
    exp_b_hist = {i: exposure(i, fte_b_at, NO_EXPOSURE_B, True) for i in active_initials}
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
            if on_serv(i, d) and get_eff_marker(d, i) not in NO_EXPOSURE_A:
                exp_a_new[i] += fte_a_at(i, d)
        d += timedelta(days=1)
    rate_a_new = ((fin_d - curr_d_start).days + 1) / sum(exp_a_new.values()) if sum(exp_a_new.values()) > 0 else 0.0
    # A standing is kept "per FTE": a 7-day block costs 7 / (FTE at that time)
    sim_a_units = {i: (surplus_a[i] - rate_a_new * exp_a_new[i]) / fte_a_at(i, curr_d_start) for i in active_initials}

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
        week_base_candidates[f] = [c for c in active_initials if on_serv(c, f)]  # leave is soft-filtered via leave_in_week_pen
        week_block_counts[f] = blocked

    prior_a_week = {ini: week_start_for_a_gap(last_date) for ini, last_date in hist_last_a_dates.items() if ini in active_initials and last_date}
    master_order = {ini: idx for idx, ini in enumerate(master_a_seq)}
    last_b_anchor = {ini: week_start_for_a_gap(dt) for ini, dt in hist_last_b_week.items() if ini in active_initials and dt}

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
            local_a_units[req_doc] += max(1, temp[f]['run']) / fte_a_at(req_doc, f)
            if week_has_bh(f):
                bg = get_week_bh_group(f)
                if bg: local_bh_groups[req_doc].add(bg)
            if temp[f]['locked']:
                assigned[f] = req_doc

        unassigned = {f for f in c_fs if f not in assigned}

        def current_candidates(f):
            return [cand for cand in week_base_candidates.get(f, []) if all(abs((f - prev_f).days) >= min_gap_days for prev_f in assigned_by_doc.get(cand, []))]

        def choose_next_week():
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
            gap_tier = 0 if closest_gap >= ideal_gap else (1 if closest_gap >= a_min_gap_days else 2)

            leave_in_week_pen = count_leave_in_week(f, cand)

            future_pressure = 0
            for other_f in unassigned:
                if other_f == f:
                    continue
                remaining = sum(1 for other_cand in week_base_candidates.get(other_f, []) if other_cand != cand and all(abs((other_f - prev_f).days) >= min_gap_days for prev_f in assigned_by_doc.get(other_cand, [])))
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
                local_a_units[cand] += 7 / fte_a_at(cand, f)
                added_bh = get_week_bh_group(f) if week_has_bh(f) else None
                if added_bh:
                    local_bh_groups[cand].add(added_bh)

                if dfs():
                    return True

                if added_bh and not any(other_f != f and assigned.get(other_f) == cand and get_week_bh_group(other_f) == added_bh for other_f in assigned) and added_bh not in hist_counts.get(cand, {}).get("BH_groups", set()):
                    local_bh_groups[cand].discard(added_bh)
                local_a_units[cand] -= 7 / fte_a_at(cand, f)
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

    counts = {ini: {"A_days": hist_counts.get(ini, {}).get("A_days", 0), "B_days": hist_counts.get(ini, {}).get("B_days", 0),
                    "Nights": hist_counts.get(ini, {}).get("Nights", 0), "BH_days": hist_counts.get(ini, {}).get("BH_days", 0),
                    "BH_groups": set(hist_counts.get(ini, {}).get("BH_groups", set()))} for ini in active_initials}
    period_counts = {ini: {"A_days": 0, "B_days": 0, "Nights": 0, "BH_days": 0} for ini in active_initials}

    recent_b_shifts, last_b_week_start = dict(hist_b_dates), dict(hist_last_b_week)
    sim_b_counts = {ini: 0 for ini in active_initials}  # B shifts given in THIS run (history is in surplus_b)
    sim_n_counts = {ini: 0 for ini in active_initials}

    def add_score(doc_ini, s_type, d_date):
        if doc_ini and doc_ini in period_counts:
            if curr_d_start <= d_date <= fin_d: period_counts[doc_ini][s_type] += 1
            if ytd_s_d <= d_date <= ytd_e_d: counts[doc_ini][s_type] += 1
            if s_type == "B_days": sim_b_counts[doc_ini] += 1
            if s_type == "Nights": sim_n_counts[doc_ini] += 1

    # --- B-SHIFT CAPACITY AND FAIRNESS TARGET SETUP ---
    weighted_cap = {ini: 0.0 for ini in active_initials}  # available weekdays on service, weighted by FTE-at-the-time
    test_d = curr_d_start
    while test_d <= fin_d:
        wd = test_d.weekday()
        if wd in B_WORK_DAYS and not (uk_holidays is not None and test_d in uk_holidays):
            for ini in active_initials:
                if on_serv(ini, test_d) and get_eff_marker(test_d, ini) not in NO_EXPOSURE_B:
                    weighted_cap[ini] += fte_b_at(ini, test_d)
        test_d += timedelta(days=1)

    total_weighted_cap = sum(weighted_cap.values())
    b_target_share = {ini: (weighted_cap[ini] / total_weighted_cap if total_weighted_cap > 0 else 1.0 / max(1, len(active_initials))) for ini in active_initials}

    hist_n = {i: hist_counts.get(i, {}).get("Nights", 0) for i in active_initials}
    rate_n = sum(hist_n.values()) / sum(exp_a.values()) if sum(exp_a.values()) > 0 else 0.0
    surplus_n = {i: 0.0 if i in level_start else hist_n[i] - rate_n * exp_a[i] for i in active_initials}
    n_target_share = {i: (exp_a_new[i] / sum(exp_a_new.values()) if sum(exp_a_new.values()) > 0 else 1.0 / max(1, len(active_initials))) for i in active_initials}

    def get_fairness_over_target(c):
        return surplus_b[c] + sim_b_counts[c] - sum(sim_b_counts.values()) * b_target_share[c]

    def get_night_over_target(c):
        return surplus_n[c] + sim_n_counts[c] - sum(sim_n_counts.values()) * n_target_share[c]

    def group_penalty(a_ini, b_ini):
        if not a_ini or not b_ini:
            return 0
        g_a, g_b = group_map.get(a_ini, "Others"), group_map.get(b_ini, "Others")
        return 1 if g_a != "Others" and g_a == g_b else 0

    while curr_d <= fin_d + timedelta(days=365):
        wd = curr_d.weekday()
        row = {"Date": curr_d.strftime("%d/%m/%Y"), "Day": DAY_NAMES[wd], "Consultant A": "", "Consultant B": "", "Nights": ""}

        if curr_d <= fin_d:
            is_bh_today = uk_holidays is not None and curr_d in uk_holidays and wd < 5

            w_start = curr_d - timedelta(days=wd)
            is_bh_week = uk_holidays is not None and any((w_start + timedelta(days=d)) in uk_holidays for d in range(5))
            active_a_nights, active_b_nights = ([0, 2, 4, 5, 6], [1, 3]) if is_bh_week else ([1, 3, 4, 5, 6], [0, 2])

            if wd == a_anchor_wd:
                anc = curr_d
                cur_a = "LOCUM_ACADEMIC" if any(locum_dict.get(anc + timedelta(days=o)) for o in range(7)) else cycle_a_schedule.get(anc, "UNCOVERED")

            blocked_count = sum(1 for c in active_initials if sum(1 for i in range(5) if get_eff_marker(w_start + timedelta(days=i), c) in LEAVE_BLOCK) >= 3)
            high_leave_mode = blocked_count >= 4

            if wd == 0 or cur_b is None or cur_b == "UNCOVERED":
                avail_week = [c for c in active_initials if c != cur_a and on_serv(c, curr_d)
                              and sum(1 for i in range(5) if get_eff_marker(w_start + timedelta(days=i), c) in LEAVE_BLOCK) < 3]
                if not avail_week:
                    avail_week = [c for c in active_initials if c != cur_a and on_serv(c, curr_d) and get_eff_marker(curr_d, c) not in B_DAY_BLOCK]
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
                        # MASTER B-WEEK TUPLE
                        return (is_runaway, monday_guard, bb_pen, leave_in_week, int(round(fot)), g_pen, b_after_a_pen, fot)
                    cur_b = min(avail_week, key=week_score)
                    last_b_week_start[cur_b] = curr_d
                else:
                    cur_b = "UNCOVERED"

            rq_d = next((i for i in active_initials if get_eff_marker(curr_d, i) == "REQ"), None)
            rqb_d = next((i for i in active_initials if get_eff_marker(curr_d, i) == "REQB"), None)
            lc_d = locum_dict.get(curr_d)

            da_n = lc_d if lc_d else (name_map.get(rq_d, rq_d) if rq_d else name_map.get(cur_a, cur_a))
            da_i = (cur_a if cur_a not in ["UNCOVERED", "LOCUM_ACADEMIC"] else None) if not rq_d and not lc_d else rq_d

            row["Consultant A"] = da_n
            if da_i: add_score(da_i, "A_days", curr_d)
            if wd in active_a_nights:
                row["Nights"] = da_n
                if da_i: add_score(da_i, "Nights", curr_d)

            if wd in B_WORK_DAYS:
                if is_bh_today:
                    row["Consultant B"] = ""
                elif rqb_d and rqb_d != da_i:  # explicit request wins over the automatic pick
                    row["Consultant B"] = name_map.get(rqb_d, rqb_d)
                    add_score(rqb_d, "B_days", curr_d)
                    recent_b_shifts[rqb_d] = curr_d
                    if wd in active_b_nights:
                        row["Nights"] = row["Consultant B"]
                        add_score(rqb_d, "Nights", curr_d)
                else:
                    if not cur_b or get_eff_marker(curr_d, cur_b) in B_DAY_BLOCK or cur_b == da_i:
                        # HARD FILTER: Cannot be B if A, or if marked with leave today
                        avail = [i for i in active_initials if i != da_i and on_serv(i, curr_d) and get_eff_marker(curr_d, i) not in B_DAY_BLOCK]
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
                                x_dil = d_map.get(x)
                                misses_other_b_night = x_dil in DAY_NAMES and DAY_NAMES.index(x_dil) in [n for n in active_b_nights if n != wd]
                                if wd in active_b_nights:
                                    # On Wednesday (and other B-night days), prioritize consultants whose DIL is Monday (or the other B-night day)
                                    # and those behind on nights so they catch up on nights
                                    dil_night_catchup = -1 if misses_other_b_night and not_over < 1.0 else 0
                                    needs_nights = 0 if not_over < 0 else 1
                                    return (is_runaway, bb_pen, dil_night_catchup, needs_nights, int(round(not_over)), int(round(fot)), g_pen, b_after_a_pen, fot)
                                else:
                                    # On non-night B days (Tue/Thu/Fri), keep Monday-DIL / night-deficit consultants free for Wednesday and B-weeks
                                    save_for_b_night = 1 if (x_dil in DAY_NAMES and DAY_NAMES.index(x_dil) in active_b_nights and not_over < 1.0) or not_over < -0.5 else 0
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

    # AL / SL days requested, read from the Rota rows exactly as the Excel formulas will count them
    leave_ct = {i: {"AL_p": 0, "SL_p": 0, "AL_y": 0, "SL_y": 0} for i in active_initials}
    for r_row in rota_data:
        r_d = datetime.strptime(r_row["Date"], "%d/%m/%Y").date()
        for i in active_initials:
            mk = r_row.get(i)
            if mk in ("AL", "SL"):
                if curr_d_start <= r_d <= fin_d: leave_ct[i][mk + "_p"] += 1
                if ytd_s_d <= r_d <= ytd_e_d: leave_ct[i][mk + "_y"] += 1

    state = {"A_MASTER_SEQ": ",".join(master_a_seq), "B_QUEUE": ",".join(active_initials),
             "CURRENT_A": final_cur_a or "", "CURRENT_B": final_cur_b or ""}
    summ = [{"Initials": k, "FTE A": fte_a_at(k, fin_d), "FTE B": fte_b_at(k, fin_d),
             "Period A": period_counts[k]["A_days"], "Period B": period_counts[k]["B_days"], "Period Nights": period_counts[k]["Nights"],
             "Period Total (A+N)": period_counts[k]["A_days"] + period_counts[k]["Nights"],
             "YTD A": counts[k]["A_days"], "YTD B": counts[k]["B_days"], "YTD Nights": counts[k]["Nights"],
             "YTD Total (A+N)": counts[k]["A_days"] + counts[k]["Nights"], "YTD BH": counts[k]["BH_days"],
             "Period AL (days)": leave_ct[k]["AL_p"], "Period SL (days)": leave_ct[k]["SL_p"],
             "YTD AL (days)": hist_counts.get(k, {}).get("AL_days", 0) + leave_ct[k]["AL_y"],
             "YTD SL (days)": hist_counts.get(k, {}).get("SL_days", 0) + leave_ct[k]["SL_y"]} for k in active_initials]
    return pd.DataFrame(rota_data), pd.DataFrame(summ), state


# =============================================================================
# V2 ADAPTER: Settings tables -> engine -> named output columns
# =============================================================================
def run_rota_v2(leave_df, staff_df, shifts_df, sel_s, sel_e, ytd_s, ytd_e, locum_dict, memory_file=None, fte_history=None):
    a_cfg, b_cfg, extras = pick_engine_shifts(shifts_df)
    if a_cfg is None or b_cfg is None:
        raise ValueError("Option 1 needs at least one '7-Day Continuous' shift (Engine A) and one '5-Day Weekday' shift (Engine B) in Settings.")
    a_name, b_name = a_cfg["Shift Name"], b_cfg["Shift Name"]

    staff = clean_staff_for_engine(sync_staff_fte_columns(staff_df, shifts_df))
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
    fte_hist_out = update_fte_history(fte_history, staff, shifts_df, sel_s)
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

    rota, summ, state = generate_rota(
        leave, consultants, sel_s, sel_e, mem["counts"], mem["master_a_seq"], mem["cycle_a"], [],
        mem["curr_a"], mem["curr_b"], mem["last_a_dates"], locum_dict, mem["hist_b_dates"], ytd_s, ytd_e, mem["last_b_week"],
        a_anchor_wd=DAY_NAMES.index(a_cfg["Start Day"]),
        a_min_gap_days=max(7, int(a_cfg["Min Rest Weeks"]) * 7),
        b_bb_gap_days=max(0, int(b_cfg["Min Rest Weeks"]) * 7),
        fte_tl_a=build_fte_timeline(fte_hist_out, a_name), fte_tl_b=build_fte_timeline(fte_hist_out, b_name),
        on_service=on_service, level_start=level_start,
    )

    rota = rota.rename(columns={"Consultant A": a_name, "Consultant B": b_name})
    insert_at = list(rota.columns).index(b_name) + 1
    for i, ex in enumerate(extras):
        rota.insert(insert_at + i, ex, "")

    summ = summ.rename(columns={
        "FTE A": f"FTE {a_name}", "FTE B": f"FTE {b_name}",
        "Period A": f"Period {a_name}", "Period B": f"Period {b_name}", "Period Total (A+N)": f"Period Total ({a_name}+Nights)",
        "YTD A": f"YTD {a_name}", "YTD B": f"YTD {b_name}", "YTD Total (A+N)": f"YTD Total ({a_name}+Nights)"})
    name_lookup = {safe_ini(c["Initials"]): c.get("Full Name", "") for c in consultants}
    summ.insert(1, "Full Name", summ["Initials"].map(name_lookup))

    # Worked BEFORE this period (within the YTD window): fixed starting point for the live Excel formulas
    prior = mem["counts"]
    summ[f"Prior YTD {a_name}"] = summ["Initials"].map(lambda i: prior.get(i, {}).get("A_days", 0))
    summ[f"Prior YTD {b_name}"] = summ["Initials"].map(lambda i: prior.get(i, {}).get("B_days", 0))
    summ["Prior YTD Nights"] = summ["Initials"].map(lambda i: prior.get(i, {}).get("Nights", 0))
    summ["Prior YTD BH"] = summ["Initials"].map(lambda i: prior.get(i, {}).get("BH_days", 0))
    summ["Prior YTD AL"] = summ["Initials"].map(lambda i: prior.get(i, {}).get("AL_days", 0))
    summ["Prior YTD SL"] = summ["Initials"].map(lambda i: prior.get(i, {}).get("SL_days", 0))

    state["PERIOD_START"] = to_date(sel_s).strftime("%d/%m/%Y")
    state["PERIOD_END"] = to_date(sel_e).strftime("%d/%m/%Y")
    state["YTD_START"] = to_date(ytd_s).strftime("%d/%m/%Y")
    state["YTD_END"] = to_date(ytd_e).strftime("%d/%m/%Y")
    state["ENGINE_A_SHIFT"], state["ENGINE_B_SHIFT"] = a_name, b_name

    if extras:
        warnings.append(f"Not auto-filled yet (Phase 2 Option 2): {', '.join(extras)}. These columns are left blank for manual entry.")
    return rota, summ, state, mem["prev_rota_df"], warnings, fte_hist_out


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


def _write_live_summary(ws, summary_df, rota_df, state_dict):
    """Overwrite the Summary sheet's count cells with formulas that read the Rota sheet, so manual edits
    to the rota (e.g. swapping people) update the fairness table instantly."""
    a_name, b_name = state_dict.get("ENGINE_A_SHIFT"), state_dict.get("ENGINE_B_SHIFT")
    ca, cb, cn = _col(rota_df, a_name), _col(rota_df, b_name), _col(rota_df, "Nights")
    ps, pe = parse_any_date(state_dict.get("PERIOD_START")), parse_any_date(state_dict.get("PERIOD_END"))
    ys, ye = parse_any_date(state_dict.get("YTD_START")), parse_any_date(state_dict.get("YTD_END"))
    if not all([ca, cb, cn, ps, pe, ys, ye]):
        return False

    sc = {c: get_column_letter(i + 1) for i, c in enumerate(summary_df.columns)}
    needed = [f"Period {a_name}", f"Period {b_name}", "Period Nights", f"Period Total ({a_name}+Nights)",
              f"YTD {a_name}", f"YTD {b_name}", "YTD Nights", f"YTD Total ({a_name}+Nights)", "YTD BH",
              f"Prior YTD {a_name}", f"Prior YTD {b_name}", "Prior YTD Nights", "Prior YTD BH",
              "Period AL (days)", "Period SL (days)", "YTD AL (days)", "YTD SL (days)", "Prior YTD AL", "Prior YTD SL", "Full Name"]
    if any(n not in sc for n in needed):
        return False

    n_rows = len(summary_df)
    last = len(rota_df) + 1
    rng = lambda col: f"Rota!${col}$2:${col}${last}"
    dr = rng("A")

    # Editable date parameters, to the right of the table
    lab_c, val_c = len(summary_df.columns) + 2, len(summary_df.columns) + 3
    vl = get_column_letter(val_c)
    for r, (label, val) in enumerate([("Period start", ps), ("Period end", pe), ("YTD start", ys), ("YTD end", ye)], start=1):
        ws.cell(row=r, column=lab_c, value=label)
        cell = ws.cell(row=r, column=val_c, value=val)
        cell.number_format = "DD/MM/YYYY"
    ws.cell(row=6, column=lab_c, value="Counts are live formulas on the Rota sheet. Edit names there (or the dates above) and this table updates.")
    ws.cell(row=7, column=lab_c, value="'Prior YTD' = shifts worked before this period (from the uploaded history).")

    def cnt(col, r, lo, hi):
        return f'COUNTIFS({rng(col)},$B{r},{dr},">="&${vl}${lo},{dr},"<="&${vl}${hi})'

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
        bh_ws.append(["Date", "Bank Holiday"])
        for d in bh_dates:
            bh_ws.append([d, uk_holidays.get(d)])
        for row in bh_ws.iter_rows(min_row=2, max_col=1):
            row[0].number_format = "DD/MM/YYYY"
        bh_ws.column_dimensions["A"].width = 14
        bh_ws.column_dimensions["B"].width = 28
        bh_rng = f"BH_Dates!$A$2:$A${len(bh_dates) + 1}"

    for r in range(2, n_rows + 2):
        ws[f"{sc[f'Period {a_name}']}{r}"] = f"={cnt(ca, r, 1, 2)}"
        ws[f"{sc[f'Period {b_name}']}{r}"] = f"={cnt(cb, r, 1, 2)}"
        ws[f"{sc['Period Nights']}{r}"] = f"={cnt(cn, r, 1, 2)}"
        ws[f"{sc[f'Period Total ({a_name}+Nights)']}{r}"] = f"={sc[f'Period {a_name}']}{r}+{sc['Period Nights']}{r}"
        ws[f"{sc[f'YTD {a_name}']}{r}"] = f"={sc[f'Prior YTD {a_name}']}{r}+{cnt(ca, r, 3, 4)}"
        ws[f"{sc[f'YTD {b_name}']}{r}"] = f"={sc[f'Prior YTD {b_name}']}{r}+{cnt(cb, r, 3, 4)}"
        ws[f"{sc['YTD Nights']}{r}"] = f"={sc['Prior YTD Nights']}{r}+{cnt(cn, r, 3, 4)}"
        ws[f"{sc[f'YTD Total ({a_name}+Nights)']}{r}"] = f"={sc[f'YTD {a_name}']}{r}+{sc['YTD Nights']}{r}"
        if bh_dates:
            def bh_cnt(col):
                return (f'SUMPRODUCT(COUNTIFS({rng(col)},$B{r},{dr},{bh_rng},'
                        f'{dr},">="&${vl}$3,{dr},"<="&${vl}$4))')
            ws[f"{sc['YTD BH']}{r}"] = f"={sc['Prior YTD BH']}{r}+{bh_cnt(ca)}+{bh_cnt(cb)}"
        else:
            ws[f"{sc['YTD BH']}{r}"] = f"={sc['Prior YTD BH']}{r}"

        l_col = _col(rota_df, summary_df.iloc[r - 2]["Initials"])  # this person's leave-marker column on the Rota sheet
        if l_col:
            for kind in ("AL", "SL"):
                lr = rng(l_col)
                ws[f"{sc[f'Period {kind} (days)']}{r}"] = f'=COUNTIFS({lr},"{kind}",{dr},">="&${vl}$1,{dr},"<="&${vl}$2)'
                ws[f"{sc[f'YTD {kind} (days)']}{r}"] = f'={sc[f"Prior YTD {kind}"]}{r}+COUNTIFS({lr},"{kind}",{dr},">="&${vl}$3,{dr},"<="&${vl}$4)'

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
        config_shifts.to_excel(writer, index=False, sheet_name="Config_Shifts")
        staff_out = config_staff.copy()
        if "FTE Changed From" in staff_out.columns:
            staff_out["FTE Changed From"] = pd.NaT  # a typed change is now recorded in Config_FTE_History
        staff_out.to_excel(writer, index=False, sheet_name="Config_Staff")
        hist_out = normalize_fte_history(fte_history)
        if hist_out.empty:
            hist_out = history_from_staff(config_staff, config_shifts)
        hist_out.to_excel(writer, index=False, sheet_name="Config_FTE_History")

        live = _write_live_summary(writer.sheets["Summary"], summary_df, rota_x, state_dict)
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
                        for cell in col_cells if cell.value is not None and not (s_name == "Summary" and cell.row in (6, 7))]
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
        "<style>.stApp {background-color: #fbfbfd !important;} "
        "p, h1, h2, h3, span, label, div {color: #1d1d1f; font-family: -apple-system, sans-serif;} "
        "div.stButton > button, div.stDownloadButton > button {background-color: #0071e3 !important; color: #fff !important; border-radius: 20px !important;}</style>",
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
        st.subheader("2. Shift Definitions")
        st.markdown("<p style='color:#86868b; font-size:14px;'>The first <b>7-Day Continuous</b> shift uses Engine A (backtracking search); "
                    "the first <b>5-Day Weekday</b> shift uses Engine B. Min Rest Weeks = minimum gap between blocks for the same consultant.</p>", unsafe_allow_html=True)
        edited_shifts = st.data_editor(
            ss.shifts_base, num_rows="dynamic", use_container_width=True, key=f"shift_editor_{ss.editor_ver}",
            column_config={
                "Shift Name": st.column_config.TextColumn("Shift Name", required=True),
                "Type": st.column_config.SelectboxColumn("Type", options=SHIFT_TYPES, required=True),
                "Start Day": st.column_config.SelectboxColumn("Start Day", options=DAY_NAMES, required=True),
                "Min Rest Weeks": st.column_config.NumberColumn("Min Rest Weeks", min_value=0, max_value=20, step=1, default=3),
            })
        config_shifts = clean_shifts(edited_shifts)
        a_cfg, b_cfg, extras = pick_engine_shifts(config_shifts)
        if a_cfg is None or b_cfg is None:
            st.warning("Add at least one '7-Day Continuous' and one '5-Day Weekday' shift so the engines can fill the rota.")
        else:
            msg = f"Engine A → **{a_cfg['Shift Name']}** (starts {a_cfg['Start Day']}) · Engine B → **{b_cfg['Shift Name']}**"
            if extras:
                msg += f" · Not auto-filled yet: {', '.join(extras)}"
            st.caption(msg)

        st.markdown("---")
        st.subheader("3. Consultant Roster & FTE Matrix")
        staff_input = sync_staff_fte_columns(ss.staff_base, config_shifts)
        staff_cfg = {
            "Initials": st.column_config.TextColumn("Initials", required=True),
            "Full Name": st.column_config.TextColumn("Full Name", required=True),
            "Active": st.column_config.CheckboxColumn("Active", default=True),
            "DIL Day": st.column_config.SelectboxColumn("DIL Day", options=["", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]),
            "Group": st.column_config.SelectboxColumn("Group", options=["Solid", "Neuro", "Others"], default="Others"),
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
            staff_cfg[f"FTE {s}"] = st.column_config.NumberColumn(f"FTE {s}", min_value=0.1, max_value=1.0, step=0.01, format="%.2f", default=1.0)
        edited_staff = st.data_editor(staff_input, num_rows="dynamic", use_container_width=True, key=f"staff_editor_{ss.editor_ver}", column_config=staff_cfg)
        config_staff = sync_staff_fte_columns(edited_staff, config_shifts)
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
            except ValueError as e:
                st.error(str(e))
                ss.rota_result = None

        if ss.rota_result:
            rota, summ, state, res_e, prev_rota_df, warns, res_shifts, res_staff, res_fte_hist = ss.rota_result
            for w in warns:
                st.warning(w)

            def highlight(row):
                dt = parse_any_date(row["Date"])
                if dt and (dt.weekday() >= 5 or (uk_holidays is not None and dt in uk_holidays)):
                    return ["background-color: #e5e5ea;"] * len(row)
                return [""] * len(row)

            st.subheader("Rota for the Selected Period")
            target = pd.to_datetime(res_e).strftime("%d/%m/%Y")
            idx = rota.index[rota["Date"] == target]
            view = rota.iloc[: idx[-1] + 1] if len(idx) else rota
            uncovered = int(view.drop(columns=["Date", "Day"]).isin(["UNCOVERED"]).sum().sum())
            if uncovered:
                st.error(f"{uncovered} cell(s) UNCOVERED. Check leave clashes or relax Min Rest Weeks.")
            st.dataframe(view.style.apply(highlight, axis=1), use_container_width=True, height=450)

            st.subheader("Fairness Calculator")
            st.dataframe(summ, use_container_width=True)
            with st.expander("FTE history recorded in the Master file"):
                st.dataframe(res_fte_hist, use_container_width=True)

            xbytes = export_master_excel(res_shifts, res_staff, rota, summ, state, prev_rota_df, res_fte_hist)
            st.download_button("Download Master Excel", xbytes.getvalue(), f"Master_Rota_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")


if __name__ == "__main__":
    main()
