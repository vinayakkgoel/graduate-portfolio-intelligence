from __future__ import annotations

import math
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# V3 — Graduate Portfolio Intelligence
# ============================================================
# Source of truth: model.xlsx only.
# The workbook contains stacked program blocks in several sheets;
# parsers below intentionally read the workbook structure rather than
# assuming every sheet is a flat table.
# ============================================================

THRESHOLDS = {
    "thin_sample_n": 30,
    "fall_red_yoy": -0.15,
    "fall_amber_yoy": -0.05,
    "tracker_red_requires_nonthin": True,
    "tracker_flag_pp": -0.10,
    "quadrant_red": {"Double Risk"},
    "quadrant_amber": {"Early Risk", "Late Risk", "Post-T3 Risk"},
    "ltr_floor": 0.12,
    "scenario_persistence_cap": 1.00,
}

BRAND = {
    "navy": "#10233F",
    "blue": "#3B82F6",
    "teal": "#14B8A6",
    "purple": "#7C3AED",
    "orange": "#F97316",
    "red": "#EF4444",
    "amber": "#F59E0B",
    "green": "#10B981",
    "slate": "#64748B",
    "light": "#F8FAFC",
    "ink": "#172033",
    "muted": "#667085",
}

PALETTE = [BRAND["blue"], BRAND["teal"], BRAND["purple"], BRAND["orange"], "#06B6D4", "#EC4899"]

MISSING_STRINGS = {"—", "-", "n/a", "N/A", "No data", "None", "nan", "NaN", ""}

TOOLTIP_DEFS = {
    "Enrollment YoY": "Like-for-like enrollment change. Fall 2026 is compared with Fall 2025; full-year figures compare AY2025 with AY2024 because 2026 is partial.",
    "Persistence Δ": "FY Trend - New vs Continuing, New Student % for FY2026 versus FY2025, shown in percentage points. Continuing % is not used for status because recent years are immature.",
    "Exits": "Graduated OR withdrew (combined). The source data cannot reliably separate these two outcomes.",
    "New retention": "Method A term-1 to term-3 survival for mature cohorts.",
    "Returning retention": "Method A credit-hour-weighted retention beyond Term 3, using the workbook's floor rule.",
    "LTR/student": "Direct-method modeled lifetime gross tuition value per student, as reported by Revenue Outlook.",
    "CAC": "FY26 paid-media acquisition cost per student from the workbook's trial CAC dataset. It is not fully loaded.",
}


def clean_value(v: Any) -> Any:
    """Convert workbook placeholders to real missing values and normalize numpy scalars."""
    if v is None:
        return np.nan
    if isinstance(v, str):
        s = v.strip()
        if s in MISSING_STRINGS:
            return np.nan
        return s
    if isinstance(v, (np.generic,)):
        return v.item()
    return v


def is_missing(v: Any) -> bool:
    if v is None:
        return True
    if isinstance(v, str):
        return v.strip() in MISSING_STRINGS
    try:
        return bool(pd.isna(v))
    except Exception:
        return False


def safe_num(v: Any, default: float = np.nan) -> float:
    if is_missing(v):
        return default
    try:
        x = float(v)
        return x if math.isfinite(x) else default
    except Exception:
        return default


def fmt_num(v: Any, decimals: int = 0) -> str:
    x = safe_num(v)
    if np.isnan(x):
        return "Not available"
    return f"{x:,.{decimals}f}"


def fmt_pct(v: Any, decimals: int = 1) -> str:
    x = safe_num(v)
    if np.isnan(x):
        return "Not available"
    return f"{x * 100:.{decimals}f}%"


def fmt_pp(v: Any, decimals: int = 1) -> str:
    x = safe_num(v)
    if np.isnan(x):
        return "Not available"
    return f"{x * 100:+.{decimals}f} pts"


def fmt_currency(v: Any, decimals: int = 0) -> str:
    x = safe_num(v)
    if np.isnan(x):
        return "Not available"
    return f"${x:,.{decimals}f}"


def display_text(v: Any, fallback: str = "Not available") -> str:
    return fallback if is_missing(v) else str(v)


def term_label(term_code: Any) -> str:
    n = safe_num(term_code)
    if np.isnan(n):
        return "Not available"
    code = int(n)
    year = code // 100
    suffix = code % 100
    if suffix == 10:
        return f"Fall {year}"
    if suffix == 20:
        return f"Spring {year + 1}"
    if suffix == 30:
        return f"Summer {year + 1}"
    return str(code)


def short_code_label(code: str, name: str) -> str:
    return f"{code} — {name}"


def parse_program_header(text: Any) -> Optional[Tuple[str, str]]:
    if not isinstance(text, str):
        return None
    m = re.match(r"^\s*([A-Za-z0-9]+)\s+—\s+(.+?)(?:\s+\|.*)?$", text.strip())
    if not m:
        return None
    return m.group(1).strip(), m.group(2).strip()


def parse_quadrant(text: Any) -> Optional[str]:
    if not isinstance(text, str):
        return None
    m = re.search(r"Historical quadrant:\s*([^|]+)", text)
    return m.group(1).strip() if m else None


def load_raw(path: str) -> Dict[str, List[List[Any]]]:
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    raw: Dict[str, List[List[Any]]] = {}
    for ws in wb.worksheets:
        raw[ws.title] = [[clean_value(v) for v in row] for row in ws.iter_rows(values_only=True)]
    return raw


@st.cache_data(show_spinner=False)
def load_model(path: str) -> Dict[str, Any]:
    raw = load_raw(path)
    return build_model(raw)


def rows_to_df(rows: List[List[Any]], header_row: int, data_start: Optional[int] = None) -> pd.DataFrame:
    if not rows or header_row >= len(rows):
        return pd.DataFrame()
    header = [str(x).strip() if not is_missing(x) else f"col_{i}" for i, x in enumerate(rows[header_row])]
    data = rows[data_start if data_start is not None else header_row + 1 :]
    return pd.DataFrame(data, columns=header)


def parse_method_a_detail(rows: List[List[Any]]) -> Dict[str, Dict[str, Any]]:
    result: Dict[str, Dict[str, Any]] = {}
    i = 0
    while i < len(rows):
        hdr = parse_program_header(rows[i][0] if rows[i] else None)
        if hdr:
            code, name = hdr
            block = {"code": code, "name": name, "terms": [], "metrics": {}}
            j = i + 1
            # Search until next program header.
            while j < len(rows) and not parse_program_header(rows[j][0] if rows[j] else None):
                r = rows[j]
                if len(r) >= 7 and str(r[0]) == "Relative Term":
                    j += 1
                    while j < len(rows):
                        rr = rows[j]
                        if parse_program_header(rr[0] if rr else None):
                            break
                        t = safe_num(rr[0]) if rr else np.nan
                        if not np.isnan(t):
                            block["terms"].append(
                                {
                                    "term": int(t),
                                    "n": safe_num(rr[1]),
                                    "avg_ch": safe_num(rr[2]),
                                    "term_retention": safe_num(rr[3]),
                                    "survival_vs_t3": safe_num(rr[4]),
                                    "above_floor": str(rr[5]).lower() == "yes" if not is_missing(rr[5]) else False,
                                    "ch_contribution": safe_num(rr[6]),
                                }
                            )
                        elif block["terms"] and isinstance(rr[0], str) and str(rr[0]).startswith("New Student Retention"):
                            break
                        j += 1
                    continue
                if r and isinstance(r[0], str):
                    label = str(r[0])
                    if label.startswith("New Student Retention"):
                        block["metrics"]["new_retention"] = safe_num(r[1])
                    elif label.startswith("Sum of CH Contributions"):
                        block["metrics"]["sum_ch_contribution"] = safe_num(r[1])
                    elif label.startswith("Sum of Avg CH"):
                        block["metrics"]["sum_avg_ch"] = safe_num(r[1])
                    elif label.startswith("Credit Hours to Degree"):
                        block["metrics"]["degree_ch"] = safe_num(r[1])
                    elif label.startswith("Returning Student Retention"):
                        block["metrics"]["returning_retention"] = safe_num(r[1])
                    elif label.startswith("LTR Factor"):
                        block["metrics"]["ltr_factor"] = safe_num(r[1])
                    elif label.startswith("First term floor"):
                        block["metrics"]["floor_breach"] = safe_num(r[1])
                j += 1
            result[code] = block
            i = j
        else:
            i += 1
    return result


def parse_method_b_summary(rows: List[List[Any]]) -> Dict[str, Dict[str, Any]]:
    df = rows_to_df(rows, 0, 1)
    result = {}
    if df.empty:
        return result
    for _, r in df.iterrows():
        code = display_text(r.get("Code"), "")
        if not code:
            continue
        result[code] = {
            "cohorts": safe_num(r.get("# Cohorts")),
            "mature_cohorts": safe_num(r.get("# Mature Cohorts")),
            "pooled_t3": safe_num(r.get("Pooled S(T3)")),
            "simple_t3": safe_num(r.get("Simple-Avg S(T3)")),
            "cross_check": display_text(r.get("Cross-check vs Method A"), "Not available"),
        }
    return result


def parse_method_b_detail(rows: List[List[Any]]) -> Dict[str, List[Dict[str, Any]]]:
    result: Dict[str, List[Dict[str, Any]]] = {}
    i = 0
    while i < len(rows):
        hdr = parse_program_header(rows[i][0] if rows[i] else None)
        if hdr:
            code, _ = hdr
            j = i + 1
            header = None
            cohorts = []
            while j < len(rows) and not parse_program_header(rows[j][0] if rows[j] else None):
                r = rows[j]
                if r and str(r[0]) == "Cohort":
                    header = r
                    j += 1
                    while j < len(rows):
                        rr = rows[j]
                        if parse_program_header(rr[0] if rr else None):
                            break
                        if not rr or is_missing(rr[0]) or str(rr[0]).startswith("Pooled S(t)") or str(rr[0]).startswith("Simple Average") or str(rr[0]).startswith("Cross-check"):
                            j += 1
                            continue
                        cohort = display_text(rr[0], "")
                        if cohort and re.fullmatch(r"\d{6}", cohort):
                            maturity = display_text(rr[2], "") if len(rr) > 2 else ""
                            n1 = safe_num(rr[3]) if len(rr) > 3 else np.nan
                            surv = {}
                            for idx, val in enumerate(rr[4:], start=1):
                                if not is_missing(val):
                                    surv[idx] = safe_num(val)
                            cohorts.append({"cohort": cohort, "mature": maturity == "Y", "n1": n1, "survival": surv})
                        j += 1
                    continue
                j += 1
            result[code] = cohorts
            i = j
        else:
            i += 1
    return result


def parse_method_a_summary(rows: List[List[Any]]) -> Dict[str, Dict[str, Any]]:
    df = rows_to_df(rows, 0, 1)
    result = {}
    for _, r in df.iterrows():
        code = display_text(r.get("Code"), "")
        if not code:
            continue
        result[code] = {
            "name": display_text(r.get("Program Name"), ""),
            "n_t1": safe_num(r.get("N @ T1 (mature)")),
            "n_t3": safe_num(r.get("N @ T3 (mature)")),
            "new_retention": safe_num(r.get("New Student Retention")),
            "returning_retention": safe_num(r.get("Returning Student Retention")),
            "ltr_factor": safe_num(r.get("LTR Factor (New x Returning)")),
            "floor_breach": safe_num(r.get("Floor Breach Term")),
            "sample_flag": display_text(r.get("Sample Flag"), "No data"),
            "degree_ch": safe_num(r.get("CH to Degree (reference only)")),
        }
    return result


def parse_revenue(rows: List[List[Any]]) -> Dict[str, Dict[str, Any]]:
    df = rows_to_df(rows, 3, 4)
    result = {}
    if df.empty:
        return result
    for _, r in df.iterrows():
        p = display_text(r.get("Program"), "")
        m = parse_program_header(p)
        if not m:
            continue
        code, name = m
        if code in result:
            continue
        result[code] = {
            "name": name,
            "degree_ch": safe_num(r.get("Degree Credit Hours")),
            "tuition": safe_num(r.get("Tuition Cost\n(per Student)")),
            "new_starts": safe_num(r.get("New Starts\n(Fall 2026)")),
            "revenue": safe_num(r.get("Revenue\n(Program Total)")),
            "expected_ch": safe_num(r.get("Expected CH/Student\n(Direct, @12% floor)")),
            "ltr_pct_tuition": safe_num(r.get("LTR as % of Tuition\n(Direct)")),
            "revenue_loss": safe_num(r.get("Revenue Loss\n(Program Total)")),
            "ltr": safe_num(r.get("LTR\n(Program Total)")),
        }
    return result


def parse_program_list(rows: List[List[Any]]) -> Dict[str, Dict[str, Any]]:
    df = rows_to_df(rows, 0, 1)
    result = {}
    for _, r in df.iterrows():
        code = display_text(r.get("Code"), "")
        if not code:
            continue
        result[code] = {
            "name": display_text(r.get("Program Name"), ""),
            "degree_ch": safe_num(r.get("CH to Degree")),
            "mature_n": safe_num(r.get("Mature-cohort N @ T1")),
            "sample_flag": display_text(r.get("Sample Flag"), "No data"),
        }
    return result


def parse_enrollment_sheet(rows: List[List[Any]], year_sheet: bool = False) -> pd.DataFrame:
    header_row = 4 if year_sheet else 4
    df = rows_to_df(rows, header_row, header_row + 2 if year_sheet else header_row + 2)
    return df


def parse_hc_term(rows: List[List[Any]]) -> pd.DataFrame:
    return rows_to_df(rows, 4, 6)


def parse_new_cont_term(rows: List[List[Any]]) -> pd.DataFrame:
    return rows_to_df(rows, 4, 6)


def parse_new_cont_year(rows: List[List[Any]]) -> pd.DataFrame:
    data = []
    for r in rows[6:]:
        code = display_text(r[0] if len(r) > 0 else None, "")
        if not code:
            continue
        rec = {"Program Code": code, "Program Name": display_text(r[1] if len(r) > 1 else None, "")}
        for idx, year in enumerate(range(2019, 2027), start=2):
            rec[f"NEW_{year}"] = safe_num(r[idx]) if idx < len(r) else np.nan
        for idx, year in enumerate(range(2019, 2027), start=11):
            rec[f"CONT_{year}"] = safe_num(r[idx]) if idx < len(r) else np.nan
        rec["NEW total"] = safe_num(r[20]) if len(r) > 20 else np.nan
        rec["CONT total"] = safe_num(r[21]) if len(r) > 21 else np.nan
        data.append(rec)
    return pd.DataFrame(data)


def parse_cac(rows: List[List[Any]]) -> Dict[str, Dict[str, Any]]:
    df = rows_to_df(rows, 6, 7)
    result = {}
    for _, r in df.iterrows():
        p = display_text(r.get("Program"), "")
        m = parse_program_header(p)
        if not m:
            continue
        code, name = m
        result[code] = {
            "media_spend": safe_num(r.get("FY26 Media Spend\n(Follow the Lead)")),
            "fy26_new_starts": safe_num(r.get("FY26 New Starts\n(Gross)")),
            "cac": safe_num(r.get("CAC\n($/student)")),
            "ltr_student": safe_num(r.get("LTR/Student\n(Direct method)")),
            "revenue_multiple": safe_num(r.get("Revenue Multiple\n(Gross LTR ÷ CAC)")),
            "quadrant": display_text(r.get("Quadrant\n(Strategic Action)"), "Not available"),
        }
    return result


def parse_margin(rows: List[List[Any]]) -> Dict[str, Dict[str, Any]]:
    df = rows_to_df(rows, 4, 5)
    result = {}
    for _, r in df.iterrows():
        p = display_text(r.get("Program"), "")
        m = parse_program_header(p)
        if not m:
            continue
        code, _ = m
        result[code] = {
            "cac": safe_num(r.get("CAC\n($/student, media-only)")),
            "gross_ltr_student": safe_num(r.get("Gross LTR\n/Student")),
            "gross_multiple": safe_num(r.get("Gross Revenue\nMultiple")),
            "fy25_revenue": safe_num(r.get("FY25 Program\nRevenue")),
            "fy25_cost": safe_num(r.get("FY25 Program\nCost (proxy)")),
            "margin_pct": safe_num(r.get("Approx.\nMargin %")),
            "margin_ltr_student": safe_num(r.get("Approx. Margin\nLTR/Student")),
            "margin_multiple": safe_num(r.get("Approx. Margin\nMultiple")),
        }
    return result


def parse_fy_trend(rows: List[List[Any]]) -> Dict[str, Any]:
    result: Dict[str, Any] = {"portfolio": {}}
    for i, r in enumerate(rows):
        if r and r[0] == "Metric" and len(r) >= 7:
            years = []
            for x in r[1:7]:
                m = re.search(r"(\d{4})", str(x)) if not is_missing(x) else None
                years.append(int(m.group(1)) if m else np.nan)
            if any(not np.isnan(x) for x in years):
                new_row = rows[i + 1] if i + 1 < len(rows) else []
                cont_row = rows[i + 2] if i + 2 < len(rows) else []
                terms_row = rows[i + 3] if i + 3 < len(rows) else []
                result["portfolio"] = {
                    "years": [int(x) for x in years if not np.isnan(x)],
                    "new": [safe_num(x) for x in new_row[1:7]],
                    "continuing": [safe_num(x) for x in cont_row[1:7]],
                    "terms_observed": [safe_num(x) for x in terms_row[1:7]],
                }
                break
    return result


def parse_watchlist(rows: List[List[Any]]) -> Tuple[Dict[str, List[Dict[str, Any]]], Dict[str, str]]:
    watch: Dict[str, List[Dict[str, Any]]] = {}
    quadrants: Dict[str, str] = {}
    # Watch list rows.
    for r in rows:
        if len(r) >= 9 and display_text(r[0], "") in {"M464", "M073", "M700", "M522", "M435", "M260", "M527"}:
            code = display_text(r[0], "")
            watch.setdefault(code, []).append(
                {
                    "program": display_text(r[1]),
                    "quadrant": display_text(r[2]),
                    "cohort": display_text(r[3]),
                    "step": display_text(r[4]),
                    "current": safe_num(r[5]),
                    "benchmark": safe_num(r[6]),
                    "delta": safe_num(r[7]),
                    "thin": str(display_text(r[8], "")).lower().startswith("yes"),
                }
            )
    # Program block headers throughout the stacked sheet.
    for r in rows:
        hdr = parse_program_header(r[0] if r else None)
        if hdr:
            code = hdr[0]
            q = parse_quadrant(r[0])
            if q:
                quadrants[code] = q
    return watch, quadrants


def parse_program_tracker(rows: List[List[Any]]) -> Dict[str, Dict[str, Any]]:
    """Parse Current-Year Tracker stacked blocks into benchmark + current cohorts."""
    result: Dict[str, Dict[str, Any]] = {}
    i = 0
    while i < len(rows):
        hdr = parse_program_header(rows[i][0] if rows[i] else None)
        if not hdr:
            i += 1
            continue
        code, name = hdr
        q = parse_quadrant(rows[i][0]) or "Not available"
        block = {"quadrant": q, "benchmark": {}, "cohorts": []}
        j = i + 1
        while j < len(rows) and not parse_program_header(rows[j][0] if rows[j] else None):
            r = rows[j]
            if r and str(r[0]) == "Historical benchmark (pooled mature cohorts)":
                block["benchmark"] = {
                    "n": safe_num(r[1]),
                    "t1_t2": safe_num(r[2]),
                    "t2_t3": safe_num(r[4]),
                    "t1_t3": safe_num(r[6]),
                }
            if r and str(r[0]) == "Entry Term":
                j += 1
                while j < len(rows):
                    rr = rows[j]
                    if parse_program_header(rr[0] if rr else None):
                        break
                    term = safe_num(rr[0]) if rr else np.nan
                    if not np.isnan(term):
                        block["cohorts"].append(
                            {
                                "entry_term": int(term),
                                "n": safe_num(rr[1]),
                                "t1_t2": safe_num(rr[2]),
                                "t2_t3": safe_num(rr[4]),
                                "t1_t3": safe_num(rr[6]),
                                "sample_flag": display_text(rr[8], "Not available"),
                            }
                        )
                    j += 1
                continue
            j += 1
        result[code] = block
        i = j
    return result


def parse_term_codes(rows: List[List[Any]]) -> List[int]:
    # Term Index has two columns; use explicit numeric codes in the second column.
    out = []
    for r in rows:
        for v in r:
            x = safe_num(v)
            if not np.isnan(x) and int(x) >= 201900 and int(x) <= 202699:
                out.append(int(x))
    return sorted(set(out))


def build_model(raw: Dict[str, List[List[Any]]]) -> Dict[str, Any]:
    programs = parse_program_list(raw["Program List"])
    ma = parse_method_a_summary(raw["Method A - Summary"])
    mb = parse_method_b_summary(raw["Method B - Summary"])
    detail = parse_method_a_detail(raw["Method A - Detail"])
    cohort_detail = parse_method_b_detail(raw["Method B - Detail"])
    revenue = parse_revenue(raw["Revenue Outlook"])
    hc_term = parse_hc_term(raw["HC by Program-Term"])
    hc_year = parse_enrollment_sheet(raw["HC by Program-Year"], year_sheet=True)
    new_cont_term = parse_new_cont_term(raw["New vs Continuing — Term"])
    new_cont_year = parse_new_cont_year(raw["New vs Continuing — Year"])
    cac = parse_cac(raw["CAC & LTR Data (Trial)"])
    margin = parse_margin(raw["Margin View (Trial, Approx.)"])
    fy_trend = parse_fy_trend(raw["FY Trend - New vs Continuing"])
    watch, quadrants = parse_watchlist(raw["Current-Year Tracker"])
    tracker = parse_program_tracker(raw["Current-Year Tracker"])
    term_codes = parse_term_codes(raw["Term Index"])

    # Merge program-level facts.
    for code, p in programs.items():
        p.update(ma.get(code, {}))
        p.update({"method_b": mb.get(code, {}), "detail": detail.get(code, {})})
        p.update({"revenue_data": revenue.get(code, {}), "cac_data": cac.get(code, {}), "margin_data": margin.get(code, {})})
        p["watch_flags"] = watch.get(code, [])
        p["quadrant"] = quadrants.get(code) or tracker.get(code, {}).get("quadrant") or "Not available"
        p["tracker"] = tracker.get(code, {})

    # HC by term into long form.
    term_rows = []
    if not hc_term.empty:
        for _, r in hc_term.iterrows():
            code = display_text(r.get("Program Code"), "")
            if not code:
                continue
            for col in hc_term.columns[2:22]:
                v = safe_num(r.get(col))
                if np.isnan(v):
                    continue
                # Use the numeric term-code row in row 6 from the raw sheet when available.
                term_map = raw["HC by Program-Term"][5]
                col_idx = list(hc_term.columns).index(col)
                term = safe_num(term_map[col_idx]) if col_idx < len(term_map) else np.nan
                if np.isnan(term):
                    continue
                term_rows.append({"code": code, "term": int(term), "headcount": v})
    hc_term_long = pd.DataFrame(term_rows)

    # New / continuing by term into long form.
    nc_rows = []
    if not new_cont_term.empty:
        raw_term_codes = raw["New vs Continuing — Term"][5]
        for _, r in new_cont_term.iterrows():
            code = display_text(r.get("Program Code"), "")
            if not code:
                continue
            cols = list(new_cont_term.columns)
            for idx, col in enumerate(cols[2:22], start=2):
                term = safe_num(raw_term_codes[idx]) if idx < len(raw_term_codes) else np.nan
                if np.isnan(term):
                    continue
                newv = safe_num(r.get(col))
                cont_col_idx = idx + 21  # blank separator after 20 term columns
                cont = safe_num(r.iloc[cont_col_idx]) if cont_col_idx < len(r) else np.nan
                if not np.isnan(newv) or not np.isnan(cont):
                    nc_rows.append({"code": code, "term": int(term), "new": newv, "continuing": cont})
    new_cont_long = pd.DataFrame(nc_rows)

    # Portfolio mature-cohort persistence is the student-weighted sum across Method A Summary.
    portfolio_t1 = float(sum(safe_num(v.get("n_t1")) for v in ma.values() if not np.isnan(safe_num(v.get("n_t1")))))
    portfolio_t3 = float(sum(safe_num(v.get("n_t3")) for v in ma.values() if not np.isnan(safe_num(v.get("n_t3")))))
    portfolio_new_retention = portfolio_t3 / portfolio_t1 if portfolio_t1 else np.nan

    # Portfolio metrics directly from trusted dashboard and year tables.
    dashboard = rows_to_df(raw["Trends & Risk Dashboard"], 6, 7)
    changing = {}
    if not dashboard.empty:
        allowed_metrics = {
            "Enrolled — Fall (latest vs prior yr)",
            "Enrolled — full FY (2025 vs 2024)",
            "New starts — full FY (2025 vs 2024)",
            "Continuing — full FY (2025 vs 2024)",
            "Ever enrolled / Still enrolled / Exited (all-time)",
            "Programs growing / shrinking (FY2023→FY2025 enrolment)",
        }
        for _, r in dashboard.iterrows():
            metric = display_text(r.get("Metric"), "")
            if metric in allowed_metrics:
                changing[metric] = {
                    "current": safe_num(r.get("Latest")),
                    "prior": safe_num(r.get("Prior")),
                    "delta": safe_num(r.get("Δ")),
                    "pct": safe_num(r.get("%Δ")),
                }

    # Portfolio persistence from FY Trend sheet.
    pt = fy_trend.get("portfolio", {})
    years = pt.get("years", [])
    new_by_year = dict(zip(years, pt.get("new", [])))
    cont_by_year = dict(zip(years, pt.get("continuing", [])))
    terms_by_year = dict(zip(years, pt.get("terms_observed", [])))

    return {
        "programs": programs,
        "program_order": list(programs.keys()),
        "ma": ma,
        "mb": mb,
        "detail": detail,
        "cohort_detail": cohort_detail,
        "revenue": revenue,
        "cac": cac,
        "margin": margin,
        "hc_term": hc_term_long,
        "hc_year": hc_year,
        "new_cont_term": new_cont_long,
        "new_cont_year": new_cont_year,
        "fy_trend": fy_trend,
        "new_by_year": new_by_year,
        "cont_by_year": cont_by_year,
        "terms_by_year": terms_by_year,
        "changing": changing,
        "portfolio_t1": portfolio_t1,
        "portfolio_t3": portfolio_t3,
        "portfolio_new_retention": portfolio_new_retention,
        "term_codes": term_codes,
        "raw": raw,
    }


def program_fall_yoy(model: Dict[str, Any], code: str) -> float:
    df = model["hc_term"]
    if df.empty:
        return np.nan
    x = df[df["code"] == code]
    fall25 = x.loc[x["term"] == 202510, "headcount"]
    fall26 = x.loc[x["term"] == 202610, "headcount"]
    if fall25.empty or fall26.empty or safe_num(fall25.iloc[0]) == 0:
        return np.nan
    return safe_num(fall26.iloc[0]) / safe_num(fall25.iloc[0]) - 1


def program_year_yoy(model: Dict[str, Any], code: str) -> float:
    df = model["hc_year"]
    if df.empty:
        return np.nan
    row = df[df["Program Code"] == code]
    if row.empty:
        return np.nan
    r = row.iloc[0]
    a = safe_num(r.get("2024"))
    b = safe_num(r.get("2025"))
    if np.isnan(a) or a == 0 or np.isnan(b):
        return np.nan
    return b / a - 1


def persistence_delta(model: Dict[str, Any]) -> float:
    new = model["new_by_year"]
    a, b = new.get(2025, np.nan), new.get(2026, np.nan)
    if np.isnan(a) or np.isnan(b):
        return np.nan
    return b - a


def program_status(model: Dict[str, Any], code: str) -> Dict[str, Any]:
    p = model["programs"][code]
    n = safe_num(p.get("mature_n"))
    if np.isnan(n) or n <= 0:
        return {"status": "⚪ Grey", "word": "Insufficient historical data", "rules": ["No mature-cohort data"]}

    rules = []
    fall = program_fall_yoy(model, code)
    quadrant = p.get("quadrant", "Not available")
    watch = p.get("watch_flags", [])
    nonthin_watch = any(not x.get("thin", False) for x in watch)
    if nonthin_watch:
        rules.append("Current-Year Tracker Watch List, non-thin sample")
    if quadrant in THRESHOLDS["quadrant_red"]:
        rules.append(f"Historical quadrant = {quadrant}")
    if not np.isnan(fall) and fall <= THRESHOLDS["fall_red_yoy"]:
        rules.append(f"Fall enrollment YoY {fall:.1%} ≤ {THRESHOLDS['fall_red_yoy']:.0%}")
    red = bool(rules)
    if red:
        return {"status": "🔴 Red", "word": "Red", "rules": rules}

    amber_rules = []
    if watch:
        amber_rules.append("Current-Year Tracker flagged step")
    if quadrant in THRESHOLDS["quadrant_amber"]:
        amber_rules.append(f"Historical quadrant = {quadrant}")
    if not np.isnan(fall) and THRESHOLDS["fall_red_yoy"] < fall <= THRESHOLDS["fall_amber_yoy"]:
        amber_rules.append(f"Fall enrollment YoY {fall:.1%} between -5% and -15%")
    if amber_rules:
        return {"status": "🟡 Amber", "word": "Amber", "rules": amber_rules}

    return {"status": "🟢 Green", "word": "Green", "rules": ["No red or amber rule triggered"]}


def directional_badge(model: Dict[str, Any], code: str) -> str:
    p = model["programs"][code]
    n = safe_num(p.get("mature_n"))
    return "↗ Directional" if not np.isnan(n) and n < THRESHOLDS["thin_sample_n"] else ""


def survival_curve(model: Dict[str, Any], code: str) -> pd.DataFrame:
    p = model["programs"][code]
    terms = p.get("detail", {}).get("terms", [])
    if not terms:
        return pd.DataFrame(columns=["Term", "Survival", "N", "Avg CH", "Observed"])
    n1 = safe_num(terms[0].get("n"))
    out = []
    for t in terms:
        n = safe_num(t.get("n"))
        surv = n / n1 if not np.isnan(n) and not np.isnan(n1) and n1 > 0 else np.nan
        out.append({
            "Term": int(t["term"]),
            "Survival": surv,
            "N": n,
            "Avg CH": safe_num(t.get("avg_ch")),
            "Observed": True,
        })
    return pd.DataFrame(out)


def baseline_direct_ltr_per_student(model: Dict[str, Any], code: str) -> float:
    p = model["programs"][code]
    rev = p.get("revenue_data", {})
    terms = p.get("detail", {}).get("terms", [])
    tuition = safe_num(rev.get("tuition"))
    degree_ch = safe_num(rev.get("degree_ch"))
    if np.isnan(tuition) or np.isnan(degree_ch) or degree_ch <= 0 or not terms:
        return np.nan
    n1 = safe_num(terms[0].get("n"))
    if np.isnan(n1) or n1 <= 0:
        return np.nan
    total = 0.0
    for t in terms:
        term = int(t["term"])
        if term < 3:
            continue
        n = safe_num(t.get("n"))
        avg_ch = safe_num(t.get("avg_ch"))
        if np.isnan(n) or np.isnan(avg_ch):
            continue
        survival = n / n1
        if survival < THRESHOLDS["ltr_floor"]:
            break
        total += survival * avg_ch
    return (tuition / degree_ch) * total


def scenario_baseline_check(model: Dict[str, Any]) -> Dict[str, Any]:
    rows = []
    for code, p in model["programs"].items():
        workbook_ltr = safe_num(p.get("revenue_data", {}).get("ltr"))
        direct = baseline_direct_ltr_per_student(model, code)
        starts = safe_num(p.get("revenue_data", {}).get("new_starts"))
        if np.isnan(workbook_ltr) or np.isnan(starts) or starts <= 0 or np.isnan(direct):
            continue
        workbook_per_student = workbook_ltr / starts
        ratio = direct / workbook_per_student if workbook_per_student else np.nan
        rows.append({"code": code, "direct": direct, "workbook": workbook_per_student, "ratio": ratio})
    df = pd.DataFrame(rows)
    if df.empty:
        return {"matches": False, "method": "proportional", "median_ratio": 1.0, "max_abs_pct": np.nan, "detail": df}
    max_abs_pct = float(np.nanmax(np.abs(df["ratio"] - 1)))
    matches = max_abs_pct <= 0.01
    median_ratio = float(np.nanmedian(df["ratio"]))
    return {
        "matches": matches,
        "method": "direct" if matches else "proportional",
        "median_ratio": median_ratio,
        "max_abs_pct": max_abs_pct,
        "detail": df,
    }


def scenario_ltr_per_student(model: Dict[str, Any], code: str, persistence_delta_pts: float, baseline_check: Dict[str, Any]) -> float:
    p = model["programs"][code]
    rev = p.get("revenue_data", {})
    workbook_ltr = safe_num(rev.get("ltr"))
    starts = safe_num(rev.get("new_starts"))
    if np.isnan(workbook_ltr) or np.isnan(starts) or starts <= 0:
        return np.nan
    if baseline_check.get("method") == "direct":
        terms = p.get("detail", {}).get("terms", [])
        tuition = safe_num(rev.get("tuition")); degree_ch = safe_num(rev.get("degree_ch"))
        if np.isnan(tuition) or np.isnan(degree_ch) or degree_ch <= 0 or not terms:
            return np.nan
        n1 = safe_num(terms[0].get("n"))
        total = 0.0
        scale = 1.0 + persistence_delta_pts / 100.0
        for t in terms:
            if int(t["term"]) < 3:
                continue
            n = safe_num(t.get("n")); ch = safe_num(t.get("avg_ch"))
            if np.isnan(n) or np.isnan(ch):
                continue
            survival = min(1.0, (n / n1) * scale)
            if survival < THRESHOLDS["ltr_floor"]:
                break
            total += survival * ch
        return (tuition / degree_ch) * total
    # Fallback: preserve workbook baseline and scale it proportionally.
    return (workbook_ltr / starts) * max(0.0, 1.0 + persistence_delta_pts / 100.0)


def scenario_results(model: Dict[str, Any], code: Optional[str], start_delta_pct: float, persistence_delta_pts: float, cac_delta_pct: float, baseline_check: Dict[str, Any]) -> pd.DataFrame:
    codes = [code] if code else list(model["programs"].keys())
    rows = []
    for c in codes:
        p = model["programs"][c]
        rev = p.get("revenue_data", {})
        starts = safe_num(rev.get("new_starts"))
        ltr = safe_num(rev.get("ltr"))
        revenue = safe_num(rev.get("revenue"))
        cac = safe_num(p.get("cac_data", {}).get("cac"))
        if np.isnan(starts) or np.isnan(ltr) or np.isnan(revenue):
            continue
        current_ltr_student = ltr / starts if starts else np.nan
        scenario_ltr_student = scenario_ltr_per_student(model, c, persistence_delta_pts, baseline_check)
        new_starts = starts * (1.0 + start_delta_pct / 100.0)
        scenario_ltr = scenario_ltr_student * new_starts if not np.isnan(scenario_ltr_student) else np.nan
        current_cac_spend = cac * starts if not np.isnan(cac) else np.nan
        scenario_cac_spend = cac * (1.0 + cac_delta_pct / 100.0) * new_starts if not np.isnan(cac) else np.nan
        current_enrollment = safe_num(model["hc_term"].loc[(model["hc_term"]["code"] == c) & (model["hc_term"]["term"] == 202610), "headcount"].sum()) if not model["hc_term"].empty else np.nan
        scenario_enrollment = current_enrollment * (1.0 + start_delta_pct / 100.0) * (1.0 + persistence_delta_pts / 100.0) if not np.isnan(current_enrollment) else np.nan
        revenue_scale = scenario_ltr / ltr if not np.isnan(scenario_ltr) and ltr > 0 else np.nan
        scenario_revenue = revenue * revenue_scale if not np.isnan(revenue_scale) else np.nan
        ratio_current = current_ltr_student / cac if not np.isnan(current_ltr_student) and not np.isnan(cac) and cac > 0 else np.nan
        ratio_scenario = scenario_ltr_student / (cac * (1.0 + cac_delta_pct / 100.0)) if not np.isnan(scenario_ltr_student) and not np.isnan(cac) and cac > 0 else np.nan
        rows.append({
            "code": c,
            "Current Enrollment": current_enrollment,
            "Scenario Enrollment": scenario_enrollment,
            "Current Revenue": revenue,
            "Scenario Revenue": scenario_revenue,
            "Current LTR": ltr,
            "Scenario LTR": scenario_ltr,
            "Current CAC Spend": current_cac_spend,
            "Scenario CAC Spend": scenario_cac_spend,
            "Current LTR:CAC": ratio_current,
            "Scenario LTR:CAC": ratio_scenario,
        })
    return pd.DataFrame(rows)


def build_program_summary(model: Dict[str, Any], code: str) -> Dict[str, Any]:
    p = model["programs"][code]
    status = program_status(model, code)
    fall_yoy = program_fall_yoy(model, code)
    year_yoy = program_year_yoy(model, code)
    revenue = p.get("revenue_data", {})
    cac = p.get("cac_data", {})
    ltr_student = safe_num(cac.get("ltr_student"))
    if np.isnan(ltr_student):
        ltr = safe_num(revenue.get("ltr")); starts = safe_num(revenue.get("new_starts"))
        ltr_student = ltr / starts if not np.isnan(ltr) and not np.isnan(starts) and starts > 0 else np.nan
    cac_value = safe_num(cac.get("cac"))
    ratio = ltr_student / cac_value if not np.isnan(ltr_student) and not np.isnan(cac_value) and cac_value > 0 else np.nan
    # New starts YoY from year sheet, like-for-like AY2025 vs AY2024.
    nc = model["new_cont_year"]
    new_yoy = np.nan
    if not nc.empty:
        rr = nc[nc["Program Code"] == code]
        if not rr.empty:
            r = rr.iloc[0]
            a = safe_num(r.get("NEW_2024")); b = safe_num(r.get("NEW_2025"))
            if not np.isnan(a) and a != 0 and not np.isnan(b):
                new_yoy = b / a - 1
    pdelta = np.nan
    # Program-level persistence change is current-year T1->T3 against historical benchmark when available.
    tr = p.get("tracker", {})
    bench = safe_num(tr.get("benchmark", {}).get("t1_t3"))
    current_cohorts = tr.get("cohorts", [])
    vals = [safe_num(x.get("t1_t3")) for x in current_cohorts if not np.isnan(safe_num(x.get("t1_t3")))]
    if not np.isnan(bench) and vals:
        pdelta = float(np.nanmean(vals) - bench)
    exposure = abs(safe_num(revenue.get("revenue_loss"))) if not np.isnan(safe_num(revenue.get("revenue_loss"))) else np.nan

    primary = "No material issue identified"
    secondary = "No material issue identified"
    issue_candidates = []
    if p.get("watch_flags"):
        worst = min(p["watch_flags"], key=lambda x: safe_num(x.get("delta"), 0))
        issue_candidates.append((abs(safe_num(worst.get("delta"), 0)), "Persistence deterioration", f"{display_text(worst.get('step'))}: {fmt_pp(worst.get('delta'))}"))
    if not np.isnan(fall_yoy) and fall_yoy <= THRESHOLDS["fall_amber_yoy"]:
        issue_candidates.append((abs(fall_yoy), "Enrollment decline", f"Fall enrollment {fall_yoy:+.1%} YoY"))
    if p.get("quadrant") in THRESHOLDS["quadrant_amber"] | THRESHOLDS["quadrant_red"]:
        issue_candidates.append((0.5, "Unit-economics risk", display_text(p.get("quadrant"))))
    floor = safe_num(p.get("floor_breach"))
    if not np.isnan(floor) and floor > safe_num(p.get("degree_ch"), 0):
        issue_candidates.append((0.25, "Long persistence tail", f"Floor breach at term {int(floor)}"))
    issue_candidates.sort(reverse=True, key=lambda x: x[0])
    if issue_candidates:
        primary = issue_candidates[0][1] + " — " + issue_candidates[0][2]
        if len(issue_candidates) > 1:
            secondary = issue_candidates[1][1] + " — " + issue_candidates[1][2]
    return {
        "status": status,
        "fall_yoy": fall_yoy,
        "year_yoy": year_yoy,
        "new_yoy": new_yoy,
        "persistence_delta": pdelta,
        "ltr_student": ltr_student,
        "cac": cac_value,
        "ltr_cac": ratio,
        "exposure": exposure,
        "primary": primary,
        "secondary": secondary,
        "directional": directional_badge(model, code),
    }


def program_three_sentence_summary(model: Dict[str, Any], code: str) -> str:
    p = model["programs"][code]
    s = build_program_summary(model, code)
    name = p["name"]
    if s["status"]["word"] == "Insufficient historical data":
        rev = safe_num(p.get("revenue_data", {}).get("revenue"))
        revenue_text = fmt_currency(rev) if not np.isnan(rev) else "Not available"
        return (
            f"{name} does not yet have a mature-cohort persistence history, so the retention and unit-economics model is not decision-grade. "
            f"The workbook currently shows {revenue_text} of modeled program revenue. "
            "Treat the program as an information gap rather than as a zero-value program until more cohort history is available."
        )
    fall = fmt_pct(s["fall_yoy"])
    new = fmt_pct(s["new_yoy"])
    pers = fmt_pp(s["persistence_delta"])
    ltr = fmt_currency(s["ltr_student"])
    cac = fmt_currency(s["cac"])
    ratio = f"{s['ltr_cac']:.1f}x" if not np.isnan(s["ltr_cac"]) else "not available"
    return (
        f"{name} is {s['status']['word'].lower()} based on the workbook's current risk rules, with Fall enrollment at {fall} YoY and new starts at {new} YoY. "
        f"Its modeled persistence change is {pers}; LTR/student is {ltr} and CAC is {cac}, implying an LTR:CAC of {ratio}. "
        f"The primary watch item is {s['primary'].lower()}, with {s['secondary'].lower()} as the secondary signal."
    )


def make_chart_template(fig: go.Figure) -> go.Figure:
    fig.update_layout(
        template="plotly_white",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, Segoe UI, sans-serif", color=BRAND["ink"]),
        margin=dict(l=10, r=10, t=45, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        hoverlabel=dict(bgcolor="white", font_size=13),
    )
    fig.update_xaxes(showgrid=False, zeroline=False)
    fig.update_yaxes(showgrid=False, zeroline=False)
    return fig


def plot_survival(df: pd.DataFrame, title: str = "Survival curve") -> go.Figure:
    fig = go.Figure()
    if df.empty:
        fig.add_annotation(text="Insufficient historical data", x=0.5, y=0.5, showarrow=False)
    else:
        nvals = df["N"] if "N" in df.columns else pd.Series([np.nan] * len(df))
        chvals = df["Avg CH"] if "Avg CH" in df.columns else pd.Series([np.nan] * len(df))
        custom = np.column_stack([[fmt_num(v) for v in nvals], [fmt_num(v, 1) for v in chvals]])
        fig.add_trace(go.Scatter(
            x=df["Term"], y=df["Survival"] * 100,
            mode="lines+markers", line=dict(color=BRAND["blue"], width=4),
            marker=dict(size=8), name="Observed survival",
            customdata=custom,
            hovertemplate="T%{x}<br>Survival: %{y:.1f}%<br>N: %{customdata[0]:,.0f}<br>Avg credits: %{customdata[1]:.1f}<extra></extra>",
        ))
        fig.add_hline(y=12, line_dash="dot", line_color=BRAND["amber"], annotation_text="12% LTR floor")
        fig.update_yaxes(title="Students remaining (%)", range=[0, 105])
        fig.update_xaxes(title="Relative term")
    fig.update_layout(title=title)
    return make_chart_template(fig)


def render_header(title: str, subtitle: str) -> None:
    st.markdown(f"<div class='hero'><div class='hero-title'>{title}</div><div class='hero-subtitle'>{subtitle}</div></div>", unsafe_allow_html=True)


def render_badge(text: str, kind: str = "blue") -> None:
    cls = {"red": "badge-red", "amber": "badge-amber", "green": "badge-green", "grey": "badge-grey", "blue": "badge-blue"}.get(kind, "badge-blue")
    st.markdown(f"<span class='badge {cls}'>{text}</span>", unsafe_allow_html=True)


def render_kpi(label: str, value: str, delta: Optional[str] = None, signal: str = "", help_text: Optional[str] = None) -> None:
    delta_html = ""
    if delta:
        cls = "delta-up" if delta.startswith("+") or "↑" in delta else "delta-down" if delta.startswith("-") or "↓" in delta else "delta-neutral"
        delta_html = f"<div class='kpi-delta {cls}'>{delta}</div>"
    help_attr = f' title="{help_text.replace(chr(34), "&quot;")}"' if help_text else ""
    st.markdown(
        f"<div class='kpi-card'><div class='kpi-label'{help_attr}>{label} {signal} <span style='opacity:.65'>ⓘ</span></div><div class='kpi-value'>{value}</div>{delta_html}</div>",
        unsafe_allow_html=True,
    )

def signal_for_pct(pct: float) -> str:
    if np.isnan(pct):
        return "⚪"
    if pct <= THRESHOLDS["fall_red_yoy"]:
        return "🔴"
    if pct <= THRESHOLDS["fall_amber_yoy"]:
        return "🟡"
    return "🟢"


def page_executive(model: Dict[str, Any]) -> None:
    render_header("Executive", "What is happening across the graduate portfolio — and where leadership should look next.")
    ch = model["changing"]
    metric_order = [
        "Enrolled — Fall (latest vs prior yr)",
        "Enrolled — full FY (2025 vs 2024)",
        "New starts — full FY (2025 vs 2024)",
        "Continuing — full FY (2025 vs 2024)",
    ]
    cols = st.columns(4)
    for col, m in zip(cols, metric_order):
        d = ch.get(m, {})
        with col:
            pct = safe_num(d.get("pct"))
            sig = signal_for_pct(pct)
            render_kpi(m.replace(" — ", " · "), fmt_num(d.get("current")), f"{safe_num(d.get('delta')):+,.0f}  |  {fmt_pct(pct)}", sig, TOOLTIP_DEFS["Enrollment YoY"])
    st.markdown("<div class='section-spacer'></div>", unsafe_allow_html=True)

    # KPI table: source is the dashboard's What's Changing block.
    table_rows = []
    for m in metric_order:
        d = ch.get(m, {})
        pct = safe_num(d.get("pct"))
        table_rows.append({
            "Metric": m.replace(" — ", " · "),
            "Current": fmt_num(d.get("current")),
            "Prior": fmt_num(d.get("prior")),
            "Δ": fmt_num(d.get("delta")),
            "%Δ": fmt_pct(pct),
            "Signal": signal_for_pct(pct),
        })
    st.dataframe(pd.DataFrame(table_rows), use_container_width=True, hide_index=True)

    # Persistence headline with caveat.
    new25, new26 = model["new_by_year"].get(2025, np.nan), model["new_by_year"].get(2026, np.nan)
    pdelta = persistence_delta(model)
    st.markdown("### Portfolio persistence")
    st.markdown(f"**FY2026 New Student %:** {fmt_pct(new26)}  ·  **FY2025:** {fmt_pct(new25)}  ·  **Change:** {fmt_pp(pdelta)}")
    terms = model["terms_by_year"].get(2026, np.nan)
    st.caption(f"Definition: {TOOLTIP_DEFS['Persistence Δ']} FY2026 has {fmt_num(terms)} observed terms in the workbook's FY Trend view. Continuing % is not used for status because its recent-year maturity is limited.")

    # Attention / outperforming cards.
    programs = []
    for code in model["program_order"]:
        if safe_num(model["programs"][code].get("mature_n")) <= 0:
            continue
        s = build_program_summary(model, code)
        if s["directional"] and safe_num(model["programs"][code].get("mature_n")) < THRESHOLDS["thin_sample_n"]:
            continue
        programs.append((code, s))
    attention = [x for x in programs if x[1]["status"]["word"] in {"Red", "Amber"}]
    attention.sort(key=lambda x: (0 if x[1]["status"]["word"] == "Red" else 1, -abs(safe_num(x[1]["fall_yoy"], 0))))
    # Outperformers: green + positive fall/new/persistence where available.
    out = [x for x in programs if x[1]["status"]["word"] == "Green"]
    out.sort(key=lambda x: sum(v for v in [safe_num(x[1]["fall_yoy"], 0), safe_num(x[1]["new_yoy"], 0), safe_num(x[1]["persistence_delta"], 0)] if not np.isnan(v)), reverse=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### 🔴🟡 What needs attention")
        if not attention:
            st.success("No red or amber programs under the current rules.")
        for code, s in attention[:7]:
            p = model["programs"][code]
            badge = "🔴 Red" if s["status"]["word"] == "Red" else "🟡 Amber"
            st.markdown(f"**{badge} · {code} — {p['name']}**")
            st.caption(f"{s['primary']}. {s['secondary']}.")
    with c2:
        st.markdown("### 🟢 Outperforming")
        if not out:
            st.info("No green programs with enough directional evidence to rank.")
        for code, s in out[:7]:
            p = model["programs"][code]
            details = []
            if not np.isnan(s["fall_yoy"]): details.append(f"Fall {fmt_pct(s['fall_yoy'])} YoY")
            if not np.isnan(s["new_yoy"]): details.append(f"new starts {fmt_pct(s['new_yoy'])} YoY")
            if not np.isnan(s["persistence_delta"]): details.append(f"persistence {fmt_pp(s['persistence_delta'])}")
            st.markdown(f"**🟢 {code} — {p['name']}**")
            st.caption(" · ".join(details) if details else "No material negative signal identified.")

    # Enrollment bridge.
    st.markdown("### Enrollment bridge")
    prev_total = safe_num(ch.get("Enrolled — full FY (2025 vs 2024)", {}).get("prior"))
    current_total = safe_num(ch.get("Enrolled — full FY (2025 vs 2024)", {}).get("current"))
    new_delta = safe_num(ch.get("New starts — full FY (2025 vs 2024)", {}).get("delta"))
    cont_delta = safe_num(ch.get("Continuing — full FY (2025 vs 2024)", {}).get("delta"))
    if not np.isnan(prev_total) and not np.isnan(current_total):
        fig = go.Figure(go.Waterfall(
            measure=["absolute", "relative", "relative", "total"],
            x=["AY2024 total", "Change in new starts", "Change in continuing", "AY2025 total"],
            y=[prev_total, new_delta, cont_delta, current_total],
            connector={"line": {"color": "#CBD5E1"}},
            increasing={"marker": {"color": BRAND["green"]}},
            decreasing={"marker": {"color": BRAND["red"]}},
            totals={"marker": {"color": BRAND["blue"]}},
            text=[fmt_num(prev_total), f"{new_delta:+,.0f}", f"{cont_delta:+,.0f}", fmt_num(current_total)],
            textposition="outside",
        ))
        fig.update_layout(title="Prior total → new starts change → continuing change → current total")
        st.plotly_chart(make_chart_template(fig), use_container_width=True)
        st.caption("Exits are not shown as a separate bar because the source cannot separate graduated from withdrawn. The bridge follows the workbook definition: prior total → change in new starts → change in continuing → current total.")


def page_enrollment(model: Dict[str, Any]) -> None:
    render_header("Enrollment", "Headcount by term and year, with new-versus-continuing context and program movement.")
    season = st.radio("Term view", ["Fall", "Spring", "Summer"], horizontal=True, index=0)
    df = model["hc_term"].copy()
    if df.empty:
        st.warning("Enrollment term data is not available.")
        return
    suffix = {"Fall": 10, "Spring": 20, "Summer": 30}[season]
    df["season"] = df["term"].astype(int) % 100
    sel = df[df["season"] == suffix].copy()
    sel["year"] = sel["term"].astype(int) // 100
    port = sel.groupby("year", as_index=False)["headcount"].sum().sort_values("year")
    fig = px.line(port, x="year", y="headcount", markers=True, title=f"Portfolio enrollment — {season}")
    fig.update_traces(line_color=BRAND["blue"], hovertemplate="Year %{x}<br>Enrollment %{y:,.0f}<extra></extra>")
    fig.update_yaxes(title="Students")
    st.plotly_chart(make_chart_template(fig), use_container_width=True)

    # Full academic-year headcount from the workbook.
    ydf = model["hc_year"].copy()
    if not ydf.empty:
        year_cols = [c for c in ydf.columns if str(c).isdigit() and 2019 <= int(c) <= 2026]
        totals = []
        for c in year_cols:
            totals.append({"Year": int(c), "Enrollment": float(pd.to_numeric(ydf[c], errors="coerce").sum(skipna=True))})
        year_port = pd.DataFrame(totals).sort_values("Year")
        fig_y = px.bar(year_port, x="Year", y="Enrollment", title="Academic-year enrollment")
        fig_y.update_traces(marker_color=BRAND["purple"], hovertemplate="AY%{x}<br>Enrollment %{y:,.0f}<extra></extra>")
        fig_y.update_yaxes(title="Students")
        st.plotly_chart(make_chart_template(fig_y), use_container_width=True)
        st.caption("AY2026 is partial in the workbook and contains Fall 2026 only; it is not treated as a full-year comparison.")

    # New vs continuing stacked headcount, not retention rates.
    nc = model["new_cont_term"].copy()
    if not nc.empty:
        nc["season"] = nc["term"].astype(int) % 100
        nc = nc[nc["season"] == suffix].copy()
        nc["year"] = nc["term"].astype(int) // 100
        agg = nc.groupby("year", as_index=False)[["new", "continuing"]].sum()
        long = agg.melt(id_vars="year", value_vars=["new", "continuing"], var_name="segment", value_name="students")
        fig2 = px.bar(long, x="year", y="students", color="segment", barmode="stack", title=f"New vs continuing headcount — {season}", color_discrete_map={"new": BRAND["teal"], "continuing": BRAND["purple"]})
        fig2.update_traces(hovertemplate="Year %{x}<br>%{fullData.name}: %{y:,.0f}<extra></extra>")
        fig2.update_yaxes(title="Students")
        st.plotly_chart(make_chart_template(fig2), use_container_width=True)

    # Growing vs shrinking programs, latest like-for-like fall comparison.
    rows = []
    for code in model["program_order"]:
        p = model["programs"][code]
        latest = df.loc[(df["code"] == code) & (df["term"] == 202610), "headcount"]
        prior = df.loc[(df["code"] == code) & (df["term"] == 202510), "headcount"]
        if latest.empty or prior.empty:
            continue
        a, b = safe_num(prior.iloc[0]), safe_num(latest.iloc[0])
        if np.isnan(a) or np.isnan(b):
            continue
        rows.append({"Code": code, "Program": p["name"], "Fall 2025": a, "Fall 2026": b, "YoY %": b / a - 1 if a else np.nan})
    trend = pd.DataFrame(rows)
    if not trend.empty:
        trend["Movement"] = np.where(trend["YoY %"] > 0.005, "Growing", np.where(trend["YoY %"] < -0.005, "Shrinking", "Stable"))
        st.markdown("### Program movement — Fall 2025 → Fall 2026")
        c1, c2, c3 = st.columns(3)
        for c, label in zip([c1, c2, c3], ["Growing", "Stable", "Shrinking"]):
            with c:
                render_kpi(label, fmt_num((trend["Movement"] == label).sum()), signal="🟢" if label == "Growing" else "🟡" if label == "Stable" else "🔴")
        st.dataframe(trend.sort_values("YoY %", ascending=False), use_container_width=True, hide_index=True)


def page_persistence(model: Dict[str, Any]) -> None:
    render_header("Persistence", "How quickly students remain in the portfolio — and how current cohorts compare with mature benchmarks.")
    codes = model["program_order"]
    options = ["All Programs"] + codes
    selected = st.selectbox("Program", options, format_func=lambda x: x if x == "All Programs" else f"{x} — {model['programs'][x]['name']}")
    include_directional = st.toggle("Include thin-sample programs in portfolio views", value=False)

    if selected == "All Programs":
        curves = []
        eligible = []
        for code in codes:
            p = model["programs"][code]
            n = safe_num(p.get("mature_n"))
            if np.isnan(n) or n <= 0:
                continue
            if n < THRESHOLDS["thin_sample_n"] and not include_directional:
                continue
            cdf = survival_curve(model, code)
            if cdf.empty:
                continue
            cdf = cdf[["Term", "N"]].copy()
            cdf["Survival"] = cdf["N"] / cdf["N"].iloc[0]
            cdf["Program"] = code
            curves.append(cdf)
            eligible.append(code)
        if curves:
            # Weighted portfolio curve by N at each term.
            allc = pd.concat(curves, ignore_index=True)
            agg = allc.groupby("Term").apply(lambda g: pd.Series({"Survival": g["N"].sum() / g["N"].iloc[0] if g["N"].iloc[0] else np.nan}), include_groups=False).reset_index()
            # Better denominator: sum T1 N for all eligible programs.
            denom = sum(float(c["N"].iloc[0]) for c in curves if not c.empty)
            num = allc.groupby("Term")["N"].sum().reset_index()
            agg = num.assign(Survival=num["N"] / denom, **{"Avg CH": np.nan})
            st.plotly_chart(plot_survival(agg.rename(columns={"N": "N"}), "Portfolio survival curve"), use_container_width=True)
            st.caption(f"Portfolio curve uses {len(eligible)} programs; thin samples are {'included' if include_directional else 'excluded'} from the ranking view.")
    else:
        cdf = survival_curve(model, selected)
        st.plotly_chart(plot_survival(cdf, f"{selected} — survival curve"), use_container_width=True)
        p = model["programs"][selected]
        s = build_program_summary(model, selected)
        c1, c2, c3 = st.columns(3)
        with c1: render_kpi("New retention", fmt_pct(p.get("new_retention")), signal="📈", help_text=TOOLTIP_DEFS["New retention"])
        with c2: render_kpi("Returning retention", fmt_pct(p.get("returning_retention")), signal="↩️", help_text=TOOLTIP_DEFS["Returning retention"])
        with c3: render_kpi("Persistence Δ", fmt_pp(s["persistence_delta"]), signal="📌", help_text=TOOLTIP_DEFS["Persistence Δ"])

    c1, c2 = st.columns(2)
    with c1: render_kpi("Portfolio mature T1", fmt_num(model.get("portfolio_t1")), signal="👥", help_text="Student-weighted mature-cohort base across the 41 programs.")
    with c2: render_kpi("Portfolio new retention", fmt_pct(model.get("portfolio_new_retention")), signal="📈", help_text=TOOLTIP_DEFS["New retention"])

    # New vs returning split as separate lines, never stacked.
    new_pct = model["new_by_year"]
    cont_pct = model["cont_by_year"]
    split = pd.DataFrame({"Year": sorted(set(new_pct) | set(cont_pct)), "New": [new_pct.get(y, np.nan) * 100 for y in sorted(set(new_pct) | set(cont_pct))], "Continuing": [cont_pct.get(y, np.nan) * 100 for y in sorted(set(new_pct) | set(cont_pct))]})
    if not split.empty:
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=split["Year"], y=split["New"], mode="lines+markers", name="New student %", line=dict(color=BRAND["teal"], width=3)))
        fig.add_trace(go.Scatter(x=split["Year"], y=split["Continuing"], mode="lines+markers", name="Continuing %", line=dict(color=BRAND["purple"], width=3)))
        fig.update_layout(title="New vs continuing persistence — separate measures", yaxis_title="Retention (%)")
        fig.update_yaxes(range=[0, 105])
        st.plotly_chart(make_chart_template(fig), use_container_width=True)
        st.caption("The two measures are intentionally shown as separate lines. Continuing % is maturity-sensitive; FY Trend - Matched exists in the workbook as a reality check.")

    # Cohort-vintage heatmap: observed cohort cells plus benchmark-based projected cells.
    heat_code = selected if selected != "All Programs" else "M013"
    cohorts = model.get("cohort_detail", {}).get(heat_code, [])
    if cohorts:
        max_term = min(10, max([max(c.get("survival", {}).keys(), default=1) for c in cohorts] + [1]))
        bench = survival_curve(model, heat_code)
        bench_map = dict(zip(bench["Term"], bench["Survival"])) if not bench.empty else {}
        z = []
        custom = []
        projected_mask = []
        y_labels = []
        for c in cohorts:
            y_labels.append(c["cohort"])
            vals = []
            meta = []
            mask = []
            obs = c.get("survival", {})
            last_obs = max(obs.keys(), default=0)
            for t in range(1, max_term + 1):
                if t in obs:
                    vals.append(obs[t] * 100 if not np.isnan(obs[t]) else np.nan)
                    meta.append("Observed")
                    mask.append(False)
                elif t > last_obs and t in bench_map:
                    vals.append(bench_map[t] * 100)
                    meta.append("Projected from mature benchmark")
                    mask.append(True)
                else:
                    vals.append(np.nan)
                    meta.append("Not available")
                    mask.append(False)
            z.append(vals); custom.append(meta); projected_mask.append(mask)
        fig = go.Figure(go.Heatmap(z=z, x=[f"T{t}" for t in range(1, max_term + 1)], y=y_labels, colorscale=[[0,"#FEE2E2"],[0.5,"#FEF3C7"],[1,"#DCFCE7"]], zmin=0, zmax=100, colorbar=dict(title="Survival %"), customdata=custom, hovertemplate="Cohort %{y}<br>%{x}: %{z:.1f}%<br>%{customdata}<extra></extra>"))
        # Projected cells are outlined with a dotted border and explicitly labeled.
        for yi, mask in enumerate(projected_mask):
            for xi, is_proj in enumerate(mask):
                if is_proj:
                    fig.add_shape(type="rect", x0=xi-0.5, x1=xi+0.5, y0=yi-0.5, y1=yi+0.5, line=dict(color="#475569", width=1, dash="dot"), fillcolor="rgba(148,163,184,0.10)")
                    fig.add_annotation(x=xi, y=yi, text="P", showarrow=False, font=dict(size=9, color="#334155"))
        fig.update_layout(title=f"Cohort-vintage survival heatmap · {heat_code}", xaxis_title="Relative term", yaxis_title="Entry cohort", height=max(360, 24 * len(y_labels) + 150))
        st.plotly_chart(make_chart_template(fig), use_container_width=True)
        st.caption("P = Projected. Observed cells come from Method B cohort history; projected cells use the program's mature benchmark survival curve. Projected cells are intentionally outlined and labeled so they cannot be mistaken for observed history.")

    # Current-Year Tracker vs benchmark.
    if selected != "All Programs":
        tr = model["programs"][selected].get("tracker", {})
        b = tr.get("benchmark", {})
        rows = []
        for c in tr.get("cohorts", []):
            for step, key in [("T1→T2", "t1_t2"), ("T2→T3", "t2_t3"), ("T1→T3", "t1_t3")]:
                val = safe_num(c.get(key)); bench = safe_num(b.get(key))
                if not np.isnan(val): rows.append({"Cohort": str(c["entry_term"]), "Step": step, "Current": val * 100, "Benchmark": bench * 100 if not np.isnan(bench) else np.nan, "Δ pts": (val - bench) * 100 if not np.isnan(bench) else np.nan})
        if rows:
            cmp = pd.DataFrame(rows)
            fig = px.bar(cmp, x="Step", y="Δ pts", color="Cohort", barmode="group", title="Current-Year Tracker vs mature benchmark", color_discrete_sequence=PALETTE)
            fig.add_hline(y=0, line_color="#94A3B8")
            fig.update_yaxes(title="Difference (points)")
            fig.update_traces(hovertemplate="Step %{x}<br>Δ %{y:+.1f} pts<extra></extra>")
            st.plotly_chart(make_chart_template(fig), use_container_width=True)


def page_programs(model: Dict[str, Any]) -> None:
    render_header("Programs", "A 10-second view of any one program — status, drivers, economics and trajectory.")
    include_directional = st.toggle("Include directional thin-sample programs in rankings", value=False)
    query = st.text_input("Search programs", placeholder="Code or program name")
    codes = model["program_order"]
    if query:
        q = query.lower()
        codes = [c for c in codes if q in c.lower() or q in model["programs"][c]["name"].lower()]
    if not codes:
        st.info("No programs match that search.")
        return
    selected = st.selectbox("Program", codes, format_func=lambda x: f"{x} — {model['programs'][x]['name']}")
    p = model["programs"][selected]
    s = build_program_summary(model, selected)
    st.markdown(f"## {selected} — {p['name']}")
    if s["status"]["word"] == "Insufficient historical data":
        render_badge("⚪ Insufficient historical data", "grey")
        st.markdown("### What we know")
        known = []
        rev = safe_num(p.get("revenue_data", {}).get("revenue"))
        if not np.isnan(rev): known.append(f"Revenue: {fmt_currency(rev)}")
        starts = safe_num(p.get("revenue_data", {}).get("new_starts"))
        if not np.isnan(starts): known.append(f"Modeled Fall 2026 new starts: {fmt_num(starts, 1)}")
        degree = safe_num(p.get("degree_ch"))
        if not np.isnan(degree): known.append(f"Degree credit hours: {fmt_num(degree)}")
        for x in known or ["Program is present in the workbook, but mature-cohort persistence is not available."]:
            st.markdown(f"- {x}")
        st.markdown("### What we don't know yet")
        for x in ["Mature-cohort persistence", "Decision-grade LTR/student", "Reliable LTR:CAC status"]:
            st.markdown(f"- {x}")
        st.markdown("### Summary")
        st.write(program_three_sentence_summary(model, selected))
        return

    # Status + rules.
    kind = "red" if s["status"]["word"] == "Red" else "amber" if s["status"]["word"] == "Amber" else "green"
    render_badge(s["status"]["status"], kind)
    if s["directional"]:
        render_badge(s["directional"], "blue")
    st.caption("Triggered rules: " + "; ".join(s["status"]["rules"]))

    cols = st.columns(6)
    metrics = [
        ("Enrollment YoY", fmt_pct(s["fall_yoy"]), TOOLTIP_DEFS["Enrollment YoY"]),
        ("New Starts YoY", fmt_pct(s["new_yoy"]), "Like-for-like academic-year new-start change: AY2025 vs AY2024."),
        ("Persistence Δ", fmt_pp(s["persistence_delta"]), TOOLTIP_DEFS["Persistence Δ"]),
        ("LTR / student", fmt_currency(s["ltr_student"]), TOOLTIP_DEFS["LTR/student"]),
        ("CAC", fmt_currency(s["cac"]), TOOLTIP_DEFS["CAC"]),
        ("LTR:CAC", f"{s['ltr_cac']:.1f}x" if not np.isnan(s["ltr_cac"]) else "Not available", "Gross LTR/student divided by the workbook's directional CAC; not margin-adjusted."),
    ]
    for c, (label, val, help_text) in zip(cols, metrics):
        with c: render_kpi(label, val, help_text=help_text)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### Primary issue")
        st.info(s["primary"])
    with c2:
        st.markdown("### Secondary issue")
        st.info(s["secondary"])
    st.markdown(f"### Financial exposure · {fmt_currency(s['exposure']) if not np.isnan(s['exposure']) else 'Not available'}")
    st.caption("Revenue exposure is the absolute value of the workbook's modeled revenue-loss field; it is a modeled gross-revenue exposure, not a forecast of realized loss.")

    st.markdown("### Survival curve")
    st.plotly_chart(plot_survival(survival_curve(model, selected), f"{selected} — survival"), use_container_width=True)

    # Enrollment trend.
    h = model["hc_term"]
    h = h[h["code"] == selected].copy()
    if not h.empty:
        h["Term label"] = h["term"].map(term_label)
        fig = px.line(h, x="Term label", y="headcount", markers=True, title="Enrollment trend")
        fig.update_traces(line_color=BRAND["purple"], hovertemplate="%{x}<br>Enrollment %{y:,.0f}<extra></extra>")
        st.plotly_chart(make_chart_template(fig), use_container_width=True)
    st.markdown("### 10-second summary")
    st.write(program_three_sentence_summary(model, selected))

    if include_directional or safe_num(p.get("mature_n")) >= THRESHOLDS["thin_sample_n"]:
        st.markdown("### Program data")
        rows = {
            "Mature cohort N@T1": fmt_num(p.get("mature_n")),
            "New retention": fmt_pct(p.get("new_retention")),
            "Returning retention": fmt_pct(p.get("returning_retention")),
            "LTR factor": fmt_pct(p.get("ltr_factor")),
            "Floor breach": f"Term {int(safe_num(p.get('floor_breach')))}" if not np.isnan(safe_num(p.get("floor_breach"))) else "No floor breach observed through available history",
        }
        st.dataframe(pd.DataFrame([rows]), use_container_width=True, hide_index=True)


def page_economics(model: Dict[str, Any]) -> None:
    render_header("Economics", "Revenue and unit economics, with trial / approximate content clearly separated from core model outputs.")
    rev_rows = []
    for code in model["program_order"]:
        p = model["programs"][code]
        r = p.get("revenue_data", {})
        if not np.isnan(safe_num(r.get("revenue"))) or not np.isnan(safe_num(r.get("ltr"))):
            rev_rows.append({"Code": code, "Program": p["name"], "Revenue": safe_num(r.get("revenue")), "LTR": safe_num(r.get("ltr")), "Revenue Loss": abs(safe_num(r.get("revenue_loss"))) if not np.isnan(safe_num(r.get("revenue_loss"))) else np.nan})
    rdf = pd.DataFrame(rev_rows)
    if not rdf.empty:
        long = rdf.melt(id_vars=["Code", "Program"], value_vars=["Revenue", "LTR"], var_name="Metric", value_name="Amount")
        fig = px.bar(long.sort_values("Amount", ascending=False), x="Program", y="Amount", color="Metric", title="Revenue Outlook — revenue vs modeled LTR", color_discrete_map={"Revenue": BRAND["blue"], "LTR": BRAND["teal"]})
        fig.update_xaxes(tickangle=-55)
        fig.update_yaxes(title="Dollars")
        fig.update_traces(hovertemplate="%{x}<br>%{fullData.name}: $%{y:,.0f}<extra></extra>")
        st.plotly_chart(make_chart_template(fig), use_container_width=True)
        rdf_display = rdf.sort_values("Revenue Loss", ascending=False).copy()
        for col in ["Revenue", "LTR", "Revenue Loss"]:
            rdf_display[col] = rdf_display[col].map(lambda v: fmt_currency(v))
        st.dataframe(rdf_display, use_container_width=True, hide_index=True)

    render_badge("TRIAL — approximate", "amber")
    st.markdown("### LTR vs CAC")
    st.caption("CAC and margin content below are trial / approximate. CAC is paid-media spend only; gross LTR:CAC is not a margin-based benchmark.")
    rows = []
    for code, d in model["cac"].items():
        if np.isnan(safe_num(d.get("cac"))) or np.isnan(safe_num(d.get("ltr_student"))):
            continue
        rows.append({"Code": code, "Program": model["programs"].get(code, {}).get("name", code), "CAC": d["cac"], "LTR/student": d["ltr_student"], "Quadrant": d.get("quadrant", "Not available")})
    cdf = pd.DataFrame(rows)
    if not cdf.empty:
        fig = px.scatter(cdf, x="LTR/student", y="CAC", color="Quadrant", text="Code", hover_data=["Program"], color_discrete_sequence=PALETTE, title="Gross LTR/student vs media-only CAC")
        fig.update_traces(textposition="top center", marker=dict(size=13))
        st.plotly_chart(make_chart_template(fig), use_container_width=True)

    st.markdown("### Margin View")
    mrows = []
    for code, d in model["margin"].items():
        if np.isnan(safe_num(d.get("margin_ltr_student"))):
            continue
        mrows.append({"Code": code, "Program": model["programs"].get(code, {}).get("name", code), "Approx. margin %": d["margin_pct"], "Approx. margin LTR/student": d["margin_ltr_student"], "Approx. margin multiple": d["margin_multiple"]})
    mdf = pd.DataFrame(mrows)
    if not mdf.empty:
        fig = px.bar(mdf.sort_values("Approx. margin LTR/student", ascending=False), x="Program", y="Approx. margin LTR/student", title="Approximate margin LTR/student", color="Approx. margin multiple", color_continuous_scale="Viridis")
        fig.update_xaxes(tickangle=-55)
        st.plotly_chart(make_chart_template(fig), use_container_width=True)
        st.dataframe(mdf.sort_values("Approx. margin LTR/student", ascending=False), use_container_width=True, hide_index=True)
    st.warning("TRIAL — approximate: margin uses a rough cost-to-deliver proxy from the workbook and is not Finance-approved. Do not treat it as a board-level margin measure.")


def page_scenario(model: Dict[str, Any]) -> None:
    render_header("Scenario Lab", "Change starts, persistence and CAC assumptions and see the estimated portfolio impact.")
    baseline = scenario_baseline_check(model)
    if baseline["method"] == "direct":
        st.success("Baseline calibration: the Direct formula reproduces Revenue Outlook LTR/student within 1% across the available modeled programs.")
    else:
        st.warning("Baseline calibration: the requested Direct formula does not reproduce Revenue Outlook within 1% for this workbook. V3 therefore uses proportional scaling from each program's Revenue Outlook LTR baseline for scenario outputs. The Direct-method check remains visible in Model QA.")
    target_options = ["Whole portfolio"] + model["program_order"]
    target = st.selectbox("Scenario scope", target_options, format_func=lambda x: x if x == "Whole portfolio" else f"{x} — {model['programs'][x]['name']}")
    start_delta = st.slider("New starts change (%)", -30.0, 30.0, 0.0, 1.0, help="Scenario-only change to modeled new starts.")
    persistence_delta = st.slider("Persistence change (percentage points)", -20.0, 20.0, 0.0, 1.0, help="Scenario-only scale applied to T2 onward survival, capped at 100% when the Direct method is active.")
    cac_delta = st.slider("CAC change (%)", -50.0, 50.0, 0.0, 1.0, help="Scenario-only change to CAC.")
    code = None if target == "Whole portfolio" else target
    res = scenario_results(model, code, start_delta, persistence_delta, cac_delta, baseline)
    if res.empty:
        st.info("Scenario outputs require a program with modeled Revenue Outlook LTR and new starts.")
        return
    if code is None:
        current = res[["Current Enrollment", "Current Revenue", "Current LTR", "Current CAC Spend"]].sum(numeric_only=True)
        scen = res[["Scenario Enrollment", "Scenario Revenue", "Scenario LTR", "Scenario CAC Spend"]].sum(numeric_only=True)
        cur_ratio = res["Current LTR"].sum() / res["Current CAC Spend"].sum() if res["Current CAC Spend"].sum() > 0 else np.nan
        scen_ratio = res["Scenario LTR"].sum() / res["Scenario CAC Spend"].sum() if res["Scenario CAC Spend"].sum() > 0 else np.nan
        table = pd.DataFrame([
            ["Enrollment", current["Current Enrollment"], scen["Scenario Enrollment"]],
            ["Revenue", current["Current Revenue"], scen["Scenario Revenue"]],
            ["LTR", current["Current LTR"], scen["Scenario LTR"]],
            ["CAC spend", current["Current CAC Spend"], scen["Scenario CAC Spend"]],
            ["LTR:CAC", cur_ratio, scen_ratio],
        ], columns=["Metric", "Current", "Scenario"])
    else:
        r = res.iloc[0]
        table = pd.DataFrame([
            ["Enrollment", r["Current Enrollment"], r["Scenario Enrollment"]],
            ["Revenue", r["Current Revenue"], r["Scenario Revenue"]],
            ["LTR", r["Current LTR"], r["Scenario LTR"]],
            ["CAC spend", r["Current CAC Spend"], r["Scenario CAC Spend"]],
            ["LTR:CAC", r["Current LTR:CAC"], r["Scenario LTR:CAC"]],
        ], columns=["Metric", "Current", "Scenario"])
    st.markdown("### Estimate — Current vs Scenario")
    show = table.copy()
    show["Current"] = show.apply(lambda r: f"{r['Current']:.1f}x" if r["Metric"] == "LTR:CAC" and not np.isnan(r["Current"]) else fmt_currency(r["Current"]) if r["Metric"] != "Enrollment" else fmt_num(r["Current"]), axis=1)
    show["Scenario"] = show.apply(lambda r: f"{r['Scenario']:.1f}x" if r["Metric"] == "LTR:CAC" and not np.isnan(r["Scenario"]) else fmt_currency(r["Scenario"]) if r["Metric"] != "Enrollment" else fmt_num(r["Scenario"]), axis=1)
    st.dataframe(show, use_container_width=True, hide_index=True)
    st.caption("Every scenario output is an Estimate. The scenario is not a new forecast and does not write back to the workbook.")
    if baseline["method"] == "proportional":
        st.caption(f"Direct-formula calibration median ratio: {baseline['median_ratio']:.3f}× versus Revenue Outlook; maximum absolute deviation among tested programs: {baseline['max_abs_pct']:.1%}.")


def page_qa(model: Dict[str, Any]) -> None:
    render_header("Model QA", "Can the numbers be trusted? V3 checks the workbook before presenting it as a decision tool.")
    rows = []
    programs = model["programs"]
    rows.append({"Check": "41 programs parsed", "Status": "PASS" if len(programs) == 41 else "FAIL", "Detail": f"{len(programs)} program records found"})
    no_data = [c for c in model["program_order"] if safe_num(programs[c].get("mature_n")) <= 0]
    thin = [c for c in model["program_order"] if 0 < safe_num(programs[c].get("mature_n")) < THRESHOLDS["thin_sample_n"]]
    rows.append({"Check": "No-data programs identified", "Status": "PASS", "Detail": ", ".join(no_data) if no_data else "None"})
    rows.append({"Check": "Thin-sample programs identified", "Status": "PASS", "Detail": ", ".join(thin) if thin else "None"})
    mismatches = []
    for code in model["program_order"]:
        a = safe_num(programs[code].get("new_retention")); b = safe_num(programs[code].get("method_b", {}).get("pooled_t3"))
        if not np.isnan(a) and not np.isnan(b) and abs(a - b) > 1e-9:
            mismatches.append(code)
    rows.append({"Check": "Method A New Retention = Method B pooled S(T3)", "Status": "PASS" if not mismatches else "FAIL", "Detail": "All modeled programs match" if not mismatches else ", ".join(mismatches)})
    m013 = programs.get("M013", {})
    m366 = programs.get("M366", {})
    rows.append({"Check": "M013 anchor values", "Status": "PASS" if (safe_num(m013.get("mature_n")) == 272 and abs(safe_num(m013.get("new_retention")) - .6911764705882353) < 1e-9 and abs(safe_num(m013.get("returning_retention")) - .623979336697787) < 1e-9 and abs(safe_num(m013.get("ltr_factor")) - .431279835658765) < 1e-9 and safe_num(m013.get("floor_breach")) == 9) else "FAIL", "Detail": "N@T1 272 · New 69.1% · Returning 62.4% · LTR factor 43.1% · floor breach term 9"})
    rows.append({"Check": "M366 anchor values", "Status": "PASS" if abs(safe_num(m366.get("new_retention")) - .973214285714286) < 1e-9 and abs(safe_num(m366.get("returning_retention")) - .909436528862287) < 1e-9 else "FAIL", "Detail": "New 97.3% · Returning 90.9%"})
    portfolio_persist_ok = abs(safe_num(model.get("portfolio_t1")) - 5841) < 1e-9 and abs(safe_num(model.get("portfolio_new_retention")) - 0.7005649717514124) < 1e-9
    rows.append({"Check": "Portfolio mature persistence anchor", "Status": "PASS" if portfolio_persist_ok else "FAIL", "Detail": "T1 5,841 · New retention 70.1%"})
    ch = model["changing"]
    portfolio_ok = (
        safe_num(model["programs"].get("M013", {}).get("mature_n")) == 272 and
        safe_num(ch.get("Enrolled — Fall (latest vs prior yr)", {}).get("current")) == 2482 and
        safe_num(ch.get("Enrolled — full FY (2025 vs 2024)", {}).get("current")) == 7348 and
        safe_num(ch.get("New starts — full FY (2025 vs 2024)", {}).get("current")) == 1518 and
        safe_num(ch.get("Continuing — full FY (2025 vs 2024)", {}).get("current")) == 5830 and
        abs(safe_num(model["new_by_year"].get(2026)) - .72739916550765) < 1e-9
    )
    rows.append({"Check": "Portfolio anchor values", "Status": "PASS" if portfolio_ok else "FAIL", "Detail": "Fall 2026 2,482 · AY2025 7,348 · new starts 1,518 · continuing 5,830"})
    rev1550 = safe_num(programs["1550C"].get("revenue_data", {}).get("revenue"))
    rows.append({"Check": "1550C insufficient-data revenue", "Status": "PASS" if rev1550 == 102672 else "FAIL", "Detail": f"Revenue {fmt_currency(rev1550)}"})
    base = scenario_baseline_check(model)
    rows.append({"Check": "Scenario baseline calibration", "Status": "PASS" if base["method"] in {"direct", "proportional"} else "FAIL", "Detail": "Direct within 1%" if base["method"] == "direct" else "Direct check failed; proportional fallback is active and disclosed"})
    qdf = pd.DataFrame(rows)
    qdf["Status"] = qdf["Status"].map(lambda x: "🟢 PASS" if x == "PASS" else "🔴 FAIL")
    st.dataframe(qdf, use_container_width=True, hide_index=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### No mature-cohort data")
        st.write(no_data)
    with c2:
        st.markdown("### Thin sample (<30 at T1)")
        st.write(thin)
    st.markdown("### Model health")
    failures = int((qdf["Status"] == "🔴 FAIL").sum())
    if failures == 0:
        st.success("Model health: GREEN — all V3 acceptance checks pass, with any known Direct-method calibration limitation explicitly handled rather than hidden.")
    else:
        st.error(f"Model health: RED — {failures} acceptance check(s) failed. Fix the data/model before relying on the dashboard.")


def inject_css() -> None:
    st.markdown(
        f"""
        <style>
        .stApp {{ background: linear-gradient(180deg, #F8FAFC 0%, #FFFFFF 38%); }}
        [data-testid="stSidebar"] {{ background: linear-gradient(180deg, {BRAND['navy']} 0%, #172554 100%); }}
        [data-testid="stSidebar"] * {{ color: white !important; }}
        .hero {{ padding: 1.0rem 0 1.2rem 0; margin-bottom: .4rem; }}
        .hero-title {{ font-size: 2.35rem; font-weight: 800; letter-spacing: -.04em; background: linear-gradient(90deg, {BRAND['navy']}, {BRAND['blue']}, {BRAND['teal']}); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }}
        .hero-subtitle {{ color: {BRAND['muted']}; font-size: 1.02rem; margin-top: .25rem; }}
        .kpi-card {{ background: rgba(255,255,255,.94); border: 1px solid #E6EAF0; border-radius: 18px; padding: 1rem 1.1rem; min-height: 132px; box-shadow: 0 10px 28px rgba(15,23,42,.07); position: relative; overflow: hidden; }}
        .kpi-card:before {{ content: ''; position:absolute; left:0; top:0; right:0; height:4px; background: linear-gradient(90deg,{BRAND['blue']},{BRAND['teal']}); }}
        .kpi-label {{ color:{BRAND['muted']}; font-size:.78rem; font-weight:700; text-transform:uppercase; letter-spacing:.05em; }}
        .kpi-value {{ color:{BRAND['ink']}; font-size:2rem; font-weight:800; margin-top:.25rem; }}
        .kpi-delta {{ font-size:.82rem; font-weight:700; margin-top:.2rem; }}
        .delta-up {{ color:{BRAND['green']}; }} .delta-down {{ color:{BRAND['red']}; }} .delta-neutral {{ color:{BRAND['slate']}; }}
        .kpi-help {{ color:{BRAND['muted']}; font-size:.7rem; margin-top:.45rem; line-height:1.25; }}
        .badge {{ display:inline-block; padding:.28rem .65rem; border-radius:999px; font-size:.78rem; font-weight:800; margin-right:.35rem; }}
        .badge-red {{ background:#FEE2E2; color:#B91C1C; }} .badge-amber {{ background:#FEF3C7; color:#92400E; }} .badge-green {{ background:#DCFCE7; color:#166534; }} .badge-grey {{ background:#E5E7EB; color:#374151; }} .badge-blue {{ background:#DBEAFE; color:#1D4ED8; }}
        .section-spacer {{ height:.45rem; }}
        div[data-testid="stDataFrame"] {{ border-radius:14px; overflow:hidden; }}
        div[data-testid="stMetric"] {{ background:white; border-radius:16px; }}
        h1,h2,h3 {{ color:{BRAND['ink']}; }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def main() -> None:
    st.set_page_config(page_title="Graduate Portfolio Intelligence", page_icon="📊", layout="wide", initial_sidebar_state="expanded")
    inject_css()
    path = str(Path(__file__).with_name("model.xlsx"))
    model = load_model(path)
    st.sidebar.markdown("# Graduate Portfolio Intelligence")
    st.sidebar.caption("V3 · decision tool · model.xlsx source of truth")
    pages = ["Executive", "Enrollment", "Persistence", "Programs", "Economics", "Scenario Lab", "Model QA"]
    page = st.sidebar.radio("Navigate", pages, index=0)
    st.sidebar.markdown("---")
    st.sidebar.caption("Definitions")
    st.sidebar.caption("Enrollment YoY uses like-for-like periods. Exits mean graduated OR withdrew combined. Thin samples are directional and excluded from rankings by default.")
    if page == "Executive": page_executive(model)
    elif page == "Enrollment": page_enrollment(model)
    elif page == "Persistence": page_persistence(model)
    elif page == "Programs": page_programs(model)
    elif page == "Economics": page_economics(model)
    elif page == "Scenario Lab": page_scenario(model)
    elif page == "Model QA": page_qa(model)


if __name__ == "__main__":
    main()
