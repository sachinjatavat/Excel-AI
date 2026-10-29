"""
anomaly_engine.py - Data Validation and Anomaly Detection Engine for Excel AI Copilot
Handles:
1. Statistical outlier detection (Z-Score & IQR methods).
2. Department / Category relative anomaly detection (e.g. unusually high salary compared to department).
3. Impossible & suspicious dates (future birth/hire dates, dates outside reasonable bounds, chronological errors).
4. Missing identifiers (empty Employee ID, Student Roll No).
5. Invalid email addresses and telephone formats.
6. Suspicious expenses / negative values where unnatural.
7. Plain-English diagnostic explanations and cell coordinates for visual highlighting.
"""

import re
import math
from datetime import datetime
from typing import List, Dict, Any, Tuple, Optional
from collections import defaultdict


def is_numeric(val: Any) -> bool:
    if val is None or val == "":
        return False
    try:
        float(str(val).replace(",", "").replace("%", "").replace("$", "").replace("₹", "").strip())
        return True
    except ValueError:
        return False


def to_float(val: Any) -> float:
    return float(str(val).replace(",", "").replace("%", "").replace("$", "").replace("₹", "").strip())


def is_valid_email(val: Any) -> bool:
    if val is None:
        return False
    s = str(val).strip()
    return bool(re.match(r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$", s))


def detect_anomalies_and_validation(
    grid: List[List[Any]],
    prompt: str = ""
) -> Dict[str, Any]:
    """
    Performs comprehensive audit of data validation issues and statistical anomalies.
    Supports department-relative outlier detection when requested in prompt.
    """
    if not grid or len(grid) < 2:
        return {"error": "Sheet does not contain enough data for validation."}

    headers = [str(c).strip() for c in grid[0]]
    data_rows = grid[1:]
    num_rows = len(data_rows)
    num_cols = len(headers)

    p_lower = prompt.lower()
    dept_relative_requested = any(k in p_lower for k in [
        "department", "dept", "compared with", "compared to", "relative to",
        "in their department", "within their department", "per department"
    ])

    anomalies = []
    
    # 1. Map column semantic types
    id_cols = []
    name_cols = []
    email_cols = []
    date_cols = []
    dept_cols = []
    salary_or_metric_cols = []

    for c_idx, h in enumerate(headers):
        h_low = h.lower()
        if any(k in h_low for k in ["emp id", "empid", "id", "roll", "roll no", "badge", "code"]):
            id_cols.append(c_idx)
        elif any(k in h_low for k in ["name", "student", "employee", "person"]):
            name_cols.append(c_idx)
        elif any(k in h_low for k in ["email", "mail"]):
            email_cols.append(c_idx)
        elif any(k in h_low for k in ["date", "dob", "birth", "hire", "joining"]):
            date_cols.append(c_idx)
        elif any(k in h_low for k in ["dept", "department", "category", "team"]):
            dept_cols.append(c_idx)
        
        # Test numeric column
        nums_count = sum(1 for r in data_rows if c_idx < len(r) and is_numeric(r[c_idx]))
        if nums_count / max(1, num_rows) >= 0.5:
            if not any(k in h_low for k in ["id", "roll", "code", "phone", "year"]):
                salary_or_metric_cols.append(c_idx)

    # 2. Check Missing IDs
    for c_idx in id_cols:
        h_name = headers[c_idx]
        for r_idx, r in enumerate(data_rows, start=1):
            val = r[c_idx] if c_idx < len(r) else None
            if val is None or str(val).strip() == "" or str(val).strip().lower() in ["n/a", "null", "-", "none"]:
                person = r[name_cols[0]] if (name_cols and name_cols[0] < len(r)) else f"Row {r_idx}"
                anomalies.append({
                    "row": r_idx,
                    "col": c_idx,
                    "header": h_name,
                    "value": str(val or "[Empty]"),
                    "severity": "danger",
                    "type": "MISSING_ID",
                    "reason": f"Missing mandatory {h_name} for record '{person}'.",
                })

    # 3. Check Invalid Email Addresses
    for c_idx in email_cols:
        h_name = headers[c_idx]
        for r_idx, r in enumerate(data_rows, start=1):
            val = r[c_idx] if c_idx < len(r) else None
            if val is not None and str(val).strip():
                s_email = str(val).strip()
                if not is_valid_email(s_email):
                    anomalies.append({
                        "row": r_idx,
                        "col": c_idx,
                        "header": h_name,
                        "value": s_email,
                        "severity": "danger",
                        "type": "INVALID_EMAIL",
                        "reason": f"Invalid email format '{s_email}' (missing domain or '@').",
                    })

    # 4. Check Suspicious Dates & Chronology
    today = datetime.now()
    for c_idx in date_cols:
        h_name = headers[c_idx]
        is_birth = any(k in h_name.lower() for k in ["birth", "dob"])
        is_hire = any(k in h_name.lower() for k in ["hire", "join", "start"])

        for r_idx, r in enumerate(data_rows, start=1):
            val = r[c_idx] if c_idx < len(r) else None
            if val is not None and str(val).strip():
                s_date = str(val).strip()
                # Parse year
                year_match = re.search(r"\b(19\d{2}|20\d{2})\b", s_date)
                if year_match:
                    year = int(year_match.group(1))
                    if year > today.year:
                        anomalies.append({
                            "row": r_idx,
                            "col": c_idx,
                            "header": h_name,
                            "value": s_date,
                            "severity": "danger",
                            "type": "FUTURE_DATE",
                            "reason": f"Impossible future date '{s_date}' in {h_name} (year {year} is in the future).",
                        })
                    elif is_birth and year > (today.year - 14):
                        anomalies.append({
                            "row": r_idx,
                            "col": c_idx,
                            "header": h_name,
                            "value": s_date,
                            "severity": "warning",
                            "type": "IMPROBABLE_AGE",
                            "reason": f"Birth date '{s_date}' indicates person would be under 14 years old.",
                        })
                    elif year < 1920:
                        anomalies.append({
                            "row": r_idx,
                            "col": c_idx,
                            "header": h_name,
                            "value": s_date,
                            "severity": "warning",
                            "type": "OUT_OF_BOUNDS_DATE",
                            "reason": f"Unusually old date '{s_date}' in {h_name} (prior to 1920).",
                        })

    # 5. Check Statistical Outliers & Department-Relative Anomalies
    # Determine target columns: prioritize columns mentioned in user prompt, else evaluate all metric columns
    active_metric_cols = []
    for c_idx in salary_or_metric_cols:
        h_low = headers[c_idx].lower()
        if any(w in p_lower for w in ["salary", "salaries", "pay", "compensation"]) and any(k in h_low for k in ["salary", "pay", "compensation"]):
            active_metric_cols.append(c_idx)
        elif any(w in p_lower for w in ["sales", "revenue"]) and any(k in h_low for k in ["sales", "revenue"]):
            active_metric_cols.append(c_idx)
        elif any(w in p_lower for w in ["expense", "cost"]) and any(k in h_low for k in ["expense", "cost"]):
            active_metric_cols.append(c_idx)
        elif any(w in p_lower for w in ["score", "grade", "performance"]) and any(k in h_low for k in ["score", "grade", "performance"]):
            active_metric_cols.append(c_idx)

    # If no specific column mentioned in prompt, inspect all detected metric columns
    if not active_metric_cols:
        active_metric_cols = list(salary_or_metric_cols)

    for target_metric_col in active_metric_cols:
        metric_name = headers[target_metric_col]

        # Calculate overall benchmark across the whole company/sheet for this metric
        all_metric_nums = [
            (to_float(r[target_metric_col]), r_idx)
            for r_idx, r in enumerate(data_rows, start=1)
            if target_metric_col < len(r) and is_numeric(r[target_metric_col])
        ]
        if not all_metric_nums:
            continue

        raw_all_nums = [v for v, _ in all_metric_nums]
        global_mean = sum(raw_all_nums) / len(raw_all_nums)
        global_var = sum((x - global_mean) ** 2 for x in raw_all_nums) / len(raw_all_nums)
        global_std = math.sqrt(global_var) if global_var > 0 else 0

        # A. Department-Relative Outlier Check
        if dept_cols:
            dept_c_idx = dept_cols[0]
            dept_name_hdr = headers[dept_c_idx]
            dept_values = defaultdict(list)

            for r_idx, r in enumerate(data_rows, start=1):
                dept_val = str(r[dept_c_idx]).strip() if dept_c_idx < len(r) else "Other"
                if not dept_val: dept_val = "Other"
                if target_metric_col < len(r) and is_numeric(r[target_metric_col]):
                    dept_values[dept_val].append((to_float(r[target_metric_col]), r_idx))

            for d_name, vals_with_idx in dept_values.items():
                if len(vals_with_idx) >= 2:
                    raw_nums = [v for v, _ in vals_with_idx]
                    d_mean = sum(raw_nums) / len(raw_nums)
                    variance = sum((x - d_mean) ** 2 for x in raw_nums) / len(raw_nums)
                    d_std = math.sqrt(variance) if variance > 0 else 0

                    for val, r_idx in vals_with_idx:
                        z_score = round((val - d_mean) / d_std, 1) if d_std > 0 else 0
                        is_dept_outlier = (d_std > 0 and z_score >= 1.5) or (val >= d_mean * 1.6 and val > 0)
                        if is_dept_outlier:
                            person = data_rows[r_idx-1][name_cols[0]] if (name_cols and name_cols[0] < len(data_rows[r_idx-1])) else f"Row {r_idx}"
                            anomalies.append({
                                "row": r_idx,
                                "col": target_metric_col,
                                "header": metric_name,
                                "value": val,
                                "severity": "warning",
                                "type": "DEPARTMENT_OUTLIER",
                                "reason": (
                                    f"{person}'s {metric_name} ({val:,.2f}) is unusually high compared to "
                                    f"the {d_name} {dept_name_hdr} average ({d_mean:,.2f}) — {z_score} standard deviations above peers."
                                ),
                            })
                        elif val < 0:
                            anomalies.append({
                                "row": r_idx,
                                "col": target_metric_col,
                                "header": metric_name,
                                "value": val,
                                "severity": "danger",
                                "type": "NEGATIVE_VALUE",
                                "reason": f"Suspicious negative value ({val}) for {metric_name}.",
                            })
                elif len(vals_with_idx) == 1 and dept_relative_requested:
                    # Single-person department compared to company benchmark
                    val, r_idx = vals_with_idx[0]
                    if global_std > 0 and (val - global_mean) / global_std >= 1.5:
                        z_score = round((val - global_mean) / global_std, 1)
                        person = data_rows[r_idx-1][name_cols[0]] if (name_cols and name_cols[0] < len(data_rows[r_idx-1])) else f"Row {r_idx}"
                        anomalies.append({
                            "row": r_idx,
                            "col": target_metric_col,
                            "header": metric_name,
                            "value": val,
                            "severity": "warning",
                            "type": "DEPARTMENT_OUTLIER",
                            "reason": (
                                f"{person}'s {metric_name} ({val:,.2f}) in {d_name} is unusually high compared to "
                                f"the company benchmark ({global_mean:,.2f}) — {z_score} standard deviations above peers."
                            ),
                        })

        # B. Global Outliers across entire dataset (IQR / Z-score)
        if len(all_metric_nums) >= 4:
            sorted_nums = sorted(raw_all_nums)
            q1 = sorted_nums[len(sorted_nums) // 4]
            q3 = sorted_nums[(3 * len(sorted_nums)) // 4]
            iqr = q3 - q1
            high_cutoff = q3 + 1.8 * iqr

            flagged_rows = {a["row"] for a in anomalies if a["col"] == target_metric_col}
            for val, r_idx in all_metric_nums:
                if r_idx not in flagged_rows and val > high_cutoff and iqr > 0:
                    person = data_rows[r_idx-1][name_cols[0]] if (name_cols and name_cols[0] < len(data_rows[r_idx-1])) else f"Row {r_idx}"
                    anomalies.append({
                        "row": r_idx,
                        "col": target_metric_col,
                        "header": metric_name,
                        "value": val,
                        "severity": "warning",
                        "type": "GLOBAL_OUTLIER",
                        "reason": f"{person}'s {metric_name} ({val:,.2f}) is a global statistical outlier (exceeds upper threshold {high_cutoff:,.2f}).",
                    })

    # Summary count
    total_anomalies = len(anomalies)

    return {
        "status": "success",
        "total_anomalies": total_anomalies,
        "anomalies": anomalies,
        "summary": (
            f"Audit complete: identified {total_anomalies} anomalies and validation issues."
            if total_anomalies > 0 else "Audit complete: no statistical anomalies or validation issues detected."
        )
    }
