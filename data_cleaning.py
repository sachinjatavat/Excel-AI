"""
data_cleaning.py - Smart Data Cleaning Engine for Excel AI Copilot
Handles:
1. Duplicate detection and removal (exact row or key-based).
2. Empty row detection and removal.
3. Inconsistent names standardization (Title Case, excess whitespace, stray symbols).
4. Inconsistent dates standardization (mixed formats -> YYYY-MM-DD).
5. Inconsistent phone numbers standardization (clean (XXX) XXX-XXXX or international).
6. Inconsistent capitalization & whitespace trimming.
7. Missing and placeholder value detection ("N/A", "null", "-", etc.).
8. Comprehensive Preview & Audit mode: "Clean this sheet, but don't delete any data without showing me first".
"""

import re
from datetime import datetime
from typing import List, Any, Dict, Tuple, Optional, Set


# Common placeholders representing missing data
PLACEHOLDER_VALUES = {
    "n/a", "na", "null", "none", "-", "--", "---", "nil", "undefined",
    "unknown", "?", "???", "#n/a", "#null!", "blank", "empty"
}


def is_empty_cell(val: Any) -> bool:
    if val is None:
        return True
    s = str(val).strip()
    return s == ""


def is_placeholder_value(val: Any) -> bool:
    if val is None:
        return True
    s = str(val).strip().lower()
    return s in PLACEHOLDER_VALUES


def is_empty_row(row: List[Any]) -> bool:
    return all(is_empty_cell(c) for c in row)


def clean_name_string(val: str) -> str:
    """Standardizes names to Title Case and cleans excessive whitespace."""
    if not val:
        return ""
    # Strip non-alphanumeric junk at ends but preserve internal hyphens/apostrophes
    cleaned = re.sub(r"[\t\r\n]+", " ", str(val))
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    
    # Capitalize parts properly (handling hyphens, apostrophes e.g. O'Connor, Mary-Jane)
    def title_word(word: str) -> str:
        if not word:
            return ""
        if "'" in word:
            return "'".join(w.capitalize() for w in word.split("'"))
        if "-" in word:
            return "-".join(w.capitalize() for w in word.split("-"))
        return word.capitalize()

    return " ".join(title_word(w) for w in cleaned.split(" "))


def parse_and_standardize_date(val: Any) -> Optional[str]:
    """
    Attempts to parse mixed date strings into standardized ISO format: YYYY-MM-DD.
    Returns standardized string if parsed, or None if not recognized as date.
    """
    if val is None or isinstance(val, (int, float)):
        return None
    s = str(val).strip()
    if len(s) < 6 or len(s) > 30:
        return None
    
    # Quick filter: must contain date separators (- / . space)
    if not re.search(r"[-/\.\s]", s):
        return None

    # Already valid ISO format YYYY-MM-DD
    if re.match(r"^\d{4}-\d{2}-\d{2}$", s):
        try:
            datetime.strptime(s, "%Y-%m-%d")
            return s
        except ValueError:
            return None

    # Common formats to try
    formats = [
        "%Y/%m/%d", "%Y.%m.%d",
        "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y",
        "%m/%d/%Y", "%m-%d-%Y",
        "%d-%b-%Y", "%d-%B-%Y", "%d %b %Y", "%d %B %Y",
        "%b %d, %Y", "%B %d, %Y", "%b %d %Y",
        "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S"
    ]

    for fmt in formats:
        try:
            dt = datetime.strptime(s, fmt)
            # Filter out extreme bogus years
            if 1900 <= dt.year <= 2100:
                return dt.strftime("%Y-%m-%d")
        except (ValueError, TypeError):
            continue

    # Regex heuristic for DD-MM-YYYY or MM-DD-YYYY
    match_3part = re.match(r"^(\d{1,2})[-/\.](\d{1,2})[-/\.](\d{2,4})$", s)
    if match_3part:
        p1, p2, p3 = int(match_3part.group(1)), int(match_3part.group(2)), int(match_3part.group(3))
        year = p3 if p3 >= 100 else (2000 + p3 if p3 < 50 else 1900 + p3)
        if 1900 <= year <= 2100:
            if p1 > 12 and 1 <= p2 <= 12: # p1 is day, p2 is month
                try:
                    dt = datetime(year, p2, p1)
                    return dt.strftime("%Y-%m-%d")
                except ValueError:
                    pass
            elif p2 > 12 and 1 <= p1 <= 12: # p1 is month, p2 is day
                try:
                    dt = datetime(year, p1, p2)
                    return dt.strftime("%Y-%m-%d")
                except ValueError:
                    pass
            elif 1 <= p1 <= 31 and 1 <= p2 <= 12:
                # Default assume DD-MM-YYYY
                try:
                    dt = datetime(year, p2, p1)
                    return dt.strftime("%Y-%m-%d")
                except ValueError:
                    pass

    return None


def parse_and_standardize_phone(val: Any) -> Optional[str]:
    """
    Standardizes phone numbers:
    - 10 digits -> (XXX) XXX-XXXX
    - 11 digits with 1 -> +1 (XXX) XXX-XXXX
    - Indian 10 digits with +91 -> +91 XXXXX XXXXX
    """
    if val is None:
        return None
    s = str(val).strip()
    # Strip common formatting
    digits = re.sub(r"[^\d+]", "", s)
    has_plus = digits.startswith("+")
    num_only = re.sub(r"\D", "", digits)

    if len(num_only) == 10:
        return f"({num_only[:3]}) {num_only[3:6]}-{num_only[6:]}"
    elif len(num_only) == 11 and num_only.startswith("1"):
        return f"+1 ({num_only[1:4]}) {num_only[4:7]}-{num_only[7:]}"
    elif len(num_only) == 12 and num_only.startswith("91"):
        return f"+91 {num_only[2:7]} {num_only[7:]}"
    elif len(num_only) >= 7 and len(num_only) <= 15 and has_plus:
        return f"+{num_only}"
    return None


def detect_column_types(grid: List[List[Any]]) -> Dict[int, str]:
    """
    Infers column types: 'name', 'date', 'phone', 'email', 'number', 'text'.
    """
    if not grid or len(grid) < 2:
        return {}
    
    headers = [str(h).lower() for h in grid[0]]
    types: Dict[int, str] = {}

    for c in range(len(grid[0])):
        h = headers[c] if c < len(headers) else ""
        
        # Check header names first
        if any(k in h for k in ["name", "student", "employee", "customer", "person", "client"]):
            types[c] = "name"
            continue
        if any(k in h for k in ["date", "dob", "birth", "joining", "hire", "timestamp"]):
            types[c] = "date"
            continue
        if any(k in h for k in ["phone", "mobile", "contact", "cell", "tel"]):
            types[c] = "phone"
            continue
        if any(k in h for k in ["email", "mail"]):
            types[c] = "email"
            continue
        
        # Check data samples
        samples = [grid[r][c] for r in range(1, min(len(grid), 15)) if c < len(grid[r]) and not is_empty_cell(grid[r][c])]
        if not samples:
            types[c] = "text"
            continue

        # Test date sample
        date_hits = sum(1 for s in samples if parse_and_standardize_date(s) is not None)
        if date_hits / len(samples) >= 0.5:
            types[c] = "date"
            continue

        # Test phone sample
        phone_hits = sum(1 for s in samples if parse_and_standardize_phone(s) is not None)
        if phone_hits / len(samples) >= 0.5:
            types[c] = "phone"
            continue

        # Test email sample
        email_hits = sum(1 for s in samples if "@" in str(s) and "." in str(s))
        if email_hits / len(samples) >= 0.5:
            types[c] = "email"
            continue

        types[c] = "text"

    return types


def analyze_data_cleaning(
    grid: List[List[Any]],
    remove_duplicates: bool = True,
    remove_empty: bool = True,
    standardize_names: bool = True,
    standardize_dates: bool = True,
    standardize_phones: bool = True,
    clean_placeholders: bool = True,
) -> Dict[str, Any]:
    """
    Scans the grid and returns a comprehensive audit and preview of all issues found
    WITHOUT altering any original data.
    """
    if not grid or len(grid) == 0:
        return {
            "total_issues_found": 0,
            "duplicate_rows_count": 0,
            "empty_rows_count": 0,
            "name_fixes_count": 0,
            "date_fixes_count": 0,
            "phone_fixes_count": 0,
            "placeholder_fixes_count": 0,
            "changes": [],
            "cleaned_grid": grid,
        }

    headers = [str(c) for c in grid[0]]
    col_types = detect_column_types(grid)
    num_cols = len(headers)

    duplicate_row_indices = []
    empty_row_indices = []
    changes = []
    seen_row_signatures: Dict[str, int] = {}

    name_fixes = 0
    date_fixes = 0
    phone_fixes = 0
    placeholder_fixes = 0

    # Build new cleaned rows
    cleaned_rows = [list(grid[0])]

    for r_idx in range(1, len(grid)):
        raw_row = grid[r_idx]
        
        # 1. Empty Row Check
        if is_empty_row(raw_row):
            empty_row_indices.append(r_idx)
            changes.append({
                "type": "delete_row",
                "row": r_idx,
                "reason": "Completely empty row",
                "old": "[Empty Row]",
                "new": "[Deleted]",
            })
            continue

        # Clean individual cells
        cleaned_row = []
        for c_idx in range(num_cols):
            val = raw_row[c_idx] if c_idx < len(raw_row) else ""
            header_name = headers[c_idx] if c_idx < len(headers) else f"Column {c_idx+1}"
            ctype = col_types.get(c_idx, "text")

            # Placeholder check
            if clean_placeholders and is_placeholder_value(val):
                if val != "":
                    placeholder_fixes += 1
                    changes.append({
                        "type": "placeholder_fix",
                        "row": r_idx,
                        "col": c_idx,
                        "header": header_name,
                        "old": str(val),
                        "new": "",
                        "reason": f"Removed placeholder missing value '{val}'",
                    })
                    cleaned_row.append("")
                    continue

            # Standardize names
            if standardize_names and ctype == "name" and not is_empty_cell(val):
                c_name = clean_name_string(str(val))
                if c_name != str(val):
                    name_fixes += 1
                    changes.append({
                        "type": "name_fix",
                        "row": r_idx,
                        "col": c_idx,
                        "header": header_name,
                        "old": str(val),
                        "new": c_name,
                        "reason": "Standardized to Proper Title Case & removed extra spaces",
                    })
                    cleaned_row.append(c_name)
                    continue

            # Standardize dates
            if standardize_dates and ctype == "date" and not is_empty_cell(val):
                std_date = parse_and_standardize_date(val)
                if std_date and std_date != str(val):
                    date_fixes += 1
                    changes.append({
                        "type": "date_fix",
                        "row": r_idx,
                        "col": c_idx,
                        "header": header_name,
                        "old": str(val),
                        "new": std_date,
                        "reason": f"Standardized date to ISO format ({std_date})",
                    })
                    cleaned_row.append(std_date)
                    continue

            # Standardize phones
            if standardize_phones and ctype == "phone" and not is_empty_cell(val):
                std_phone = parse_and_standardize_phone(val)
                if std_phone and std_phone != str(val):
                    phone_fixes += 1
                    changes.append({
                        "type": "phone_fix",
                        "row": r_idx,
                        "col": c_idx,
                        "header": header_name,
                        "old": str(val),
                        "new": std_phone,
                        "reason": f"Standardized phone number format ({std_phone})",
                    })
                    cleaned_row.append(std_phone)
                    continue

            # General string cleanup (trim trailing/leading whitespace)
            if isinstance(val, str):
                trimmed = re.sub(r"\s+", " ", val).strip()
                if trimmed != val and not is_empty_cell(val):
                    changes.append({
                        "type": "whitespace_fix",
                        "row": r_idx,
                        "col": c_idx,
                        "header": header_name,
                        "old": val,
                        "new": trimmed,
                        "reason": "Trimmed redundant whitespace",
                    })
                    cleaned_row.append(trimmed)
                    continue

            cleaned_row.append(val)

        # 2. Duplicate Row Check (on cleaned values)
        row_signature = "|".join(str(c).lower().strip() for c in cleaned_row)
        if row_signature in seen_row_signatures:
            original_row = seen_row_signatures[row_signature]
            duplicate_row_indices.append(r_idx)
            changes.append({
                "type": "delete_row",
                "row": r_idx,
                "reason": f"Exact duplicate of row {original_row}",
                "old": f"Duplicate row: {cleaned_row[:3]}...",
                "new": "[Deleted]",
            })
            continue

        seen_row_signatures[row_signature] = r_idx
        cleaned_rows.append(cleaned_row)

    total_issues = (
        len(duplicate_row_indices)
        + len(empty_row_indices)
        + name_fixes
        + date_fixes
        + phone_fixes
        + placeholder_fixes
    )

    return {
        "total_issues_found": total_issues,
        "duplicate_rows_count": len(duplicate_row_indices),
        "duplicate_row_indices": duplicate_row_indices,
        "empty_rows_count": len(empty_row_indices),
        "empty_row_indices": empty_row_indices,
        "name_fixes_count": name_fixes,
        "date_fixes_count": date_fixes,
        "phone_fixes_count": phone_fixes,
        "placeholder_fixes_count": placeholder_fixes,
        "changes": changes,
        "cleaned_grid": cleaned_rows,
    }


def execute_smart_clean(
    grid: List[List[Any]],
    prompt: str = ""
) -> Tuple[List[List[Any]], str, Dict[str, Any]]:
    """
    High-level entry point for smart data cleaning.
    Detects if user requested preview/approval ("don't delete without showing me first").
    """
    p_lower = prompt.lower()
    preview_requested = any(k in p_lower for k in [
        "don't delete", "dont delete", "without showing", "show me first",
        "preview", "review first", "check first", "ask me first"
    ])

    report = analyze_data_cleaning(grid)
    
    parts = []
    if report["duplicate_rows_count"] > 0:
        parts.append(f"{report['duplicate_rows_count']} duplicate rows")
    if report["empty_rows_count"] > 0:
        parts.append(f"{report['empty_rows_count']} empty rows")
    if report["name_fixes_count"] > 0:
        parts.append(f"{report['name_fixes_count']} name capitalizations")
    if report["date_fixes_count"] > 0:
        parts.append(f"{report['date_fixes_count']} date formats")
    if report["phone_fixes_count"] > 0:
        parts.append(f"{report['phone_fixes_count']} phone numbers")
    if report["placeholder_fixes_count"] > 0:
        parts.append(f"{report['placeholder_fixes_count']} missing/placeholder cells")

    if report["total_issues_found"] == 0:
        msg = "Worksheet is already clean! No duplicates, empty rows, or formatting inconsistencies found."
        return grid, msg, report

    summary_str = ", ".join(parts) if parts else "formatting adjustments"
    
    if preview_requested:
        msg = f"Found {report['total_issues_found']} cleaning issues ({summary_str}). Review preview below before applying."
        report["requires_confirmation"] = True
        return grid, msg, report

    # Otherwise return cleaned grid
    msg = f"Cleaned sheet: resolved {report['total_issues_found']} issues ({summary_str})."
    return report["cleaned_grid"], msg, report

