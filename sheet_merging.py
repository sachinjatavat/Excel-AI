"""
sheet_merging.py - Advanced Sheet Merging Engine for Excel AI Copilot
Handles:
1. Matching records across different sheets/workbooks using Names, IDs, Emails, or Phone numbers.
2. Adding selected columns from another sheet into the active sheet.
3. Flagging conflicting values between sheets when a column exists in both with different values.
4. Flagging unmatched records (records in Target not found in Source, and records in Source not in Target).
5. Comprehensive Merge Analysis & Preview report before applying.
"""

import re
import difflib
from typing import List, Any, Dict, Tuple, Optional, Set


def normalize_key_value(val: Any) -> str:
    """Normalizes string for matching (lowercased, spaces stripped, punctuation normalized)."""
    if val is None:
        return ""
    s = str(val).strip().lower()
    # Normalize email
    if "@" in s:
        return s
    # Normalize phone: extract numbers
    digits = re.sub(r"\D", "", s)
    if len(digits) >= 7:
        return digits
    # General string: collapse spaces
    return re.sub(r"\s+", " ", s)


KEY_SYNONYMS = {
    "email": ["email", "mail", "e-mail", "e_mail"],
    "id": ["id", "emp id", "empid", "emp_id", "employee id", "employee_id", "roll", "roll no", "roll_no", "code", "badge", "s.no"],
    "phone": ["phone", "mobile", "contact", "cell", "tel", "phone number"],
    "name": ["name", "student name", "employee name", "full name", "person", "client", "customer"]
}


def detect_key_concept(header: str) -> Optional[str]:
    h = header.strip().lower()
    for concept, terms in KEY_SYNONYMS.items():
        if any(term == h or term in h for term in terms):
            return concept
    return None


def find_best_key_pair(target_headers: List[str], source_headers: List[str], preferred_concept: Optional[str] = None) -> Tuple[int, int]:
    """Finds best matching key column indices (target_idx, source_idx)."""
    t_concepts = {detect_key_concept(h): idx for idx, h in enumerate(target_headers) if detect_key_concept(h)}
    s_concepts = {detect_key_concept(h): idx for idx, h in enumerate(source_headers) if detect_key_concept(h)}

    if preferred_concept:
        pref = preferred_concept.strip().lower()
        # Find matching concept
        for concept in KEY_SYNONYMS:
            if pref == concept or any(t in pref for t in KEY_SYNONYMS[concept]):
                if concept in t_concepts and concept in s_concepts:
                    return t_concepts[concept], s_concepts[concept]

    # Priority order: id, email, phone, name
    for priority in ["id", "email", "phone", "name"]:
        if priority in t_concepts and priority in s_concepts:
            return t_concepts[priority], s_concepts[priority]

    # Fallback to direct name match
    for t_idx, t_h in enumerate(target_headers):
        for s_idx, s_h in enumerate(source_headers):
            if t_h.strip().lower() == s_h.strip().lower():
                return t_idx, s_idx

    return 0, 0


def analyze_sheet_merge(
    target_grid: List[List[Any]],
    source_grid: List[List[Any]],
    key_column_target: Optional[str] = None,
    key_column_source: Optional[str] = None,
    selected_columns: Optional[List[str]] = None,
    overwrite_conflicts: bool = False
) -> Dict[str, Any]:
    """
    Analyzes merge between target_grid (current sheet) and source_grid (second sheet/workbook).
    Returns matched count, unmatched counts, conflicting fields, and proposed merged grid.
    """
    if not target_grid or len(target_grid) == 0:
        return {"error": "Target worksheet is empty."}
    if not source_grid or len(source_grid) == 0:
        return {"error": "Source sheet is empty."}

    target_headers = [str(c).strip() for c in target_grid[0]]
    source_headers = [str(c).strip() for c in source_grid[0]]

    # 1. Determine key column pair for target & source
    target_key_idx, source_key_idx = find_best_key_pair(
        target_headers, source_headers, key_column_target or key_column_source
    )
    target_key_name = target_headers[target_key_idx]
    source_key_name = source_headers[source_key_idx]

    # Index source rows by normalized key
    source_index: Dict[str, Tuple[int, List[Any]]] = {}
    for r_idx in range(1, len(source_grid)):
        row = source_grid[r_idx]
        key_val = normalize_key_value(row[source_key_idx]) if source_key_idx < len(row) else ""
        if key_val:
            source_index[key_val] = (r_idx, row)

    # Determine which columns to import from source
    # If selected_columns specified, filter to those; otherwise pick columns that don't already exist in target
    source_col_map: List[Tuple[int, str]] = [] # (source_col_index, source_col_header)
    target_headers_lower = [h.lower() for h in target_headers]

    for c_idx, h in enumerate(source_headers):
        if c_idx == source_key_idx:
            continue
        h_clean = h.strip()
        if selected_columns:
            if any(sel.lower() == h_clean.lower() or sel.lower() in h_clean.lower() for sel in selected_columns):
                source_col_map.append((c_idx, h_clean))
        else:
            # Default: add columns not in target, or if specified
            if h_clean.lower() not in target_headers_lower:
                source_col_map.append((c_idx, h_clean))

    # If no new columns detected, import all non-key columns from source
    if not source_col_map:
        for c_idx, h in enumerate(source_headers):
            if c_idx != source_key_idx:
                source_col_map.append((c_idx, h.strip()))

    # Track matches, unmatched, and conflicts
    matched_target_rows = 0
    unmatched_target_rows = []
    matched_source_keys: Set[str] = set()
    conflicts = []

    # Check for shared columns (columns existing in both target and source)
    shared_cols = []
    for s_c_idx, s_h in enumerate(source_headers):
        if s_c_idx == source_key_idx:
            continue
        for t_c_idx, t_h in enumerate(target_headers):
            if t_c_idx == target_key_idx:
                continue
            if s_h.lower() == t_h.lower():
                shared_cols.append((t_c_idx, s_c_idx, t_h))

    # Build merged grid
    # New columns added to headers
    merged_headers = list(target_headers)
    for _, s_h in source_col_map:
        if s_h.lower() not in target_headers_lower:
            merged_headers.append(s_h)

    merged_rows = [merged_headers]

    for r_idx in range(1, len(target_grid)):
        t_row = list(target_grid[r_idx])
        t_key = normalize_key_value(t_row[target_key_idx]) if target_key_idx < len(t_row) else ""
        
        # Check fuzzy match if exact match fails
        matched_source = None
        if t_key in source_index:
            matched_source = source_index[t_key]
            matched_source_keys.add(t_key)
        elif len(t_key) > 3:
            # Try fuzzy match on keys
            best_match = difflib.get_close_matches(t_key, list(source_index.keys()), n=1, cutoff=0.85)
            if best_match:
                matched_source = source_index[best_match[0]]
                matched_source_keys.add(best_match[0])

        if matched_source is not None:
            matched_target_rows += 1
            s_r_idx, s_row = matched_source

            # Check conflicts on shared columns
            for t_c_idx, s_c_idx, col_name in shared_cols:
                t_val = str(t_row[t_c_idx]).strip() if t_c_idx < len(t_row) else ""
                s_val = str(s_row[s_c_idx]).strip() if s_c_idx < len(s_row) else ""
                if t_val and s_val and t_val.lower() != s_val.lower():
                    conflicts.append({
                        "target_row": r_idx,
                        "source_row": s_r_idx,
                        "key": str(t_row[target_key_idx]),
                        "column": col_name,
                        "target_value": t_val,
                        "source_value": s_val,
                    })
                    if overwrite_conflicts:
                        t_row[t_c_idx] = s_val

            # Append new columns
            for s_c_idx, s_h in source_col_map:
                if s_h.lower() not in target_headers_lower:
                    val = s_row[s_c_idx] if s_c_idx < len(s_row) else ""
                    t_row.append(val)
        else:
            unmatched_target_rows.append({
                "row": r_idx,
                "key": str(t_row[target_key_idx]) if target_key_idx < len(t_row) else "",
            })
            # Pad empty values for new columns
            for _, s_h in source_col_map:
                if s_h.lower() not in target_headers_lower:
                    t_row.append("")

        merged_rows.append(t_row)

    # Determine unmatched source rows
    unmatched_source_rows = []
    for s_key, (s_r_idx, s_row) in source_index.items():
        if s_key not in matched_source_keys:
            unmatched_source_rows.append({
                "source_row": s_r_idx,
                "key": str(s_row[source_key_idx]) if source_key_idx < len(s_row) else s_key,
            })

    columns_added_names = [s_h for _, s_h in source_col_map if s_h.lower() not in target_headers_lower]

    return {
        "status": "success",
        "key_column_target": target_key_name,
        "key_column_source": source_key_name,
        "columns_added": columns_added_names,
        "matched_count": matched_target_rows,
        "unmatched_target_count": len(unmatched_target_rows),
        "unmatched_source_count": len(unmatched_source_rows),
        "conflicts_count": len(conflicts),
        "conflicts": conflicts[:20], # top 20 for preview
        "unmatched_target_rows": unmatched_target_rows[:15],
        "unmatched_source_rows": unmatched_source_rows[:15],
        "merged_grid": merged_rows,
    }


def parse_merge_prompt(prompt: str, target_headers: List[str]) -> Tuple[Optional[str], Optional[List[str]]]:
    """
    Extracts key column name and target columns to add from user natural language prompt.
    e.g. "Match employees by email and add their salaries from the second workbook"
    -> key = "email", columns = ["salary", "salaries"]
    """
    p_lower = prompt.lower()
    
    # 1. Extract key column
    key_match = re.search(r"(?:match|join|merge|by|using|on)\s+(?:the\s+)?(?:column\s+|attribute\s+)?([a-zA-Z0-9_\s]+?)(?:\s+and|\s+then|\s+to|\s+from|$)", p_lower)
    key_candidate = None
    if key_match:
        cand = key_match.group(1).strip()
        for k in ["email", "id", "roll", "phone", "name", "employee", "code"]:
            if k in cand:
                key_candidate = k
                break

    # 2. Extract columns to add
    cols_candidate = []
    add_match = re.search(r"(?:add|include|import|bring|append|merge)\s+(?:their\s+|the\s+)?([a-zA-Z0-9_,\s]+?)(?:\s+from|\s+column|\s+to|$)", p_lower)
    if add_match:
        raw_cols = add_match.group(1).strip()
        parts = re.split(r",|\band\b", raw_cols)
        for part in parts:
            c = part.strip()
            if c and c not in ["data", "sheet", "record", "records", "column", "columns", "second", "another"]:
                cols_candidate.append(c)

    return key_candidate, cols_candidate if cols_candidate else None
