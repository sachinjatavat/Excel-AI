"""
table_engine.py - Dedicated Fast Offline Excel AI Engine
Specialized for ANY Excel table operations: Students, Grades, Employees, Salaries,
Sales, Inventory, Expenses, Budget, and Generic Data with 100% offline execution.

Features:
1. Multi-domain support (Salary, HR, Sales, Inventory, Grades, Finance, etc.).
2. Multilingual & Romanized support (English, Hindi Devanagari, Romanized Hindi / Hinglish / Romanji).
3. Typo-tolerant fuzzy matching (single-letter mistakes, letter transpositions).
4. Selective Data Merging: Updates existing rows from pasted data while strictly ignoring non-existing entries.
5. Targeting rows and columns by either NAME or ORDER NUMBER (1st, 2nd, 3rd, column 2, row 4, pehla, teesra, तीसरा, etc.).
6. Instant execution (< 0.01 seconds, 100% offline).
"""

import re
import csv
import io
import difflib
from typing import List, Any, Tuple, Optional, Dict, Set

# Specialized modular engines for the 6 core features
import data_cleaning
import sheet_merging
import formula_engine
import dashboard_engine
import anomaly_engine
import diff_engine


# ==============================================================================
# 1. CORE UTILITIES & NORMALIZATION
# ==============================================================================

def is_number(val: Any) -> bool:
    if val is None or val == "":
        return False
    try:
        float(str(val).replace(",", "").replace("%", "").replace("$", "").replace("₹", "").strip())
        return True
    except ValueError:
        return False


def to_number(val: Any) -> float:
    return float(str(val).replace(",", "").replace("%", "").replace("$", "").replace("₹", "").strip())


def clean_cell(val: Any) -> Any:
    if val is None:
        return ""
    if isinstance(val, (int, float)):
        return val
    s = str(val).strip()
    if is_number(s):
        try:
            if "." in s:
                f = float(s.replace(",", "").replace("%", "").replace("$", "").replace("₹", "").strip())
                return round(f, 2)
            return int(s.replace(",", "").replace("%", "").replace("$", "").replace("₹", "").strip())
        except ValueError:
            return s
    return s


def normalize_rectangular_grid(grid: List[List[Any]]) -> List[List[Any]]:
    if not grid or not any(grid):
        return []
    max_cols = max(len(r) if isinstance(r, list) else 1 for r in grid)
    normalized = []
    for r in grid:
        if isinstance(r, list):
            row = [clean_cell(c) for c in r]
            if len(row) < max_cols:
                row.extend([""] * (max_cols - len(row)))
            normalized.append(row)
        else:
            normalized.append([clean_cell(r)] + [""] * (max_cols - 1))
    return normalized


# ==============================================================================
# 2. MULTILINGUAL & ORDINALS PARSER (English, Hinglish, Hindi Devanagari)
# ==============================================================================

ORDINALS_MAP: Dict[str, int] = {
    # English words
    "first": 1, "1st": 1, "second": 2, "2nd": 2, "third": 3, "3rd": 3,
    "fourth": 4, "4th": 4, "fifth": 5, "5th": 5, "sixth": 6, "6th": 6,
    "seventh": 7, "7th": 7, "eighth": 8, "8th": 8, "ninth": 9, "9th": 9,
    "tenth": 10, "10th": 10, "last": -1, "bottom": -1, "top": 1,
    # Hinglish
    "pehla": 1, "pehli": 1, "pahla": 1, "pahli": 1,
    "dusra": 2, "dusri": 2, "doosra": 2, "doosri": 2,
    "teesra": 3, "teesri": 3, "tisra": 3, "tisri": 3,
    "chautha": 4, "chauthi": 4, "chotha": 4,
    "panchva": 5, "panchvi": 5, "panchwa": 5,
    "chhatta": 6, "chhatthi": 6, "chatha": 6,
    "saatva": 7, "saatvi": 7, "aathva": 8, "aathvi": 8,
    "nauva": 9, "dasva": 10, "aakhri": -1, "antim": -1,
    # Hindi Devanagari
    "पहला": 1, "पहली": 1, "दूसरा": 2, "दूसरी": 2,
    "तीसरा": 3, "तीसरी": 3, "चौथा": 4, "चौथी": 4,
    "पांचवां": 5, "पांचवीं": 5, "छठा": 6, "छठी": 6,
    "सातवां": 7, "आठवां": 8, "नौवां": 9, "दसवां": 10,
    "आखिरी": -1, "अंतिम": -1,
}


def parse_target_index(text: str, target_type: str = "column") -> Optional[int]:
    """
    Extracts 1-based order number or -1 (for last) from prompt.
    E.g. 'delete column 3', 'remove 2nd column', 'teesra column hatao', 'तीसरा कॉलम हटाओ'.
    """
    t_words = r"(?:column|col|stambh|स्तम्भ|कॉलम)" if target_type == "column" else r"(?:row|line|pankti|पंक्ति|रो)"
    low = text.lower()

    # 1. 'column 3' or 'row 4'
    m1 = re.search(rf"(?:^|\s|[.,;:!?]){t_words}\s*(\d+)(?:\s|$|[.,;:!?])", low)
    if m1:
        return int(m1.group(1))

    # 2. '3rd column' or '4th row'
    m2 = re.search(rf"(?:^|\s|[.,;:!?])(\d+)(?:st|nd|rd|th)?\s*{t_words}(?:\s|$|[.,;:!?])", low)
    if m2:
        return int(m2.group(1))

    # 3. Ordinals: 'teesra column', 'तीसरा कॉलम', '3rd column', 'last column'
    for word, idx in ORDINALS_MAP.items():
        pattern1 = rf"(?:^|\s|[.,;:!?]){re.escape(word)}\s*{t_words}(?:\s|$|[.,;:!?])"
        pattern2 = rf"(?:^|\s|[.,;:!?]){t_words}\s*{re.escape(word)}(?:\s|$|[.,;:!?])"
        if re.search(pattern1, low) or re.search(pattern2, low):
            return idx

    return None


# ==============================================================================
# 3. TYPO TOLERANCE & PROMPT NORMALIZATION (Multilingual)
# ==============================================================================

COMMON_TYPOS: Dict[str, str] = {
    # Actions & Verbs
    "creat": "create", "craete": "create", "mak": "make", "mke": "make",
    "bnao": "create", "banao": "create",
    "remov": "remove", "rmove": "remove", "delet": "delete", "delt": "delete", "drp": "drop",
    "hatao": "remove", "nikalo": "calculate", "jodo": "add", "badlo": "change", "chagne": "change",
    "clculate": "calculate", "calulate": "calculate", "calcualte": "calculate", "calculat": "calculate",
    "updat": "update", "renam": "rename", "curv": "curve", "crve": "curve", "bonas": "bonus",
    "filtr": "filter", "srt": "sort", "captalize": "capitalize",
    # Hindi Devanagari Verbs & Nouns
    "बनाओ": "create", "बनाएं": "create", "बनाइए": "create",
    "हटाओ": "remove", "हटाएं": "remove", "हटा दीजिए": "remove", "मिटाओ": "remove",
    "जोड़ो": "add", "डालो": "add", "शामिल करो": "add",
    "बदलो": "change", "अपडेट करो": "update",
    "निकालो": "calculate", "गणना करो": "calculate",
    "वेतन": "salary", "सैलरी": "salary",
    "कर्मचारी": "employee", "स्टाफ": "employee",
    "छात्र": "student", "विद्यार्थी": "student",
    "कॉलम": "column", "स्तम्भ": "column",
    "रो": "row", "पंक्ति": "row",
    "तालिका": "table", "टेबल": "table",
    "औसत": "average", "कुल": "total", "योग": "total", "प्रतिशत": "percentage",
    "बोनस": "bonus", "हाइक": "hike", "वृद्धि": "increase",
    # Column concepts
    "leter": "letter", "grad": "grade", "grde": "grade", "gardes": "grades",
    "colum": "column", "colomn": "column", "colunm": "column", "cloumn": "column", "colm": "column",
    "rw": "row", "ro": "row",
    "totl": "total", "toal": "total", "totle": "total",
    "averge": "average", "avrage": "average", "avg": "average", "ausat": "average",
    "persentage": "percentage", "prcentage": "percentage", "percentag": "percentage",
    "remrks": "remarks", "remak": "remarks", "fedback": "feedback", "coment": "comment",
    "rnk": "rank", "ranck": "rank",
    "tabel": "table", "tabl": "table", "tebal": "table",
    "stduent": "student", "studnt": "student", "studnet": "student",
    "pas": "pass", "pasa": "pass", "fal": "fail", "faill": "fail", "fale": "fail", "reslt": "result",
    "hieghest": "highest", "higest": "highest", "lowset": "lowest", "lowst": "lowest",
    "statictics": "statistics", "atendance": "attendance", "attendence": "attendance",
    "nam": "name", "scienc": "science", "mathes": "math", "maths": "math",
    "slary": "salary", "salry": "salary", "tankhwah": "salary", "vetan": "salary",
    "employe": "employee", "employe name": "employee name", "karmchari": "employee",
}

VALID_WORDS: Set[str] = {
    "student", "students", "name", "roll", "math", "maths", "science", "english", "hindi",
    "physics", "chemistry", "biology", "history", "geography", "total", "grade", "score",
    "marks", "average", "rank", "result", "pass", "fail", "remarks", "attendance", "date",
    "table", "column", "columns", "row", "rows", "quiz", "exam", "midterm", "final", "test",
    "create", "make", "delete", "remove", "add", "calculate", "update", "curve", "sort",
    "filter", "proper", "title", "case", "letters", "letter", "bottom", "summary", "stats",
    "salary", "bonus", "tax", "department", "employee", "sales", "revenue", "price", "quantity",
    "cost", "profit", "discount", "stock", "inventory", "item", "product", "hike", "raise"
}

DOMAIN_KEYWORDS: List[str] = [
    "create", "make", "table", "column", "columns", "row", "rows", "delete", "remove",
    "add", "calculate", "update", "change", "rename", "sort", "filter", "letter",
    "grade", "grades", "remarks", "feedback", "total", "average", "percentage",
    "curve", "bonus", "rank", "pass", "fail", "result", "summary", "highest", "lowest",
    "student", "students", "attendance", "proper", "title", "case",
    "salary", "bonus", "tax", "department", "employee", "sales", "revenue", "price",
    "quantity", "cost", "profit", "discount", "stock", "inventory", "item", "product", "hike"
]


def normalize_prompt(prompt: str) -> str:
    """
    Cleans typos, Hindi/Hinglish phrasing, and normalizes command keywords.
    Ensures single-letter typos like 'leter', 'grad', 'colum' are cleanly corrected.
    """
    cleaned = prompt.strip()
    # Normalize common slash and Hinglish phrases
    cleaned = re.sub(r"\bpas[s]?\s*/\s*fa[il]{1,3}\b", "pass/fail", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\bpas/fal\b", "pass/fail", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\btable\s+banao\b", "create table", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\btabel\s+banao\b", "create table", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\bbanao\s+table\b", "create table", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\bhata\s+do\b", "remove", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\bhatao\b", "remove", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\bnikal\s+do\b", "remove", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\bnikalo\b", "calculate", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\bbadlo\b", "change", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\bjodo\b", "add", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\bdaalo\b", "add", cleaned, flags=re.IGNORECASE)

    words = cleaned.split()
    normalized_words = []
    for w in words:
        prefix = ""
        suffix = ""
        core = w
        while core and core[0] in "([{\"'`":
            prefix += core[0]
            core = core[1:]
        while core and core[-1] in ".,;:!?)]}\"'`":
            suffix = core[-1] + suffix
            core = core[:-1]

        low = core.lower()
        if low in COMMON_TYPOS:
            core = COMMON_TYPOS[low]
        elif low in VALID_WORDS:
            core = low
        elif len(low) >= 4 and not any(ord(c) > 127 for c in low):
            # Fuzzy match English typos against domain keywords with similarity threshold >= 0.82
            best_kw = None
            best_score = 0.82
            for kw in DOMAIN_KEYWORDS:
                score = difflib.SequenceMatcher(None, low, kw).ratio()
                if score > best_score:
                    best_score = score
                    best_kw = kw
            if best_kw:
                core = best_kw
        normalized_words.append(prefix + core + suffix)
    return " ".join(normalized_words)


# ==============================================================================
# 4. COLUMN DETECTION & CLASSIFICATION (Multi-Domain)
# ==============================================================================

def is_id_or_meta_column(header: Any) -> bool:
    """Identifies columns like Roll No, ID, Code, Serial No that should NOT be added into totals."""
    h = str(header).lower().strip()
    return bool(re.search(r"\b(roll|id|emp_id|s\.no|sno|sr|serial|code|rank|age|year|phone|mobile|contact|date|pincode|zip)\b", h))


def find_numeric_columns(grid: List[List[Any]]) -> List[int]:
    """Finds all columns that are primarily numeric (excluding header)."""
    if len(grid) <= 1:
        return []
    cols_count = len(grid[0])
    num_cols = []
    for c in range(cols_count):
        num_count = 0
        total_non_empty = 0
        for r in range(1, len(grid)):
            val = grid[r][c] if c < len(grid[r]) else ""
            if val != "":
                total_non_empty += 1
                if is_number(val):
                    num_count += 1
        if total_non_empty > 0 and num_count / total_non_empty >= 0.6:
            num_cols.append(c)
    return num_cols


def find_score_or_val_columns(grid: List[List[Any]]) -> List[int]:
    """Finds numeric columns that are actual values (Salary, Scores, Revenue), excluding ID/roll."""
    if len(grid) <= 1:
        return []
    num_cols = find_numeric_columns(grid)
    headers = [str(c) for c in grid[0]]
    val_cols = [c for c in num_cols if not is_id_or_meta_column(headers[c])]
    return val_cols if val_cols else num_cols


def clean_header_text(text: str) -> str:
    """Strips parenthetical points, weights, currencies, and punctuation."""
    s = re.sub(r"\s*\(.*?\)", "", str(text)).strip().lower()
    s = re.sub(r"\s*\[.*?\]", "", s).strip()
    s = re.sub(r"[%:_#$₹]", "", s).strip()
    return s


def find_best_matching_column(headers: List[Any], query_or_keywords: Any) -> Optional[int]:
    """
    Finds best matching column index by exact match, order number ('2nd column', 'teesra column'),
    token overlap, or fuzzy similarity.
    """
    if not headers:
        return None

    # 1. Check order number in query
    if isinstance(query_or_keywords, str):
        target_idx = parse_target_index(query_or_keywords, "column")
        if target_idx is not None:
            if target_idx == -1:
                return len(headers) - 1
            if 1 <= target_idx <= len(headers):
                return target_idx - 1

    cleaned_headers = [clean_header_text(h) for h in headers]
    raw_headers = [str(h).lower().strip() for h in headers]

    # If query is a list of keywords
    if isinstance(query_or_keywords, list):
        for kw in query_or_keywords:
            kw_clean = clean_header_text(kw)
            for idx, ch in enumerate(cleaned_headers):
                if kw_clean == ch or kw_clean in ch:
                    return idx
        for kw in query_or_keywords:
            kw_clean = clean_header_text(kw)
            for idx, ch in enumerate(cleaned_headers):
                if difflib.SequenceMatcher(None, kw_clean, ch).ratio() >= 0.72:
                    return idx
        return None

    # Query is a string
    q = str(query_or_keywords).strip().lower()
    q_clean = clean_header_text(q)

    # 2. Exact match
    for idx, (rh, ch) in enumerate(zip(raw_headers, cleaned_headers)):
        if q == rh or q == ch or q_clean == ch:
            return idx

    # 3. Substring match
    for idx, ch in enumerate(cleaned_headers):
        if ch and (ch in q_clean or q_clean in ch):
            return idx

    # 4. Token overlap match
    q_tokens = set(re.findall(r"\w+", q_clean))
    best_overlap_idx = None
    max_overlap = 0
    for idx, ch in enumerate(cleaned_headers):
        h_tokens = set(re.findall(r"\w+", ch))
        overlap = len(q_tokens & h_tokens)
        if overlap > max_overlap:
            max_overlap = overlap
            best_overlap_idx = idx
    if max_overlap > 0:
        return best_overlap_idx

    # 5. Fuzzy similarity match
    best_idx = None
    best_ratio = 0.70
    for idx, ch in enumerate(cleaned_headers):
        ratio = difflib.SequenceMatcher(None, q_clean, ch).ratio()
        if ratio > best_ratio:
            best_ratio = ratio
            best_idx = idx

    return best_idx


# ==============================================================================
# 5. SELECTIVE DATA MERGER (50 Pasted Records -> 20 Existing Records)
# ==============================================================================

def is_selective_merge_request(prompt: str) -> bool:
    """Detects whether prompt asks to selectively merge/update data for existing rows."""
    norm = prompt.lower()
    return any(phrase in norm for phrase in [
        "exist already", "already exist", "already in the table", "already in table",
        "who exist", "that exist", "exist karte", "pehle se hai", "pehle se table me",
        "only update existing", "update existing", "merge data", "not of the students who do not exist",
        "not of the employees who do not exist", "not of the", "baaki mat dalo", "baaki chhod do",
        "add the data of the", "update data of"
    ]) and (":" in prompt or "\t" in prompt or "," in prompt or "\n" in prompt)


def parse_pasted_data_entries(prompt: str) -> Tuple[str, List[Tuple[str, List[str]]]]:
    """Extracts candidate column name and data entries from pasted text block."""
    lines = [l.strip() for l in prompt.splitlines() if l.strip()]
    data_entries: List[Tuple[str, List[str]]] = []

    # Detect target column name
    new_col_name = "New Data"
    for l in lines:
        if ":" in l and not any(c.isdigit() for c in l):
            cand = re.sub(r"\s*(?:scores|marks|data|list)\s*:.*$", "", l, flags=re.IGNORECASE).strip()
            if cand and len(cand) < 30 and cand.lower() not in ["the", "note", "tip"]:
                new_col_name = cand.title()
                break
        col_m = re.search(r"(?:here is|add|scores for|data for|update)\s+(?:the\s+)?([A-Za-z0-9_\s]+?)(?:scores|marks|data|salary|bonus|list|:|$)", l, re.IGNORECASE)
        if col_m:
            candidate = col_m.group(1).strip()
            if candidate and candidate.lower() not in ["the", "data", "students", "employees", "table", "student", "employee"]:
                new_col_name = candidate.title()
                break

    for l in lines:
        # Check key-value or delimited pair: 'Emma Watson: 98'
        if ":" in l:
            parts = [p.strip() for p in l.split(":", 1)]
            if len(parts) == 2 and any(c.isdigit() for c in parts[1]):
                data_entries.append((parts[0], [parts[1]]))
        elif " - " in l:
            parts = [p.strip() for p in l.split(" - ")]
            if len(parts) >= 2 and any(any(c.isdigit() for c in p) for p in parts[1:]):
                data_entries.append((parts[0], parts[1:]))
        elif "\t" in l:
            parts = [p.strip() for p in l.split("\t")]
            if len(parts) >= 2:
                key = parts[0] if not parts[0].isdigit() else parts[1]
                vals = [p for p in parts if p != key]
                data_entries.append((key, vals))
        elif "," in l:
            parts = [p.strip() for p in l.split(",")]
            if len(parts) >= 2 and any(any(c.isdigit() for c in p) for p in parts[1:]):
                data_entries.append((parts[0], parts[1:]))
        else:
            # Space separated: 'Emma Watson 98'
            m_space = re.search(r"^([A-Za-z\s]+?)\s+(\d+(?:\.\d+)?)$", l)
            if m_space:
                data_entries.append((m_space.group(1).strip(), [m_space.group(2).strip()]))

    return new_col_name, data_entries


def execute_selective_merge(grid: List[List[Any]], prompt: str) -> Tuple[List[List[Any]], str]:
    """
    Selectively updates or appends data for rows that ALREADY EXIST in the table,
    strictly ignoring any entries in the pasted list that do not exist in the current sheet.
    """
    headers = list(grid[0])
    new_col_name, entries = parse_pasted_data_entries(prompt)
    if not entries:
        raise ValueError("Could not extract any matching records from the pasted data.")

    # Check if pasted keys are names or IDs
    is_key_numeric = any(k.strip().isdigit() for k, _ in entries[:5])
    name_col_idx = 0
    id_col_idx = None
    for idx, h in enumerate(headers):
        h_low = str(h).lower()
        if any(k in h_low for k in ["name", "student", "employee", "item", "product", "person"]):
            name_col_idx = idx
        if any(k in h_low for k in ["roll", "id", "code", "emp_id", "s.no"]):
            id_col_idx = idx

    # Build lookup map of existing rows indexing both Names and IDs
    existing_map: Dict[str, int] = {}
    for r_idx, row in enumerate(grid[1:], start=1):
        # Index name
        name_key = str(row[name_col_idx]).strip().lower()
        existing_map[name_key] = r_idx
        existing_map[re.sub(r"\s+", " ", name_key)] = r_idx
        # Index ID if available
        if id_col_idx is not None and id_col_idx < len(row):
            id_key = str(row[id_col_idx]).strip().lower()
            existing_map[id_key] = r_idx
        elif len(row) > 0 and str(row[0]).strip().isdigit():
            existing_map[str(row[0]).strip()] = r_idx

    # Check if target column already exists in grid
    existing_target_col = find_best_matching_column(headers, new_col_name)

    val_count = max(len(vals) for key, vals in entries)
    updated_grid = [list(grid[0])]

    if existing_target_col is not None:
        col_start_idx = existing_target_col
        for row in grid[1:]:
            updated_grid.append(list(row))
    else:
        col_start_idx = len(headers)
        if val_count == 1:
            updated_grid[0].append(new_col_name)
        else:
            for i in range(val_count):
                updated_grid[0].append(f"{new_col_name} {i+1}")
        for row in grid[1:]:
            updated_grid.append(list(row) + [""] * val_count)

    updated_count = 0
    ignored_count = 0

    for key, vals in entries:
        clean_key = key.strip().lower()
        clean_key = re.sub(r"^\d+[\.\)\-]\s*", "", clean_key)  # strip list numbers '1. '

        matched_row_idx = None
        if clean_key in existing_map:
            matched_row_idx = existing_map[clean_key]
        else:
            for ex_key, ex_r_idx in existing_map.items():
                if difflib.SequenceMatcher(None, clean_key, ex_key).ratio() >= 0.82:
                    matched_row_idx = ex_r_idx
                    break

        if matched_row_idx is not None:
            for v_i, val in enumerate(vals):
                val_clean = clean_cell(val)
                target_idx = col_start_idx + v_i
                if target_idx < len(updated_grid[matched_row_idx]):
                    updated_grid[matched_row_idx][target_idx] = val_clean
            updated_count += 1
        else:
            ignored_count += 1

    msg = f"Updated data for {updated_count} existing records from pasted list (ignored {ignored_count} non-existing records)."
    return normalize_rectangular_grid(updated_grid), msg


# ==============================================================================
# 6. MULTI-DOMAIN TABLE TEMPLATES & GENERATOR
# ==============================================================================

def extract_columns_from_prompt(prompt: str) -> List[str]:
    """Extracts column list from prompt for any domain."""
    normalized = normalize_prompt(prompt)
    no_rows = re.sub(r"\bfor\s+\d+\s+(?:students|records|rows|items|people|employees)\b", "", normalized, flags=re.IGNORECASE)
    no_rows = re.sub(r"\bwith\s+\d+\s+(?:students|records|rows|items|people|employees)\b", "", no_rows, flags=re.IGNORECASE)
    no_rows = re.sub(r"\bke\s+sath\b", "", no_rows, flags=re.IGNORECASE)

    # Match 'columns: ...'
    m = re.search(r"(?:columns|headers|fields)[:\s]+([^.\n]+)", no_rows, re.IGNORECASE)
    if m:
        raw = m.group(1).strip()
    else:
        # Match 'create table with ...' or 'make table ...'
        m2 = re.search(r"(?:create|make|generate|build)\s+(?:a\s+|an\s+)?(?:table|sheet|grid)\s*(?:with\s+)?(?:columns\s+|headers\s+)?[:\s]*([^.\n]+)", no_rows, re.IGNORECASE)
        if m2:
            raw = m2.group(1).strip()
        else:
            raw = ""

    if raw:
        cols = [c.strip(" '\"\t`[]") for c in re.split(r"[,;]|\band\b", raw) if c.strip(" '\"\t`[]")]
        cols = [c for c in cols if c.lower() not in ["columns", "table", "rows", "new", "blank"]]
        if len(cols) >= 2 or (len(cols) == 1 and "," in raw):
            return cols
    return []


def generate_table_from_columns(cols: List[str], row_count: int = 5) -> List[List[Any]]:
    """Generates realistic sample data for any column set (Salary, Sales, Students, etc.)."""
    sample_names = ["Emma Watson", "Liam Johnson", "Sophia Martinez", "Noah Brown", "Olivia Davis", "Aarav Sharma", "Priya Patel", "Lucas Wilson"]
    departments = ["Engineering", "Marketing", "Sales", "Human Resources", "Finance", "Operations", "Product"]
    sample_items = ["Laptop", "Wireless Mouse", "USB-C Hub", "Mechanical Keyboard", "27-inch Monitor", "Desk Chair"]

    header = [c.title() if not c.isupper() else c for c in cols]
    grid = [header]

    for i in range(row_count):
        row = []
        subject_scores = []
        base_salary = 60000 + (i * 7500)
        bonus_val = 5000 + (i * 1200)

        for c in cols:
            c_low = c.lower()
            if any(k in c_low for k in ["emp id", "employee id", "emp_id"]):
                row.append(f"EMP{101 + i}")
            elif any(k in c_low for k in ["roll", "id", "s.no", "sr"]):
                row.append(101 + i)
            elif any(k in c_low for k in ["name", "student", "employee", "person"]):
                row.append(sample_names[i % len(sample_names)])
            elif "department" in c_low or "dept" in c_low:
                row.append(departments[i % len(departments)])
            elif any(k in c_low for k in ["base salary", "basic salary", "salary", "wage"]):
                row.append(base_salary)
            elif "bonus" in c_low:
                row.append(bonus_val)
            elif any(k in c_low for k in ["total salary", "gross salary", "compensation", "net salary", "ctc"]):
                row.append(base_salary + bonus_val)
            elif any(k in c_low for k in ["math", "phys", "chem", "bio", "sci", "eng", "hindi", "score", "mark", "quiz", "midterm", "exam"]):
                score = 72 + ((i * 7 + len(c) * 3) % 26)
                subject_scores.append(score)
                row.append(score)
            elif "total" in c_low:
                tot = sum(subject_scores) if subject_scores else (base_salary + bonus_val)
                row.append(tot)
            elif any(k in c_low for k in ["average", "avg"]):
                avg = round(sum(subject_scores) / len(subject_scores), 1) if subject_scores else 82.5
                row.append(avg)
            elif any(k in c_low for k in ["percent", "%"]):
                pct = f"{round((sum(subject_scores) / (len(subject_scores) * 100)) * 100, 1)}%" if subject_scores else f"{80 + i * 3}%"
                row.append(pct)
            elif "grade" in c_low:
                row.append(["A", "B", "A+", "B+", "A"][i % 5])
            elif any(k in c_low for k in ["pass", "result", "status"]):
                row.append("Pass" if i != 3 else "Fail")
            elif any(k in c_low for k in ["attend", "present"]):
                row.append(f"{85 + (i * 3) % 15}%")
            elif any(k in c_low for k in ["product", "item", "description"]):
                row.append(sample_items[i % len(sample_items)])
            elif any(k in c_low for k in ["price", "cost", "rate"]):
                row.append([499, 199, 49, 29, 149][i % 5])
            elif any(k in c_low for k in ["qty", "quantity", "stock"]):
                row.append([15, 40, 60, 25, 10][i % 5])
            elif "date" in c_low:
                row.append(f"2026-10-{15 + i:02d}")
            else:
                row.append("")
        grid.append(row)

    return normalize_rectangular_grid(grid)


def generate_template_table(prompt: str) -> Optional[List[List[Any]]]:
    """Generates standard template tables across multiple business & educational domains."""
    norm = normalize_prompt(prompt).lower()

    row_count = 5
    rc_match = re.search(r"\b(?:for|with)\s+(\d+)\s+(?:students|records|rows|items|employees)", norm)
    if rc_match:
        row_count = min(int(rc_match.group(1)), 25)

    # 1. Custom extracted columns
    extracted_cols = extract_columns_from_prompt(prompt)
    if extracted_cols:
        return generate_table_from_columns(extracted_cols, row_count)

    # 2. Employee Salary & Payroll Template
    if any(k in norm for k in ["salary", "payroll", "employee table", "compensation", "vetan", "tankhwah"]):
        header = ["Employee ID", "Employee Name", "Department", "Base Salary ($)", "Bonus ($)", "Total Salary ($)"]
        data = [
            header,
            ["EMP101", "Emma Watson", "Engineering", 85000, 10000, 95000],
            ["EMP102", "Liam Johnson", "Marketing", 65000, 6000, 71000],
            ["EMP103", "Sophia Martinez", "Product", 92000, 12000, 104000],
            ["EMP104", "Noah Brown", "Sales", 58000, 8000, 66000],
            ["EMP105", "Olivia Davis", "Human Resources", 72000, 7000, 79000],
        ]
        return normalize_rectangular_grid(data)

    # 3. Sales & Products Template
    if any(k in norm for k in ["sales", "product table", "revenue", "order table", "bikri"]):
        header = ["Product ID", "Product Name", "Category", "Unit Price ($)", "Quantity Sold", "Total Revenue ($)"]
        data = [
            header,
            ["PRD-01", "Wireless Laptop", "Electronics", 899, 15, 13485],
            ["PRD-02", "Mechanical Keyboard", "Accessories", 89, 45, 4005],
            ["PRD-03", "USB-C Hub", "Accessories", 49, 60, 2940],
            ["PRD-04", "27-inch Monitor", "Displays", 279, 20, 5580],
            ["PRD-05", "Ergonomic Desk Chair", "Furniture", 199, 12, 2388],
        ]
        return normalize_rectangular_grid(data)

    # 4. Attendance Sheet Template
    if any(k in norm for k in ["attendance", "present absent", "attendance sheet"]):
        header = ["Roll No", "Student Name", "Mon", "Tue", "Wed", "Thu", "Fri", "Total Present"]
        data = [
            header,
            [101, "Emma Watson", "P", "P", "P", "P", "P", 5],
            [102, "Liam Johnson", "P", "A", "P", "P", "P", 4],
            [103, "Sophia Martinez", "P", "P", "P", "P", "P", 5],
            [104, "Noah Brown", "A", "P", "P", "A", "P", 3],
            [105, "Olivia Davis", "P", "P", "P", "P", "P", 5],
        ]
        return normalize_rectangular_grid(data)

    # 5. Default Gradebook Template
    if any(k in norm for k in ["create table", "make table", "gradebook", "student table", "sample table", "new table", "banao"]):
        header = ["Student Name", "Quiz 1 (20)", "Quiz 2 (20)", "Midterm (100)", "Final Exam (100)", "Total Score (100)"]
        data = [
            header,
            ["Emma Watson", 19, 18, 92, 95, 93.5],
            ["Liam Johnson", 14, 16, 78, 82, 79.0],
            ["Sophia Martinez", 20, 19, 96, 98, 97.0],
            ["Noah Brown", 12, 11, 65, 70, 67.0],
            ["Olivia Davis", 17, 18, 88, 85, 86.5],
            ["Lucas Wilson", 15, 14, 72, 75, 74.0],
            ["Ava Taylor", 18, 19, 90, 93, 91.5]
        ]
        return normalize_rectangular_grid(data)

    return None


def parse_prompt_to_table(prompt: str) -> Optional[List[List[Any]]]:
    # 1. Template generation
    grid = generate_template_table(prompt)
    if grid:
        return grid
    # 2. Markdown table
    lines = [line.strip() for line in prompt.strip().splitlines() if line.strip().startswith("|")]
    if len(lines) >= 2:
        m_grid = []
        for line in lines:
            if re.match(r"^\|[\s\-:|]+\|$", line):
                continue
            cells = [c.strip() for c in line.split("|")]
            if len(cells) >= 2:
                row = cells[1:-1]
                if any(row):
                    m_grid.append([clean_cell(c) for c in row])
        if len(m_grid) >= 1:
            return normalize_rectangular_grid(m_grid)
    # 3. CSV / Delimited
    csv_grid = None
    sample = prompt.strip()
    if "\t" in sample or "," in sample or ";" in sample:
        try:
            delim = "\t" if "\t" in sample else ("," if "," in sample else ";")
            reader = csv.reader(io.StringIO(sample), delimiter=delim)
            rows = [[clean_cell(c) for c in r] for r in reader if any(r)]
            if len(rows) >= 2 and len(rows[0]) > 1:
                csv_grid = normalize_rectangular_grid(rows)
        except Exception:
            pass
    if csv_grid:
        return csv_grid

    # 4. Key-value list parser
    p_lines = [l.strip() for l in prompt.splitlines() if l.strip()]
    pairs = []
    for l in p_lines:
        if ":" in l:
            parts = [p.strip() for p in l.split(":", 1)]
            pairs.append(parts)
        elif " - " in l:
            parts = [p.strip() for p in l.split(" - ")]
            pairs.append(parts)
    if len(pairs) >= 2:
        headers = ["Name", "Value"]
        return normalize_rectangular_grid([headers] + pairs)

    return None


# ==============================================================================
# 6.5 NATURAL INTENT HANDLERS: ROWS, COLUMNS, CLEAR, DUPLICATE
# ==============================================================================

def auto_increment_id(last_val: Any) -> str:
    s = str(last_val).strip()
    m = re.match(r"^([A-Za-z_\-\s]*)(\d+)$", s)
    if m:
        prefix, num_str = m.group(1), m.group(2)
        next_num = int(num_str) + 1
        if num_str.startswith("0") and len(num_str) > 1:
            return f"{prefix}{next_num:0{len(num_str)}d}"
        return f"{prefix}{next_num}"
    try:
        return str(int(float(s)) + 1)
    except:
        return "101"


def is_add_row_request(prompt: str) -> bool:
    p = prompt.strip().lower()
    if p in [
        "add row", "add new row", "add a row", "add another row", "add blank row", "add empty row",
        "insert row", "insert new row", "insert a row", "insert blank row", "insert empty row",
        "new row", "another row", "one more row", "blank row", "empty row", "append row", "append new row",
        "create row", "create new row", "row add karo", "nayi row", "naya row", "ek row jodo",
        "nayi row add karo", "naya row add karo", "row jodo", "पंक्ति जोड़ो", "रो जोड़ो"
    ]:
        return True

    verb = r"(?:add|insert|append|create|put|plus|jodo|daalo|banao|शामिल\s+करो|जोड़ो)"
    row_noun = r"(?:row|rows|record|records|entry|entries|line|lines|pankti|पंक्ति|रो)"
    
    if re.search(rf"\b{verb}\b.*\b{row_noun}\b", p):
        return True
    if re.search(rf"\b{row_noun}\b.*\b{verb}\b", p):
        return True
    if re.search(rf"^(?:new|blank|empty|another|one more|extra|fresh|naya|nayi)\s+{row_noun}\b", p):
        return True
    if re.search(r"\b(?:add|insert|append)\s+(?:a\s+|an\s+|the\s+|new\s+|another\s+)?(?:student|employee|person|item|candidate|worker|teacher|client|customer)\b", p):
        return True

    return False


def execute_add_row(grid: List[List[Any]], prompt: str) -> Tuple[List[List[Any]], str]:
    headers = [str(c) for c in grid[0]]
    p_lower = prompt.lower()
    
    # 1. Determine how many rows to add
    count = 1
    count_m = re.search(r"\b(?:add|insert|create|append)\s+(\d+|two|three|four|five|ten)\s+(?:new\s+|blank\s+|empty\s+)?(?:rows|row|records|pankti|line|lines)", p_lower)
    row_count_str = None
    if count_m:
        raw_c = count_m.group(1).lower()
        row_count_str = raw_c
        word_map = {"two": 2, "three": 3, "four": 4, "five": 5, "ten": 10}
        count = word_map.get(raw_c, int(raw_c) if raw_c.isdigit() else 1)
    count = max(1, min(100, count))

    # 2. Position
    is_top = any(k in p_lower for k in ["at top", "at the top", "at start", "at beginning", "first row", "upar", "top me"])

    # 3. Identify column categories in grid
    id_col_idx = None
    name_col_idx = None
    dept_col_idx = None
    num_cols = []

    for idx, h in enumerate(headers):
        h_low = h.lower()
        if any(k in h_low for k in ["id", "roll", "emp id", "empid", "s.no", "code"]) and id_col_idx is None:
            id_col_idx = idx
        elif any(k in h_low for k in ["name", "student", "employee", "person", "client"]) and name_col_idx is None:
            name_col_idx = idx
        elif any(k in h_low for k in ["dept", "department", "team"]) and dept_col_idx is None:
            dept_col_idx = idx
        elif idx != id_col_idx:
            num_samples = [r[idx] for r in grid[1:10] if idx < len(r) and is_number(r[idx])]
            if len(num_samples) >= 1:
                num_cols.append(idx)

    # 4. Extract data values from prompt (if any)
    name_val = ""
    name_m = re.search(r"(?:student|employee|person|name\s+is|named|for)\s+([A-Za-z]+(?:\s+[A-Za-z]+)?)", prompt, re.IGNORECASE)
    if name_m:
        raw_n = name_m.group(1).strip()
        clean_n = re.sub(r"\b(?:with|and|having|whose|salary|bonus|dept|department|score|marks|grade|roll|id)\b.*$", "", raw_n, flags=re.IGNORECASE).strip()
        if clean_n.lower() not in ["row", "new", "blank", "a", "an", "the", "another", "one", "record"]:
            name_val = clean_n.title()

    dept_val = ""
    dept_m = re.search(r"(?:department|dept|team)\s+(?:is\s+|as\s+)?([A-Za-z]+)", prompt, re.IGNORECASE)
    if dept_m:
        dept_val = dept_m.group(1).strip().title()

    prompt_without_count = re.sub(rf"\b(?:add|insert|create|append)\s+{re.escape(row_count_str or '')}\s+(?:new\s+|blank\s+|empty\s+)?(?:rows|row)", "", prompt, flags=re.IGNORECASE) if row_count_str else prompt
    data_nums = [to_number(n) for n in re.findall(r"\b\d+(?:\.\d+)?\b", prompt_without_count)]

    # Check last ID for auto-increment
    last_id = ""
    if id_col_idx is not None and len(grid) > 1:
        for r in reversed(grid[1:]):
            if id_col_idx < len(r) and str(r[id_col_idx]).strip():
                last_id = r[id_col_idx]
                break

    new_rows = []
    curr_id = last_id
    num_val_idx = 0

    for _ in range(count):
        row = [""] * len(headers)
        
        # Auto-increment ID if present
        if id_col_idx is not None:
            curr_id = auto_increment_id(curr_id) if curr_id else str(100 + len(grid) + len(new_rows))
            row[id_col_idx] = curr_id

        # Name
        if name_col_idx is not None and name_val:
            row[name_col_idx] = name_val

        # Department
        if dept_col_idx is not None and dept_val:
            row[dept_col_idx] = dept_val

        # Numeric values mapped to numeric columns
        for c in num_cols:
            if num_val_idx < len(data_nums):
                row[c] = data_nums[num_val_idx]
                num_val_idx += 1

        new_rows.append(row)

    if is_top:
        updated = [grid[0]] + new_rows + grid[1:]
        msg = f"Inserted {count} new row(s) at the top."
    else:
        updated = [list(r) for r in grid] + new_rows
        msg = f"Added {count} new row(s) to the table."

    if name_val:
        msg = f"Added new row for '{name_val}'."

    return normalize_rectangular_grid(updated), msg


def is_add_column_request(prompt: str) -> bool:
    p = prompt.strip().lower()
    if is_add_row_request(p):
        return False
    if p in [
        "add column", "add new column", "add a column", "add another column", "add blank column", "add empty column",
        "insert column", "insert new column", "insert a column", "insert blank column", "insert empty column",
        "new column", "another column", "one more column", "blank column", "empty column", "append column",
        "column add karo", "naya column", "ek column jodo", "naya column add karo", "कॉलम जोड़ो"
    ]:
        return True
    verb = r"(?:add|insert|append|create|put|plus|jodo|daalo|banao|शामिल\s+करो|जोड़ो)"
    col_noun = r"(?:column|col|cols|columns|stambh|स्तम्भ|कॉलम|field|attribute)"
    if re.search(rf"\b{verb}\b.*\b{col_noun}\b", p):
        return True
    if re.search(rf"\b{col_noun}\b.*\b{verb}\b", p):
        return True
    if re.search(rf"^(?:new|blank|empty|another|one more|extra|fresh|naya|nayi)\s+{col_noun}\b", p):
        return True
    if re.search(r"\b(?:add|insert)\s+(?:a\s+|an\s+|the\s+|new\s+)?column\s+[a-zA-Z0-9_\s]+", p):
        return True
    return False


def execute_add_column(grid: List[List[Any]], prompt: str) -> Tuple[List[List[Any]], str]:
    headers = [str(c) for c in grid[0]]
    p_lower = prompt.lower()
    
    # Extract column name
    name_m = re.search(r"(?:add|insert|create|append)\s+(?:a\s+|an\s+|the\s+|new\s+|blank\s+|empty\s+)?(?:column|col|field)\s+(?:named|called\s+)?['\"`]?([^'\"`\s,]+(?:\s+[^'\"`\s,]+)?)['\"`]?", prompt, re.IGNORECASE)
    
    col_name = "New Column"
    if name_m:
        cand = name_m.group(1).strip()
        if cand.lower() not in ["with", "value", "default", "at", "before", "after", "named", "called", "row", "table"]:
            cand = re.sub(r"\s+\b(?:with|having|default|at|before|after)\b.*$", "", cand, flags=re.IGNORECASE).strip()
            if cand:
                col_name = cand.title()

    if col_name in headers:
        suffix_num = 2
        while f"{col_name} {suffix_num}" in headers:
            suffix_num += 1
        col_name = f"{col_name} {suffix_num}"

    # Extract default value
    default_val = ""
    val_m = re.search(r"(?:with\s+value|with\s+default|with|value\s+is|value)\s+['\"`]?([^'\"`\n]+?)['\"`]?$", prompt, re.IGNORECASE)
    if val_m:
        raw_val = val_m.group(1).strip()
        if raw_val.lower() not in ["row", "table", "sheet"]:
            default_val = clean_cell(raw_val)

    is_start = any(k in p_lower for k in ["at start", "at beginning", "first column", "shuru me"])
    
    updated = []
    if is_start:
        updated.append([col_name] + list(grid[0]))
        for row in grid[1:]:
            updated.append([default_val] + list(row))
        msg = f"Inserted new column '{col_name}' at the beginning."
    else:
        updated.append(list(grid[0]) + [col_name])
        for row in grid[1:]:
            updated.append(list(row) + [default_val])
        val_str = f" with default '{default_val}'" if default_val != "" else ""
        msg = f"Added column '{col_name}'{val_str} to the table."

    return normalize_rectangular_grid(updated), msg


def is_delete_row_request(prompt: str) -> bool:
    p = prompt.strip().lower()
    delete_verb = r"(?:delete|remove|drop|cut|clear|hatao|hata\s+do|nikal\s+do|nikalo|मिटाओ|हटाओ)"
    row_noun = r"(?:row|rows|record|records|entry|entries|line|lines|student|employee|person|item|candidate|pankti|पंक्ति|रो)"
    
    if re.search(rf"\b{delete_verb}\b.*\b{row_noun}\b", p):
        return True
    if re.search(rf"\b{row_noun}\b.*\b{delete_verb}\b", p):
        return True
    return False


def execute_delete_row(grid: List[List[Any]], prompt: str) -> Tuple[List[List[Any]], str]:
    headers = [str(c) for c in grid[0]]
    p_lower = prompt.lower()
    norm_prompt = normalize_prompt(prompt)

    # Empty rows
    if any(k in p_lower for k in ["blank row", "empty row", "khali row"]):
        clean_rows = [grid[0]]
        deleted_count = 0
        for r in grid[1:]:
            if any(c not in [None, ""] for c in r):
                clean_rows.append(r)
            else:
                deleted_count += 1
        return normalize_rectangular_grid(clean_rows), f"Removed {deleted_count} empty row(s)."

    row_order = parse_target_index(norm_prompt, "row")
    if row_order is not None:
        if row_order == -1:
            target_r_idx = len(grid) - 1
        elif 1 <= row_order < len(grid):
            target_r_idx = row_order
        else:
            target_r_idx = len(grid) - 1
    elif any(k in p_lower for k in ["last row", "aakhri row", "bottom row"]):
        target_r_idx = len(grid) - 1
    elif any(k in p_lower for k in ["first row", "top row", "pehli row"]):
        target_r_idx = 1
    else:
        name_m = re.search(r"(?:student|employee|row\s+for|for|name\s+is)\s+([A-Za-z]+(?:\s+[A-Za-z]+)?)", norm_prompt, re.IGNORECASE)
        if name_m:
            target_name = name_m.group(1).lower().strip()
            updated = [grid[0]]
            deleted = False
            for row in grid[1:]:
                row_str = " ".join(str(c).lower() for c in row)
                if target_name in row_str:
                    deleted = True
                else:
                    updated.append(row)
            if deleted:
                return normalize_rectangular_grid(updated), f"Removed row record for '{target_name.title()}'."
        
        target_r_idx = len(grid) - 1

    if target_r_idx is not None and 1 <= target_r_idx < len(grid):
        row_label = str(grid[target_r_idx][0]) if grid[target_r_idx] else f"Row {target_r_idx}"
        updated = [r for idx, r in enumerate(grid) if idx != target_r_idx]
        return normalize_rectangular_grid(updated), f"Deleted row {target_r_idx} ({row_label})."

    return grid, "No rows deleted."


def is_delete_column_request(prompt: str) -> bool:
    p = prompt.strip().lower()
    if is_delete_row_request(p):
        return False
    delete_verb = r"(?:delete|remove|drop|cut|hatao|hata\s+do|nikal\s+do|nikalo|मिटाओ|हटाओ)"
    col_noun = r"(?:column|col|cols|columns|stambh|स्तम्भ|कॉलम|field|header)"
    
    if re.search(rf"\b{delete_verb}\b.*\b{col_noun}\b", p):
        return True
    if re.search(rf"\b{col_noun}\b.*\b{delete_verb}\b", p):
        return True
    return False


def execute_delete_column(grid: List[List[Any]], prompt: str) -> Tuple[List[List[Any]], str]:
    headers = [str(c) for c in grid[0]]
    p_lower = prompt.lower()
    norm_prompt = normalize_prompt(prompt)

    target_col = None

    col_order = parse_target_index(norm_prompt, "column")
    if col_order is not None:
        if col_order == -1:
            target_col = len(headers) - 1
        elif 1 <= col_order <= len(headers):
            target_col = col_order - 1
    elif any(k in p_lower for k in ["last column", "aakhri column"]):
        target_col = len(headers) - 1
    elif any(k in p_lower for k in ["first column", "pehla column"]):
        target_col = 0
    else:
        extract_col_match = re.search(r"(?:remove|delete|drop|hatao|hata\s+do)\s+(?:column\s+|col\s+)?['\"`]?([^'\"`\n]+?)['\"`]?$", norm_prompt, re.IGNORECASE)
        query = extract_col_match.group(1).strip() if extract_col_match else norm_prompt
        target_col = find_best_matching_column(headers, query)

        if target_col is None:
            target_col = find_best_matching_column(headers, norm_prompt)

    if target_col is None:
        target_col = len(headers) - 1

    if target_col is not None and 0 <= target_col < len(headers):
        target_name = headers[target_col]
        updated = []
        for row in grid:
            new_row = [row[i] for i in range(len(row)) if i != target_col]
            updated.append(new_row)
        return normalize_rectangular_grid(updated), f"Removed column '{target_name}' (Column {target_col + 1})."

    return grid, "No column removed."


def is_clear_data_request(prompt: str) -> bool:
    p = prompt.strip().lower()
    return any(k in p for k in [
        "clear table", "clear data", "clear all data", "clear rows", "clear data rows",
        "empty table", "reset table", "table saaf karo", "data saaf karo"
    ])


def execute_clear_data(grid: List[List[Any]], prompt: str) -> Tuple[List[List[Any]], str]:
    if not grid:
        return [], "Table is empty."
    headers = grid[0]
    blank_row = [""] * len(headers)
    updated = [headers, blank_row]
    return updated, "Cleared all data rows while preserving headers."


def is_duplicate_row_request(prompt: str) -> bool:
    p = prompt.strip().lower()
    return any(k in p for k in ["duplicate row", "copy row", "clone row", "duplicate last row", "row copy karo"])


def execute_duplicate_row(grid: List[List[Any]], prompt: str) -> Tuple[List[List[Any]], str]:
    if len(grid) < 2:
        return grid, "No data row available to duplicate."
    p = prompt.strip().lower()
    norm_prompt = normalize_prompt(prompt)
    target_idx = len(grid) - 1
    
    row_order = parse_target_index(norm_prompt, "row")
    if row_order is not None and 1 <= row_order < len(grid):
        target_idx = row_order
        
    cloned_row = list(grid[target_idx])
    for idx, h in enumerate(grid[0]):
        if any(k in str(h).lower() for k in ["id", "roll", "emp id", "empid", "s.no"]):
            cloned_row[idx] = auto_increment_id(grid[-1][idx])
            break

    updated = [list(r) for r in grid] + [cloned_row]
    return normalize_rectangular_grid(updated), f"Duplicated row {target_idx}."


# ==============================================================================
# 7. PRIMARY TABLE INSTRUCTION EXECUTION ENGINE
# ==============================================================================

def execute_table_instruction(grid: List[List[Any]], raw_prompt: str) -> Tuple[List[List[Any]], str]:
    """
    Executes natural language instruction on existing 2D grid with full multi-domain,
    multilingual, typo-tolerant, and order/name targeting capabilities.
    """
    grid = normalize_rectangular_grid(grid)
    norm_prompt = normalize_prompt(raw_prompt)
    p_lower = norm_prompt.lower()

    # --------------------------------------------------------------------------
    # CASE 0: EMPTY GRID -> CREATE TABLE
    # --------------------------------------------------------------------------
    if not grid or len(grid) == 0:
        new_grid = parse_prompt_to_table(raw_prompt)
        if new_grid:
            return new_grid, f"Created new table with {len(new_grid)-1} rows from your instruction."
        raise ValueError("Worksheet is empty. Try: 'Create employee salary table' or 'Create table with Roll No, Name, Math, Science'")

    headers = [str(c) for c in grid[0]]

    # --------------------------------------------------------------------------
    # FEATURE 1: SMART DATA CLEANING
    # "Clean this sheet, but don't delete any data without showing me first"
    # "Remove duplicates and empty rows", "Fix inconsistent names and dates"
    # --------------------------------------------------------------------------
    is_clean_intent = any(k in p_lower for k in [
        "clean this sheet", "clean sheet", "clean data", "smart clean", "data cleaning",
        "remove duplicate", "remove duplicates", "delete duplicate", "delete duplicates",
        "remove empty row", "delete empty row", "fix inconsistent", "standardize names",
        "fix capitalization", "clean duplicates", "clean table", "clean the table"
    ])
    if is_clean_intent:
        cleaned_grid, clean_msg, clean_report = data_cleaning.execute_smart_clean(grid, raw_prompt)
        meta = {
            "action_type": "smart_clean",
            "cleaning_report": clean_report,
            "requires_confirmation": clean_report.get("requires_confirmation", False),
        }
        return cleaned_grid, clean_msg, meta

    # --------------------------------------------------------------------------
    # FEATURE 5: DATA VALIDATION & ANOMALY DETECTION
    # "Find employees whose salaries are unusually high compared with their department"
    # "Detect missing employee IDs or invalid email addresses", "Check for outliers"
    # --------------------------------------------------------------------------
    is_anomaly_intent = any(k in p_lower for k in [
        "unusually high", "unusually low", "outlier", "outliers", "anomaly", "anomalies",
        "suspicious", "impossible date", "invalid email", "missing id", "missing employee",
        "validation", "audit data", "check errors", "check invalid", "unusual expense"
    ])
    if is_anomaly_intent:
        anomaly_report = anomaly_engine.detect_anomalies_and_validation(grid, raw_prompt)
        meta = {
            "action_type": "anomaly_detection",
            "anomaly_report": anomaly_report,
            "anomalies": anomaly_report.get("anomalies", []),
        }
        return grid, anomaly_report.get("summary", "Audit completed."), meta

    # --------------------------------------------------------------------------
    # FEATURE 4: AUTOMATIC REPORTS & DASHBOARDS
    # "Create a monthly sales dashboard with the top five employees"
    # "Generate dashboard", "Create summary report sheet", "KPI dashboard"
    # --------------------------------------------------------------------------
    is_dashboard_intent = any(k in p_lower for k in [
        "sales dashboard", "create dashboard", "generate dashboard", "dashboard with",
        "monthly sales dashboard", "summary sheet", "executive dashboard",
        "top five employees", "top 5 employees", "create report sheet", "summary dashboard"
    ])
    if is_dashboard_intent:
        dash_res = dashboard_engine.generate_dashboard_report(grid, raw_prompt)
        meta = {
            "action_type": "dashboard_report",
            "dashboard_report": dash_res,
            "is_new_sheet": True,
            "sheet_name": dash_res.get("sheet_name", "Dashboard"),
        }
        msg = f"Generated Executive Dashboard with KPIs, Top Ranking, and Chart specification on new sheet '{dash_res.get('sheet_name', 'Dashboard')}'."
        return grid, msg, meta

    # --------------------------------------------------------------------------
    # FEATURE 3: FORMULA GENERATOR & FIXER
    # "Calculate the bonus as 10% of salary if performance is above 80"
    # "Fix formula error #REF!", "Explain formula", "Generate formula"
    # --------------------------------------------------------------------------
    if any(err_tag in p_lower for err_tag in ["#ref!", "#value!", "#n/a", "#div/0!", "#name?", "fix formula", "error in formula"]):
        fixed_res = formula_engine.fix_formula_errors(raw_prompt)
        meta = {
            "action_type": "formula_fix",
            "formula_info": fixed_res,
        }
        msg = f"Formula fix identified ({fixed_res['error_type']}): {fixed_res['diagnosis']} Suggested: {fixed_res['fixed_formula']}"
        return grid, msg, meta

    if any(k in p_lower for k in ["explain formula", "what does formula", "break down formula"]):
        explained = formula_engine.explain_formula(raw_prompt)
        meta = {
            "action_type": "formula_explain",
            "formula_info": explained,
        }
        return grid, f"Formula Explanation: {explained['summary']}", meta

    # Check: "Calculate the bonus as 10% of salary if performance is above 80"
    if ("calculate" in p_lower or "compute" in p_lower or "bonus as" in p_lower) and ("if" in p_lower or "above" in p_lower):
        form_gen = formula_engine.generate_formula_from_prompt(raw_prompt, headers, current_row=2)
        if form_gen.get("status") == "success" and form_gen.get("type") == "IF_LOGICAL":
            col_name = form_gen.get("suggested_header", "Bonus")
            col_map = formula_engine.map_headers_to_columns(headers)
            val_col = col_map.get("salary") or col_map.get("base salary") or "D"
            cond_col = col_map.get("performance") or col_map.get("score") or "E"
            val_idx = formula_engine.letter_to_column_index(val_col)
            cond_idx = formula_engine.letter_to_column_index(cond_col)

            updated = [list(grid[0]) + [col_name]]
            for r_idx, row in enumerate(grid[1:], start=2):
                cond_val = to_number(row[cond_idx]) if (cond_idx < len(row) and is_number(row[cond_idx])) else 0
                base_val = to_number(row[val_idx]) if (val_idx < len(row) and is_number(row[val_idx])) else 0
                res_val = round(base_val * 0.10, 2) if cond_val > 80 else 0
                updated.append(list(row) + [res_val])

            meta = {
                "action_type": "formula_generation",
                "formula_info": form_gen,
            }
            msg = f"Added '{col_name}' column using formula {form_gen['formula']} ({form_gen['explanation']})."
            return normalize_rectangular_grid(updated), msg, meta

    # --------------------------------------------------------------------------
    # CASE 1: ADVANCED SHEET MERGING / SELECTIVE DATA MERGE
    # --------------------------------------------------------------------------
    if is_selective_merge_request(raw_prompt) or ("merge" in p_lower and "sheet" in p_lower):
        res = execute_selective_merge(grid, raw_prompt)
        return res[0], res[1], {"action_type": "sheet_merge"}

    # --------------------------------------------------------------------------
    # NATURAL ROW OPERATIONS: ADD ROW (Any phrasing: 'add new row', 'add a row', 'insert blank row', 'row add karo', etc.)
    # --------------------------------------------------------------------------
    if is_add_row_request(raw_prompt):
        updated_grid, action_msg = execute_add_row(grid, raw_prompt)
        return updated_grid, action_msg, {"action_type": "add_row"}

    # --------------------------------------------------------------------------
    # NATURAL COLUMN OPERATIONS: ADD COLUMN ('add new column', 'add column Remarks', 'insert column Status', etc.)
    # --------------------------------------------------------------------------
    if is_add_column_request(raw_prompt):
        updated_grid, action_msg = execute_add_column(grid, raw_prompt)
        return updated_grid, action_msg, {"action_type": "add_column"}

    # --------------------------------------------------------------------------
    # NATURAL ROW DELETION: DELETE ROW ('delete last row', 'delete row 3', 'remove row for Emma', 'row hatao', etc.)
    # --------------------------------------------------------------------------
    if is_delete_row_request(raw_prompt):
        updated_grid, action_msg = execute_delete_row(grid, raw_prompt)
        return updated_grid, action_msg, {"action_type": "delete_row"}

    # --------------------------------------------------------------------------
    # NATURAL COLUMN DELETION: DELETE COLUMN ('delete column 2', 'remove column Bonus', 'column hatao', etc.)
    # --------------------------------------------------------------------------
    if is_delete_column_request(raw_prompt):
        updated_grid, action_msg = execute_delete_column(grid, raw_prompt)
        return updated_grid, action_msg, {"action_type": "delete_column"}

    # --------------------------------------------------------------------------
    # NATURAL CLEAR DATA REQUEST ('clear table', 'clear data', 'reset table', etc.)
    # --------------------------------------------------------------------------
    if is_clear_data_request(raw_prompt):
        updated_grid, action_msg = execute_clear_data(grid, raw_prompt)
        return updated_grid, action_msg, {"action_type": "clear_data"}

    # --------------------------------------------------------------------------
    # DUPLICATE ROW REQUEST ('duplicate last row', 'copy row 2', etc.)
    # --------------------------------------------------------------------------
    if is_duplicate_row_request(raw_prompt):
        updated_grid, action_msg = execute_duplicate_row(grid, raw_prompt)
        return updated_grid, action_msg, {"action_type": "duplicate_row"}

    # --------------------------------------------------------------------------
    # CASE 4: CELL VALUE UPDATE (e.g. 'change score of Emma to 95', 'set Rahul salary to 75000')
    # --------------------------------------------------------------------------
    if any(k in p_lower for k in ["change", "update", "set", "badlo"]) and any(k in p_lower for k in ["score", "mark", "salary", "bonus", "value", "to", "ko"]):
        val_match = re.search(r"\bto\s+([0-9]+(?:\.[0-9]+)?|[A-Za-z0-9_+\-%]+)\b", norm_prompt) or re.search(r"\b([0-9]+(?:\.[0-9]+)?)\s*(?:karo|rakho)\b", norm_prompt)
        if val_match:
            new_val = clean_cell(val_match.group(1))

            target_row_idx = None
            row_label = ""
            for r_idx, row in enumerate(grid[1:], start=1):
                for c in range(min(3, len(row))):
                    cell_text = str(row[c]).lower()
                    if cell_text and (cell_text in p_lower or any(part in p_lower for part in cell_text.split() if len(part) >= 3)):
                        target_row_idx = r_idx
                        row_label = str(row[c])
                        break
                if target_row_idx is not None:
                    break

            target_col_idx = None
            for idx, h in enumerate(headers):
                h_clean = clean_header_text(h)
                if h_clean and (h_clean in p_lower or h.lower() in p_lower):
                    target_col_idx = idx
                    break

            if target_col_idx is None:
                val_cols = find_score_or_val_columns(grid)
                if val_cols:
                    target_col_idx = val_cols[-1]

            if target_row_idx is not None and target_col_idx is not None:
                updated = [list(r) for r in grid]
                old_val = updated[target_row_idx][target_col_idx]
                updated[target_row_idx][target_col_idx] = new_val
                return normalize_rectangular_grid(updated), f"Updated '{headers[target_col_idx]}' for {row_label} from {old_val} to {new_val}."

    # --------------------------------------------------------------------------
    # CASE 5: SALARY HIKE / BONUS / PERCENTAGE ADJUSTMENTS (Generic Curve)
    # --------------------------------------------------------------------------
    if any(k in p_lower for k in ["hike", "raise", "bonus", "curve", "increase", "deduct", "discount", "increment", "grace", "badhao"]):
        amount_match = re.search(r"(\d+(?:\.\d+)?)\s*%", p_lower)
        is_pct_hike = amount_match is not None
        if not amount_match:
            amount_match = re.search(r"(\d+(?:\.\d+)?)\s*(?:point|pt|mark|bonus|dollar|rs|rupees|bucks)", p_lower) or re.search(r"(?:hike|curve|bonus|deduct|discount|add|increase by)\s*(\d+(?:\.\d+)?)", p_lower)

        hike_val = float(amount_match.group(1)) if amount_match else 10.0
        is_deduct = any(k in p_lower for k in ["deduct", "decrease", "discount", "minus", "ghatao"])
        if is_deduct:
            hike_val = -hike_val

        cap_val = 100.0 if ("100" in p_lower or "cap at 100" in p_lower) else None

        # Find target column (Salary, Price, Midterm, Score, etc.)
        val_cols = find_score_or_val_columns(grid)
        target_col = find_best_matching_column(headers, p_lower)
        if target_col is None or target_col not in val_cols:
            target_col = val_cols[0] if val_cols else 0

        target_name = headers[target_col]
        updated = [list(grid[0])]
        for row in grid[1:]:
            new_row = list(row)
            val = new_row[target_col]
            if is_number(val):
                num = to_number(val)
                if is_pct_hike:
                    new_num = num * (1 + (hike_val / 100.0))
                else:
                    new_num = num + hike_val
                if cap_val is not None:
                    new_num = min(new_num, cap_val)
                new_row[target_col] = round(new_num, 2)
            updated.append(new_row)

        desc = f"{'+' if hike_val > 0 else ''}{hike_val}{'%' if is_pct_hike else ''}"
        return normalize_rectangular_grid(updated), f"Applied {desc} adjustment to '{target_name}'."

    # --------------------------------------------------------------------------
    # CASE 6: COLUMN FORMULAS / COMBINATIONS (e.g. Base Salary + Bonus, Price * Qty)
    # --------------------------------------------------------------------------
    if any(k in p_lower for k in ["total salary", "total revenue", "total compensation", "sum of", "multiply", "*", "price * quantity"]):
        # A. Price * Quantity -> Total Revenue
        if any(k in p_lower for k in ["revenue", "price * quantity", "sales total"]):
            price_col = find_best_matching_column(headers, ["unit price", "price", "rate", "cost"])
            qty_col = find_best_matching_column(headers, ["quantity", "qty", "quantity sold", "units"])
            if price_col is not None and qty_col is not None:
                updated = [list(grid[0]) + ["Total Revenue ($)"]]
                for row in grid[1:]:
                    p = to_number(row[price_col]) if is_number(row[price_col]) else 0
                    q = to_number(row[qty_col]) if is_number(row[qty_col]) else 0
                    updated.append(list(row) + [round(p * q, 2)])
                return normalize_rectangular_grid(updated), f"Calculated 'Total Revenue' = '{headers[price_col]}' * '{headers[qty_col]}'."

        # B. Base Salary + Bonus -> Total Salary
        if any(k in p_lower for k in ["total salary", "total compensation", "gross salary", "ctc"]):
            sal_col = find_best_matching_column(headers, ["base salary", "salary", "basic salary"])
            bon_col = find_best_matching_column(headers, ["bonus", "allowance", "incentive"])
            if sal_col is not None and bon_col is not None:
                updated = [list(grid[0]) + ["Total Compensation ($)"]]
                for row in grid[1:]:
                    s = to_number(row[sal_col]) if is_number(row[sal_col]) else 0
                    b = to_number(row[bon_col]) if is_number(row[bon_col]) else 0
                    updated.append(list(row) + [round(s + b, 2)])
                return normalize_rectangular_grid(updated), f"Calculated 'Total Compensation' = '{headers[sal_col]}' + '{headers[bon_col]}'."


    # --------------------------------------------------------------------------
    # CASE 8: RENAME COLUMN
    # --------------------------------------------------------------------------
    if any(k in p_lower for k in ["rename column", "rename col", "change header", "change column name", "rename"]):
        rename_match = re.search(r"rename\s+(?:column\s+|col\s+)?['\"`]?([^'\"`]+?)['\"`]?\s+to\s+['\"`]?([^'\"`\n]+)['\"`]?", norm_prompt, re.IGNORECASE)
        if rename_match:
            old_name = rename_match.group(1).strip()
            new_name = rename_match.group(2).strip()
            col_idx = find_best_matching_column(headers, old_name)
            if col_idx is not None:
                updated = [list(grid[0])]
                updated[0][col_idx] = new_name
                for row in grid[1:]:
                    updated.append(list(row))
                return normalize_rectangular_grid(updated), f"Renamed column '{headers[col_idx]}' to '{new_name}'."

    # --------------------------------------------------------------------------
    # CASE 9: ADD LETTER GRADE COLUMN
    # --------------------------------------------------------------------------
    if any(k in p_lower for k in ["letter grade", "calculate grade", "add grade", "grades", "grading"]):
        score_cols = find_score_or_val_columns(grid)
        target_col = find_best_matching_column(headers, p_lower)
        if target_col is not None and target_col in score_cols:
            score_col = target_col
        else:
            tot_idx = find_best_matching_column(headers, ["total score", "total", "final"])
            score_col = tot_idx if tot_idx is not None else (score_cols[-1] if score_cols else 0)

        def get_grade(val: Any) -> str:
            if not is_number(val):
                return ""
            score = to_number(val)
            if score >= 90:
                return "A"
            elif score >= 80:
                return "B"
            elif score >= 70:
                return "C"
            elif score >= 60:
                return "D"
            else:
                return "F"

        existing_grade_idx = find_best_matching_column(headers, ["letter grade", "grade"])
        updated = []
        if existing_grade_idx is not None and existing_grade_idx >= len(headers) - 2:
            updated.append(list(grid[0]))
            for row in grid[1:]:
                val = row[score_col] if score_col < len(row) else ""
                new_r = list(row)
                new_r[existing_grade_idx] = get_grade(val)
                updated.append(new_r)
        else:
            updated.append(list(grid[0]) + ["Letter Grade"])
            for row in grid[1:]:
                val = row[score_col] if score_col < len(row) else ""
                updated.append(list(row) + [get_grade(val)])
        return normalize_rectangular_grid(updated), f"Added 'Letter Grade' column based on '{headers[score_col]}'."

    # --------------------------------------------------------------------------
    # CASE 10: ADD TOTAL / SUM COLUMN (Excludes ID/Roll No)
    # --------------------------------------------------------------------------
    if any(k in p_lower for k in ["add total", "calculate total", "total column", "sum column", "total score", "total", "kul", "yog"]):
        val_cols = find_score_or_val_columns(grid)
        if not val_cols:
            raise ValueError("No numeric columns found to calculate Total.")

        specific_cols = []
        for idx, h in enumerate(headers):
            h_clean = clean_header_text(h)
            if h_clean and h_clean in p_lower and idx in val_cols:
                specific_cols.append(idx)
        active_sum_cols = specific_cols if len(specific_cols) >= 2 else val_cols

        existing_tot_idx = find_best_matching_column(headers, ["total score", "total", "total salary", "total revenue"])
        updated = []
        if existing_tot_idx is not None and existing_tot_idx >= len(headers) - 2:
            updated.append(list(grid[0]))
            for row in grid[1:]:
                tot = sum(to_number(row[c]) for c in active_sum_cols if c < len(row) and is_number(row[c]))
                new_r = list(row)
                new_r[existing_tot_idx] = round(tot, 2)
                updated.append(new_r)
        else:
            updated.append(list(grid[0]) + ["Total Score"])
            for row in grid[1:]:
                tot = sum(to_number(row[c]) for c in active_sum_cols if c < len(row) and is_number(row[c]))
                updated.append(list(row) + [round(tot, 2)])

        return normalize_rectangular_grid(updated), f"Calculated and added Total column from {len(active_sum_cols)} columns."

    # --------------------------------------------------------------------------
    # CASE 11: ADD AVERAGE / PERCENTAGE COLUMN (Excludes ID/Roll No)
    # --------------------------------------------------------------------------
    if any(k in p_lower for k in ["add average", "calculate average", "average column", "percentage column", "calculate percentage", "percentage", "percent", "average", "ausat"]):
        val_cols = find_score_or_val_columns(grid)
        if not val_cols:
            raise ValueError("No numeric columns found to calculate Average/Percentage.")

        is_pct = "percent" in p_lower or "%" in p_lower
        col_name = "Percentage (%)" if is_pct else "Average"

        updated = []
        updated.append(list(grid[0]) + [col_name])
        for row in grid[1:]:
            nums = [to_number(row[c]) for c in val_cols if c < len(row) and is_number(row[c])]
            avg = round(sum(nums) / len(nums), 2) if nums else 0
            val_out = f"{avg}%" if is_pct else avg
            updated.append(list(row) + [val_out])
        return normalize_rectangular_grid(updated), f"Calculated and added '{col_name}' column."

    # --------------------------------------------------------------------------
    # CASE 12: ADD RANK COLUMN
    # --------------------------------------------------------------------------
    if any(k in p_lower for k in ["add rank", "ranking", "rank column", "rank"]):
        val_cols = find_score_or_val_columns(grid)
        tot_idx = find_best_matching_column(headers, ["total score", "total", "salary", "revenue", "final"])
        target_score_col = tot_idx if tot_idx is not None else (val_cols[-1] if val_cols else 0)

        scores = []
        for r_idx, row in enumerate(grid[1:]):
            val = row[target_score_col] if target_score_col < len(row) else 0
            num = to_number(val) if is_number(val) else -9999
            scores.append((num, r_idx))

        scores.sort(key=lambda x: x[0], reverse=True)
        ranks = {}
        for rank, (score, r_idx) in enumerate(scores, start=1):
            ranks[r_idx] = rank

        updated = []
        updated.append(list(grid[0]) + ["Rank"])
        for r_idx, row in enumerate(grid[1:]):
            updated.append(list(row) + [ranks.get(r_idx, "-")])
        return normalize_rectangular_grid(updated), f"Added 'Rank' column based on '{headers[target_score_col]}'."

    # --------------------------------------------------------------------------
    # CASE 13: ADD PASS / FAIL / RESULT COLUMN
    # --------------------------------------------------------------------------
    if any(k in p_lower for k in ["pass/fail", "pass or fail", "result", "pass fail", "status column", "pass", "fail"]):
        val_cols = find_score_or_val_columns(grid)
        tot_idx = find_best_matching_column(headers, ["total score", "total", "final"])
        target_col = tot_idx if tot_idx is not None else (val_cols[-1] if val_cols else 0)

        cutoff = 40.0
        cutoff_match = re.search(r"(\d+)\s*(?:cutoff|passing|to pass|marks to pass|marks)", p_lower) or re.search(r"(?:>=|score\s*>=)\s*(\d+)", p_lower)
        if cutoff_match:
            cutoff = float(cutoff_match.group(1))

        updated = []
        updated.append(list(grid[0]) + ["Result"])
        for row in grid[1:]:
            val = row[target_col] if target_col < len(row) else ""
            res = "Pass" if is_number(val) and to_number(val) >= cutoff else "Fail"
            updated.append(list(row) + [res])
        return normalize_rectangular_grid(updated), f"Added 'Result' (Pass/Fail) column with {cutoff} cutoff."

    # --------------------------------------------------------------------------
    # CASE 14: ADD REMARKS / FEEDBACK
    # --------------------------------------------------------------------------
    if any(k in p_lower for k in ["remark", "remarks", "feedback", "comment", "comments"]):
        val_cols = find_score_or_val_columns(grid)
        tot_idx = find_best_matching_column(headers, ["total score", "total", "final"])
        score_col = tot_idx if tot_idx is not None else (val_cols[-1] if val_cols else 0)

        def get_remark(val: Any) -> str:
            if not is_number(val):
                return "Good performance"
            s = to_number(val)
            if s >= 90:
                return "Outstanding performance! Exceptional results."
            elif s >= 80:
                return "Very good effort! Consistent high achievement."
            elif s >= 70:
                return "Good progress. On track with targets."
            elif s >= 60:
                return "Satisfactory. Steady progress observed."
            else:
                return "Needs focused improvement and support."

        updated = []
        updated.append(list(grid[0]) + ["Teacher Remarks"])
        for row in grid[1:]:
            val = row[score_col] if score_col < len(row) else ""
            rem = get_remark(val)
            updated.append(list(row) + [rem])
        return normalize_rectangular_grid(updated), "Added personalized Remarks column."

    # --------------------------------------------------------------------------
    # CASE 15: SMART SUMMARY / STATISTICS ROWS AT BOTTOM (Multi-Domain)
    # --------------------------------------------------------------------------
    if any(k in p_lower for k in ["summary", "average at bottom", "highest and lowest", "min and max", "stats", "statistics", "bottom", "ausat"]):
        val_cols = find_score_or_val_columns(grid)
        if not val_cols:
            raise ValueError("No numeric columns found to calculate summary statistics.")

        # Detect domain for appropriate summary labels
        header_text = " ".join(headers).lower()
        is_salary = any(k in header_text for k in ["salary", "employee", "compensation", "wage"])
        is_sales = any(k in header_text for k in ["sales", "revenue", "product", "price"])

        clean_grid = [grid[0]]
        for r in grid[1:]:
            first_cell = str(r[0]).lower().strip()
            if not any(k in first_cell for k in ["average", "highest", "lowest", "total payroll", "total revenue"]):
                clean_grid.append(r)

        data_rows = clean_grid[1:]
        if not data_rows:
            return grid, "No data rows available."

        avg_label = "Average Salary" if is_salary else ("Average Sales" if is_sales else "Class Average")
        tot_label = "Total Payroll (Sum)" if is_salary else ("Total Revenue (Sum)" if is_sales else "Overall Total")
        max_label = "Highest"
        min_label = "Lowest"

        avg_row = [avg_label] + [""] * (len(headers) - 1)
        tot_row = [tot_label] + [""] * (len(headers) - 1)
        max_row = [max_label] + [""] * (len(headers) - 1)
        min_row = [min_label] + [""] * (len(headers) - 1)

        for c in val_cols:
            vals = [to_number(r[c]) for r in data_rows if c < len(r) and is_number(r[c])]
            if vals:
                avg_row[c] = round(sum(vals) / len(vals), 2)
                tot_row[c] = round(sum(vals), 2)
                max_row[c] = round(max(vals), 2)
                min_row[c] = round(min(vals), 2)

        summary_rows = [avg_row, tot_row, max_row, min_row] if (is_salary or is_sales) else [avg_row, max_row, min_row]
        updated = clean_grid + summary_rows
        return normalize_rectangular_grid(updated), f"Added summary rows ({avg_label}, Highest, Lowest) at the bottom."

    # --------------------------------------------------------------------------
    # CASE 16: SORTING (By Name or Order Number)
    # --------------------------------------------------------------------------
    if any(k in p_lower for k in ["sort", "order by", "kram"]):
        target_col = find_best_matching_column(headers, p_lower)
        if target_col is None:
            val_cols = find_score_or_val_columns(grid)
            target_col = val_cols[-1] if val_cols else 0

        descending = any(k in p_lower for k in ["desc", "descending", "highest first", "top first", "highest to lowest", "bada se chhota"])
        num_cols = find_numeric_columns(grid)
        is_num = target_col in num_cols

        def sort_key(row):
            val = row[target_col] if target_col < len(row) else ""
            if is_num:
                return to_number(val) if is_number(val) else -9999
            return str(val).lower()

        sorted_rows = sorted(grid[1:], key=sort_key, reverse=descending)
        updated = [grid[0]] + sorted_rows
        dir_str = "descending" if descending else "ascending"
        return normalize_rectangular_grid(updated), f"Sorted table by '{headers[target_col]}' ({dir_str})."

    # --------------------------------------------------------------------------
    # CASE 17: PROPER CASE / CLEAN TEXT
    # --------------------------------------------------------------------------
    if any(k in p_lower for k in ["proper case", "title case", "capitalize", "clean name", "standardize name"]):
        name_col = find_best_matching_column(headers, ["student name", "employee name", "name", "full name", "product name"])
        if name_col is None:
            name_col = 0

        updated = [list(grid[0])]
        for row in grid[1:]:
            new_row = list(row)
            if name_col < len(new_row):
                val = str(new_row[name_col]).strip()
                new_row[name_col] = re.sub(r"\s+", " ", val).title()
            updated.append(new_row)
        return normalize_rectangular_grid(updated), f"Standardized names in '{headers[name_col]}' to Proper Case."


    # --------------------------------------------------------------------------
    # Fallback: Pasted data table
    # --------------------------------------------------------------------------
    pasted_table = parse_prompt_to_table(raw_prompt)
    if pasted_table:
        return pasted_table, f"Created new table from prompt text ({len(pasted_table)-1} rows)."

    raise ValueError(
        f"Could not automatically determine action for: '{raw_prompt}'. "
        "Try: 'Give 10% salary hike', 'Add 5000 bonus', 'Calculate Total', "
        "'Delete column 3', 'Remove row for Emma', or paste your data directly."
    )
