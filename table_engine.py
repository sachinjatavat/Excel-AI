"""
table_engine.py - Specialized Offline Excel Table AI Engine
Processes table operations, text-to-table parsing, column additions,
deletions, modifications, curving, formulas, and remarks with ZERO external downloads.
"""

import re
import csv
import io
import statistics
from typing import List, Any, Tuple, Optional, Dict


def is_number(val: Any) -> bool:
    if val is None or val == "":
        return False
    try:
        float(str(val).replace(",", "").replace("%", "").strip())
        return True
    except ValueError:
        return False


def to_number(val: Any) -> float:
    return float(str(val).replace(",", "").replace("%", "").strip())


def clean_cell(val: Any) -> Any:
    if val is None:
        return ""
    if isinstance(val, (int, float)):
        return val
    s = str(val).strip()
    if is_number(s):
        try:
            if "." in s:
                f = float(s)
                return round(f, 2)
            return int(s)
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
# 1. TEXT-TO-TABLE PARSER
# Converts raw pasted data, comma/tab/colon lists, or creation requests into 2D grid
# ==============================================================================

def parse_markdown_table(text: str) -> Optional[List[List[Any]]]:
    lines = [line.strip() for line in text.strip().splitlines() if line.strip().startswith("|")]
    if len(lines) >= 2:
        grid = []
        for line in lines:
            if re.match(r"^\|[\s\-:|]+\|$", line):
                continue
            cells = [c.strip() for c in line.split("|")]
            if len(cells) >= 2:
                row = cells[1:-1]
                if any(row):
                    grid.append([clean_cell(c) for c in row])
        if len(grid) >= 1:
            return normalize_rectangular_grid(grid)
    return None


def parse_delimited_text(text: str) -> Optional[List[List[Any]]]:
    lines = [l.strip() for l in text.strip().splitlines() if l.strip()]
    if not lines:
        return None

    # Detect delimiter: comma, tab, semicolon, pipe
    sample = "\n".join(lines[:10])
    comma_count = sample.count(",")
    tab_count = sample.count("\t")
    semi_count = sample.count(";")
    pipe_count = sample.count("|")

    chosen_delim = None
    if tab_count >= len(lines) - 1 and tab_count > 0:
        chosen_delim = "\t"
    elif comma_count >= len(lines) - 1 and comma_count > 0:
        chosen_delim = ","
    elif semi_count >= len(lines) - 1 and semi_count > 0:
        chosen_delim = ";"
    elif pipe_count >= len(lines) - 1 and pipe_count > 0:
        chosen_delim = "|"

    if chosen_delim:
        try:
            reader = csv.reader(io.StringIO(text), delimiter=chosen_delim)
            rows = [[clean_cell(c) for c in row] for row in reader if any(row)]
            if len(rows) > 0 and len(rows[0]) > 1:
                return normalize_rectangular_grid(rows)
        except Exception:
            pass

    return None


def parse_key_value_lines(text: str) -> Optional[List[List[Any]]]:
    # Look for patterns like "Name: 85" or "1. Rahul - 85 - A" or "Rahul, 85, 90"
    lines = [l.strip() for l in text.strip().splitlines() if l.strip()]
    
    # Try comma-separated items within single line: "Rahul: 85, Priya: 92, Amit: 78"
    if len(lines) == 1 and (":" in lines[0] or "-" in lines[0]) and ("," in lines[0] or ";" in lines[0]):
        parts = re.split(r"[,;]\s*", lines[0])
        if len(parts) >= 2:
            lines = parts

    pairs = []
    for line in lines:
        # Strip numbering: "1.", "1 -", "1)"
        cleaned = re.sub(r"^\d+[\.\)\-]\s*", "", line)
        if ":" in cleaned:
            parts = [p.strip() for p in cleaned.split(":", 1)]
            pairs.append(parts)
        elif " - " in cleaned:
            parts = [p.strip() for p in cleaned.split(" - ")]
            pairs.append(parts)
        elif "\t" in cleaned:
            parts = [p.strip() for p in cleaned.split("\t")]
            pairs.append(parts)
        elif "," in cleaned:
            parts = [p.strip() for p in cleaned.split(",")]
            pairs.append(parts)
        else:
            # Space separated: "Rahul 85"
            parts = cleaned.split()
            if len(parts) >= 2:
                pairs.append(parts)

    if len(pairs) >= 2:
        # Check if first row looks like a header or if we should add a header
        first = pairs[0]
        has_header = not any(is_number(c) for c in first) and all(isinstance(c, str) for c in first)
        
        # If no header, synthesize header
        col_count = max(len(p) for p in pairs)
        grid = []
        if not has_header:
            if col_count == 2:
                grid.append(["Name", "Value"])
            else:
                grid.append(["Item"] + [f"Field {i+1}" for i in range(col_count - 1)])
        
        for p in pairs:
            grid.append([clean_cell(c) for c in p])
        return normalize_rectangular_grid(grid)

    return None


def generate_template_table(prompt: str) -> Optional[List[List[Any]]]:
    lower = prompt.lower()
    
    # Check if prompt specifies columns: e.g. "columns: Roll No, Name, Math, Science"
    col_match = re.search(r"(?:columns|fields|headers)[:\s]+([^.]+)", prompt, re.IGNORECASE)
    if col_match:
        raw_cols = col_match.group(1).strip()
        cols = [c.strip(" '\"`[]") for c in re.split(r"[,;]|\band\b", raw_cols) if c.strip(" '\"`[]")]
        if cols:
            # Generate header with 5 sample rows or empty rows
            sample_names = ["Emma Watson", "Liam Johnson", "Sophia Martinez", "Noah Brown", "Olivia Davis"]
            rows = [cols]
            for i, name in enumerate(sample_names):
                row = []
                for c in cols:
                    c_low = c.lower()
                    if "roll" in c_low or "id" in c_low:
                        row.append(101 + i)
                    elif "name" in c_low or "student" in c_low:
                        row.append(name)
                    elif any(sub in c_low for sub in ["math", "phys", "chem", "sci", "eng", "score", "mark", "exam", "quiz"]):
                        row.append(70 + (i * 5) % 28)
                    elif "grade" in c_low:
                        row.append("A" if i % 2 == 0 else "B")
                    elif "attend" in c_low:
                        row.append(85 + (i * 3) % 15)
                    else:
                        row.append("")
                rows.append(row)
            return normalize_rectangular_grid(rows)

    # General gradebook / attendance table creation
    if any(k in lower for k in ["create table", "new table", "gradebook", "attendance sheet", "sample table", "student table", "make a table"]):
        if "attendance" in lower:
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
        else:
            # Default Student Gradebook
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
    p_low = prompt.lower()
    # 1. Try template generation if explicitly requested
    if any(k in p_low for k in ["create", "new table", "generate", "make a table", "columns:", "headers:", "gradebook", "attendance sheet"]):
        grid = generate_template_table(prompt)
        if grid:
            return grid
    # 2. Try markdown table
    grid = parse_markdown_table(prompt)
    if grid:
        return grid
    # 3. Try key-value lists (e.g. Rahul: 85, Priya: 92 or Rahul - 85)
    if ":" in prompt or " - " in prompt:
        grid = parse_key_value_lines(prompt)
        if grid:
            return grid
    # 4. Try CSV / TSV / Delimited
    grid = parse_delimited_text(prompt)
    if grid:
        return grid
    # 5. Try key-value lists general
    grid = parse_key_value_lines(prompt)
    if grid:
        return grid
    # 6. Final attempt template generation
    grid = generate_template_table(prompt)
    if grid:
        return grid
    return None


# ==============================================================================
# 2. TABLE MANIPULATION ENGINE
# Handles Add Column, Remove Column, Curving, Statistics, Remarks, Filtering, Sorting
# ==============================================================================

def find_column_index(headers: List[Any], keywords: List[str]) -> Optional[int]:
    """Finds best matching column index by keyword list"""
    for idx, h in enumerate(headers):
        h_str = str(h).lower()
        for kw in keywords:
            if kw.lower() == h_str:
                return idx
    for idx, h in enumerate(headers):
        h_str = str(h).lower()
        for kw in keywords:
            if kw.lower() in h_str:
                return idx
    return None


def find_numeric_columns(grid: List[List[Any]]) -> List[int]:
    """Finds all columns that are primarily numeric (excluding header)"""
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


def execute_table_instruction(grid: List[List[Any]], prompt: str) -> Tuple[List[List[Any]], str]:
    """
    Executes natural language instruction on existing 2D grid.
    Returns (updated_grid, action_description).
    """
    grid = normalize_rectangular_grid(grid)
    if not grid or len(grid) == 0:
        # Sheet is empty: try generating from prompt
        new_grid = parse_prompt_to_table(prompt)
        if new_grid:
            return new_grid, f"Created new table with {len(new_grid)-1} rows from your prompt."
        raise ValueError("Worksheet is empty and no table data could be detected in your prompt.")

    headers = [str(c) for c in grid[0]]
    p_lower = prompt.lower()

    # --------------------------------------------------------------------------
    # Case A: REMOVE / DELETE COLUMN
    # --------------------------------------------------------------------------
    if any(w in p_lower for w in ["remove column", "delete column", "drop column", "remove", "delete"]):
        # Find which column to remove
        target_col = None
        target_name = ""
        for idx, h in enumerate(headers):
            h_clean = re.sub(r"\s*\(.*?\)", "", h).strip().lower()
            if h_clean and (h_clean in p_lower or h.lower() in p_lower):
                target_col = idx
                target_name = h
                break

        # Check for column by number: "remove column 2"
        if target_col is None:
            col_num_match = re.search(r"column\s+(\d+)", p_lower)
            if col_num_match:
                c_num = int(col_num_match.group(1)) - 1
                if 0 <= c_num < len(headers):
                    target_col = c_num
                    target_name = headers[c_num]

        if target_col is not None:
            updated = []
            for row in grid:
                new_row = [row[i] for i in range(len(row)) if i != target_col]
                updated.append(new_row)
            return normalize_rectangular_grid(updated), f"Removed column '{target_name}'."

    # --------------------------------------------------------------------------
    # Case B: REMOVE / FILTER ROWS
    # --------------------------------------------------------------------------
    if "remove empty" in p_lower or "delete empty" in p_lower:
        updated = [grid[0]]
        for row in grid[1:]:
            if any(str(c).strip() != "" for c in row):
                updated.append(row)
        return normalize_rectangular_grid(updated), "Removed empty rows."

    if any(k in p_lower for k in ["remove failing", "delete fail", "remove rows where", "delete rows where"]):
        # Find score column
        score_col = find_column_index(headers, ["total", "score", "marks", "final", "grade"])
        num_cols = find_numeric_columns(grid)
        if score_col is None and num_cols:
            score_col = num_cols[-1]

        cutoff = 40.0
        cutoff_match = re.search(r"[<]\s*(\d+)", prompt) or re.search(r"less than\s*(\d+)", p_lower) or re.search(r"below\s*(\d+)", p_lower)
        if cutoff_match:
            cutoff = float(cutoff_match.group(1))

        if score_col is not None:
            updated = [grid[0]]
            removed_count = 0
            for row in grid[1:]:
                val = row[score_col]
                if is_number(val) and to_number(val) < cutoff:
                    removed_count += 1
                else:
                    updated.append(row)
            return normalize_rectangular_grid(updated), f"Filtered out {removed_count} rows with score below {cutoff}."

    # --------------------------------------------------------------------------
    # Case C: ADD LETTER GRADE COLUMN
    # --------------------------------------------------------------------------
    if any(k in p_lower for k in ["letter grade", "calculate grade", "add grade", "grades", "grading"]):
        score_col = find_column_index(headers, ["total score", "total", "score", "marks", "final exam", "percentage"])
        num_cols = find_numeric_columns(grid)
        if score_col is None and num_cols:
            score_col = num_cols[-1]

        if score_col is None:
            raise ValueError("Could not find a numeric score column to calculate letter grades.")

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

        updated = []
        updated.append(list(grid[0]) + ["Letter Grade"])
        for row in grid[1:]:
            val = row[score_col] if score_col < len(row) else ""
            grade = get_grade(val)
            updated.append(list(row) + [grade])
        return normalize_rectangular_grid(updated), f"Added 'Letter Grade' column based on '{headers[score_col]}'."

    # --------------------------------------------------------------------------
    # Case D: ADD TEACHER REMARKS / FEEDBACK
    # --------------------------------------------------------------------------
    if any(k in p_lower for k in ["remark", "remarks", "feedback", "comment", "comments", "student note"]):
        score_col = find_column_index(headers, ["total score", "total", "score", "marks", "final", "grade"])
        num_cols = find_numeric_columns(grid)
        if score_col is None and num_cols:
            score_col = num_cols[-1]

        def get_remark(val: Any) -> str:
            if not is_number(val):
                return "Keep working hard!"
            s = to_number(val)
            if s >= 90:
                return "Excellent performance! Outstanding grasp of concepts."
            elif s >= 80:
                return "Very good effort! Consistent work and strong potential."
            elif s >= 70:
                return "Good progress. With extra review, can reach top marks."
            elif s >= 60:
                return "Satisfactory understanding. Daily practice recommended."
            else:
                return "Needs additional support and focused revision."

        updated = []
        updated.append(list(grid[0]) + ["Teacher Remarks"])
        for row in grid[1:]:
            val = row[score_col] if score_col is not None and score_col < len(row) else ""
            rem = get_remark(val)
            updated.append(list(row) + [rem])
        return normalize_rectangular_grid(updated), "Added personalized 'Teacher Remarks' column."

    # --------------------------------------------------------------------------
    # Case E: ADD PASS / FAIL / RESULT
    # --------------------------------------------------------------------------
    if any(k in p_lower for k in ["pass/fail", "pass or fail", "result", "pass fail", "status column"]):
        score_col = find_column_index(headers, ["total", "score", "marks", "final", "total score"])
        num_cols = find_numeric_columns(grid)
        if score_col is None and num_cols:
            score_col = num_cols[-1]

        cutoff = 40.0
        cutoff_match = re.search(r"(\d+)\s*(?:cutoff|passing|to pass|marks to pass)", p_lower)
        if cutoff_match:
            cutoff = float(cutoff_match.group(1))

        updated = []
        updated.append(list(grid[0]) + ["Result"])
        for row in grid[1:]:
            val = row[score_col] if score_col is not None and score_col < len(row) else ""
            res = "Pass" if is_number(val) and to_number(val) >= cutoff else "Fail"
            updated.append(list(row) + [res])
        return normalize_rectangular_grid(updated), f"Added 'Result' (Pass/Fail) column with {cutoff} cutoff."

    # --------------------------------------------------------------------------
    # Case F: ADD TOTAL / SUM COLUMN
    # --------------------------------------------------------------------------
    if any(k in p_lower for k in ["add total", "calculate total", "total column", "sum column"]):
        num_cols = find_numeric_columns(grid)
        if not num_cols:
            raise ValueError("No numeric columns found to calculate Total.")

        updated = []
        updated.append(list(grid[0]) + ["Total Score"])
        for row in grid[1:]:
            total = sum(to_number(row[c]) for c in num_cols if c < len(row) and is_number(row[c]))
            updated.append(list(row) + [round(total, 2)])
        return normalize_rectangular_grid(updated), f"Calculated and added 'Total Score' column from {len(num_cols)} columns."

    # --------------------------------------------------------------------------
    # Case G: ADD AVERAGE / PERCENTAGE COLUMN
    # --------------------------------------------------------------------------
    if any(k in p_lower for k in ["add average", "calculate average", "percentage column", "calculate percentage"]):
        num_cols = find_numeric_columns(grid)
        if not num_cols:
            raise ValueError("No numeric columns found to calculate Average.")

        updated = []
        col_name = "Percentage (%)" if "percent" in p_lower else "Average"
        updated.append(list(grid[0]) + [col_name])
        for row in grid[1:]:
            nums = [to_number(row[c]) for c in num_cols if c < len(row) and is_number(row[c])]
            avg = round(sum(nums) / len(nums), 2) if nums else 0
            updated.append(list(row) + [avg])
        return normalize_rectangular_grid(updated), f"Calculated and added '{col_name}' column."

    # --------------------------------------------------------------------------
    # Case H: ADD RANK COLUMN
    # --------------------------------------------------------------------------
    if any(k in p_lower for k in ["add rank", "ranking", "rank column"]):
        score_col = find_column_index(headers, ["total score", "total", "score", "marks", "final"])
        num_cols = find_numeric_columns(grid)
        if score_col is None and num_cols:
            score_col = num_cols[-1]

        if score_col is None:
            raise ValueError("Could not find a score column to calculate Rank.")

        # Extract scores with row indices
        scores = []
        for r_idx, row in enumerate(grid[1:]):
            val = row[score_col] if score_col < len(row) else 0
            num = to_number(val) if is_number(val) else -9999
            scores.append((num, r_idx))

        # Rank descending
        scores.sort(key=lambda x: x[0], reverse=True)
        ranks = {}
        for rank, (score, r_idx) in enumerate(scores, start=1):
            ranks[r_idx] = rank

        updated = []
        updated.append(list(grid[0]) + ["Rank"])
        for r_idx, row in enumerate(grid[1:]):
            updated.append(list(row) + [ranks.get(r_idx, "-")])
        return normalize_rectangular_grid(updated), "Added student 'Rank' column."

    # --------------------------------------------------------------------------
    # Case I: GRADE CURVING / BONUS POINTS
    # --------------------------------------------------------------------------
    if any(k in p_lower for k in ["curve", "add", "bonus", "increase", "points"]) and any(c in p_lower for c in ["midterm", "quiz", "final", "exam", "score", "marks"]):
        # Extract curve amount: e.g. "5-point curve", "add 5 points", "curve by 10"
        amount_match = re.search(r"(\d+(?:\.\d+)?)\s*(?:point|pt|mark|%)", p_lower) or re.search(r"(?:curve|add|bonus|increase by)\s*(\d+(?:\.\d+)?)", p_lower)
        curve_val = float(amount_match.group(1)) if amount_match else 5.0
        cap_val = 100.0 if "100" in p_lower or "cap" in p_lower else None

        # Find target column
        target_col = None
        target_name = ""
        for idx, h in enumerate(headers):
            h_clean = re.sub(r"\s*\(.*?\)", "", h).strip().lower()
            if h_clean and h_clean in p_lower:
                target_col = idx
                target_name = h
                break

        num_cols = find_numeric_columns(grid)
        if target_col is None and num_cols:
            target_col = num_cols[0]
            target_name = headers[target_col]

        if target_col is not None:
            updated = [list(grid[0])]
            for row in grid[1:]:
                new_row = list(row)
                val = new_row[target_col]
                if is_number(val):
                    curved = to_number(val) + curve_val
                    if cap_val is not None:
                        curved = min(curved, cap_val)
                    new_row[target_col] = round(curved, 2)
                updated.append(new_row)
            cap_str = f" (capped at {cap_val})" if cap_val else ""
            return normalize_rectangular_grid(updated), f"Applied +{curve_val} curve to '{target_name}'{cap_str}."

    # --------------------------------------------------------------------------
    # Case J: CLEAN / PROPER CASE NAMES
    # --------------------------------------------------------------------------
    if any(k in p_lower for k in ["proper case", "title case", "capitalize", "clean name", "standardize name"]):
        name_col = find_column_index(headers, ["student name", "name", "student", "full name"])
        if name_col is None:
            name_col = 0

        updated = [list(grid[0])]
        for row in grid[1:]:
            new_row = list(row)
            if name_col < len(new_row):
                val = str(new_row[name_col]).strip()
                # Proper title casing
                new_row[name_col] = re.sub(r"\s+", " ", val).title()
            updated.append(new_row)
        return normalize_rectangular_grid(updated), f"Standardized names in '{headers[name_col]}' to Proper Case."

    # --------------------------------------------------------------------------
    # Case K: SUMMARY / STATISTICS ROWS AT BOTTOM
    # --------------------------------------------------------------------------
    if any(k in p_lower for k in ["summary", "average at bottom", "highest and lowest", "min and max", "class stats", "statistics"]):
        num_cols = find_numeric_columns(grid)
        if not num_cols:
            raise ValueError("No numeric columns found to calculate summary statistics.")

        # Check existing summary rows so we don't duplicate
        existing_first_cols = [str(r[0]).lower().strip() for r in grid[1:]]
        clean_grid = [grid[0]]
        for r in grid[1:]:
            if str(r[0]).lower().strip() not in ["average", "class average", "highest", "lowest", "median"]:
                clean_grid.append(r)

        data_rows = clean_grid[1:]
        if not data_rows:
            return grid, "No data rows available."

        avg_row = ["Class Average"] + [""] * (len(headers) - 1)
        max_row = ["Highest Score"] + [""] * (len(headers) - 1)
        min_row = ["Lowest Score"] + [""] * (len(headers) - 1)

        for c in num_cols:
            vals = [to_number(r[c]) for r in data_rows if c < len(r) and is_number(r[c])]
            if vals:
                avg_row[c] = round(sum(vals) / len(vals), 2)
                max_row[c] = round(max(vals), 2)
                min_row[c] = round(min(vals), 2)

        updated = clean_grid + [avg_row, max_row, min_row]
        return normalize_rectangular_grid(updated), "Added summary rows (Class Average, Highest, Lowest) at the bottom."

    # --------------------------------------------------------------------------
    # Case L: SORTING
    # --------------------------------------------------------------------------
    if any(k in p_lower for k in ["sort", "order by"]):
        target_col = None
        for idx, h in enumerate(headers):
            h_clean = re.sub(r"\s*\(.*?\)", "", h).strip().lower()
            if h_clean and h_clean in p_lower:
                target_col = idx
                break

        if target_col is None:
            num_cols = find_numeric_columns(grid)
            target_col = num_cols[-1] if num_cols else 0

        descending = any(k in p_lower for k in ["desc", "descending", "highest first", "top first", "highest to lowest"])
        is_num_col = target_col in find_numeric_columns(grid)

        def sort_key(row):
            val = row[target_col] if target_col < len(row) else ""
            if is_num_col:
                return to_number(val) if is_number(val) else -9999
            return str(val).lower()

        sorted_rows = sorted(grid[1:], key=sort_key, reverse=descending)
        updated = [grid[0]] + sorted_rows
        dir_str = "descending" if descending else "ascending"
        return normalize_rectangular_grid(updated), f"Sorted table by '{headers[target_col]}' ({dir_str})."

    # --------------------------------------------------------------------------
    # Case M: GENERIC COLUMN ADDITION (e.g. "Add column Section with value A")
    # --------------------------------------------------------------------------
    add_col_match = re.search(r"add\s+column\s+['\"`]?([^'\"`]+?)['\"`]?\s+(?:with\s+value|with)?\s*['\"`]?([^'\"`\n]+)?", prompt, re.IGNORECASE)
    if add_col_match:
        new_col_name = add_col_match.group(1).strip()
        default_val = (add_col_match.group(2) or "").strip()
        updated = []
        updated.append(list(grid[0]) + [new_col_name])
        for row in grid[1:]:
            updated.append(list(row) + [default_val])
        return normalize_rectangular_grid(updated), f"Added column '{new_col_name}' with value '{default_val}'."

    # --------------------------------------------------------------------------
    # Fallback: If nothing matched, check if prompt itself contains pasted table
    # --------------------------------------------------------------------------
    pasted_table = parse_prompt_to_table(prompt)
    if pasted_table:
        return pasted_table, f"Created new table from prompt text ({len(pasted_table)-1} rows)."

    raise ValueError(f"Could not automatically determine action for: '{prompt}'. Try specifying 'Add Letter Grades', 'Add Remarks', 'Add 5-point curve', 'Add Summary Rows', or paste your data directly.")
