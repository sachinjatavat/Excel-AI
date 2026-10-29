"""
formula_engine.py - Excel Formula Generator, Explainer, and Fixer
Handles:
1. Generating Excel formulas from plain English with dynamic column letter mapping.
2. Explaining existing Excel formulas in clear, plain English.
3. Detecting and fixing Excel formula errors (#REF!, #VALUE!, #N/A, #DIV/0!, #NAME?).
"""

import re
from typing import List, Dict, Any, Tuple, Optional


def column_index_to_letter(col_idx: int) -> str:
    """Converts 0-based column index to Excel column letter (0 -> A, 1 -> B, 26 -> AA)."""
    result = ""
    col_idx += 1
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def letter_to_column_index(letter: str) -> int:
    """Converts Excel column letter to 0-based index ('A' -> 0, 'B' -> 1)."""
    result = 0
    for char in letter.upper():
        result = result * 26 + (ord(char) - 64)
    return result - 1


def map_headers_to_columns(headers: List[str]) -> Dict[str, str]:
    """
    Creates a mapping from lowercase header names and keywords to Excel column letters.
    e.g. {'salary': 'C', 'base salary': 'C', 'performance': 'D', ...}
    """
    mapping: Dict[str, str] = {}
    for idx, h in enumerate(headers):
        let = column_index_to_letter(idx)
        h_clean = str(h).strip().lower()
        if h_clean:
            mapping[h_clean] = let
            # Also store individual words for compound headers like "Total Score (100)" -> "total score", "score"
            cleaned_simple = re.sub(r"[\(\)\[\]0-9%\$₹\.]", "", h_clean).strip()
            if cleaned_simple and cleaned_simple != h_clean:
                mapping[cleaned_simple] = let
            for word in cleaned_simple.split():
                if len(word) >= 3 and word not in ["the", "and", "for", "with"]:
                    if word not in mapping:
                        mapping[word] = let
    return mapping


def generate_formula_from_prompt(
    prompt: str,
    headers: List[str],
    current_row: int = 2
) -> Dict[str, Any]:
    """
    Translates plain English prompt into an Excel formula dynamically mapped to headers.
    e.g. "Calculate the bonus as 10% of salary if performance is above 80"
    -> =IF(C2>80, B2*0.1, 0)
    """
    p_lower = prompt.lower().strip()
    col_map = map_headers_to_columns(headers)
    r = current_row

    # Helper to resolve column letter from word
    def get_col(term: str) -> Optional[str]:
        t = term.lower().strip()
        if t in col_map:
            return col_map[t]
        for k, v in col_map.items():
            if t == k or t in k or k in t:
                return v
        return None

    formula = ""
    explanation = ""
    suggested_header = "Calculated Column"

    # 1. Condition: IF ... THEN ... ELSE
    # e.g. "Calculate the bonus as 10% of salary if performance is above 80"
    # or "if score > 50 then Pass else Fail"
    if "if" in p_lower:
        # Check: bonus as 10% of salary if performance is above 80
        match_bonus_if = re.search(
            r"(?:calculate\s+)?(?:the\s+)?([a-zA-Z\s]+?)\s+as\s+(\d+(?:\.\d+)?%?)\s+of\s+([a-zA-Z\s]+?)\s+if\s+([a-zA-Z\s]+?)\s+(?:is\s+)?(above|greater than|>|>=|at least|below|less than|<|<=)\s*(\d+(?:\.\d+)?)",
            p_lower
        )
        if match_bonus_if:
            col_target_name = match_bonus_if.group(1).strip().title()
            pct_raw = match_bonus_if.group(2).strip()
            val_col_name = match_bonus_if.group(3).strip()
            cond_col_name = match_bonus_if.group(4).strip()
            op_raw = match_bonus_if.group(5).strip()
            threshold = match_bonus_if.group(6).strip()

            val_col = get_col(val_col_name) or "B"
            cond_col = get_col(cond_col_name) or "C"

            # Parse percentage
            if "%" in pct_raw:
                pct_mult = float(pct_raw.replace("%", "")) / 100.0
            else:
                pct_mult = float(pct_raw) / 100.0 if float(pct_raw) > 1 else float(pct_raw)

            op = ">" if op_raw in ["above", "greater than", ">"] else (">=" if op_raw in [">=", "at least"] else ("<" if op_raw in ["below", "less than", "<"] else "<="))

            formula = f"=IF({cond_col}{r}{op}{threshold}, {val_col}{r}*{pct_mult}, 0)"
            explanation = (
                f"Checks if cell {cond_col}{r} ({cond_col_name.title()}) {op} {threshold}. "
                f"If True, multiplies {val_col}{r} ({val_col_name.title()}) by {pct_mult} ({pct_raw}); "
                f"otherwise returns 0."
            )
            suggested_header = col_target_name or "Bonus"
            return {
                "status": "success",
                "formula": formula,
                "suggested_header": suggested_header,
                "explanation": explanation,
                "type": "IF_LOGICAL"
            }

        # Check: If score >= 60 then Pass else Fail
        match_generic_if = re.search(
            r"if\s+([a-zA-Z\s]+?)\s*(>=|<=|>|<|=|is greater than|is less than|equals)\s*(\d+|['\"][^'\"]+['\"])\s*(?:then\s+)?(['\"][^'\"]+['\"]|\d+|[a-zA-Z]+)(?:\s*(?:else|otherwise)\s*(['\"][^'\"]+['\"]|\d+|[a-zA-Z]+))?",
            p_lower
        )
        if match_generic_if:
            cond_col_name = match_generic_if.group(1).strip()
            op = match_generic_if.group(2).strip()
            val_cmp = match_generic_if.group(3).strip()
            true_val = match_generic_if.group(4).strip()
            false_val = (match_generic_if.group(5) or '""').strip()

            if op in ["is greater than"]: op = ">"
            elif op in ["is less than"]: op = "<"
            elif op in ["equals"]: op = "="

            cond_col = get_col(cond_col_name) or "A"
            if not true_val.isdigit() and not (true_val.startswith('"') or true_val.startswith("'")):
                true_val = f'"{true_val}"'
            if not false_val.isdigit() and not (false_val.startswith('"') or false_val.startswith("'")):
                false_val = f'"{false_val}"'

            formula = f"=IF({cond_col}{r}{op}{val_cmp}, {true_val}, {false_val})"
            explanation = f"Evaluates if {cond_col}{r} satisfies {op} {val_cmp}. Returns {true_val} if true, else {false_val}."
            return {
                "status": "success",
                "formula": formula,
                "suggested_header": "Status",
                "explanation": explanation,
                "type": "IF_LOGICAL"
            }

    # 2. SUM / TOTAL
    # e.g. "Sum of Base Salary and Bonus" or "Total of columns B to E"
    if any(k in p_lower for k in ["sum of", "total of", "add columns", "sum columns"]):
        cols_found = []
        for word in p_lower.split():
            c = get_col(word)
            if c and c not in cols_found:
                cols_found.append(c)
        if len(cols_found) >= 2:
            formula = f"=SUM({cols_found[0]}{r}:{cols_found[-1]}{r})" if len(cols_found) > 2 else f"={cols_found[0]}{r}+{cols_found[1]}{r}"
            explanation = f"Calculates the sum of cells {', '.join(f'{c}{r}' for c in cols_found)}."
            return {
                "status": "success",
                "formula": formula,
                "suggested_header": "Total",
                "explanation": explanation,
                "type": "SUM_MATH"
            }

    # 3. AVERAGE
    if "average of" in p_lower or "calculate average" in p_lower:
        cols_found = []
        for word in p_lower.split():
            c = get_col(word)
            if c and c not in cols_found:
                cols_found.append(c)
        if cols_found:
            formula = f"=AVERAGE({cols_found[0]}{r}:{cols_found[-1]}{r})"
            explanation = f"Calculates the average across {cols_found[0]}{r} to {cols_found[-1]}{r}."
            return {
                "status": "success",
                "formula": formula,
                "suggested_header": "Average",
                "explanation": explanation,
                "type": "AVERAGE_MATH"
            }

    # 4. MULTIPLICATION / REVENUE (e.g. Price * Quantity)
    if any(k in p_lower for k in ["price * quantity", "multiply", "rate * hours", "revenue"]):
        c1 = get_col("price") or get_col("unit price") or get_col("rate") or "A"
        c2 = get_col("quantity") or get_col("qty") or get_col("hours") or "B"
        formula = f"={c1}{r}*{c2}{r}"
        explanation = f"Multiplies cell {c1}{r} by cell {c2}{r} to calculate total value."
        return {
            "status": "success",
            "formula": formula,
            "suggested_header": "Total Amount",
            "explanation": explanation,
            "type": "MULTIPLICATION"
        }

    # 5. XLOOKUP / VLOOKUP
    if any(k in p_lower for k in ["lookup", "vlookup", "xlookup", "find in sheet", "match from"]):
        key_col = get_col("email") or get_col("id") or "A"
        formula = f'=XLOOKUP({key_col}{r}, Sheet2!A:A, Sheet2!B:B, "Not Found")'
        explanation = f"Searches for value in {key_col}{r} inside Sheet2 column A and returns the corresponding value from Sheet2 column B."
        return {
            "status": "success",
            "formula": formula,
            "suggested_header": "Lookup Result",
            "explanation": explanation,
            "type": "LOOKUP"
        }

    # 6. COUNTIF / COUNTIFS
    if any(k in p_lower for k in ["count if", "count how many", "count passed"]):
        target_col = get_col("score") or get_col("grade") or "C"
        formula = f'=COUNTIF({target_col}2:{target_col}100, ">=60")'
        explanation = f"Counts the number of cells in range {target_col}2:{target_col}100 that meet the criteria '>=60'."
        return {
            "status": "success",
            "formula": formula,
            "suggested_header": "Count",
            "explanation": explanation,
            "type": "COUNTIF"
        }

    # Fallback generic formula generator
    # Attempt to extract column references
    first_col = column_index_to_letter(0)
    second_col = column_index_to_letter(1) if len(headers) > 1 else first_col
    formula = f"={first_col}{r}+{second_col}{r}"
    explanation = f"Combines values from column {first_col} and column {second_col} for row {r}."
    return {
        "status": "success",
        "formula": formula,
        "suggested_header": "Result",
        "explanation": explanation,
        "type": "GENERIC"
    }


def explain_formula(formula: str) -> Dict[str, Any]:
    """
    Parses and explains an Excel formula in plain, human-friendly English.
    """
    raw = formula.strip()
    if not raw.startswith("="):
        raw = "=" + raw
    
    f_upper = raw.upper()
    steps = []

    # Detect functions
    if "IFERROR(" in f_upper:
        steps.append("Shields the formula: if any error occurs, it returns a safe default value instead of #ERROR.")

    if "XLOOKUP(" in f_upper:
        match = re.search(r"XLOOKUP\s*\(([^,]+),\s*([^,]+),\s*([^,\)]+)", f_upper)
        if match:
            lookup_val, lookup_arr, return_arr = match.group(1), match.group(2), match.group(3)
            steps.append(f"Searches for value '{lookup_val}' inside range '{lookup_arr}' and returns the matching item from '{return_arr}'.")
        else:
            steps.append("Performs modern dynamic lookup across rows/columns.")

    elif "VLOOKUP(" in f_upper:
        match = re.search(r"VLOOKUP\s*\(([^,]+),\s*([^,]+),\s*([^,]+)", f_upper)
        if match:
            lookup_val, tbl, col_idx = match.group(1), match.group(2), match.group(3)
            steps.append(f"Looks up value '{lookup_val}' in the leftmost column of table '{tbl}' and retrieves value from column index {col_idx}.")

    if "IF(" in f_upper:
        match = re.search(r"IF\s*\(([^,]+),\s*([^,]+),\s*([^\)]+)\)", f_upper)
        if match:
            cond, if_true, if_false = match.group(1), match.group(2), match.group(3)
            steps.append(f"Evaluates condition [{cond}]: if TRUE, computes [{if_true}]; if FALSE, computes [{if_false}].")

    if "SUM(" in f_upper or "SUMIFS(" in f_upper:
        steps.append("Calculates the total mathematical sum of the specified cell ranges.")

    if "AVERAGE(" in f_upper or "AVERAGEIFS(" in f_upper:
        steps.append("Computes the arithmetic mean (average) of the referenced numbers.")

    if "COUNTIF(" in f_upper or "COUNTIFS(" in f_upper:
        steps.append("Counts cells within a range that meet the given criteria.")

    if not steps:
        steps.append(f"Computes mathematical expression: {raw[1:]}")

    return {
        "formula": raw,
        "summary": " ".join(steps),
        "breakdown": steps,
    }


def fix_formula_errors(
    broken_formula_or_cell: str,
    context_hint: str = ""
) -> Dict[str, Any]:
    """
    Diagnoses and generates fixes for Excel formula errors (#REF!, #VALUE!, #N/A, #DIV/0!, #NAME?).
    """
    s = broken_formula_or_cell.strip()
    s_upper = s.upper()

    error_type = "UNKNOWN"
    diagnosis = ""
    fixed_formula = ""

    if "#REF!" in s_upper:
        error_type = "#REF!"
        diagnosis = "Invalid Cell Reference: A cell, row, or column referred to in this formula was deleted or moved."
        # Replace #REF! with valid range or wrap in IFERROR
        fixed = re.sub(r"#REF!", "A2", s)
        fixed_formula = f"=IFERROR({fixed.lstrip('=')}, 0)" if fixed.startswith("=") else f"=IFERROR({fixed}, 0)"

    elif "#DIV/0!" in s_upper:
        error_type = "#DIV/0!"
        diagnosis = "Division by Zero: The denominator cell is zero or empty."
        fixed_formula = f"=IFERROR({s.lstrip('=')}, 0)"

    elif "#N/A" in s_upper:
        error_type = "#N/A"
        diagnosis = "Value Not Available: A lookup formula (XLOOKUP, VLOOKUP, MATCH) could not find the lookup value."
        fixed_formula = f'=IFNA({s.lstrip("=")}, "Not Found")'

    elif "#VALUE!" in s_upper:
        error_type = "#VALUE!"
        diagnosis = "Wrong Data Type: A mathematical formula expected a number but encountered text, or mismatched matrix dimensions."
        fixed_formula = f"=IFERROR({s.lstrip('=')}, 0)"

    elif "#NAME?" in s_upper:
        error_type = "#NAME?"
        diagnosis = "Unrecognized Name: A function name is misspelled or text wasn't enclosed in quotation marks."
        fixed_formula = s

    else:
        # Wrap any formula safely
        error_type = "GENERAL_ERROR"
        diagnosis = "Formula error or potential failure condition detected."
        fixed_formula = f"=IFERROR({s.lstrip('=')}, 0)"

    return {
        "error_type": error_type,
        "diagnosis": diagnosis,
        "original": s,
        "fixed_formula": fixed_formula,
        "recommendation": f"Wrap with IFERROR or check source cell values to prevent {error_type}."
    }

