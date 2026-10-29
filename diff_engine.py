"""
diff_engine.py - Grid Diff, Preview, and Change History Engine for Excel AI Copilot
Handles:
1. Calculating exact cell-by-cell diffs between original and updated Excel grids.
2. Generating human-readable change explanations ("Salary modified from 75000 to 82500").
3. Preview generation for user approval before applying changes to Excel.
4. History tracking and snapshot rollback structures.
"""

from typing import List, Dict, Any, Tuple, Optional


def column_letter(col_idx: int) -> str:
    """Converts 0-based column index to Excel column letter."""
    result = ""
    col_idx += 1
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def compute_grid_diff(
    original_grid: List[List[Any]],
    updated_grid: List[List[Any]],
    reason_hint: str = ""
) -> Dict[str, Any]:
    """
    Compares original_grid and updated_grid and returns a structured diff
    with cell coordinates, old/new values, and human-readable reasons.
    """
    if not original_grid:
        return {
            "is_new_table": True,
            "total_cells_changed": sum(len(r) for r in updated_grid),
            "rows_added": len(updated_grid),
            "rows_deleted": 0,
            "cols_added": len(updated_grid[0]) if updated_grid else 0,
            "cols_deleted": 0,
            "cell_diffs": [],
            "summary": f"Created new table with {len(updated_grid)} rows and {len(updated_grid[0]) if updated_grid else 0} columns.",
        }

    orig_rows = len(original_grid)
    orig_cols = len(original_grid[0]) if original_grid else 0
    new_rows = len(updated_grid)
    new_cols = len(updated_grid[0]) if updated_grid else 0

    headers = [str(c).strip() for c in (original_grid[0] if orig_rows > 0 else updated_grid[0])]

    cell_diffs = []
    total_changed = 0

    # Compare existing cells
    common_rows = min(orig_rows, new_rows)
    common_cols = min(orig_cols, new_cols)

    for r in range(common_rows):
        for c in range(common_cols):
            old_val = original_grid[r][c] if c < len(original_grid[r]) else ""
            new_val = updated_grid[r][c] if c < len(updated_grid[r]) else ""
            
            # Normalize for comparison
            s_old = "" if old_val is None else str(old_val).strip()
            s_new = "" if new_val is None else str(new_val).strip()

            if s_old != s_new:
                total_changed += 1
                col_let = column_letter(c)
                excel_cell = f"{col_let}{r+1}"
                header_name = headers[c] if c < len(headers) else f"Column {col_let}"
                
                # Formulate reason
                reason = reason_hint
                if not reason:
                    if r == 0:
                        reason = "Renamed column header"
                    elif s_old == "":
                        reason = "Populated previously empty cell"
                    elif s_new == "":
                        reason = "Cleared cell value"
                    else:
                        reason = f"Updated value in {header_name}"

                cell_diffs.append({
                    "cell": excel_cell,
                    "row": r,
                    "col": c,
                    "header": header_name,
                    "old_value": old_val,
                    "new_value": new_val,
                    "reason": reason,
                })

    # Added columns
    if new_cols > orig_cols:
        for c in range(orig_cols, new_cols):
            col_let = column_letter(c)
            new_header = updated_grid[0][c] if c < len(updated_grid[0]) else f"Column {col_let}"
            for r in range(new_rows):
                val = updated_grid[r][c] if c < len(updated_grid[r]) else ""
                total_changed += 1
                cell_diffs.append({
                    "cell": f"{col_let}{r+1}",
                    "row": r,
                    "col": c,
                    "header": new_header,
                    "old_value": "[New Column]",
                    "new_value": val,
                    "reason": f"Added column '{new_header}'",
                })

    # Added rows
    if new_rows > orig_rows:
        for r in range(orig_rows, new_rows):
            for c in range(new_cols):
                col_let = column_letter(c)
                val = updated_grid[r][c] if c < len(updated_grid[r]) else ""
                total_changed += 1
                header_name = headers[c] if c < len(headers) else f"Column {col_let}"
                cell_diffs.append({
                    "cell": f"{col_let}{r+1}",
                    "row": r,
                    "col": c,
                    "header": header_name,
                    "old_value": "[New Row]",
                    "new_value": val,
                    "reason": "Added new row record",
                })

    rows_diff = new_rows - orig_rows
    cols_diff = new_cols - orig_cols

    summary_parts = []
    if total_changed > 0:
        summary_parts.append(f"{total_changed} cells modified")
    if rows_diff > 0:
        summary_parts.append(f"{rows_diff} rows added")
    elif rows_diff < 0:
        summary_parts.append(f"{abs(rows_diff)} rows removed")
    if cols_diff > 0:
        summary_parts.append(f"{cols_diff} columns added")
    elif cols_diff < 0:
        summary_parts.append(f"{abs(cols_diff)} columns removed")

    summary = ", ".join(summary_parts) if summary_parts else "No changes detected."

    return {
        "is_new_table": False,
        "total_cells_changed": total_changed,
        "rows_added": max(0, rows_diff),
        "rows_deleted": max(0, -rows_diff),
        "cols_added": max(0, cols_diff),
        "cols_deleted": max(0, -cols_diff),
        "cell_diffs": cell_diffs[:50], # Limit to first 50 cells for fast preview rendering
        "total_diff_count": len(cell_diffs),
        "summary": summary,
    }

