"""
dashboard_engine.py - Automatic Reports and Dashboards Engine for Excel AI Copilot
Handles:
1. Creating a brand new report sheet without modifying the original data.
2. Calculating KPIs: Totals, Averages, Maximums, Minimums, and Record Counts.
3. Top N rankings (e.g. Top 5 Employees or Products).
4. Category/Department aggregations (pivot summary).
5. Chart specification for Excel native charts (Column/Bar/Line).
"""

import re
from typing import List, Dict, Any, Tuple, Optional
from collections import defaultdict


def is_num(val: Any) -> bool:
    if val is None or val == "":
        return False
    try:
        float(str(val).replace(",", "").replace("%", "").replace("$", "").replace("₹", "").strip())
        return True
    except ValueError:
        return False


def to_num(val: Any) -> float:
    return float(str(val).replace(",", "").replace("%", "").replace("$", "").replace("₹", "").strip())


def generate_dashboard_report(
    grid: List[List[Any]],
    prompt: str = ""
) -> Dict[str, Any]:
    """
    Analyzes active sheet data and compiles an executive Dashboard & Report
    designed to be placed in a separate, dedicated Excel sheet.
    """
    if not grid or len(grid) < 2:
        return {"error": "Sheet does not contain enough data to generate a dashboard."}

    headers = [str(c).strip() for c in grid[0]]
    data_rows = [r for r in grid[1:] if any(c not in [None, ""] for c in r)]
    
    if not data_rows:
        return {"error": "No data rows available."}

    num_rows = len(data_rows)
    num_cols = len(headers)

    # 1. Identify key columns: Label column (Name/Product/Item) and Metric column (Sales/Salary/Score/Revenue)
    metric_cols = []
    label_cols = []
    category_cols = []

    for c_idx, h in enumerate(headers):
        h_low = h.lower()
        # Test numeric ratio in data
        nums_count = sum(1 for r in data_rows if c_idx < len(r) and is_num(r[c_idx]))
        if nums_count / num_rows >= 0.5:
            # Check if this is an ID or Roll No
            if not any(k in h_low for k in ["id", "roll", "code", "s.no", "year", "phone"]):
                metric_cols.append(c_idx)
        else:
            if any(k in h_low for k in ["dept", "department", "category", "region", "team", "type", "role"]):
                category_cols.append(c_idx)
            elif not any(k in h_low for k in ["id", "roll", "code", "email", "phone", "date", "birth", "hire", "s.no"]):
                if any(k in h_low for k in ["name", "student", "employee", "product", "item", "title"]):
                    label_cols.append(c_idx)

    # Check if prompt specifically targets a metric (e.g. "sales dashboard" -> Monthly Sales)
    p_low = prompt.lower()
    primary_metric_idx = None
    for c_idx in metric_cols:
        h_low = headers[c_idx].lower()
        if ("sales" in p_low or "revenue" in p_low) and any(k in h_low for k in ["sales", "revenue"]):
            primary_metric_idx = c_idx
            break
        elif ("salary" in p_low or "compensation" in p_low) and any(k in h_low for k in ["salary", "pay"]):
            primary_metric_idx = c_idx
            break

    if primary_metric_idx is None:
        primary_metric_idx = metric_cols[0] if metric_cols else (num_cols - 1)
    primary_metric_name = headers[primary_metric_idx]

    primary_label_idx = label_cols[0] if label_cols else 0
    primary_label_name = headers[primary_label_idx]

    primary_cat_idx = category_cols[0] if category_cols else None

    # Check Top N requested in prompt (e.g. "top five", "top 5", "top 10")
    top_n = 5
    top_n_match = re.search(r"top\s+(\d+|five|three|ten)", prompt.lower())
    if top_n_match:
        n_str = top_n_match.group(1)
        if n_str.isdigit():
            top_n = int(n_str)
        elif n_str == "three": top_n = 3
        elif n_str == "five": top_n = 5
        elif n_str == "ten": top_n = 10

    # 2. Compute KPIs for Primary Metric
    metric_values = [to_num(r[primary_metric_idx]) for r in data_rows if primary_metric_idx < len(r) and is_num(r[primary_metric_idx])]
    total_val = round(sum(metric_values), 2) if metric_values else 0
    avg_val = round(sum(metric_values) / len(metric_values), 2) if metric_values else 0
    max_val = round(max(metric_values), 2) if metric_values else 0
    min_val = round(min(metric_values), 2) if metric_values else 0

    # 3. Compute Top N Leaderboard
    ranked_rows = []
    for r in data_rows:
        val = to_num(r[primary_metric_idx]) if primary_metric_idx < len(r) and is_num(r[primary_metric_idx]) else -999999
        lbl = str(r[primary_label_idx]) if primary_label_idx < len(r) else "Unknown"
        cat = str(r[primary_cat_idx]) if (primary_cat_idx is not None and primary_cat_idx < len(r)) else "-"
        ranked_rows.append((val, lbl, cat))

    ranked_rows.sort(key=lambda x: x[0], reverse=True)
    top_items = ranked_rows[:top_n]

    # 4. Compute Category / Department Aggregations
    cat_summary = {}
    if primary_cat_idx is not None:
        cat_groups = defaultdict(list)
        for r in data_rows:
            cat_name = str(r[primary_cat_idx]).strip() if primary_cat_idx < len(r) else "Other"
            if not cat_name: cat_name = "Other"
            if primary_metric_idx < len(r) and is_num(r[primary_metric_idx]):
                cat_groups[cat_name].append(to_num(r[primary_metric_idx]))
        
        for c_name, vals in cat_groups.items():
            cat_summary[c_name] = {
                "count": len(vals),
                "total": round(sum(vals), 2),
                "avg": round(sum(vals) / len(vals), 2),
                "pct_of_total": round((sum(vals) / total_val * 100), 1) if total_val > 0 else 0
            }

    # 5. Build New Sheet Layout 2D Grid
    sheet_grid = []
    
    # Title Banner
    sheet_grid.append(["EXECUTIVE DASHBOARD & REPORT", "", "", "", "", ""])
    sheet_grid.append([f"Auto-generated based on '{primary_metric_name}'", "", "", "", "", ""])
    sheet_grid.append(["", "", "", "", "", ""]) # blank row

    # KPI Summary Cards Table
    sheet_grid.append(["KEY METRIC", "VALUE", "", "STATISTIC", "VALUE", ""])
    sheet_grid.append(["Total " + primary_metric_name, total_val, "", "Active Records Count", num_rows, ""])
    sheet_grid.append(["Average " + primary_metric_name, avg_val, "", "Highest Recorded", max_val, ""])
    sheet_grid.append(["Lowest Recorded", min_val, "", "Spread (Max - Min)", round(max_val - min_val, 2), ""])
    sheet_grid.append(["", "", "", "", "", ""]) # blank row

    # Top N Leaderboard
    sheet_grid.append([f"TOP {top_n} RANKING: {primary_metric_name}", "", "", "", "", ""])
    sheet_grid.append(["Rank", primary_label_name, primary_metric_name, "Category/Dept", "% of Total", ""])
    for rank, (val, lbl, cat) in enumerate(top_items, start=1):
        pct = f"{round((val / total_val * 100), 1)}%" if total_val > 0 else "0%"
        sheet_grid.append([f"#{rank}", lbl, val, cat, pct, ""])

    sheet_grid.append(["", "", "", "", "", ""]) # blank row

    # Category / Department Table (if exists)
    if cat_summary:
        sheet_grid.append([f"BREAKDOWN BY {headers[primary_cat_idx].upper()}", "", "", "", "", ""])
        sheet_grid.append([headers[primary_cat_idx], "Headcount", "Total", "Average", "% Share", ""])
        for c_name, data in sorted(cat_summary.items(), key=lambda x: x[1]["total"], reverse=True):
            sheet_grid.append([c_name, data["count"], data["total"], data["avg"], f"{data['pct_of_total']}%", ""])

    # Pad rectangular structure
    max_w = max(len(r) for r in sheet_grid)
    for r in sheet_grid:
        if len(r) < max_w:
            r.extend([""] * (max_w - len(r)))

    # Chart recommendation
    chart_config = {
        "recommended": True,
        "chart_type": "ColumnClustered", # Office.js ChartType
        "title": f"Top {top_n} {primary_label_name} by {primary_metric_name}",
        # Target range coordinates for chart data: the Top N table
        "data_start_row": 9,
        "data_start_col": 1,
        "data_row_count": len(top_items) + 1,
        "data_col_count": 2,
    }

    return {
        "status": "success",
        "sheet_name": "Executive Dashboard",
        "dashboard_grid": sheet_grid,
        "kpis": {
            "total": total_val,
            "average": avg_val,
            "maximum": max_val,
            "minimum": min_val,
            "count": num_rows,
            "primary_metric": primary_metric_name,
            "primary_label": primary_label_name,
        },
        "top_n_items": [{"rank": idx+1, "label": lbl, "value": val, "category": cat} for idx, (val, lbl, cat) in enumerate(top_items)],
        "chart_config": chart_config,
    }
