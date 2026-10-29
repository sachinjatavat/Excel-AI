import json
import re
import os
import logging
from typing import Any, List, Optional, Dict
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
import requests
import uvicorn

import table_engine
import data_cleaning
import sheet_merging
import formula_engine
import dashboard_engine
import anomaly_engine
import diff_engine

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("excel_ai_backend")

app = FastAPI(
    title="Local Excel AI Copilot",
    description="Dedicated Fast Offline Excel AI Engine with Smart Cleaning, Merging, Formulas, Dashboards, Anomaly Detection & Undo/Preview",
    version="3.0.0",
)

# Enable CORS for all origins
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ==============================================================================
# REQUEST MODELS
# ==============================================================================

class ExecuteRequest(BaseModel):
    user_prompt: str = Field(..., description="Natural language instruction or pasted data")
    grid_data: Optional[List[List[Any]]] = Field(default=[], description="2D array representing Excel worksheet rows and columns")
    secondary_grid: Optional[List[List[Any]]] = Field(default=None, description="Optional secondary sheet grid for merging")
    api_key: Optional[str] = Field(default=None, description="Optional Google Gemini API key")
    provider: Optional[str] = Field(default="auto", description="Execution engine: auto, builtin, or gemini")
    preview_only: Optional[bool] = Field(default=False, description="Whether to preview changes before applying")


class CleanAuditRequest(BaseModel):
    grid_data: List[List[Any]] = Field(..., description="2D table grid")
    remove_duplicates: bool = Field(default=True)
    remove_empty: bool = Field(default=True)
    standardize_names: bool = Field(default=True)
    standardize_dates: bool = Field(default=True)
    standardize_phones: bool = Field(default=True)
    clean_placeholders: bool = Field(default=True)


class MergeRequest(BaseModel):
    target_grid: List[List[Any]] = Field(..., description="Active sheet grid")
    source_grid: List[List[Any]] = Field(..., description="Secondary sheet or workbook grid")
    key_column_target: Optional[str] = Field(default=None)
    key_column_source: Optional[str] = Field(default=None)
    selected_columns: Optional[List[str]] = Field(default=None)
    overwrite_conflicts: bool = Field(default=False)


class FormulaGenRequest(BaseModel):
    prompt: str = Field(..., description="Plain English description of formula")
    headers: List[str] = Field(..., description="Worksheet column headers")
    current_row: int = Field(default=2)


class FormulaExplainRequest(BaseModel):
    formula: str = Field(..., description="Excel formula (e.g. =IF(C2>80, B2*0.1, 0))")


class FormulaFixRequest(BaseModel):
    broken_formula: str = Field(..., description="Formula or error cell (e.g. #REF!, #DIV/0!)")


class ValidateRequest(BaseModel):
    grid_data: List[List[Any]] = Field(..., description="2D table grid")
    prompt: Optional[str] = Field(default="", description="Optional focus e.g. department outliers")


class DashboardRequest(BaseModel):
    grid_data: List[List[Any]] = Field(..., description="2D table grid")
    prompt: Optional[str] = Field(default="", description="Optional dashboard parameters")


# ==============================================================================
# CLOUD AI HELPER (OPTIONAL GEMINI)
# ==============================================================================

def extract_and_parse_json(raw_text: str) -> dict:
    cleaned = raw_text.strip()
    cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned, flags=re.MULTILINE)
    cleaned = re.sub(r"\s*```$", "", cleaned, flags=re.MULTILINE)
    cleaned = cleaned.strip()

    try:
        data = json.loads(cleaned)
        if isinstance(data, dict):
            return data
    except Exception:
        pass

    match = re.search(r"(\{[\s\S]*\})", cleaned)
    if match:
        try:
            data = json.loads(match.group(1).strip())
            if isinstance(data, dict):
                return data
        except Exception:
            pass

    array_match = re.search(r"(\[[\s\S]*\])", cleaned)
    if array_match:
        try:
            arr = json.loads(array_match.group(1).strip())
            if isinstance(arr, list) and len(arr) > 0 and isinstance(arr[0], list):
                return {"status": "success", "updated_grid": arr}
        except Exception:
            pass

    raise ValueError(f"Unable to parse valid JSON from AI response. Raw output:\n{raw_text[:500]}")


def call_gemini(user_prompt: str, grid_data: List[List[Any]], api_key: str) -> dict:
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={api_key}"
    system_instruction = (
        "You are an expert Excel AI Copilot. Your task is to update or generate a 2D Excel table "
        "based on the user's instruction or pasted data.\n"
        "Guidelines:\n"
        "- Maintain rectangular table structure (all rows must have identical number of columns).\n"
        "- Add or remove columns/rows strictly as requested.\n"
        "- STRICT JSON RESPONSE FORMAT ONLY:\n"
        "{\n"
        '  "status": "success",\n'
        '  "updated_grid": [["Col1", "Col2"], ["val1", "val2"]]\n'
        "}\n"
        "Do not include any markdown fences or conversational text. Output raw JSON only."
    )
    prompt = (
        f"CURRENT 2D EXCEL GRID:\n{json.dumps(grid_data, ensure_ascii=False)}\n\n"
        f"INSTRUCTION / PASTED DATA:\n{user_prompt}\n\n"
        f"Return ONLY valid JSON with updated_grid."
    )
    body = {
        "contents": [{"parts": [{"text": system_instruction + "\n\n" + prompt}]}],
        "generationConfig": {
            "temperature": 0.1,
            "responseMimeType": "application/json",
        },
    }
    res = requests.post(url, json=body, timeout=30)
    if not res.ok:
        raise Exception(f"Gemini API returned HTTP {res.status_code}: {res.text}")
    res_data = res.json()
    candidates = res_data.get("candidates", [])
    if not candidates:
        raise Exception("No response generated by Gemini API.")
    raw_text = candidates[0]["content"]["parts"][0]["text"]
    parsed = extract_and_parse_json(raw_text)
    grid = table_engine.normalize_rectangular_grid(parsed.get("updated_grid", []))
    diff = diff_engine.compute_grid_diff(grid_data, grid)
    return {
        "status": "success",
        "updated_grid": grid,
        "rows_count": len(grid),
        "columns_count": len(grid[0]) if grid else 0,
        "message": "Processed successfully with Google Gemini AI.",
        "engine": "Google Gemini AI (Cloud)",
        "diff": diff,
    }


# ==============================================================================
# ENDPOINTS
# ==============================================================================

@app.get("/")
@app.get("/health")
def health_check():
    env_gemini_key = bool(os.environ.get("GEMINI_API_KEY"))
    return {
        "status": "healthy",
        "service": "Local Excel AI Copilot Backend",
        "version": "3.0.0",
        "features": [
            "Smart Data Cleaning (Duplicates, Empty Rows, Names, Dates, Phones)",
            "Advanced Sheet Merging (Key Matching, Conflicts, Unmatched)",
            "Formula Generator & Fixer (#REF!, #VALUE!, Plain English)",
            "Automatic Dashboards & Reports (KPIs, Top N, Office Charts)",
            "Data Validation & Anomaly Detection (Department Outliers, Dates, Emails)",
            "Undo, Live Preview & Change History"
        ],
        "builtin_engine": "Ready (100% Offline, Fast & Dedicated)",
        "gemini_configured": env_gemini_key,
    }


@app.post("/execute")
def execute_instruction(payload: ExecuteRequest):
    """
    Unified smart natural language execution endpoint.
    Handles all 6 core features + general Excel table transformations.
    """
    user_prompt = payload.user_prompt.strip()
    grid_data = payload.grid_data if isinstance(payload.grid_data, list) else []
    secondary_grid = payload.secondary_grid
    api_key = payload.api_key or os.environ.get("GEMINI_API_KEY")
    provider = (payload.provider or "auto").lower()
    user_preview_only = payload.preview_only or any(k in user_prompt.lower() for k in [
        "preview", "show me first", "show changes", "don't delete", "without showing", "show diff"
    ])

    if not user_prompt:
        raise HTTPException(status_code=400, detail="The instruction/data prompt cannot be empty.")

    # Convert None/null inside grid_data to empty strings
    sanitized_grid = []
    for row in grid_data:
        if isinstance(row, list):
            sanitized_grid.append(["" if cell is None else cell for cell in row])
        else:
            sanitized_grid.append([row])

    # If secondary grid provided for merge
    if secondary_grid and any(k in user_prompt.lower() for k in ["merge", "match", "join"]):
        merge_res = sheet_merging.analyze_sheet_merge(sanitized_grid, secondary_grid)
        if merge_res.get("status") == "success":
            merged = merge_res["merged_grid"]
            diff = diff_engine.compute_grid_diff(sanitized_grid, merged, "Sheet merge")
            msg = f"Merged {merge_res['matched_count']} records using key '{merge_res['key_column_target']}'. Added: {', '.join(merge_res['columns_added'])}."
            return {
                "status": "success",
                "updated_grid": merged,
                "rows_count": len(merged),
                "columns_count": len(merged[0]) if merged else 0,
                "message": msg,
                "engine": "Built-in Merge Engine",
                "diff": diff,
                "merge_report": merge_res,
                "preview_required": user_preview_only,
            }

    # 1. Explicit Gemini Cloud AI requested with API key
    if provider == "gemini" and api_key:
        try:
            return call_gemini(user_prompt, sanitized_grid, api_key)
        except Exception as e:
            logger.warning(f"Gemini API execution failed: {e}")
            raise HTTPException(status_code=502, detail=f"Gemini API error: {str(e)}")

    # 2. Execute via Built-in Excel AI Engine
    try:
        exec_res = table_engine.execute_table_instruction(sanitized_grid, user_prompt)
        metadata = {}
        if len(exec_res) == 3:
            updated_grid, action_msg, metadata = exec_res
        else:
            updated_grid, action_msg = exec_res

        # Compute exact diff
        diff = diff_engine.compute_grid_diff(sanitized_grid, updated_grid, action_msg)
        
        # Check if preview is required
        requires_conf = metadata.get("requires_confirmation", False) or user_preview_only

        return {
            "status": "success",
            "updated_grid": updated_grid,
            "rows_count": len(updated_grid),
            "columns_count": len(updated_grid[0]) if updated_grid else 0,
            "message": action_msg,
            "engine": "Built-in Excel AI Engine (Fast & Offline)",
            "diff": diff,
            "preview_required": requires_conf,
            "metadata": metadata,
            "action_type": metadata.get("action_type", "general"),
            "cleaning_report": metadata.get("cleaning_report"),
            "anomaly_report": metadata.get("anomaly_report"),
            "dashboard_report": metadata.get("dashboard_report"),
            "formula_info": metadata.get("formula_info"),
            "is_new_sheet": metadata.get("is_new_sheet", False),
            "sheet_name": metadata.get("sheet_name"),
        }
    except Exception as local_err:
        logger.warning(f"Execution notice: {local_err}")
        if api_key:
            try:
                return call_gemini(user_prompt, sanitized_grid, api_key)
            except Exception as gemini_err:
                logger.warning(f"Gemini fallback failed: {gemini_err}")

        raise HTTPException(status_code=400, detail=str(local_err))


# ------------------------------------------------------------------------------
# DEDICATED FEATURE ENDPOINTS
# ------------------------------------------------------------------------------

@app.post("/clean/audit")
def clean_audit(payload: CleanAuditRequest):
    """Audits sheet for cleaning issues and returns preview diff without deleting data."""
    report = data_cleaning.analyze_data_cleaning(
        payload.grid_data,
        remove_duplicates=payload.remove_duplicates,
        remove_empty=payload.remove_empty,
        standardize_names=payload.standardize_names,
        standardize_dates=payload.standardize_dates,
        standardize_phones=payload.standardize_phones,
        clean_placeholders=payload.clean_placeholders,
    )
    diff = diff_engine.compute_grid_diff(payload.grid_data, report["cleaned_grid"], "Smart data cleaning")
    return {
        "status": "success",
        "report": report,
        "diff": diff,
    }


@app.post("/clean/apply")
def clean_apply(payload: CleanAuditRequest):
    """Applies smart data cleaning to the grid."""
    report = data_cleaning.analyze_data_cleaning(
        payload.grid_data,
        remove_duplicates=payload.remove_duplicates,
        remove_empty=payload.remove_empty,
        standardize_names=payload.standardize_names,
        standardize_dates=payload.standardize_dates,
        standardize_phones=payload.standardize_phones,
        clean_placeholders=payload.clean_placeholders,
    )
    cleaned = report["cleaned_grid"]
    diff = diff_engine.compute_grid_diff(payload.grid_data, cleaned, "Applied data cleaning")
    return {
        "status": "success",
        "updated_grid": cleaned,
        "diff": diff,
        "message": f"Successfully resolved {report['total_issues_found']} cleaning issues.",
    }


@app.post("/merge/analyze")
def merge_analyze(payload: MergeRequest):
    """Analyzes merge between active sheet and second sheet/workbook."""
    res = sheet_merging.analyze_sheet_merge(
        payload.target_grid,
        payload.source_grid,
        key_column_target=payload.key_column_target,
        key_column_source=payload.key_column_source,
        selected_columns=payload.selected_columns,
        overwrite_conflicts=payload.overwrite_conflicts,
    )
    if "error" in res:
        raise HTTPException(status_code=400, detail=res["error"])
    diff = diff_engine.compute_grid_diff(payload.target_grid, res["merged_grid"], "Sheet merge")
    return {
        "status": "success",
        "merge_report": res,
        "diff": diff,
    }


@app.post("/formula/generate")
def formula_generate(payload: FormulaGenRequest):
    """Generates dynamic Excel formula from plain English."""
    res = formula_engine.generate_formula_from_prompt(
        payload.prompt, payload.headers, payload.current_row
    )
    return res


@app.post("/formula/explain")
def formula_explain(payload: FormulaExplainRequest):
    """Explains an Excel formula in plain English."""
    return formula_engine.explain_formula(payload.formula)


@app.post("/formula/fix")
def formula_fix(payload: FormulaFixRequest):
    """Diagnoses and fixes Excel formula errors (#REF!, #DIV/0!, #N/A, #VALUE!)."""
    return formula_engine.fix_formula_errors(payload.broken_formula)


@app.post("/report/generate")
def report_generate(payload: DashboardRequest):
    """Generates an executive dashboard with KPIs, Top N, and Chart configuration."""
    res = dashboard_engine.generate_dashboard_report(payload.grid_data, payload.prompt)
    if "error" in res:
        raise HTTPException(status_code=400, detail=res["error"])
    return res


@app.post("/validate/anomalies")
def validate_anomalies(payload: ValidateRequest):
    """Performs anomaly detection and data validation audit."""
    return anomaly_engine.detect_anomalies_and_validation(payload.grid_data, payload.prompt)


@app.post("/diff")
def calculate_diff(payload: Dict[str, Any]):
    """Calculates diff between two grids."""
    orig = payload.get("original_grid", [])
    upd = payload.get("updated_grid", [])
    reason = payload.get("reason", "")
    return diff_engine.compute_grid_diff(orig, upd, reason)


if __name__ == "__main__":
    print("=" * 65)
    print("      LOCAL EXCEL AI COPILOT - ADVANCED BACKEND v3.0")
    print("=" * 65)
    print(" * Engine:       Built-in Excel AI Engine (100% Offline, Fast)")
    print(" * Features:     Smart Clean | Merging | Formulas | Dashboards | Anomaly | Undo")
    if os.environ.get("GEMINI_API_KEY"):
        print(" * Cloud AI:     Google Gemini Configured (Optional)")
    print(" * Server URL:   http://127.0.0.1:8000")
    print("=" * 65)
    uvicorn.run(app, host="127.0.0.1", port=8000)
