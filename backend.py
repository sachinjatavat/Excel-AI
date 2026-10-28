import json
import re
import logging
from typing import Any, List
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
import requests
import uvicorn

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("excel_ai_backend")

app = FastAPI(
    title="Local Excel AI Copilot for Teachers",
    description="100% Offline Microsoft Excel AI Assistant powered by Ollama (qwen2.5:7b)",
    version="1.0.0",
)

# Enable CORS for all origins
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

OLLAMA_BASE_URL = "http://localhost:11434"
OLLAMA_GENERATE_URL = f"{OLLAMA_BASE_URL}/api/generate"
OLLAMA_TAGS_URL = f"{OLLAMA_BASE_URL}/api/tags"
DEFAULT_MODEL = "qwen2.5:7b"


def get_installed_models() -> List[str]:
    try:
        res = requests.get(OLLAMA_TAGS_URL, timeout=3)
        if res.status_code == 200:
            return [m.get("name") for m in res.json().get("models", []) if "name" in m]
    except Exception:
        pass
    return []


def resolve_active_model() -> str:
    installed = get_installed_models()
    if not installed:
        return DEFAULT_MODEL
    # Prefer qwen2.5:7b if installed
    for m in installed:
        if DEFAULT_MODEL in m or m.startswith("qwen2.5"):
            return m
    # Otherwise use the first installed model (e.g. gemma4:12b)
    return installed[0]


class ExecuteRequest(BaseModel):
    user_prompt: str = Field(..., description="Teacher's natural language instruction")
    grid_data: List[List[Any]] = Field(..., description="2D array representing Excel worksheet rows and columns")


def extract_and_parse_json(raw_text: str) -> dict:
    """
    Cleans markdown formatting, code fences, and extracts a valid JSON object
    from the LLM's raw response.
    """
    cleaned = raw_text.strip()

    # Remove standard markdown code fences ```json ... ``` or ``` ... ```
    cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned, flags=re.MULTILINE)
    cleaned = re.sub(r"\s*```$", "", cleaned, flags=re.MULTILINE)
    cleaned = cleaned.strip()

    # Try direct parse first
    try:
        data = json.loads(cleaned)
        if isinstance(data, dict):
            return data
    except Exception:
        pass

    # Find the outermost JSON block: from the first '{' to the last '}'
    match = re.search(r"(\{[\s\S]*\})", cleaned)
    if match:
        json_candidate = match.group(1).strip()
        try:
            data = json.loads(json_candidate)
            if isinstance(data, dict):
                return data
        except Exception as e:
            logger.warning(f"Failed regex-extracted JSON parse: {e}")

    # Fallback: Check if the model only returned a 2D array [[...]]
    array_match = re.search(r"(\[[\s\S]*\])", cleaned)
    if array_match:
        try:
            arr = json.loads(array_match.group(1).strip())
            if isinstance(arr, list) and len(arr) > 0 and isinstance(arr[0], list):
                return {"status": "success", "updated_grid": arr}
        except Exception:
            pass

    raise ValueError(f"Unable to parse valid JSON from LLM response. Raw output:\n{raw_text[:500]}")


@app.get("/")
@app.get("/health")
def health_check():
    """
    Health check endpoint that inspects local Ollama connectivity
    and identifies the active model.
    """
    ollama_online = False
    available_models = []

    try:
        res = requests.get(OLLAMA_TAGS_URL, timeout=3)
        if res.status_code == 200:
            ollama_online = True
            models_data = res.json().get("models", [])
            available_models = [m.get("name") for m in models_data if "name" in m]
    except Exception as e:
        logger.debug(f"Ollama health check probe failed: {e}")

    active_model = resolve_active_model()
    model_installed = len(available_models) > 0

    return {
        "status": "healthy",
        "service": "Local Excel AI Copilot Backend",
        "ollama_url": OLLAMA_BASE_URL,
        "ollama_online": ollama_online,
        "target_model": active_model,
        "model_installed": model_installed,
        "installed_models": available_models,
    }


@app.post("/execute")
def execute_instruction(payload: ExecuteRequest):
    """
    Accepts user natural language prompt and 2D grid data from Excel.
    Sends formatted prompt to local Ollama (qwen2.5:7b).
    Returns cleaned JSON with the updated 2D array.
    """
    user_prompt = payload.user_prompt.strip()
    grid_data = payload.grid_data

    if not user_prompt:
        raise HTTPException(status_code=400, detail="The 'user_prompt' cannot be empty.")

    if not isinstance(grid_data, list) or len(grid_data) == 0:
        raise HTTPException(
            status_code=400,
            detail="The 'grid_data' must be a non-empty 2D array representing Excel rows and columns.",
        )

    # Convert None/null inside grid_data to empty strings for cleaner processing
    sanitized_grid = []
    for row in grid_data:
        if isinstance(row, list):
            sanitized_grid.append(["" if cell is None else cell for cell in row])
        else:
            sanitized_grid.append([row])

    # Construct system and user prompt for Ollama
    system_prompt = (
        "You are an expert offline Excel AI Copilot and teacher's data assistant. "
        "Your task is to update or modify an Excel 2D grid based on the teacher's instruction.\n"
        "Guidelines:\n"
        "- Common tasks: Grade calculations (A/B/C/D/F), grade curves, statistics (Average, Median, Min, Max), "
        "attendance rates, personalized student feedback/remarks, sorting, formatting names.\n"
        "- Maintain rectangular integrity: every row in the output 2D grid must have the same number of columns.\n"
        "- If adding new columns (e.g. 'Grade' or 'Remarks'), include a header in the first row and values in subsequent rows.\n"
        "- Preserve all existing data unless specifically asked to edit or replace it.\n"
        "- For numerical values, format cleanly (e.g., round percentages/averages to 1 or 2 decimals).\n"
        "- STRICT OUTPUT FORMAT: Respond ONLY with a valid JSON object matching this schema:\n"
        "{\n"
        '  "status": "success",\n'
        '  "updated_grid": [[row1_col1, row1_col2, ...], [row2_col1, row2_col2, ...]]\n'
        "}\n"
        "- Do NOT output markdown ticks (no ```json or ```). Do NOT include conversational filler, explanations, or notes. "
        "Only output the raw JSON."
    )

    prompt = (
        f"INPUT 2D EXCEL GRID:\n"
        f"{json.dumps(sanitized_grid, ensure_ascii=False)}\n\n"
        f"TEACHER'S INSTRUCTION:\n"
        f"{user_prompt}\n\n"
        f"Return ONLY the JSON response with the updated grid."
    )

    logger.info(f"Received instruction: {user_prompt[:80]}... for grid of {len(sanitized_grid)} rows.")

    # Call Ollama API
    active_model = resolve_active_model()
    ollama_payload = {
        "model": active_model,
        "prompt": prompt,
        "system": system_prompt,
        "stream": False,
        "options": {
            "temperature": 0.1,  # Low temperature for strict adherence to structure and calculation accuracy
            "top_p": 0.9,
        },
    }

    try:
        response = requests.post(
            OLLAMA_GENERATE_URL,
            json=ollama_payload,
            timeout=120,  # 2 minute timeout for local model inference
        )
    except requests.exceptions.ConnectionError:
        error_msg = (
            f"Cannot connect to local Ollama at {OLLAMA_BASE_URL}. "
            f"Please ensure Ollama is installed and running."
        )
        logger.error(error_msg)
        raise HTTPException(status_code=503, detail=error_msg)
    except requests.exceptions.Timeout:
        error_msg = (
            f"Ollama took longer than 120 seconds to respond. "
            f"The sheet might be very large or your CPU/GPU is under heavy load."
        )
        logger.error(error_msg)
        raise HTTPException(status_code=504, detail=error_msg)
    except Exception as e:
        logger.error(f"Unexpected error communicating with Ollama: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Ollama communication error: {str(e)}")

    if response.status_code != 200:
        logger.error(f"Ollama returned HTTP {response.status_code}: {response.text}")
        raise HTTPException(
            status_code=502,
            detail=f"Ollama error (HTTP {response.status_code}): {response.text[:300]}",
        )

    try:
        ollama_json = response.json()
        raw_llm_output = ollama_json.get("response", "")
        logger.info(f"Raw response length: {len(raw_llm_output)} chars.")
    except Exception as e:
        raise HTTPException(
            status_code=500, detail=f"Failed to parse response envelope from Ollama: {str(e)}"
        )

    # Safely parse JSON from LLM output
    try:
        parsed_result = extract_and_parse_json(raw_llm_output)
    except ValueError as ve:
        logger.warning(f"JSON parsing error: {ve}")
        raise HTTPException(
            status_code=500,
            detail=f"The AI model did not return valid JSON. Error: {str(ve)}",
        )

    # Validate updated_grid structure
    updated_grid = parsed_result.get("updated_grid")
    if not isinstance(updated_grid, list) or len(updated_grid) == 0:
        raise HTTPException(
            status_code=500,
            detail="The AI model's output did not contain a valid 2D array in 'updated_grid'.",
        )

    # Normalize rectangular grid: Ensure every row has identical column length
    max_cols = max(len(r) if isinstance(r, list) else 1 for r in updated_grid)
    normalized_grid = []
    for r in updated_grid:
        if isinstance(r, list):
            row_copy = list(r)
            if len(row_copy) < max_cols:
                row_copy.extend([""] * (max_cols - len(row_copy)))
            normalized_grid.append(row_copy)
        else:
            normalized_grid.append([r] + [""] * (max_cols - 1))

    logger.info(
        f"Successfully processed grid. Returning {len(normalized_grid)} rows x {max_cols} columns."
    )

    return {
        "status": "success",
        "updated_grid": normalized_grid,
        "rows_count": len(normalized_grid),
        "columns_count": max_cols,
    }


if __name__ == "__main__":
    model = resolve_active_model()
    print(f"[*] Starting Excel AI Backend on http://127.0.0.1:8000 ...")
    print(f"[*] Active Ollama model: {model}")
    uvicorn.run(app, host="127.0.0.1", port=8000)
