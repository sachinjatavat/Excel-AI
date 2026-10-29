# 🎓 Local Excel AI Copilot for Teachers

> **100% Offline, Private, and Zero-Setup AI Assistant for Microsoft Excel.**  
> Powered by a dedicated **Built-in Excel AI Engine** with optional **Google Gemini** support.

---

## 🌟 Overview

The **Local Excel AI Copilot** is a taskpane add-in designed specifically for educators and teachers. It runs right inside Microsoft Excel's sidebar, parses raw pasted data into tables, executes natural language instructions, and updates your gradebook and student records automatically.

- **⚡ Zero Setup & No Ollama Required:** You do **not** need to download multi-gigabyte AI models or install external software. The built-in Excel engine runs out-of-the-box!
- **📋 Copy-Paste to Table:** Paste any raw unstructured text (e.g., `Rahul: 85, Priya: 92, Amit: 78` or CSV/lists) and instantly turn it into a formatted Excel table.
- **➕ Column & Row Management:** Add or remove columns on the fly (Letter Grades, Pass/Fail, Teacher Remarks, Attendance %, Rank, Totals, Averages).
- **🔒 100% Offline & Private:** Student grades, names, and academic records **never leave your computer**. Fully compliant with student data privacy guidelines (FERPA / GDPR).
- **🌐 Optional Cloud AI:** Supports Google Gemini API (free tier) for open-ended questions without downloading any local models.
- **✨ Native Excel Design:** Polished Microsoft Fluent UI sidebar that fits seamlessly into your everyday Excel workflow.

---

## 📋 System Requirements

1. **Operating System:** Windows 10 or Windows 11 (64-bit).
2. **Microsoft Excel:** Excel 2016, 2019, 2021, or Microsoft 365 desktop app (also works in Excel Online).
3. **Python:** Python 3.10 or higher ([Download from python.org](https://www.python.org/downloads/) — *make sure to check "Add python.exe to PATH"* during installation).

---

## 🚀 Quick Start Guide (3 Simple Steps)

### Step 1: Install Python Dependencies
Open Command Prompt in this project folder (`excel`) and run:
```bash
pip install -r requirements.txt
```
*This installs FastAPI, Uvicorn, Requests, and Pydantic (takes ~10 seconds).*

---

### Step 2: Run the One-Time Windows Installer
1. Find the file **`install.bat`** in this folder.
2. Simply **double-click** `install.bat` (administrator privileges are not required).
3. The script will automatically:
   - Copy `manifest.xml` to `%LOCALAPPDATA%\Microsoft\Office\AIAddIn\`.
   - Register the add-in in the Windows Registry (`WEF\Developer` and `WEF\TrustedCatalogs`).
4. You will see an **"INSTALLATION SUCCESSFUL!"** message. Press any key to close the window.

---

### Step 3: Start the Assistant
1. Simply **double-click** the file:
   ```text
   start_assistant.bat
   ```
2. This script automatically:
   - Starts the AI backend on `http://127.0.0.1:8000`.
   - Starts the local web server on `http://localhost:3000` to serve the sidebar.
   - Waits 2 seconds and automatically opens **Microsoft Excel**.

---

### Step 4: Open the Add-in Inside Excel
1. In Excel, open any workbook or create a new blank workbook.
2. If you don't see the **"Teacher Copilot"** button on the **Home** ribbon yet, insert it once:
   - **Modern Excel (Office 365 / 2024):**
     1. On the **Home** tab ribbon, click the **Add-ins** button (far right).
     2. Click **More Add-ins** (or browse add-ins).
     3. Click the **SHARED FOLDER** tab or **MY ADD-INS**.
     4. Select **Local Excel AI** and click **Add**.
   - **Classic Excel (2016 / 2019 / 2021):**
     1. Click the **Insert** tab on the top menu.
     2. Click **Add-ins** (or **My Add-ins**).
     3. Select the **SHARED FOLDER** tab.
     4. Select **Local Excel AI** and click **Add**.
   - **Excel Online (Browser):**
     1. Click **Insert** > **Office Add-ins** > **Upload My Add-in**.
     2. Browse to this folder and pick `manifest.xml`.
3. The **Teacher Copilot** button will now be on your **Home** ribbon, and the **Teacher AI Copilot** taskpane will appear on the right side of your spreadsheet!

> 💡 **First Time Testing?**  
> Click the **"📥 Load Demo Sheet"** button inside the sidebar. It will instantly populate the active sheet with a sample student gradebook so you can test prompts immediately!

---

## 🚀 6 Flagship Features (The Best Capabilities)

The Local Excel AI Copilot comes equipped with dedicated, offline engines for everyday data workflows:

### 1. 🧹 Smart Data Cleaning (High Priority)
- **Automatic Audit & Remediation:** Detects and removes exact duplicate rows and blank lines.
- **Inconsistent Name & Text Cleaning:** Converts messy casings (`john DOE`, `mCDONALD`, `ALICE`) to Proper Title Case (`John Doe`, `McDonald`, `Alice`) and removes redundant internal/trailing whitespace.
- **Date & Phone Standardization:** Formats mixed date strings (`15/11/1988`, `Nov 15 1988`, `1988.11.15`) into uniform ISO `YYYY-MM-DD`. Formats mixed phone numbers into standard `(XXX) XXX-XXXX` or international format.
- **Missing & Placeholder Values:** Flags or cleans `"N/A"`, `"null"`, `"-"`, `"None"` placeholders.
- **Approval / Preview Mode:** Full support for *"Clean this sheet, but don't delete any data without showing me first"* — presents a cell-by-cell preview with checkboxes before applying any modification.

### 2. 🔀 Advanced Sheet Merging
- **Multi-Workbook Matching:** Intelligently joins records across workbooks or sheets using **Email, Employee/Student ID, Phone Number, or Name**.
- **Selective Column Addition:** Import only the columns you want (e.g., *"Match employees by email and add their salaries from the second workbook"*).
- **Conflict & Mismatch Flags:** Automatically flags conflicting values (e.g. `Department: Sales vs Marketing`) and reports unmatched records from both source and target.

### 3. ⚡ Formula Generator, Explainer & Fixer
- **Plain English to Excel Formula:** Translates natural instructions into dynamic formulas mapped to your worksheet's actual column letters:
  - *"Calculate the bonus as 10% of salary if performance is above 80"* $\rightarrow$ `=IF(I2>80, H2*0.1, 0)`.
  - Supports `IF`, `SUMIFS`, `COUNTIF`, `AVERAGEIFS`, `XLOOKUP`, `VLOOKUP`, `DATEDIF`, `CONCATENATE`, and multi-condition logic.
- **Formula Explainer:** Translates cryptic nested Excel formulas into plain English step-by-step breakdowns.
- **Formula Error Fixer:** Automatically diagnoses `#REF!`, `#VALUE!`, `#N/A`, `#DIV/0!`, `#NAME?` and produces safe corrected formulas (e.g. wrapped in `=IFERROR(...)` or `=IFNA(...)`).

### 4. 📊 Automatic Reports & Executive Dashboards
- **100% Non-Destructive:** Creates a brand **new report sheet** (`Executive Dashboard`) leaving your raw original data completely untouched!
- **Executive KPIs:** Calculates Total, Average, Spread, Highest, Lowest, and Record Counts.
- **Top N Rankings:** Automatically ranks top performers (e.g. *"Create a monthly sales dashboard with the top five employees"*).
- **Native Excel Charts:** Generates formatted KPI summary cards, pivot tables, and inserts native Office.js Excel Clustered Column/Bar charts directly onto the dashboard sheet.

### 5. 🔍 Data Validation & Anomaly Detection
- **Department & Category Relative Outliers:** Identifies employees whose salaries or expenses deviate significantly from peers in the same department (e.g. *"Find employees whose salaries are unusually high compared with their department"*).
- **Chronology & Date Validation:** Detects impossible future birth dates, out-of-bounds years, and chronological violations.
- **Format Integrity:** Flags missing mandatory IDs, invalid email formats (missing `@` or domain), and abnormal negative numbers.
- **Visual Highlighting in Excel:** One-click button to highlight anomalous cells directly in Excel with soft warning colors (`#FEE2E2` / `#FEF3C7`).

### 6. ↩️ Undo, Live Preview & Change History
- **Interactive Live Preview Drawer:** Displays exact diffs before writing to the sheet: `Cell | Header | Before | After | Reason`.
- **Change Log:** Tracks every operation in the session with timestamps and cell counts.
- **One-Click Rollback / Undo:** Click **"↩️ Undo"** to instantly revert your Excel sheet back to its prior state.

---

## 🎯 Universal Prompt Cheat-Sheet (Any Domain & Language)

You can type prompts in **English, हिन्दी (Devanagari), or Hinglish / Romanized Hindi**:

| Category | Example Prompt | What the Copilot Does |
| :--- | :--- | :--- |
| **Add Row (Any Phrasing)** | `"add new row"`, `"insert row"`, `"add 3 rows"`, or `"row add karo"` | Adds clean new row(s) to the table with auto-incremented ID. |
| **Add Row with Data** | `"add employee Rahul with salary 75000 and department Engineering"` | Inserts new record and populates respective columns automatically. |
| **Add / Insert Column** | `"add new column Remarks"` or `"add column Status with default Pending"` | Appends column with optional default value across all data rows. |
| **Delete Row / Column** | `"delete last row"`, `"delete row 2"`, or `"delete column Bonus"` | Deletes target row or column cleanly by order number or name. |
| **Smart Cleaning** | `"Clean this sheet, but don't delete any data without showing me first."` | Audits sheet, finds duplicate rows, formats dates/names, and opens interactive preview for approval. |
| **Outlier Detection** | `"Find employees whose salaries are unusually high compared with their department."` | Calculates department-relative Z-scores and flags employees with abnormal compensation. |
| **Formula Generator** | `"Calculate the bonus as 10% of salary if performance is above 80."` | Generates `=IF(I2>80, H2*0.1, 0)` dynamically mapped to your column letters and inserts column. |
| **Formula Fixer** | `"Fix formula error #DIV/0!"` or `"=A2/#DIV/0!"` | Diagnoses root cause and wraps formula safely in `=IFERROR(A2/B2, 0)`. |
| **Executive Dashboard** | `"Create a monthly sales dashboard with the top five employees."` | Generates a new `Executive Dashboard` sheet with KPI cards, Top 5 ranking, and Excel chart. |
| **Sheet Merging** | `"Match employees by email and add their salaries from the second workbook."` | Matches across sheets using email key, imports salary column, and flags conflicting fields. |
| **Salary Hike & Bonuses** | `"Give 10% salary hike to Base Salary"` or `"salary me 10% hike do"` | Increases salary by 10% instantly across all employees. |

---

## 🏗️ Project File Structure

```text
excel/
├── backend.py            # FastAPI backend (port 8000) with endpoints for all 6 features
├── table_engine.py       # Table instruction engine connecting prompt parser to specialized modules
├── data_cleaning.py      # Feature 1: Smart Data Cleaning (duplicates, empty rows, names, dates, phones)
├── sheet_merging.py      # Feature 2: Advanced Sheet Merging (key matching, column import, conflict detection)
├── formula_engine.py     # Feature 3: Formula Generator, Explainer, and Fixer (#REF!, #DIV/0!, #N/A)
├── dashboard_engine.py   # Feature 4: Automatic Dashboards & Reports (KPIs, Top 5, Excel native charts)
├── anomaly_engine.py     # Feature 5: Anomaly Detection (department outliers, impossible dates, invalid emails)
├── diff_engine.py        # Feature 6: Grid Diff, Live Preview drawer, and Undo rollback snapshots
├── taskpane.html         # Office.js HTML5/JS sidebar with Fluent UI, dedicated tabs, and Undo bar
├── manifest.xml          # Office Add-in XML manifest (v1.1) with ribbon button integration
├── test_all_features.py  # Comprehensive test suite validating all 6 core features
├── requirements.txt      # Python dependencies (fastapi, uvicorn, requests, pydantic)
├── install.bat           # Setup batch script: registers Trusted Catalog in Windows Registry
├── start_assistant.bat   # Launcher script: starts backend, HTTP server, and launches Excel
└── assets/               # High-resolution add-in icons (16px, 32px, 64px, 80px)
```

---

## 🔧 Technical Details

### Architecture Flow
```
┌─────────────────────────┐          HTTP POST          ┌─────────────────────────┐
│   Microsoft Excel       │ ──────────────────────────> │   FastAPI Backend       │
│   (taskpane.html)       │                             │   (backend.py :8000)    │
│   - Reads getUsedRange  │ <────────────────────────── │   - Strips code fences  │
│   - Writes updated_grid │     Clean JSON Response     │   - Normalizes matrix   │
└─────────────────────────┘                             └────────────┬────────────┘
                                                                     │ HTTP POST
                                                                     │ (stream=False)
                                                                     ▼
                                                        ┌─────────────────────────┐
                                                        │   Ollama Local Server   │
                                                        │   (:11434 / qwen2.5:7b) │
                                                        └─────────────────────────┘
```

1. **`taskpane.html`:** Uses `Office.onReady` and `Excel.run(async (context) => ...)` to read `sheet.getUsedRange().values`.
2. **`backend.py`:** Formats a structured prompt with the 2D grid and teacher instruction.
3. **Ollama (`qwen2.5:7b`):** Runs locally on CPU/GPU, processes the table, and returns a JSON payload:
   ```json
   {
     "status": "success",
     "updated_grid": [["Student", "Score", "Grade"], ["Alice", 95, "A"]]
   }
   ```
4. **Excel Update:** The JavaScript layer measures the new grid dimensions and updates `sheet.getRangeByIndexes(...)`, automatically expanding new columns and auto-fitting column widths.

---

## ❓ Frequently Asked Questions & Troubleshooting

### Q1: The sidebar status says "Ollama: Offline".
- **Solution:** Ensure Ollama is running on your machine. Open a terminal and run `ollama serve` or `ollama run qwen2.5:7b`. You can also verify by visiting `http://localhost:11434` in your web browser (it should display *"Ollama is running"*).

### Q2: The sidebar status says "Backend: Offline".
- **Solution:** Make sure `start_assistant.bat` is running, or start the backend manually:
  ```bash
  python backend.py
  ```

### Q3: "Local Excel AI" does not show up in Excel under Insert > Add-ins > Shared Folder.
1. Make sure you ran `install.bat` by **right-clicking and selecting "Run as administrator"**.
2. Fully close Microsoft Excel and re-open it.
3. In Excel, go to: **File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs**.
4. Confirm that `%LOCALAPPDATA%\Microsoft\Office\AIAddIn` is listed and checked under **"Show in Menu"**.
5. Alternatively, in the Add-ins dialog, click the **Refresh** icon in the top-right corner.

### Q4: Can I use a different Ollama model (e.g., `llama3.1:8b` or `mistral`)?
- Yes! Open `backend.py` and change the line:
  ```python
  OLLAMA_MODEL = "qwen2.5:7b"
  ```
  to your preferred model (e.g., `llama3.1:8b`). Make sure you have pulled that model using `ollama pull <model_name>`.

---

## 🔒 Privacy & Compliance Guarantee

- **No telemetry or data tracking:** No requests are sent over the public internet.
- **Works completely offline:** You can disconnect your Wi-Fi or unplug your Ethernet cable, and the Copilot will continue functioning at 100% capacity.
- **Safe for student records:** Fully compliant with school policies regarding personally identifiable information (PII).
