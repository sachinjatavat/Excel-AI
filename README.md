# 🎓 Local Excel AI Copilot for Teachers

> **100% Offline, Private, and Zero-Setup AI Assistant for Microsoft Excel.**  
> Powered by a dedicated **Built-in Excel Table Engine** with optional **Google Gemini / Ollama** support.

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

## 🎯 Teacher Prompt Cheat-Sheet

Here are real-world teacher prompts you can type or click directly from the sidebar shortcuts:

| Task | Example Prompt | What the Copilot Does |
| :--- | :--- | :--- |
| **Letter Grades** | `"Calculate letter grades (A, B, C, D, F) in a new column based on Total Score (90+ A, 80-89 B, 70-79 C, 60-69 D, below 60 F)."` | Adds a "Letter Grade" column and computes grades for every student. |
| **Grade Curving** | `"Add a 5-point curve to the Midterm column, ensuring no score exceeds 100."` | Updates test scores with the curve while capping at 100. |
| **Class Summary** | `"Add summary rows at the bottom with Average, Highest, and Lowest for all test columns."` | Calculates class statistics and appends them at the bottom. |
| **Student Remarks** | `"Add a 'Teacher Remarks' column with encouraging, personalized feedback for each student based on their score."` | Generates constructive, encouraging progress notes for report cards. |
| **Attendance %** | `"Calculate attendance percentage for each student in a new column based on Present days vs Total classes."` | Adds an attendance percentage column rounded to 1 decimal. |
| **Clean Student Names**| `"Standardize all student names to Proper Title Case and remove extra spaces."` | Cleans up messy rosters (e.g., `john DOE` $\rightarrow$ `John Doe`). |

---

## 🏗️ Project File Structure

```text
excel/
├── backend.py            # FastAPI backend (port 8000) communicating with Ollama (qwen2.5:7b)
├── taskpane.html         # Office.js HTML5/JS sidebar with Fluent UI & teacher shortcuts
├── manifest.xml          # Office Add-in XML manifest (v1.1) with ribbon button integration
├── requirements.txt      # Python dependencies (fastapi, uvicorn, requests, pydantic)
├── install.bat           # Admin batch script: registers Trusted Catalog in Windows Registry
├── start_assistant.bat   # Launcher script: starts backend, HTTP server, and launches Excel
├── assets/               # High-resolution add-in icons (16px, 32px, 64px, 80px)
└── README.md             # Complete user and setup documentation
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
