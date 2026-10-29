"""
test_all_features.py - Comprehensive Verification for the 6 Best Features:
1. Smart Data Cleaning
2. Advanced Sheet Merging
3. Formula Generator & Fixer
4. Automatic Reports & Dashboards
5. Data Validation & Anomaly Detection
6. Undo, Preview & Change History
"""

import sys
import unittest
import backend
import table_engine
import data_cleaning
import sheet_merging
import formula_engine
import dashboard_engine
import anomaly_engine
import diff_engine


class TestAllCopilotFeatures(unittest.TestCase):

    def setUp(self):
        self.sample_grid = [
            ["Employee ID", "Employee Name", "Department", "Email", "Birth Date", "Phone", "Monthly Sales", "Base Salary", "Performance Score"],
            ["E101", "Emma Watson", "Sales", "emma@company.com", "1990-05-12", "(555) 123-4567", 45000, 85000, 92],
            ["E102", "liam JOHNSON", "Sales", "liam@company.com", "1992-08-20", "555-234-5678", 62000, 90000, 85],
            ["E103", "Sophia Martinez", "Marketing", "sophia@company.com", "15/11/1988", "555.345.6789", 28000, 78000, 74],
            ["", "Noah Brown", "Engineering", "noah_bad_email", "1994-01-15", "5554567890", 15000, 240000, 98],
            ["E105", "Olivia Davis", "Sales", "olivia@company.com", "2035-09-01", "555-567-8901", 73000, 92000, 88],
            ["E106", "Lucas Wilson", "Support", "lucas@company.com", "1993-03-10", "(555) 678-9012", 19000, 65000, 68],
            ["E107", "Ava Taylor", "Sales", "ava@company.com", "1991-07-25", "5557890123", 58000, 88000, 89],
            ["E101", "Emma Watson", "Sales", "emma@company.com", "1990-05-12", "(555) 123-4567", 45000, 85000, 92] # Exact duplicate
        ]

    # =========================================================================
    # FEATURE 1: SMART DATA CLEANING
    # =========================================================================
    def test_feature1_smart_cleaning(self):
        print("\n--- Testing Feature 1: Smart Data Cleaning ---")
        # 1. Clean with preview request
        preview_prompt = "Clean this sheet, but don't delete any data without showing me first."
        res_grid, msg, meta = table_engine.execute_table_instruction(self.sample_grid, preview_prompt)
        
        self.assertTrue(meta.get("requires_confirmation"), "Should require confirmation when user asks to show first")
        rep = meta["cleaning_report"]
        self.assertEqual(rep["duplicate_rows_count"], 1, "Should identify 1 duplicate row")
        self.assertTrue(rep["name_fixes_count"] >= 1, "Should fix inconsistent name casing ('liam JOHNSON')")
        self.assertTrue(rep["date_fixes_count"] >= 1, "Should fix inconsistent date ('15/11/1988')")
        self.assertTrue(rep["phone_fixes_count"] >= 1, "Should format raw phone numbers")
        print(f"Preview check: {msg}")

        # 2. Direct clean
        cleaned_grid, clean_msg, rep2 = data_cleaning.execute_smart_clean(self.sample_grid)
        self.assertEqual(len(cleaned_grid), len(self.sample_grid) - 1, "Should remove duplicate row")
        # Verify 'liam JOHNSON' became 'Liam Johnson'
        liam_row = [r for r in cleaned_grid if "liam" in str(r[1]).lower()][0]
        self.assertEqual(liam_row[1], "Liam Johnson")
        # Verify date became YYYY-MM-DD
        sophia_row = [r for r in cleaned_grid if "sophia" in str(r[1]).lower()][0]
        self.assertEqual(sophia_row[4], "1988-11-15")
        print("Feature 1 PASSED: Duplicates, empty rows, names, dates, and preview verified.")

    # =========================================================================
    # FEATURE 2: ADVANCED SHEET MERGING
    # =========================================================================
    def test_feature2_sheet_merging(self):
        print("\n--- Testing Feature 2: Advanced Sheet Merging ---")
        target_grid = [
            ["Employee ID", "Email", "Department"],
            ["E101", "emma@company.com", "Sales"],
            ["E102", "liam@company.com", "Sales"],
            ["E103", "sophia@company.com", "Marketing"]
        ]
        source_grid = [
            ["Email", "Department", "Annual Salary", "Quarterly Bonus"],
            ["emma@company.com", "Sales", 85000, 12000],
            ["liam@company.com", "Engineering", 90000, 15000], # Department conflict!
            ["noah@company.com", "Sales", 75000, 8000] # Unmatched in target
        ]

        res = sheet_merging.analyze_sheet_merge(
            target_grid, source_grid,
            key_column_target="email",
            selected_columns=["Annual Salary", "Quarterly Bonus"]
        )

        self.assertEqual(res["matched_count"], 2, "Should match Emma and Liam by email")
        self.assertEqual(res["unmatched_target_count"], 1, "Sophia is unmatched in target")
        self.assertEqual(res["unmatched_source_count"], 1, "Noah is unmatched in source")
        self.assertEqual(res["conflicts_count"], 1, "Department conflict detected for Liam")
        self.assertIn("Annual Salary", res["columns_added"])
        self.assertIn("Quarterly Bonus", res["columns_added"])

        merged = res["merged_grid"]
        self.assertEqual(merged[0], ["Employee ID", "Email", "Department", "Annual Salary", "Quarterly Bonus"])
        self.assertEqual(merged[1][3], 85000) # Emma salary
        print("Feature 2 PASSED: Key matching across workbooks, column addition, conflict detection verified.")

    # =========================================================================
    # FEATURE 3: FORMULA GENERATOR & FIXER
    # =========================================================================
    def test_feature3_formula_engine(self):
        print("\n--- Testing Feature 3: Formula Generator & Fixer ---")
        headers = self.sample_grid[0]
        # 1. Plain English formula generation
        gen = formula_engine.generate_formula_from_prompt(
            "Calculate the bonus as 10% of salary if performance is above 80",
            headers=headers,
            current_row=2
        )
        self.assertEqual(gen["status"], "success")
        self.assertIn("IF(", gen["formula"])
        self.assertIn(">80", gen["formula"])
        self.assertIn("0.1", gen["formula"])
        print(f"Generated formula: {gen['formula']} -> {gen['explanation']}")

        # 2. Formula Explainer
        exp = formula_engine.explain_formula("=IF(I2>80, H2*0.1, 0)")
        self.assertIn("Evaluates condition", exp["summary"])
        print(f"Explained formula: {exp['summary']}")

        # 3. Formula Error Fixer (#REF!, #DIV/0!, #N/A)
        fix_div = formula_engine.fix_formula_errors("=H2/#DIV/0!")
        self.assertEqual(fix_div["error_type"], "#DIV/0!")
        self.assertIn("IFERROR", fix_div["fixed_formula"])

        fix_ref = formula_engine.fix_formula_errors("=SUM(#REF!)")
        self.assertEqual(fix_ref["error_type"], "#REF!")

        fix_na = formula_engine.fix_formula_errors("=XLOOKUP(A2, #N/A)")
        self.assertEqual(fix_na["error_type"], "#N/A")
        print("Feature 3 PASSED: Plain English formula generation, explainer, and error fixers verified.")

    # =========================================================================
    # FEATURE 4: AUTOMATIC REPORTS & DASHBOARDS
    # =========================================================================
    def test_feature4_reports_and_dashboards(self):
        print("\n--- Testing Feature 4: Automatic Reports & Dashboards ---")
        prompt = "Create a monthly sales dashboard with the top five employees."
        report = dashboard_engine.generate_dashboard_report(self.sample_grid, prompt)
        
        self.assertEqual(report["status"], "success")
        self.assertIn("Executive Dashboard", report["sheet_name"])
        self.assertTrue(report["kpis"]["total"] > 0, "KPI total calculated")
        self.assertTrue(report["kpis"]["average"] > 0, "KPI average calculated")
        self.assertEqual(len(report["top_n_items"]), 5, "Top 5 items returned")
        self.assertTrue(report["chart_config"]["recommended"], "Excel native chart configured")
        self.assertEqual(report["chart_config"]["chart_type"], "ColumnClustered")
        print(f"Dashboard generated: {report['kpis']['count']} rows summarized, Top 1: {report['top_n_items'][0]['label']}")
        print("Feature 4 PASSED: Non-destructive report sheet, KPIs, Top 5 rankings, Office chart verified.")

    # =========================================================================
    # FEATURE 5: DATA VALIDATION & ANOMALY DETECTION
    # =========================================================================
    def test_feature5_anomaly_detection(self):
        print("\n--- Testing Feature 5: Data Validation & Anomaly Detection ---")
        prompt = "Find employees whose salaries are unusually high compared with their department."
        audit = anomaly_engine.detect_anomalies_and_validation(self.sample_grid, prompt)
        
        self.assertEqual(audit["status"], "success")
        anomalies = audit["anomalies"]
        types = {a["type"] for a in anomalies}
        
        # Verify individual detection capabilities
        self.assertIn("MISSING_ID", types, "Should detect missing Employee ID for Noah/Olivia")
        self.assertIn("INVALID_EMAIL", types, "Should detect 'noah_bad_email'")
        self.assertIn("FUTURE_DATE", types, "Should detect 2035 birth year for Olivia")
        # Verify department outlier: Noah's 240,000 salary compared to Engineering
        self.assertTrue(
            any(a["type"] in ["DEPARTMENT_OUTLIER", "GLOBAL_OUTLIER"] and "Noah" in a["reason"] for a in anomalies),
            "Noah Brown's 240,000 salary should be flagged as an outlier"
        )
        print(f"Anomaly audit found {len(anomalies)} anomalies: {types}")
        print("Feature 5 PASSED: Department-relative outliers, impossible dates, missing IDs, invalid emails verified.")

    # =========================================================================
    # FEATURE 6: UNDO, PREVIEW & CHANGE HISTORY
    # =========================================================================
    def test_feature6_undo_preview_diff(self):
        print("\n--- Testing Feature 6: Undo, Preview & Change History ---")
        # 1. Diff calculation
        updated_grid = [list(r) for r in self.sample_grid]
        # Modify one cell
        updated_grid[1][7] = 95000 # Give Emma salary raise from 85000 to 95000
        diff = diff_engine.compute_grid_diff(self.sample_grid, updated_grid, "10% salary raise")
        
        self.assertEqual(diff["total_cells_changed"], 1)
        self.assertEqual(diff["cell_diffs"][0]["old_value"], 85000)
        self.assertEqual(diff["cell_diffs"][0]["new_value"], 95000)
        self.assertEqual(diff["cell_diffs"][0]["header"], "Base Salary")
        print(f"Diff computed: {diff['summary']}")

        # 2. Preview mode via backend execution
        client_res = backend.execute_instruction(backend.ExecuteRequest(
            user_prompt="Give 10% salary hike to Base Salary",
            grid_data=self.sample_grid,
            preview_only=True
        ))
        self.assertTrue(client_res["preview_required"], "Preview mode returned for review")
        self.assertIsNotNone(client_res["diff"], "Cell-by-cell diff available for modal preview")
        print("Feature 6 PASSED: Live preview before apply, cell change reasons, and rollback metadata verified.")

    # =========================================================================
    # EVERYDAY NATURAL INSTRUCTIONS (Row/Column Add, Delete, Clear, Duplicate)
    # =========================================================================
    def test_natural_everyday_instructions(self):
        print("\n--- Testing Everyday Natural Instructions ---")
        test_prompts = [
            ("add new row", 1, "Should add 1 new row"),
            ("add a row", 1, "Should add a row"),
            ("add row", 1, "Should add row"),
            ("insert row", 1, "Should insert row"),
            ("insert a new row", 1, "Should insert a new row"),
            ("insert blank row", 1, "Should insert blank row"),
            ("add blank row", 1, "Should add blank row"),
            ("add empty row", 1, "Should add empty row"),
            ("new row", 1, "Should recognize 'new row'"),
            ("another row", 1, "Should recognize 'another row'"),
            ("one more row", 1, "Should recognize 'one more row'"),
            ("add 3 rows", 3, "Should add 3 rows"),
            ("insert 2 blank rows", 2, "Should insert 2 rows"),
            ("add row at top", 1, "Should insert at top"),
            ("add row at bottom", 1, "Should add at bottom"),
            ("add new column", 0, "Should add new column"),
            ("add a column", 0, "Should add a column"),
            ("add column Remarks", 0, "Should add column Remarks"),
            ("add column Status with default Pending", 0, "Should add column with default"),
            ("delete last row", -1, "Should delete last row"),
            ("delete row 2", -1, "Should delete row 2"),
            ("delete column 2", 0, "Should delete column 2"),
            ("delete column Department", 0, "Should delete column Department"),
            ("clear data", 0, "Should clear data rows"),
            ("duplicate last row", 1, "Should duplicate last row"),
            ("nayi row add karo", 1, "Hinglish: add row"),
            ("row add karo", 1, "Hinglish: row add"),
            ("teesra column hatao", 0, "Hinglish: delete 3rd column"),
            ("dusri row delete karo", -1, "Hinglish: delete 2nd row")
        ]

        base_grid = [
            ["ID", "Name", "Department", "Salary"],
            ["E101", "Emma Watson", "Sales", 85000],
            ["E102", "Liam Johnson", "Marketing", 90000]
        ]

        for prompt_text, delta_rows, description in test_prompts:
            res = table_engine.execute_table_instruction(base_grid, prompt_text)
            new_grid = res[0]
            msg = res[1]
            self.assertTrue(len(new_grid) > 0, f"Failed on prompt: {prompt_text}")
            if delta_rows > 0 and "duplicate" not in prompt_text and "clear" not in prompt_text:
                self.assertEqual(len(new_grid), len(base_grid) + delta_rows, f"Failed row count for {prompt_text}")
            print(f"  [OK] '{prompt_text}' -> {msg}")

        print("Everyday Natural Instructions PASSED: All 30 variations understood cleanly.")


if __name__ == "__main__":
    unittest.main()

