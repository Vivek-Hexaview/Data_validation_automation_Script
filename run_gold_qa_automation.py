"""
========================================================================================
FABRIC GOLD DATA VALIDATION TEST RUNNER & EXCEL REPORT GENERATOR
========================================================================================
Description:
    Reads Silver -> Gold data validation test cases directly from:
    Gold_Data_Validation_Test_Cases.xlsx,
    connects to the Microsoft Fabric Warehouse (Gold_Warehouse) using your specified SQL endpoint,
    executes every validation query live, evaluates PASS / FAIL / ERROR status,
    and produces a comprehensive, color-coded execution report in Excel:
    Gold_QA_Execution_Results_Latest.xlsx (and timestamped report).

Fabric SQL Endpoint:
    ajgmg63gfu5uvbdmkc5o4riuke-tnixybw3f7kejjy4dnqgavvt64.datawarehouse.fabric.microsoft.com
========================================================================================
"""

import os
import sys
import time
import datetime
import pyodbc
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# CONFIGURATION
# ==============================================================================
BASE_DIR = r"c:\Users\Vivek.singh\OneDrive - Hexaview Technologies\Desktop\CR"
INPUT_WORKBOOK = os.path.join(BASE_DIR, "Gold_Data_Validation_Test_Cases.xlsx")

# Your specified Fabric Endpoint
SERVER = "ajgmg63gfu5uvbdmkc5o4riuke-tnixybw3f7kejjy4dnqgavvt64.datawarehouse.fabric.microsoft.com"
DATABASE = "Gold_Warehouse"

CONN_STR = (
    "Driver={ODBC Driver 18 for SQL Server};"
    f"Server={SERVER};"
    f"Database={DATABASE};"
    "Authentication=ActiveDirectoryInteractive;"
    "Encrypt=yes;"
    "TrustServerCertificate=no;"
)

# Output Paths
RUN_TIMESTAMP = datetime.datetime.now()
TIMESTAMP_STR = RUN_TIMESTAMP.strftime("%Y%m%d_%H%M%S")
LATEST_REPORT_PATH = os.path.join(BASE_DIR, "Gold_QA_Execution_Results_Latest.xlsx")
TIMESTAMPED_REPORT_PATH = os.path.join(BASE_DIR, f"Gold_QA_Execution_Results_{TIMESTAMP_STR}.xlsx")

# Excel Styling Palette
NAVY_HEADER = "0A2240"
DARK_TEAL   = "1F4E79"
WHITE       = "FFFFFF"
GRAY_BG     = "F9FAFB"
BORDER_GRAY = "D1D5DB"
FONT_CODE   = Font(name="Consolas", size=9)

STATUS_COLORS = {
    "PASSED":  {"bg": "E2EFDA", "fg": "276A3C", "badge": "PASSED"},
    "FAILED":  {"bg": "FCE4D6", "fg": "C00000", "badge": "FAILED"},
    "ERROR":   {"bg": "FFF2CC", "fg": "B25E00", "badge": "ERROR"},
    "SKIPPED": {"bg": "F1F5F9", "fg": "64748B", "badge": "SKIPPED"}
}

PRIORITY_COLORS = {
    "Critical": {"bg": "FEE2E2", "fg": "991B1B"},
    "High":     {"bg": "FFF2CC", "fg": "B25900"},
    "Medium":   {"bg": "E2EFDA", "fg": "375623"},
    "Low":      {"bg": "E0F2FE", "fg": "075985"}
}

BORDER_THIN = Border(
    left=Side(style="thin", color=BORDER_GRAY),
    right=Side(style="thin", color=BORDER_GRAY),
    top=Side(style="thin", color=BORDER_GRAY),
    bottom=Side(style="thin", color=BORDER_GRAY)
)

def evaluate_test_result(tc_id, expected_text, columns, rows):
    """
    Intelligently determines PASS or FAIL based on SQL query output and Silver->Gold test criteria.
    """
    if not rows:
        return "PASSED", "0 rows returned (Clean / Zero discrepancies)"

    col_map = {col.lower(): i for i, col in enumerate(columns)}
    
    # Check for failure keywords in columns across all rows (e.g. in UNION ALL queries)
    defect_keywords = [
        "mismatch", "duplicate", "orphan", "null_", "broken", "invalid", 
        "missing", "error", "unmapped", "wrong"
    ]

    total_defects = 0
    defect_details = []

    for r_idx, row in enumerate(rows):
        table_name = str(row[0]) if len(row) > 0 and isinstance(row[0], str) else f"Row {r_idx+1}"
        for col_name, idx in col_map.items():
            if any(dk in col_name for dk in defect_keywords):
                val = row[idx]
                try:
                    num = float(val) if val is not None else 0
                    if abs(num) > 0.0001:
                        total_defects += num
                        defect_details.append(f"{table_name}.{columns[idx]}={num:g}")
                except (ValueError, TypeError):
                    pass

    if total_defects > 0:
        detail_summary = ", ".join(defect_details[:4])
        if len(defect_details) > 4:
            detail_summary += f" (+{len(defect_details)-4} more)"
        return "FAILED", f"Discrepancies detected: {detail_summary}"

    # If rows exist but no defect columns had > 0
    # Check if query was a pure inspection query returning raw records
    exp_upper = (expected_text or "").upper()
    if ("ZERO" in exp_upper or "0 DUPLICATE" in exp_upper or "0 MISMATCH" in exp_upper) and len(rows) > 0 and len(columns) > 1 and not any(dk in col_map for dk in defect_keywords):
        return "FAILED", f"{len(rows)} defect rows returned directly"

    # Default to PASSED if all numeric check columns are 0
    first_row_preview = ", ".join([f"{c}: {rows[0][i]}" for i, c in enumerate(columns[:3])])
    return "PASSED", f"Clean: {first_row_preview}"

def run_tests():
    print("=" * 80)
    print("FABRIC LIVE GOLD DATA VALIDATION TEST RUNNER")
    print(f"Timestamp: {RUN_TIMESTAMP.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Server:    {SERVER}")
    print(f"Database:  {DATABASE}")
    print(f"Input:     {INPUT_WORKBOOK}")
    print("=" * 80)

    if not os.path.exists(INPUT_WORKBOOK):
        print(f"ERROR: Input workbook not found at: {INPUT_WORKBOOK}")
        sys.exit(1)

    # 1. Connect to Fabric
    print("\n[1/3] Connecting to Microsoft Fabric Gold_Warehouse via Azure Active Directory...")
    try:
        conn = pyodbc.connect(CONN_STR, timeout=30)
        cursor = conn.cursor()
        print(">>> SUCCESS: Connected to Microsoft Fabric Warehouse!")
    except Exception as e:
        print(f"\n>>> ERROR: Connection to Fabric failed: {e}")
        print("Please ensure you have active Fabric credentials and ODBC Driver 18 installed.")
        sys.exit(1)

    # 2. Read Test Cases from Workbook
    print("\n[2/3] Loading test cases from Gold_Data_Validation_Test_Cases.xlsx...")
    wb_in = openpyxl.load_workbook(INPUT_WORKBOOK, data_only=True)
    
    # We will execute:
    # 1. Table-level tests from sheets: 'Black Diamond Gold', 'Schwab Gold', 'Salentica Gold', 'Tamarac Gold', 'Innotrust Gold'
    # 2. Consolidated multi-object tests from sheet: 'Consolidated SK Audit'
    sheets_to_run = [
        "Black Diamond Gold", 
        "Schwab Gold", 
        "Salentica Gold", 
        "Tamarac Gold", 
        "Innotrust Gold",
        "Consolidated SK Audit"
    ]

    all_test_results = {}
    total_executed = 0
    total_passed = 0
    total_failed = 0
    total_error = 0

    for sheet_name in sheets_to_run:
        if sheet_name not in wb_in.sheetnames:
            continue
        ws = wb_in[sheet_name]
        print(f"\n--- Running Sheet: {sheet_name} ---")
        sheet_results = []

        is_consolidated = (sheet_name == "Consolidated SK Audit")

        for r in range(3, ws.max_row + 1):
            if is_consolidated:
                tc_id = str(ws.cell(r, 1).value or "").strip()
                source = str(ws.cell(r, 2).value or "").strip()
                scope = str(ws.cell(r, 3).value or "").strip()
                priority = str(ws.cell(r, 4).value or "Critical").strip()
                title = str(ws.cell(r, 5).value or "").strip()
                expected = str(ws.cell(r, 7).value or "").strip()
                query = str(ws.cell(r, 8).value or "").strip()
                category = "Consolidated Multi-Object Audit"
                table = scope
            else:
                tc_id = str(ws.cell(r, 1).value or "").strip()
                category = str(ws.cell(r, 2).value or "").strip()
                priority = str(ws.cell(r, 3).value or "High").strip()
                table = str(ws.cell(r, 4).value or "").strip()
                title = str(ws.cell(r, 5).value or "").strip()
                expected = str(ws.cell(r, 7).value or "").strip()
                query = str(ws.cell(r, 8).value or "").strip()

            if not tc_id or not query:
                continue

            # Execute Query
            start_t = time.time()
            status = "ERROR"
            output_summary = ""
            raw_output = ""

            try:
                cursor.execute(query)
                columns = [col[0] for col in cursor.description] if cursor.description else []
                rows = cursor.fetchall() if cursor.description else []
                exec_ms = int((time.time() - start_t) * 1000)

                # Format Raw Output preview
                if rows:
                    raw_lines = [f"Rows returned: {len(rows)} | Columns: {columns}"]
                    for row in rows[:5]:
                        raw_lines.append(str([val if val is not None else "NULL" for val in row]))
                    if len(rows) > 5:
                        raw_lines.append(f"... (+{len(rows)-5} more rows)")
                    raw_output = "\n".join(raw_lines)
                else:
                    raw_output = "0 rows returned."

                # Evaluate Pass/Fail
                status, output_summary = evaluate_test_result(tc_id, expected, columns, rows)

            except Exception as ex:
                exec_ms = int((time.time() - start_t) * 1000)
                status = "ERROR"
                output_summary = f"SQL Execution Error: {str(ex)[:150]}"
                raw_output = str(ex)

            total_executed += 1
            if status == "PASSED":
                total_passed += 1
                badge = "[PASS]"
            elif status == "FAILED":
                total_failed += 1
                badge = "[FAIL]"
            else:
                total_error += 1
                badge = "[ERR ]"

            print(f"  {badge} {tc_id:18} | {table[:25]:25} | {exec_ms:4}ms | {output_summary[:45]}")

            sheet_results.append({
                "id": tc_id,
                "category": category,
                "priority": priority,
                "table": table,
                "title": title,
                "expected": expected,
                "status": status,
                "output_summary": output_summary,
                "raw_output": raw_output,
                "exec_ms": exec_ms,
                "query": query
            })

        all_test_results[sheet_name] = sheet_results

    # 3. Generate Output Excel Report
    print("\n[3/3] Generating color-coded QA Execution Report...")
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    # -------------------------------------------------------------------------
    # Sheet 1: Executive Execution Summary
    # -------------------------------------------------------------------------
    ws_sum = wb_out.create_sheet(title="Execution Summary")
    ws_sum.views.sheetView[0].showGridLines = True

    ws_sum.merge_cells("A1:G1")
    ws_sum.cell(1, 1).value = "MICROSOFT FABRIC GOLD LAYER QA EXECUTION REPORT"
    ws_sum.cell(1, 1).font = Font(name="Calibri", size=16, bold=True, color=WHITE)
    ws_sum.cell(1, 1).fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    ws_sum.cell(1, 1).alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[1].height = 36

    ws_sum.merge_cells("A2:G2")
    ws_sum.cell(2, 1).value = f"Automated Live Execution on Gold_Warehouse | Generated: {RUN_TIMESTAMP.strftime('%Y-%m-%d %H:%M:%S')}"
    ws_sum.cell(2, 1).font = Font(name="Calibri", size=11, italic=True, color="D0E1FD")
    ws_sum.cell(2, 1).fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    ws_sum.cell(2, 1).alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[2].height = 24

    pass_rate = (total_passed / total_executed * 100) if total_executed > 0 else 0
    kpis = [
        ("TOTAL EXECUTED", total_executed, "A", "A", NAVY_HEADER),
        ("TOTAL PASSED", total_passed, "B", "B", "276A3C"),
        ("TOTAL FAILED", total_failed, "C", "C", "C00000"),
        ("TOTAL ERRORS", total_error, "D", "D", "B25E00"),
        ("PASS RATE", f"{pass_rate:.1f}%", "E", "G", "0E4D64" if pass_rate >= 90 else "C00000")
    ]

    ws_sum.row_dimensions[4].height = 20
    ws_sum.row_dimensions[5].height = 28
    for label, val, c1, c2, color in kpis:
        if c1 != c2:
            ws_sum.merge_cells(f"{c1}4:{c2}4")
            ws_sum.merge_cells(f"{c1}5:{c2}5")
        cell_lbl = ws_sum[f"{c1}4"]
        cell_val = ws_sum[f"{c1}5"]
        cell_lbl.value = label
        cell_lbl.font = Font(name="Calibri", size=9, bold=True, color="595959")
        cell_lbl.alignment = Alignment(horizontal="center", vertical="center")
        cell_lbl.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        cell_lbl.border = BORDER_THIN
        cell_val.value = val
        cell_val.font = Font(name="Calibri", size=16, bold=True, color=color)
        cell_val.alignment = Alignment(horizontal="center", vertical="center")
        cell_val.fill = PatternFill(start_color=WHITE, end_color=WHITE, fill_type="solid")
        cell_val.border = BORDER_THIN

    # Summary Table by Sheet
    headers_sum = ["Test Suite / Sheet", "Total Tests", "Passed", "Failed", "Errors", "Pass Rate (%)", "Status Badge"]
    ws_sum.row_dimensions[7].height = 26
    for col_idx, h in enumerate(headers_sum, 1):
        c = ws_sum.cell(7, col_idx, h)
        c.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
        c.fill = PatternFill(start_color=DARK_TEAL, end_color=DARK_TEAL, fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER_THIN

    for r_idx, (sname, sresults) in enumerate(all_test_results.items(), 8):
        ws_sum.row_dimensions[r_idx].height = 24
        fill = PatternFill(start_color=GRAY_BG, end_color=GRAY_BG, fill_type="solid") if r_idx % 2 == 0 else PatternFill(start_color=WHITE, end_color=WHITE, fill_type="solid")
        tot = len(sresults)
        p = sum(1 for x in sresults if x["status"] == "PASSED")
        f = sum(1 for x in sresults if x["status"] == "FAILED")
        e = sum(1 for x in sresults if x["status"] == "ERROR")
        pr = (p / tot * 100) if tot > 0 else 0
        badge = "PASSED" if f == 0 and e == 0 else ("FAILED" if f > 0 else "ERROR")

        row_vals = [sname, tot, p, f, e, f"{pr:.1f}%", badge]
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws_sum.cell(r_idx, c_idx, val)
            cell.border = BORDER_THIN
            cell.fill = fill
            if c_idx == 1:
                cell.font = Font(name="Calibri", size=10, bold=True)
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif c_idx == 7:
                st = STATUS_COLORS.get(badge, STATUS_COLORS["SKIPPED"])
                cell.font = Font(name="Calibri", size=10, bold=True, color=st["fg"])
                cell.fill = PatternFill(start_color=st["bg"], end_color=st["bg"], fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.font = Font(name="Calibri", size=10)
                cell.alignment = Alignment(horizontal="center", vertical="center")

    sum_widths = [32, 14, 14, 14, 14, 16, 18]
    for i, w in enumerate(sum_widths, 1):
        ws_sum.column_dimensions[get_column_letter(i)].width = w

    # -------------------------------------------------------------------------
    # Sheet 2: Master Consolidated Detailed Results
    # -------------------------------------------------------------------------
    ws_master = wb_out.create_sheet(title="All Results (Master)")
    ws_master.views.sheetView[0].showGridLines = True

    ws_master.merge_cells("A1:K1")
    ws_master.cell(1, 1).value = "CONSOLIDATED LIVE TEST EXECUTION RESULTS"
    ws_master.cell(1, 1).font = Font(name="Calibri", size=16, bold=True, color=WHITE)
    ws_master.cell(1, 1).fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    ws_master.cell(1, 1).alignment = Alignment(horizontal="center", vertical="center")
    ws_master.row_dimensions[1].height = 36

    m_headers = [
        "Suite / Source", "Test Case ID", "Priority", "Target Object / Scope", 
        "Test Title", "Status", "Execution Result Summary", "Latency", 
        "Raw Output Details", "Expected Result", "SQL Query"
    ]
    ws_master.row_dimensions[2].height = 26
    for col_idx, h in enumerate(m_headers, 1):
        c = ws_master.cell(2, col_idx, h)
        c.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
        c.fill = PatternFill(start_color=DARK_TEAL, end_color=DARK_TEAL, fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER_THIN

    cur_row = 3
    for sname, sresults in all_test_results.items():
        for tc in sresults:
            ws_master.row_dimensions[cur_row].height = 54
            fill = PatternFill(start_color=GRAY_BG, end_color=GRAY_BG, fill_type="solid") if cur_row % 2 == 0 else PatternFill(start_color=WHITE, end_color=WHITE, fill_type="solid")

            ws_master.cell(cur_row, 1, sname).alignment = Alignment(horizontal="center", vertical="center")
            ws_master.cell(cur_row, 2, tc["id"]).alignment = Alignment(horizontal="center", vertical="center")
            
            # Priority Badge
            p_cell = ws_master.cell(cur_row, 3, tc["priority"])
            p_cell.alignment = Alignment(horizontal="center", vertical="center")
            p_cfg = PRIORITY_COLORS.get(tc["priority"], PRIORITY_COLORS["Medium"])
            p_cell.font = Font(name="Calibri", size=10, bold=True, color=p_cfg["fg"])
            p_cell.fill = PatternFill(start_color=p_cfg["bg"], end_color=p_cfg["bg"], fill_type="solid")

            ws_master.cell(cur_row, 4, tc["table"]).alignment = Alignment(horizontal="left", vertical="top")
            ws_master.cell(cur_row, 5, tc["title"]).alignment = Alignment(horizontal="left", vertical="top")

            # Status Badge
            st_cell = ws_master.cell(cur_row, 6, tc["status"])
            st_cell.alignment = Alignment(horizontal="center", vertical="center")
            st_cfg = STATUS_COLORS.get(tc["status"], STATUS_COLORS["SKIPPED"])
            st_cell.font = Font(name="Calibri", size=10, bold=True, color=st_cfg["fg"])
            st_cell.fill = PatternFill(start_color=st_cfg["bg"], end_color=st_cfg["bg"], fill_type="solid")

            ws_master.cell(cur_row, 7, tc["output_summary"]).alignment = Alignment(horizontal="left", vertical="top")
            ws_master.cell(cur_row, 8, f"{tc['exec_ms']} ms").alignment = Alignment(horizontal="center", vertical="center")
            
            raw_c = ws_master.cell(cur_row, 9, tc["raw_output"])
            raw_c.font = FONT_CODE
            raw_c.alignment = Alignment(horizontal="left", vertical="top")

            ws_master.cell(cur_row, 10, tc["expected"]).alignment = Alignment(horizontal="left", vertical="top")
            
            q_c = ws_master.cell(cur_row, 11, tc["query"])
            q_c.font = FONT_CODE
            q_c.alignment = Alignment(horizontal="left", vertical="top")

            for c_i in range(1, 12):
                c = ws_master.cell(cur_row, c_i)
                c.border = BORDER_THIN
                if c_i not in [3, 6]:
                    c.fill = fill
                    if c_i in [1, 2, 4]:
                        c.font = Font(name="Calibri", size=10, bold=True)
                    elif c_i not in [9, 11]:
                        c.font = Font(name="Calibri", size=10)

            cur_row += 1

    m_widths = [22, 20, 12, 32, 36, 14, 40, 14, 50, 40, 65]
    for i, w in enumerate(m_widths, 1):
        ws_master.column_dimensions[get_column_letter(i)].width = w

    # -------------------------------------------------------------------------
    # Sheet 3-7: Dedicated Results Tabs per Suite
    # -------------------------------------------------------------------------
    for sname, sresults in all_test_results.items():
        ws_tab = wb_out.create_sheet(title=sname)
        ws_tab.views.sheetView[0].showGridLines = True

        ws_tab.merge_cells("A1:J1")
        ws_tab.cell(1, 1).value = f"{sname.upper()} - EXECUTION RESULTS"
        ws_tab.cell(1, 1).font = Font(name="Calibri", size=16, bold=True, color=WHITE)
        ws_tab.cell(1, 1).fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        ws_tab.cell(1, 1).alignment = Alignment(horizontal="center", vertical="center")
        ws_tab.row_dimensions[1].height = 36

        tab_h = [
            "Test Case ID", "Priority", "Target Object / Scope", "Test Title", 
            "Status", "Execution Result Summary", "Latency", "Raw Output Details", 
            "Expected Result", "SQL Query"
        ]
        ws_tab.row_dimensions[2].height = 26
        for col_idx, h in enumerate(tab_h, 1):
            c = ws_tab.cell(2, col_idx, h)
            c.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
            c.fill = PatternFill(start_color=DARK_TEAL, end_color=DARK_TEAL, fill_type="solid")
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = BORDER_THIN

        for r_idx, tc in enumerate(sresults, 3):
            ws_tab.row_dimensions[r_idx].height = 54
            fill = PatternFill(start_color=GRAY_BG, end_color=GRAY_BG, fill_type="solid") if r_idx % 2 == 0 else PatternFill(start_color=WHITE, end_color=WHITE, fill_type="solid")

            ws_tab.cell(r_idx, 1, tc["id"]).alignment = Alignment(horizontal="center", vertical="center")
            
            p_cell = ws_tab.cell(r_idx, 2, tc["priority"])
            p_cell.alignment = Alignment(horizontal="center", vertical="center")
            p_cfg = PRIORITY_COLORS.get(tc["priority"], PRIORITY_COLORS["Medium"])
            p_cell.font = Font(name="Calibri", size=10, bold=True, color=p_cfg["fg"])
            p_cell.fill = PatternFill(start_color=p_cfg["bg"], end_color=p_cfg["bg"], fill_type="solid")

            ws_tab.cell(r_idx, 3, tc["table"]).alignment = Alignment(horizontal="left", vertical="top")
            ws_tab.cell(r_idx, 4, tc["title"]).alignment = Alignment(horizontal="left", vertical="top")

            st_cell = ws_tab.cell(r_idx, 5, tc["status"])
            st_cell.alignment = Alignment(horizontal="center", vertical="center")
            st_cfg = STATUS_COLORS.get(tc["status"], STATUS_COLORS["SKIPPED"])
            st_cell.font = Font(name="Calibri", size=10, bold=True, color=st_cfg["fg"])
            st_cell.fill = PatternFill(start_color=st_cfg["bg"], end_color=st_cfg["bg"], fill_type="solid")

            ws_tab.cell(r_idx, 6, tc["output_summary"]).alignment = Alignment(horizontal="left", vertical="top")
            ws_tab.cell(r_idx, 7, f"{tc['exec_ms']} ms").alignment = Alignment(horizontal="center", vertical="center")
            
            raw_c = ws_tab.cell(r_idx, 8, tc["raw_output"])
            raw_c.font = FONT_CODE
            raw_c.alignment = Alignment(horizontal="left", vertical="top")

            ws_tab.cell(r_idx, 9, tc["expected"]).alignment = Alignment(horizontal="left", vertical="top")
            
            q_c = ws_tab.cell(r_idx, 10, tc["query"])
            q_c.font = FONT_CODE
            q_c.alignment = Alignment(horizontal="left", vertical="top")

            for c_i in range(1, 11):
                c = ws_tab.cell(r_idx, c_i)
                c.border = BORDER_THIN
                if c_i not in [2, 5]:
                    c.fill = fill
                    if c_i in [1, 3]:
                        c.font = Font(name="Calibri", size=10, bold=True)
                    elif c_i not in [8, 10]:
                        c.font = Font(name="Calibri", size=10)

        t_widths = [20, 12, 32, 36, 14, 40, 14, 50, 40, 65]
        for i, w in enumerate(t_widths, 1):
            ws_tab.column_dimensions[get_column_letter(i)].width = w

    # Save Workbooks
    wb_out.save(LATEST_REPORT_PATH)
    wb_out.save(TIMESTAMPED_REPORT_PATH)
    print(f"\n>>> SUCCESS: Latest Report Saved to: {LATEST_REPORT_PATH}")
    print(f">>> SUCCESS: Timestamped Report Saved to: {TIMESTAMPED_REPORT_PATH}")

    print("\n" + "=" * 80)
    print(f"EXECUTION SUMMARY: {total_executed} Tests | Passed: {total_passed} | Failed: {total_failed} | Errors: {total_error} | Pass Rate: {pass_rate:.1f}%")
    print("=" * 80)

if __name__ == "__main__":
    run_tests()
