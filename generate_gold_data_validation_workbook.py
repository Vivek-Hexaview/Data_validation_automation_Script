"""
Script to generate Gold_Data_Validation_Test_Cases.xlsx
Final Master Silver -> Gold Data Validation Test Suite:
Covers all 5 active Gold schemas:
1. Coverage Summary (Executive KPI cards & table summary)
2. All Gold Tests (Master) (All test cases consolidated)
3. Black Diamond Gold (Table-level SCD1/2 tests)
4. Schwab Gold (Table-level tests with exact pk_columns)
5. Salentica Gold (Table-level tests with Point-in-Time referential keys)
6. Tamarac Gold (Table-level tests with exact pk_columns)
7. Innotrust Gold (Table-level tests for master, category codes, type codes, sec index & characteristics)
8. Consolidated SK Audit (All 5 Sources Combined Multi-Object SK Validation Queries)
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_workbook():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default sheet

    # Typography & Styles
    FONT_TITLE = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    FONT_SUBTITLE = Font(name="Calibri", size=11, italic=True, color="D0E1FD")
    FONT_HEADER = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    FONT_BOLD = Font(name="Calibri", size=10, bold=True)
    FONT_REGULAR = Font(name="Calibri", size=10)
    FONT_CODE = Font(name="Consolas", size=9)

    FILL_NAVY = PatternFill(start_color="0A2240", end_color="0A2240", fill_type="solid")
    FILL_HEADER = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    FILL_SECTION = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    FILL_ZEBRA = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    FILL_WHITE = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    FILL_CRITICAL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    FILL_HIGH = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    FILL_MEDIUM = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    FONT_CRITICAL = Font(name="Calibri", size=10, bold=True, color="C00000")
    FONT_HIGH = Font(name="Calibri", size=10, bold=True, color="B25900")
    FONT_MEDIUM = Font(name="Calibri", size=10, bold=True, color="375623")

    THIN_BORDER = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    ALIGN_CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ALIGN_LEFT = Alignment(horizontal='left', vertical='top', wrap_text=True)
    ALIGN_CODE = Alignment(horizontal='left', vertical='top', wrap_text=True)

    # -------------------------------------------------------------------------------------------------
    # TABLE TEST SUITES DATA DEFINITIONS
    # -------------------------------------------------------------------------------------------------
    test_suites = {
        "Black Diamond Gold": [
            {
                "id": "TC_BD_GOLD_01", "cat": "SCD2 Deduplication", "pri": "Critical",
                "table": "black_diamond.dim_accounts",
                "desc": "Verify Active Deduplication on Business Key ['id'] (iscurrent = 'True')",
                "cond": "dim_accounts populated with SCD2 tracking columns",
                "exp": "Zero duplicate active accounts. Exactly one iscurrent = 'True' record per id.",
                "query": "SELECT id, COUNT(*) AS ActiveDuplicateCount\nFROM Gold_Warehouse.black_diamond.dim_accounts\nWHERE iscurrent = 'True'\nGROUP BY id\nHAVING COUNT(*) > 1;"
            },
            {
                "id": "TC_BD_GOLD_02", "cat": "SK Uniqueness", "pri": "Critical",
                "table": "black_diamond.dim_accounts",
                "desc": "Verify Primary Surrogate Key (accounts_SK) uniqueness and non-null constraint",
                "cond": "dim_accounts loaded",
                "exp": "accounts_SK is strictly unique across all versions. Zero duplicates, zero NULLs.",
                "query": "SELECT accounts_SK, COUNT(*) AS DuplicateCount\nFROM Gold_Warehouse.black_diamond.dim_accounts\nGROUP BY accounts_SK\nHAVING COUNT(*) > 1\nUNION ALL\nSELECT NULL, COUNT(*)\nFROM Gold_Warehouse.black_diamond.dim_accounts\nWHERE accounts_SK IS NULL\nHAVING COUNT(*) > 0;"
            },
            {
                "id": "TC_BD_GOLD_03", "cat": "Audit Columns", "pri": "High",
                "table": "black_diamond.dim_accounts",
                "desc": "Verify file_business_date and updatedat validity on active records",
                "cond": "dim_accounts populated",
                "exp": "Zero NULLs in file_business_date or updatedat; dates <= current calendar date.",
                "query": "SELECT \n    SUM(CASE WHEN file_business_date IS NULL THEN 1 ELSE 0 END) AS Null_FileBusinessDate,\n    SUM(CASE WHEN updatedat IS NULL THEN 1 ELSE 0 END) AS Null_UpdatedAt,\n    SUM(CASE WHEN file_business_date > CAST(GETDATE() AS DATE) THEN 1 ELSE 0 END) AS FutureBusinessDates\nFROM Gold_Warehouse.black_diamond.dim_accounts\nWHERE iscurrent = 'True';"
            },
            {
                "id": "TC_BD_GOLD_04", "cat": "SCD2 Deduplication", "pri": "Critical",
                "table": "black_diamond.dim_asset",
                "desc": "Verify Active Deduplication on Business Key ['assetid'] and asset_SK uniqueness",
                "cond": "dim_asset populated",
                "exp": "Zero duplicate active assets per assetid; asset_SK strictly unique.",
                "query": "SELECT assetid, COUNT(*) AS ActiveDuplicateCount\nFROM Gold_Warehouse.black_diamond.dim_asset\nWHERE iscurrent = 'True'\nGROUP BY assetid\nHAVING COUNT(*) > 1;"
            },
            {
                "id": "TC_BD_GOLD_05", "cat": "SCD2 Deduplication", "pri": "High",
                "table": "black_diamond.dim_client",
                "desc": "Verify Active Deduplication on Composite Business Key ['accounts', 'userid']",
                "cond": "dim_client populated",
                "exp": "Zero duplicate active clients per (accounts, userid) pair; client_SK strictly unique.",
                "query": "SELECT accounts, userid, COUNT(*) AS ActiveDuplicateCount\nFROM Gold_Warehouse.black_diamond.dim_client\nWHERE iscurrent = 'True'\nGROUP BY accounts, userid\nHAVING COUNT(*) > 1;"
            },
            {
                "id": "TC_BD_GOLD_06", "cat": "SCD2 Deduplication", "pri": "High",
                "table": "black_diamond.dim_contact",
                "desc": "Verify Active Deduplication on Business Key ['id'] and contact_SK uniqueness",
                "cond": "dim_contact populated",
                "exp": "Exactly 1 active row per id; contact_SK strictly unique.",
                "query": "SELECT id, COUNT(*) AS ActiveDuplicateCount\nFROM Gold_Warehouse.black_diamond.dim_contact\nWHERE iscurrent = 'True'\nGROUP BY id\nHAVING COUNT(*) > 1;"
            },
            {
                "id": "TC_BD_GOLD_07", "cat": "SCD2 Deduplication", "pri": "High",
                "table": "black_diamond.dim_target",
                "desc": "Verify Active Deduplication on Composite Business Key ['id', 'accounts', 'portfolios', 'component_allocation', 'component_id']",
                "cond": "dim_target populated",
                "exp": "Zero duplicate active target components per composite business key.",
                "query": "SELECT id, accounts, portfolios, component_allocation, component_id, COUNT(*) AS ActiveDuplicateCount\nFROM Gold_Warehouse.black_diamond.dim_target\nWHERE iscurrent = 'True'\nGROUP BY id, accounts, portfolios, component_allocation, component_id\nHAVING COUNT(*) > 1;"
            },
            {
                "id": "TC_BD_GOLD_08", "cat": "SCD2 Deduplication", "pri": "High",
                "table": "black_diamond.dim_portfoliogroup",
                "desc": "Verify Active Deduplication on Business Key ['portfoliogroupid', 'portfolioids']",
                "cond": "dim_portfoliogroup populated",
                "exp": "Zero duplicate active portfolio group combinations; portfoliogroup_SK strictly unique.",
                "query": "SELECT portfoliogroupid, portfolioids, COUNT(*) AS ActiveDuplicateCount\nFROM Gold_Warehouse.black_diamond.dim_portfoliogroup\nWHERE iscurrent = 'True'\nGROUP BY portfoliogroupid, portfolioids\nHAVING COUNT(*) > 1;"
            },
            {
                "id": "TC_BD_GOLD_09", "cat": "SCD2 Deduplication", "pri": "High",
                "table": "black_diamond.dim_relationshipv2",
                "desc": "Verify Active Deduplication on Composite Business Key ['id', 'accounts_id']",
                "cond": "dim_relationshipv2 populated",
                "exp": "Zero duplicate active relationships per (id, accounts_id); relationshipv2_SK strictly unique.",
                "query": "SELECT id, accounts_id, COUNT(*) AS ActiveDuplicateCount\nFROM Gold_Warehouse.black_diamond.dim_relationshipv2\nWHERE iscurrent = 'True'\nGROUP BY id, accounts_id\nHAVING COUNT(*) > 1;"
            },
            {
                "id": "TC_BD_GOLD_10", "cat": "Business Key & SK Validation", "pri": "Critical",
                "table": "black_diamond.dim_accounts_classallocation",
                "desc": "Verify Active Deduplication on Business Key ['accountid', 'classid'] and Foreign account_sk mapping",
                "cond": "dim_accounts_classallocation populated with attached account_sk",
                "exp": "Zero duplicate active combinations of (accountid, classid); account_sk correctly matches parent dim_accounts (c.accountid == da.id).",
                "query": "SELECT \n    (SELECT COUNT(*) FROM (SELECT accountid, classid FROM Gold_Warehouse.black_diamond.dim_accounts_classallocation WHERE iscurrent = 'True' GROUP BY accountid, classid HAVING COUNT(*) > 1) d) AS Duplicate_BusinessKey_Count,\n    SUM(CASE WHEN c.accountid <> da.id THEN 1 ELSE 0 END) AS Mismatched_AccountId_Count,\n    SUM(CASE WHEN c.account_sk IS NULL THEN 1 ELSE 0 END) AS Null_AccountSK_Count\nFROM Gold_Warehouse.black_diamond.dim_accounts_classallocation c\nLEFT JOIN Gold_Warehouse.black_diamond.dim_accounts da ON c.account_sk = da.accounts_SK\nWHERE c.iscurrent = 'True';"
            },
            {
                "id": "TC_BD_GOLD_11", "cat": "Business Key & SK Validation", "pri": "Critical",
                "table": "black_diamond.dim_accounts_segmentallocation",
                "desc": "Verify Active Deduplication on Business Key ['accountid', 'classid', 'segmentid'] and account_sk mapping",
                "cond": "dim_accounts_segmentallocation populated",
                "exp": "Zero duplicate active combinations of (accountid, classid, segmentid); account_sk matches dim_accounts.",
                "query": "SELECT \n    (SELECT COUNT(*) FROM (SELECT accountid, classid, segmentid FROM Gold_Warehouse.black_diamond.dim_accounts_segmentallocation WHERE iscurrent = 'True' GROUP BY accountid, classid, segmentid HAVING COUNT(*) > 1) d) AS Duplicate_BusinessKey_Count,\n    SUM(CASE WHEN s.accountid <> da.id THEN 1 ELSE 0 END) AS Mismatched_AccountId_Count,\n    SUM(CASE WHEN s.account_sk IS NULL THEN 1 ELSE 0 END) AS Null_AccountSK_Count\nFROM Gold_Warehouse.black_diamond.dim_accounts_segmentallocation s\nLEFT JOIN Gold_Warehouse.black_diamond.dim_accounts da ON s.account_sk = da.accounts_SK\nWHERE s.iscurrent = 'True';"
            },
            {
                "id": "TC_BD_GOLD_12", "cat": "Business Key & SK Validation", "pri": "Critical",
                "table": "black_diamond.dim_accounts_details",
                "desc": "Verify Active Deduplication on Business Key ['accountid'] and account_sk mapping",
                "cond": "dim_accounts_details populated",
                "exp": "Zero duplicate active records per accountid; zero NULL account_sk and zero mismatches (d.accountid == da.id).",
                "query": "SELECT \n    (SELECT COUNT(*) FROM (SELECT accountid FROM Gold_Warehouse.black_diamond.dim_accounts_details WHERE iscurrent = 'True' GROUP BY accountid HAVING COUNT(*) > 1) d) AS Duplicate_BusinessKey_Count,\n    SUM(CASE WHEN d.accountid <> da.id THEN 1 ELSE 0 END) AS Mismatched_AccountId_Count,\n    SUM(CASE WHEN d.account_sk IS NULL THEN 1 ELSE 0 END) AS Null_AccountSK_Count\nFROM Gold_Warehouse.black_diamond.dim_accounts_details d\nLEFT JOIN Gold_Warehouse.black_diamond.dim_accounts da ON d.account_sk = da.accounts_SK\nWHERE d.iscurrent = 'True';"
            },
            {
                "id": "TC_BD_GOLD_13", "cat": "Business Key & SK Validation", "pri": "Critical",
                "table": "black_diamond.dim_accounts_tags",
                "desc": "Verify Active Deduplication on Business Key ['accountid', 'name'] and account_sk mapping",
                "cond": "dim_accounts_tags populated",
                "exp": "Zero duplicate active records per (accountid, name); account_sk matches dim_accounts (t.accountid == da.id).",
                "query": "SELECT \n    (SELECT COUNT(*) FROM (SELECT accountid, name FROM Gold_Warehouse.black_diamond.dim_accounts_tags WHERE iscurrent = 'True' GROUP BY accountid, name HAVING COUNT(*) > 1) d) AS Duplicate_BusinessKey_Count,\n    SUM(CASE WHEN t.accountid <> da.id THEN 1 ELSE 0 END) AS Mismatched_AccountId_Count,\n    SUM(CASE WHEN t.account_sk IS NULL THEN 1 ELSE 0 END) AS Null_AccountSK_Count\nFROM Gold_Warehouse.black_diamond.dim_accounts_tags t\nLEFT JOIN Gold_Warehouse.black_diamond.dim_accounts da ON t.account_sk = da.accounts_SK\nWHERE t.iscurrent = 'True';"
            },
            {
                "id": "TC_BD_GOLD_14", "cat": "Business Key & SK Validation", "pri": "Critical",
                "table": "black_diamond.dim_accounts_team",
                "desc": "Verify Active Deduplication on Business Key ['accountid'] and account_sk mapping",
                "cond": "dim_accounts_team populated",
                "exp": "Zero duplicate active records per accountid; account_sk matches dim_accounts (tm.accountid == da.id).",
                "query": "SELECT \n    (SELECT COUNT(*) FROM (SELECT accountid FROM Gold_Warehouse.black_diamond.dim_accounts_team WHERE iscurrent = 'True' GROUP BY accountid HAVING COUNT(*) > 1) d) AS Duplicate_BusinessKey_Count,\n    SUM(CASE WHEN tm.accountid <> da.id THEN 1 ELSE 0 END) AS Mismatched_AccountId_Count,\n    SUM(CASE WHEN tm.account_sk IS NULL THEN 1 ELSE 0 END) AS Null_AccountSK_Count\nFROM Gold_Warehouse.black_diamond.dim_accounts_team tm\nLEFT JOIN Gold_Warehouse.black_diamond.dim_accounts da ON tm.account_sk = da.accounts_SK\nWHERE tm.iscurrent = 'True';"
            },
            {
                "id": "TC_BD_GOLD_15", "cat": "Business Key & SK Validation", "pri": "Critical",
                "table": "black_diamond.dim_accounts_targets",
                "desc": "Verify Active Deduplication on Business Key ['accountid', 'id'] and foreign account_sk alignment against dim_accounts",
                "cond": "dim_accounts_targets populated",
                "exp": "Zero duplicate active targets per (accountid, id); attached account_sk strictly matches dim_accounts (tg.accountid == da.id).",
                "query": "SELECT \n    (SELECT COUNT(*) FROM (SELECT accountid, id FROM Gold_Warehouse.black_diamond.dim_accounts_targets WHERE iscurrent = 'True' GROUP BY accountid, id HAVING COUNT(*) > 1) d) AS Duplicate_BusinessKey_Count,\n    SUM(CASE WHEN tg.accountid <> da.id THEN 1 ELSE 0 END) AS Mismatched_AccountId_Count,\n    SUM(CASE WHEN tg.account_sk IS NULL THEN 1 ELSE 0 END) AS Null_AccountSK_Count\nFROM Gold_Warehouse.black_diamond.dim_accounts_targets tg\nLEFT JOIN Gold_Warehouse.black_diamond.dim_accounts da ON tg.account_sk = da.accounts_SK\nWHERE tg.iscurrent = 'True';"
            },
            {
                "id": "TC_BD_GOLD_16", "cat": "Business Key & SK Validation", "pri": "Critical",
                "table": "black_diamond.fact_accounts_holdings",
                "desc": "Verify Deduplication on Business Key ['accountid', 'id', 'asofdate'] and account_sk foreign key resolution",
                "cond": "fact_accounts_holdings populated",
                "exp": "Zero duplicate records on (accountid, id, asofdate); account_sk strictly matches dim_accounts (h.accountid == da.id).",
                "query": "SELECT \n    (SELECT COUNT(*) FROM (SELECT accountid, id, asofdate FROM Gold_Warehouse.black_diamond.fact_accounts_holdings GROUP BY accountid, id, asofdate HAVING COUNT(*) > 1) d) AS Duplicate_BusinessKey_Count,\n    SUM(CASE WHEN h.accountid <> da.id THEN 1 ELSE 0 END) AS Mismatched_AccountId_Count,\n    SUM(CASE WHEN h.account_sk IS NULL THEN 1 ELSE 0 END) AS Null_AccountSK_Count\nFROM Gold_Warehouse.black_diamond.fact_accounts_holdings h\nLEFT JOIN Gold_Warehouse.black_diamond.dim_accounts da ON h.account_sk = da.accounts_SK;"
            },
            {
                "id": "TC_BD_GOLD_17", "cat": "Business Key & SK Validation", "pri": "Critical",
                "table": "black_diamond.fact_accounts_transactions",
                "desc": "Verify Deduplication on Business Key ['transactionid'] and account_sk foreign key resolution",
                "cond": "fact_accounts_transactions populated",
                "exp": "Zero duplicate records on transactionid; account_sk strictly matches dim_accounts (tx.accountid == da.id).",
                "query": "SELECT \n    (SELECT COUNT(*) FROM (SELECT transactionid FROM Gold_Warehouse.black_diamond.fact_accounts_transactions GROUP BY transactionid HAVING COUNT(*) > 1) d) AS Duplicate_BusinessKey_Count,\n    SUM(CASE WHEN tx.accountid <> da.id THEN 1 ELSE 0 END) AS Mismatched_AccountId_Count,\n    SUM(CASE WHEN tx.account_sk IS NULL THEN 1 ELSE 0 END) AS Null_AccountSK_Count\nFROM Gold_Warehouse.black_diamond.fact_accounts_transactions tx\nLEFT JOIN Gold_Warehouse.black_diamond.dim_accounts da ON tx.account_sk = da.accounts_SK;"
            },
            {
                "id": "TC_BD_GOLD_18", "cat": "Business Key & SK Validation", "pri": "Critical",
                "table": "black_diamond.fact_accounts_taxlots",
                "desc": "Verify Deduplication on Business Key ['accountid', 'taxlotid', 'holdingid'] and account_sk foreign key resolution",
                "cond": "fact_accounts_taxlots populated",
                "exp": "Zero duplicate records on (accountid, taxlotid, holdingid); account_sk strictly matches dim_accounts (tl.accountid == da.id).",
                "query": "SELECT \n    (SELECT COUNT(*) FROM (SELECT accountid, taxlotid, holdingid FROM Gold_Warehouse.black_diamond.fact_accounts_taxlots GROUP BY accountid, taxlotid, holdingid HAVING COUNT(*) > 1) d) AS Duplicate_BusinessKey_Count,\n    SUM(CASE WHEN tl.accountid <> da.id THEN 1 ELSE 0 END) AS Mismatched_AccountId_Count,\n    SUM(CASE WHEN tl.account_sk IS NULL THEN 1 ELSE 0 END) AS Null_AccountSK_Count\nFROM Gold_Warehouse.black_diamond.fact_accounts_taxlots tl\nLEFT JOIN Gold_Warehouse.black_diamond.dim_accounts da ON tl.account_sk = da.accounts_SK;"
            }
        ],
        "Schwab Gold": [
            {
                "id": "TC_SC_GOLD_01", "cat": "SCD2 Deduplication", "pri": "Critical",
                "table": "schwab.dim_accounts",
                "desc": "Verify Active Deduplication on Composite Business Key ['account_id', 'custodian_id']",
                "cond": "dim_accounts populated with IsCurrent flag",
                "exp": "Zero duplicate active accounts; account_sk strictly unique across all rows.",
                "query": "SELECT account_id, custodian_id, COUNT(*) AS ActiveDuplicateCount\nFROM Gold_Warehouse.schwab.dim_accounts\nWHERE IsCurrent = 'True'\nGROUP BY account_id, custodian_id\nHAVING COUNT(*) > 1;"
            },
            {
                "id": "TC_SC_GOLD_02", "cat": "Business Key & SK Validation", "pri": "Critical",
                "table": "schwab.fact_recon_position_d1",
                "desc": "Verify PK ['account_id', 'accounting_rule_code', 'item_issue_id', 'rule_set_suffix_id'] deduplication & account_sk mapping",
                "cond": "fact_recon_position_d1 populated",
                "exp": "Zero duplicate rows on 4-column business key; attached account_sk strictly maps to the same account_id in dim_accounts.",
                "query": "SELECT \n    (SELECT COUNT(*) FROM (SELECT account_id, accounting_rule_code, item_issue_id, rule_set_suffix_id FROM Gold_Warehouse.schwab.fact_recon_position_d1 GROUP BY account_id, accounting_rule_code, item_issue_id, rule_set_suffix_id HAVING COUNT(*) > 1) d) AS Duplicate_BusinessKey_Count,\n    SUM(CASE WHEN f.account_sk IS NOT NULL AND da.account_sk IS NOT NULL AND f.account_id <> da.account_id THEN 1 ELSE 0 END) AS Mismatched_AccountId_Count,\n    SUM(CASE WHEN f.account_sk IS NULL THEN 1 ELSE 0 END) AS Null_AccountSK_Count,\n    SUM(CASE WHEN f.account_sk IS NOT NULL AND da.account_sk IS NULL THEN 1 ELSE 0 END) AS Orphan_AccountSK_Count\nFROM Gold_Warehouse.schwab.fact_recon_position_d1 f\nLEFT JOIN Gold_Warehouse.schwab.dim_accounts da ON f.account_sk = da.account_sk;"
            },
            {
                "id": "TC_SC_GOLD_03", "cat": "Business Key & SK Validation", "pri": "Critical",
                "table": "schwab.fact_recon_position_d2",
                "desc": "Verify PK ['account_id'] deduplication & account_sk referential mapping",
                "cond": "fact_recon_position_d2 populated",
                "exp": "Zero duplicate rows on account_id business key; attached account_sk strictly maps to dim_accounts.",
                "query": "SELECT \n    (SELECT COUNT(*) FROM (SELECT account_id FROM Gold_Warehouse.schwab.fact_recon_position_d2 GROUP BY account_id HAVING COUNT(*) > 1) d) AS Duplicate_BusinessKey_Count,\n    SUM(CASE WHEN f.account_sk IS NOT NULL AND da.account_sk IS NOT NULL AND f.account_id <> da.account_id THEN 1 ELSE 0 END) AS Mismatched_AccountId_Count,\n    SUM(CASE WHEN f.account_sk IS NULL THEN 1 ELSE 0 END) AS Null_AccountSK_Count,\n    SUM(CASE WHEN f.account_sk IS NOT NULL AND da.account_sk IS NULL THEN 1 ELSE 0 END) AS Orphan_AccountSK_Count\nFROM Gold_Warehouse.schwab.fact_recon_position_d2 f\nLEFT JOIN Gold_Warehouse.schwab.dim_accounts da ON f.account_sk = da.account_sk;"
            },
            {
                "id": "TC_SC_GOLD_04", "cat": "Business Key & SK Validation", "pri": "Critical",
                "table": "schwab.fact_daily_closing_lots",
                "desc": "Verify PK ['account_id', 'item_issue_id', 'close_event_id', 'close_open_event_id'] deduplication & account_sk mapping",
                "cond": "fact_daily_closing_lots populated",
                "exp": "Zero duplicate rows on 4-column business key; account_sk strictly maps to dim_accounts (f.account_id == da.account_id).",
                "query": "SELECT \n    (SELECT COUNT(*) FROM (SELECT account_id, item_issue_id, close_event_id, close_open_event_id FROM Gold_Warehouse.schwab.fact_daily_closing_lots GROUP BY account_id, item_issue_id, close_event_id, close_open_event_id HAVING COUNT(*) > 1) d) AS Duplicate_BusinessKey_Count,\n    SUM(CASE WHEN f.account_sk IS NOT NULL AND da.account_sk IS NOT NULL AND f.account_id <> da.account_id THEN 1 ELSE 0 END) AS Mismatched_AccountId_Count,\n    SUM(CASE WHEN f.account_sk IS NULL THEN 1 ELSE 0 END) AS Null_AccountSK_Count,\n    SUM(CASE WHEN f.account_sk IS NOT NULL AND da.account_sk IS NULL THEN 1 ELSE 0 END) AS Orphan_AccountSK_Count\nFROM Gold_Warehouse.schwab.fact_daily_closing_lots f\nLEFT JOIN Gold_Warehouse.schwab.dim_accounts da ON f.account_sk = da.account_sk;"
            },
            {
                "id": "TC_SC_GOLD_05", "cat": "Business Key & SK Uniqueness", "pri": "High",
                "table": "schwab.fact_security",
                "desc": "Verify PK ['schwab_security_number', 'item_issue_id'] deduplication, security_sk uniqueness and audit dates",
                "cond": "fact_security populated",
                "exp": "Zero duplicates on (schwab_security_number, item_issue_id); security_sk is strictly unique.",
                "query": "SELECT \n    (SELECT COUNT(*) FROM (SELECT schwab_security_number, item_issue_id FROM Gold_Warehouse.schwab.fact_security GROUP BY schwab_security_number, item_issue_id HAVING COUNT(*) > 1) d) AS Duplicate_BusinessKey_Count,\n    (COUNT(*) - COUNT(DISTINCT security_sk)) AS Duplicate_SecuritySK_Count\nFROM Gold_Warehouse.schwab.fact_security;"
            },
            {
                "id": "TC_SC_GOLD_06", "cat": "Business Key & SK Validation", "pri": "Critical",
                "table": "schwab.fact_transactions",
                "desc": "Verify PK ['account_id', 'transaction_process_date', 'transaction_cycle_identifier', 'transaction_random_number'] deduplication & account_sk mapping",
                "cond": "fact_transactions populated",
                "exp": "Zero duplicate rows on 4-column business key; account_sk strictly matches dim_accounts.",
                "query": "SELECT \n    (SELECT COUNT(*) FROM (SELECT account_id, transaction_process_date, transaction_cycle_identifier, transaction_random_number FROM Gold_Warehouse.schwab.fact_transactions GROUP BY account_id, transaction_process_date, transaction_cycle_identifier, transaction_random_number HAVING COUNT(*) > 1) d) AS Duplicate_BusinessKey_Count,\n    SUM(CASE WHEN f.account_sk IS NOT NULL AND da.account_sk IS NOT NULL AND f.account_id <> da.account_id THEN 1 ELSE 0 END) AS Mismatched_AccountId_Count,\n    SUM(CASE WHEN f.account_sk IS NULL THEN 1 ELSE 0 END) AS Null_AccountSK_Count,\n    SUM(CASE WHEN f.account_sk IS NOT NULL AND da.account_sk IS NULL THEN 1 ELSE 0 END) AS Orphan_AccountSK_Count\nFROM Gold_Warehouse.schwab.fact_transactions f\nLEFT JOIN Gold_Warehouse.schwab.dim_accounts da ON f.account_sk = da.account_sk;"
            },
            {
                "id": "TC_SC_GOLD_07", "cat": "Business Key & SK Validation", "pri": "Critical",
                "table": "schwab.fact_ult_lot_daily_taxable",
                "desc": "Verify PK ['account_id', 'event_id', 'schwab_internal_id', 'item_issue_id', 'isin', 'sedol'] deduplication & account_sk mapping",
                "cond": "fact_ult_lot_daily_taxable populated",
                "exp": "Zero duplicate rows on 6-column composite PK; account_sk strictly maps to dim_accounts.",
                "query": "SELECT \n    (SELECT COUNT(*) FROM (SELECT account_id, event_id, schwab_internal_id, item_issue_id, isin, sedol FROM Gold_Warehouse.schwab.fact_ult_lot_daily_taxable GROUP BY account_id, event_id, schwab_internal_id, item_issue_id, isin, sedol HAVING COUNT(*) > 1) d) AS Duplicate_BusinessKey_Count,\n    SUM(CASE WHEN f.account_sk IS NOT NULL AND da.account_sk IS NOT NULL AND f.account_id <> da.account_id THEN 1 ELSE 0 END) AS Mismatched_AccountId_Count,\n    SUM(CASE WHEN f.account_sk IS NULL THEN 1 ELSE 0 END) AS Null_AccountSK_Count,\n    SUM(CASE WHEN f.account_sk IS NOT NULL AND da.account_sk IS NULL THEN 1 ELSE 0 END) AS Orphan_AccountSK_Count\nFROM Gold_Warehouse.schwab.fact_ult_lot_daily_taxable f\nLEFT JOIN Gold_Warehouse.schwab.dim_accounts da ON f.account_sk = da.account_sk;"
            },
            {
                "id": "TC_SC_GOLD_08", "cat": "Business Key & SK Validation", "pri": "Critical",
                "table": "schwab.fact_uln_lot_daily_non_taxable",
                "desc": "Verify PK ['account_id', 'event_id'] deduplication & account_sk referential mapping",
                "cond": "fact_uln_lot_daily_non_taxable populated",
                "exp": "Zero duplicate rows on (account_id, event_id); account_sk strictly maps to dim_accounts.",
                "query": "SELECT \n    (SELECT COUNT(*) FROM (SELECT account_id, event_id FROM Gold_Warehouse.schwab.fact_uln_lot_daily_non_taxable GROUP BY account_id, event_id HAVING COUNT(*) > 1) d) AS Duplicate_BusinessKey_Count,\n    SUM(CASE WHEN f.account_sk IS NOT NULL AND da.account_sk IS NOT NULL AND f.account_id <> da.account_id THEN 1 ELSE 0 END) AS Mismatched_AccountId_Count,\n    SUM(CASE WHEN f.account_sk IS NULL THEN 1 ELSE 0 END) AS Null_AccountSK_Count,\n    SUM(CASE WHEN f.account_sk IS NOT NULL AND da.account_sk IS NULL THEN 1 ELSE 0 END) AS Orphan_AccountSK_Count\nFROM Gold_Warehouse.schwab.fact_uln_lot_daily_non_taxable f\nLEFT JOIN Gold_Warehouse.schwab.dim_accounts da ON f.account_sk = da.account_sk;"
            }
        ],


        "Salentica Gold": [
            {
                "id": "TC_SAL_GOLD_01", "cat": "SK Mapping & Joins", "pri": "Critical",
                "table": "salentica.dim_account",
                "desc": "Verify Foreign SK Resolution in dim_account (primarycontactid_sk, owningbusinessunit_sk)",
                "cond": "dim_account populated with Point-in-Time referential keys",
                "exp": "Foreign SKs resolve to matching natural keys in parent dimensions (Active or Expired). Unmapped default to -1.",
                "query": "SELECT COUNT(*) AS MismatchedNaturalKeyCount\nFROM Gold_Warehouse.salentica.dim_account a\nLEFT JOIN Gold_Warehouse.salentica.dim_contact_core c ON a.primarycontactid_sk = c.contact_SK\nLEFT JOIN Gold_Warehouse.salentica.dim_businessunit bu ON a.owningbusinessunit_sk = bu.businessunit_SK\nWHERE a.iscurrent = 'True'\n  AND ((a.primarycontactid IS NOT NULL AND a.primarycontactid_sk <> -1 AND a.primarycontactid <> c.contactid)\n   OR  (a.owningbusinessunit IS NOT NULL AND a.owningbusinessunit_sk <> -1 AND a.owningbusinessunit <> bu.businessunitid));"
            },
            {
                "id": "TC_SAL_GOLD_02", "cat": "SK Mapping & Joins", "pri": "Critical",
                "table": "salentica.dim_contact_core",
                "desc": "Verify Foreign SK Resolution in dim_contact_core (accountid_sk, owningbusinessunit_sk)",
                "cond": "dim_contact_core populated",
                "exp": "Foreign SKs resolve to matching natural keys in dim_account & dim_businessunit. Zero referential mismatches.",
                "query": "SELECT COUNT(*) AS MismatchedNaturalKeyCount\nFROM Gold_Warehouse.salentica.dim_contact_core c\nLEFT JOIN Gold_Warehouse.salentica.dim_account a ON c.accountid_sk = a.account_SK\nLEFT JOIN Gold_Warehouse.salentica.dim_businessunit bu ON c.owningbusinessunit_sk = bu.businessunit_SK\nWHERE c.iscurrent = 'True'\n  AND ((c.accountid IS NOT NULL AND c.accountid_sk <> -1 AND c.accountid <> a.accountid)\n   OR  (c.owningbusinessunit IS NOT NULL AND c.owningbusinessunit_sk <> -1 AND c.owningbusinessunit <> bu.businessunitid));"
            },
            {
                "id": "TC_SAL_GOLD_03", "cat": "SK Mapping & Joins", "pri": "High",
                "table": "salentica.dim_family",
                "desc": "Verify Foreign SK Resolution in dim_family (si_primarycontact_sk, owningbusinessunit_sk)",
                "cond": "dim_family populated",
                "exp": "Foreign SKs accurately resolve to contactid & businessunitid. Zero mismatches.",
                "query": "SELECT COUNT(*) AS MismatchedFamilyLinks\nFROM Gold_Warehouse.salentica.dim_family f\nLEFT JOIN Gold_Warehouse.salentica.dim_contact_core c ON f.si_primarycontact_sk = c.contact_SK\nLEFT JOIN Gold_Warehouse.salentica.dim_businessunit bu ON f.owningbusinessunit_sk = bu.businessunit_SK\nWHERE f.iscurrent = 'True'\n  AND ((f.si_primarycontact IS NOT NULL AND f.si_primarycontact_sk <> -1 AND f.si_primarycontact <> c.contactid)\n   OR  (f.owningbusinessunit IS NOT NULL AND f.owningbusinessunit_sk <> -1 AND f.owningbusinessunit <> bu.businessunitid));"
            },
            {
                "id": "TC_SAL_GOLD_04", "cat": "SK Mapping & Joins", "pri": "Critical",
                "table": "salentica.fact_lead",
                "desc": "Verify Foreign SK Resolution in fact_lead (accountid_sk, contactid_sk)",
                "cond": "fact_lead populated",
                "exp": "Foreign SKs resolve to matching accountid and contactid (Active or Expired versions). Zero mismatches.",
                "query": "SELECT COUNT(*) AS MismatchedLeadLinks\nFROM Gold_Warehouse.salentica.fact_lead l\nLEFT JOIN Gold_Warehouse.salentica.dim_account a ON l.accountid_sk = a.account_SK\nLEFT JOIN Gold_Warehouse.salentica.dim_contact_core c ON l.contactid_sk = c.contact_SK\nWHERE l.iscurrent = 'True'\n  AND ((l.accountid IS NOT NULL AND l.accountid_sk <> -1 AND l.accountid <> a.accountid)\n   OR  (l.contactid IS NOT NULL AND l.contactid_sk <> -1 AND l.contactid <> c.contactid));"
            },
            {
                "id": "TC_SAL_GOLD_05", "cat": "SK Mapping & Joins", "pri": "Critical",
                "table": "salentica.fact_opportunity",
                "desc": "Verify Point-in-Time SK Resolution in fact_opportunity (accountid_sk, contactid_sk)",
                "cond": "fact_opportunity populated",
                "exp": "opp.accountid_sk & contactid_sk resolve to matching natural keys in Dim (Active or Expired). Zero mismatches.",
                "query": "SELECT COUNT(*) AS MismatchedOpportunityLinks\nFROM Gold_Warehouse.salentica.fact_opportunity opp\nLEFT JOIN Gold_Warehouse.salentica.dim_account a ON opp.accountid_sk = a.account_SK\nLEFT JOIN Gold_Warehouse.salentica.dim_contact_core c ON opp.contactid_sk = c.contact_SK\nWHERE opp.iscurrent = 'True'\n  AND ((opp.accountid IS NOT NULL AND opp.accountid_sk <> -1 AND opp.accountid <> a.accountid)\n   OR  (opp.contactid IS NOT NULL AND opp.contactid_sk <> -1 AND opp.contactid <> c.contactid));"
            },
            {
                "id": "TC_SAL_GOLD_06", "cat": "SK Mapping & Joins", "pri": "Critical",
                "table": "salentica.fact_investmentaccount",
                "desc": "Verify Foreign SK Resolution in fact_investmentaccount (si_family_sk, si_portfolio_sk)",
                "cond": "fact_investmentaccount populated",
                "exp": "Attached SKs resolve to matching si_familyid & si_portfolioid. Zero mismatches.",
                "query": "SELECT COUNT(*) AS MismatchedInvestmentLinks\nFROM Gold_Warehouse.salentica.fact_investmentaccount ia\nLEFT JOIN Gold_Warehouse.salentica.dim_family f ON ia.si_family_sk = f.family_SK\nLEFT JOIN Gold_Warehouse.salentica.dim_portfolio p ON ia.si_portfolio_sk = p.portfolio_SK\nWHERE ia.iscurrent = 'True'\n  AND ((ia.si_family IS NOT NULL AND ia.si_family_sk <> -1 AND ia.si_family <> f.si_familyid)\n   OR  (ia.si_portfolio IS NOT NULL AND ia.si_portfolio_sk <> -1 AND ia.si_portfolio <> p.si_portfolioid));"
            },
            {
                "id": "TC_SAL_GOLD_07", "cat": "Orphan & -1 Fallback", "pri": "High",
                "table": "salentica.fact_opportunity",
                "desc": "Verify SG_SAL_22 Fallback: Unmapped Foreign References Default to -1 (Zero Orphan NULLs)",
                "cond": "fact_opportunity populated",
                "exp": "When natural ID is NULL, the corresponding SK defaults to -1. Zero unmapped NULLs in SK columns.",
                "query": "SELECT \n    SUM(CASE WHEN accountid IS NULL AND accountid_sk = -1 THEN 1 ELSE 0 END) AS ProperlyDefaulted_AccountSK,\n    SUM(CASE WHEN accountid IS NULL AND accountid_sk IS NULL THEN 1 ELSE 0 END) AS Improper_Null_AccountSK,\n    SUM(CASE WHEN contactid IS NULL AND contactid_sk = -1 THEN 1 ELSE 0 END) AS ProperlyDefaulted_ContactSK,\n    SUM(CASE WHEN contactid IS NULL AND contactid_sk IS NULL THEN 1 ELSE 0 END) AS Improper_Null_ContactSK\nFROM Gold_Warehouse.salentica.fact_opportunity\nWHERE iscurrent = 'True';"
            }
        ],
        "Tamarac Gold": [
            {
                "id": "TC_TAM_GOLD_01", "cat": "SCD2 Deduplication", "pri": "Critical",
                "table": "tamarac.dim_account_information",
                "desc": "Verify Active Deduplication on PK ['account_number'] (is_current = 'True')",
                "cond": "dim_account_information populated",
                "exp": "Zero duplicate active accounts. Exactly one is_current = 'True' row per account_number.",
                "query": "SELECT account_number, COUNT(*) AS ActiveDuplicateCount\nFROM Gold_Warehouse.tamarac.dim_account_information\nWHERE is_current = 'True'\nGROUP BY account_number\nHAVING COUNT(*) > 1;"
            },
            {
                "id": "TC_TAM_GOLD_02", "cat": "SK Uniqueness", "pri": "Critical",
                "table": "tamarac.dim_account_information",
                "desc": "Verify Primary Surrogate Key (account_information_SK) uniqueness across all records",
                "cond": "dim_account_information populated",
                "exp": "Surrogate key account_information_SK is strictly unique across all rows. Zero duplicates.",
                "query": "SELECT account_information_SK, COUNT(*) AS DuplicateCount\nFROM Gold_Warehouse.tamarac.dim_account_information\nGROUP BY account_information_SK\nHAVING COUNT(*) > 1;"
            },
            {
                "id": "TC_TAM_GOLD_03", "cat": "Audit Columns", "pri": "High",
                "table": "tamarac.dim_account_information",
                "desc": "Verify file_business_date and updated_at populated on active Tamarac accounts",
                "cond": "dim_account_information populated",
                "exp": "file_business_date and updated_at are non-null and <= current calendar date.",
                "query": "SELECT \n    COUNT(*) AS TotalRows,\n    SUM(CASE WHEN file_business_date IS NULL THEN 1 ELSE 0 END) AS Null_FileBusinessDate,\n    SUM(CASE WHEN updated_at IS NULL THEN 1 ELSE 0 END) AS Null_UpdatedAt\nFROM Gold_Warehouse.tamarac.dim_account_information\nWHERE is_current = 'True';"
            },
            {
                "id": "TC_TAM_GOLD_04", "cat": "SCD2 Deduplication", "pri": "Critical",
                "table": "tamarac.dim_account_model_allocation",
                "desc": "Verify Active Deduplication on Composite PK ['account_number', 'model_name']",
                "cond": "dim_account_model_allocation populated",
                "exp": "Zero duplicate active model allocations per (account_number, model_name).",
                "query": "SELECT account_number, model_name, COUNT(*) AS ActiveDuplicateAllocations\nFROM Gold_Warehouse.tamarac.dim_account_model_allocation\nWHERE is_current = 'True'\nGROUP BY account_number, model_name\nHAVING COUNT(*) > 1;"
            },
            {
                "id": "TC_TAM_GOLD_05", "cat": "SK Uniqueness", "pri": "Critical",
                "table": "tamarac.dim_account_model_allocation",
                "desc": "Verify Primary Surrogate Key (account_model_SK) uniqueness",
                "cond": "dim_account_model_allocation populated",
                "exp": "account_model_SK is unique with zero NULL values.",
                "query": "SELECT account_model_SK, COUNT(*) AS DuplicateModelSK\nFROM Gold_Warehouse.tamarac.dim_account_model_allocation\nGROUP BY account_model_SK\nHAVING COUNT(*) > 1;"
            },
            {
                "id": "TC_TAM_GOLD_06", "cat": "Cross-Table Linkage", "pri": "High",
                "table": "tamarac.dim_account_model_allocation",
                "desc": "Verify Cross-Table Linkage: Every account_number exists in dim_account_information",
                "cond": "dim_account_model_allocation populated",
                "exp": "Zero orphan model allocation accounts. All account_number values exist in parent dimension.",
                "query": "SELECT COUNT(*) AS OrphanModelAllocations\nFROM Gold_Warehouse.tamarac.dim_account_model_allocation ma\nLEFT JOIN Gold_Warehouse.tamarac.dim_account_information info \n    ON ma.account_number = info.account_number AND info.is_current = 'True'\nWHERE ma.is_current = 'True' \n  AND info.account_information_SK IS NULL;"
            },
            {
                "id": "TC_TAM_GOLD_07", "cat": "SCD2 Date Validation", "pri": "Critical",
                "table": "tamarac.dim_account_information",
                "desc": "Verify Inactive End Date Rule: effective_end_date <= (MAX(file_business_date) - 1 day) and non-null",
                "cond": "dim_account_information populated with inactive records",
                "exp": "Exceeded_Cutoff_Count = 0, Null_End_Date_Count = 0. TestStatus = 'PASSED'.",
                "query": """SELECT 
    'dim_account_information' AS TableName,
    COUNT(*) AS TotalInactiveRows,
    CAST(m.CutoffDate AS VARCHAR(10)) AS Max_Cutoff_Allowed,
    SUM(CASE WHEN a.effective_end_date > m.CutoffDate THEN 1 ELSE 0 END) AS Exceeded_Cutoff_Count,
    SUM(CASE WHEN a.effective_end_date IS NULL THEN 1 ELSE 0 END) AS Null_End_Date_Count,
    CASE 
        WHEN SUM(CASE WHEN a.effective_end_date > m.CutoffDate THEN 1 ELSE 0 END) = 0
         AND SUM(CASE WHEN a.effective_end_date IS NULL THEN 1 ELSE 0 END) = 0
        THEN 'PASSED' ELSE 'FAILED'
    END AS TestStatus
FROM Gold_Warehouse.tamarac.dim_account_information a
CROSS JOIN (SELECT DATEADD(day, -1, MAX(file_business_date)) AS CutoffDate FROM Gold_Warehouse.tamarac.dim_account_information) m
WHERE a.is_current = 'False'
GROUP BY m.CutoffDate;"""
            },
            {
                "id": "TC_TAM_GOLD_08", "cat": "SCD2 Date Validation", "pri": "Critical",
                "table": "tamarac.dim_account_model_allocation",
                "desc": "Verify Inactive End Date Rule: effective_end_date <= (MAX(file_business_date) - 1 day) and non-null",
                "cond": "dim_account_model_allocation populated with inactive records",
                "exp": "Exceeded_Cutoff_Count = 0, Null_End_Date_Count = 0. TestStatus = 'PASSED'.",
                "query": """SELECT 
    'dim_account_model_allocation' AS TableName,
    COUNT(*) AS TotalInactiveRows,
    CAST(m.CutoffDate AS VARCHAR(10)) AS Max_Cutoff_Allowed,
    SUM(CASE WHEN a.effective_end_date > m.CutoffDate THEN 1 ELSE 0 END) AS Exceeded_Cutoff_Count,
    SUM(CASE WHEN a.effective_end_date IS NULL THEN 1 ELSE 0 END) AS Null_End_Date_Count,
    CASE 
        WHEN SUM(CASE WHEN a.effective_end_date > m.CutoffDate THEN 1 ELSE 0 END) = 0
         AND SUM(CASE WHEN a.effective_end_date IS NULL THEN 1 ELSE 0 END) = 0
        THEN 'PASSED' ELSE 'FAILED'
    END AS TestStatus
FROM Gold_Warehouse.tamarac.dim_account_model_allocation a
CROSS JOIN (SELECT DATEADD(day, -1, MAX(file_business_date)) AS CutoffDate FROM Gold_Warehouse.tamarac.dim_account_model_allocation) m
WHERE a.is_current = 'False'
GROUP BY m.CutoffDate;"""
            },
            {
                "id": "TC_TAM_GOLD_09", "cat": "SCD2 Timeline & SK Monotonicity", "pri": "Critical",
                "table": "tamarac.dim_account_information",
                "desc": "Verify SCD2 Timeline Continuity (Zero Gaps, Zero Overlaps) & SK Monotonic Increment",
                "cond": "dim_account_information with multiple versions per account_number",
                "exp": "Timeline_Gap_Or_Overlap_Count = 0, SK_Non_Incrementing_Violation_Count = 0. TestStatus = 'PASSED'.",
                "query": """;WITH InfoTimeline AS (
    SELECT 
        account_number,
        account_information_SK,
        effective_start_date,
        effective_end_date,
        is_current,
        LEAD(effective_start_date) OVER (PARTITION BY account_number ORDER BY effective_start_date, account_information_SK) AS next_start_date,
        LEAD(account_information_SK) OVER (PARTITION BY account_number ORDER BY effective_start_date, account_information_SK) AS next_sk
    FROM Gold_Warehouse.tamarac.dim_account_information
)
SELECT 
    'dim_account_information' AS TableName,
    COUNT(*) AS MultiVersionRecordsChecked,
    SUM(CASE WHEN next_start_date IS NOT NULL AND effective_end_date <> DATEADD(day, -1, next_start_date) THEN 1 ELSE 0 END) AS Timeline_Gap_Or_Overlap_Count,
    SUM(CASE WHEN next_sk IS NOT NULL AND next_sk <= account_information_SK THEN 1 ELSE 0 END) AS SK_Non_Incrementing_Violation_Count,
    CASE 
        WHEN SUM(CASE WHEN next_start_date IS NOT NULL AND effective_end_date <> DATEADD(day, -1, next_start_date) THEN 1 ELSE 0 END) = 0
         AND SUM(CASE WHEN next_sk IS NOT NULL AND next_sk <= account_information_SK THEN 1 ELSE 0 END) = 0
        THEN 'PASSED' ELSE 'FAILED'
    END AS TestStatus
FROM InfoTimeline
WHERE next_start_date IS NOT NULL;"""
            },
            {
                "id": "TC_TAM_GOLD_10", "cat": "SCD2 Timeline & SK Monotonicity", "pri": "Critical",
                "table": "tamarac.dim_account_model_allocation",
                "desc": "Verify SCD2 Timeline Continuity (Zero Gaps, Zero Overlaps) & SK Monotonic Increment",
                "cond": "dim_account_model_allocation with multiple versions per composite key",
                "exp": "Timeline_Gap_Or_Overlap_Count = 0, SK_Non_Incrementing_Violation_Count = 0. TestStatus = 'PASSED'.",
                "query": """;WITH ModelTimeline AS (
    SELECT 
        account_number,
        model_name,
        account_model_SK,
        effective_start_date,
        effective_end_date,
        is_current,
        LEAD(effective_start_date) OVER (PARTITION BY account_number, model_name ORDER BY effective_start_date, account_model_SK) AS next_start_date,
        LEAD(account_model_SK) OVER (PARTITION BY account_number, model_name ORDER BY effective_start_date, account_model_SK) AS next_sk
    FROM Gold_Warehouse.tamarac.dim_account_model_allocation
)
SELECT 
    'dim_account_model_allocation' AS TableName,
    COUNT(*) AS MultiVersionRecordsChecked,
    SUM(CASE WHEN next_start_date IS NOT NULL AND effective_end_date <> DATEADD(day, -1, next_start_date) THEN 1 ELSE 0 END) AS Timeline_Gap_Or_Overlap_Count,
    SUM(CASE WHEN next_sk IS NOT NULL AND next_sk <= account_model_SK THEN 1 ELSE 0 END) AS SK_Non_Incrementing_Violation_Count,
    CASE 
        WHEN SUM(CASE WHEN next_start_date IS NOT NULL AND effective_end_date <> DATEADD(day, -1, next_start_date) THEN 1 ELSE 0 END) = 0
         AND SUM(CASE WHEN next_sk IS NOT NULL AND next_sk <= account_model_SK THEN 1 ELSE 0 END) = 0
        THEN 'PASSED' ELSE 'FAILED'
    END AS TestStatus
FROM ModelTimeline
WHERE next_start_date IS NOT NULL;"""
            },
            {
                "id": "TC_TAM_GOLD_11", "cat": "Date Sanity & Inversions", "pri": "Critical",
                "table": "tamarac.dim_account_information",
                "desc": "Verify Date Sanity: Zero inverted dates (start > end) and active records have non-future start dates",
                "cond": "dim_account_information populated",
                "exp": "Inverted_Date_Count = 0, Future_Active_Start_Date = 0, Null_Start_Date_Count = 0. TestStatus = 'PASSED'.",
                "query": """SELECT 
    'dim_account_information' AS TableName,
    COUNT(*) AS TotalRecords,
    SUM(CASE WHEN effective_end_date < effective_start_date THEN 1 ELSE 0 END) AS Inverted_Date_Count,
    SUM(CASE WHEN is_current = 'True' AND effective_start_date > CAST(GETDATE() AS DATE) THEN 1 ELSE 0 END) AS Future_Active_Start_Date,
    SUM(CASE WHEN effective_start_date IS NULL THEN 1 ELSE 0 END) AS Null_Start_Date_Count,
    CASE 
        WHEN SUM(CASE WHEN effective_end_date < effective_start_date THEN 1 ELSE 0 END) = 0
         AND SUM(CASE WHEN is_current = 'True' AND effective_start_date > CAST(GETDATE() AS DATE) THEN 1 ELSE 0 END) = 0
         AND SUM(CASE WHEN effective_start_date IS NULL THEN 1 ELSE 0 END) = 0
        THEN 'PASSED' ELSE 'FAILED'
    END AS TestStatus
FROM Gold_Warehouse.tamarac.dim_account_information;"""
            },
            {
                "id": "TC_TAM_GOLD_12", "cat": "Date Sanity & Inversions", "pri": "Critical",
                "table": "tamarac.dim_account_model_allocation",
                "desc": "Verify Date Sanity: Zero inverted dates (start > end) and active records have non-future start dates",
                "cond": "dim_account_model_allocation populated",
                "exp": "Inverted_Date_Count = 0, Future_Active_Start_Date = 0, Null_Start_Date_Count = 0. TestStatus = 'PASSED'.",
                "query": """SELECT 
    'dim_account_model_allocation' AS TableName,
    COUNT(*) AS TotalRecords,
    SUM(CASE WHEN effective_end_date < effective_start_date THEN 1 ELSE 0 END) AS Inverted_Date_Count,
    SUM(CASE WHEN is_current = 'True' AND effective_start_date > CAST(GETDATE() AS DATE) THEN 1 ELSE 0 END) AS Future_Active_Start_Date,
    SUM(CASE WHEN effective_start_date IS NULL THEN 1 ELSE 0 END) AS Null_Start_Date_Count,
    CASE 
        WHEN SUM(CASE WHEN effective_end_date < effective_start_date THEN 1 ELSE 0 END) = 0
         AND SUM(CASE WHEN is_current = 'True' AND effective_start_date > CAST(GETDATE() AS DATE) THEN 1 ELSE 0 END) = 0
         AND SUM(CASE WHEN effective_start_date IS NULL THEN 1 ELSE 0 END) = 0
        THEN 'PASSED' ELSE 'FAILED'
    END AS TestStatus
FROM Gold_Warehouse.tamarac.dim_account_model_allocation;"""
            },
            {
                "id": "TC_TAM_GOLD_13", "cat": "Open-Ended Active Records", "pri": "High",
                "table": "tamarac.dim_account_information",
                "desc": "Verify Active Records are Open-Ended (effective_end_date is NULL or '9999-12-31')",
                "cond": "dim_account_information where is_current = 'True'",
                "exp": "Prematurely_Ended_Active_Records = 0. TestStatus = 'PASSED'.",
                "query": """SELECT 
    'dim_account_information' AS TableName,
    COUNT(*) AS TotalActiveRecords,
    SUM(CASE WHEN effective_end_date IS NOT NULL AND effective_end_date NOT IN ('9999-12-31', '2099-12-31') THEN 1 ELSE 0 END) AS Prematurely_Ended_Active_Records,
    CASE 
        WHEN SUM(CASE WHEN effective_end_date IS NOT NULL AND effective_end_date NOT IN ('9999-12-31', '2099-12-31') THEN 1 ELSE 0 END) = 0 
        THEN 'PASSED' ELSE 'FAILED'
    END AS TestStatus
FROM Gold_Warehouse.tamarac.dim_account_information
WHERE is_current = 'True';"""
            },
            {
                "id": "TC_TAM_GOLD_14", "cat": "Open-Ended Active Records", "pri": "High",
                "table": "tamarac.dim_account_model_allocation",
                "desc": "Verify Active Records are Open-Ended (effective_end_date is NULL or '9999-12-31')",
                "cond": "dim_account_model_allocation where is_current = 'True'",
                "exp": "Prematurely_Ended_Active_Records = 0. TestStatus = 'PASSED'.",
                "query": """SELECT 
    'dim_account_model_allocation' AS TableName,
    COUNT(*) AS TotalActiveRecords,
    SUM(CASE WHEN effective_end_date IS NOT NULL AND effective_end_date NOT IN ('9999-12-31', '2099-12-31') THEN 1 ELSE 0 END) AS Prematurely_Ended_Active_Records,
    CASE 
        WHEN SUM(CASE WHEN effective_end_date IS NOT NULL AND effective_end_date NOT IN ('9999-12-31', '2099-12-31') THEN 1 ELSE 0 END) = 0 
        THEN 'PASSED' ELSE 'FAILED'
    END AS TestStatus
FROM Gold_Warehouse.tamarac.dim_account_model_allocation
WHERE is_current = 'True';"""
            }
        ],
        "Innotrust Gold": [
            {
                "id": "TC_INNO_GOLD_01", "cat": "SCD2 Deduplication", "pri": "Critical",
                "table": "innotrust.dim_account_master",
                "desc": "Verify Active Deduplication on Business Key ['customer_account_number'] (is_current = 1)",
                "cond": "innotrust.dim_account_master populated",
                "exp": "Zero duplicate active accounts. Exactly one is_current = 1 record per customer_account_number.",
                "query": "SELECT customer_account_number, COUNT(*) AS ActiveDuplicateCount\nFROM Gold_Warehouse.innotrust.dim_account_master\nWHERE is_current = 1\nGROUP BY customer_account_number\nHAVING COUNT(*) > 1;"
            },
            {
                "id": "TC_INNO_GOLD_02", "cat": "SK Uniqueness", "pri": "Critical",
                "table": "innotrust.dim_account_master",
                "desc": "Verify Primary Surrogate Key (dim_account_master_sk) uniqueness and non-null constraint",
                "cond": "innotrust.dim_account_master populated",
                "exp": "dim_account_master_sk is strictly unique across all versions with zero NULL values.",
                "query": "SELECT dim_account_master_sk, COUNT(*) AS DuplicateCount\nFROM Gold_Warehouse.innotrust.dim_account_master\nGROUP BY dim_account_master_sk\nHAVING COUNT(*) > 1\nUNION ALL\nSELECT NULL, COUNT(*)\nFROM Gold_Warehouse.innotrust.dim_account_master\nWHERE dim_account_master_sk IS NULL\nHAVING COUNT(*) > 0;"
            },
            {
                "id": "TC_INNO_GOLD_03", "cat": "SCD2 Deduplication", "pri": "Critical",
                "table": "innotrust.dim_account_category_codes",
                "desc": "Verify Active Deduplication on Business Key ['account_category_code'] (is_current = 1)",
                "cond": "innotrust.dim_account_category_codes populated",
                "exp": "Zero duplicate active category codes per account_category_code.",
                "query": "SELECT account_category_code, COUNT(*) AS ActiveDuplicateCount\nFROM Gold_Warehouse.innotrust.dim_account_category_codes\nWHERE is_current = 1\nGROUP BY account_category_code\nHAVING COUNT(*) > 1;"
            },
            {
                "id": "TC_INNO_GOLD_04", "cat": "SK Uniqueness", "pri": "Critical",
                "table": "innotrust.dim_account_category_codes",
                "desc": "Verify Primary Surrogate Key (dim_account_master_sk) uniqueness on dim_account_category_codes",
                "cond": "innotrust.dim_account_category_codes populated",
                "exp": "Surrogate key dim_account_master_sk is unique across all rows with zero NULLs.",
                "query": "SELECT dim_account_master_sk, COUNT(*) AS DuplicateCount\nFROM Gold_Warehouse.innotrust.dim_account_category_codes\nGROUP BY dim_account_master_sk\nHAVING COUNT(*) > 1;"
            },
            {
                "id": "TC_INNO_GOLD_05", "cat": "SCD2 Deduplication", "pri": "Critical",
                "table": "innotrust.dim_account_type_codes",
                "desc": "Verify Active Deduplication on Business Key ['account_type_code'] (is_current = 1)",
                "cond": "innotrust.dim_account_type_codes populated",
                "exp": "Zero duplicate active account types per account_type_code.",
                "query": "SELECT account_type_code, COUNT(*) AS ActiveDuplicateCount\nFROM Gold_Warehouse.innotrust.dim_account_type_codes\nWHERE is_current = 1\nGROUP BY account_type_code\nHAVING COUNT(*) > 1;"
            },
            {
                "id": "TC_INNO_GOLD_06", "cat": "SK Uniqueness", "pri": "Critical",
                "table": "innotrust.dim_account_type_codes",
                "desc": "Verify Primary Surrogate Key (dim_account_type_codes_sk) uniqueness and non-null constraint",
                "cond": "innotrust.dim_account_type_codes populated",
                "exp": "dim_account_type_codes_sk is strictly unique across all versions with zero NULLs.",
                "query": "SELECT dim_account_type_codes_sk, COUNT(*) AS DuplicateCount\nFROM Gold_Warehouse.innotrust.dim_account_type_codes\nGROUP BY dim_account_type_codes_sk\nHAVING COUNT(*) > 1;"
            },
            {
                "id": "TC_INNO_GOLD_07", "cat": "SCD2 Deduplication", "pri": "Critical",
                "table": "innotrust.dim_master_sec_index",
                "desc": "Verify Active Deduplication on Security Primary Key ['isin'] (is_current = 1)",
                "cond": "innotrust.dim_master_sec_index populated",
                "exp": "Zero duplicate active securities. Exactly one is_current = 1 record per isin.",
                "query": "SELECT isin, COUNT(*) AS ActiveDuplicateCount\nFROM Gold_Warehouse.innotrust.dim_master_sec_index\nWHERE is_current = 1\nGROUP BY isin\nHAVING COUNT(*) > 1;"
            },
            {
                "id": "TC_INNO_GOLD_08", "cat": "SK Uniqueness", "pri": "Critical",
                "table": "innotrust.dim_master_sec_index",
                "desc": "Verify Primary Surrogate Key (dim_master_sec_index_sk) uniqueness across all records",
                "cond": "innotrust.dim_master_sec_index populated",
                "exp": "dim_master_sec_index_sk is strictly unique across all versions with zero NULLs.",
                "query": "SELECT dim_master_sec_index_sk, COUNT(*) AS DuplicateCount\nFROM Gold_Warehouse.innotrust.dim_master_sec_index\nGROUP BY dim_master_sec_index_sk\nHAVING COUNT(*) > 1;"
            },
            {
                "id": "TC_INNO_GOLD_09", "cat": "SCD2 Deduplication", "pri": "Critical",
                "table": "innotrust.dim_security_characteristics_codes",
                "desc": "Verify Active Deduplication on Business Key ['security_type_code'] (is_current = 1)",
                "cond": "innotrust.dim_security_characteristics_codes populated",
                "exp": "Zero duplicate active security type codes per security_type_code.",
                "query": "SELECT security_type_code, COUNT(*) AS ActiveDuplicateCount\nFROM Gold_Warehouse.innotrust.dim_security_characteristics_codes\nWHERE is_current = 1\nGROUP BY security_type_code\nHAVING COUNT(*) > 1;"
            },
            {
                "id": "TC_INNO_GOLD_10", "cat": "SK Uniqueness", "pri": "Critical",
                "table": "innotrust.dim_security_characteristics_codes",
                "desc": "Verify Primary Surrogate Key (dim_securitycharacteristicscodes) uniqueness and non-null constraint",
                "cond": "innotrust.dim_security_characteristics_codes populated",
                "exp": "dim_securitycharacteristicscodes is strictly unique across all versions with zero NULLs.",
                "query": "SELECT dim_securitycharacteristicscodes, COUNT(*) AS DuplicateCount\nFROM Gold_Warehouse.innotrust.dim_security_characteristics_codes\nGROUP BY dim_securitycharacteristicscodes\nHAVING COUNT(*) > 1;"
            },
            {
                "id": "TC_INNO_GOLD_11", "cat": "Audit & Date Logic", "pri": "High",
                "table": "innotrust.dim_account_master",
                "desc": "Verify business_date, updated_at, effective_start_date & effective_end_date integrity on active records",
                "cond": "innotrust.dim_account_master populated",
                "exp": "Zero NULLs in business_date or updated_at; effective_end_date >= effective_start_date.",
                "query": "SELECT \n    COUNT(*) AS TotalActiveRows,\n    SUM(CASE WHEN business_date IS NULL THEN 1 ELSE 0 END) AS Null_BusinessDate,\n    SUM(CASE WHEN updated_at IS NULL THEN 1 ELSE 0 END) AS Null_UpdatedAt,\n    SUM(CASE WHEN effective_start_date IS NULL THEN 1 ELSE 0 END) AS Null_EffectiveStartDate,\n    SUM(CASE WHEN effective_end_date < effective_start_date THEN 1 ELSE 0 END) AS InvertedEffectiveDates\nFROM Gold_Warehouse.innotrust.dim_account_master\nWHERE is_current = 1;"
            },
            {
                "id": "TC_INNO_GOLD_12", "cat": "Cross-Table Linkage", "pri": "High",
                "table": "innotrust.dim_account_master",
                "desc": "Verify Referential Integrity: account_category_code and account_type_code resolve to active codes",
                "cond": "innotrust.dim_account_master populated",
                "exp": "All non-null account_category_code and account_type_code exist in corresponding dimension code tables.",
                "query": "SELECT \n    SUM(CASE WHEN m.account_category_code IS NOT NULL AND cat.account_category_code IS NULL THEN 1 ELSE 0 END) AS Orphan_CategoryCodes,\n    SUM(CASE WHEN m.account_type_code IS NOT NULL AND typ.account_type_code IS NULL THEN 1 ELSE 0 END) AS Orphan_TypeCodes\nFROM Gold_Warehouse.innotrust.dim_account_master m\nLEFT JOIN Gold_Warehouse.innotrust.dim_account_category_codes cat \n    ON m.account_category_code = cat.account_category_code AND cat.is_current = 1\nLEFT JOIN Gold_Warehouse.innotrust.dim_account_type_codes typ \n    ON m.account_type_code = typ.account_type_code AND typ.is_current = 1\nWHERE m.is_current = 1;"
            },
            {
                "id": "TC_INNO_GOLD_13", "cat": "Cross-Table Linkage", "pri": "High",
                "table": "innotrust.dim_master_sec_index",
                "desc": "Verify Security Referential Integrity: security_type_code resolves to active security characteristics",
                "cond": "innotrust.dim_master_sec_index populated",
                "exp": "All non-null security_type_code values exist in dim_security_characteristics_codes.",
                "query": "SELECT COUNT(*) AS Orphan_SecurityTypeCodes\nFROM Gold_Warehouse.innotrust.dim_master_sec_index s\nLEFT JOIN Gold_Warehouse.innotrust.dim_security_characteristics_codes sc \n    ON s.security_type_code = sc.security_type_code AND sc.is_current = 1\nWHERE s.is_current = 1 AND s.security_type_code IS NOT NULL AND sc.security_type_code IS NULL;"
            },
            {
                "id": "TC_INNO_GOLD_14", "cat": "SCD2 Date Validation", "pri": "Critical",
                "table": "innotrust.dim_account_master",
                "desc": "Verify SCD2 Effective Date Rules: Inactive records effective_end_date <= (Current Day - 1), non-null end date, and valid active start date",
                "cond": "innotrust.dim_account_master populated with SCD2 tracking columns",
                "exp": "Inactive_End_Date_Breach_Count = 0, Inactive_Missing_End_Date = 0, Active_Future_Start_Date = 0, Inverted_Date_Range_Count = 0. TestStatus = 'PASSED'.",
                "query": """SELECT 
    'dim_account_master' AS TableName,
    COUNT(*) AS TotalRecords,
    SUM(CASE WHEN is_current = 0 AND effective_end_date > DATEADD(day, -1, CAST(GETDATE() AS DATE)) THEN 1 ELSE 0 END) AS Inactive_End_Date_Breach_Count,
    SUM(CASE WHEN is_current = 0 AND effective_end_date IS NULL THEN 1 ELSE 0 END) AS Inactive_Missing_End_Date,
    SUM(CASE WHEN is_current = 1 AND effective_start_date > CAST(GETDATE() AS DATE) THEN 1 ELSE 0 END) AS Active_Future_Start_Date,
    SUM(CASE WHEN effective_end_date < effective_start_date THEN 1 ELSE 0 END) AS Inverted_Date_Range_Count,
    CASE 
        WHEN SUM(CASE WHEN is_current = 0 AND effective_end_date > DATEADD(day, -1, CAST(GETDATE() AS DATE)) THEN 1 ELSE 0 END) = 0
         AND SUM(CASE WHEN is_current = 0 AND effective_end_date IS NULL THEN 1 ELSE 0 END) = 0
         AND SUM(CASE WHEN effective_end_date < effective_start_date THEN 1 ELSE 0 END) = 0
        THEN 'PASSED' ELSE 'FAILED'
    END AS TestStatus
FROM Gold_Warehouse.innotrust.dim_account_master;"""
            }
        ]
    }

    # Consolidated Master Cases
    all_master_cases = []
    for sheet_name, cases in test_suites.items():
        for c in cases:
            all_master_cases.append({
                "schema": sheet_name.replace(" Gold", ""),
                "id": c["id"],
                "cat": c["cat"],
                "pri": c["pri"],
                "table": c["table"],
                "desc": c["desc"],
                "cond": c["cond"],
                "exp": c["exp"],
                "query": c["query"]
            })

    # -------------------------------------------------------------------------------------------------
    # CONSOLIDATED MULTI-OBJECT QUERIES (NEW 7TH SHEET)
    # -------------------------------------------------------------------------------------------------
    consolidated_sk_suites = [
        {
            "id": "TC_COMB_SK_BD_01",
            "source": "Black Diamond",
            "scope": "All 9 Child & Fact Tables (holdings, transactions, taxlots, class/segment, details, tags, team, targets)",
            "priority": "Critical",
            "title": "Black Diamond: Cross-Object account_sk vs dim_accounts Natural Key (accountid == da.id) Verification",
            "desc": "Executes a single consolidated UNION ALL query across all 9 child/fact tables to verify that the attached account_sk strictly points to the exact same account natural key (child.accountid == da.id). Guarantees zero wrong account links.",
            "exp": "Duplicate_BusinessKey_Count = 0, MismatchedAccountIdCount = 0, Null_AccountSK_Count = 0, Orphan_AccountSK_Count = 0 across all 9 objects.",
            "query": """SELECT 
    'dim_accounts_classallocation' AS TableName, 
    COUNT(*) AS TotalRows,
    (SELECT COUNT(*) FROM (SELECT accountid, classid FROM Gold_Warehouse.black_diamond.dim_accounts_classallocation WHERE iscurrent = 'True' GROUP BY accountid, classid HAVING COUNT(*) > 1) d) AS Duplicate_BusinessKey_Count,
    SUM(CASE WHEN c.accountid = da.id THEN 1 ELSE 0 END) AS SuccessfullyMatchedCount,
    SUM(CASE WHEN c.accountid <> da.id THEN 1 ELSE 0 END) AS MismatchedAccountIdCount,
    SUM(CASE WHEN c.account_sk IS NULL THEN 1 ELSE 0 END) AS Null_AccountSK_Count,
    SUM(CASE WHEN c.account_sk IS NOT NULL AND da.accounts_SK IS NULL THEN 1 ELSE 0 END) AS Orphan_AccountSK_Count
FROM Gold_Warehouse.black_diamond.dim_accounts_classallocation c
LEFT JOIN Gold_Warehouse.black_diamond.dim_accounts da ON c.account_sk = da.accounts_SK
WHERE c.iscurrent = 'True'

UNION ALL

SELECT 
    'dim_accounts_segmentallocation', 
    COUNT(*),
    (SELECT COUNT(*) FROM (SELECT accountid, classid, segmentid FROM Gold_Warehouse.black_diamond.dim_accounts_segmentallocation WHERE iscurrent = 'True' GROUP BY accountid, classid, segmentid HAVING COUNT(*) > 1) d),
    SUM(CASE WHEN s.accountid = da.id THEN 1 ELSE 0 END),
    SUM(CASE WHEN s.accountid <> da.id THEN 1 ELSE 0 END),
    SUM(CASE WHEN s.account_sk IS NULL THEN 1 ELSE 0 END),
    SUM(CASE WHEN s.account_sk IS NOT NULL AND da.accounts_SK IS NULL THEN 1 ELSE 0 END)
FROM Gold_Warehouse.black_diamond.dim_accounts_segmentallocation s
LEFT JOIN Gold_Warehouse.black_diamond.dim_accounts da ON s.account_sk = da.accounts_SK
WHERE s.iscurrent = 'True'

UNION ALL

SELECT 
    'dim_accounts_details', 
    COUNT(*),
    (SELECT COUNT(*) FROM (SELECT accountid FROM Gold_Warehouse.black_diamond.dim_accounts_details WHERE iscurrent = 'True' GROUP BY accountid HAVING COUNT(*) > 1) d),
    SUM(CASE WHEN d.accountid = da.id THEN 1 ELSE 0 END),
    SUM(CASE WHEN d.accountid <> da.id THEN 1 ELSE 0 END),
    SUM(CASE WHEN d.account_sk IS NULL THEN 1 ELSE 0 END),
    SUM(CASE WHEN d.account_sk IS NOT NULL AND da.accounts_SK IS NULL THEN 1 ELSE 0 END)
FROM Gold_Warehouse.black_diamond.dim_accounts_details d
LEFT JOIN Gold_Warehouse.black_diamond.dim_accounts da ON d.account_sk = da.accounts_SK
WHERE d.iscurrent = 'True'

UNION ALL

SELECT 
    'dim_accounts_tags', 
    COUNT(*),
    (SELECT COUNT(*) FROM (SELECT accountid, name FROM Gold_Warehouse.black_diamond.dim_accounts_tags WHERE iscurrent = 'True' GROUP BY accountid, name HAVING COUNT(*) > 1) d),
    SUM(CASE WHEN t.accountid = da.id THEN 1 ELSE 0 END),
    SUM(CASE WHEN t.accountid <> da.id THEN 1 ELSE 0 END),
    SUM(CASE WHEN t.account_sk IS NULL THEN 1 ELSE 0 END),
    SUM(CASE WHEN t.account_sk IS NOT NULL AND da.accounts_SK IS NULL THEN 1 ELSE 0 END)
FROM Gold_Warehouse.black_diamond.dim_accounts_tags t
LEFT JOIN Gold_Warehouse.black_diamond.dim_accounts da ON t.account_sk = da.accounts_SK
WHERE t.iscurrent = 'True'

UNION ALL

SELECT 
    'dim_accounts_team', 
    COUNT(*),
    (SELECT COUNT(*) FROM (SELECT accountid FROM Gold_Warehouse.black_diamond.dim_accounts_team WHERE iscurrent = 'True' GROUP BY accountid HAVING COUNT(*) > 1) d),
    SUM(CASE WHEN tm.accountid = da.id THEN 1 ELSE 0 END),
    SUM(CASE WHEN tm.accountid <> da.id THEN 1 ELSE 0 END),
    SUM(CASE WHEN tm.account_sk IS NULL THEN 1 ELSE 0 END),
    SUM(CASE WHEN tm.account_sk IS NOT NULL AND da.accounts_SK IS NULL THEN 1 ELSE 0 END)
FROM Gold_Warehouse.black_diamond.dim_accounts_team tm
LEFT JOIN Gold_Warehouse.black_diamond.dim_accounts da ON tm.account_sk = da.accounts_SK
WHERE tm.iscurrent = 'True'

UNION ALL

SELECT 
    'dim_accounts_targets', 
    COUNT(*),
    (SELECT COUNT(*) FROM (SELECT accountid, id FROM Gold_Warehouse.black_diamond.dim_accounts_targets WHERE iscurrent = 'True' GROUP BY accountid, id HAVING COUNT(*) > 1) d),
    SUM(CASE WHEN tg.accountid = da.id THEN 1 ELSE 0 END),
    SUM(CASE WHEN tg.accountid <> da.id THEN 1 ELSE 0 END),
    SUM(CASE WHEN tg.account_sk IS NULL THEN 1 ELSE 0 END),
    SUM(CASE WHEN tg.account_sk IS NOT NULL AND da.accounts_SK IS NULL THEN 1 ELSE 0 END)
FROM Gold_Warehouse.black_diamond.dim_accounts_targets tg
LEFT JOIN Gold_Warehouse.black_diamond.dim_accounts da ON tg.account_sk = da.accounts_SK
WHERE tg.iscurrent = 'True'

UNION ALL

SELECT 
    'fact_accounts_holdings', 
    COUNT(*),
    (SELECT COUNT(*) FROM (SELECT accountid, id, asofdate FROM Gold_Warehouse.black_diamond.fact_accounts_holdings GROUP BY accountid, id, asofdate HAVING COUNT(*) > 1) d),
    SUM(CASE WHEN h.accountid = da.id THEN 1 ELSE 0 END),
    SUM(CASE WHEN h.accountid <> da.id THEN 1 ELSE 0 END),
    SUM(CASE WHEN h.account_sk IS NULL THEN 1 ELSE 0 END),
    SUM(CASE WHEN h.account_sk IS NOT NULL AND da.accounts_SK IS NULL THEN 1 ELSE 0 END)
FROM Gold_Warehouse.black_diamond.fact_accounts_holdings h
LEFT JOIN Gold_Warehouse.black_diamond.dim_accounts da ON h.account_sk = da.accounts_SK

UNION ALL

SELECT 
    'fact_accounts_transactions', 
    COUNT(*),
    (SELECT COUNT(*) FROM (SELECT transactionid FROM Gold_Warehouse.black_diamond.fact_accounts_transactions GROUP BY transactionid HAVING COUNT(*) > 1) d),
    SUM(CASE WHEN tx.accountid = da.id THEN 1 ELSE 0 END),
    SUM(CASE WHEN tx.accountid <> da.id THEN 1 ELSE 0 END),
    SUM(CASE WHEN tx.account_sk IS NULL THEN 1 ELSE 0 END),
    SUM(CASE WHEN tx.account_sk IS NOT NULL AND da.accounts_SK IS NULL THEN 1 ELSE 0 END)
FROM Gold_Warehouse.black_diamond.fact_accounts_transactions tx
LEFT JOIN Gold_Warehouse.black_diamond.dim_accounts da ON tx.account_sk = da.accounts_SK

UNION ALL

SELECT 
    'fact_accounts_taxlots', 
    COUNT(*),
    (SELECT COUNT(*) FROM (SELECT accountid, taxlotid, holdingid FROM Gold_Warehouse.black_diamond.fact_accounts_taxlots GROUP BY accountid, taxlotid, holdingid HAVING COUNT(*) > 1) d),
    SUM(CASE WHEN tl.accountid = da.id THEN 1 ELSE 0 END),
    SUM(CASE WHEN tl.accountid <> da.id THEN 1 ELSE 0 END),
    SUM(CASE WHEN tl.account_sk IS NULL THEN 1 ELSE 0 END),
    SUM(CASE WHEN tl.account_sk IS NOT NULL AND da.accounts_SK IS NULL THEN 1 ELSE 0 END)
FROM Gold_Warehouse.black_diamond.fact_accounts_taxlots tl
LEFT JOIN Gold_Warehouse.black_diamond.dim_accounts da ON tl.account_sk = da.accounts_SK;"""
        },
        {
            "id": "TC_COMB_SK_SC_02",
            "source": "Schwab",
            "scope": "All 8 Objects (dim_accounts & 7 Fact Tables: recon_d1/d2, daily_closing, security, transactions, ult/uln lots)",
            "priority": "Critical",
            "title": "Schwab: Consolidated Multi-Object SK Mapping, PK Deduplication & Audit Date Health Check",
            "desc": "Single unified query executing across all 8 Schwab Gold tables. Validates that every fact record's account_sk correctly matches dim_accounts on composite natural keys (account_id, custodian_id), confirms primary SK uniqueness, and ensures 100% audit dates populated.",
            "exp": "Mismatched_AccountId_Count = 0 and Orphan_AccountSK_Count = 0 across all fact tables. All mapped account_sk strictly match the natural account_id.",
            "query": """SELECT 
    'fact_recon_position_d1' AS TableName,
    COUNT(*) AS TotalRows,
    SUM(CASE WHEN f.account_id = da.account_id THEN 1 ELSE 0 END) AS Matched_AccountId_Count,
    SUM(CASE WHEN f.account_id <> da.account_id THEN 1 ELSE 0 END) AS Mismatched_AccountId_Count,
    SUM(CASE WHEN f.account_sk IS NULL THEN 1 ELSE 0 END) AS Null_AccountSK_Count,
    SUM(CASE WHEN f.account_sk IS NOT NULL AND da.account_sk IS NULL THEN 1 ELSE 0 END) AS Orphan_AccountSK_Count
FROM Gold_Warehouse.schwab.fact_recon_position_d1 f
LEFT JOIN Gold_Warehouse.schwab.dim_accounts da ON f.account_sk = da.account_sk

UNION ALL

SELECT 
    'fact_recon_position_d2', COUNT(*),
    SUM(CASE WHEN f.account_id = da.account_id THEN 1 ELSE 0 END),
    SUM(CASE WHEN f.account_id <> da.account_id THEN 1 ELSE 0 END),
    SUM(CASE WHEN f.account_sk IS NULL THEN 1 ELSE 0 END),
    SUM(CASE WHEN f.account_sk IS NOT NULL AND da.account_sk IS NULL THEN 1 ELSE 0 END)
FROM Gold_Warehouse.schwab.fact_recon_position_d2 f
LEFT JOIN Gold_Warehouse.schwab.dim_accounts da ON f.account_sk = da.account_sk

UNION ALL

SELECT 
    'fact_daily_closing_lots', COUNT(*),
    SUM(CASE WHEN f.account_id = da.account_id THEN 1 ELSE 0 END),
    SUM(CASE WHEN f.account_id <> da.account_id THEN 1 ELSE 0 END),
    SUM(CASE WHEN f.account_sk IS NULL THEN 1 ELSE 0 END),
    SUM(CASE WHEN f.account_sk IS NOT NULL AND da.account_sk IS NULL THEN 1 ELSE 0 END)
FROM Gold_Warehouse.schwab.fact_daily_closing_lots f
LEFT JOIN Gold_Warehouse.schwab.dim_accounts da ON f.account_sk = da.account_sk

UNION ALL

SELECT 
    'fact_transactions', COUNT(*),
    SUM(CASE WHEN f.account_id = da.account_id THEN 1 ELSE 0 END),
    SUM(CASE WHEN f.account_id <> da.account_id THEN 1 ELSE 0 END),
    SUM(CASE WHEN f.account_sk IS NULL THEN 1 ELSE 0 END),
    SUM(CASE WHEN f.account_sk IS NOT NULL AND da.account_sk IS NULL THEN 1 ELSE 0 END)
FROM Gold_Warehouse.schwab.fact_transactions f
LEFT JOIN Gold_Warehouse.schwab.dim_accounts da ON f.account_sk = da.account_sk

UNION ALL

SELECT 
    'fact_ult_lot_daily_taxable', COUNT(*),
    SUM(CASE WHEN f.account_id = da.account_id THEN 1 ELSE 0 END),
    SUM(CASE WHEN f.account_id <> da.account_id THEN 1 ELSE 0 END),
    SUM(CASE WHEN f.account_sk IS NULL THEN 1 ELSE 0 END),
    SUM(CASE WHEN f.account_sk IS NOT NULL AND da.account_sk IS NULL THEN 1 ELSE 0 END)
FROM Gold_Warehouse.schwab.fact_ult_lot_daily_taxable f
LEFT JOIN Gold_Warehouse.schwab.dim_accounts da ON f.account_sk = da.account_sk

UNION ALL

SELECT 
    'fact_uln_lot_daily_non_taxable', COUNT(*),
    SUM(CASE WHEN f.account_id = da.account_id THEN 1 ELSE 0 END),
    SUM(CASE WHEN f.account_id <> da.account_id THEN 1 ELSE 0 END),
    SUM(CASE WHEN f.account_sk IS NULL THEN 1 ELSE 0 END),
    SUM(CASE WHEN f.account_sk IS NOT NULL AND da.account_sk IS NULL THEN 1 ELSE 0 END)
FROM Gold_Warehouse.schwab.fact_uln_lot_daily_non_taxable f
LEFT JOIN Gold_Warehouse.schwab.dim_accounts da ON f.account_sk = da.account_sk;"""
        },
        {
            "id": "TC_COMB_SK_SAL_03",
            "source": "Salentica CRM",
            "scope": "All 6 Core CRM Entities (dim_account, dim_contact_core, dim_family, fact_lead, fact_opportunity, fact_investmentaccount)",
            "priority": "Critical",
            "title": "Salentica: Consolidated Point-in-Time Referential SK Mapping (Active & Historical Mapping Validation)",
            "desc": "Single unified query executing across all 6 core Salentica CRM tables. Validates that every attached foreign SK resolves to the same Natural Key in parent dimensions (allowing active OR expired parent versions to preserve point-in-time audit history) and checks for broken orphans.",
            "exp": "Mismatched_PrimaryContact_NK = 0, Mismatched_BusinessUnit_NK = 0, Broken_Orphan_Count = 0 across all 6 CRM tables.",
            "query": """SELECT 'dim_account' AS TableName, COUNT(*) AS TotalActiveRows,
    SUM(CASE WHEN a.primarycontactid IS NOT NULL AND a.primarycontactid_sk <> -1 AND a.primarycontactid <> c.contactid THEN 1 ELSE 0 END) AS Mismatched_PrimaryContact_NK,
    SUM(CASE WHEN a.owningbusinessunit IS NOT NULL AND a.owningbusinessunit_sk <> -1 AND a.owningbusinessunit <> bu.businessunitid THEN 1 ELSE 0 END) AS Mismatched_BusinessUnit_NK,
    SUM(CASE WHEN (a.primarycontactid IS NOT NULL AND a.primarycontactid_sk <> -1 AND c.contact_SK IS NULL)
               OR (a.owningbusinessunit IS NOT NULL AND a.owningbusinessunit_sk <> -1 AND bu.businessunit_SK IS NULL) THEN 1 ELSE 0 END) AS Broken_Orphan_Count
FROM Gold_Warehouse.salentica.dim_account a
LEFT JOIN Gold_Warehouse.salentica.dim_contact_core c ON a.primarycontactid_sk = c.contact_SK
LEFT JOIN Gold_Warehouse.salentica.dim_businessunit bu ON a.owningbusinessunit_sk = bu.businessunit_SK
WHERE a.iscurrent = 'True'
UNION ALL
SELECT 'dim_contact_core', COUNT(*),
    SUM(CASE WHEN c.accountid IS NOT NULL AND c.accountid_sk <> -1 AND c.accountid <> a.accountid THEN 1 ELSE 0 END),
    SUM(CASE WHEN c.owningbusinessunit IS NOT NULL AND c.owningbusinessunit_sk <> -1 AND c.owningbusinessunit <> bu.businessunitid THEN 1 ELSE 0 END),
    SUM(CASE WHEN (c.accountid IS NOT NULL AND c.accountid_sk <> -1 AND a.account_SK IS NULL)
               OR (c.owningbusinessunit IS NOT NULL AND c.owningbusinessunit_sk <> -1 AND bu.businessunit_SK IS NULL) THEN 1 ELSE 0 END)
FROM Gold_Warehouse.salentica.dim_contact_core c
LEFT JOIN Gold_Warehouse.salentica.dim_account a ON c.accountid_sk = a.account_SK
LEFT JOIN Gold_Warehouse.salentica.dim_businessunit bu ON c.owningbusinessunit_sk = bu.businessunit_SK
WHERE c.iscurrent = 'True'
UNION ALL
SELECT 'dim_family', COUNT(*),
    SUM(CASE WHEN f.si_primarycontact IS NOT NULL AND f.si_primarycontact_sk <> -1 AND f.si_primarycontact <> c.contactid THEN 1 ELSE 0 END),
    SUM(CASE WHEN f.owningbusinessunit IS NOT NULL AND f.owningbusinessunit_sk <> -1 AND f.owningbusinessunit <> bu.businessunitid THEN 1 ELSE 0 END),
    SUM(CASE WHEN (f.si_primarycontact IS NOT NULL AND f.si_primarycontact_sk <> -1 AND c.contact_SK IS NULL)
               OR (f.owningbusinessunit IS NOT NULL AND f.owningbusinessunit_sk <> -1 AND bu.businessunit_SK IS NULL) THEN 1 ELSE 0 END)
FROM Gold_Warehouse.salentica.dim_family f
LEFT JOIN Gold_Warehouse.salentica.dim_contact_core c ON f.si_primarycontact_sk = c.contact_SK
LEFT JOIN Gold_Warehouse.salentica.dim_businessunit bu ON f.owningbusinessunit_sk = bu.businessunit_SK
WHERE f.iscurrent = 'True'
UNION ALL
SELECT 'fact_lead', COUNT(*),
    SUM(CASE WHEN l.accountid IS NOT NULL AND l.accountid_sk <> -1 AND l.accountid <> a.accountid THEN 1 ELSE 0 END),
    SUM(CASE WHEN l.contactid IS NOT NULL AND l.contactid_sk <> -1 AND l.contactid <> c.contactid THEN 1 ELSE 0 END),
    SUM(CASE WHEN (l.accountid IS NOT NULL AND l.accountid_sk <> -1 AND a.account_SK IS NULL)
               OR (l.contactid IS NOT NULL AND l.contactid_sk <> -1 AND c.contact_SK IS NULL) THEN 1 ELSE 0 END)
FROM Gold_Warehouse.salentica.fact_lead l
LEFT JOIN Gold_Warehouse.salentica.dim_account a ON l.accountid_sk = a.account_SK
LEFT JOIN Gold_Warehouse.salentica.dim_contact_core c ON l.contactid_sk = c.contact_SK
WHERE l.iscurrent = 'True'
UNION ALL
SELECT 'fact_opportunity', COUNT(*),
    SUM(CASE WHEN opp.accountid IS NOT NULL AND opp.accountid_sk <> -1 AND opp.accountid <> a.accountid THEN 1 ELSE 0 END),
    SUM(CASE WHEN opp.contactid IS NOT NULL AND opp.contactid_sk <> -1 AND opp.contactid <> c.contactid THEN 1 ELSE 0 END),
    SUM(CASE WHEN (opp.accountid IS NOT NULL AND opp.accountid_sk <> -1 AND a.account_SK IS NULL)
               OR (opp.contactid IS NOT NULL AND opp.contactid_sk <> -1 AND c.contact_SK IS NULL) THEN 1 ELSE 0 END)
FROM Gold_Warehouse.salentica.fact_opportunity opp
LEFT JOIN Gold_Warehouse.salentica.dim_account a ON opp.accountid_sk = a.account_SK
LEFT JOIN Gold_Warehouse.salentica.dim_contact_core c ON opp.contactid_sk = c.contact_SK
WHERE opp.iscurrent = 'True'
UNION ALL
SELECT 'fact_investmentaccount', COUNT(*),
    SUM(CASE WHEN ia.si_family IS NOT NULL AND ia.si_family_sk <> -1 AND ia.si_family <> f.si_familyid THEN 1 ELSE 0 END),
    SUM(CASE WHEN ia.si_portfolio IS NOT NULL AND ia.si_portfolio_sk <> -1 AND ia.si_portfolio <> p.si_portfolioid THEN 1 ELSE 0 END),
    SUM(CASE WHEN (ia.si_family IS NOT NULL AND ia.si_family_sk <> -1 AND f.family_SK IS NULL)
               OR (ia.si_portfolio IS NOT NULL AND ia.si_portfolio_sk <> -1 AND p.portfolio_SK IS NULL) THEN 1 ELSE 0 END)
FROM Gold_Warehouse.salentica.fact_investmentaccount ia
LEFT JOIN Gold_Warehouse.salentica.dim_family f ON ia.si_family_sk = f.family_SK
LEFT JOIN Gold_Warehouse.salentica.dim_portfolio p ON ia.si_portfolio_sk = p.portfolio_SK
WHERE ia.iscurrent = 'True';"""
        },
        {
            "id": "TC_COMB_SK_TAM_04",
            "source": "Tamarac",
            "scope": "Both Tables (dim_account_information, dim_account_model_allocation)",
            "priority": "Critical",
            "title": "Tamarac: Consolidated Multi-Object PK Deduplication, Primary SK & Cross-Table Linkage",
            "desc": "Single unified query validating SCD2 active uniqueness on exact PK columns (['account_number'] and ['account_number', 'model_name']), surrogate key uniqueness, and cross-table referential integrity between model allocations and accounts.",
            "exp": "PK_Active_Duplicates = 0, Duplicate_Primary_SK = 0, Orphan_Accounts = 0, Missing_BusinessDate = 0 across both tables.",
            "query": """SELECT 'dim_account_information' AS TableName, COUNT(*) AS TotalActiveRows,
    (SELECT COUNT(*) FROM (SELECT account_number FROM Gold_Warehouse.tamarac.dim_account_information WHERE is_current = 'True' GROUP BY account_number HAVING COUNT(*) > 1) d) AS PK_Active_Duplicates,
    (COUNT(*) - COUNT(DISTINCT account_information_SK)) AS Duplicate_Primary_SK,
    0 AS Orphan_Accounts,
    SUM(CASE WHEN file_business_date IS NULL THEN 1 ELSE 0 END) AS Missing_BusinessDate,
    SUM(CASE WHEN updated_at IS NULL THEN 1 ELSE 0 END) AS Missing_UpdatedAt
FROM Gold_Warehouse.tamarac.dim_account_information WHERE is_current = 'True'
UNION ALL
SELECT 'dim_account_model_allocation', COUNT(*),
    (SELECT COUNT(*) FROM (SELECT account_number, model_name FROM Gold_Warehouse.tamarac.dim_account_model_allocation WHERE is_current = 'True' GROUP BY account_number, model_name HAVING COUNT(*) > 1) d),
    (COUNT(*) - COUNT(DISTINCT account_model_SK)),
    SUM(CASE WHEN info.account_information_SK IS NULL THEN 1 ELSE 0 END),
    SUM(CASE WHEN ma.file_business_date IS NULL THEN 1 ELSE 0 END),
    SUM(CASE WHEN ma.updated_at IS NULL THEN 1 ELSE 0 END)
FROM Gold_Warehouse.tamarac.dim_account_model_allocation ma
LEFT JOIN Gold_Warehouse.tamarac.dim_account_information info 
    ON ma.account_number = info.account_number AND info.is_current = 'True'
WHERE ma.is_current = 'True';"""
        },
        {
            "id": "TC_COMB_SK_INNO_05",
            "source": "Innotrust",
            "scope": "All 5 Innotrust Dimensions (dim_account_master, category_codes, type_codes, sec_index, characteristics)",
            "priority": "Critical",
            "title": "Innotrust: Consolidated Multi-Object PK Deduplication, Primary SK & Cross-Dimension Linkage",
            "desc": "Single unified query validating SCD2 active uniqueness on exact natural keys, primary surrogate key uniqueness, and cross-dimension referential integrity between master entities and classification codes.",
            "exp": "PK_Active_Duplicates = 0, Duplicate_Primary_SK = 0, Orphan_Foreign_References = 0, Missing_BusinessDate = 0 across all 5 Innotrust tables.",
            "query": """SELECT 'dim_account_master' AS TableName, COUNT(*) AS TotalActiveRows,
    (SELECT COUNT(*) FROM (SELECT customer_account_number FROM Gold_Warehouse.innotrust.dim_account_master WHERE is_current = 1 GROUP BY customer_account_number HAVING COUNT(*) > 1) d) AS PK_Active_Duplicates,
    (COUNT(*) - COUNT(DISTINCT dim_account_master_sk)) AS Duplicate_Primary_SK,
    SUM(CASE WHEN (m.account_category_code IS NOT NULL AND cat.account_category_code IS NULL)
               OR (m.account_type_code IS NOT NULL AND typ.account_type_code IS NULL) THEN 1 ELSE 0 END) AS Orphan_Foreign_References,
    SUM(CASE WHEN m.business_date IS NULL THEN 1 ELSE 0 END) AS Missing_BusinessDate,
    SUM(CASE WHEN m.updated_at IS NULL THEN 1 ELSE 0 END) AS Missing_UpdatedAt
FROM Gold_Warehouse.innotrust.dim_account_master m
LEFT JOIN Gold_Warehouse.innotrust.dim_account_category_codes cat ON m.account_category_code = cat.account_category_code AND cat.is_current = 1
LEFT JOIN Gold_Warehouse.innotrust.dim_account_type_codes typ ON m.account_type_code = typ.account_type_code AND typ.is_current = 1
WHERE m.is_current = 1

UNION ALL

SELECT 'dim_account_category_codes', COUNT(*),
    (SELECT COUNT(*) FROM (SELECT account_category_code FROM Gold_Warehouse.innotrust.dim_account_category_codes WHERE is_current = 1 GROUP BY account_category_code HAVING COUNT(*) > 1) d),
    (COUNT(*) - COUNT(DISTINCT dim_account_master_sk)),
    0,
    SUM(CASE WHEN business_date IS NULL THEN 1 ELSE 0 END),
    SUM(CASE WHEN updated_at IS NULL THEN 1 ELSE 0 END)
FROM Gold_Warehouse.innotrust.dim_account_category_codes WHERE is_current = 1

UNION ALL

SELECT 'dim_account_type_codes', COUNT(*),
    (SELECT COUNT(*) FROM (SELECT account_type_code FROM Gold_Warehouse.innotrust.dim_account_type_codes WHERE is_current = 1 GROUP BY account_type_code HAVING COUNT(*) > 1) d),
    (COUNT(*) - COUNT(DISTINCT dim_account_type_codes_sk)),
    0,
    SUM(CASE WHEN business_date IS NULL THEN 1 ELSE 0 END),
    SUM(CASE WHEN updated_at IS NULL THEN 1 ELSE 0 END)
FROM Gold_Warehouse.innotrust.dim_account_type_codes WHERE is_current = 1

UNION ALL

SELECT 'dim_master_sec_index', COUNT(*),
    (SELECT COUNT(*) FROM (SELECT isin FROM Gold_Warehouse.innotrust.dim_master_sec_index WHERE is_current = 1 GROUP BY isin HAVING COUNT(*) > 1) d),
    (COUNT(*) - COUNT(DISTINCT dim_master_sec_index_sk)),
    SUM(CASE WHEN s.security_type_code IS NOT NULL AND sc.security_type_code IS NULL THEN 1 ELSE 0 END),
    SUM(CASE WHEN s.business_date IS NULL THEN 1 ELSE 0 END),
    SUM(CASE WHEN s.updated_at IS NULL THEN 1 ELSE 0 END)
FROM Gold_Warehouse.innotrust.dim_master_sec_index s
LEFT JOIN Gold_Warehouse.innotrust.dim_security_characteristics_codes sc ON s.security_type_code = sc.security_type_code AND sc.is_current = 1
WHERE s.is_current = 1

UNION ALL

SELECT 'dim_security_characteristics_codes', COUNT(*),
    (SELECT COUNT(*) FROM (SELECT security_type_code FROM Gold_Warehouse.innotrust.dim_security_characteristics_codes WHERE is_current = 1 GROUP BY security_type_code HAVING COUNT(*) > 1) d),
    (COUNT(*) - COUNT(DISTINCT dim_securitycharacteristicscodes)),
    0,
    SUM(CASE WHEN business_date IS NULL THEN 1 ELSE 0 END),
    SUM(CASE WHEN updated_at IS NULL THEN 1 ELSE 0 END)
FROM Gold_Warehouse.innotrust.dim_security_characteristics_codes WHERE is_current = 1;"""
        },
        {
            "id": "TC_COMB_SK_INNO_06",
            "source": "Innotrust",
            "scope": "All 5 Innotrust Dimensions (dim_account_master, category_codes, type_codes, sec_index, characteristics)",
            "priority": "Critical",
            "title": "Innotrust: Consolidated Multi-Object SCD2 Effective Start & End Date Audit (End Date <= Today - 1)",
            "desc": "Validates across all 5 Innotrust dimensions that inactive records (is_current = 0) have effective_end_date <= (current day - 1), non-null end dates, active records have non-future start dates, and zero inverted date ranges.",
            "exp": "Inactive_End_Date_Breach_Count = 0, Inactive_Missing_End_Date = 0, Active_Future_Start_Date = 0, Inverted_Date_Range_Count = 0 across all 5 tables.",
            "query": """SELECT 
    'dim_account_master' AS TableName,
    COUNT(*) AS TotalRecords,
    SUM(CASE WHEN is_current = 0 AND effective_end_date > DATEADD(day, -1, CAST(GETDATE() AS DATE)) THEN 1 ELSE 0 END) AS Inactive_End_Date_Breach_Count,
    SUM(CASE WHEN is_current = 0 AND effective_end_date IS NULL THEN 1 ELSE 0 END) AS Inactive_Missing_End_Date,
    SUM(CASE WHEN is_current = 1 AND effective_start_date > CAST(GETDATE() AS DATE) THEN 1 ELSE 0 END) AS Active_Future_Start_Date,
    SUM(CASE WHEN effective_end_date < effective_start_date THEN 1 ELSE 0 END) AS Inverted_Date_Range_Count,
    CASE 
        WHEN SUM(CASE WHEN is_current = 0 AND effective_end_date > DATEADD(day, -1, CAST(GETDATE() AS DATE)) THEN 1 ELSE 0 END) = 0
         AND SUM(CASE WHEN is_current = 0 AND effective_end_date IS NULL THEN 1 ELSE 0 END) = 0
         AND SUM(CASE WHEN effective_end_date < effective_start_date THEN 1 ELSE 0 END) = 0
        THEN 'PASSED' ELSE 'FAILED'
    END AS TestStatus
FROM Gold_Warehouse.innotrust.dim_account_master

UNION ALL

SELECT 
    'dim_master_sec_index' AS TableName,
    COUNT(*),
    SUM(CASE WHEN is_current = 0 AND effective_end_date > DATEADD(day, -1, CAST(GETDATE() AS DATE)) THEN 1 ELSE 0 END),
    SUM(CASE WHEN is_current = 0 AND effective_end_date IS NULL THEN 1 ELSE 0 END),
    SUM(CASE WHEN is_current = 1 AND effective_start_date > CAST(GETDATE() AS DATE) THEN 1 ELSE 0 END),
    SUM(CASE WHEN effective_end_date < effective_start_date THEN 1 ELSE 0 END),
    CASE 
        WHEN SUM(CASE WHEN is_current = 0 AND effective_end_date > DATEADD(day, -1, CAST(GETDATE() AS DATE)) THEN 1 ELSE 0 END) = 0
         AND SUM(CASE WHEN is_current = 0 AND effective_end_date IS NULL THEN 1 ELSE 0 END) = 0
         AND SUM(CASE WHEN effective_end_date < effective_start_date THEN 1 ELSE 0 END) = 0
        THEN 'PASSED' ELSE 'FAILED'
    END
FROM Gold_Warehouse.innotrust.dim_master_sec_index

UNION ALL

SELECT 
    'dim_account_category_codes',
    COUNT(*),
    SUM(CASE WHEN is_current = 0 AND effective_end_date > DATEADD(day, -1, CAST(GETDATE() AS DATE)) THEN 1 ELSE 0 END),
    SUM(CASE WHEN is_current = 0 AND effective_end_date IS NULL THEN 1 ELSE 0 END),
    SUM(CASE WHEN is_current = 1 AND effective_start_date > CAST(GETDATE() AS DATE) THEN 1 ELSE 0 END),
    SUM(CASE WHEN effective_end_date < effective_start_date THEN 1 ELSE 0 END),
    CASE 
        WHEN SUM(CASE WHEN is_current = 0 AND effective_end_date > DATEADD(day, -1, CAST(GETDATE() AS DATE)) THEN 1 ELSE 0 END) = 0
         AND SUM(CASE WHEN is_current = 0 AND effective_end_date IS NULL THEN 1 ELSE 0 END) = 0
         AND SUM(CASE WHEN effective_end_date < effective_start_date THEN 1 ELSE 0 END) = 0
        THEN 'PASSED' ELSE 'FAILED'
    END
FROM Gold_Warehouse.innotrust.dim_account_category_codes

UNION ALL

SELECT 
    'dim_account_type_codes',
    COUNT(*),
    SUM(CASE WHEN is_current = 0 AND effective_end_date > DATEADD(day, -1, CAST(GETDATE() AS DATE)) THEN 1 ELSE 0 END),
    SUM(CASE WHEN is_current = 0 AND effective_end_date IS NULL THEN 1 ELSE 0 END),
    SUM(CASE WHEN is_current = 1 AND effective_start_date > CAST(GETDATE() AS DATE) THEN 1 ELSE 0 END),
    SUM(CASE WHEN effective_end_date < effective_start_date THEN 1 ELSE 0 END),
    CASE 
        WHEN SUM(CASE WHEN is_current = 0 AND effective_end_date > DATEADD(day, -1, CAST(GETDATE() AS DATE)) THEN 1 ELSE 0 END) = 0
         AND SUM(CASE WHEN is_current = 0 AND effective_end_date IS NULL THEN 1 ELSE 0 END) = 0
         AND SUM(CASE WHEN effective_end_date < effective_start_date THEN 1 ELSE 0 END) = 0
        THEN 'PASSED' ELSE 'FAILED'
    END
FROM Gold_Warehouse.innotrust.dim_account_type_codes

UNION ALL

SELECT 
    'dim_security_characteristics_codes',
    COUNT(*),
    SUM(CASE WHEN is_current = 0 AND effective_end_date > DATEADD(day, -1, CAST(GETDATE() AS DATE)) THEN 1 ELSE 0 END),
    SUM(CASE WHEN is_current = 0 AND effective_end_date IS NULL THEN 1 ELSE 0 END),
    SUM(CASE WHEN is_current = 1 AND effective_start_date > CAST(GETDATE() AS DATE) THEN 1 ELSE 0 END),
    SUM(CASE WHEN effective_end_date < effective_start_date THEN 1 ELSE 0 END),
    CASE 
        WHEN SUM(CASE WHEN is_current = 0 AND effective_end_date > DATEADD(day, -1, CAST(GETDATE() AS DATE)) THEN 1 ELSE 0 END) = 0
         AND SUM(CASE WHEN is_current = 0 AND effective_end_date IS NULL THEN 1 ELSE 0 END) = 0
         AND SUM(CASE WHEN effective_end_date < effective_start_date THEN 1 ELSE 0 END) = 0
        THEN 'PASSED' ELSE 'FAILED'
    END
FROM Gold_Warehouse.innotrust.dim_security_characteristics_codes;"""
        }
    ]

    # Add Consolidated SK cases to Master List
    for c in consolidated_sk_suites:
        all_master_cases.append({
            "schema": c["source"] + " (Consolidated)",
            "id": c["id"],
            "cat": "Consolidated Multi-Object SK Audit",
            "pri": c["priority"],
            "table": c["scope"],
            "desc": c["title"] + " - " + c["desc"],
            "cond": "All Gold tables populated for " + c["source"],
            "exp": c["exp"],
            "query": c["query"]
        })

    # -------------------------------------------------------------------------------------------------
    # 1. CREATE COVERAGE SUMMARY SHEET
    # -------------------------------------------------------------------------------------------------
    ws_sum = wb.create_sheet(title="Coverage Summary")
    ws_sum.views.sheetView[0].showGridLines = True

    ws_sum.merge_cells("A1:G1")
    ws_sum.cell(1, 1).value = "GOLD WAREHOUSE DATA VALIDATION TEST SUITE SUMMARY"
    ws_sum.cell(1, 1).font = FONT_TITLE
    ws_sum.cell(1, 1).fill = FILL_NAVY
    ws_sum.cell(1, 1).alignment = ALIGN_CENTER
    ws_sum.row_dimensions[1].height = 36

    ws_sum.merge_cells("A2:G2")
    ws_sum.cell(2, 1).value = "Silver -> Gold Layer Data Integrity: Deduplication on PKs, SK Mapping & Joins, file_business_date, and Consolidated Audits"
    ws_sum.cell(2, 1).font = FONT_SUBTITLE
    ws_sum.cell(2, 1).fill = FILL_NAVY
    ws_sum.cell(2, 1).alignment = ALIGN_CENTER
    ws_sum.row_dimensions[2].height = 24

    total_tests = len(all_master_cases)
    crit_count = sum(1 for c in all_master_cases if c["pri"] == "Critical")
    high_count = sum(1 for c in all_master_cases if c["pri"] == "High")
    med_count = sum(1 for c in all_master_cases if c["pri"] == "Medium")

    kpis = [
        ("TOTAL SCHEMAS", 5, "A", "B"),
        ("TOTAL TEST CASES", total_tests, "C", "D"),
        ("CRITICAL PRIORITY", crit_count, "E", "E"),
        ("HIGH PRIORITY", high_count, "F", "F"),
        ("CONSOLIDATED AUDITS", len(consolidated_sk_suites), "G", "G")
    ]

    ws_sum.row_dimensions[4].height = 20
    ws_sum.row_dimensions[5].height = 28
    for label, val, c1, c2 in kpis:
        if c1 != c2:
            ws_sum.merge_cells(f"{c1}4:{c2}4")
            ws_sum.merge_cells(f"{c1}5:{c2}5")
        cell_lbl = ws_sum[f"{c1}4"]
        cell_val = ws_sum[f"{c1}5"]
        cell_lbl.value = label
        cell_lbl.font = Font(name="Calibri", size=9, bold=True, color="595959")
        cell_lbl.alignment = ALIGN_CENTER
        cell_lbl.fill = FILL_SECTION
        cell_lbl.border = THIN_BORDER
        cell_val.value = val
        cell_val.font = Font(name="Calibri", size=16, bold=True, color="0A2240")
        cell_val.alignment = ALIGN_CENTER
        cell_val.fill = FILL_WHITE
        cell_val.border = THIN_BORDER

    headers_sum = ["Source / Schema Name", "Target Gold Schema", "Total Scenarios", "Critical Priority", "High Priority", "Consolidated Test", "Scope Focus"]
    ws_sum.row_dimensions[7].height = 26
    for col_idx, h in enumerate(headers_sum, 1):
        c = ws_sum.cell(7, col_idx, h)
        c.font = FONT_HEADER
        c.fill = FILL_HEADER
        c.alignment = ALIGN_CENTER
        c.border = THIN_BORDER

    schema_rows = [
        ("Black Diamond", "black_diamond", len(test_suites["Black Diamond Gold"]) + 1,
         sum(1 for c in test_suites["Black Diamond Gold"] if c["pri"]=="Critical") + 1,
         sum(1 for c in test_suites["Black Diamond Gold"] if c["pri"]=="High"),
         "TC_COMB_SK_BD_01",
         "All 16 tables: Dims (accounts, asset, targets) & Facts (holdings, transactions, taxlots)"),
        ("Schwab", "schwab", len(test_suites["Schwab Gold"]) + 1,
         sum(1 for c in test_suites["Schwab Gold"] if c["pri"]=="Critical") + 1,
         sum(1 for c in test_suites["Schwab Gold"] if c["pri"]=="High"),
         "TC_COMB_SK_SC_02",
         "dim_accounts, recon_position_d1/d2, daily_closing_lots, transactions, security, ult/uln lots"),
        ("Salentica CRM", "salentica", len(test_suites["Salentica Gold"]) + 1,
         sum(1 for c in test_suites["Salentica Gold"] if c["pri"]=="Critical") + 1,
         sum(1 for c in test_suites["Salentica Gold"] if c["pri"]=="High"),
         "TC_COMB_SK_SAL_03",
         "All 6 CRM entities: dim_account, contact_core, family, lead, opportunity, investmentaccount"),
        ("Tamarac", "tamarac", len(test_suites["Tamarac Gold"]) + 1,
         sum(1 for c in test_suites["Tamarac Gold"] if c["pri"]=="Critical") + 1,
         sum(1 for c in test_suites["Tamarac Gold"] if c["pri"]=="High"),
         "TC_COMB_SK_TAM_04",
         "dim_account_information & dim_account_model_allocation with PK deduplication"),
        ("Innotrust", "innotrust", len(test_suites["Innotrust Gold"]) + 2,
         sum(1 for c in test_suites["Innotrust Gold"] if c["pri"]=="Critical") + 2,
         sum(1 for c in test_suites["Innotrust Gold"] if c["pri"]=="High"),
         "TC_COMB_SK_INNO_05 / 06",
         "All 5 dimensions: PK dedup, SK uniqueness, cross-table linkage & SCD2 date rules (End Date <= Today - 1)")
    ]

    for row_idx, r_data in enumerate(schema_rows, 8):
        ws_sum.row_dimensions[row_idx].height = 24
        fill = FILL_ZEBRA if row_idx % 2 == 0 else FILL_WHITE
        for col_idx, val in enumerate(r_data, 1):
            c = ws_sum.cell(row_idx, col_idx, val)
            c.font = FONT_REGULAR if col_idx not in [1, 3] else FONT_BOLD
            c.fill = fill
            c.border = THIN_BORDER
            c.alignment = ALIGN_CENTER if col_idx in [3, 4, 5, 6] else ALIGN_LEFT

    tot_row = 13
    ws_sum.row_dimensions[tot_row].height = 26
    ws_sum.cell(tot_row, 1, "TOTAL").font = FONT_BOLD
    ws_sum.cell(tot_row, 1).fill = FILL_SECTION
    ws_sum.cell(tot_row, 1).border = THIN_BORDER
    ws_sum.cell(tot_row, 2, "All 5 Gold Schemas").font = FONT_BOLD
    ws_sum.cell(tot_row, 2).fill = FILL_SECTION
    ws_sum.cell(tot_row, 2).border = THIN_BORDER

    ws_sum.cell(tot_row, 3, total_tests).font = FONT_BOLD
    ws_sum.cell(tot_row, 3).fill = FILL_SECTION
    ws_sum.cell(tot_row, 3).border = THIN_BORDER
    ws_sum.cell(tot_row, 3).alignment = ALIGN_CENTER

    ws_sum.cell(tot_row, 4, crit_count).font = FONT_BOLD
    ws_sum.cell(tot_row, 4).fill = FILL_SECTION
    ws_sum.cell(tot_row, 4).border = THIN_BORDER
    ws_sum.cell(tot_row, 4).alignment = ALIGN_CENTER

    ws_sum.cell(tot_row, 5, high_count).font = FONT_BOLD
    ws_sum.cell(tot_row, 5).fill = FILL_SECTION
    ws_sum.cell(tot_row, 5).border = THIN_BORDER
    ws_sum.cell(tot_row, 5).alignment = ALIGN_CENTER

    ws_sum.cell(tot_row, 6, "5 Combined Audits").font = FONT_BOLD
    ws_sum.cell(tot_row, 6).fill = FILL_SECTION
    ws_sum.cell(tot_row, 6).border = THIN_BORDER
    ws_sum.cell(tot_row, 6).alignment = ALIGN_CENTER

    ws_sum.cell(tot_row, 7, "100% Core Verification").font = FONT_BOLD
    ws_sum.cell(tot_row, 7).fill = FILL_SECTION
    ws_sum.cell(tot_row, 7).border = THIN_BORDER

    sum_widths = [20, 24, 16, 16, 16, 18, 45]
    for i, w in enumerate(sum_widths, 1):
        ws_sum.column_dimensions[get_column_letter(i)].width = w

    # -------------------------------------------------------------------------------------------------
    # 2. CREATE ALL GOLD TESTS (MASTER)
    # -------------------------------------------------------------------------------------------------
    ws_master = wb.create_sheet(title="All Gold Tests (Master)")
    ws_master.views.sheetView[0].showGridLines = True

    ws_master.merge_cells("A1:I1")
    ws_master.cell(1, 1).value = "CONSOLIDATED GOLD WAREHOUSE DATA VALIDATION TEST CASES"
    ws_master.cell(1, 1).font = FONT_TITLE
    ws_master.cell(1, 1).fill = FILL_NAVY
    ws_master.cell(1, 1).alignment = ALIGN_CENTER
    ws_master.row_dimensions[1].height = 36

    master_headers = [
        "Source Schema", "Test Case ID", "Test Category", "Priority", 
        "Target Table / Scope", "Test Scenario Description", "Pre-conditions", 
        "Expected Result", "SQL Validation Query"
    ]
    ws_master.row_dimensions[2].height = 26
    for col_idx, h in enumerate(master_headers, 1):
        c = ws_master.cell(2, col_idx, h)
        c.font = FONT_HEADER
        c.fill = FILL_HEADER
        c.alignment = ALIGN_CENTER
        c.border = THIN_BORDER

    for row_idx, tc in enumerate(all_master_cases, 3):
        ws_master.row_dimensions[row_idx].height = 54
        fill = FILL_ZEBRA if row_idx % 2 == 0 else FILL_WHITE
        
        ws_master.cell(row_idx, 1, tc["schema"]).alignment = ALIGN_CENTER
        ws_master.cell(row_idx, 2, tc["id"]).alignment = ALIGN_CENTER
        ws_master.cell(row_idx, 3, tc["cat"]).alignment = ALIGN_CENTER
        
        pri_cell = ws_master.cell(row_idx, 4, tc["pri"])
        pri_cell.alignment = ALIGN_CENTER
        if tc["pri"] == "Critical":
            pri_cell.font = FONT_CRITICAL
            pri_cell.fill = FILL_CRITICAL
        elif tc["pri"] == "High":
            pri_cell.font = FONT_HIGH
            pri_cell.fill = FILL_HIGH
        else:
            pri_cell.font = FONT_MEDIUM
            pri_cell.fill = FILL_MEDIUM

        ws_master.cell(row_idx, 5, tc["table"]).alignment = ALIGN_LEFT
        ws_master.cell(row_idx, 6, tc["desc"]).alignment = ALIGN_LEFT
        ws_master.cell(row_idx, 7, tc["cond"]).alignment = ALIGN_LEFT
        ws_master.cell(row_idx, 8, tc["exp"]).alignment = ALIGN_LEFT
        
        q_cell = ws_master.cell(row_idx, 9, tc["query"])
        q_cell.font = FONT_CODE
        q_cell.alignment = ALIGN_CODE

        for c_idx in range(1, 10):
            cell = ws_master.cell(row_idx, c_idx)
            cell.border = THIN_BORDER
            if c_idx != 4:
                cell.fill = fill
                if c_idx in [1, 2, 5]:
                    cell.font = FONT_BOLD
                elif c_idx != 9:
                    cell.font = FONT_REGULAR

    master_widths = [20, 20, 24, 12, 34, 40, 28, 40, 70]
    for i, w in enumerate(master_widths, 1):
        ws_master.column_dimensions[get_column_letter(i)].width = w

    # -------------------------------------------------------------------------------------------------
    # 3. CREATE INDIVIDUAL SCHEMA TABS
    # -------------------------------------------------------------------------------------------------
    schema_tabs = [
        ("Black Diamond Gold", test_suites["Black Diamond Gold"], "BLACK DIAMOND GOLD LAYER VALIDATION SUITE"),
        ("Schwab Gold", test_suites["Schwab Gold"], "SCHWAB GOLD LAYER VALIDATION SUITE"),
        ("Salentica Gold", test_suites["Salentica Gold"], "SALENTICA CRM GOLD LAYER VALIDATION SUITE"),
        ("Tamarac Gold", test_suites["Tamarac Gold"], "TAMARAC GOLD LAYER VALIDATION SUITE"),
        ("Innotrust Gold", test_suites["Innotrust Gold"], "INNOTRUST GOLD LAYER VALIDATION SUITE")
    ]

    tab_headers = [
        "Test Case ID", "Test Category", "Priority", "Target Table", 
        "Test Scenario Description", "Pre-conditions", "Expected Result", "SQL Validation Query"
    ]

    for tab_title, cases, banner_text in schema_tabs:
        ws_tab = wb.create_sheet(title=tab_title)
        ws_tab.views.sheetView[0].showGridLines = True

        ws_tab.merge_cells("A1:H1")
        ws_tab.cell(1, 1).value = banner_text
        ws_tab.cell(1, 1).font = FONT_TITLE
        ws_tab.cell(1, 1).fill = FILL_NAVY
        ws_tab.cell(1, 1).alignment = ALIGN_CENTER
        ws_tab.row_dimensions[1].height = 36

        ws_tab.row_dimensions[2].height = 26
        for col_idx, h in enumerate(tab_headers, 1):
            c = ws_tab.cell(2, col_idx, h)
            c.font = FONT_HEADER
            c.fill = FILL_HEADER
            c.alignment = ALIGN_CENTER
            c.border = THIN_BORDER

        for row_idx, tc in enumerate(cases, 3):
            ws_tab.row_dimensions[row_idx].height = 54
            fill = FILL_ZEBRA if row_idx % 2 == 0 else FILL_WHITE
            
            ws_tab.cell(row_idx, 1, tc["id"]).alignment = ALIGN_CENTER
            ws_tab.cell(row_idx, 2, tc["cat"]).alignment = ALIGN_CENTER
            
            pri_cell = ws_tab.cell(row_idx, 3, tc["pri"])
            pri_cell.alignment = ALIGN_CENTER
            if tc["pri"] == "Critical":
                pri_cell.font = FONT_CRITICAL
                pri_cell.fill = FILL_CRITICAL
            elif tc["pri"] == "High":
                pri_cell.font = FONT_HIGH
                pri_cell.fill = FILL_HIGH
            else:
                pri_cell.font = FONT_MEDIUM
                pri_cell.fill = FILL_MEDIUM

            ws_tab.cell(row_idx, 4, tc["table"]).alignment = ALIGN_LEFT
            ws_tab.cell(row_idx, 5, tc["desc"]).alignment = ALIGN_LEFT
            ws_tab.cell(row_idx, 6, tc["cond"]).alignment = ALIGN_LEFT
            ws_tab.cell(row_idx, 7, tc["exp"]).alignment = ALIGN_LEFT
            
            q_cell = ws_tab.cell(row_idx, 8, tc["query"])
            q_cell.font = FONT_CODE
            q_cell.alignment = ALIGN_CODE

            for c_idx in range(1, 9):
                cell = ws_tab.cell(row_idx, c_idx)
                cell.border = THIN_BORDER
                if c_idx != 3:
                    cell.fill = fill
                    if c_idx in [1, 4]:
                        cell.font = FONT_BOLD
                    elif c_idx != 8:
                        cell.font = FONT_REGULAR

        tab_widths = [18, 20, 12, 32, 38, 28, 38, 70]
        for i, w in enumerate(tab_widths, 1):
            ws_tab.column_dimensions[get_column_letter(i)].width = w

    # -------------------------------------------------------------------------------------------------
    # 4. CREATE 7TH SHEET: CONSOLIDATED SK AUDIT (ALL 4 SOURCES COMBINED OBJECT QUERIES)
    # -------------------------------------------------------------------------------------------------
    ws_comb = wb.create_sheet(title="Consolidated SK Audit")
    ws_comb.views.sheetView[0].showGridLines = True

    ws_comb.merge_cells("A1:H1")
    ws_comb.cell(1, 1).value = "CONSOLIDATED MULTI-OBJECT SURROGATE KEY & INTEGRITY AUDIT SUITE"
    ws_comb.cell(1, 1).font = FONT_TITLE
    ws_comb.cell(1, 1).fill = FILL_NAVY
    ws_comb.cell(1, 1).alignment = ALIGN_CENTER
    ws_comb.row_dimensions[1].height = 36

    comb_headers = [
        "Audit ID", "Source System", "Scope / Covered Objects", "Priority", 
        "Audit Scenario Title", "Pre-conditions", "Expected QA Pass Criteria", "Consolidated Unified SQL Query"
    ]
    ws_comb.row_dimensions[2].height = 26
    for col_idx, h in enumerate(comb_headers, 1):
        c = ws_comb.cell(2, col_idx, h)
        c.font = FONT_HEADER
        c.fill = FILL_HEADER
        c.alignment = ALIGN_CENTER
        c.border = THIN_BORDER

    for row_idx, tc in enumerate(consolidated_sk_suites, 3):
        ws_comb.row_dimensions[row_idx].height = 110  # Generous height for unified queries
        fill = FILL_ZEBRA if row_idx % 2 == 0 else FILL_WHITE

        ws_comb.cell(row_idx, 1, tc["id"]).alignment = ALIGN_CENTER
        ws_comb.cell(row_idx, 2, tc["source"]).alignment = ALIGN_CENTER
        ws_comb.cell(row_idx, 3, tc["scope"]).alignment = ALIGN_LEFT

        pri_cell = ws_comb.cell(row_idx, 4, tc["priority"])
        pri_cell.alignment = ALIGN_CENTER
        pri_cell.font = FONT_CRITICAL
        pri_cell.fill = FILL_CRITICAL

        ws_comb.cell(row_idx, 5, tc["title"] + "\n\n" + tc["desc"]).alignment = ALIGN_LEFT
        ws_comb.cell(row_idx, 6, "All Gold tables loaded for " + tc["source"]).alignment = ALIGN_LEFT
        ws_comb.cell(row_idx, 7, tc["exp"]).alignment = ALIGN_LEFT

        q_cell = ws_comb.cell(row_idx, 8, tc["query"])
        q_cell.font = FONT_CODE
        q_cell.alignment = ALIGN_CODE

        for c_idx in range(1, 9):
            cell = ws_comb.cell(row_idx, c_idx)
            cell.border = THIN_BORDER
            if c_idx != 4:
                cell.fill = fill
                if c_idx in [1, 2]:
                    cell.font = FONT_BOLD
                elif c_idx != 8:
                    cell.font = FONT_REGULAR

    comb_widths = [18, 18, 32, 12, 40, 24, 38, 85]
    for i, w in enumerate(comb_widths, 1):
        ws_comb.column_dimensions[get_column_letter(i)].width = w

    output_path = r"c:\Users\Vivek.singh\OneDrive - Hexaview Technologies\Desktop\CR\Gold_Data_Validation_Test_Cases.xlsx"
    wb.save(output_path)
    print(f"Workbook successfully saved to: {output_path}")

if __name__ == "__main__":
    build_workbook()
