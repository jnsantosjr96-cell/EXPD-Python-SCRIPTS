import os
from pathlib import Path
from datetime import date

import oracledb
import pandas as pd
from openpyxl import load_workbook

# Outlook (pywin32) for automatic email sending
try:
    import win32com.client as win32
except ImportError:
    win32 = None

# ==========================
# GENERAL CONFIG
# ==========================

DOWNLOADS_DIR = Path(os.path.join(os.path.expanduser("~"), "Downloads"))

CONCAT_PATH = Path(
    r"C:\Users\josenjr\OneDrive - Expedia Group\Desktop\Daily activities\COA vs WO Process\Concat.xlsx"
)

BASE_OUTPUT_DIR = Path(
    r"C:\Users\josenjr\OneDrive - Expedia Group\Desktop\Daily activities\COA vs WO Process"
)

today_str = date.today().strftime("%m.%d.%Y")  # folder name in MM.DD.YYYY format
OUTPUT_FOLDER = BASE_OUTPUT_DIR / today_str
OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

# RELO template (keeps macros)
RELO_DM_TEMPLATE = Path(
    r"C:\Users\josenjr\OneDrive - Expedia Group\Desktop\Daily activities\COA vs WO Process\4 Non-Lodging Relo CM.xlsm"
)

# ==========================
# ORACLE CONFIG
# ==========================

ORACLE_CLIENT_DIR = r"C:\Users\josenjr\OneDrive - Expedia Group\Desktop\Oracle Instant Client\instantclient_23_0"
USERNAME = "josenjr"
PASSWORD = "eyX057UWzLnZTl3w"      # CHANGE ME
DSN      = "ashworaebsdb02-vip.datawarehouse.expecn.com:1526/ORAPRD_UI"

# WO adjustments query
BASE_QUERY = """
SELECT
    CTA.TRX_NUMBER              AS "Transaction Number",
    CTTA.NAME                   AS "Transaction Type",
    HCA.ACCOUNT_NUMBER          AS "Account Number",
    CTA.INVOICE_CURRENCY_CODE   AS "Entered Currency",
    AAA.AMOUNT                  AS "Entered Amount",
    AAA.ADJUSTMENT_NUMBER       AS "Adjustment Number"
FROM AR.AR_ADJUSTMENTS_ALL      AAA
JOIN AR.RA_CUSTOMER_TRX_ALL     CTA  ON CTA.CUSTOMER_TRX_ID   = AAA.CUSTOMER_TRX_ID
JOIN AR.RA_CUST_TRX_TYPES_ALL   CTTA ON CTTA.CUST_TRX_TYPE_ID = CTA.CUST_TRX_TYPE_ID
JOIN AR.HZ_CUST_ACCOUNTS        HCA  ON HCA.CUST_ACCOUNT_ID   = CTA.BILL_TO_CUSTOMER_ID
JOIN APPS.HR_OPERATING_UNITS    HOU  ON HOU.ORGANIZATION_ID   = CTA.ORG_ID
JOIN APPS.GL_LEDGER_LE_V        GLL  ON GLL.LEGAL_ENTITY_ID   = HOU.DEFAULT_LEGAL_CONTEXT_ID
WHERE
    GLL.LEDGER_CATEGORY_CODE = 'PRIMARY'
    AND CTTA.NAME NOT LIKE '%EAC%'
    AND (
        INSTR(HOU.NAME, '11105') > 0
        OR INSTR(HOU.NAME, '11115') > 0
        OR INSTR(HOU.NAME, '12305') > 0
        OR INSTR(HOU.NAME, '14101') > 0
        OR INSTR(HOU.NAME, '23310') > 0
        OR INSTR(HOU.NAME, '72110') > 0
    )
    AND HCA.ACCOUNT_NUMBER IN ({placeholders})
ORDER BY
    HCA.ACCOUNT_NUMBER,
    CTA.TRX_NUMBER,
    AAA.ADJUSTMENT_NUMBER
"""

# RELO CMs query (with Reference)
RELO_QUERY = """
SELECT
    CTA.TRX_NUMBER                  AS "Transaction Number",
    CTTA.NAME                       AS "Transaction Type",
    HCA.ACCOUNT_NUMBER              AS "Account Number",
    CTA.INVOICE_CURRENCY_CODE       AS "Entered Currency",
    PSA.AMOUNT_DUE_ORIGINAL         AS "Entered Amount",
    PSA.AMOUNT_DUE_REMAINING        AS "Open Balance",
    CTA.REASON_CODE                 AS "Reason Code",
    CTA.TRX_DATE                    AS "Transaction Date",
    CTA.INTERFACE_HEADER_ATTRIBUTE1 AS "Reference"
FROM AR.RA_CUSTOMER_TRX_ALL      CTA
JOIN AR.RA_CUST_TRX_TYPES_ALL    CTTA ON CTTA.CUST_TRX_TYPE_ID = CTA.CUST_TRX_TYPE_ID
JOIN AR.HZ_CUST_ACCOUNTS         HCA  ON HCA.CUST_ACCOUNT_ID   = CTA.BILL_TO_CUSTOMER_ID
JOIN AR.AR_PAYMENT_SCHEDULES_ALL PSA  ON PSA.CUSTOMER_TRX_ID   = CTA.CUSTOMER_TRX_ID
JOIN APPS.HR_OPERATING_UNITS     HOU  ON HOU.ORGANIZATION_ID   = CTA.ORG_ID
JOIN APPS.GL_LEDGER_LE_V         GLL  ON GLL.LEGAL_ENTITY_ID   = HOU.DEFAULT_LEGAL_CONTEXT_ID
WHERE
    GLL.LEDGER_CATEGORY_CODE = 'PRIMARY'
    AND PSA.CLASS = 'CM'
    AND CTA.REASON_CODE LIKE '%OFFSET_ACCRUED_AR%'
    AND CTTA.NAME IN (
        'CH_DIR_CM_RELO_USD',
        'US_DIR_CM_RELO_USD',
        'TS_DIR_CM_RELO_USD',
        'BR_DIR_CM_RELO_BRL',
        'TS_CM_RELO'
    )
    AND HCA.ACCOUNT_NUMBER IN ({placeholders})
ORDER BY
    CTA.TRX_DATE DESC,
    CTA.TRX_NUMBER
"""

# ==========================
# HELPER FUNCTIONS
# ==========================

def find_latest_coa_file() -> Path:
    """Find the most recent 'RPA-306-001 Cash On Account*' file in Downloads."""
    candidates = []
    for f in DOWNLOADS_DIR.iterdir():
        if f.is_file() and f.name.startswith("RPA-306-001 Cash On Account"):
            candidates.append(f)
    if not candidates:
        raise FileNotFoundError("No 'RPA-306-001 Cash On Account*' file found in Downloads.")
    return max(candidates, key=lambda p: p.stat().st_mtime)

def chunk_list(lst, n):
    """Yield successive n-sized chunks from a list."""
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

def update_concat_column_b(customers):
    """Populate column B in Concat.xlsx with customer list and copy formulas down."""
    wb = load_workbook(CONCAT_PATH)
    ws = wb.active  # single sheet

    start_row = 2  # header
    for i, cust in enumerate(customers, start=start_row):
        ws.cell(row=i, column=2, value=cust)

    max_row = start_row + len(customers) - 1
    template_row = 2

    # copy formulas from row 2 downwards
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=template_row, column=col)
        if isinstance(cell.value, str) and cell.value.startswith("="):
            formula = cell.value
            for r in range(template_row + 1, max_row + 1):
                if ws.cell(row=r, column=col).value is None:
                    ws.cell(row=r, column=col).value = formula

    wb.save(CONCAT_PATH)
    print(f"Concat.xlsx updated (column B + formulas down to row {max_row}).")

def map_operating_unit(trx_type: str) -> str:
    """Map transaction type prefix to Operating Unit name."""
    if not isinstance(trx_type, str):
        return ""
    if trx_type.startswith("CH"):
        return "12305 Expedia Lodging Partner Services Sarl"
    if trx_type.startswith("US"):
        return "11105 Expedia, Inc."
    if trx_type.startswith("TS"):
        return "14101 Travelscape, LLC"
    if trx_type.startswith("BR"):
        return "11115 Expedia do Brasil Agencia"
    return ""

def map_activity_name(oper_unit: str) -> str:
    """Map Operating Unit to Activity Name."""
    if not isinstance(oper_unit, str):
        return ""
    if oper_unit.startswith("12305"):
        return "CH WRITE-OFF BAD DEBT"
    if oper_unit.startswith("14101"):
        return "TS WRITE-OFF BAD DEBT"
    if oper_unit.startswith("11105"):
        return "US WRITE-OFF BAD DEBT"
    if oper_unit.startswith("11115"):
        return "BR WRITE-OFF BAD DEBT"
    return ""

def send_outlook_email_with_attachments(to_addr, subject, body, attachments):
    """
    Send an e-mail via Outlook (if pywin32 and Outlook are available).
    attachments: list of Paths
    """
    if win32 is None:
        print("pywin32 is not installed. Automatic Outlook email sending is not available.")
        return

    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = MailItem
        mail.To = to_addr
        mail.Subject = subject
        mail.Body = body

        for att in attachments:
            att_path = str(att)
            if os.path.isfile(att_path):
                mail.Attachments.Add(att_path)
            else:
                print(f"Warning: attachment not found, will not be included: {att_path}")

        mail.Send()
        print(f"Email sent to {to_addr} with subject '{subject}'.")
    except Exception as e:
        print(f"Failed to send email via Outlook: {e}")

# ==========================
# RELO DM (.xlsm) BUILD HELPERS
# ==========================

def _find_col_letter_by_header(ws, header_text, search_rows=range(1, 20)):
    """
    Search for a header text in the first rows and return the column letter.
    Exact match (case-insensitive) on the cell value.
    """
    header_text = str(header_text).strip().lower()
    for r in search_rows:
        for cell in ws[r]:
            val = cell.value
            if val is None:
                continue
            if str(val).strip().lower() == header_text:
                return cell.column_letter
    raise KeyError(f"Header '{header_text}' not found in template.")

def _copy_default_from_row(ws, col_letter, from_row, to_row):
    """
    Copy only the cell value from from_row to to_row for a given column letter.
    Styles/colors are not copied so new rows remain visually "clean".
    """
    src = ws[f"{col_letter}{from_row}"]
    dst = ws[f"{col_letter}{to_row}"]
    dst.value = src.value

def build_ts_relo_dm_file(df_match_relo, template_path: Path, out_dir: Path):
    """
    Generate a .xlsm file from the template, with rows only for RELO MATCHES of type TS_CM_RELO,
    transforming the type into TS_DM_RELO and filling columns as per rules.
    Keeps template macros (keep_vba=True).

    Requirements in df_match_relo:
    - columns: ['RECEIPT_NUMBER','LOCAL_RECEIPT_AMOUNT','Transaction Number',
                'Account Number','Transaction Type']
    """
    # 1) Filter only TS_CM_RELO
    if df_match_relo is None or df_match_relo.empty:
        print("Match RELO is empty. No rows to generate RELO DM file.")
        return None

    df_src = (
        df_match_relo
        .loc[lambda d: d["Transaction Type"].astype(str).str.upper() == "TS_CM_RELO"]
        .copy()
    )
    if df_src.empty:
        print("No RELO MATCH rows of type TS_CM_RELO. Nothing to generate.")
        return None

    # Consolidate 1 row per (receipt, trx, account)
    df_src = (
        df_src
        .drop_duplicates(subset=["RECEIPT_NUMBER", "Transaction Number", "Account Number"])
        .loc[:, ["RECEIPT_NUMBER", "Transaction Number", "Account Number", "LOCAL_RECEIPT_AMOUNT"]]
        .copy()
    )
    df_src["AMOUNT_POS"] = df_src["LOCAL_RECEIPT_AMOUNT"].abs().round(2)

    # 2) Open template .xlsm with macros
    if not template_path.exists():
        raise FileNotFoundError(f"Template .xlsm not found: {template_path}")

    wb = load_workbook(template_path, keep_vba=True, data_only=False)

    if "EVC RELO CM" not in wb.sheetnames:
        raise ValueError("Sheet 'EVC RELO CM' not found in template.")

    ws = wb["EVC RELO CM"]  # contains headers and defaults in row 10

    # 3) Map columns by header
    col_customer   = _find_col_letter_by_header(ws, "Customer Number")
    col_shipto     = _find_col_letter_by_header(ws, "Ship To Cust Numb")
    col_trx_date   = _find_col_letter_by_header(ws, "Transaction Date")
    col_txn_type   = _find_col_letter_by_header(ws, "Txn Type")
    col_comments   = _find_col_letter_by_header(ws, "Comments")
    col_reference  = _find_col_letter_by_header(ws, "Reference")
    col_context    = _find_col_letter_by_header(ws, "Context")
    col_line_num   = _find_col_letter_by_header(ws, "Line Number")  # sequential

    col_unit_price = "AM"  # Unit Selling Price
    col_qty        = "AN"  # Quantity
    col_total_amt  = "AP"  # Total Amt

    # Columns that must preserve default value from row 10 (includes AH)
    # Ajustado: removidas colunas T (Credit Reason) e V (skip workflow),
    # para que fiquem em branco nas novas linhas.
    preserve_letters = ["X", "AH", "AI", "AJ", "AK", "AQ", "AR", "AS"]

    # 4) Insert rows starting at row 11 (row 10 = template defaults)
    start_row = 11
    today_dt = date.today()

    for i, (_, rec) in enumerate(df_src.iterrows(), start=0):
        receipt_number = rec["RECEIPT_NUMBER"]
        trx_number     = rec["Transaction Number"]
        account_number = rec["Account Number"]
        amount_pos     = float(rec["AMOUNT_POS"])

        r = start_row + i

        # Basic headers
        ws[f"{col_customer}{r}"] = str(account_number)
        ws[f"{col_shipto}{r}"] = str(account_number)
        ws[f"{col_trx_date}{r}"] = today_dt

        # Txn Type = TS_DM_RELO (also explicitly in column P)
        ws[f"{col_txn_type}{r}"] = "TS_DM_RELO"
        ws[f"P{r}"] = "TS_DM_RELO"

        # Comments formula
        ws[f"{col_comments}{r}"] = '="COA vs WO " & TEXT(TODAY(),"mm/dd/yyyy")'

        # Reference: DM ISSUED TO OFFSET CM <Transaction Number>
        ws[f"{col_reference}{r}"] = f"DM ISSUED TO OFFSET CM {trx_number}"

        # Copy default values from row 10 for specific columns (no T/V)
        for colL in preserve_letters:
            _copy_default_from_row(ws, colL, 10, r)

        # Context also preserves default from row 10
        _copy_default_from_row(ws, col_context, 10, r)

        # Line values (positive)
        ws[f"{col_unit_price}{r}"] = amount_pos   # AM
        ws[f"{col_qty}{r}"] = 1                   # AN
        ws[f"{col_total_amt}{r}"] = amount_pos    # AP

        # Line Number sequential, BUT NOT if "Line Number" column is AH (we keep row 10 value)
        if col_line_num != "AH":
            ws[f"{col_line_num}{r}"] = i + 1

    # 5) Save .xlsm in today's folder
    out_name = f"4 Non-Lodging Relo DM_{today_dt.strftime('%m%d%Y')}.xlsm"
    out_path = out_dir / out_name
    wb.save(out_path)
    print(f"RELO DM file generated: {out_path}")
    return out_path

# ==========================
# MAIN PIPELINE
# ==========================

def main():
    # 1) Latest COA file
    coa_path = find_latest_coa_file()
    print(f"COA file found: {coa_path}")

    # 2) Read 'Data' sheet and extract accounts
    df_coa = pd.read_excel(coa_path, sheet_name="Data")

    if "CUSTOMER_NBR" not in df_coa.columns:
        raise ValueError("Column CUSTOMER_NBR not found in 'Data' sheet of COA.")

    customers = (
        df_coa["CUSTOMER_NBR"]
        .dropna().astype(str).str.strip()
        .loc[lambda s: s != ""].unique().tolist()
    )
    print(f"Total distinct CUSTOMER_NBR: {len(customers)}")

    # Sheet 'CustList'
    df_cust = pd.DataFrame({"CUSTOMER_NBR": customers})
    with pd.ExcelWriter(coa_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        df_cust.to_excel(w, sheet_name="CustList", index=False)
    print("Sheet 'CustList' created/updated in COA.")

    # Update Concat.xlsx
    update_concat_column_b(customers)

    # 3) Connect to Oracle (THICK)
    print("Initializing Oracle Client...")
    oracledb.init_oracle_client(lib_dir=ORACLE_CLIENT_DIR)

    all_wo_rows, wo_columns = [], None
    all_relo_rows, relo_columns = [], None

    print("Connecting to Oracle...")
    with oracledb.connect(user=USERNAME, password=PASSWORD, dsn=DSN) as conn:
        print("Connected.")
        with conn.cursor() as cur:
            cur.execute("ALTER SESSION SET CURRENT_SCHEMA = APPS")
            print("ALTER SESSION SET CURRENT_SCHEMA = APPS executed.")

        total_chunks = (len(customers) + 999) // 1000

        for idx, acc_chunk in enumerate(chunk_list(customers, 1000), start=1):
            print(f"\n[Batch {idx}/{total_chunks}] Accounts: {len(acc_chunk)}")
            print("First accounts:", acc_chunk[:5])

            placeholders = ", ".join(f":acc{i}" for i in range(len(acc_chunk)))
            params = {f"acc{i}": v for i, v in enumerate(acc_chunk)}

            # WO
            sql_wo = BASE_QUERY.format(placeholders=placeholders)
            try:
                with conn.cursor() as cur:
                    cur.execute(sql_wo, params)
                    rows = cur.fetchall()
                    if rows:
                        if wo_columns is None:
                            wo_columns = [d[0] for d in cur.description]
                        all_wo_rows.extend(rows)
                    print(f"  -> WO: {len(rows)} rows")
            except oracledb.DatabaseError as e:
                error, = e.args
                print(f"  !! ERROR in WO batch {idx}: {error.message}")
                break

            # RELO
            sql_relo = RELO_QUERY.format(placeholders=placeholders)
            try:
                with conn.cursor() as cur:
                    cur.execute(sql_relo, params)
                    rows = cur.fetchall()
                    if rows:
                        if relo_columns is None:
                            relo_columns = [d[0] for d in cur.description]
                        all_relo_rows.extend(rows)
                    print(f"  -> RELO: {len(rows)} rows")
            except oracledb.DatabaseError as e:
                error, = e.args
                print(f"  !! ERROR in RELO batch {idx}: {error.message}")
                break

    if not all_wo_rows:
        print("No rows returned for WO. Exiting.")
        return

    df_wo = pd.DataFrame(all_wo_rows, columns=wo_columns)
    with pd.ExcelWriter(coa_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        df_wo.to_excel(w, sheet_name="WO", index=False)
    print(f"Sheet 'WO' created with {len(df_wo)} rows.")

    df_relo = None
    if all_relo_rows:
        df_relo = pd.DataFrame(all_relo_rows, columns=relo_columns)
        with pd.ExcelWriter(coa_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
            df_relo.to_excel(w, sheet_name="Relo CMs", index=False)
        print(f"Sheet 'Relo CMs' created with {len(df_relo)} rows.")
    else:
        print("No rows returned for Relo CMs.")

    # 4) Normalize Data and GENERATE OUTPUT FILES
    req_data = ["RECEIPT_NUMBER", "LOCAL_RECEIPT_AMOUNT", "CUSTOMER_NBR", "CURRENCY_CODE", "RECEIPT_STATUS"]
    miss_data = [c for c in req_data if c not in df_coa.columns]
    if miss_data:
        print(f"Missing columns in Data: {miss_data}.")
        return

    df_coa_match = df_coa.copy()
    df_coa_match["CUSTOMER_NBR"] = df_coa_match["CUSTOMER_NBR"].astype(str).str.strip()
    df_coa_match["LOCAL_RECEIPT_AMOUNT"] = pd.to_numeric(df_coa_match["LOCAL_RECEIPT_AMOUNT"], errors="coerce")
    df_coa_match["CURRENCY_CODE"] = df_coa_match["CURRENCY_CODE"].astype(str).str.strip()
    df_coa_match["receipt_amount"] = df_coa_match["LOCAL_RECEIPT_AMOUNT"].round(2)

    # ========= WO MATCHES (SUMIFS by TRANSACTION NUMBER) + CSV =========
    req_wo = ["Transaction Number", "Transaction Type", "Account Number", "Entered Amount", "Entered Currency"]
    miss_wo = [c for c in req_wo if c not in df_wo.columns]
    if miss_wo:
        print(f"Missing columns in WO: {miss_wo}. Skipping WO matches / CSV.")
        df_matches = None
        df_matches_agg = None
        csv_path = None
        cash_xlsx_path = None
    else:
        df_wo_m = df_wo.copy()
        df_wo_m["Account Number"] = df_wo_m["Account Number"].astype(str).str.strip()
        df_wo_m["Entered Amount"] = pd.to_numeric(df_wo_m["Entered Amount"], errors="coerce")
        df_wo_m["Entered Currency"] = df_wo_m["Entered Currency"].astype(str).str.strip()
        df_wo_m["Transaction Type"] = df_wo_m["Transaction Type"].astype(str)

        # Exclude any TRANSACTION NUMBER having TYPE containing 'GROUP'
        group_trx = df_wo_m.loc[
            df_wo_m["Transaction Type"].str.contains("GROUP", case=False, na=False),
            "Transaction Number",
        ].unique()
        if len(group_trx) > 0:
            print(f"Excluding {len(group_trx)} Transaction Number(s) with Transaction Type containing 'GROUP' from WO matching.")
            df_wo_m = df_wo_m[~df_wo_m["Transaction Number"].isin(group_trx)].copy()

        # SUMIFS by (Account Number, Transaction Number, Currency)
        df_wo_sum_trx = (
            df_wo_m
            .groupby(["Account Number", "Transaction Number", "Entered Currency"], as_index=False)["Entered Amount"]
            .sum()
        )
        df_wo_sum_trx["sum_amount"] = df_wo_sum_trx["Entered Amount"].round(2)

        merged_trx = df_coa_match.merge(
            df_wo_sum_trx[["Account Number", "Transaction Number", "Entered Currency", "sum_amount"]],
            left_on=["CUSTOMER_NBR", "CURRENCY_CODE"],
            right_on=["Account Number", "Entered Currency"],
            how="inner",
        )

        # RECEIPT (positive) == - SUMIFS(WO by TRX) (negative)
        cond_trx = merged_trx["receipt_amount"] == (-merged_trx["sum_amount"])
        keys_ok_trx = merged_trx.loc[
            cond_trx,
            [
                "CUSTOMER_NBR",
                "CURRENCY_CODE",
                "receipt_amount",
                "RECEIPT_NUMBER",
                "RECEIPT_STATUS",
                "Account Number",
                "Transaction Number",
                "Entered Currency",
                "sum_amount",
            ],
        ].drop_duplicates()

        if keys_ok_trx.empty:
            print("No WO matches by Transaction Number.")
            df_matches = None
            df_matches_agg = None
            csv_path = None
            cash_xlsx_path = None
        else:
            # ensure only one match per RECEIPT_NUMBER in Matches sheet
            before_len = len(keys_ok_trx)
            keys_ok_trx = (
                keys_ok_trx
                .sort_values(
                    ["CUSTOMER_NBR", "CURRENCY_CODE", "RECEIPT_NUMBER", "Transaction Number"]
                )
                .drop_duplicates(subset=["RECEIPT_NUMBER"], keep="first")
            )
            after_len = len(keys_ok_trx)
            if after_len < before_len:
                print(
                    "Receipts with multiple WO matches found. "
                    f"Before: {before_len} rows, after limiting to 1 per RECEIPT_NUMBER: {after_len}."
                )

            # matches_wo still at adjustment line level
            matches_wo = df_wo_m.merge(
                keys_ok_trx,
                on=["Account Number", "Transaction Number", "Entered Currency"],
                how="inner",
            )

            # Aggregate by RECEIPT + TRANSACTION NUMBER (remove duplicate adjustments)
            group_cols = [
                "RECEIPT_NUMBER",
                "RECEIPT_STATUS",
                "Transaction Number",
                "CURRENCY_CODE",
                "Entered Currency",
                "Transaction Type",
                "Account Number",
            ]
            df_matches_agg = (
                matches_wo
                .groupby(group_cols, as_index=False)
                .agg({
                    "receipt_amount": "first",   # same value in the group
                    "sum_amount": "first",       # SUMIFS by TRX
                    "Entered Amount": "sum",     # sum of adjustments for that invoice
                    "RECEIPT_STATUS": "first",
                })
            )

            df_matches = df_matches_agg.rename(columns={
                "receipt_amount": "LOCAL_RECEIPT_AMOUNT",
                "sum_amount": "WO_SUMIFS",
                "CURRENCY_CODE": "Receipt Currency Code",
                "Entered Currency": "WO Currency",
            })

            # Write Matches sheet back into COA
            with pd.ExcelWriter(coa_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
                df_matches.to_excel(w, sheet_name="Matches", index=False)
            print(f"Sheet 'Matches' created with {len(df_matches)} rows (1 row per RECEIPT_NUMBER).")

            # Dates for file naming and comments
            today = date.today()
            today_csv_name = today.strftime("%m%d%Y")
            today_comment = today.strftime("%m/%d/%Y")

            # Separate file for Cash (mirroring Matches sheet)
            cash_xlsx_name = f"COAvsWO_Cash_{today_csv_name}.xlsx"
            cash_xlsx_path = OUTPUT_FOLDER / cash_xlsx_name
            df_matches.to_excel(cash_xlsx_path, index=False)
            print(f"Cash file generated with {len(df_matches)} rows at:\n{cash_xlsx_path}")

            # COAvsWO CSV from aggregated matches (1 row per TRX / RECEIPT)
            df_csv = pd.DataFrame()
            df_csv["Index"] = range(1, len(df_matches_agg) + 1)
            oper_units = df_matches_agg["Transaction Type"].map(map_operating_unit)
            df_csv["Operating Unit"] = oper_units
            df_csv["Transaction Number"] = df_matches_agg["Transaction Number"]
            df_csv["BFB Number"] = ""
            df_csv["Activity Name"] = oper_units.map(map_activity_name)
            df_csv["Adjustment Type"] = "Line"
            # positive value (invert sign of Entered Amount)
            df_csv["Amount to be Adjusted"] = (-df_matches_agg["Entered Amount"]).round(2)
            df_csv["Reason"] = "Uneconomical to Collect"
            df_csv["Comments"] = "COA vs WO " + today_comment
            df_csv["GL Date"] = ""
            df_csv["Adjust Date"] = ""

            csv_name = f"COAvsWO{today_csv_name}.csv"
            csv_path = OUTPUT_FOLDER / csv_name
            df_csv.to_csv(csv_path, index=False, encoding="cp1252")
            print(f"CSV generated with {len(df_csv)} rows at:\n{csv_path}")

    # ========= RELO MATCHES (SUMIFS by TRANSACTION NUMBER) + RELO DM =========
    relo_dm_path = None
    if df_relo is not None:
        req_relo = [
            "Transaction Number", "Transaction Type",
            "Account Number", "Entered Amount",
            "Entered Currency", "Reference",
        ]
        miss_relo = [c for c in req_relo if c not in df_relo.columns]
        if miss_relo:
            print(f"Missing columns in Relo CMs: {miss_relo}.")
        else:
            df_relo_m = df_relo.copy()
            df_relo_m["Account Number"] = df_relo_m["Account Number"].astype(str).str.strip()
            df_relo_m["Entered Amount"] = pd.to_numeric(df_relo_m["Entered Amount"], errors="coerce")
            df_relo_m["Entered Currency"] = df_relo_m["Entered Currency"].astype(str).str.strip()

            df_relo_sum_trx = (
                df_relo_m
                .groupby(["Account Number", "Transaction Number", "Entered Currency"], as_index=False)["Entered Amount"]
                .sum()
            )
            df_relo_sum_trx["sum_amount"] = df_relo_sum_trx["Entered Amount"].round(2)

            merged_relo_trx = df_coa_match.merge(
                df_relo_sum_trx[["Account Number", "Transaction Number", "Entered Currency", "sum_amount"]],
                left_on=["CUSTOMER_NBR", "CURRENCY_CODE"],
                right_on=["Account Number", "Entered Currency"],
                how="inner",
            )

            cond_relo_trx = merged_relo_trx["receipt_amount"] == (-merged_relo_trx["sum_amount"])
            keys_relo_trx = merged_relo_trx.loc[
                cond_relo_trx,
                [
                    "CUSTOMER_NBR", "CURRENCY_CODE", "receipt_amount", "RECEIPT_NUMBER",
                    "Account Number", "Transaction Number", "Entered Currency", "sum_amount",
                ],
            ].drop_duplicates()

            if keys_relo_trx.empty:
                print("No RELO matches by Transaction Number.")
            else:
                matches_relo = df_relo_m.merge(
                    keys_relo_trx,
                    on=["Account Number", "Transaction Number", "Entered Currency"],
                    how="inner",
                )

                df_match_relo = matches_relo[[
                    "RECEIPT_NUMBER",
                    "receipt_amount",
                    "Transaction Number",
                    "Entered Amount",
                    "CURRENCY_CODE",
                    "Entered Currency",
                    "Reference",
                    "Account Number",
                    "Transaction Type",
                ]].rename(columns={
                    "receipt_amount": "LOCAL_RECEIPT_AMOUNT",
                    "CURRENCY_CODE": "Receipt Currency Code",
                    "Entered Currency": "WO Currency",
                })

                # Write Match RELO sheet in COA
                with pd.ExcelWriter(coa_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
                    df_match_relo.to_excel(w, sheet_name="Match RELO", index=False)
                print(f"Sheet 'Match RELO' created with {len(df_match_relo)} rows.")

                # Also write Match RELO into the Cash file, if generated
                if 'cash_xlsx_path' in locals() and cash_xlsx_path is not None:
                    with pd.ExcelWriter(cash_xlsx_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
                        df_match_relo.to_excel(w, sheet_name="Match RELO", index=False)
                    print(f"Sheet 'Match RELO' added to Cash file: {cash_xlsx_path}")

                # Create RELO DM .xlsm file (only TS_CM_RELO)
                try:
                    relo_dm_path = build_ts_relo_dm_file(df_match_relo, RELO_DM_TEMPLATE, OUTPUT_FOLDER)
                    if relo_dm_path:
                        print(f"RELO DM (.xlsm) created: {relo_dm_path}")
                except Exception as e:
                    print(f"Failed to generate RELO DM file: {e}")

    # Send email via Outlook with output files (including RELO DM if generated)
    if csv_path is not None and 'cash_xlsx_path' in locals() and cash_xlsx_path is not None:
        today_comment = date.today().strftime("%m/%d/%Y")
        email_to = "hotelcollectbilling@expedia.com"
        email_subject = f"COA vs WO {today_comment}"
        email_body = (
            f"Hi team,\n\n"
            f"Please see attached the COA vs WO outputs for {today_comment}.\n\n"
            f"Best regards,\n"
            f"José\n"
        )
        attachments = [csv_path, cash_xlsx_path]
        if 'relo_dm_path' in locals() and relo_dm_path:
            attachments.append(relo_dm_path)

        send_outlook_email_with_attachments(email_to, email_subject, email_body, attachments)

if __name__ == "__main__":
    main()
