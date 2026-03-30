import os
from pathlib import Path
from datetime import date
import re

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

# 3 HC Relo Upload DM template (keeps macros)
HC_DM_TEMPLATE = Path(
    r"C:\Users\josenjr\OneDrive - Expedia Group\Desktop\Daily activities\COA vs WO Process\3 HC Relo Upload DM.xlsm"
)

# ==========================
# ORACLE CONFIG
# ==========================

ORACLE_CLIENT_DIR = r"C:\Users\josenjr\OneDrive - Expedia Group\Desktop\Oracle Instant Client\instantclient_23_0"
USERNAME = "josenjr"
PASSWORD = "qyuxYQZFs13"      # CHANGE ME
DSN      = "ashworaebsdb02-vip.datawarehouse.expecn.com:1526/ORAPRD_UI"

# WO adjustments query
BASE_QUERY = """
SELECT
CTA.TRX_NUMBER              AS "Transaction Number",
CTTA.NAME                   AS "Transaction Type",
CTA.TRX_DATE                AS "Transaction Date",
HCA.ACCOUNT_NUMBER          AS "Account Number",
CTA.INVOICE_CURRENCY_CODE   AS "Entered Currency",
AAA.AMOUNT                  AS "Entered Amount",
AAA.ADJUSTMENT_NUMBER       AS "Adjustment Number",
AAA.REASON_CODE             AS "Reason Code"
FROM AR.AR_ADJUSTMENTS_ALL      AAA
JOIN AR.RA_CUSTOMER_TRX_ALL     CTA  ON CTA.CUSTOMER_TRX_ID   = AAA.CUSTOMER_TRX_ID
JOIN AR.RA_CUST_TRX_TYPES_ALL   CTTA ON CTTA.CUST_TRX_TYPE_ID = CTA.CUST_TRX_TYPE_ID
JOIN AR.HZ_CUST_ACCOUNTS        HCA  ON HCA.CUST_ACCOUNT_ID   = CTA.BILL_TO_CUSTOMER_ID
JOIN APPS.HR_OPERATING_UNITS    HOU  ON HOU.ORGANIZATION_ID   = CTA.ORG_ID
JOIN APPS.GL_LEDGER_LE_V        GLL  ON GLL.LEGAL_ENTITY_ID   = HOU.DEFAULT_LEGAL_CONTEXT_ID
WHERE
GLL.LEDGER_CATEGORY_CODE = 'PRIMARY'
AND CTTA.NAME NOT LIKE '%EAC%'
AND NOT (UPPER(CTTA.NAME) LIKE '%GROUP%' AND UPPER(CTTA.NAME) NOT LIKE '%XLR%')
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

# Nova query: Transaction Register p/ CMs da aba Match RELO + invoice aplicada
CM_APPLIED_QUERY = """
/*
* Oracle Transaction Register - UAT
* Author: Jose Santos (josenjr@expediagroup.com)
*/

SELECT  

CASE 
  WHEN HZP.COUNTRY in ('AS','AU','BD','BN','BT','CC','CK','CN','FJ','FM','GU','HK','ID','IN','JP','KH','KP','KR','LA','LK','MH','MM','MN','MO','MP','MV','MY','NC','NF','NP','NU','NZ','PF','PG','PH','PK','PW','SB','SG','TH','TK','TL','TO','TW','VN','VU','WF','WS')
  THEN 'APAC'
  WHEN HZP.COUNTRY in ('AG','AI','AN','AR','AW','BB','BL','BM','BO','BQ','BR','BS','BZ','CL','CO','CR','CU','CW','DM','DO','EC','FK','GD','GF','GP','GT','GY','HN','HT','JM','KN','KY','LC','MF','MQ','MS','MX','NI','PA','PE','PN','PR','PY','SR','SV','SX','TC','TT','UY','VC','VE','VG','VI')
  THEN 'LATAM'
  WHEN HZP.COUNTRY in ('AL','AM','AT','AZ','BA','BE','BG','BY','CH','CZ','DE','DK','EE','FI','FO','FR','GB','GE','GG','GL','HR','HU','IE','IS','JE','KG','KZ','LI','LT','LU','LV','MD','ME','MK','NL','NO','PL','PM','RO','RS','RU','SE','SH','SI','SJ','SK','TJ','TM','UA','UZ')
  THEN 'N-EMEA'
  WHEN HZP.COUNTRY in ('AD','AE','AF','AO','BF','BH','BI','BJ','BV','BW','CD','CF','CG','CI','CM','CV','CY','DJ','DZ','EG','ER','ES','ET','GA','GH','GI','GM','GN','GQ','GR','GW','IL','IQ','IR','IT','JO','KE','KM','KW','LB','LR','LS','LY','MA','MC','MG','ML','MR','MT','MU','MW','MZ','NA','NE','NG','OM','PS','PT','QA','RE','RW','SA','SC','SD','SL','SM','SN','SS','ST','SY','SZ','TD','TG','TN','TR','TZ','UG','VA','YE','YT','ZA','ZM','ZW')
  THEN 'S-EMEA'
  WHEN HZP.COUNTRY in ('CA','US')
  THEN 'NAMER'
  ELSE 'NOT_DEFINED'
  END AS "POAR Super Region"

,GLL.LEDGER_NAME "Ledger"
,HOU.NAME "OU NAME"
,SUBSTR(HOU.NAME,6,51) "Legal Entity"
,SUBSTR(HOU.NAME,1,5) "Company"
,PSA.INVOICE_CURRENCY_CODE "Currency"
,PSA.CLASS "CLASS"
,CTA.ATTRIBUTE4 "XLR ID"
,CTA.TRX_NUMBER "Transaction Number"      -- CM number
,CTTA.NAME "TRX Type Name"
,HZP.PARTY_NAME "Customer Name"
,HCA.ACCOUNT_NUMBER "Account Number"
,CTA.TRX_DATE "TRX DATE"
,PSA.GL_DATE "GL Date"
,TO_CHAR(CTA.TRX_DATE,'YYYY-MM')  "Year / Month"
,PSA.AMOUNT_DUE_ORIGINAL "Entered Amt"
,PSA.AMOUNT_DUE_REMAINING "Open Balance"
,CASE 
    WHEN SUBSTR(CTTA.NAME,1,2) in ('CH','US')
    THEN ROUND(CTA.EXCHANGE_RATE * PSA.AMOUNT_DUE_ORIGINAL,2)
    ELSE PSA.AMOUNT_DUE_ORIGINAL
    END AS "Functional Amt"
,CASE
    WHEN PSA.INVOICE_CURRENCY_CODE = 'USD'
    THEN 1
    ELSE ROUND(GLDR.CONVERSION_RATE,15)
    END AS "Conversion Rate"
,CASE WHEN PSA.INVOICE_CURRENCY_CODE = 'USD'
    THEN PSA.AMOUNT_DUE_ORIGINAL
    ELSE ROUND(ROUND(GLDR.CONVERSION_RATE,15) * PSA.AMOUNT_DUE_ORIGINAL,11) 
    END AS "USD Amount"
,TO_CHAR(CTA.PRINTING_LAST_PRINTED ,'MM/DD/YYYY HH:MI')  "Invoice Print Date"
,CTA.INTERFACE_HEADER_ATTRIBUTE1 "Reference"
,CTA.COMMENTS "Comments"
,CTA.INTERFACE_HEADER_CONTEXT "Batch Name"
,CTA.REASON_CODE "Reason Code"
,ARM.NAME "Receipt Method"
,HZP.CITY "City"
,HZP.COUNTRY "Country"
,ROUND(CTLA.TAX_AMOUNT,2) "Tax Amount"
,CASE  WHEN PSA.AMOUNT_IN_DISPUTE IS NULL
    THEN 0
    ELSE PSA.AMOUNT_IN_DISPUTE
    END AS "Amount In Dispute"
,CASE WHEN PSA.AMOUNT_IN_DISPUTE IS NULL or PSA.AMOUNT_IN_DISPUTE = '0.00'
    THEN 'NO'
    ELSE 'YES'
    END AS "TRX ON DISPUTE"
,CASE
    WHEN PSA.DISPUTE_DATE IS NULL
    THEN NULL
    ELSE PSA.DISPUTE_DATE
    END AS "Dispute Date"
,HCA.ATTRIBUTE1 "Hotel ID"
,TT.NAME "Term Name"
,FNU.USER_NAME "Created By User"

,CASE 
WHEN HCA.CUSTOMER_CLASS_CODE = 'TA'
    THEN 'Travel Ad'
    WHEN HCA.CUSTOMER_CLASS_CODE = 'HOTEL'
            AND CTTA.NAME NOT LIKE '%GROUP%'
            THEN 'Hotel Collect Independent'
    WHEN HCA.CUSTOMER_CLASS_CODE = 'HOTEL'
            AND CTTA.NAME LIKE '%GROUP%'
            THEN 'Hotel Collect Corporate'
    WHEN HCA.CUSTOMER_CLASS_CODE = 'MESODB'
    THEN 'Meso Direct Bill'
    WHEN HCA.CUSTOMER_CLASS_CODE = 'SUPPLIER'
            OR HCA.CUSTOMER_CLASS_CODE = 'OTHER'
            OR CTTA.NAME LIKE '%SO%'
            THEN 'Supplier Other'
    WHEN HCA.CUSTOMER_CLASS_CODE = 'GROUP PARENT'
    THEN 'Hotel Collect Corporate'
    WHEN HCA.CUSTOMER_CLASS_CODE = 'MESOMF'
    THEN 'Meso Marketing Funds'
    WHEN HCA.CUSTOMER_CLASS_CODE = 'ELE'
    THEN 'Expedia Local Expert'
    ELSE 'UNLABLED ITEM - PLEASE CONTACT ICORT'
    END AS "Business Model"

,CASE
WHEN CTTA.NAME LIKE '%RELO%'
    THEN 'RELO'
    WHEN CTTA.NAME LIKE '%VOIDED%'
    THEN 'VOIDED'
    WHEN CTTA.NAME LIKE '%FRAUD%'
    THEN 'FRAUD'
    WHEN CTTA.NAME LIKE '%REBILL%'
    THEN 'REBILL'
    ELSE 'Standard TRX'
    END AS "TRX Type"

,CASE
WHEN ARM.RECEIPT_METHOD_ID IN ('4004','56065','4009','4005','4003','4008','56066','39002','4002','4006','4000','4011','56064','32000','37000','39003','4010','72021','72020','72022')
    THEN 'AP'
    ELSE 'Non AP'
    END AS "AP Status"

/* >>> INVOICE EM QUE O CM ESTÁ APLICADO <<< */
,INV.TRX_NUMBER  AS "Applied Invoice Number"
,INV.TRX_DATE    AS "Applied Invoice Date"
/* <<< FIM DO BLOCO >>> */

FROM AR.RA_CUSTOMER_TRX_ALL CTA

LEFT JOIN AR.HZ_CUST_ACCOUNTS HCA
  ON CTA.BILL_TO_CUSTOMER_ID = HCA.CUST_ACCOUNT_ID

LEFT JOIN AR.HZ_PARTIES HZP
  ON HCA.PARTY_ID = HZP.PARTY_ID

LEFT JOIN APPS.HR_OPERATING_UNITS HOU
  ON HOU.ORGANIZATION_ID = CTA.ORG_ID

LEFT JOIN APPS.GL_LEDGER_LE_V GLL
  ON GLL.LEGAL_ENTITY_ID = HOU.DEFAULT_LEGAL_CONTEXT_ID

LEFT JOIN AR.AR_PAYMENT_SCHEDULES_ALL PSA
  ON PSA.CUSTOMER_TRX_ID = CTA.CUSTOMER_TRX_ID

LEFT JOIN AR.AR_RECEIPT_METHODS ARM
  ON CTA.RECEIPT_METHOD_ID = ARM.RECEIPT_METHOD_ID

LEFT JOIN AR.RA_CUST_TRX_TYPES_ALL CTTA
  ON CTTA.CUST_TRX_TYPE_ID = CTA.CUST_TRX_TYPE_ID

LEFT JOIN APPS.AR_CUSTOMER_PROFILES_V CPV
  ON HCA.CUST_ACCOUNT_ID = CPV.CUSTOMER_ID

LEFT JOIN APPS.FND_USER FNU
  ON FNU.USER_ID = CTA.CREATED_BY

LEFT JOIN (
  SELECT SUM(AR.RA_CUSTOMER_TRX_LINES_ALL.EXTENDED_AMOUNT) "TAX_AMOUNT",
         AR.RA_CUSTOMER_TRX_LINES_ALL.CUSTOMER_TRX_ID,
         AR.RA_CUSTOMER_TRX_LINES_ALL.LINE_TYPE
  FROM AR.RA_CUSTOMER_TRX_LINES_ALL
  WHERE AR.RA_CUSTOMER_TRX_LINES_ALL.LINE_TYPE = 'TAX'
  GROUP BY AR.RA_CUSTOMER_TRX_LINES_ALL.CUSTOMER_TRX_ID,
           AR.RA_CUSTOMER_TRX_LINES_ALL.LINE_TYPE
) CTLA
  ON CTA.CUSTOMER_TRX_ID = CTLA.CUSTOMER_TRX_ID

LEFT JOIN (
  SELECT GL.GL_DAILY_RATES.CONVERSION_RATE,
         GL.GL_DAILY_RATES.FROM_CURRENCY,
         GL.GL_DAILY_RATES.TO_CURRENCY,
         GL.GL_DAILY_RATES.CONVERSION_DATE
  FROM GL.GL_DAILY_RATES 
  WHERE GL.GL_DAILY_RATES.TO_CURRENCY = 'USD'
    AND GL.GL_DAILY_RATES.CONVERSION_TYPE = 'Corporate'
) GLDR
  ON (GLDR.FROM_CURRENCY = PSA.INVOICE_CURRENCY_CODE
      AND GLDR.CONVERSION_DATE = PSA.GL_DATE)

LEFT JOIN AR.RA_TERMS_TL TT
  ON PSA.TERM_ID = TT.TERM_ID

/* >>> JOINS CM -> INVOICE APLICADA <<< */
LEFT JOIN AR.AR_RECEIVABLE_APPLICATIONS_ALL RAA
  ON RAA.CUSTOMER_TRX_ID = CTA.CUSTOMER_TRX_ID   -- este CTA é o CM
 AND RAA.APPLICATION_TYPE = 'CM'
 AND RAA.DISPLAY = 'Y'

LEFT JOIN AR.RA_CUSTOMER_TRX_ALL INV
  ON INV.CUSTOMER_TRX_ID = RAA.APPLIED_CUSTOMER_TRX_ID   -- invoice onde o CM foi aplicado
/* <<< FIM DAS NOVAS JOINS >>> */

WHERE 
  GLL.LEDGER_CATEGORY_CODE = 'PRIMARY'  -- NEVER REMOVE
  AND CTTA.NAME NOT LIKE '%EAC%'

  AND CTA.TRX_NUMBER IN ({placeholders})
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

def normalize_reason(reason):
    """Normalize reason codes like 'UNECONOMICAL_TO_COLLECT' -> 'Uneconomical to Collect'."""
    if pd.isna(reason) or not isinstance(reason, str) or not reason.strip():
        return "Uneconomical to Collect"
    reason_clean = reason.strip().upper()
    overrides = {
        "UNECONOMICAL_TO_COLLECT": "Uneconomical to Collect",
        "SMALL_AMT_REMAINING": "Small Amt Remaining",
    }
    if reason_clean in overrides:
        return overrides[reason_clean]
    return reason_clean.replace("_", " ").title()

def send_outlook_email_with_attachments(to_addr, subject, body, attachments):
    """
    Send an e-mail via Outlook (if pywin32 and Outlook are available).
    attachments: list of Paths
    """
    if win32 is None:
        print("pywin32 is not installed.\nAutomatic Outlook email sending is not available.")
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
    Styles/colors are not copied so new rows remain visually clean.
    """
    src = ws[f"{col_letter}{from_row}"]
    dst = ws[f"{col_letter}{to_row}"]
    dst.value = src.value

def build_ts_relo_dm_file(df_match_relo, template_path: Path, out_dir: Path):
    """
    Generate a .xlsm file from the template, with rows only for RELO MATCHES of type TS_CM_RELO,
    transforming the type into TS_DM_RELO and filling columns as per rules.
    Keeps template macros (keep_vba=True).
    """
    # 1) Filter only TS_CM_RELO
    if df_match_relo is None or df_match_relo.empty:
        print("Match RELO is empty.\nNo rows to generate RELO DM file.")
        return None

    df_src = (
        df_match_relo
        .loc[lambda d: d["Transaction Type"].astype(str).str.upper() == "TS_CM_RELO"]
        .copy()
    )

    if df_src.empty:
        print("No RELO MATCH rows of type TS_CM_RELO.\nNothing to generate.")
        return None

    # Consolidate 1 row per (receipt, trx, account)
    df_src = (
        df_src
        .drop_duplicates(subset=["RECEIPT_NUMBER", "Transaction Number", "Account Number"])
        .loc[:, [
            "RECEIPT_NUMBER",
            "Transaction Number",
            "Account Number",
            "LOCAL_RECEIPT_AMOUNT",
            "WO Currency",
        ]]
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

    # tenta achar a coluna de moeda no template
    try:
        col_currency = _find_col_letter_by_header(ws, "Currency")
    except KeyError:
        col_currency = None  # se não existir, só não preenche

    col_uom        = "AM"  # Unit of Measure
    col_unit_price = "AN"  # Unit Selling Price
    col_total_amt  = "AP"  # Total Amt

    # Columns that must preserve default value from row 10 (includes AH)
    # T (Credit Reason) and V (skip workflow) removed to keep them blank
    preserve_letters = ["X", "AH", "AI", "AJ", "AK", "AQ", "AR", "AS"]

    # 4) Insert rows starting at row 11 (row 10 = template defaults)
    start_row = 11
    today_dt = date.today()

    for i, (_, rec) in enumerate(df_src.iterrows(), start=0):
        trx_number     = rec["Transaction Number"]
        account_number = rec["Account Number"]
        amount_pos     = float(rec["AMOUNT_POS"])
        wo_currency    = str(rec["WO Currency"]).strip() if pd.notna(rec["WO Currency"]) else ""

        r = start_row + i

        # Basic headers
        ws[f"{col_customer}{r}"] = str(account_number)
        ws[f"{col_shipto}{r}"]   = str(account_number)
        ws[f"{col_trx_date}{r}"] = today_dt

        # Currency vinda da Match RELO (WO Currency)
        if col_currency is not None:
            ws[f"{col_currency}{r}"] = wo_currency

        # Txn Type = TS_DM_RELO (also explicitly in column P)
        ws[f"{col_txn_type}{r}"] = "TS_DM_RELO"
        ws[f"P{r}"]              = "TS_DM_RELO"

        # Comments formula
        ws[f"{col_comments}{r}"] = '="COA vs WO " & TEXT(TODAY(),"mm/dd/yyyy")'

        # Reference: DM ISSUED TO OFFSET CM <Transaction Number>
        ws[f"{col_reference}{r}"] = f"DM ISSUED TO OFFSET CM {trx_number}"

        # Copy default values from row 10 for specific columns (no T/V)
        for colL in preserve_letters:
            _copy_default_from_row(ws, colL, 10, r)

        # Context also preserves default from row 10
        _copy_default_from_row(ws, col_context, 10, r)

        # Line values
        ws[f"{col_uom}{r}"]        = "EA"        # AM: Unit of Measure
        ws[f"{col_unit_price}{r}"] = amount_pos  # AN: Unit Selling Price
        ws[f"{col_total_amt}{r}"]  = amount_pos  # AP: Total Amt

        # Line Number sequential, BUT NOT if "Line Number" column is AH (we keep row 10 value)
        if col_line_num != "AH":
            ws[f"{col_line_num}{r}"] = i + 1

    # 5) Save .xlsm in today's folder
    out_name = f"4 Non-Lodging Relo DM_{today_dt.strftime('%m%d%Y')}.xlsm"
    out_path = out_dir / out_name
    wb.save(out_path)

    print(f"RELO DM file generated: {out_path}")
    return out_path

def build_hc_relo_dm_file(df_match_relo: pd.DataFrame, template_path: Path, out_dir: Path):
    """
    Generate a .xlsm file from the template '3 HC Relo Upload DM.xlsm', using only
    MATCH RELO rows whose Transaction Type is different from TS_CM_RELO.

    Mantém a estrutura do arquivo base e escreve o Transaction Type
    SEMPRE na coluna P (Txn type), sem criar coluna nova.
    """
    if df_match_relo is None or df_match_relo.empty:
        print("Match RELO is empty.\nNo rows to generate 3 HC Relo Upload DM file.")
        return None

    # Filter only Transaction Type != TS_CM_RELO
    df_src = (
        df_match_relo
        .loc[lambda d: d["Transaction Type"].astype(str).str.upper() != "TS_CM_RELO"]
        .copy()
    )

    if df_src.empty:
        print("No RELO MATCH rows with Transaction Type different from TS_CM_RELO.\nNothing to generate for 3 HC Relo Upload DM.")
        return None

    # Consolidate 1 row per (receipt, trx, account)
    df_src = (
        df_src
        .drop_duplicates(subset=["RECEIPT_NUMBER", "Transaction Number", "Account Number"])
        .loc[:, [
            "RECEIPT_NUMBER",
            "Transaction Number",
            "Account Number",
            "LOCAL_RECEIPT_AMOUNT",
            "WO Currency",
            "Transaction Type",
        ]]
        .copy()
    )

    df_src["AMOUNT_POS"] = df_src["LOCAL_RECEIPT_AMOUNT"].abs().round(2)
    df_src["WO Currency"] = df_src["WO Currency"].astype(str).str.strip()

    if not template_path.exists():
        raise FileNotFoundError(f"Template .xlsm not found: {template_path}")

    wb = load_workbook(template_path, keep_vba=True, data_only=False)
    today_dt = date.today()

    # Cache per-entity sheet configuration (US, CH, TS)
    sheet_cfg = {}

    for _, rec in df_src.iterrows():
        trx_type = str(rec["Transaction Type"] or "")
        trx_type_up = trx_type.upper()

        if trx_type_up.startswith("US"):
            entity_key = "US"
            new_trx_type = "US_DIR_DM_RELO_USD"
        elif trx_type_up.startswith("CH"):
            entity_key = "CH"
            new_trx_type = "CH_DIR_DM_RELO_USD"
        elif trx_type_up.startswith("TS"):
            entity_key = "TS"
            new_trx_type = "TS_DIR_DM_RELO_USD"
        else:
            # Not mapped to this file (e.g., BR)
            continue

        # Initialize / get sheet config
        if entity_key not in sheet_cfg:
            target_ws = None
            for name in wb.sheetnames:
                if name.upper().startswith(entity_key):
                    target_ws = wb[name]
                    break

            if target_ws is None:
                raise ValueError(
                    f"Sheet for entity '{entity_key}' not found in 3 HC Relo Upload DM template."
                )

            ws = target_ws
            next_row = ws.max_row + 1

            def safe_find(header):
                try:
                    return _find_col_letter_by_header(ws, header)
                except KeyError:
                    return None

            cfg = {
                "ws": ws,
                "next_row": next_row,
                "col_trx_date": safe_find("Transaction Date"),
                "col_gl_date": safe_find("GL Date"),
                "col_currency": safe_find("Currency Code"),
                "col_comments": safe_find("Comments"),
                "col_line_number": safe_find("Line Number"),
                "col_line_type": safe_find("Line Type"),
                "col_item": safe_find("Item"),
                "col_description": safe_find("Description"),
                "col_quantity": safe_find("Quantity"),
                # Transaction Type será SEMPRE coluna P (Txn type)
            }
            sheet_cfg[entity_key] = cfg

        cfg = sheet_cfg[entity_key]
        ws = cfg["ws"]
        r = cfg["next_row"]
        cfg["next_row"] += 1

        account_number = str(rec["Account Number"])
        trx_number = str(rec["Transaction Number"])
        amount_pos = float(rec["AMOUNT_POS"])
        currency = str(rec["WO Currency"])

        # Fixed letters from spec
        ws[f"I{r}"] = account_number   # Customer Number
        ws[f"L{r}"] = account_number   # Ship To Customer Number

        # Datas (Transaction Date + GL Date) = hoje
        if cfg["col_trx_date"]:
            ws[f"{cfg['col_trx_date']}{r}"] = today_dt
        if cfg["col_gl_date"]:
            ws[f"{cfg['col_gl_date']}{r}"] = today_dt

        # Transaction Type NOVO, SEMPRE na coluna P (Txn type)
        ws[f"P{r}"] = new_trx_type

        # Comments formula (inclui CM / Transaction Number)
        if cfg["col_comments"]:
            ws[f"{cfg['col_comments']}{r}"] = (
                f'="DM issued to offset Relocation CM - Coa vs WO " & '
                f'TEXT(TODAY(),"mm/dd/yyyy") & " " & "{trx_number}"'
            )

        # Print option (W)
        ws[f"W{r}"] = "Do Not Print"

        # Currency code
        if cfg["col_currency"]:
            ws[f"{cfg['col_currency']}{r}"] = currency

        # Context / Billing / etc.
        ws[f"AC{r}"] = "DIRECT AGENCY"
        ws[f"AD{r}"] = currency           # Billing Currency
        ws[f"AG{r}"] = "INDEPENDENTS"     # Hotel Category
        ws[f"AI{r}"] = "More4Apps"        # Int Header
        ws[f"AJ{r}"] = f"DM Issued to offset CM {trx_number}"  # Reference

        # Line number, type, item, description, qty
        if cfg["col_line_number"]:
            ws[f"{cfg['col_line_number']}{r}"] = 1
        if cfg["col_line_type"]:
            ws[f"{cfg['col_line_type']}{r}"] = "Line"
        if cfg["col_item"]:
            ws[f"{cfg['col_item']}{r}"] = "EXPWA_COMP_CM"
        if cfg["col_description"]:
            ws[f"{cfg['col_description']}{r}"] = "COMPENSATION CREDIT"
        if cfg["col_quantity"]:
            ws[f"{cfg['col_quantity']}{r}"] = 1

        # UOM + amounts
        ws[f"AP{r}"] = "EA"
        ws[f"AQ{r}"] = amount_pos
        ws[f"AS{r}"] = amount_pos

        # Line context / Management Unit / Traveller
        ws[f"AT{r}"] = "DIRECT AGENCY"
        ws[f"AU{r}"] = "1097"
        ws[f"AV{r}"] = trx_number

        # Source currency/amount
        ws[f"AW{r}"] = currency
        ws[f"AX{r}"] = amount_pos

        # Datas AZ, BA, BC
        ws[f"AZ{r}"] = today_dt
        ws[f"BA{r}"] = today_dt
        ws[f"BC{r}"] = today_dt

        # Reservation ID / Int Line context / Business model
        ws[f"BD{r}"] = "REQUESTED"
        ws[f"BE{r}"] = "EG_INVOICING"
        ws[f"BF{r}"] = "DIR"

    out_name = f"3 HC Relo Upload DM_{today_dt.strftime('%m%d%Y')}.xlsm"
    out_path = out_dir / out_name
    wb.save(out_path)
    print(f"3 HC Relo Upload DM file generated: {out_path}")
    return out_path

# ==========================
# NOVAS FUNÇÕES: comparação Reference x Applied Invoice Number
# ==========================

def _extract_invoice_from_reference(ref: str) -> str:
    """
    Remove o prefixo 'CREDIT INVOICE' e retorna só o número da invoice
    que está na Reference da query de Relo CMs.
    """
    if pd.isna(ref):
        return ""
    s = str(ref).strip()
    # remove 'CREDIT INVOICE' (case-insensitive)
    s = re.sub(r"(?i)^CREDIT\s+INVOICE\s*", "", s).strip()
    # pega a primeira sequência de dígitos (se existir)
    m = re.search(r"\d+", s)
    return m.group(0) if m else s

def check_relo_cms_already_reversed(df_match_relo: pd.DataFrame, coa_path: Path):
    """
    1) Pega todos os CMs da aba Match RELO
    2) Roda a CM_APPLIED_QUERY no Oracle (só para esses CMs)
    3) Compara Reference (limpa) x Applied Invoice Number
    4) CMs onde não bate (e têm Applied Invoice Number preenchido) vão
       para a aba 'CMs Already Reversed' na planilha-mãe.
    """
    if df_match_relo is None or df_match_relo.empty:
        print("Match RELO vazio. Nada para comparar com Applied Invoice Number.")
        return

    # Lista de CMs únicos da aba Match RELO
    cm_list = (
        df_match_relo["Transaction Number"]
        .dropna()
        .astype(str)
        .str.strip()
        .unique()
        .tolist()
    )
    if not cm_list:
        print("Nenhum CM encontrado em Match RELO.")
        return

    print(f"Rodando CM_APPLIED_QUERY para {len(cm_list)} CM(s) da aba Match RELO...")

    # Conecta no Oracle só para essa query (oracledb.init_oracle_client já foi chamado no main)
    with oracledb.connect(user=USERNAME, password=PASSWORD, dsn=DSN) as conn:
        with conn.cursor() as cur:
            placeholders = ", ".join(f":cm{i}" for i in range(len(cm_list)))
            sql_cm = CM_APPLIED_QUERY.format(placeholders=placeholders)
            params = {f"cm{i}": v for i, v in enumerate(cm_list)}
            cur.execute(sql_cm, params)
            rows = cur.fetchall()
            if not rows:
                print("CM_APPLIED_QUERY não retornou linhas.")
                return
            cols = [d[0] for d in cur.description]
            df_cm = pd.DataFrame(rows, columns=cols)

    # Garantir tipos string nas colunas-chave
    df_cm["Transaction Number"] = df_cm["Transaction Number"].astype(str).str.strip()
    df_cm["Applied Invoice Number"] = (
        df_cm["Applied Invoice Number"].astype(str).str.strip()
    )

    # Referência original da query de Relo CMs (Match RELO já tem a coluna Reference)
    df_ref = df_match_relo.copy()
    df_ref["Transaction Number"] = df_ref["Transaction Number"].astype(str).str.strip()
    df_ref["Reference_Invoice"] = (
        df_ref["Reference"]
        .apply(_extract_invoice_from_reference)
        .astype(str)
        .str.strip()
    )

    # Trazer Applied Invoice Number para as linhas de Match RELO
    df_applied_small = (
        df_cm[["Transaction Number", "Applied Invoice Number"]]
        .drop_duplicates(subset=["Transaction Number"])
        .copy()
    )

    df_check = df_ref.merge(
        df_applied_small,
        on="Transaction Number",
        how="left",
    )

    # Considerar "já revertido" apenas quando existe Applied Invoice Number
    # e ele NÃO é igual ao número de invoice que está na Reference
    mask_has_applied = df_check["Applied Invoice Number"].notna() & (
        df_check["Applied Invoice Number"].astype(str).str.strip() != ""
    )
    mask_diff = (
        df_check["Reference_Invoice"].astype(str).str.strip()
        != df_check["Applied Invoice Number"].astype(str).str.strip()
    )
    df_reversed = df_check[mask_has_applied & mask_diff].copy()

    if df_reversed.empty:
        print(
            "Nenhum CM da aba Match RELO foi identificado como 'Already Reversed' "
            "(Reference bate com Applied Invoice Number ou não há aplicação)."
        )
        return

    # Marca explicitamente o status
    df_reversed["Status"] = "CMs Already Reversed"

    # Grava nova aba na planilha-mãe
    with pd.ExcelWriter(
        coa_path, engine="openpyxl", mode="a", if_sheet_exists="replace"
    ) as w:
        df_reversed.to_excel(w, sheet_name="CMs Already Reversed", index=False)

    print(
        f"Aba 'CMs Already Reversed' criada com {len(df_reversed)} linha(s) "
        "onde Reference != Applied Invoice Number."
    )

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
        print("No rows returned for WO.\nExiting.")
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

    # Payment date from column T (20th column), if it exists
    if df_coa.shape[1] > 19:
        payment_raw = df_coa.iloc[:, 19]
        df_coa_match["PAYMENT_DATE"] = pd.to_datetime(payment_raw, errors="coerce")
    else:
        df_coa_match["PAYMENT_DATE"] = pd.NaT

    # ========= WO MATCHES (SUMIFS by TRANSACTION NUMBER) + CSV =========
    req_wo = [
        "Transaction Number",
        "Transaction Type",
        "Transaction Date",
        "Account Number",
        "Entered Amount",
        "Entered Currency",
        "Reason Code",
    ]
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

        # Exclude Transaction Numbers whose Transaction Type contains 'GROUP' but NOT 'XLR'
        mask_group = df_wo_m["Transaction Type"].str.contains("GROUP", case=False, na=False)
        mask_xlr = df_wo_m["Transaction Type"].str.contains("XLR", case=False, na=False)
        group_only_trx = df_wo_m.loc[mask_group & ~mask_xlr, "Transaction Number"].unique()
        if len(group_only_trx) > 0:
            print(
                f"Excluding {len(group_only_trx)} Transaction Number(s) whose "
                "Transaction Type contains 'GROUP' but not 'XLR' from WO matching."
            )
        df_wo_m = df_wo_m[~df_wo_m["Transaction Number"].isin(group_only_trx)].copy()

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
                "PAYMENT_DATE",
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
                "Reason Code",
            ]
            df_matches_agg = (
                matches_wo
                .groupby(group_cols, as_index=False)
                .agg({
                    "receipt_amount": "first",   # same value in the group
                    "sum_amount": "first",       # SUMIFS by TRX
                    "Entered Amount": "sum",     # sum of adjustments for that invoice
                    "RECEIPT_STATUS": "first",
                    "Transaction Date": "first",
                    "PAYMENT_DATE": "first",
                    "Reason Code": "first",
                })
            )

            # Days between payment and invoice
            if "PAYMENT_DATE" in df_matches_agg.columns and "Transaction Date" in df_matches_agg.columns:
                df_matches_agg["Days Between Payment and Invoice"] = (
                    (df_matches_agg["PAYMENT_DATE"] - df_matches_agg["Transaction Date"]).dt.days
                )
            else:
                df_matches_agg["Days Between Payment and Invoice"] = pd.NA

            # Split negative days APENAS para aba separada (sem remover de df_matches_agg)
            neg_mask = df_matches_agg["Days Between Payment and Invoice"] < 0
            df_negative_days = df_matches_agg[neg_mask].copy()

            df_matches = df_matches_agg.rename(columns={
                "receipt_amount": "LOCAL_RECEIPT_AMOUNT",
                "sum_amount": "WO_SUMIFS",
                "CURRENCY_CODE": "Receipt Currency Code",
                "Entered Currency": "WO Currency",
            })

            # --- Split DIR vs non-DIR for CSV and NL review ---
            dir_mask = df_matches["Transaction Type"].astype(str).str.contains("DIR", case=False, na=False)
            df_matches_dir = df_matches[dir_mask].copy()
            df_nl_review = df_matches[~dir_mask].copy()

            # --- UNKNOWN / UNKNWON OID currency matches ---
            df_unknown_matches = None
            unknown_mask = df_coa_match["CUSTOMER_NBR"].astype(str).str.upper().isin(["UNKNOWN", "UNKNWON"])
            df_unknown = df_coa_match[unknown_mask].copy()
            if not df_unknown.empty:
                df_unknown_tmp = df_unknown[["RECEIPT_NUMBER", "receipt_amount", "CURRENCY_CODE"]].copy()
                df_unknown_tmp = df_unknown_tmp.rename(columns={"CURRENCY_CODE": "Entered Currency"})
                merged_unknown = df_unknown_tmp.merge(
                    df_wo_sum_trx[["Account Number", "Transaction Number", "Entered Currency", "sum_amount"]],
                    on="Entered Currency",
                    how="inner",
                )
                cond_unknown = merged_unknown["receipt_amount"] == (-merged_unknown["sum_amount"])
                df_unknown_matches = merged_unknown.loc[cond_unknown].drop_duplicates()
            else:
                df_unknown_matches = None

            # Write Matches, NL review, Unknown OID Matches and Negative Days back into COA
            with pd.ExcelWriter(coa_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
                df_matches.to_excel(w, sheet_name="Matches", index=False)
                print(f"Sheet 'Matches' created with {len(df_matches)} rows (1 row per RECEIPT_NUMBER).")

                if not df_nl_review.empty:
                    df_nl_review.to_excel(w, sheet_name="NL review", index=False)
                    print(f"Sheet 'NL review' created with {len(df_nl_review)} rows (non-DIR Transaction Type).")

                if df_unknown_matches is not None and not df_unknown_matches.empty:
                    df_unknown_matches.to_excel(w, sheet_name="Unknown OID Matches", index=False)
                    print(f"Sheet 'Unknown OID Matches' created with {len(df_unknown_matches)} rows.")

                if not df_negative_days.empty:
                    df_negative_days_renamed = df_negative_days.rename(columns={
                        "receipt_amount": "LOCAL_RECEIPT_AMOUNT",
                        "sum_amount": "WO_SUMIFS",
                        "CURRENCY_CODE": "Receipt Currency Code",
                        "Entered Currency": "WO Currency",
                    })
                    df_negative_days_renamed.to_excel(w, sheet_name="Negative Days", index=False)
                    print(f"Sheet 'Negative Days' created with {len(df_negative_days_renamed)} rows (payment < invoice date).")

            # Dates for file naming and comments
            today = date.today()
            today_csv_name = today.strftime("%m%d%Y")
            today_comment = today.strftime("%m/%d/%Y")

            # Separate file for Cash (mirrors Matches; keeps all rows)
            cash_xlsx_name = f"COAvsWO_Cash_{today_csv_name}.xlsx"
            cash_xlsx_path = OUTPUT_FOLDER / cash_xlsx_name
            df_matches.to_excel(cash_xlsx_path, index=False)
            print(f"Cash file generated with {len(df_matches)} rows at:\n{cash_xlsx_path}")

            # COAvsWO CSV from aggregated matches, DIR only
            df_matches_agg_dir = df_matches_agg[dir_mask.values].copy()
            if df_matches_agg_dir.empty:
                print("No DIR transactions found in Matches.\nCSV will not be generated.")
                csv_path = None
            else:
                df_csv = pd.DataFrame()
                df_csv["Index"] = range(1, len(df_matches_agg_dir) + 1)

                oper_units = df_matches_agg_dir["Transaction Type"].map(map_operating_unit)
                df_csv["Operating Unit"] = oper_units
                df_csv["Transaction Number"] = df_matches_agg_dir["Transaction Number"]
                df_csv["BFB Number"] = ""
                df_csv["Activity Name"] = oper_units.map(map_activity_name)
                df_csv["Adjustment Type"] = "Line"
                # positive value (invert sign of Entered Amount)
                df_csv["Amount to be Adjusted"] = (-df_matches_agg_dir["Entered Amount"]).round(2)
                # Normalized Reason from WO Reason Code
                df_csv["Reason"] = df_matches_agg_dir["Reason Code"].apply(normalize_reason)
                df_csv["Comments"] = "COA vs WO " + today_comment
                df_csv["GL Date"] = ""
                df_csv["Adjust Date"] = ""

                # Drop any rows without a Transaction Number (defensive)
                df_csv = df_csv[
                    df_csv["Transaction Number"].notna()
                    & (df_csv["Transaction Number"].astype(str).str.strip() != "")
                ].reset_index(drop=True)

                csv_name = f"COAvsWO{today_csv_name}.csv"
                csv_path = OUTPUT_FOLDER / csv_name
                df_csv.to_csv(csv_path, index=False, encoding="cp1252")
                print(f"CSV generated with {len(df_csv)} rows at:\n{csv_path}")

    # ========= RELO MATCHES (SUMIFS by TRANSACTION NUMBER) + RELO DM / 3 HC =========
    relo_dm_path = None
    hc_dm_path = None
    df_match_relo = None

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
                if "cash_xlsx_path" in locals() and cash_xlsx_path is not None:
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

                # Create 3 HC Relo Upload DM .xlsm file (Transaction Type != TS_CM_RELO)
                try:
                    hc_dm_path = build_hc_relo_dm_file(df_match_relo, HC_DM_TEMPLATE, OUTPUT_FOLDER)
                    if hc_dm_path:
                        print(f"3 HC Relo Upload DM (.xlsm) created: {hc_dm_path}")
                except Exception as e:
                    print(f"Failed to generate 3 HC Relo Upload DM file: {e}")

                # Comparar Reference x Applied Invoice Number
                try:
                    check_relo_cms_already_reversed(df_match_relo, coa_path)
                except Exception as e:
                    print(f"Falha ao rodar checagem de 'CMs Already Reversed': {e}")

    # Send email via Outlook with output files (including RELO DM / 3 HC DM if generated)
    attachments = []

    if "cash_xlsx_path" in locals() and cash_xlsx_path is not None:
        attachments.append(cash_xlsx_path)
    if "csv_path" in locals() and csv_path is not None:
        attachments.append(csv_path)
    if "relo_dm_path" in locals() and relo_dm_path:
        attachments.append(relo_dm_path)
    if "hc_dm_path" in locals() and hc_dm_path:
        attachments.append(hc_dm_path)

    if attachments:
        today_comment = date.today().strftime("%m/%d/%Y")
        email_to = "hotelcollectbilling@expedia.com"
        email_subject = f"COA vs WO {today_comment}"
        email_body = (
            f"Hi team,\n\n"
            f"Please see attached the COA vs WO outputs for {today_comment}.\n\n"
            f"Best regards,\n"
            f"Global Billing\n"
        )
        send_outlook_email_with_attachments(email_to, email_subject, email_body, attachments)

if __name__ == "__main__":
    main()
