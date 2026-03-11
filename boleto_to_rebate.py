import sys
from pathlib import Path
from datetime import date  # <--- NOVO

import oracledb
import pandas as pd

# ==========================
# IMPORT DO OUTLOOK (pywin32)
# ==========================
try:
    import win32com.client as win32
except ImportError as e:
    raise SystemExit(
        "ERRO: pywin32 não está instalado neste ambiente.\n"
        "No terminal do PyCharm, rode:\n"
        "  pip install pywin32\n"
        "  python -m pywin32_postinstall install\n"
        "Depois rode o script novamente."
    ) from e

# ==========================
# CONFIGURAÇÕES
# ==========================

# Oracle Instant Client
ORACLE_CLIENT_DIR = r"C:\Users\josenjr\OneDrive - Expedia Group\Desktop\Oracle Instant Client\instantclient_23_0"

# Credenciais Oracle
USERNAME = "josenjr"
PASSWORD = "INSERT PASSWORD"  # TROCAR: ideal ler de variável de ambiente
DSN      = "ashworaebsdb02-vip.datawarehouse.expecn.com:1526/ORAPRD_UI"

# Pasta/arquivo de saída
OUTPUT_FOLDER = Path(r"C:\Users\josenjr\OneDrive - Expedia Group\Desktop\Resultados_SQL")

_TODAY_STR = date.today().strftime("%Y-%m-%d")  # ex.: 2026-03-06
OUTPUT_FILE = OUTPUT_FOLDER / f"rebate_boleto_{_TODAY_STR}.xlsx"

# --- DATA DE REFERÊNCIA PARA O E-MAIL: DIA 5 DO MÊS CORRENTE ---
_REF_DATE = date.today().replace(day=5)              # ex.: 2026-02-05
_EMAIL_DATE_STR = _REF_DATE.strftime("%d/%m/%Y")     # ex.: "05/02/2026"

# Dados do e-mail
EMAIL_TO   = "hotelcollectbilling@expedia.com"
EMAIL_SUBJ = f"Rebate to Boleto - {_EMAIL_DATE_STR}"
EMAIL_BODY = (
    "Hello,\n\n"
    f"Please see attached Boleto to Rebate report created on {_EMAIL_DATE_STR},\n"
    "Thank you,\n"
    "Global Billing"
)

# Query Oracle
QUERY = """
SELECT 
    PSA.CLASS "CLASS",
    CTTA.NAME "TRX Type Name",
    HZP.PARTY_NAME "Customer Name",
    HCA.ACCOUNT_NUMBER "Account Number",
    HCA.ATTRIBUTE1 "Hotel ID",
    CTA.TRX_NUMBER "Transaction Number",
    CTA.TRX_DATE "TRX DATE",
    PSA.GL_DATE "GL Date",
    PSA.INVOICE_CURRENCY_CODE "Currency",
    PSA.AMOUNT_DUE_ORIGINAL "Entered Amt",
    PSA.AMOUNT_DUE_REMAINING "Open Balance",
    PSA.AMOUNT_DUE_ORIGINAL - PSA.AMOUNT_DUE_REMAINING "Rebate",
    CASE
       WHEN PSA.AMOUNT_DUE_REMAINING in ('0')
       THEN 'Write off boleto'
       ELSE 'Apply rebate'
    END AS "Billing actions"
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
WHERE
    GLL.LEDGER_CATEGORY_CODE = 'PRIMARY'
AND CTTA.NAME NOT LIKE '%EAC%'
AND GLL.LEDGER_NAME IN ('Brazil PL')
AND CTA.TRX_DATE BETWEEN '05-MAR-26' AND '05-MAR-26'
AND PSA.CLASS IN ('INV')
AND PSA.AMOUNT_DUE_ORIGINAL NOT LIKE PSA.AMOUNT_DUE_REMAINING
ORDER BY 13, 6
"""

# ==========================
# FUNÇÕES
# ==========================

def init_oracle():
    oracledb.init_oracle_client(lib_dir=ORACLE_CLIENT_DIR)

def run_query_and_save():
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

    with oracledb.connect(user=USERNAME, password=PASSWORD, dsn=DSN) as conn:
        # garante schema APPS
        with conn.cursor() as cur:
            cur.execute("ALTER SESSION SET CURRENT_SCHEMA = APPS")

        df = pd.read_sql(QUERY, conn)

    df.to_excel(OUTPUT_FILE, index=False)
    print(f"Linhas retornadas: {len(df)}")
    print(f"Arquivo salvo em: {OUTPUT_FILE}")

def send_outlook_email(to, subject, body, attachment_path):
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # olMailItem

    mail.To = to
    mail.Subject = subject
    mail.Body = body

    apath = Path(attachment_path)
    if apath.exists():
        mail.Attachments.Add(str(apath))

    # Se quiser revisar antes, troque para mail.Display()
    mail.Send()
    print("E-mail enviado via Outlook.")

# ==========================
# MAIN
# ==========================

if __name__ == "__main__":
    init_oracle()
    run_query_and_save()
    send_outlook_email(EMAIL_TO, EMAIL_SUBJ, EMAIL_BODY, OUTPUT_FILE)
