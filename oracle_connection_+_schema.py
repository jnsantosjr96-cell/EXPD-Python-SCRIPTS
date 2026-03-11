import oracledb
import pandas as pd
from pathlib import Path

# 1) THICK MODE (Instant Client)
oracledb.init_oracle_client(
    lib_dir=r"C:\Users\josenjr\OneDrive - Expedia Group\Desktop\Oracle Instant Client\instantclient_23_0"
)

# 2) Credenciais e DSN
USERNAME = "josenjr"
PASSWORD = "INSERT PASSWORD"  # ideal é depois colocar em variável de ambiente
DSN      = "ashworaebsdb02-vip.datawarehouse.expecn.com:1526/ORAPRD_UI"

# 3) Caminho de saída
output_dir = Path(r"C:\Users\josenjr\Downloads\Brazil report")
output_dir.mkdir(parents=True, exist_ok=True)
output_file = output_dir / "Brazil_AR_2025.xlsx"

# 4) Query completa
SQL = """
select
    hca.account_number "OID",
    rctl.attribute12 "Booking ID",
    rct.trx_number "AR Transaction",
    APSA.AMOUNT_DUE_REMAINING "AR Open Balance",
    rct.invoice_currency_code "AR Currency",
    rctl.INVENTORY_ITEM_ID "Item Code",
    CASE
      WHEN rctl.INVENTORY_ITEM_ID in ('16')
      THEN 'HLC1955'
      WHEN rctl.INVENTORY_ITEM_ID in ('160011')
      THEN 'HLX1955'
      WHEN rctl.INVENTORY_ITEM_ID in ('43003')
      THEN 'EXPWA_COMP_CM'
      WHEN rctl.INVENTORY_ITEM_ID in ('181011')
      THEN 'EXPWA_COMP_CM'
      WHEN rctl.INVENTORY_ITEM_ID in ('179016')
      THEN 'HLC19113'
      WHEN rctl.INVENTORY_ITEM_ID in ('179017')
      THEN 'HLX19113'
      ELSE 'NOT_DEFINED'
    END AS "Item",
    rctl.description "Description",
    extended_amount "Unit Price",
    rcta.type "Transaction Class",
    rcta.name "Transaction Type",
    rct.ATTRIBUTE5 "Hotel Category",
    rct.ATTRIBUTE4 "Transaction Subtype",
    rct.TRX_DATE "Transaction Date",
    rct.INTERFACE_HEADER_ATTRIBUTE1 "Transaction Reference",
    rct.reason_code "CM Reason",
    rct.comments "Comments",
    rctl.INTERFACE_LINE_CONTEXT "Context Value",
    rctl.interface_line_attribute15 "Business Model",
    rctl.ATTRIBUTE1 "Management Unit",
    rctl.ATTRIBUTE2 "Traveler Name",
    rctl.ATTRIBUTE7 "Check-in date",
    rctl.ATTRIBUTE8 "Check-out date",
    rctl.ATTRIBUTE3 "Original Source Currency",
    rctl.ATTRIBUTE4 "Original Source Amount",
    rctl.ATTRIBUTE10 "Recon Date",
    rctl.attribute12 "Reservation ID"
from ra_customer_trx_all rct,
     ra_customer_trx_lines_all rctl,
     ra_cust_trx_types_all rcta,
     AR_PAYMENT_SCHEDULES_ALL APSA,
     hz_cust_accounts hca
where rct.customer_trx_id = rctl.customer_trx_id
  and rct.cust_trx_type_id = rcta.cust_trx_type_id
  and rct.CUSTOMER_TRX_ID = APSA.CUSTOMER_TRX_ID
  and rct.bill_to_customer_id = hca.cust_account_id
  and rct.TRX_DATE between '01-JAN-25' AND '31-DEC-25'
  and rcta.name in ('BR_DIR_INVOICE_BRL')
  and rctl.attribute12 is not NULL
order by rctl.attribute12, rct.creation_date, rct.trx_number
"""

# 5) Conexão + ALTER SESSION + leitura da query
with oracledb.connect(user=USERNAME, password=PASSWORD, dsn=DSN) as conn:
    with conn.cursor() as cur:
        cur.execute("ALTER SESSION SET CURRENT_SCHEMA = APPS")

    # Lê direto para DataFrame
    df = pd.read_sql(SQL, conn)

# 6) Exporta para Excel
df.to_excel(output_file, index=False)

print("Conectado! Thick mode?", not oracledb.is_thin_mode())
print(f"Linhas retornadas: {len(df)}")
print(f"Arquivo gerado em: {output_file}")
