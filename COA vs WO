import os
from pathlib import Path
from datetime import date

import oracledb
import pandas as pd
from openpyxl import load_workbook

# Tentativa de import do Outlook (pywin32) para envio de e-mail automático
try:
    import win32com.client as win32
except ImportError:
    win32 = None

# ==========================
# CONFIG GERAL
# ==========================

DOWNLOADS_DIR = Path(os.path.join(os.path.expanduser("~"), "Downloads"))

CONCAT_PATH = Path(
    r"C:\Users\josenjr\OneDrive - Expedia Group\Desktop\Daily activities\COA vs WO Process\Concat.xlsx"
)

BASE_OUTPUT_DIR = Path(
    r"C:\Users\josenjr\OneDrive - Expedia Group\Desktop\Daily activities\COA vs WO Process"
)
today_str = date.today().strftime("%m.%d.%Y")  # pasta no formato MM.DD.YYYY
OUTPUT_FOLDER = BASE_OUTPUT_DIR / today_str
OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

# Template RELO (mantém macros)
RELO_DM_TEMPLATE = Path(
    r"C:\Users\josenjr\OneDrive - Expedia Group\Desktop\Daily activities\COA vs WO Process\4 Non-Lodging Relo CM.xlsm"
)

# ==========================
# CONFIG ORACLE
# ==========================

ORACLE_CLIENT_DIR = r"C:\Users\josenjr\OneDrive - Expedia Group\Desktop\Oracle Instant Client\instantclient_23_0"
USERNAME = "josenjr"
PASSWORD = "eyX057UWzLnZTl3w"      # TROCAR
DSN      = "ashworaebsdb02-vip.datawarehouse.expecn.com:1526/ORAPRD_UI"

# WO (ajustes)
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

# RELO CMs – COM Reference
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
# FUNÇÕES AUXILIARES
# ==========================

def find_latest_coa_file() -> Path:
    candidates = []
    for f in DOWNLOADS_DIR.iterdir():
        if f.is_file() and f.name.startswith("RPA-306-001 Cash On Account"):
            candidates.append(f)
    if not candidates:
        raise FileNotFoundError("Nenhum arquivo 'RPA-306-001 Cash On Account*' encontrado em Downloads.")
    return max(candidates, key=lambda p: p.stat().st_mtime)

def chunk_list(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

def update_concat_column_b(customers):
    wb = load_workbook(CONCAT_PATH)
    ws = wb.active  # 1ª aba

    start_row = 2  # cabeçalho
    for i, cust in enumerate(customers, start=start_row):
        ws.cell(row=i, column=2, value=cust)

    max_row = start_row + len(customers) - 1
    template_row = 2

    # copiar fórmulas da linha 2 para baixo
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=template_row, column=col)
        if isinstance(cell.value, str) and cell.value.startswith("="):
            formula = cell.value
            for r in range(template_row + 1, max_row + 1):
                if ws.cell(row=r, column=col).value is None:
                    ws.cell(row=r, column=col).value = formula

    wb.save(CONCAT_PATH)
    print(f"Concat.xlsx atualizado (coluna B + fórmulas até a linha {max_row}).")

def map_operating_unit(trx_type: str) -> str:
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
    Envia um e-mail via Outlook (se pywin32 e Outlook estiverem disponíveis).
    attachments: lista de Paths
    """
    if win32 is None:
        print("pywin32 não está instalado. Não foi possível enviar e-mail automático pelo Outlook.")
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
                print(f"Aviso: anexo não encontrado, não será incluído: {att_path}")

        mail.Send()
        print(f"E-mail enviado para {to_addr} com assunto '{subject}'.")
    except Exception as e:
        print(f"Falha ao enviar e-mail pelo Outlook: {e}")

# ==========================
# NOVO PASSO: construir .xlsm RELO DM (TS_CM_RELO -> TS_DM_RELO)
# ==========================

def _find_col_letter_by_header(ws, header_text, search_rows=range(1, 20)):
    """
    Procura o texto exato do cabeçalho nas primeiras linhas e retorna a letra da coluna.
    """
    header_text = str(header_text).strip().lower()
    for r in search_rows:
        for cell in ws[r]:
            val = cell.value
            if val is None:
                continue
            if str(val).strip().lower() == header_text:
                return cell.column_letter
    raise KeyError(f"Cabeçalho '{header_text}' não encontrado no template.")

def _copy_default_from_row(ws, col_letter, from_row, to_row):
    ws[f"{col_letter}{to_row}"] = ws[f"{col_letter}{from_row}"].value

def build_ts_relo_dm_file(df_match_relo, template_path: Path, out_dir: Path):
    """
    Gera arquivo .xlsm a partir do template, com linhas somente dos MATCHES RELO de TS_CM_RELO,
    transformando o tipo em TS_DM_RELO e preenchendo as colunas conforme regras.
    Mantém macros/estrutura do template (.xlsm) usando keep_vba=True.

    Pré-requisitos em df_match_relo:
      - colunas: ['RECEIPT_NUMBER','LOCAL_RECEIPT_AMOUNT','Transaction Number','Account Number','Transaction Type']
    """
    # 1) Filtrar apenas TS_CM_RELO
    if df_match_relo is None or df_match_relo.empty:
        print("Match RELO vazio. Nenhuma linha para o arquivo RELO DM.")
        return None

    df_src = (
        df_match_relo
        .loc[lambda d: d["Transaction Type"].astype(str).str.upper() == "TS_CM_RELO"]
        .copy()
    )
    if df_src.empty:
        print("Nenhum MATCH RELO do tipo TS_CM_RELO. Nada a gerar.")
        return None

    # Consolidar 1 linha por (receipt, trx, account)
    df_src = (
        df_src
        .drop_duplicates(subset=["RECEIPT_NUMBER", "Transaction Number", "Account Number"])
        .loc[:, ["RECEIPT_NUMBER", "Transaction Number", "Account Number", "LOCAL_RECEIPT_AMOUNT"]]
        .copy()
    )
    # Valor positivo
    df_src["AMOUNT_POS"] = df_src["LOCAL_RECEIPT_AMOUNT"].abs().round(2)

    # 2) Abrir template .xlsm com macros
    if not template_path.exists():
        raise FileNotFoundError(f"Template .xlsm não encontrado: {template_path}")
    wb = load_workbook(template_path, keep_vba=True, data_only=False)
    if "EVC RELO CM" not in wb.sheetnames:
        raise ValueError("A planilha 'EVC RELO CM' não foi encontrada no template.")
    ws = wb["EVC RELO CM"]  # contém cabeçalhos e defaults na linha 10

    # 3) Mapear colunas por cabeçalho (robusto)
    col_customer     = _find_col_letter_by_header(ws, "Customer Number")
    col_shipto       = _find_col_letter_by_header(ws, "Ship To Cust Numb")
    col_trx_date     = _find_col_letter_by_header(ws, "Transaction Date")
    col_txn_type     = _find_col_letter_by_header(ws, "Txn Type")
    col_comments     = _find_col_letter_by_header(ws, "Comments")
    col_reference    = _find_col_letter_by_header(ws, "Reference")
    col_context      = _find_col_letter_by_header(ws, "Context")

    # Colunas de linha (valores):
    col_line_number  = _find_col_letter_by_header(ws, "Line Number")  # sequencial
    col_unit_price   = "AM"   # Unit Selling Price (pedido explícito: AM)
    col_qty          = "AN"   # Quantity (pedido explícito: AN)
    col_total_amt    = "AP"   # Total Amt (pedido explícito: AP)

    # Colunas que devem preservar a informação da linha 10 (por letra):
    preserve_letters = ["T", "V", "X", "AH", "AI", "AJ", "AK", "AQ", "AR", "AS"]

    # 4) Inserir linhas a partir da 11 (linha 10 = defaults do template)
    start_row = 11
    today_dt = date.today()

    for i, row in enumerate(df_src.itertuples(index=False), start=0):
        receipt_number   = getattr(row, "RECEIPT_NUMBER")
        trx_number       = getattr(row, "Transaction_Number")
        account_number   = getattr(row, "Account_Number")
        amount_pos       = float(getattr(row, "AMOUNT_POS"))

        r = start_row + i

        # Cabeçalhos básicos
        ws[f"{col_customer}{r}"] = str(account_number)
        ws[f"{col_shipto}{r}"]   = str(account_number)
        ws[f"{col_trx_date}{r}"] = today_dt  # data de hoje

        # Txn Type = TS_DM_RELO (inclui coluna P explicitamente)
        ws[f"{col_txn_type}{r}"] = "TS_DM_RELO"
        ws[f"P{r}"]              = "TS_DM_RELO"  # pedido: “coluna P deve ser TS_DM_RELO”

        # Comments com FÓRMULA
        ws[f"{col_comments}{r}"] = '="COA vs WO " & TEXT(TODAY(),"mm/dd/yyyy")'

        # Reference
        ws[f"{col_reference}{r}"] = f"DM ISSUED TO OFFSET {trx_number}"

        # Copiar defaults da LINHA 10 para as colunas pedidas
        for colL in preserve_letters:
            _copy_default_from_row(ws, colL, 10, r)
        # Context também preserva default da linha 10
        _copy_default_from_row(ws, col_context, 10, r)

        # Valores da linha (positivos)
        ws[f"{col_unit_price}{r}"] = amount_pos  # AM
        ws[f"{col_qty}{r}"]        = 1           # AN
        ws[f"{col_total_amt}{r}"]  = amount_pos  # AP

        # Line Number sequencial (1..N)
        ws[f"{col_line_number}{r}"] = i + 1

    # 5) Salvar .xlsm na pasta do dia
    out_name = f"4 Non-Lodging Relo DM_{today_dt.strftime('%m%d%Y')}.xlsm"
    out_path = out_dir / out_name
    wb.save(out_path)
    print(f"Arquivo RELO DM gerado: {out_path}")
    return out_path

# ==========================
# PIPELINE COMPLETO
# ==========================

def main():
    # 1) COA mais recente
    coa_path = find_latest_coa_file()
    print(f"Arquivo COA encontrado: {coa_path}")

    # 2) Ler aba Data e extrair contas
    df_coa = pd.read_excel(coa_path, sheet_name="Data")
    if "CUSTOMER_NBR" not in df_coa.columns:
        raise ValueError("Coluna CUSTOMER_NBR não encontrada na aba 'Data' do COA.")

    customers = (
        df_coa["CUSTOMER_NBR"]
        .dropna().astype(str).str.strip()
        .loc[lambda s: s != ""].unique().tolist()
    )
    print(f"Total de CUSTOMER_NBR distintos: {len(customers)}")

    # CustList
    df_cust = pd.DataFrame({"CUSTOMER_NBR": customers})
    with pd.ExcelWriter(coa_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        df_cust.to_excel(w, sheet_name="CustList", index=False)
    print("Aba 'CustList' criada/atualizada no COA.")

    # Atualizar Concat.xlsx
    update_concat_column_b(customers)

    # 3) Conectar Oracle (THICK)
    print("Inicializando Oracle Client...")
    oracledb.init_oracle_client(lib_dir=ORACLE_CLIENT_DIR)

    all_wo_rows, wo_columns = [], None
    all_relo_rows, relo_columns = [], None

    print("Conectando ao Oracle...")
    with oracledb.connect(user=USERNAME, password=PASSWORD, dsn=DSN) as conn:
        print("Conectado.")
        with conn.cursor() as cur:
            cur.execute("ALTER SESSION SET CURRENT_SCHEMA = APPS")
            print("ALTER SESSION SET CURRENT_SCHEMA = APPS executado.")

        total_chunks = (len(customers) + 999) // 1000

        for idx, acc_chunk in enumerate(chunk_list(customers, 1000), start=1):
            print(f"\n[Lote {idx}/{total_chunks}] Contas: {len(acc_chunk)}")
            print("Primeiras contas:", acc_chunk[:5])

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
                print(f"  -> WO: {len(rows)} linhas")
            except oracledb.DatabaseError as e:
                error, = e.args
                print(f"  !! ERRO WO no lote {idx}: {error.message}")
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
                print(f"  -> RELO: {len(rows)} linhas")
            except oracledb.DatabaseError as e:
                error, = e.args
                print(f"  !! ERRO RELO no lote {idx}: {error.message}")
                break

    if not all_wo_rows:
        print("Nenhuma linha em WO. Encerrando.")
        return

    df_wo = pd.DataFrame(all_wo_rows, columns=wo_columns)
    with pd.ExcelWriter(coa_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        df_wo.to_excel(w, sheet_name="WO", index=False)
    print(f"Aba 'WO' criada com {len(df_wo)} linhas.")

    df_relo = None
    if all_relo_rows:
        df_relo = pd.DataFrame(all_relo_rows, columns=relo_columns)
        with pd.ExcelWriter(coa_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
            df_relo.to_excel(w, sheet_name="Relo CMs", index=False)
        print(f"Aba 'Relo CMs' criada com {len(df_relo)} linhas.")
    else:
        print("Nenhuma linha retornada em Relo CMs.")

    # 4) Normalizar Data
    req_data = ["RECEIPT_NUMBER", "LOCAL_RECEIPT_AMOUNT", "CUSTOMER_NBR", "CURRENCY_CODE", "RECEIPT_STATUS"]
    miss_data = [c for c in req_data if c not in df_coa.columns]
    if miss_data:
        print(f"Faltam colunas em Data: {miss_data}.")
        return

    df_coa_match = df_coa.copy()
    df_coa_match["CUSTOMER_NBR"] = df_coa_match["CUSTOMER_NBR"].astype(str).str.strip()
    df_coa_match["LOCAL_RECEIPT_AMOUNT"] = pd.to_numeric(df_coa_match["LOCAL_RECEIPT_AMOUNT"], errors="coerce")
    df_coa_match["CURRENCY_CODE"] = df_coa_match["CURRENCY_CODE"].astype(str).str.strip()
    df_coa_match["receipt_amount"] = df_coa_match["LOCAL_RECEIPT_AMOUNT"].round(2)

    # ========= MATCHES WO (SUMIFS por TRANSACTION NUMBER) + CSV =========
    req_wo = ["Transaction Number", "Transaction Type", "Account Number", "Entered Amount", "Entered Currency"]
    miss_wo = [c for c in req_wo if c not in df_wo.columns]
    if miss_wo:
        print(f"Faltam colunas em WO: {miss_wo}. Pulando Matches/CSV.")
    else:
        df_wo_m = df_wo.copy()
        df_wo_m["Account Number"] = df_wo_m["Account Number"].astype(str).str.strip()
        df_wo_m["Entered Amount"] = pd.to_numeric(df_wo_m["Entered Amount"], errors="coerce")
        df_wo_m["Entered Currency"] = df_wo_m["Entered Currency"].astype(str).str.strip()
        df_wo_m["Transaction Type"] = df_wo_m["Transaction Type"].astype(str)

        # Excluir qualquer TRANSACTION NUMBER que tenha TYPE contendo 'GROUP'
        group_trx = df_wo_m.loc[
            df_wo_m["Transaction Type"].str.contains("GROUP", case=False, na=False),
            "Transaction Number"
        ].unique()

        if len(group_trx) > 0:
            print(f"Excluindo {len(group_trx)} Transaction Number(s) com Transaction Type contendo 'GROUP' do match de WO.")
            df_wo_m = df_wo_m[~df_wo_m["Transaction Number"].isin(group_trx)].copy()

        # SUMIFS por (Account Number, Transaction Number, Currency)
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
            how="inner"
        )

        # RECEIPT (positivo) == - SUMIFS(WO por TRX) (negativo)
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
            ]
        ].drop_duplicates()

        if keys_ok_trx.empty:
            print("Nenhum match WO por Transaction Number.")
        else:
            # garantir apenas um match por RECEIPT_NUMBER na aba Matches
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
                    f"Receipts com múltiplos matches WO encontrados. "
                    f"Antes: {before_len} linhas, depois de limitar 1 por RECEIPT_NUMBER: {after_len}."
                )

            # matches_wo ainda no nível de linha de ajuste
            matches_wo = df_wo_m.merge(
                keys_ok_trx,
                on=["Account Number", "Transaction Number", "Entered Currency"],
                how="inner"
            )

            # Agregar por RECEIPT + TRANSACTION NUMBER (remover duplicados de ajuste)
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
                    "receipt_amount": "first",   # mesmo valor dentro do grupo
                    "sum_amount": "first",       # SUMIFS por TRX
                    "Entered Amount": "sum",     # soma dos ajustes dessa invoice
                    "RECEIPT_STATUS": "first",
                })
            )

            df_matches = df_matches_agg.rename(columns={
                "receipt_amount": "LOCAL_RECEIPT_AMOUNT",
                "sum_amount": "WO_SUMIFS",
                "CURRENCY_CODE": "Receipt Currency Code",
                "Entered Currency": "WO Currency"
            })

            # Gravar aba Matches no próprio COA
            with pd.ExcelWriter(coa_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
                df_matches.to_excel(w, sheet_name="Matches", index=False)
            print(f"Aba 'Matches' criada com {len(df_matches)} linhas (1 linha por RECEIPT_NUMBER).")

            # Datas para nome dos arquivos e comentário
            today = date.today()
            today_csv_name = today.strftime("%m%d%Y")
            today_comment = today.strftime("%m/%d/%Y")

            # Arquivo separado para Cash, espelhando a aba Matches
            cash_xlsx_name = f"COAvsWO_Cash_{today_csv_name}.xlsx"
            cash_xlsx_path = OUTPUT_FOLDER / cash_xlsx_name
            df_matches.to_excel(cash_xlsx_path, index=False)
            print(f"Arquivo para Cash gerado com {len(df_matches)} linhas em:\n{cash_xlsx_path}")

            # CSV COAvsWO a partir do agregado (1 linha por TRX / RECEIPT)
            df_csv = pd.DataFrame()
            df_csv["Index"] = range(1, len(df_matches_agg) + 1)

            oper_units = df_matches_agg["Transaction Type"].map(map_operating_unit)
            df_csv["Operating Unit"] = oper_units
            df_csv["Transaction Number"] = df_matches_agg["Transaction Number"]
            df_csv["BFB Number"] = ""
            df_csv["Activity Name"] = oper_units.map(map_activity_name)
            df_csv["Adjustment Type"] = "Line"
            # valor positivo (invertendo o sinal do Entered Amount)
            df_csv["Amount to be Adjusted"] = (-df_matches_agg["Entered Amount"]).round(2)
            df_csv["Reason"] = "Uneconomical to Collect"
            df_csv["Comments"] = "COA vs WO " + today_comment
            df_csv["GL Date"] = ""
            df_csv["Adjust Date"] = ""

            csv_name = f"COAvsWO{today_csv_name}.csv"
            csv_path = OUTPUT_FOLDER / csv_name
            df_csv.to_csv(csv_path, index=False, encoding="cp1252")
            print(f"CSV gerado com {len(df_csv)} linhas em:\n{csv_path}")

            # ========= MATCH RELO (SUMIFS por TRANSACTION NUMBER) =========
            relo_dm_path = None
            if df_relo is not None:
                req_relo = ["Transaction Number", "Transaction Type",
                            "Account Number", "Entered Amount",
                            "Entered Currency", "Reference"]
                miss_relo = [c for c in req_relo if c not in df_relo.columns]
                if miss_relo:
                    print(f"Faltam colunas em Relo CMs: {miss_relo}.")
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
                        how="inner"
                    )

                    cond_relo_trx = merged_relo_trx["receipt_amount"] == (-merged_relo_trx["sum_amount"])
                    keys_relo_trx = merged_relo_trx.loc[
                        cond_relo_trx,
                        ["CUSTOMER_NBR", "CURRENCY_CODE", "receipt_amount", "RECEIPT_NUMBER",
                         "Account Number", "Transaction Number", "Entered Currency", "sum_amount"]
                    ].drop_duplicates()

                    if keys_relo_trx.empty:
                        print("Nenhum match RELO por Transaction Number.")
                    else:
                        matches_relo = df_relo_m.merge(
                            keys_relo_trx,
                            on=["Account Number", "Transaction Number", "Entered Currency"],
                            how="inner"
                        )

                        df_match_relo = matches_relo[[
                            "RECEIPT_NUMBER",
                            "receipt_amount",
                            "Transaction Number",
                            "Entered Amount",
                            "CURRENCY_CODE",
                            "Entered Currency",
                            "Reference",
                            "Account Number",        # ADICIONADO
                            "Transaction Type"       # ADICIONADO
                        ]].rename(columns={
                            "receipt_amount": "LOCAL_RECEIPT_AMOUNT",
                            "CURRENCY_CODE": "Receipt Currency Code",
                            "Entered Currency": "WO Currency"
                        })

                        with pd.ExcelWriter(coa_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
                            df_match_relo.to_excel(w, sheet_name="Match RELO", index=False)

                        print(f"Aba 'Match RELO' criada com {len(df_match_relo)} linhas.")

                        # ---- NOVO ARQUIVO .xlsm RELO DM (somente TS_CM_RELO) ----
                        try:
                            relo_dm_path = build_ts_relo_dm_file(df_match_relo, RELO_DM_TEMPLATE, OUTPUT_FOLDER)
                            if relo_dm_path:
                                print(f"RELO DM (.xlsm) criado: {relo_dm_path}")
                        except Exception as e:
                            print(f"Falha ao gerar arquivo RELO DM: {e}")

            # Enviar e-mail pelo Outlook com os arquivos de saída (INCLUINDO o RELO DM, se gerado)
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
