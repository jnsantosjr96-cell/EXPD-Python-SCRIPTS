
import os
import glob
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule

# Dynamic path to Downloads
downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")

# Search for specific Oracle file
oracle_files = glob.glob(os.path.join(downloads_folder, "*EXPD_AR_Aging_7_Buckets*.xls"))
if not oracle_files:
    raise FileNotFoundError("File EXPD_AR_Aging_7_Buckets not found in Downloads folder.")

oracle_path = max(oracle_files, key=os.path.getmtime)
print(f"Using file: {oracle_path}")

# Read file as HTML
df = pd.read_html(oracle_path, skiprows=19)[0]
print("File read as HTML.")

# Select correct columns by actual indexes
col_indices = [1, 5, 9, 10, 11, 14, 15, 16, 17]
df = df.iloc[:, col_indices]

# Rename to desired headers
columns_to_keep = [
    "Company", "Bill-To Number", "Transaction Type", "Transaction Class",
    "Transaction #", "Transaction Date", "Transactional Currency",
    "Transactional Outstanding Amount", "Aging Category"
]
df.columns = columns_to_keep

# Remove rows where Transaction Type = "payment"
df = df[df["Transaction Type"].astype(str).str.lower() != "payment"]

# ✅ Logic to remove Bill-To Numbers with only one type of Transaction Class
valid_bill_to = df.groupby("Bill-To Number")["Transaction Class"].apply(lambda x: len(set(x)) > 1)
df = df[df["Bill-To Number"].isin(valid_bill_to[valid_bill_to].index)]

# Sort according to logic
df = df.sort_values(by=["Bill-To Number", "Transaction Class", "Transaction Date", "Aging Category"])

# Create new clean workbook
wb = Workbook()
ws = wb.active
ws.title = "AR Aging Clean"

# Styles
header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
header_font = Font(color="FFFFFF", bold=True)
thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                     top=Side(style="thin"), bottom=Side(style="thin"))

# Headers
for col_num, col_name in enumerate(columns_to_keep, start=1):
    cell = ws.cell(row=1, column=col_num, value=col_name)
    cell.fill = header_fill
    cell.font = header_font
    cell.border = thin_border

# Insert data
for i, row in enumerate(df.values, start=2):
    for j, value in enumerate(row, start=1):
        cell = ws.cell(row=i, column=j, value=value)
        cell.alignment = Alignment(wrap_text=True)
        cell.border = thin_border

# Adjust column widths
for col_num in range(1, len(columns_to_keep) + 1):
    max_length = max(len(str(ws.cell(row=row, column=col_num).value)) for row in range(1, ws.max_row + 1))
    ws.column_dimensions[get_column_letter(col_num)].width = max_length + 2

# Format monetary values column
currency_col = 8  # "Transactional Outstanding Amount"
for row in range(2, ws.max_row + 1):
    ws.cell(row=row, column=currency_col).number_format = '#,##0.00'

# Freeze pane and add filter
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:{get_column_letter(len(columns_to_keep))}{ws.max_row}"

# ✅ Conditional formatting only if there is data
if ws.max_row > 1:
    aging_col_letter = get_column_letter(9)
    ws.conditional_formatting.add(
        f"{aging_col_letter}2:{aging_col_letter}{ws.max_row}",
        FormulaRule(
            formula=[f'ISNUMBER(SEARCH("Current",{aging_col_letter}2))'],
            fill=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        )
    )
    ws.conditional_formatting.add(
        f"{aging_col_letter}2:{aging_col_letter}{ws.max_row}",
        FormulaRule(
            formula=[f'ISNUMBER(SEARCH("1-30 Days",{aging_col_letter}2))'],
            fill=PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
        )
    )
    ws.conditional_formatting.add(
        f"{aging_col_letter}2:{aging_col_letter}{ws.max_row}",
        FormulaRule(
            formula=[f'ISNUMBER(SEARCH("181-360 Days",{aging_col_letter}2))'],
            fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        )
    )

# Save initial file
output_path = os.path.join(downloads_folder, "AR_Aging_Clean.xlsx")
wb.save(output_path)
print(f"✅ Processed file saved at: {output_path}")

# =========================
# ✅ NEW STEP: Matching CM to Invoices
# =========================

# Reload workbook and data
wb = load_workbook(output_path)
df = pd.DataFrame(df)  # Ensure DataFrame is ready

# Filter only Credit Memo and Invoice
df_match = df[df["Transaction Class"].isin(["Credit Memo", "Invoice"])].copy()
df_match["Transactional Outstanding Amount"] = pd.to_numeric(df_match["Transactional Outstanding Amount"], errors="coerce")

# Separate CMs and Invoices
credit_memos = df_match[df_match["Transaction Class"] == "Credit Memo"].copy()
invoices = df_match[df_match["Transaction Class"] == "Invoice"].copy()

# Sort for processing
credit_memos.sort_values(by=["Bill-To Number", "Transaction Date"], inplace=True)
invoices.sort_values(by=["Bill-To Number", "Transaction Date"], inplace=True)

matches = []

grouped_cm = credit_memos.groupby(["Bill-To Number", "Transactional Currency"])
grouped_inv = invoices.groupby(["Bill-To Number", "Transactional Currency"])

for (bill_to, currency), cm_group in grouped_cm:
    if (bill_to, currency) not in grouped_inv.groups:
        continue
    inv_group = grouped_inv.get_group((bill_to, currency))

    cm_list = cm_group[["Transaction #", "Transactional Outstanding Amount"]].values.tolist()
    inv_list = inv_group[["Transaction #", "Transactional Outstanding Amount"]].values.tolist()

    for cm_id, cm_amount in cm_list:
        cm_balance = abs(cm_amount)
        for inv_id, inv_amount in inv_list:
            if cm_balance <= 0:
                break
            if inv_amount <= 0:
                continue

            if cm_balance >= inv_amount:
                applied = inv_amount
                cm_balance -= inv_amount
                inv_amount = 0
            else:
                applied = cm_balance
                inv_amount -= cm_balance
                cm_balance = 0

            # Update invoice amount
            for i in range(len(inv_list)):
                if inv_list[i][0] == inv_id:
                    inv_list[i][1] = inv_amount
                    break

            matches.append([bill_to, cm_id, inv_id, applied, cm_balance])

# Create new sheet
ws2 = wb.create_sheet("CM_Invoice_Match")
headers = ["Bill-To Number", "Credit Memo ID", "Invoice ID", "Amount Applied", "Remaining Balance"]

# Headers
for col_num, header in enumerate(headers, start=1):
    cell = ws2.cell(row=1, column=col_num, value=header)
    cell.font = Font(bold=True)

# Insert data
for i, row in enumerate(matches, start=2):
    for j, value in enumerate(row, start=1):
        ws2.cell(row=i, column=j, value=value)

# Adjust column widths
for col_num in range(1, len(headers) + 1):
    max_length = max(len(str(ws2.cell(row=row, column=col_num).value)) for row in range(1, ws2.max_row + 1))
    ws2.column_dimensions[get_column_letter(col_num)].width = max_length + 2

# Save final file
wb.save(output_path)
print(f"✅ Matching completed. New sheet 'CM_Invoice_Match' added to {output_path}")
