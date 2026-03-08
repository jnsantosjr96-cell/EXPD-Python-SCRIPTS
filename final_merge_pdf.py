
"""
Simple Merge: Invoice (2 pages) + Boleto (1 page)
No OCR, no text extraction, no regex.
Documents are matched STRICTLY by the order of pages.

You only provide the STARTING NUMBER and the script generates:

invoice_00{num}.pdf
invoice_00{num+1}.pdf
invoice_00{num+2}.pdf
...
"""

from PyPDF2 import PdfReader, PdfWriter
from pathlib import Path

# ================================
# CONFIGURATION
# ================================

INVOICES_PDF = r"C:\Users\josenjr\OneDrive - Expedia Group\Desktop\Daily activities\2026\January\01.20.2026\Invoices.pdf"
BOLETOS_PDF  = r"C:\Users\josenjr\OneDrive - Expedia Group\Desktop\Daily activities\2026\January\01.20.2026\Boletos.pdf"

OUT_DIR = Path(r"C:\Users\josenjr\OneDrive - Expedia Group\Desktop\Daily activities\2026\January\01.20.2026\Outcome Folder")
OUT_DIR.mkdir(exist_ok=True)

# How many pages each document contains
INVOICE_PAGES = 2
BOLETO_PAGES  = 1

# ⚠️⚠️ PLACE THE FIRST NUMBER HERE ⚠️⚠️
STARTING_NUMBER = 25300560


# ================================
# PROCESSING
# ================================

invoice_reader = PdfReader(INVOICES_PDF)
boleto_reader  = PdfReader(BOLETOS_PDF)

total_invoice_docs = len(invoice_reader.pages) // INVOICE_PAGES
total_boleto_docs  = len(boleto_reader.pages) // BOLETO_PAGES

pairs = min(total_invoice_docs, total_boleto_docs)

print(f"Total pairs available: {pairs}\n")

current_number = STARTING_NUMBER

for i in range(pairs):
    writer = PdfWriter()

    # invoice pages
    inv_start = i * INVOICE_PAGES
    for p in range(inv_start, inv_start + INVOICE_PAGES):
        writer.add_page(invoice_reader.pages[p])

    # boleto page
    bol_start = i * BOLETO_PAGES
    writer.add_page(boleto_reader.pages[bol_start])

    # standardized name: invoice_00XXXXXXXX.pdf
    final_name = f"invoice_00{current_number}.pdf"
    output_path = OUT_DIR / final_name

    with open(output_path, "wb") as f:
        writer.write(f)

    print(f"Generated: {final_name}")

    current_number += 1  # increments sequence

print("\nProcess completed!")
print(f"Files saved at: {OUT_DIR.resolve()}")
