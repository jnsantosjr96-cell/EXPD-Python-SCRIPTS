
import os
import fitz  # PyMuPDF

# Define source and destination folders
source_folder = r"C:\Users\josenjr\OneDrive - Expedia Group\Desktop\Daily activities\2026\March\03.04.2026_2\Source Folder"
destination_folder = os.path.join(source_folder, "Compressed")

# Create destination folder if it doesn't exist
os.makedirs(destination_folder, exist_ok=True)

# Reference range and mapping from Boleto index to reference number
# Consider one DM number below the first in the range
# and one more DM number above the last number of the sequence
reference_range = list(range(26300085, 26300139))

# Loop through each reference number and merge corresponding PDFs - Update BOLETO INDEX NUMBER AND MONTH IN FILE NAME
for ref in reference_range:
    broubrl_filename = f"2026_03_BROUBRL_{ref}.pdf"
    boleto_index = ref - 26300085
    boleto_filename = f"Boleto-{boleto_index}.pdf"

    broubrl_path = os.path.join(source_folder, broubrl_filename)
    boleto_path = os.path.join(source_folder, boleto_filename)

    # Check if both files exist
    if os.path.exists(broubrl_path) and os.path.exists(boleto_path):
        merged_pdf = fitz.open()

        # Merge BROUBRL PDF
        with fitz.open(broubrl_path) as pdf1:
            merged_pdf.insert_pdf(pdf1)

        # Merge Boleto PDF
        with fitz.open(boleto_path) as pdf2:
            merged_pdf.insert_pdf(pdf2)

        # Save merged PDF with updated name format
        output_filename = f"Invoice_00{ref}.pdf"
        output_path = os.path.join(destination_folder, output_filename)
        merged_pdf.save(output_path)
        merged_pdf.close()

print("PDF merging completed. Merged files saved in 'Compressed' folder with updated names.")
