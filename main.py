import pandas as pd
import os
import config

from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

from utility import replace_title, may_empty, fil_form_tri, fil_form_to_eg, create_pdf

# Load the Excel file
data = pd.read_excel(config.EXCEL_FILE)

# Register the Thai font if needed
pdfmetrics.registerFont(TTFont('THSarabun', config.TH_SARABUN))
pdfmetrics.registerFont(TTFont('DejaVuSans', config.DEJAVU_SANS))

# Function to read template from a file
def load_template(template_file):
    with open(template_file, "r", encoding="utf-8") as file:
        return file.read()

# Load the template
template = load_template(config.TEMPLATE_FILE)

# Generate PDFs for each row
os.makedirs(config.OUTPUT_FOLDER, exist_ok=True)

def generate_form(row, template):
    if config.TO or config.EG:
        return fil_form_to_eg(row, template)
    elif config.TRI:
        return fil_form_tri(row, template)
    else:
        print('Error, invalid level.......')
        
# Run the program
for index, row in data.iterrows():
    # filled_form = fil_form(row, template)
    filled_form = generate_form(row, template)

    output_path = os.path.join(config.OUTPUT_FOLDER, f"{config.FILENAME}_{index + 1}.pdf")
    create_pdf(filled_form, output_path, config.IMAGE_PATH)

print(f"PDFs have been generated in the folder: {config.OUTPUT_FOLDER}")

