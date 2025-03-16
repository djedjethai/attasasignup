import sys
import argparse
import pandas as pd
import os
import config

from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

from utility import replace_title, may_empty, fil_form_tri, fil_form_to_eg, create_pdf

# Set up argument parsing
parser = argparse.ArgumentParser(description="Process an Excel file and generate a PDF report.")
parser.add_argument("args", nargs="*", help="Optional arguments for processing")

cli_args = parser.parse_args()

# Allowed first arguments
valid_options = {"tri", "to", "eg"}

if not cli_args.args or cli_args.args[0] not in valid_options:
    print("Error: First argument must be one of 'tri', 'to', or 'eg'.")
    sys.exit(1)  # Exit the program with an error code

if cli_args.args[0] == 'tri':
    config.TRI = True
if cli_args.args[0] == 'to':
    config.TO = True
if cli_args.args[0] == 'eg':
    config.EG = True

# Debug: Print received arguments
print(f"Students level: {cli_args.args[0]}")

# Get the list of files in the data folder
files_in_data_folder = [f for f in os.listdir(config.DATA_FOLDER) if os.path.isfile(os.path.join(config.DATA_FOLDER, f))]

# Check if there is more than one file in the 'data' folder
if len(files_in_data_folder) != 1:
    print(f"Error: More than one or no file found in the '{config.DATA_FOLDER}' folder")
    sys.exit(1)

excel_file = os.path.join(config.DATA_FOLDER, files_in_data_folder[0])
data = pd.read_excel(excel_file)
print(f"Loaded file: {excel_file}")

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
    formatted_index = str(index + 1).zfill(3)

    # output_path = os.path.join(config.OUTPUT_FOLDER, f"{config.FILENAME}_{index + 1}.pdf")
    output_path = os.path.join(config.OUTPUT_FOLDER, f"{formatted_index}_{config.FILENAME}.pdf")
    create_pdf(filled_form, output_path, config.IMAGE_PATH)

print(f"PDFs have been generated in the folder: {config.OUTPUT_FOLDER}")

