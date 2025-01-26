import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
import os
from reportlab.pdfbase import pdfmetrics

from utility import replace_title, may_empty, fil_form

image_path = './data/picture.png'

tri = False
to = True
eg = False

# Load the Excel file
excel_file = "./data/responses.xlsx"  # Replace with your file name
data = pd.read_excel(excel_file)

# Register the Thai font if needed
pdfmetrics.registerFont(TTFont('THSarabun', './fonts/THSarabun-New-Regular.ttf'))
pdfmetrics.registerFont(TTFont('DejaVuSans', './fonts/DejaVuSans.ttf'))

# Function to read template from a file
def load_template(template_file):
    with open(template_file, "r", encoding="utf-8") as file:
        return file.read()

# Load the template
template_file = "./forms/formTxt.txt"  # Replace with the path to your template file
template = load_template(template_file)

# Function to replace placeholders
def generate_form(row, template):
    
    # NOTE in case there is various form, otherwise no need to embed the func
    form = fil_form(row, template)
    return form


def create_pdf(content, output_path, image_path=None):
    content = content.replace('$ti', '$X' if tri else '$Y')
    content = content.replace('$to', '$X' if to else '$Y')
    content = content.replace('$eg', '$X' if eg else '$Y')

    line_spacing = 2

    c = canvas.Canvas(output_path, pagesize=letter)
    width, height = letter
    c.setFont("THSarabun", 18)  # Use the Thai font here
    # c.setFont("DejaVuSans", 18) 
    lines = content.split("\n")

    y_position = height - 40  # Starting position from the top
    
    for line in lines:
        y_position -= line_spacing
        x_position = 40


        # Check for placeholders in the line
        if '$X' in line or '$Y' in line:
            final_parts = []  # Store processed parts

            # Process the line for '$Y' first
            parts = line.split('$Y')
            for i, part in enumerate(parts):
                # Handle text before or after '$Y'
                if part:
                    final_parts.append((part, "THSarabun"))
                if i < len(parts) - 1:  # Add □ after every '$Y', except the last part
                    final_parts.append(('□', "DejaVuSans"))

            # Now process '$X' in the updated `final_parts`
            processed_parts = []
            for text, font in final_parts:
                if '$X' in text:
                    # Split on '$X' and add ☒ where needed
                    sub_parts = text.split('$X')
                    for j, sub_part in enumerate(sub_parts):
                        if sub_part:
                            processed_parts.append((sub_part, font))
                        if j < len(sub_parts) - 1:  # Add ☒ after every '$X', except the last part
                            processed_parts.append(('☒', "DejaVuSans"))
                else:
                    # No '$X', keep the part as is
                    processed_parts.append((text, font))

            # Draw the final processed parts
            for text, font in processed_parts:
                c.setFont(font, 18)
                if text == '☒':
                    # Adjust vertical position for ☒
                    c.drawString(x_position, y_position - 2.5, text)  # Shift ☒ slightly downward
                else:
                    c.drawString(x_position, y_position, text)
                x_position += c.stringWidth(text, font, 18)

        else:
            # Draw lines without '$X' or '$Y' directly
            c.drawString(x_position, y_position, line)

        if 'ธรรมศึกษาชั้นตรี' in line:
            picture_x = 480  # X position of image (adjust as needed)
            picture_y = y_position - 75  # Y position (adjust as needed)
            c.drawImage(image_path, picture_x, picture_y, width=100, height=100)  # Adjust size and position
            y_position = picture_y + 75  # Adjust y_position to account for the height of the image

        if 'ธรรมศึกษาชั้นเอก' in line:
            picture_x = 480  # X position of image (adjust as needed)
            picture_y = y_position - 75  # Y position (adjust as needed)
            c.drawImage(image_path, picture_x, picture_y, width=100, height=100)  # Adjust size and position
            y_position = picture_y + 75  # Adjust y_position to account for the height of the image

        y_position -= 14
        
        if y_position < 40:  # Check if we're at the bottom of the page
            c.showPage()
            c.setFont("THSarabun", 10)
            y_position = height - 40

    c.save()

# Generate PDFs for each row
output_folder = "output_pdfs"  # Folder to save the PDFs
os.makedirs(output_folder, exist_ok=True)

# Assuming there's a column in the Excel file with image paths
for index, row in data.iterrows():
    filled_form = generate_form(row, template)

    output_path = os.path.join(output_folder, f"form_{index + 1}.pdf")
    create_pdf(filled_form, output_path, image_path)

print(f"PDFs have been generated in the folder: {output_folder}")

