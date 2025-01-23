import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
import os
from reportlab.pdfbase import pdfmetrics

image_path = './data/images.png'

# Load the Excel file
excel_file = "./data/responses.xlsx"  # Replace with your file name
data = pd.read_excel(excel_file)

# Register the Thai font if needed
pdfmetrics.registerFont(TTFont('THSarabun', './THSarabun/THSarabun-New-Regular.ttf'))

# Function to read template from a file
def load_template(template_file):
    with open(template_file, "r", encoding="utf-8") as file:
        return file.read()

# Load the template
template_file = "./forms/formTxt.txt"  # Replace with the path to your template file
template = load_template(template_file)

# Function to replace placeholders
def generate_form(row, template):
    form = template
    for col in row.index:
        print(col)
        placeholder = ''
        if col.strip() == 'เพศ':
            if row[col].strip() == 'นางสาว':
                placeholder = '$02'
                print(f"Column: {col}, Value: {row[col]}")
                print('grrrrr in pet')
                # form = form.replace(placeholder, str(row[col]) if pd.notna(row[col]) else "")
                form = form.replace(placeholder, 'X' if pd.notna(row[col]) else "")
        elif col.strip() == 'กรุณาตรวจสอบคำตอบของคุณก่อนที่จะส่ง':
            print('.............')
            # add a pic


    # print("Generated form:", form)
    return form

def create_pdf(content, output_path, image_path=None):
    c = canvas.Canvas(output_path, pagesize=letter)
    width, height = letter
    c.setFont("THSarabun", 18)  # Use the Thai font here
    lines = content.split("\n")

    y_position = height - 40  # Starting position from the top
    
    for line in lines:
        c.drawString(40, y_position, line)
        y_position -= 14  # Move down for the next line
        
        if 'ธรรมศึกษาชั้นตรี' in line:
            picture_x = 500  # X position of image (adjust as needed)
            picture_y = y_position - 60  # Y position (adjust as needed)
            c.drawImage(image_path, picture_x, picture_y, width=100, height=100)  # Adjust size and position
            # y_position -= 120  # Adjust space for the image
            y_position = picture_y + 30  # Adjust y_position to account for the height of the image

        if 'ธรรมศึกษาชั้นเอก' in line:
            picture_x = 500  # X position of image (adjust as needed)
            picture_y = y_position - 60  # Y position (adjust as needed)
            c.drawImage(image_path, picture_x, picture_y, width=100, height=100)  # Adjust size and position
            # y_position -= 120  # Adjust space for the image
            y_position = picture_y + 60  # Adjust y_position to account for the height of the image



        
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

    # Assuming the column for image paths is 'image_path' in the Excel file
    # image_path = row.get('./data/images.png', None)  # Modify this based on your Excel column name
    
    output_path = os.path.join(output_folder, f"form_{index + 1}.pdf")
    create_pdf(filled_form, output_path, image_path)

print(f"PDFs have been generated in the folder: {output_folder}")

