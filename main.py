import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
import os
from reportlab.pdfbase import pdfmetrics

from utility import replace_title, may_empty

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

placeholder = ''

# Function to replace placeholders
def generate_form(row, template):
    form = template
    for col in row.index:
        print(col)
        if col.strip() == 'เพศ':
            form = replace_title(form, row[col])
        elif col.strip() == 'ชื่อ':
                placeholder = '$2'
                form = form.replace(placeholder, str(row[col]))
        elif col.strip() == 'นามสกุล':
                placeholder = '$3'
                form = form.replace(placeholder, str(row[col]))
        elif col.strip() == 'วัน/เดือน/ปี (พ.ศ.) เกิด':
                placeholder = '$4'
                form = form.replace(placeholder, str(row[col]))
        elif col.strip() == 'ที่อยู่ปัจจุบัน เลขที่':
                placeholder = '$5'
                form = form.replace(placeholder, str(row[col]))
        elif col.strip() == 'หมู่ที่':
                placeholder = '$6'
                form = may_empty(form, row[col], placeholder)
        elif col.strip() == 'หมู่บ้าน':
                placeholder = '$7'
                form = may_empty(form, row[col], placeholder)
        elif col.strip() == 'ซอย':
                placeholder = '$8'
                form = may_empty(form, row[col], placeholder)
        elif col.strip() == 'ถนน':
                placeholder = '$9'
                form = may_empty(form, row[col], placeholder)
        elif col.strip() == 'แขวง':
                placeholder = '$10'
                form = form.replace(placeholder, str(row[col]))
        elif col.strip() == 'เขต':
                placeholder = '$11'
                form = form.replace(placeholder, str(row[col]))
        elif col.strip() == 'จังหวัด':
                placeholder = '$12'
                form = form.replace(placeholder, str(row[col]))
        elif col.strip() == 'รหัสไปรษณีย์':
                placeholder = '$13'
                form = form.replace(placeholder, str(row[col]))
        elif col.strip() == 'รุ่นที่': # krou samathi
                placeholder = '$19'
                form = may_empty(form, row[col], placeholder)
        elif col.strip() == 'เลขสาขา':
                placeholder = '$-2'
                form = may_empty(form, row[col], placeholder)
        elif col.strip() == 'เบอร์โทรศัพท์ (รูปแบบ xxx-xxx-xxxx)':
                placeholder = '$14'
                form = form.replace(placeholder, str(row[col]))
        elif col.strip() == 'กรณีฉุกเฉินติดต่อ':
                placeholder = '$17'
                form = form.replace(placeholder, str(row[col]))
        elif col.strip() == 'เบอร์โทรศัพท์ผู้ติดต่อฉุกเฉิน (รูปแบบ xxx-xxx-xxxx)':
                placeholder = '$18'
                form = form.replace(placeholder, str(row[col]))
        elif col.strip() == 'กรุณาตรวจสอบคำตอบของคุณก่อนที่จะส่ง':
            print('.............')
            # add a pic

    # print("Generated form:", form)
    return form


def create_pdf(content, output_path, image_path=None):
    content = content.replace('$ti', 'X' if tri else '□')
    content = content.replace('$to', 'X' if to else '□')
    content = content.replace('$eg', 'X' if eg else '□')

    line_spacing = 2

    c = canvas.Canvas(output_path, pagesize=letter)
    width, height = letter
    c.setFont("THSarabun", 18)  # Use the Thai font here
    # c.setFont("DejaVuSans", 18) 
    lines = content.split("\n")

    y_position = height - 40  # Starting position from the top
    
    for line in lines:
        y_position -= line_spacing
        
        if 'X' in line:
            parts = line.split('X')
            x_position = 40
            for part in parts[:-1]:
                c.drawString(x_position, y_position, part)
                x_position += c.stringWidth(part, "THSarabun", 18)
                c.setFont("DejaVuSans", 18)
                c.drawString(x_position, y_position, '☒')
                x_position += c.stringWidth('☒', "DejaVuSans", 18)
                c.setFont("THSarabun", 18)
            c.drawString(x_position, y_position, parts[-1])
        else:
            c.drawString(40, y_position, line)
        
                
        if 'ธรรมศึกษาชั้นตรี' in line:
            picture_x = 480  # X position of image (adjust as needed)
            picture_y = y_position - 75  # Y position (adjust as needed)
            c.drawImage(image_path, picture_x, picture_y, width=100, height=100)  # Adjust size and position
            # y_position -= 120  # Adjust space for the image
            y_position = picture_y + 75  # Adjust y_position to account for the height of the image

        if 'ธรรมศึกษาชั้นเอก' in line:
            picture_x = 480  # X position of image (adjust as needed)
            picture_y = y_position - 75  # Y position (adjust as needed)
            c.drawImage(image_path, picture_x, picture_y, width=100, height=100)  # Adjust size and position
            # y_position -= 120  # Adjust space for the image
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

    # Assuming the column for image paths is 'image_path' in the Excel file
    # image_path = row.get('./data/images.png', None)  # Modify this based on your Excel column name
    
    output_path = os.path.join(output_folder, f"form_{index + 1}.pdf")
    create_pdf(filled_form, output_path, image_path)

print(f"PDFs have been generated in the folder: {output_folder}")

