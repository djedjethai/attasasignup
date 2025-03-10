import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import config

def create_pdf(content, output_path, image_path=None):
    content = content.replace('$ti', '$X' if config.TRI else '$Y')
    content = content.replace('$to', '$X' if config.TO else '$Y')
    content = content.replace('$eg', '$X' if config.EG else '$Y')

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

def fil_form_to_eg(row, form):
    placeholder = ''
    for col in row.index:
        # if col.strip() == 'เพศ':
        if col.strip() == 'คำนำหน้าฃื่อ':
                form = replace_title(form, row[col])
        elif col.strip() == 'ชื่อ':
                placeholder = '$2'
                form = form.replace(placeholder, str(row[col]))
        elif col.strip() == 'นามสกุล':
                placeholder = '$3'
                strVal = str(row[col])
                config.FILENAME = strVal
                form = form.replace(placeholder, strVal)
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
        # elif col.strip() == 'รุ่นที่': # krou samathi
        elif col.strip() == 'รุ่น': # krou samathi
                placeholder = '$19'
                form = may_empty(form, row[col], placeholder)
        elif col.strip() == 'เลขสาขา':
                placeholder = '$-2'
                form = may_empty(form, row[col], placeholder)
        elif col.strip() == 'เบอร์โทรศัพท์ (รูปแบบ xxx-xxx-xxxx)':
                placeholder = '$14'
                form = form.replace(placeholder, str(row[col]))
        # elif col.strip() == 'กรณีฉุกเฉินติดต่อ':
        elif col.strip() == 'ชื่อผู้ติดต่อกรณีฉุกเฉิน':
                placeholder = '$17'
                form = form.replace(placeholder, str(row[col]))
        elif col.strip() == 'เบอร์โทรศัพท์ผู้ติดต่อฉุกเฉิน (รูปแบบ xxx-xxx-xxxx)':
                placeholder = '$18'
                form = form.replace(placeholder, str(row[col]))
        
    return form

def fil_form_tri(row, form):
    placeholder = ''
    for col in row.index:
        # if col.strip() == 'เพศ':
        if col.strip() == 'คำนำหน้าฃื่อ':
                form = replace_title(form, row[col])
        elif col.strip() == 'ชื่อ':
                placeholder = '$2'
                form = form.replace(placeholder, str(row[col]))
        elif col.strip() == 'นามสกุล':
                placeholder = '$3'
                strVal = str(row[col])
                config.FILENAME = strVal
                form = form.replace(placeholder, strVal)
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
        # TODO to change next round, to align with to and eg
        elif col.strip() == 'รุ่นที่': # krou samathi
                placeholder = '$19'
                form = may_empty(form, row[col], placeholder)
        elif col.strip() == 'เลขสาขา':
                placeholder = '$-2'
                form = may_empty(form, row[col], placeholder)
        elif col.strip() == 'เบอร์โทรศัพท์ (รูปแบบ xxx-xxx-xxxx)':
                placeholder = '$14'
                form = form.replace(placeholder, str(row[col]))
        # TODO to change next round, to align with to and eg
        elif col.strip() == 'กรณีฉุกเฉินติดต่อ':
                placeholder = '$17'
                form = form.replace(placeholder, str(row[col]))
        elif col.strip() == 'เบอร์โทรศัพท์ผู้ติดต่อฉุกเฉิน (รูปแบบ xxx-xxx-xxxx)':
                placeholder = '$18'
                form = form.replace(placeholder, str(row[col]))
        
    return form

def may_empty(form, value, placeholder):
    if pd.isna(value):
        form = form.replace(placeholder, '.........')
    else:
        valueStr = str(value)
        valueStr = valueStr.split('.')[0]
        form = form.replace(placeholder, valueStr)
    return form

def replace_title(form, gender_value):
    # Define the placeholders to be replaced
    placeholders = ['$00', '$01', '$02', '$03']
     
    # Replace the correct placeholder with 'X' based on gender value
    if gender_value == 'นาย':
        form = form.replace('$00', '$X' if pd.notna(gender_value) else "")
        placeholders.remove('$00')
        form = form.replace('$04', "")
    elif gender_value == 'นาง':
        form = form.replace('$01', '$X' if pd.notna(gender_value) else "")
        placeholders.remove('$01')
        form = form.replace('$04', "")
    elif gender_value == 'นางสาว':
        form = form.replace('$02', '$X' if pd.notna(gender_value) else "")
        placeholders.remove('$02')
        form = form.replace('$04', "")
    else:
        form = form.replace('$03', '$X' if pd.notna(gender_value) else "")
        placeholders.remove('$03')
        form = form.replace('$04', f"{gender_value}" if pd.notna(gender_value) else "")

    return clear_placeholders(form, placeholders)

def clear_placeholders(form, placeholders):
    # Remove placeholders that were not replaced yet
    for placeholder in placeholders:
        form = form.replace(placeholder, '$Y')
    return form
