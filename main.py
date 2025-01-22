import pandas as pd
from docx import Document

from utility import replace_title, may_empty

# Load the responses from the Excel file
input_file = "./data/responses.xlsx"
responses = pd.read_excel(input_file)

# Define the template file (can be an empty form with placeholders)
template_file = "./forms/formTest.docx"

placeholder = ''

# Process each row and generate a form
for index, row in responses.iterrows():
    # Load the template
    doc = Document(template_file)

    
    # Replace placeholders with actual data
    for paragraph in doc.paragraphs:
        print('the paragraph is: ', paragraph.text)
        for col_index, (key, value) in enumerate(row.items()):  # Use enumerate
            print(f"Processing column {col_index}: key = {key}, value = {value}")

            # if pd.isna(value):  # Check if the value is missing
            #     continue  # Skip this placeholder if the value is missing
            # print('key is: ', key)
            # print('value is: ', value)
            # placeholder = f"{{{{{key}}}}}"  # Placeholder format: {{ColumnName}}
            # print('the placeholder is: ', placeholder)
            if key == 'เพศ':    
                paragraph = replace_title(paragraph, str(value)) 
                
            elif key == 'ชื่อ':  # Correct syntax for equality check
                placeholder = '2'
                paragraph.text = paragraph.text.replace(placeholder, str(value))

            elif key == 'นามสกุล':  # Correct syntax for equality check
                placeholder = '3'
                paragraph.text = paragraph.text.replace(placeholder, str(value))

            elif key == 'วัน/เดือน/ปี (พ.ศ.) เกิด':  # Correct syntax for equality check
                placeholder = '4'
                paragraph.text = paragraph.text.replace(placeholder, str(value))
            
            elif key == 'ที่อยู่ปัจจุบัน เลขที่':  # Correct syntax for equality check
                placeholder = '5'
                paragraph.text = paragraph.text.replace(placeholder, str(value))

            elif key == 'หมู่ที่':  # Correct syntax for equality check
                placeholder = '6'
                paragraph = may_empty(paragraph, value, placeholder)

            elif key == 'หมู่บ้าน':  # Correct syntax for equality check
                placeholder = '7'
                paragraph = may_empty(paragraph, value, placeholder)

            elif key == 'ซอย':  # Correct syntax for equality check
                placeholder = '8'
                paragraph = may_empty(paragraph, value, placeholder)

            elif key == 'ถนน':  # Correct syntax for equality check
                placeholder = '9'
                paragraph = may_empty(paragraph, value, placeholder)

            elif key == 'แขวง':  # Correct syntax for equality check
                placeholder = '10'
                paragraph.text = paragraph.text.replace(placeholder, str(value))

            else:
                print(f"In else {col_index}: key = {key}, value = {value}")
    
    # Save the filled form
    output_directory = "./results/"
    output_file = f"{output_directory}form_{index + 1}.docx"
    # output_file = f"form_{index + 1}.docx"  # Name the file based on the row index
    doc.save(output_file)

print("Forms generated successfully!")


