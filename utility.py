import pandas as pd

def may_empty(paragraph, value, placeholder):
    if pd.isna(value):
        paragraph.text = paragraph.text.replace(placeholder, '...............')
    else:
        paragraph.text = paragraph.text.replace(placeholder, str(value))

    return paragraph

def replace_title(paragraph, gender_value):
    # Define the placeholders to be replaced
    placeholders = ['$00', '$01', '$02', '$03']
     
    # Replace the correct placeholder with 'X' based on gender value
    if gender_value == 'นาย':
        paragraph.text = paragraph.text.replace('$00', '☒')
        placeholders.remove('$00')
    elif gender_value == 'นาง':
        paragraph.text = paragraph.text.replace('$01', '☒')
        placeholders.remove('$01')
    elif gender_value == 'นางสาว':
        paragraph.text = paragraph.text.replace('$02', '☒')
        placeholders.remove('$02')
    else:
        paragraph.text = paragraph.text.replace('$03', f"{gender_value}")
        placeholders.remove('$03')

    return clear_placeholders(paragraph, placeholders)

def clear_placeholders(paragraph, placeholders):
    # Remove placeholders that were not replaced yet
    for placeholder in placeholders:
        paragraph.text = paragraph.text.replace(placeholder, '□')  # Replace the placeholder with a space
    return paragraph
