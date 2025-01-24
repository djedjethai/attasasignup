import pandas as pd

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
        form = form.replace('$00', 'X' if pd.notna(gender_value) else "")
        placeholders.remove('$00')
    elif gender_value == 'นาง':
        form = form.replace('$01', 'X' if pd.notna(gender_value) else "")
        placeholders.remove('$01')
    elif gender_value == 'นางสาว':
        form = form.replace('$02', 'X' if pd.notna(gender_value) else "")
        placeholders.remove('$02')
    else:
        form = form.replace('$03', f"{gender_value}" if pd.notna(gender_value) else "")
        placeholders.remove('$03')

    return clear_placeholders(form, placeholders)

def clear_placeholders(form, placeholders):
    # Remove placeholders that were not replaced yet
    for placeholder in placeholders:
        form = form.replace(placeholder, '□')
    return form
