import pandas as pd

def fil_form(row, form):
    placeholder = ''
    for col in row.index:
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
    elif gender_value == 'นาง':
        form = form.replace('$01', '$X' if pd.notna(gender_value) else "")
        placeholders.remove('$01')
    elif gender_value == 'นางสาว':
        form = form.replace('$02', '$X' if pd.notna(gender_value) else "")
        placeholders.remove('$02')
    else:
        form = form.replace('$03', f"{gender_value}" if pd.notna(gender_value) else "")
        placeholders.remove('$03')

    return clear_placeholders(form, placeholders)

def clear_placeholders(form, placeholders):
    # Remove placeholders that were not replaced yet
    for placeholder in placeholders:
        form = form.replace(placeholder, '$Y')
    return form
