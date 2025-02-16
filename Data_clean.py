import pandas as pd
import re
import random
import os

def generate_student_id():
    return str(random.randint(10000, 9999999))

def grade(grade):
    if pd.isna(grade):
        return ''
    grade = str(grade).lower().strip()
    if grade == 'kinder':
        return 'K'
    # Handle numbers with suffixes (1st, 2nd, 3rd, 4th, etc.)
    match = re.match(r'^(\d+)(st|nd|rd|th)?$', grade)
    if match:
        return int(match.group(1))
    # If it's not a special case, return uppercase version
    return grade.upper()

def clean_language(lang):
    if pd.isna(lang):
        return 'English'
    lang = lang.strip().upper()
    lang_map = {
        'ENG': 'English',
        'SPAN': 'Spanish',
        'SPN': 'Spanish',
        'SPANISH': 'Spanish',
        'ARA': 'Arabic',
        'FRENCH': 'French',
        'THAI': 'Thai',
        'RUSSIAN': 'Russian'
    }
    return lang_map.get(lang, 'English')

def clean_phone(phone):
    if pd.isna(phone):
        return ''
    digits = re.sub(r'\D', '', str(phone))
    if len(digits) == 10:
        return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
    return ''

def add_contact_record(records, student_id, row, contact_type, phone_col, first_name_col, last_name_col, email_col, guardianship_col):
    phone = clean_phone(row.get(phone_col, ''))
    email = row.get(email_col, '')
    if isinstance(email, str):
        email = email.strip()
    else:
        email = ''
    if phone and email:
        records.append({
            'student_id': student_id,
            'student_first_name': row.get('Student First Name', ''),
            'student_last_name': row.get('Student Last Name', ''),
            'student_grade': grade(row.get('Student Grade', '')),
            'contact_first_name': row.get(first_name_col, ''),
            'contact_last_name': row.get(last_name_col, ''),
            'guardianship': row.get(guardianship_col, '').title() if pd.notna(row.get(guardianship_col, '')) else '',
            'contact_email': email,
            'contact_phone': phone,
            'language': clean_language(row.get('Language', '')),
            'contact_type': contact_type  
        })


def process_data(df):
    records = []
    student_ids = {}
    for _, row in df.iterrows():
        student_key = (row.get('Student First Name', ''), row.get('Student Last Name', ''))
        if student_key not in student_ids:
            student_ids[student_key] = generate_student_id()
        student_id = student_ids[student_key]
        add_contact_record(records, student_id, row, 'Parent #1', 'Parent #1 Phone number',
                           'Parent #1 First Name', 'Parent #1 Last Name',
                           'Parent #1 Parent Email Address', 'Guardianship')
        
        add_contact_record(records, student_id, row, 'Parent #2', 'Parent #2 Phone Number',
                           'Parent #2 First Name', 'Parent #2 Last Name',
                           'Parent #2 Email Address', 'Parent #2 Guardianship')
    
    return pd.DataFrame(records)


def main():
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    filename = os.path.join(downloads_folder, 'Data Cleaning Project.xlsx')
    
    # Read from 'Data Cleaning Project' sheet
    data = pd.read_excel(filename, sheet_name='School Data')
    
    # Process the data
    cleaned_data = process_data(data)
    
    # Define column order
    column_order = ['student_id', 'student_first_name', 'student_last_name', 'student_grade',
                    'contact_first_name', 'contact_last_name', 'guardianship', 'contact_email',
                    'contact_phone', 'language']
    
    # Reorder columns
    cleaned_data = cleaned_data[column_order]
    
    # Write to a new worksheet 'Desired Import Format'
    with pd.ExcelWriter(filename, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        cleaned_data.to_excel(writer, sheet_name='Processed Data', index=False)
    
    print("Data processing complete. The processed data has been written to 'Desired Import Format' sheet.")
if __name__ == "__main__":
    main()