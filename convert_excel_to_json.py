import pandas as pd
import json
from datetime import datetime

# Read the Excel file
excel_file = r'c:\Users\pavan\Downloads\Y-23  CRT  SUMMER TRAINING  Attendance (23.05.2026) (1).xlsx'

try:
    # Read the Excel file
    df = pd.read_excel(excel_file, engine='openpyxl')
    
    # Skip the header rows and find the actual data
    # Row 1 contains the title, Row 2 contains the column headers
    df.columns = df.iloc[1]
    df = df[2:].reset_index(drop=True)
    
    # Get the column names
    print("Column names:", df.columns.tolist())
    
    # Rename the first column to S.NO if it's not already
    df.columns.values[0] = 'S.NO'
    
    # Print first few rows to understand the data
    print("\nFirst few rows:")
    print(df.head(10))
    
    # Get the time slot columns (the numeric columns at the end)
    time_slots = []
    for col in df.columns:
        if str(col).isdigit() or ('-' in str(col) and ':' in str(col)):
            time_slots.append(col)
    
    print(f"\nTime slots found: {time_slots}")
    
    # Extract student data and attendance
    students_data = []
    
    for index, row in df.iterrows():
        try:
            # Get student information
            s_no = row.get('S.NO', '')
            name = row.get('NAME', '')
            branch = row.get('BRANCH', '')
            dept = row.get('DEPT', '')
            cluster = row.get('CLUSTER', '')
            crt_sec = row.get('CRT SEC', '')
            regd_no = row.get('REGD.NO', '')
            
            # Skip rows without essential data
            if pd.isna(name) or pd.isna(regd_no):
                continue
            
            # Use REGD.NO as the student ID (this is the actual registration number)
            student_id = str(int(regd_no)) if pd.notna(regd_no) else ''
            
            # Ensure ID starts with 23 (it should from the Excel data)
            if not student_id.startswith('23'):
                print(f"Warning: Student ID {student_id} doesn't start with 23, skipping")
                continue
            
            # Extract attendance for each time slot
            attendance_records = []
            for i, time_slot in enumerate(time_slots):
                status = row.get(time_slot, '')
                if pd.notna(status):
                    # Convert P/A to present/absent
                    status_lower = str(status).upper()
                    if status_lower == 'P':
                        attendance_status = 'present'
                    elif status_lower == 'A':
                        attendance_status = 'absent'
                    else:
                        attendance_status = 'absent'  # Default to absent if unclear
                    
                    # Create a date (using sequential dates starting from a base date)
                    # Since we don't have actual dates, we'll use the time slot index
                    base_date = datetime(2024, 5, 11)  # Starting from May 11, 2024 as mentioned in Excel
                    attendance_date = (base_date + pd.Timedelta(days=i)).strftime('%Y-%m-%d')
                    
                    attendance_records.append({
                        'date': attendance_date,
                        'status': attendance_status
                    })
            
            # Only add students who have attendance records
            if attendance_records:
                students_data.append({
                    'id': student_id,
                    'name': str(name),
                    'attendance': attendance_records
                })
                
        except Exception as e:
            print(f"Error processing row {index}: {e}")
            continue
    
    # Create the final JSON structure
    attendance_json = {
        'students': students_data
    }
    
    # Save to JSON file
    output_file = r'c:\Users\pavan\OneDrive\pavan\crt attendence webiste\attendance.json'
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(attendance_json, f, indent=2, ensure_ascii=False)
    
    print(f"\nSuccessfully converted {len(students_data)} students to JSON format")
    print(f"Output saved to: {output_file}")
    
    # Print sample of converted data
    print("\nSample of converted data (first 3 students):")
    for student in students_data[:3]:
        print(f"ID: {student['id']}, Name: {student['name']}")
        print(f"Attendance records: {len(student['attendance'])}")
        print(f"First 3 records: {student['attendance'][:3]}")
        print()
    
except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()
