import pandas as pd
import json
from datetime import datetime
from collections import defaultdict

# Excel files to merge
excel_files = [
    r'c:\Users\pavan\Downloads\Y23 CRT Attendance (11.05.26).xlsx',
    r'c:\Users\pavan\Downloads\Y-23  SUMMER CRT Attendance (12.05.26).xlsx',
    r'c:\Users\pavan\Downloads\Y-23  CRT  SUMMER TRAINING  Attendance Report As oN 13.05.2026.xlsx',
    r'c:\Users\pavan\Downloads\Y-23  CRT  SUMMER TRAINING  Attendance Report As oN 14.05.2026.xlsx',
    r'c:\Users\pavan\Downloads\Y-23  SUMMER CRT Attendance(15.05.2026).xlsx',
    r'c:\Users\pavan\Downloads\Y-23  CRT  SUMMER TRAINING  Attendance(16.05.2026).xlsx',
    r'c:\Users\pavan\Downloads\Y-23  CRT  SUMMER TRAINING  Attemndance(18.05.26).xlsx',
    r'c:\Users\pavan\Downloads\Y-23  CRT  SUMMER TRAINING  Attendance (19.05.26).xlsx',
    r'c:\Users\pavan\Downloads\Y-23  CRT  SUMMER TRAINING  Attendance(20.05.2026) (2).xlsx',
    r'c:\Users\pavan\Downloads\Y-23  CRT  SUMMER TRAINING  Attendance(21.05.2026).xlsx',
    r'c:\Users\pavan\Downloads\Y-23  CRT  SUMMER TRAINING  Attendance (22.05.2026).xlsx',
    r'c:\Users\pavan\Downloads\Y-23  CRT  SUMMER TRAINING  Attendance (23.05.2026).xlsx'
]

# Dictionary to store all student data
all_students = defaultdict(lambda: {
    'name': '',
    'branch': '',
    'department': '',
    'cluster': '',
    'crt_sec': '',
    'crt_room': '',
    'attendance': []
})

# Process each Excel file
for excel_file in excel_files:
    try:
        print(f"\nProcessing: {excel_file}")
        
        # Read the Excel file
        df = pd.read_excel(excel_file, engine='openpyxl')
        
        # Skip the header rows and find the actual data
        df.columns = df.iloc[1]
        df = df[2:].reset_index(drop=True)
        
        # Rename the first column to S.NO if it's not already
        df.columns.values[0] = 'S.NO'
        
        # Get the time slot columns (the numeric columns at the end)
        time_slots = []
        for col in df.columns:
            if str(col).isdigit() or ('-' in str(col) and ':' in str(col)):
                time_slots.append(col)
        
        print(f"Time slots found: {time_slots}")
        
        # Determine base date from filename
        if '11.05.26' in excel_file or '11.05.2026' in excel_file:
            base_date = datetime(2026, 5, 11)
        elif '12.05.26' in excel_file or '12.05.2026' in excel_file:
            base_date = datetime(2026, 5, 12)
        elif '13.05.2026' in excel_file:
            base_date = datetime(2026, 5, 13)
        elif '14.05.2026' in excel_file:
            base_date = datetime(2026, 5, 14)
        elif '15.05.2026' in excel_file:
            base_date = datetime(2026, 5, 15)
        elif '16.05.2026' in excel_file:
            base_date = datetime(2026, 5, 16)
        elif '18.05.26' in excel_file or '18.05.2026' in excel_file:
            base_date = datetime(2026, 5, 18)
        elif '19.05.26' in excel_file or '19.05.2026' in excel_file:
            base_date = datetime(2026, 5, 19)
        elif '20.05.2026' in excel_file:
            base_date = datetime(2026, 5, 20)
        elif '21.05.2026' in excel_file:
            base_date = datetime(2026, 5, 21)
        elif '22.05.2026' in excel_file:
            base_date = datetime(2026, 5, 22)
        elif '23.05.2026' in excel_file:
            base_date = datetime(2026, 5, 23)
        else:
            base_date = datetime(2024, 5, 11)
        
        print(f"Base date: {base_date.strftime('%Y-%m-%d')}")
        
        # Extract student data and attendance
        for index, row in df.iterrows():
            try:
                # Get student information
                name = row.get('NAME', '')
                regd_no = row.get('REGD.NO', '')
                branch = row.get('BRANCH', '')
                dept = row.get('DEPT', '')
                cluster = row.get('CLUSTER', '')
                crt_sec = row.get('CRT SEC', '')
                crt_room = row.get('CRT ROOM', '')
                
                # Skip rows without essential data
                if pd.isna(name) or pd.isna(regd_no):
                    continue
                
                # Use REGD.NO as the student ID
                student_id = str(int(regd_no)) if pd.notna(regd_no) else ''
                
                # Ensure ID starts with 23
                if not student_id.startswith('23'):
                    continue
                
                # Store/update student details (only if not already set)
                if not all_students[student_id]['name']:
                    all_students[student_id]['name'] = str(name)
                if not all_students[student_id]['branch'] and pd.notna(branch):
                    all_students[student_id]['branch'] = str(branch)
                if not all_students[student_id]['department'] and pd.notna(dept):
                    all_students[student_id]['department'] = str(dept)
                if not all_students[student_id]['cluster'] and pd.notna(cluster):
                    all_students[student_id]['cluster'] = str(cluster)
                if not all_students[student_id]['crt_sec'] and pd.notna(crt_sec):
                    all_students[student_id]['crt_sec'] = str(crt_sec)
                if not all_students[student_id]['crt_room'] and pd.notna(crt_room):
                    all_students[student_id]['crt_room'] = str(crt_room)
                
                # Extract attendance for each time slot - all time slots belong to the same date
                for time_slot in time_slots:
                    status = row.get(time_slot, '')
                    if pd.notna(status):
                        # Convert P/A to present/absent
                        status_lower = str(status).upper()
                        if status_lower == 'P':
                            attendance_status = 'present'
                        elif status_lower == 'A':
                            attendance_status = 'absent'
                        else:
                            attendance_status = 'absent'
                        
                        # Use the base date for all time slots (not sequential dates)
                        attendance_date = base_date.strftime('%Y-%m-%d')
                        
                        # Add attendance record with time slot information
                        all_students[student_id]['attendance'].append({
                            'date': attendance_date,
                            'time_slot': str(time_slot),
                            'status': attendance_status
                        })
                        
            except Exception as e:
                print(f"Error processing row {index}: {e}")
                continue
        
        print(f"Processed students from this file")
        
    except Exception as e:
        print(f"Error processing {excel_file}: {e}")
        import traceback
        traceback.print_exc()

# Convert to the required JSON structure
students_data = []
for student_id, data in all_students.items():
    if data['attendance']:  # Only include students with attendance records
        # Sort attendance by date
        data['attendance'].sort(key=lambda x: x['date'])
        
        students_data.append({
            'id': student_id,
            'name': data['name'],
            'branch': data['branch'],
            'department': data['department'],
            'cluster': data['cluster'],
            'crt_sec': data['crt_sec'],
            'crt_room': data['crt_room'],
            'attendance': data['attendance']
        })

# Create the final JSON structure
attendance_json = {
    'students': students_data
}

# Save to JSON file
output_file = r'c:\Users\pavan\OneDrive\pavan\crt attendence webiste\attendance.json'
with open(output_file, 'w', encoding='utf-8') as f:
    json.dump(attendance_json, f, indent=2, ensure_ascii=False)

print(f"\nSuccessfully merged data from {len(excel_files)} Excel files")
print(f"Total students: {len(students_data)}")
print(f"Output saved to: {output_file}")

# Print sample of merged data
print("\nSample of merged data (first 3 students):")
for student in students_data[:3]:
    print(f"ID: {student['id']}, Name: {student['name']}")
    print(f"Total attendance records: {len(student['attendance'])}")
    print(f"Date range: {student['attendance'][0]['date']} to {student['attendance'][-1]['date']}")
    print()
