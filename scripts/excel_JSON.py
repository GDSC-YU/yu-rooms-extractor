import pandas as pd
import json

excel_file_path = 'data/TimeTable 20241.xlsx'
xls = pd.ExcelFile(excel_file_path)

rooms = {}

for sheet_name in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

    df.columns = ['Section Seq.', 'Course Code', 'Course Name', 'Crd Hrs.', 'Activity', 'Enrolled', 'Days', 'From', 'To', 'Room', 'Instructor']

    print(f"Columns in {sheet_name}:", df.columns)

    room_schedule = {
        "name": sheet_name,
        "sunday": [],
        "monday": [],
        "tuesday": [],
        "wednesday": [],
        "thursday": [],
    }

    for _, row in df.iterrows():
        if isinstance(row['From'], str) and row['From'].lower() == 'from':
            continue
        
        if pd.isnull(row['From']) or pd.isnull(row['To']):
            continue 

        time_start = row['From']
        time_end = row['To']

        if isinstance(time_start, pd.Timestamp):
            time_start = time_start.time()
        if isinstance(time_end, pd.Timestamp):
            time_end = time_end.time()

        days = str(row['Days']).strip()
        course_name = row['Course Name']

        day_map = {
            '1': 'sunday',
            '2': 'monday',
            '3': 'tuesday',
            '4': 'wednesday',
            '5': 'thursday'
        }
        
        day_numbers = days.split()

        for day_num in day_numbers:
            if day_num in day_map:
                day_name = day_map[day_num]
                room_schedule[day_name].append({
                    "timeStart": {
                        "hour": time_start.hour,
                        "minute": time_start.minute
                    },
                    "timeEnd": {
                        "hour": time_end.hour,
                        "minute": time_end.minute
                    },
                    "courseName": course_name
                })

    rooms[sheet_name] = room_schedule

json_file_path = 'output.json'
with open(json_file_path, 'w') as json_file:
    json.dump(rooms, json_file, indent=4)

print("JSON is successfully saved to:", json_file_path)
