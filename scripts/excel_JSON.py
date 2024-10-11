import pandas as pd
import json
import re

def split_by_common_words(course_name):
    if isinstance(course_name, str):
        course_name = re.sub(r'(?<!\s)(In|The|Of|And|To|At|On)', r' \1', course_name)
        course_name = re.sub(r'(?<!\s)([a-z])([A-Z])', r'\1 \2', course_name)
        course_name = re.sub(r'([A-Z]+)([A-Z][a-z])', r'\1 \2', course_name) 
        
        course_name = re.sub(r'\s+', ' ', course_name).strip()
        course_name = ' '.join(word.capitalize() for word in course_name.split())
    else:
        course_name = "Unknown"
    
    return course_name

def correct_known_course_names(course_name):
    known_courses = {
        "HistoryofArchitecture": "History of Architecture",
        "SustainabilityinTheBuiltEnvironment": "Sustainability in The Built Environment",
        "TheoryofArchitecture2": "Theory of Architecture 2"
    }
    
    if course_name in known_courses:
        return known_courses[course_name]
    
    return split_by_common_words(course_name)

def format_course_name(course_name):
    course_name = correct_known_course_names(course_name)

    course_name = split_by_common_words(course_name)
    
    return course_name

def clean_room_name(room_name):
    if isinstance(room_name, str):
        room_name = room_name.encode('ascii', 'ignore').decode('ascii')
        room_name = room_name.strip()
    return room_name

excel_file_path = 'data/TimeTable 20241.xlsx'
xls = pd.ExcelFile(excel_file_path)

df = pd.read_excel(xls, sheet_name='Sheet2', header=3)

df.columns = ['Section Seq.', 'Course Code', 'Course Name', 'Crd Hrs.', 'Activity', 'Enrolled', 'Days', 'From', 'To', 'Room', 'Instructor']

rooms = {}

day_map = {
    '1': 'sunday',
    '2': 'monday',
    '3': 'tuesday',
    '4': 'wednesday',
    '5': 'thursday'
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
    course_name = format_course_name(row['Course Name'])  
    room_name = clean_room_name(row['Room'])  

    if room_name not in rooms:
        rooms[room_name] = {
            "name": room_name,
            "sunday": [],
            "monday": [],
            "tuesday": [],
            "wednesday": [],
            "thursday": []
        }

    day_numbers = days.split()

    for day_num in day_numbers:
        if day_num in day_map:
            day_name = day_map[day_num]
            rooms[room_name][day_name].append({
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
            
json_file_path = 'output.json'
with open(json_file_path, 'w') as json_file:
    json.dump(rooms, json_file, indent=4)

print(f"JSON is successfully saved to: {json_file_path}")
