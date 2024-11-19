import pandas as pd 
import json
import re
from datetime import datetime, timedelta
import os

def split_by_common_words(name):
    if isinstance(name, str):
        name = re.sub(r'(?<!\s)(In|The|Of|And|To|At|On)', r' \1', name)
        name = re.sub(r'(?<!\s)([a-z])([A-Z])', r'\1 \2', name)
        name = re.sub(r'([A-Z]+)([A-Z][a-z])', r'\1 \2', name)
        name = re.sub(r'([0-9])([A-Za-z])', r'\1 \2', name)
        
        name = ' '.join(word.capitalize() for word in name.split())
        
        roman_numerals = {"Ii": "II", "Iii": "III", "Iv": "IV", "Vi": "VI", "Vii": "VII", "Viii": "VIII", "Ix": "IX", "X": "X"}
        name = ' '.join(roman_numerals.get(word, word) for word in name.split())
    else:
        name = "Unknown"
    return name

def format_course_name(course_name):
    return split_by_common_words(course_name) if pd.notnull(course_name) else "Unknown"

def format_room_name(room_name):
    return split_by_common_words(room_name) if pd.notnull(room_name) else "Unknown"

def clean_room_name(room_name):
    if isinstance(room_name, str):
        room_name = room_name.encode('ascii', 'ignore').decode('ascii')
        room_name = room_name.strip()
    return room_name

def add_free_periods(sorted_schedule):
    if not sorted_schedule:
        return []
    result = []
    for i in range(len(sorted_schedule) - 1):
        result.append(sorted_schedule[i])
        end_time = datetime.strptime(f"{sorted_schedule[i]['timeEnd']['hour']}:{sorted_schedule[i]['timeEnd']['minute']}", "%H:%M")
        next_start_time = datetime.strptime(f"{sorted_schedule[i + 1]['timeStart']['hour']}:{sorted_schedule[i + 1]['timeStart']['minute']}", "%H:%M")
        if next_start_time > end_time:
            result.append({
                "courseName": "Free",
                "timeStart": {"hour": end_time.hour, "minute": end_time.minute},
                "timeEnd": {"hour": next_start_time.hour, "minute": next_start_time.minute}
            })
    result.append(sorted_schedule[-1])  
    return result

########################
script_dir = os.path.dirname(os.path.realpath(__file__))
excel_file_path = os.path.join(script_dir, "../data/TimeTable 20241.xlsx")

xls = pd.ExcelFile(excel_file_path)
df = pd.read_excel(xls, sheet_name='Sheet2', header=3)

df.columns = ['Section Seq.', 'Course Code', 'Course Name', 'Crd Hrs.', 'Activity', 'Enrolled', 'Days', 'From', 'To', 'Room', 'Instructor']
day_map = {'1': 'sunday', '2': 'monday', '3': 'tuesday', '4': 'wednesday', '5': 'thursday'}
rooms = {}

for _, row in df.iterrows():
    if row[['Days', 'From', 'To', 'Room']].isnull().all():
        continue

    time_start = row['From']
    time_end = row['To']

    if isinstance(time_start, pd.Timestamp):
        time_start = time_start.time()
    elif isinstance(time_start, float):
        continue

    if isinstance(time_end, pd.Timestamp):
        time_end = time_end.time()
    elif isinstance(time_end, float):
        continue
    
    days = str(row['Days']).strip()

    course_name = format_course_name(row['Course Name'])
    course_code = row['Course Code'] if pd.notnull(row['Course Code']) else "Unknown"

    if course_name == "Unknown" or course_code == "Unknown":
        continue
    
    room_name = row['Room']
    if pd.isnull(room_name):
        continue  

    room_name = format_room_name(clean_room_name(str(room_name)))
    room_names = room_name.split('/')

    for individual_room in room_names:
        if individual_room not in rooms:
            rooms[individual_room] = {"name": individual_room, **{day: [] for day in day_map.values()}}

        day_numbers = days.split()
        for day_num in day_numbers:
            if day_num in day_map:
                day_name = day_map[day_num]
            
            if not any(
                entry['timeStart'] == {"hour": time_start.hour, "minute": time_start.minute} and
                entry['timeEnd'] == {"hour": time_end.hour, "minute": time_end.minute}
                for entry in rooms[individual_room][day_name]
            ):
                rooms[individual_room][day_name].append({
                    "timeStart": {"hour": time_start.hour, "minute": time_start.minute},
                    "timeEnd": {"hour": time_end.hour, "minute": time_end.minute},
                    "courseName": course_name
                })


for room, schedule in rooms.items():
    for day, courses in schedule.items():
        if day != "name":  
            sorted_courses = sorted(courses, key=lambda x: (x['timeStart']['hour'], x['timeStart']['minute']))
            rooms[room][day] = add_free_periods(sorted_courses)

#################
json_file_path = 'output.json'
with open(json_file_path, 'w') as json_file:
    json.dump(rooms, json_file, indent=4)

print(f"JSON is successfully saved to: {json_file_path}")
