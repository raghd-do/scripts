import pandas as pd
from collections import defaultdict
import json

# Read Excel data
excel_data = pd.read_excel("السجل الاكاديمي.xlsx")

# Convert to JSON
json_data = excel_data.to_dict(orient='records')

# Group data by student
student_data = defaultdict(lambda: defaultdict(list))
for record in json_data:
    student_id = record['الإسم']  # Replace with the actual column name
    student_info = {
        # Replace with your column names
        'الإسم': record['الإسم'],
        'السجل المدني': record['السجل المدني'],
        'الفئة': record['الفئة'],
        # Add other columns here
    }
    exam_info = {
        # Replace with your exam info
        'اسم الدار': record['اسم الدار'],
        'السنة': record['السنة'],
        'المنهج': record['المنهج'],
        'المستوى': record['المستوى'],
        'المعدل': record['المعدل'],
        # Add other exam columns here
    }
    student_data[student_id]['student_info'] = student_info
    student_data[student_id]['exams'].append(exam_info)

# Convert defaultdict to regular dictionary
student_data = {k: dict(v) for k, v in student_data.items()}

# Save the grouped data to a JSON file
with open('grouped_data.json', 'w') as json_file:
    json.dump(student_data, json_file, indent=4)
