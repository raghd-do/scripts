import pandas as pd
import json

def excel_to_json(excel_file):
    df = pd.read_excel(excel_file)
    results = []

    students = {}

    for index, row in df.iterrows():

        student_id = row['السجل المدني']

        if student_id not in students :

            student = {
                'serializableId': student_id,
                'student_info': {
                    'الإسم': row['الإسم'],
                    'السجل المدني': student_id,
                    'الفئة': row['الفئة']
                },
                'exams': []
            }
            students[student_id] = student

        exam = {
            'اسم الدار': row['اسم الدار'],
            'المستوى': row['المستوى'],
            'المعدل': row['المعدل'],
            'المنهج': row['المنهج'],
            'السنة': row['السنة'],  # يمكنك تعديل السنة والفصل حسب الحاجة
            'الفصل': row['الفصل']

        }
        students[student_id]['exams'].append(exam)

    return list(students.values())

# استبدل 'your_file.xlsx' باسم ملف الإكسل الخاص بك
excel_file = 'database.xlsx'
json_data = excel_to_json(excel_file)

# تصدير البيانات إلى ملف JSON
with open('output.json', 'w', encoding='utf-8') as f:
    json.dump(json_data, f, ensure_ascii=False, indent=4)

print("تم تحويل ملف الإكسل إلى JSON بنجاح.")