from openpyxl import *
import os

tests = os.listdir("test_templates")
database = load_workbook("database\قاعدة الفصل الأول - ١٤٤٤هـ.xlsx")

for test in tests:
    student_test = load_workbook("test_templates/" + test)
    sheet = student_test.active
    student_name = sheet["C4"].value
    student_id = round(sheet["I4"].value)
    student_track = sheet.title

    print(student_name, student_id)

    hefeth_sheet = database["حفظ"]

    found = False
    for row in hefeth_sheet.iter_rows(min_row=2, max_row=len(list(hefeth_sheet.rows))):
        # (dar, halaqah, teacher, student, id, school, track, level, status) = row
        if row[3].value == student_name and row[4].value == student_id:
            print("met")
            found = True

    if not found:
        isExist = os.path.exists("not_found")
        if not isExist:
            os.makedirs("not_found")

            os.rename(f"test_templates\{test}", f"not_found\{test}")
