from openpyxl import *
import os
import itertools

# تحويل مجلد الاستمارات إلى لسته
tests = os.listdir("test_templates")
# فتح القاعدة
database = load_workbook("database\قاعدة الدورة الرمضانية.xlsx")

# فتح استمارة استمارة
for test in tests:
    student_test = load_workbook("test_templates/" + test, data_only=True)
    sheet = student_test.active
    student_name = sheet["C4"].value
    student_track = sheet["E4"].value
    student_level = sheet["E5"].value
    student_id = sheet["G4"].value
    student_teacher = sheet["G5"].value
    student_grade = sheet['G13'].value
    student_taqdeer = sheet['G14'].value

    print("coping:", student_name, student_id)

    student = (
        "لآلئ الرمضانية", student_teacher, student_name, student_id, student_track, student_level, "تلاوة", student_grade, student_taqdeer
    )

    #  فتح الورقات للقاعدة
    hodory_sheet = database["حضوري"]
    online_sheet = database["عن بعد"]

    # hodory_sheet.append(student)
    online_sheet.append(student)

    # move to done file
    isExist = os.path.exists("done")
    if not isExist:
        os.makedirs("done")
    os.rename(f"test_templates\{test}", f"done\{test}")

    print("copying is done :)")
    database.save(filename="database\قاعدة الدورة الرمضانية.xlsx")
