from openpyxl import Workbook, load_workbook
import os

# فتح القاعدة
# all
database = load_workbook(filename="القاعدة\قاعدة الفصل الأول - ١٤٤٤هـ.xlsx")
# some
# database = load_workbook(filename="القاعدة\قاعدة الفصل الأول - ١٤٤٤هـ.xlsx")

# # ورقة الحفظ
# hefeth_sheet = database['حفظ']

# for value in hefeth_sheet.iter_rows(min_row=1082,
#                                     max_row=1745,
#                                     min_col=1,
#                                     max_col=9,
#                                     values_only=True):
#     (dar, halagah, teacher, student, id, school, track, level, status) = value

#     # تجواز الأميات
#     if school == "أمية":
#         continue

#     test_template = load_workbook(
#         filename="قوالب\قالب - استمارات اختبارات المناهج.xlsx")

#     # delete unwanted sheets
#     for i in range(1, 7):
#         if test_template[str(i)].title != str(track):
#             if str(track) == "1-ب" and i == 1:
#                 continue
#             test_template.remove(test_template[str(i)])

#     # print(test_template.sheetnames)

#     sheet = test_template.active
#     sheet['C4'] = student
#     sheet['C5'] = dar

#     if track == 6:
#         sheet['E5'] = level
#         sheet['G4'] = id
#     else:
#         sheet['G4'] = track
#         sheet['G5'] = level
#         sheet['I4'] = id

#     test_template.save(filename=f"حفظ\{dar}\{student}.xlsx")
#     print(f"{student} {id} is done")

# ورقة التعاهد
# taahod_sheet = database['تعاهد']

# for value in taahod_sheet.iter_rows(min_row=2,
#                                     max_row=233,
#                                     min_col=1,
#                                     max_col=9,
#                                     values_only=True):
#     (dar, halagah, teacher, student, id, school, track, level, status) = value

# #     # تجواز الأميات
#     if school == "أمية":
#         continue

#     if status == "خاتمة":
#         test_template = load_workbook(
#             filename="قوالب\قالب - استمارة تعاهد للدورات .xlsx")

#         print("خاتمة")

#         sheet = test_template.active
#         sheet['C4'] = student
#         sheet['C5'] = dar
#         sheet['G4'] = track
#         sheet['G5'] = level
#         sheet['I4'] = id

#         test_template.save(filename=f"خاتمات\{student}.xlsx")

#         continue
#     else:
#         test_template = load_workbook(
#             filename="قوالب\قالب - استمارات اختبارات المناهج.xlsx")

#     # delete unwanted sheets
#     for i in range(1, 7):
#         if test_template[str(i)].title != str(track):
#             if str(track) == "1-ب" and i == 1:
#                 continue
#             test_template.remove(test_template[str(i)])

#     sheet = test_template.active
#     sheet['C4'] = student
#     sheet['C5'] = dar

#     if track == 6:
#         sheet['E5'] = level
#         sheet['G4'] = id
#     else:
#         sheet['G4'] = track
#         sheet['G5'] = level
#         sheet['I4'] = id

#     test_template.save(filename=f"تعاهد\{dar}\{student}.xlsx")
#     print(f"{student} {id} is done")

# ورقة التلقين
# talqeen_sheet = database['منهج التلقين']

# for value in talqeen_sheet.iter_rows(min_row=3,
#                                      max_row=68,
#                                      min_col=1,
#                                      max_col=8,
#                                      values_only=True):
#     (dar, halagah, teacher, student, id, school, track, level) = value

#     test_template = load_workbook(
#         filename="قوالب\قالب - استمارة اختبار التلقين التعليمي.xlsx")

#     sheet = test_template.active
#     sheet['B4'] = student
#     sheet['B5'] = dar
#     sheet['D4'] = track
#     sheet['D5'] = level
#     sheet['F4'] = id

#     test_template.save(filename=f"استمارات التلقين\{dar}\{student}.xlsx")
#     print(f"{student} {id} is done")

# الأميات
hefeth_sheet = database['تعاهد']

for value in hefeth_sheet.iter_rows(min_row=2,
                                    max_row=4,
                                    min_col=1,
                                    max_col=9,
                                    values_only=True):
    (dar, halagah, teacher, student, id, school, track, level, status) = value

    # تجواز غير الأميات
    # if school != "أمية":
    #     continue

    test_template = load_workbook(
        filename="قوالب\قالب - استمارة اختبار الأمهات والأميات.xlsx")

    # delete unwanted sheets
    for i in range(1, 7):
        if test_template[str(i)].title != str(track):
            test_template.remove(test_template[str(i)])

    sheet = test_template.active
    sheet['C4'] = student
    sheet['C5'] = dar

    if track == 6:
        sheet['E5'] = level
        sheet['G4'] = id
    else:
        sheet['F5'] = level
        sheet['H4'] = id

    isExist = os.path.exists(f"أميات\{dar}\تعاهد")

    # create the teacher folder if is not exsist
    if not isExist:
        os.makedirs(f"أميات\{dar}\تعاهد")

    test_template.save(filename=f"أميات\{dar}\تعاهد\{student}.xlsx")

    # test_template.save(filename=f"أميات\{teacher}\تعاهد\{student}.xlsx")
    print(f"{student} {id} is done")

# ورقة التلاوة
# telawah_sheet = database['منهج التلاوة']

# for value in telawah_sheet.iter_rows(min_row=18,
#                                      max_row=23,
#                                      min_col=1,
#                                      max_col=8,
#                                      values_only=True):
#     (dar, halagah, teacher, student, id, school, track, level) = value

#     test_template = load_workbook(
#         filename="قوالب\استمارات اختبارات منهج التلاوة.xlsx")

#     # delete unwanted sheets
#     for i in range(1, 9):
#         if test_template[str(i)].title != str(level):
#             test_template.remove(test_template[str(i)])

#     sheet = test_template[str(level)]
#     print(sheet)
#     sheet['C4'] = student
#     sheet['C5'] = dar
#     sheet['G4'] = id

#     Exist = os.path.exists(f"استمارات التلاوة\{dar}")

#     # create the teacher folder if is not exsist
#     if not Exist:
#         os.makedirs(f"استمارات التلاوة\{dar}")

#     test_template.save(filename=f"استمارات التلاوة\{dar}\{student}.xlsx")
#     print(f"{student} {id} is done")

print('Done :D')
