from openpyxl import *
import os
import itertools

tests = os.listdir("test_templates")
database = load_workbook("database\قاعدة الفصل الأول - ١٤٤٤هـ.xlsx")


for test in tests:
    student_test = load_workbook("test_templates/" + test)
    sheet = student_test.active
    student_name = sheet["C4"].value
    student_id = round(sheet["I4"].value)
    student_track = sheet.title

    print("Loking for:", student_name, student_id)

    hefeth_sheet = database["حفظ"]

    found = False
    print("in hefeth")
    for row in hefeth_sheet.iter_rows(min_row=2, max_row=len(list(hefeth_sheet.rows))):
        if row[3].value == student_name and row[4].value == student_id:
            print("we found her :D")
            found = True

            print("now we'r copying her marks")
            if row[5].value == "أمية":
                continue
            
            if student_track == '1':
                row[10].value = sheet.cell(row=20, column=9).value
                for (i, j) in itertools.zip_longest(range(11, 44), range(27, 76, 6)):
                    if sheet.cell(row=j, column=9).value == None:
                        continue
                    row[i] = sheet.cell(row=j, column=9).value

            elif student_track == '2':
                row[10] = sheet.cell(row=15, column=9).value
                for (i, j) in itertools.zip_longest(range(11, 44), range(21, 112, 5)):
                    if sheet.cell(row=j, column=9).value == None:
                        continue
                    row[i] = sheet.cell(row=j, column=9).value
            
            elif student_track == '3':
                row[10] = sheet.cell(row=14, column=9).value
                for (i, j) in itertools.zip_longest(range(11, 44), range(19, 132, 4)):
                    if sheet.cell(row=j, column=9).value == None:
                        continue
                    row[i] = sheet.cell(row=j, column=9).value
            
            elif student_track == '4':
                row[10] = sheet.cell(row=14, column=9).value
                for (i, j) in itertools.zip_longest(range(11, 44), range(19, 216, 4)):
                    if sheet.cell(row=j, column=9).value == None:
                        continue
                    row[i] = sheet.cell(row=j, column=9).value

            elif student_track == '5':
                row[10] = sheet.cell(row=16, column=9).value
                for (i, j) in itertools.zip_longest(range(11, 44), range(23, 125, 6)):
                    if sheet.cell(row=j, column=9).value == None:
                        continue
                    row[i] = sheet.cell(row=j, column=9).value

            elif student_track == '6':
                row[10] = sheet.cell(row=13, column=9).value
                for (i, j) in itertools.zip_longest(range(11, 44), range(18, 135, 4)):
                    if sheet.cell(row=j, column=9).value == None:
                        continue
                    row[i] = sheet.cell(row=j, column=9).value
            
            print("copying is done :)")

            # move to done file
            isExist = os.path.exists("done")
            if not isExist:
                os.makedirs("done")

            os.rename(f"test_templates\{test}", f"done\{test}")


    if not found:
        isExist = os.path.exists("not_found")
        if not isExist:
            os.makedirs("not_found")

        os.rename(f"test_templates\{test}", f"not_found\{test}")
