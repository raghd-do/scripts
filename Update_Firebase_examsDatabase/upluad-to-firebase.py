import firebase_admin
from firebase_admin import credentials, firestore
import json

#Write the [File Path] to your firebase credential file
private_key = "test-department-firebase-adminsdk-8nj29-770bc9b9ba.json"

# Initialize Firebase with your project's credentials
cred = credentials.Certificate(private_key)
firebase_admin.initialize_app(cred)

# Initialize Firestore
db = firestore.client()

# استبدل 'output.json' بمسار ملف JSON الخاص بك
with open('output.json', 'r', encoding='utf-8') as f:
    json_data = json.load(f)

for student in json_data:
    student_id = str(student['serializableId'])
    student_ref = db.collection('students').document(student_id)
    student_doc = student_ref.get()

    if student_doc.exists:
        # المستند موجود، قم بإضافة قائمة exams فقط
        existing_exams = student_doc.to_dict().get('exams', [])
        new_exams = student['exams']
        updated_exams = existing_exams + new_exams
        student_ref.update({'exams': updated_exams})

        print(student["student_info"]["الإسم"] + " updated")
    else:
        # المستند غير موجود، قم بإنشاء مستند جديد
        student_ref.set(student)
        print(student["student_info"]["الإسم"] + " created")

print("تم رفع البيانات إلى Firestore بنجاح.")