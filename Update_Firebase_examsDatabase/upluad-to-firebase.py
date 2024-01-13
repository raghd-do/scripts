import firebase_admin
from firebase_admin import credentials, firestore
import json

#Write the [File Path] to your firebase credential file
private_key = "fir-web-codelab-79cbb-firebase-adminsdk-4ebbq-fc38a1d45d.json"

# Initialize Firebase with your project's credentials
cred = credentials.Certificate(private_key)
firebase_admin.initialize_app(cred)

# Initialize Firestore
db = firestore.client()

# Replace 'your-collection' with the name of your Firestore collection
collection_name = 'students'


# Read the grouped JSON data [local data to be apploaded]
with open('grouped_data.json') as json_file:
    grouped_data = json.load(json_file)

# Retrieve all documents in the Firestore collection
docs = db.collection(collection_name).stream()


# UPDATE part
for doc in docs:
    # Get the data from the Firestore document
    doc_data = doc.to_dict()

    # Store the student full name in a var for searching purposes
    student_name = doc_data.get("student_info", {}).get("الإسم")

    # Check if the name field in Firestore matches the corresponding student_id in JSON data
    if student_name in grouped_data:

        # Update the "exams" list in the Firestore document by appending the list objects from the JSON data
        doc_data["exams"].extend(grouped_data[student_name]["exams"])
        # Delte the updated student data from the JSON file
        del grouped_data[student_name]

        print(f'{student_name} has been updated @_@')


# Function to generate a unique serializable ID (you can customize this)
def generate_unique_id():
    import uuid
    return str(uuid.uuid4())


# CREATE part
# if grouped_data:
# Student does not exist, create a new document
for student_name, student_data in grouped_data.items():
    print(1)
    # Generate a unique serializable ID
    student_data["serializableId"] = generate_unique_id()

    # Create a new document with the student_name as the document ID
    doc_ref = db.collection(collection_name).document(student_name)
    # Set the data for the document
    doc_ref.set(student_data)

    print(f'{student_name} has been created !_!')

# Close the Firebase Admin SDK (optional)
firebase_admin.delete_app(firebase_admin.get_app())

