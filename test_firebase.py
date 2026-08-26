from firebase_config import db

test_ref = db.collection("test").document("connection")

test_ref.set(
    {"message": "B'ELITEZ Firebase connection successful", "status": "success"}
)

print("Firebase connection successful!")
