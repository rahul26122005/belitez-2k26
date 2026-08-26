import os
import firebase_admin
from firebase_admin import credentials, firestore

# Firebase service account JSON file
SERVICE_ACCOUNT_FILE = os.path.join(
    os.path.dirname(__file__), "firebase_service_account.json"
)


# Initialize Firebase only once
if not firebase_admin._apps:

    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        raise FileNotFoundError(
            "firebase_service_account.json was not found " "in the project folder."
        )

    cred = credentials.Certificate(SERVICE_ACCOUNT_FILE)

    firebase_admin.initialize_app(cred)


# Firestore database
db = firestore.client()
