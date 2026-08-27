import os
import json
import firebase_admin
from firebase_admin import credentials, firestore

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, "firebase_service_account.json")

if not os.path.exists(SERVICE_ACCOUNT_FILE):
    raise FileNotFoundError(
        f"Firebase service account file not found:\n{SERVICE_ACCOUNT_FILE}"
    )

try:
    with open(SERVICE_ACCOUNT_FILE, "r", encoding="utf-8") as f:
        service_account_info = json.load(f)
except json.JSONDecodeError as e:
    raise RuntimeError(f"firebase_service_account.json contains invalid JSON: {e}")

try:
    cred = credentials.Certificate(service_account_info)

    if not firebase_admin._apps:
        firebase_admin.initialize_app(cred)

    db = firestore.client()

except Exception as e:
    raise RuntimeError(f"Firebase initialization failed: {e}")
