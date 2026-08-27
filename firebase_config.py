import os
import json
import firebase_admin
from firebase_admin import credentials, firestore

firebase_json = os.environ.get("FIREBASE_SERVICE_ACCOUNT")

if not firebase_json:
    raise RuntimeError("FIREBASE_SERVICE_ACCOUNT environment variable is not set.")

try:
    firebase_config = json.loads(firebase_json)
except json.JSONDecodeError as e:
    raise RuntimeError(f"FIREBASE_SERVICE_ACCOUNT contains invalid JSON: {e}")

if not firebase_admin._apps:
    cred = credentials.Certificate(firebase_config)
    firebase_admin.initialize_app(cred)

db = firestore.client()
