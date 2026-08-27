from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    flash,
    session,
    send_file,
    abort,
)

from firebase_config import db


from datetime import datetime, timezone
from functools import wraps

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

import os
import io
import uuid

# ============================================================
# FLASK APPLICATION
# ============================================================

app = Flask(__name__)

app.secret_key = os.environ.get("FLASK_SECRET_KEY", "belitez-2k26")

# Maximum HTTP request size = 10 MB.
# Actual payment image stored in Firestore is limited to 800 KB.
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024


# ============================================================
# PAYMENT FILE SETTINGS
# ============================================================

MAX_PAYMENT_FILE_SIZE = 10 * 1024 * 1024

ALLOWED_PAYMENT_EXTENSIONS = {
    "jpg",
    "jpeg",
    "png",
    "webp",
}

ALLOWED_PAYMENT_MIME_TYPES = {
    "image/jpeg",
    "image/jpg",
    "image/png",
    "image/webp",
}


# ============================================================
# FIRESTORE PAYMENT IMAGE SETTINGS
#
# Payment screenshots are stored directly INSIDE the same
# Firestore registration document as a bytes field.
#
# Firestore has a 1 MiB document limit, so the uploaded image
# is limited to 800 KB to leave room for the other registration
# fields.
# ============================================================

MAX_FIRESTORE_IMAGE_SIZE = 800 * 1024


# ============================================================
# ADMIN CREDENTIALS
# ============================================================

ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "admin")

ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "bmebelitez")


# ============================================================
# EVENT DATA
# ============================================================

EVENTS = [
    (
        "ppt-presentation",
        "PPT Presentation",
        "Technical",
        "📊",
        "Present your innovative idea, research or project with confidence.",
        [
            "1. Topic: Select a paper related to the given theme or technical field.",
            "2. Originality: Present original work and avoid copied content.",
            "3. Time Limit: Complete the presentation within the allotted 5 minutes.",
            "4. PPT Format: Use clear headings, diagrams, images, and key points",
            "5. Slides: Keep the presentation concise, preferably within 10 slides",
            "6. Team Size: Follow the specified team size of 1–4 members.",
            "7. Submission: Submit the paper/PPT before the given deadline.",
            "8. Language & Q&A: Present in English and answer judges’ questions confidently.",
            "9. References & Plagiarism: Include references and ensure the paper is plagiarism-free.",
            "10. Discipline & Decision: Maintain professional conduct; judges’ decision is final, and rule violations may lead to disqualification.",
        ],
    ),
    (
        "poster-presentation",
        "Poster Presentation",
        "Technical",
        "🧬",
        "Showcase your research through an attractive scientific poster.",
        [
            "1. Maximum 3 members.",
            "2. Participants must use poster size from A3 to A1 for Hardcopy.",
            "3. Complete the presentation within time limit of 5 minutes.",
            "4. Coordinator number - 7448665022",
        ],
    ),
    (
        "spot-to-solve",
        "Spot to Solve",
        "Technical",
        "🧠",
        "Think fast, solve smart and demonstrate technical problem-solving.",
        [
            "Team Event - 3 to 4 members per team.",
            "Round 1 : knowledge knockout.",
            "Round 2 : mystery medical box.",
            "Round 3 : think about case study.",
            "Rules:",
            "Should not communicate with other team.",
            "Mobile phone necessary for first round.",
            "Coordinator number - 9363082703 / 75300 20522",
        ],
    ),
    (
        "medicomind",
        "Medicomind",
        "Technical",
        "💡",
        "Challenge your medical knowledge, logic and awareness.",
        [
            "Round 1 – Memory Blast.",
            "Round 2 – Clue Connect. ",
            "Round 3 – Ultimate Mind Lock. ",
            "General Rules",
            "   1. Time Limit: Each round must be completed within the given time.",
            "   2. No Cheating: Mobile phones, internet, or outside assistance are not allowed.",
            "   3. Follow the Challenge: Participants must follow the instructions given by the event coordinators for each round.",
            "   4. Points: Correct answers and successful completion earn points. The team with the highest overall score will be declared the winner.",
            "Winner Announcement: Auditorium:",
            "Team Size: 3 - 4Members per Team",
            "Contact: 6379840842 / 9176394047",
        ],
    ),
    (
        "cinephoria-short-film",
        "Cinephoria & Short Film",
        "Non-Technical",
        "🎬",
        "Bring storytelling, creativity and cinematic vision to the screen.",
        [
            "Each Team Has 4 Members (Max) ",
            "Strictly Smartphone Not Allowed ",
            "Winners Will Be Move To Next Round",
            "Contact: 9344032349",
        ],
    ),
    (
        "minutes-to-win",
        "Minutes to Win",
        "Non-Technical",
        "⏱️",
        "Fast challenges, quick decisions and maximum fun.",
        [
            "Round 1 – Fake or Fact",
            "Round 2 – Think Tank",
            "Round 3 – Sketch & Guess",
            "General Rules",
            "1.Time Limit: Each challenge must be completed within the given time.",
            "2.No Cheating: No discussion, mobile phones, or outside help is allowed.",
            "3.Follow the Challenge: Participants must follow the rules and instructions given for each round.",
            "4.Points: Correct completion earns points.",
            " Contact: 8072583996 / 8667564780",
        ],
    ),
    (
        "team-building-activity",
        "Team Building Activity",
        "Non-Technical",
        "🤝",
        "Collaborate, communicate and complete exciting team challenges.",
        [
            "3 Members per Team | 4 Games | Timed Rounds",
            "       1. Catch the Ball – One hand only; ball drops = game over. 40 sec",
            "       2. Cup Pyramid Challenge – No hands; use mouth only. 2–3 min",
            "       3. Ball & Pen Challenge – Handle the ball using only a pen. 1 min 30 sec",
            "       4. Balloon Cup Pyramid – Build the pyramid while keeping the balloon in the air. 1 min 30 sec",
            "       5. Winner: Overall performance & timing will decide the winner.",
            "Note: No elimination; may be introduced only if participation is very high.",
            "Contact - 9025792732 / 6379762101",
        ],
    ),
    (
        "squad-wars",
        "Squad-Wars",
        "Non-Technical",
        "⚔️",
        "Compete as a squad through exciting collaborative challenges.",
        [
            "1.Team Event – 4 members per team, full map",
            "2. No hacks. NO EMOTE",
        ],
    ),
    (
        "workshop",
        "Workshop",
        "Workshop",
        "🛠️",
        "Learn practical concepts through an interactive hands-on session.",
        [
            "Title : advanced respiratory care equipment hands-on training and technology",
            "   1. Hands-on Equipment Training – Practical training on modern respiratory care devices and their operation.",
            "   2. Mechanical Ventilation – Learn the basics of ventilators, ventilation modes, and patient monitoring",
            "   3. Respiratory Monitoring – Understand pulse oximetry, capnography, and other respiratory monitoring technologies.",
            "   4. Clinical Application & Safety – Practice proper equipment handling, troubleshooting, infection control, and patient safety.",
        ],
    ),
]


# ============================================================
# GLOBAL EVENT DATA
# ============================================================


@app.context_processor
def global_data():

    return {
        "events": [
            {
                "slug": event[0],
                "name": event[1],
                "category": event[2],
                "icon": event[3],
                "short": event[4],
                "rules": event[5],
            }
            for event in EVENTS
        ]
    }


# ============================================================
# FILE VALIDATION
# ============================================================


def allowed_payment_file(file):

    if file is None:
        return False

    if not file.filename:
        return False

    filename = file.filename.lower().strip()

    if "." not in filename:
        return False

    extension = filename.rsplit(".", 1)[1]

    if extension not in ALLOWED_PAYMENT_EXTENSIONS:
        return False

    content_type = (file.content_type or "").lower()

    if content_type not in ALLOWED_PAYMENT_MIME_TYPES:
        return False

    return True


# ============================================================
# UPLOAD PAYMENT SCREENSHOT
# ============================================================


def upload_payment_screenshot(payment_file, registration_id):
    """
    Validate the payment screenshot and return information that can
    be stored directly in the SAME Firestore document.

    No Firebase Storage bucket is used.
    No local file is created.

    Firestore stores the actual image as a bytes field:
        payment_screenshot["image_bytes"]
    """

    if not allowed_payment_file(payment_file):
        raise ValueError("Only JPG, JPEG, PNG and WEBP images are allowed.")

    # --------------------------------------------------------
    # FILE SIZE
    # --------------------------------------------------------

    payment_file.seek(0, os.SEEK_END)
    file_size = payment_file.tell()
    payment_file.seek(0)

    if file_size <= 0:
        raise ValueError("The payment screenshot is empty.")

    # The HTTP request can still be up to 10 MB, but the actual
    # image stored in Firestore must remain below the document limit.
    if file_size > MAX_FIRESTORE_IMAGE_SIZE:
        raise ValueError(
            "Payment screenshot must be 800 KB or smaller because "
            "the image is stored directly inside Firestore."
        )

    # --------------------------------------------------------
    # READ IMAGE INTO MEMORY
    # --------------------------------------------------------

    image_bytes = payment_file.read()

    if not image_bytes:
        raise ValueError("The payment screenshot could not be read.")

    if len(image_bytes) > MAX_FIRESTORE_IMAGE_SIZE:
        raise ValueError(
            "Payment screenshot must be 800 KB or smaller because "
            "the image is stored directly inside Firestore."
        )

    # --------------------------------------------------------
    # FILE INFORMATION
    # --------------------------------------------------------

    original_filename = payment_file.filename
    content_type = payment_file.content_type or "application/octet-stream"

    # --------------------------------------------------------
    # FIRESTORE PAYMENT IMAGE DATA
    # --------------------------------------------------------
    #
    # IMPORTANT:
    # image_bytes is stored directly in Firestore.
    # There is no storage_path, bucket, blob, or Firebase Storage.
    # --------------------------------------------------------

    return {
        "file_name": original_filename,
        "stored_file_name": f"{registration_id}_{uuid.uuid4().hex}",
        "content_type": content_type,
        "size_bytes": len(image_bytes),
        "storage_type": "firestore",
        "image_bytes": image_bytes,
    }


# ============================================================
# PUBLIC ROUTES
# ============================================================


@app.route("/")
def department():

    return render_template("department.html")


@app.route("/register-home")
def home():

    return render_template("index.html")


@app.route("/events")
def events():

    return render_template("events.html")


# ============================================================
# EVENT DETAILS
# ============================================================


@app.route("/event/<slug>")
def detail(slug):

    event = next(
        (event for event in EVENTS if event[0] == slug),
        None,
    )

    if not event:

        return render_template("404.html"), 404

    event_data = {
        "slug": event[0],
        "name": event[1],
        "category": event[2],
        "icon": event[3],
        "short": event[4],
        "rules": event[5],
    }

    return render_template("event_detail.html", event=event_data)


# ============================================================
# REGISTRATION
# ============================================================


@app.route("/register", methods=["GET", "POST"])
def register():

    if request.method == "GET":

        return render_template("register.html", form={})

    # ========================================================
    # FORM DATA
    # ========================================================

    form_data = {
        "name": request.form.get("name", "").strip(),
        "college": request.form.get("college", "").strip(),
        "department": request.form.get("department", "").strip(),
        "contact": request.form.get("contact", "").strip(),
        "email": request.form.get("email", "").strip(),
        "year": request.form.get("year", "").strip(),
        "technical_event": request.form.get("technical_event", "").strip(),
        "nontechnical_event": request.form.get("nontechnical_event", "").strip(),
        "workshop": request.form.get("workshop", "").strip(),
        "food": request.form.get("food", "").strip(),
        "transaction_id": request.form.get(
            "transaction_id", request.form.get("Transaction", "")
        ).strip(),
    }

    # ========================================================
    # REQUIRED FIELDS
    # ========================================================

    required_fields = [
        "name",
        "college",
        "department",
        "contact",
        "email",
        "year",
        "workshop",
        "food",
        "transaction_id",
    ]

    missing = any(not form_data[field] for field in required_fields)

    no_event = not form_data["technical_event"] and not form_data["nontechnical_event"]

    if missing or no_event:

        flash(
            "Please complete all required fields and select at least one event.",
            "error",
        )

        return render_template("register.html", form=form_data)

    # ========================================================
    # CONTACT
    # ========================================================

    if len(form_data["contact"]) != 10 or not form_data["contact"].isdigit():

        flash("Please enter a valid 10-digit contact number.", "error")

        return render_template("register.html", form=form_data)

    # ========================================================
    # PAYMENT SCREENSHOT
    # ========================================================

    payment_file = request.files.get("payment_screenshot")

    if not payment_file:

        flash("Payment screenshot is required.", "error")

        return render_template("register.html", form=form_data)

    if not allowed_payment_file(payment_file):

        flash("Only JPG, JPEG, PNG and WEBP images are allowed.", "error")

        return render_template("register.html", form=form_data)

    # ========================================================
    # REGISTRATION ID
    # ========================================================

    registration_id = "BEL26-" + uuid.uuid4().hex[:8].upper()

    # ========================================================
    # UPLOAD PAYMENT SCREENSHOT
    # ========================================================

    try:

        payment_info = upload_payment_screenshot(payment_file, registration_id)

    except ValueError as error:

        flash(str(error), "error")

        return render_template("register.html", form=form_data)

    except Exception as error:

        print("PAYMENT STORAGE ERROR:", error)

        flash("Payment screenshot upload failed. Please try again.", "error")

        return render_template("register.html", form=form_data)

    # ========================================================
    # ONE REGISTRATION DATA DICTIONARY
    # ========================================================

    registration_data = {
        # ----------------------------------------------------
        # REGISTRATION
        # ----------------------------------------------------
        "registration_id": registration_id,
        # ----------------------------------------------------
        # PARTICIPANT
        # ----------------------------------------------------
        "name": form_data["name"],
        "college": form_data["college"],
        "department": form_data["department"],
        "contact": form_data["contact"],
        "email": form_data["email"],
        "year": form_data["year"],
        # ----------------------------------------------------
        # EVENTS
        # ----------------------------------------------------
        "technical_event": form_data["technical_event"],
        "nontechnical_event": form_data["nontechnical_event"],
        # ----------------------------------------------------
        # WORKSHOP / FOOD
        # ----------------------------------------------------
        "workshop": form_data["workshop"],
        "food": form_data["food"],
        # ----------------------------------------------------
        # PAYMENT
        # ----------------------------------------------------
        "transaction_id": form_data["transaction_id"],
        # ----------------------------------------------------
        # PAYMENT SCREENSHOT
        #
        # Nested INSIDE the SAME document.
        # ----------------------------------------------------
        "payment_screenshot": {
            "file_name": payment_info["file_name"],
            "stored_file_name": payment_info["stored_file_name"],
            "content_type": payment_info["content_type"],
            "size_bytes": payment_info["size_bytes"],
            "storage_type": "firestore",
            "image_bytes": payment_info["image_bytes"],
        },
        # ----------------------------------------------------
        # STATUS
        # ----------------------------------------------------
        "status": "registered",
        "payment_status": "pending",
        # ----------------------------------------------------
        # TIMESTAMP
        # ----------------------------------------------------
        "registered_at": datetime.now(timezone.utc),
    }

    # ========================================================
    # SAVE TO FIRESTORE
    #
    # ONLY ONE COLLECTION
    # ========================================================

    try:

        db.collection("Registrations").document(registration_id).set(registration_data)

    except Exception as error:

        print("FIRESTORE ERROR:", error)

        flash("Registration could not be completed. Please try again.", "error")

        return render_template("register.html", form=form_data)

    # ========================================================
    # SUCCESS
    # ========================================================

    return redirect(url_for("success", rid=registration_id, name=form_data["name"]))


# ============================================================
# SUCCESS
# ============================================================


@app.route("/success")
def success():

    return render_template(
        "success.html",
        rid=request.args.get("rid", ""),
        name=request.args.get("name", ""),
    )


# ============================================================
# ADMIN AUTHENTICATION
# ============================================================


def admin_required(view):

    @wraps(view)
    def wrapped_view(*args, **kwargs):

        if not session.get("admin_logged_in"):

            return redirect(url_for("admin_login"))

        return view(*args, **kwargs)

    return wrapped_view


# ============================================================
# ADMIN LOGIN
# ============================================================


@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():

    if session.get("admin_logged_in"):

        return redirect(url_for("admin"))

    if request.method == "POST":

        username = request.form.get("username", "").strip()

        password = request.form.get("password", "")

        if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:

            session.clear()

            session["admin_logged_in"] = True

            return redirect(url_for("admin"))

        flash("Invalid username or password.", "error")

    return render_template("admin_login.html")


# ============================================================
# ADMIN LOGOUT
# ============================================================


@app.route("/admin/logout")
def admin_logout():

    session.clear()

    return redirect(url_for("admin_login"))


# ============================================================
# GET ALL REGISTRATIONS
# ============================================================


def get_all_registrations():

    registrations = []

    try:

        documents = db.collection("Registrations").stream()

        for document in documents:

            data = document.to_dict()

            data["id"] = document.id

            # ------------------------------------------------
            # NORMALIZE PAYMENT DATA
            # ------------------------------------------------

            payment = data.get("payment_screenshot")

            if not isinstance(payment, dict):

                payment = {}

            data["payment_screenshot_data"] = payment

            # ------------------------------------------------
            # Payment display values
            # ------------------------------------------------

            data["payment_file_name"] = payment.get(
                "file_name", data.get("payment_screenshot_name", "")
            )

            data["payment_storage_path"] = (
                "Firestore document" if payment.get("image_bytes") else ""
            )

            data["payment_file_type"] = payment.get(
                "content_type", data.get("payment_screenshot_type", "")
            )

            data["payment_file_size"] = payment.get(
                "size_bytes", data.get("payment_screenshot_size", 0)
            )

            data["payment_storage_type"] = payment.get(
                "storage_type",
                "firestore" if payment.get("image_bytes") else "",
            )

            registrations.append(data)

    except Exception as error:

        print("FIRESTORE READ ERROR:", error)

    return registrations


# ============================================================
# FORMAT FILE SIZE
# ============================================================


def format_file_size(size):

    try:

        size = int(size)

    except Exception:

        return "-"

    if size <= 0:
        return "-"

    if size < 1024:

        return f"{size} B"

    if size < 1024 * 1024:

        return f"{size / 1024:.1f} KB"

    return f"{size / (1024 * 1024):.2f} MB"


# ============================================================
# FORMAT TIMESTAMP
# ============================================================


def format_timestamp(value):

    if value is None:
        return "-"

    try:

        if hasattr(value, "to_datetime"):

            value = value.to_datetime()

        if hasattr(value, "strftime"):

            return value.strftime("%d-%m-%Y %I:%M %p")

    except Exception:

        pass

    return str(value)


# ============================================================
# ADMIN FILTERS
# ============================================================


def get_admin_filters():

    return {
        "search": request.args.get("search", "").strip().lower(),
        "college": request.args.get("college", "").strip(),
        "department": request.args.get("department", "").strip(),
        "year": request.args.get("year", "").strip(),
        "technical_event": request.args.get("technical_event", "").strip(),
        "nontechnical_event": request.args.get("nontechnical_event", "").strip(),
        "workshop": request.args.get("workshop", "").strip(),
        "food": request.args.get("food", "").strip(),
        "payment_status": request.args.get("payment_status", "").strip(),
        "status": request.args.get("status", "").strip(),
    }


# ============================================================
# FILTER REGISTRATIONS
# ============================================================


def filter_registrations(registrations, filters):

    filtered = []

    for registration in registrations:

        # ----------------------------------------------------
        # SEARCH
        # ----------------------------------------------------

        if filters["search"]:

            payment = registration.get("payment_screenshot_data", {})

            searchable = " ".join(
                [
                    str(registration.get("registration_id", "")),
                    str(registration.get("name", "")),
                    str(registration.get("college", "")),
                    str(registration.get("department", "")),
                    str(registration.get("contact", "")),
                    str(registration.get("email", "")),
                    str(registration.get("year", "")),
                    str(registration.get("technical_event", "")),
                    str(registration.get("nontechnical_event", "")),
                    str(registration.get("workshop", "")),
                    str(registration.get("food", "")),
                    str(registration.get("transaction_id", "")),
                    str(registration.get("payment_status", "")),
                    str(payment.get("file_name", "")),
                ]
            ).lower()

            if filters["search"] not in searchable:

                continue

        # ----------------------------------------------------
        # COLLEGE
        # ----------------------------------------------------

        if filters["college"] and registration.get("college") != filters["college"]:

            continue

        # ----------------------------------------------------
        # DEPARTMENT
        # ----------------------------------------------------

        if (
            filters["department"]
            and registration.get("department") != filters["department"]
        ):

            continue

        # ----------------------------------------------------
        # YEAR
        # ----------------------------------------------------

        if filters["year"] and registration.get("year") != filters["year"]:

            continue

        # ----------------------------------------------------
        # TECHNICAL EVENT
        # ----------------------------------------------------

        if (
            filters["technical_event"]
            and registration.get("technical_event") != filters["technical_event"]
        ):

            continue

        # ----------------------------------------------------
        # NON TECHNICAL EVENT
        # ----------------------------------------------------

        if (
            filters["nontechnical_event"]
            and registration.get("nontechnical_event") != filters["nontechnical_event"]
        ):

            continue

        # ----------------------------------------------------
        # WORKSHOP
        # ----------------------------------------------------

        if filters["workshop"] and registration.get("workshop") != filters["workshop"]:

            continue

        # ----------------------------------------------------
        # FOOD
        # ----------------------------------------------------

        if filters["food"] and registration.get("food") != filters["food"]:

            continue

        # ----------------------------------------------------
        # PAYMENT STATUS
        # ----------------------------------------------------

        if (
            filters["payment_status"]
            and registration.get("payment_status", "") != filters["payment_status"]
        ):

            continue

        # ----------------------------------------------------
        # REGISTRATION STATUS
        # ----------------------------------------------------

        if filters["status"] and registration.get("status", "") != filters["status"]:

            continue

        filtered.append(registration)

    return filtered


# ============================================================
# FILTER OPTIONS
# ============================================================


def get_filter_options(registrations):

    return {
        "colleges": sorted(
            {
                str(r.get("college", "")).strip()
                for r in registrations
                if r.get("college")
            }
        ),
        "departments": sorted(
            {
                str(r.get("department", "")).strip()
                for r in registrations
                if r.get("department")
            }
        ),
        "years": sorted(
            {str(r.get("year", "")).strip() for r in registrations if r.get("year")}
        ),
        "technical_events": sorted(
            {
                str(r.get("technical_event", "")).strip()
                for r in registrations
                if r.get("technical_event")
            }
        ),
        "nontechnical_events": sorted(
            {
                str(r.get("nontechnical_event", "")).strip()
                for r in registrations
                if r.get("nontechnical_event")
            }
        ),
        "payment_statuses": sorted(
            {
                str(r.get("payment_status", "")).strip()
                for r in registrations
                if r.get("payment_status")
            }
        ),
        "statuses": sorted(
            {str(r.get("status", "")).strip() for r in registrations if r.get("status")}
        ),
    }


# ============================================================
# SORT
# ============================================================


def registration_sort_key(registration):

    value = registration.get("registered_at")

    if value is None:
        return 0

    if hasattr(value, "timestamp"):

        try:

            return value.timestamp()

        except Exception:

            return 0

    if isinstance(value, datetime):

        try:

            if value.tzinfo is None:

                value = value.replace(tzinfo=timezone.utc)

            return value.timestamp()

        except Exception:

            return 0

    return 0


# ============================================================
# ADMIN DASHBOARD
# ============================================================


@app.route("/admin")
@admin_required
def admin():

    registrations = get_all_registrations()

    filters = get_admin_filters()

    filtered_registrations = filter_registrations(registrations, filters)

    # Newest first
    filtered_registrations.sort(key=registration_sort_key, reverse=True)

    filter_options = get_filter_options(registrations)

    # --------------------------------------------------------
    # PAYMENT COUNTS
    # --------------------------------------------------------

    pending_count = sum(
        1 for r in registrations if r.get("payment_status") == "pending"
    )

    verified_count = sum(
        1 for r in registrations if r.get("payment_status") == "verified"
    )

    rejected_count = sum(
        1 for r in registrations if r.get("payment_status") == "rejected"
    )

    # --------------------------------------------------------
    # TEMPLATE
    # --------------------------------------------------------

    return render_template(
        "admin.html",
        registrations=filtered_registrations,
        total_registrations=len(registrations),
        filtered_count=len(filtered_registrations),
        pending_count=pending_count,
        verified_count=verified_count,
        rejected_count=rejected_count,
        colleges=filter_options["colleges"],
        departments=filter_options["departments"],
        years=filter_options["years"],
        technical_events=filter_options["technical_events"],
        nontechnical_events=filter_options["nontechnical_events"],
        payment_statuses=filter_options["payment_statuses"],
        statuses=filter_options["statuses"],
        current_filters=filters,
        format_file_size=format_file_size,
        format_timestamp=format_timestamp,
    )


# ============================================================
# ADMIN PAYMENT SCREENSHOT
#
# The image is retrieved directly from Firestore.
# It is NOT stored locally and does NOT use Firebase Storage.
# ============================================================


@app.route("/admin/payment/<registration_id>")
@admin_required
def admin_payment(registration_id):
    """
    Retrieve the payment screenshot directly from the same
    Firestore registration document.

    No Firebase Storage or local file is used.
    """

    document = db.collection("Registrations").document(registration_id).get()

    if not document.exists:
        abort(404)

    data = document.to_dict()
    payment = data.get("payment_screenshot", {})

    if not isinstance(payment, dict):
        payment = {}

    image_bytes = payment.get("image_bytes")

    if not image_bytes:
        abort(404)

    try:
        if not isinstance(image_bytes, bytes):
            image_bytes = bytes(image_bytes)

        content_type = payment.get("content_type") or "image/jpeg"

        return send_file(
            io.BytesIO(image_bytes),
            mimetype=content_type,
            download_name=payment.get(
                "file_name",
                "payment_screenshot",
            ),
        )

    except Exception as error:
        print("PAYMENT IMAGE ERROR:", error)
        abort(404)


# ============================================================
# ADMIN PAYMENT SCREENSHOT COMPATIBILITY ROUTE
# ============================================================
#
# The current admin.html may call the endpoint name
# "admin_payment_screenshot". Keep this compatibility endpoint
# so the existing dashboard works without changing its other code.
# The actual image still comes directly from Firestore.
# ============================================================


@app.route(
    "/admin/payment-screenshot/<registration_id>", endpoint="admin_payment_screenshot"
)
@admin_required
def admin_payment_screenshot(registration_id):
    return admin_payment(registration_id)


# ============================================================
# EXCEL EXPORT
# ============================================================


@app.route("/admin/export")
@admin_required
def admin_export():

    registrations = get_all_registrations()

    filters = get_admin_filters()

    registrations = filter_registrations(registrations, filters)

    registrations.sort(key=registration_sort_key, reverse=True)

    workbook = Workbook()

    worksheet = workbook.active

    worksheet.title = "Registrations"

    # ========================================================
    # HEADERS
    # ========================================================

    headers = [
        "Registration ID",
        "Name",
        "College",
        "Department",
        "Contact",
        "Email",
        "Year",
        "Technical Event",
        "Non-Technical Event",
        "Workshop",
        "Food",
        "Transaction ID",
        "Payment Screenshot Name",
        "Payment Screenshot Type",
        "Payment Screenshot Size",
        "Payment Image Storage",
        "Payment Status",
        "Registration Status",
        "Registered At",
    ]

    worksheet.append(headers)

    # ========================================================
    # HEADER STYLE
    # ========================================================

    for cell in worksheet[1]:

        cell.font = Font(bold=True)

        cell.alignment = Alignment(horizontal="center", vertical="center")

    # ========================================================
    # DATA
    # ========================================================

    for registration in registrations:

        payment = registration.get("payment_screenshot_data", {})

        registered_at = registration.get("registered_at")

        worksheet.append(
            [
                registration.get("registration_id", registration.get("id", "")),
                registration.get("name", ""),
                registration.get("college", ""),
                registration.get("department", ""),
                registration.get("contact", ""),
                registration.get("email", ""),
                registration.get("year", ""),
                registration.get("technical_event", ""),
                registration.get("nontechnical_event", ""),
                registration.get("workshop", ""),
                registration.get("food", ""),
                registration.get("transaction_id", ""),
                payment.get("file_name", ""),
                payment.get("content_type", ""),
                format_file_size(payment.get("size_bytes", 0)),
                payment.get("storage_path", ""),
                registration.get("payment_status", "pending"),
                registration.get("status", ""),
                format_timestamp(registered_at),
            ]
        )

    worksheet.freeze_panes = "A2"

    # ========================================================
    # COLUMN WIDTH
    # ========================================================

    for column in worksheet.columns:

        max_length = 0

        column_letter = column[0].column_letter

        for cell in column:

            try:

                max_length = max(max_length, len(str(cell.value or "")))

            except Exception:

                pass

        worksheet.column_dimensions[column_letter].width = min(max_length + 3, 60)

    # ========================================================
    # SEND EXCEL
    # ========================================================

    output = io.BytesIO()

    workbook.save(output)

    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=("BELITEZ_2K26_" "Registrations.xlsx"),
        mimetype=(
            "application/vnd.openxmlformats-" "officedocument.spreadsheetml.sheet"
        ),
    )


# ============================================================
# ERROR HANDLERS
# ============================================================


@app.errorhandler(404)
def page_not_found(error):

    return render_template("404.html"), 404


@app.errorhandler(413)
def file_too_large(error):

    flash("The uploaded payment screenshot cannot exceed 10 MB.", "error")

    return redirect(url_for("register"))


@app.errorhandler(500)
def internal_server_error(error):

    return render_template("500.html"), 500


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
