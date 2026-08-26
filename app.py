from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    flash,
    session,
    send_file,
)

from firebase_config import db

from datetime import datetime, timezone
from functools import wraps

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

import io
import os
import uuid

# ============================================================
# FLASK APPLICATION
# ============================================================

app = Flask(__name__)

# Secret key
# Local fallback is only for development.
# In Render, set FLASK_SECRET_KEY as an environment variable.
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "belitez-2k26")


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
            "Maximum presentation time: 8 minutes.",
            "Presentation followed by a short Q&A session.",
            "Carry the presentation in PPT/PDF format.",
            "Content should be original.",
        ],
    ),
    (
        "poster-presentation",
        "Poster Presentation",
        "Technical",
        "🧬",
        "Showcase your research through an attractive scientific poster.",
        [
            "Poster should be clear and readable.",
            "Participants should be available during evaluation.",
            "Originality and presentation quality will be considered.",
            "Follow the announced poster-size instructions.",
        ],
    ),
    (
        "spot-to-solve",
        "Spot to Solve",
        "Technical",
        "🧠",
        "Think fast, solve smart and demonstrate technical problem-solving.",
        [
            "Problem statements are provided at the venue.",
            "Complete the task within the given time.",
            "Judges decision is final.",
            "External assistance is not permitted.",
        ],
    ),
    (
        "medicomind",
        "Medicomind",
        "Technical",
        "💡",
        "Challenge your medical knowledge, logic and awareness.",
        [
            "Each team must consist of 3 to 4 members.",
            "Round 1 uses displayed images for 30 seconds.",
            "Rounds 2 and 3 use questions, clues and a final code.",
            "Mobile phones, internet and external assistance are prohibited.",
        ],
    ),
    (
        "cinephoria-short-film",
        "Cinephoria & Short Film",
        "Non-Technical",
        "🎬",
        "Bring storytelling, creativity and cinematic vision to the screen.",
        [
            "Film must be suitable for the symposium audience.",
            "Submit in the specified format.",
            "Original content is preferred.",
            "Story, creativity, editing and impact may be evaluated.",
        ],
    ),
    (
        "minutes-to-win",
        "Minutes to Win",
        "Non-Technical",
        "⏱️",
        "Fast challenges, quick decisions and maximum fun.",
        [
            "Participants compete in timed rounds.",
            "Rules can vary between rounds.",
            "Participation may be individual or team-based.",
            "Coordinator decision is final.",
        ],
    ),
    (
        "team-building-activity",
        "Team Building Activity",
        "Non-Technical",
        "🤝",
        "Collaborate, communicate and complete exciting team challenges.",
        [
            "Teams follow the event instructions.",
            "Every member should participate.",
            "Respectful teamwork is mandatory.",
            "Coordinator may modify rounds when necessary.",
        ],
    ),
    (
        "squad-wars",
        "Squad-Wars",
        "Non-Technical",
        "⚔️",
        "Compete as a squad through exciting collaborative challenges.",
        [
            "Teams follow the event instructions.",
            "Every member should participate.",
            "Respectful teamwork is mandatory.",
            "Coordinator decision is final.",
        ],
    ),
    (
        "workshop",
        "Workshop",
        "Workshop",
        "🛠️",
        "Learn practical concepts through an interactive hands-on session.",
        [
            "Registration is required.",
            "Report before the session begins.",
            "Seats may be limited.",
            "Workshop instructions are shared by coordinators.",
        ],
    ),
]


# ============================================================
# GLOBAL EVENT DATA FOR TEMPLATES
# ============================================================


@app.context_processor
def globals():

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
# PUBLIC WEBSITE
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
# INDIVIDUAL EVENT DETAILS
# ============================================================


@app.route("/event/<slug>")
def detail(slug):

    event = next((event for event in EVENTS if event[0] == slug), None)

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

    # --------------------------------------------------------
    # GET REQUEST
    # --------------------------------------------------------

    if request.method == "GET":

        return render_template("register.html", form={})

    # --------------------------------------------------------
    # COLLECT FORM DATA
    # --------------------------------------------------------

    fields = [
        "name",
        "department",
        "year",
        "college",
        "contact",
        "email",
        "technical_event",
        "nontechnical_event",
        "workshop",
        "food",
    ]

    form_data = {field: request.form.get(field, "").strip() for field in fields}

    # --------------------------------------------------------
    # REQUIRED FIELDS
    # --------------------------------------------------------

    required_fields = [
        "name",
        "department",
        "year",
        "college",
        "contact",
        "workshop",
        "food",
    ]

    missing_required = any(not form_data[field] for field in required_fields)

    # --------------------------------------------------------
    # EVENT VALIDATION
    # --------------------------------------------------------

    no_event_selected = (
        not form_data["technical_event"] and not form_data["nontechnical_event"]
    )

    if missing_required or no_event_selected:

        flash(
            "Please complete all required fields and select at least one event.",
            "error",
        )

        return render_template("register.html", form=form_data)

    # --------------------------------------------------------
    # CREATE UNIQUE REGISTRATION ID
    # --------------------------------------------------------

    registration_id = "BEL26-" + uuid.uuid4().hex[:8].upper()

    # --------------------------------------------------------
    # REGISTRATION DATA
    # --------------------------------------------------------

    registration_data = {
        "registration_id": registration_id,
        "name": form_data["name"],
        "college": form_data["college"],
        "department": form_data["department"],
        "contact": form_data["contact"],
        "email": form_data["email"],
        "year": form_data["year"],
        "technical_event": form_data["technical_event"],
        "nontechnical_event": form_data["nontechnical_event"],
        "workshop": form_data["workshop"],
        "food": form_data["food"],
        "status": "registered",
        "registered_at": datetime.now(timezone.utc),
    }

    # --------------------------------------------------------
    # SAVE TO FIRESTORE
    # --------------------------------------------------------

    try:

        db.collection("Registrations").document(registration_id).set(registration_data)

    except Exception as error:

        print("Firestore registration error:", error)

        flash("Registration could not be completed. Please try again.", "error")

        return render_template("register.html", form=form_data)

    # --------------------------------------------------------
    # SUCCESS
    # --------------------------------------------------------

    return redirect(url_for("success", rid=registration_id, name=form_data["name"]))


# ============================================================
# REGISTRATION SUCCESS PAGE
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

    # Already logged in
    if session.get("admin_logged_in"):

        return redirect(url_for("admin"))

    # Login submitted
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
# HELPER: GET ALL REGISTRATIONS
# ============================================================


def get_all_registrations():

    registrations = []

    try:

        documents = db.collection("Registrations").stream()

        for document in documents:

            data = document.to_dict()

            data["id"] = document.id

            registrations.append(data)

    except Exception as error:

        print("Firestore admin read error:", error)

    return registrations


# ============================================================
# HELPER: GET FILTER VALUES
# ============================================================


def get_filter_options(registrations):

    colleges = sorted(
        {
            str(registration.get("college", "")).strip()
            for registration in registrations
            if registration.get("college")
        }
    )

    departments = sorted(
        {
            str(registration.get("department", "")).strip()
            for registration in registrations
            if registration.get("department")
        }
    )

    years = sorted(
        {
            str(registration.get("year", "")).strip()
            for registration in registrations
            if registration.get("year")
        }
    )

    technical_events = sorted(
        {
            str(registration.get("technical_event", "")).strip()
            for registration in registrations
            if registration.get("technical_event")
        }
    )

    nontechnical_events = sorted(
        {
            str(registration.get("nontechnical_event", "")).strip()
            for registration in registrations
            if registration.get("nontechnical_event")
        }
    )

    statuses = sorted(
        {
            str(registration.get("status", "")).strip()
            for registration in registrations
            if registration.get("status")
        }
    )

    return {
        "colleges": colleges,
        "departments": departments,
        "years": years,
        "technical_events": technical_events,
        "nontechnical_events": nontechnical_events,
        "statuses": statuses,
    }


# ============================================================
# HELPER: READ FILTERS
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
        "status": request.args.get("status", "").strip(),
    }


# ============================================================
# HELPER: APPLY ADMIN FILTERS
# ============================================================


def filter_registrations(registrations, filters):

    filtered = []

    for registration in registrations:

        # ----------------------------------------------------
        # SEARCH
        # ----------------------------------------------------

        if filters["search"]:

            searchable_text = " ".join(
                [
                    str(registration.get("registration_id", "")),
                    str(registration.get("id", "")),
                    str(registration.get("name", "")),
                    str(registration.get("college", "")),
                    str(registration.get("department", "")),
                    str(registration.get("contact", "")),
                    str(registration.get("email", "")),
                    str(registration.get("technical_event", "")),
                    str(registration.get("nontechnical_event", "")),
                ]
            ).lower()

            if filters["search"] not in searchable_text:

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
        # NON-TECHNICAL EVENT
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
        # STATUS
        # ----------------------------------------------------

        if filters["status"] and registration.get("status") != filters["status"]:

            continue

        # ----------------------------------------------------
        # PASSED
        # ----------------------------------------------------

        filtered.append(registration)

    return filtered


# ============================================================
# HELPER: SORT REGISTRATIONS
# ============================================================


def registration_sort_key(registration):

    value = registration.get("registered_at")

    if value is None:

        return 0

    # Firestore Timestamp
    if hasattr(value, "timestamp"):

        try:

            return value.timestamp()

        except Exception:

            return 0

    # Python datetime
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

    # Get all Firestore registrations
    registrations = get_all_registrations()

    # Get filters
    filters = get_admin_filters()

    # Apply filters
    filtered_registrations = filter_registrations(registrations, filters)

    # Newest first
    filtered_registrations.sort(key=registration_sort_key, reverse=False)

    # Filter options
    filter_options = get_filter_options(registrations)

    return render_template(
        "admin.html",
        registrations=filtered_registrations,
        total_registrations=len(registrations),
        filtered_count=len(filtered_registrations),
        colleges=filter_options["colleges"],
        departments=filter_options["departments"],
        years=filter_options["years"],
        technical_events=filter_options["technical_events"],
        nontechnical_events=filter_options["nontechnical_events"],
        statuses=filter_options["statuses"],
        current_filters=filters,
    )


# ============================================================
# ADMIN EXCEL EXPORT
# ============================================================


@app.route("/admin/export")
@admin_required
def admin_export():

    # Get all registrations
    registrations = get_all_registrations()

    # Get current filters
    filters = get_admin_filters()

    # Apply same filters
    filtered_registrations = filter_registrations(registrations, filters)

    # Sort newest first
    filtered_registrations.sort(key=registration_sort_key, reverse=False)

    # ========================================================
    # CREATE WORKBOOK
    # ========================================================

    workbook = Workbook()

    worksheet = workbook.active

    worksheet.title = "Registrations"

    # ========================================================
    # EXCEL HEADERS
    # ========================================================

    headers = [
        "Registration ID",
        "Name",
        "College",
        "Department",
        "Year",
        "Contact",
        "Email",
        "Technical Event",
        "Non-Technical Event",
        "Workshop",
        "Food",
        "Status",
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
    # ADD DATA
    # ========================================================

    for registration in filtered_registrations:

        registered_at = registration.get("registered_at")

        if hasattr(registered_at, "strftime"):

            try:

                registered_at = registered_at.strftime("%Y-%m-%d %H:%M:%S")

            except Exception:

                registered_at = str(registered_at)

        worksheet.append(
            [
                registration.get("registration_id", registration.get("id", "")),
                registration.get("name", ""),
                registration.get("college", ""),
                registration.get("department", ""),
                registration.get("year", ""),
                registration.get("contact", ""),
                registration.get("email", ""),
                registration.get("technical_event", ""),
                registration.get("nontechnical_event", ""),
                registration.get("workshop", ""),
                registration.get("food", ""),
                registration.get("status", ""),
                registered_at or "",
            ]
        )

    # ========================================================
    # FREEZE HEADER
    # ========================================================

    worksheet.freeze_panes = "A2"

    # ========================================================
    # AUTO COLUMN WIDTH
    # ========================================================

    for column in worksheet.columns:

        max_length = 0

        column_letter = column[0].column_letter

        for cell in column:

            try:

                length = len(str(cell.value if cell.value is not None else ""))

                if length > max_length:

                    max_length = length

            except Exception:

                pass

        worksheet.column_dimensions[column_letter].width = min(max_length + 3, 40)

    # ========================================================
    # CREATE EXCEL IN MEMORY
    # ========================================================

    output = io.BytesIO()

    workbook.save(output)

    output.seek(0)

    # ========================================================
    # FILE NAME
    # ========================================================

    has_filters = any(value for value in filters.values())

    if has_filters:

        filename = "BELITEZ_2K26_" "Filtered_Registrations.xlsx"

    else:

        filename = "BELITEZ_2K26_" "All_Registrations.xlsx"

    # ========================================================
    # SEND FILE
    # ========================================================

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
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


@app.errorhandler(500)
def internal_server_error(error):

    return render_template("500.html"), 500


# ============================================================
# APPLICATION START
# ============================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 5000)),
        debug=True,
    )
