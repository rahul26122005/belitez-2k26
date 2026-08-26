# B'ELITEZ 2K26 — Python/Flask Responsive Symposium Website

## Brand hierarchy

Every major page follows:

1. MAHENDRA COLLEGE OF ENGINEERING — largest
2. DEPARTMENT OF BIOMEDICAL ENGINEERING — smaller
3. B'ELIEZ 2K26 — smallest

## Event exploration

Every Technical, Non-Technical and Workshop event has:

**Explore Event →**

Clicking it opens the individual event details page containing:
- Event description
- Rules & guidelines
- Registration CTA
- Event category
- Biomedical branding

The registration page also has **Explore Event →** links on every event selection card so participants can read the rules before choosing.

## Run

### First time

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

If PowerShell blocks activation, use:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
```

### Every later time

```powershell
.\.venv\Scripts\Activate.ps1
python app.py
```

Or double-click:

`run_windows.bat`

## URLs

Department website:

http://127.0.0.1:5000/

Symposium home:

http://127.0.0.1:5000/register-home

All events:

http://127.0.0.1:5000/events

Registration:

http://127.0.0.1:5000/register

Admin dashboard:

http://127.0.0.1:5000/admin

## Database

Registration data is stored locally in:

`belitez.db`

This version remains Python/Flask + SQLite and does not require Java/Maven.
