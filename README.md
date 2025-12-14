# AI-DRIVEN-VOICE-ASSISTANT-FOR-DESKTOP-AUTOMATION

A local, permission-aware voice assistant built with `Heimdall.py` (main) and `secondry.py` (helpers).

## Quick setup (Windows)

1. Create a Python virtual environment and activate it:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

2. Install dependencies:

```powershell
pip install -r requirements.txt
```

3. Optional: set Twilio environment variables if you plan to use SMS features:

```powershell
setx TWILIO_SID "your_sid"
setx TWILIO_AUTH "your_auth_token"
setx TWILIO_FROM "+1234567890"
```

4. Run environment checker (helps identify missing packages):

```powershell
python check_env.py
```

5. Run the assistant:

```powershell
python Heimdall.py
```

Or double-click `run_assistant.bat` on Windows.

6. Run tests (requires `pytest` installed):

```powershell
pip install pytest
python -m pytest -q
```

## Files
- `Heimdall.py`: main GUI and voice orchestration.
- `secondry.py`: helpers and external integrations.
- `check_env.py`: helper to check installed packages and env vars.
- `requirements.txt`: pinned dependency list.

## Notes
- This project uses system hardware (microphone, speakers, screen capture) and third-party binaries (tesseract). Ensure those are installed.
- For cloud features (GDrive/Outlook/Gmail), OAuth credentials are required and not included.
- If you want, I can add CI (GitHub Actions) and more exhaustive tests next.
