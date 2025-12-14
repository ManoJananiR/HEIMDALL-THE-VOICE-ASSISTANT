import importlib
import os

pkgs = [
    'speech_recognition', 'pyttsx3', 'pyaudio', 'pycaw', 'comtypes', 'pywhatkit',
    'wikipedia', 'pyautogui', 'requests', 'bs4', 'psutil', 'numpy', 'matplotlib',
    'googletrans', 'pytesseract', 'cv2', 'twilio', 'PIL'
]

print('Checking Python packages (imports)')
missing = []
for p in pkgs:
    try:
        importlib.import_module(p)
        print(f'OK: {p}')
    except Exception as e:
        print(f'MISSING: {p} -> {e}')
        missing.append(p)

print('\nEnvironment variables')
for v in ('TWILIO_SID', 'TWILIO_AUTH', 'TWILIO_FROM'):
    val = os.environ.get(v)
    print(f'{v}:', 'SET' if val else 'NOT SET')

if missing:
    print('\nSome packages are missing. Install with: pip install -r requirements.txt')
else:
    print('\nAll checked packages appear importable.')
