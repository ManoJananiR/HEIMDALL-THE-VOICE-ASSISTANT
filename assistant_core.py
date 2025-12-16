import os
import platform
import queue
import re
import subprocess
import threading
from datetime import datetime, timedelta
import psutil
import pyautogui
import pyttsx3
import pywhatkit
import requests
import speech_recognition as sr
import wikipedia
from googletrans import Translator
from serpapi import GoogleSearch
from PIL import Image
import pytesseract
import webbrowser

# ---------- CONFIG ----------
WEATHER_API_KEY = "836849a8608be336d996262cd8fd2c40"
NEWS_API_KEY = "5fdb52e91386ee5a9e3d53a6b3e4a62fd78ab61d5dc9f8550e83c985f63944a7"
PRICE_API_KEY = NEWS_API_KEY  # same key used

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# ---------- GLOBALS ----------
_engine = pyttsx3.init()
_engine.setProperty("rate", 165)
_engine.setProperty("volume", 1.0)

_speech_queue = queue.Queue()      # type: queue.Queue[str]
_translator = Translator()
_screen_record_process = None      # type: subprocess.Popen

# ============================================================
# TTS / STT
# ============================================================

def _speech_worker():
    while True:
        text = _speech_queue.get()
        if not text:
            continue
        _engine.say(text)
        _engine.runAndWait()


threading.Thread(target=_speech_worker, daemon=True).start()


def speak(text):
    """Queue text for speaking (non‑blocking)."""
    if not text:
        return
    _speech_queue.put(text)


def listen(timeout=5, phrase_limit=6):
    """Listen once from the default microphone and return lowercase text or None."""
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source, duration=0.4)
        try:
            audio = recognizer.listen(source, timeout=timeout, phrase_time_limit=phrase_limit)
        except sr.WaitTimeoutError:
            print("STT timeout")
            return None

    try:
        text = recognizer.recognize_google(audio).lower()
        print("You said:", text)
        return text
    except Exception as e:
        print("STT error:", repr(e))
        return None

# ============================================================
# NETWORK / WIFI
# ============================================================

def _get_wifi_name():
    """Return current Wi‑Fi SSID on Windows, else None."""
    if platform.system() != "Windows":
        return None
    try:
        out = subprocess.check_output(
            ["netsh", "wlan", "show", "interfaces"],
            encoding="utf-8",
            stderr=subprocess.DEVNULL,
        )
        for line in out.splitlines():
            if "SSID" in line and "BSSID" not in line:
                ssid = line.split(":", 1)[1].strip()
                return ssid or None
    except Exception:
        return None

def describe_wifi_connection():
    ssid = _get_wifi_name()
    msg = (
        f"Your laptop is connected to the Wi‑Fi network {ssid}."
        if ssid
        else "Your laptop is not connected to any Wi‑Fi network."
    )
    speak(msg)
    return msg

# ============================================================
# WEATHER / NEWS / PRICE
# ============================================================

def _extract_city(cmd):
    stop_words = {"weather", "temperature", "in", "today", "now",
                  "what", "is", "the", "tell", "me"}
    words = cmd.lower().split()
    city_words = [w for w in words if w not in stop_words]
    return " ".join(city_words).strip()


def get_weather(cmd):
    city = _extract_city(cmd)
    if not city:
        return "Please tell me the city name."

    url = ("https://api.openweathermap.org/data/2.5/weather"
           f"?q={city}&appid={WEATHER_API_KEY}&units=metric")
    res = requests.get(url, timeout=10).json()

    if res.get("cod") != 200:
        return f"I couldn't find weather for {city}"

    temp = res["main"]["temp"]
    desc = res["weather"][0]["description"]
    return f"The weather in {city} is {desc}. Temperature is {temp} degrees Celsius."


def get_news(cmd):
    match = re.search(r"news.*?(?:about|on)?\s([a-zA-Z ]+)", cmd)
    topic = match.group(1).strip() if match else "latest news"

    params = {
        "engine": "google_news",
        "q": topic,
        "api_key": NEWS_API_KEY,
    }
    res = requests.get("https://serpapi.com/search.json", params=params, timeout=10).json()
    articles = res.get("news_results", [])

    if not articles:
        return f"No news found about {topic}"

    titles = [a.get("title") for a in articles[:3] if a.get("title")]
    return "Here are today's top news about {}. {}".format(
        topic, " ".join(f"{i+1}. {t}." for i, t in enumerate(titles))
    )


def _extract_product_and_city(cmd):
    cmd = cmd.lower()
    city_match = re.search(r"in\s+([a-z ]+)$", cmd)
    if not city_match:
        return None, None
    city = city_match.group(1).strip()
    product_part = cmd[:city_match.start()]
    for w in ["price", "of", "what is", "the"]:
        product_part = product_part.replace(w, "")
    product = product_part.strip()
    return product or None, city


def get_price(product, city):
    try:
        params = {
            "engine": "google",
            "q": f"{product} price in {city} INR",
            "hl": "en",
            "gl": "in",
            "api_key": PRICE_API_KEY,
        }
        results = GoogleSearch(params).get_dict()
        for result in results.get("organic_results", []):
            snippet = result.get("snippet", "")
            match = re.search(r"₹\s?[\d,]+", snippet)
            if match:
                return f"The price of {product} in {city} is around {match.group()}."
        return None
    except Exception as e:
        print("Price error:", e)
        return None

# ============================================================
# SCREEN / FILE / APPS
# ============================================================

def take_screenshot():
    folder = "screenshots"
    os.makedirs(folder, exist_ok=True)
    filename = os.path.join(folder, "screenshot_{}.png".format(datetime.now().strftime("%Y%m%d_%H%M%S")))
    pyautogui.screenshot(filename)
    speak("Screenshot captured")
    return "Screenshot saved at {}".format(filename)


def start_screen_recording():
    global _screen_record_process
    os.makedirs("recordings", exist_ok=True)
    filename = os.path.join("recordings", "record_{}.mp4".format(datetime.now().strftime("%Y%m%d_%H%M%S")))
    _screen_record_process = subprocess.Popen(
        [
            "ffmpeg", "-y",
            "-f", "gdigrab",
            "-framerate", "30",
            "-i", "desktop",
            "-pix_fmt", "yuv420p",
            filename,
        ],
        stdin=subprocess.PIPE,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    return "Screen recording started"


def stop_screen_recording():
    global _screen_record_process
    if not _screen_record_process:
        return "No active recording"
    try:
        _screen_record_process.stdin.write(b"q")
        _screen_record_process.stdin.flush()
        _screen_record_process.wait()
        _screen_record_process = None
        return "Screen recording stopped and saved"
    except Exception:
        return "Failed to stop recording properly"

def open_file_by_name(filename):
    for root, dirs, files in os.walk("C:\\"):
        if filename.lower() in files:
            os.startfile(os.path.join(root, filename))
            speak(f"Opening {filename}")
            return f"Opened {filename}"
    return "File not found"

def open_folder(name):
    folders = {
        "desktop": os.path.join(os.path.expanduser("~"), "Desktop"),
        "downloads": os.path.join(os.path.expanduser("~"), "Downloads"),
        "documents": os.path.join(os.path.expanduser("~"), "Documents"),
        "pictures": os.path.join(os.path.expanduser("~"), "Pictures"),
        "videos": os.path.join(os.path.expanduser("~"), "Videos"),
        "music": os.path.join(os.path.expanduser("~"), "Music"),
    }
    path = folders.get(name)
    if path and os.path.exists(path):
        os.startfile(path)
        return "Opened {} folder".format(name)
    return "Folder not found"


def open_application(name):
    apps = {
        "chrome": "chrome",
        "vs code": "code",
        "visual studio code": "code",
        "notepad": "notepad",
        "calculator": "calc",
    }
    cmd = apps.get(name)
    if not cmd:
        return None
    subprocess.Popen(cmd, shell=True)
    return "Opening {}".format(name)


def create_notepad():
    subprocess.Popen(["notepad.exe"])
    return "Notepad opened"

def _find_office_exe(possible_names):
    """
    Try to locate an Office executable (Word/PowerPoint) by checking
    common install folders and finally relying on PATH.
    possible_names: list like ["WINWORD.EXE"] or ["POWERPNT.EXE"].
    """
    office_roots = [
        r"C:\Program Files\Microsoft Office",
        r"C:\Program Files (x86)\Microsoft Office",
        r"C:\Program Files\Microsoft Office\root",
        r"C:\Program Files (x86)\Microsoft Office\root",
    ]

    # 1) search common Office folders
    for root in office_roots:
        if not os.path.exists(root):
            continue
        for dirpath, dirnames, filenames in os.walk(root):
            for exe_name in possible_names:
                if exe_name.lower() in [f.lower() for f in filenames]:
                    return os.path.join(dirpath, exe_name)

    # 2) try plain name (if in PATH)
    for exe_name in possible_names:
        try:
            subprocess.Popen(exe_name)
            return exe_name
        except Exception:
            continue

    return None


def create_word():
    exe_path = _find_office_exe(["WINWORD.EXE", "winword.exe"])
    if not exe_path:
        return "Microsoft Word not found"
    try:
        subprocess.Popen(exe_path)
        return "Word opened"
    except Exception:
        return "Failed to open Microsoft Word"


def create_ppt():
    exe_path = _find_office_exe(["POWERPNT.EXE", "powerpnt.exe"])
    if not exe_path:
        return "Microsoft PowerPoint not found"
    try:
        subprocess.Popen(exe_path)
        return "PowerPoint opened"
    except Exception:
        return "Failed to open Microsoft PowerPoint"

# ============================================================
# WHATSAPP
# ============================================================

def send_whatsapp_voice_assisted(full_cmd):
    """
    full_cmd like: 'send whatsapp message i am here 9940 999567'
    Extract last 10 digits as phone, text before that as message.
    """
    text = full_cmd.replace("send whatsapp message", "").strip()
    if not text:
        speak("Please say your message and the ten digit phone number.")
        return "WhatsApp message failed"

    print("WhatsApp raw text:", repr(text))
    digits_only = re.sub(r"[^\d]", "", text)
    if len(digits_only) < 10:
        speak("I could not hear a ten digit phone number. Please try again.")
        return "WhatsApp message failed"

    raw_number = digits_only[-10:]
    phone = "+91" + raw_number

    no_space = text.replace(" ", "")
    idx = no_space.rfind(raw_number)
    if idx == -1:
        message = text.strip()
    else:
        count = 0
        cut_index = len(text)
        for i, ch in enumerate(text):
            if ch != " ":
                if count == idx:
                    cut_index = i
                    break
                count += 1
        message = text[:cut_index].strip()

    if not message:
        speak("I only heard the phone number, not the message. Please try again.")
        return "WhatsApp message failed"

    print("Phone interpreted as:", phone)
    print("Message interpreted as:", message)

    try:
        send_time = datetime.now() + timedelta(minutes=2)
        hour, minute = send_time.hour, send_time.minute
        speak("Scheduling your WhatsApp message to {} at {}:{}.".format(phone, hour, minute))
        pywhatkit.sendwhatmsg(phone, message, hour, minute, wait_time=20, tab_close=True)
        speak("Your WhatsApp message has been scheduled.")
        return "WhatsApp message scheduled"
    except Exception as e:
        print("WhatsApp error:", e)
        speak("Something went wrong while sending the WhatsApp message.")
        return "WhatsApp message failed"

# ============================================================
# OCR SCREEN READER
# ============================================================

def read_screen():
    """Screenshot screen, run OCR, speak the text."""
    try:
        img = pyautogui.screenshot()
        text = pytesseract.image_to_string(img).strip()
        if not text:
            msg = "I could not detect any readable text on the screen."
            speak(msg)
            return msg
        speak("Reading the screen content.")
        speak(text)
        return text
    except Exception as e:
        print("Screen read error:", e)
        msg = "Something went wrong while reading the screen."
        speak(msg)
        return msg

# ============================================================
# COMMAND ROUTER
# ============================================================

def execute_command(cmd):
    """Route a natural‑language command to the right handler."""
    cmd = cmd.lower().strip()

    # WhatsApp
    if "send whatsapp message" in cmd:
        return send_whatsapp_voice_assisted(cmd)

    # Apps / windows
    if "open" in cmd and "chrome" in cmd:
        speak("Opening Chrome")
        paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        ]
        for p in paths:
            if os.path.exists(p):
                subprocess.Popen(p)
                return "Opening Chrome"
        return "Chrome not found"

    if "refresh" in cmd:
        subprocess.Popen(
            "powershell -cmd \"Add-Type -AssemblyName System.Windows.Forms;"
            "[System.Windows.Forms.SendKeys]::SendWait('{F5}')\"",
            shell=True,
        )
        speak("Refreshing the window")
        return "Refreshing the current window"

    # Time / battery / wifi
    if "time" in cmd:
        now_str = datetime.now().strftime("%I:%M %p")
        speak("The time is {}".format(now_str))
        return "Time: {}".format(now_str)

    if "battery" in cmd or "charge" in cmd:
        b = psutil.sensors_battery()
        status = "charging" if b.power_plugged else "not charging"
        speak("Battery is {} percent and {}".format(b.percent, status))
        return "Battery: {}% ({})".format(b.percent, status)

    if "wi-fi" in cmd or "wifi" in cmd:
        return describe_wifi_connection()

    # Search / volume
    if "search" in cmd:
        q = cmd.replace("search", "").strip()
        webbrowser.open("https://www.google.com/search?q={}".format(q))
        speak("Searching for {}".format(q))
        return "Searching Google for {}".format(q)

    if "volume" in cmd:
        if "high" in cmd or "up" in cmd:
            subprocess.Popen(
                "powershell (New-Object -ComObject WScript.Shell).SendKeys([char]175)",
                shell=True,
            )
            speak("Volume increased")
            return "Volume increased"
        if "low" in cmd or "down" in cmd:
            subprocess.Popen(
                "powershell (New-Object -ComObject WScript.Shell).SendKeys([char]174)",
                shell=True,
            )
            speak("Volume decreased")
            return "Volume decreased"

    # Weather / news / price
    if "weather" in cmd:
        result = get_weather(cmd)
        speak(result)
        return result

    if "news" in cmd:
        result = get_news(cmd)
        speak(result)
        return result

    if "price" in cmd:
        product, city = _extract_product_and_city(cmd)
        if not product or not city:
            msg = "Please tell me the product and the city."
            speak(msg)
            return msg
        price = get_price(product, city)
        if price:
            speak(price)
            return price
        msg = "I couldn't find the price for {} in {}.".format(product, city)
        speak(msg)
        return msg

    # Screenshot / recording
    if any(x in cmd for x in ("take screenshot", "take a screenshot")):
        return take_screenshot()

    if any(x in cmd for x in ("record screen", "start recording")):
        return start_screen_recording()

    if any(x in cmd for x in ("stop recording", "stop screen recording")):
        return stop_screen_recording()

    # Create files
    if "create notepad" in cmd or "open notepad" in cmd:
        return create_notepad()

    if "open word" in cmd or "create word" in cmd:
        return create_word()

    if "open powerpoint" in cmd or "create ppt" in cmd:
        return create_ppt()

    # Open folder
    if cmd.startswith("open "):
        app_name = cmd.replace("open", "").strip()
        app_result = open_application(app_name)
        if app_result:
            return app_result

    if "open desktop" in cmd:
        return open_folder("desktop")
    if "open downloads" in cmd:
        return open_folder("downloads")
    if "open documents" in cmd:
        return open_folder("documents")
    
    # OPEN FILE BY NAME
    if cmd.startswith("open file"):
        # examples: "open file resume.docx", "open file project report.pdf"
        filename = cmd.replace("open file", "", 1).strip()
        if not filename:
            return "Please tell me the file name, for example: open file resume.docx"
        return open_file_by_name(filename)

    # Wikipedia
    if "wikipedia" in cmd:
        topic = cmd.replace("wikipedia", "").strip()
        try:
            summary = wikipedia.summary(topic, sentences=2)
            speak(summary)
            return summary
        except Exception:
            return "I couldn't find information on that topic"

    # Translate
    if "translate" in cmd and "to" in cmd:
        text = cmd.split("translate", 1)[1].split("to")[0].strip()
        lang = cmd.split("to")[-1].strip()
        try:
            translated = _translator.translate(text, dest=lang)
            speak(translated.text)
            return translated.text
        except Exception:
            return "Translation failed"

    # Read screen
    if any(x in cmd for x in ("read screen", "read this screen", "screen reader")):
        return read_screen()

    # Exit
    if "exit" in cmd or "stop assistant" in cmd:
        speak("Stopping assistant. Goodbye")
        os._exit(0)

    speak("Sorry, I didn't understand")
    return "Sorry, I didn't understand"
