# ===================== assistant_core.py ====================
import speech_recognition as sr
import pyttsx3
import subprocess
import pywhatkit
import time
import platform
import os
import webbrowser
import psutil
from datetime import datetime, timedelta
import requests
import re
import pyaudio
import queue
import threading
from serpapi import GoogleSearch
import pyautogui
import wikipedia
from googletrans import Translator

# ---------------- CONFIG (API KEYS) ----------------
WEATHER_API_KEY = "836849a8608be336d996262cd8fd2c40"
NEWS_API_KEY = "5fdb52e91386ee5a9e3d53a6b3e4a62fd78ab61d5dc9f8550e83c985f63944a7"
PRICE_API_KEY = "5fdb52e91386ee5a9e3d53a6b3e4a62fd78ab61d5dc9f8550e83c985f63944a7"

screen_record_process = None

engine = pyttsx3.init()
engine.setProperty('rate', 175)
engine.setProperty("volume", 1.0)

speech_queue = queue.Queue()
translator = Translator()

recording = False
screen_record_process = None


# ---------------- TTS THREAD ----------------
def speech_worker():
    while True:
        text = speech_queue.get()
        if not text:
            continue
        engine.say(text)
        engine.runAndWait()


threading.Thread(target=speech_worker, daemon=True).start()


def speak(text):
    if not text:
        return
    speech_queue.put(text)


# ---------------- STT ----------------
def listen():
    """Listen once and return recognized text (lowercase) or None."""
    r = sr.Recognizer()

    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source, duration=0.4)
        try:
            audio = r.listen(source, timeout=6, phrase_time_limit=8)
        except sr.WaitTimeoutError:
            print("STT timeout")
            return None

    try:
        text = r.recognize_google(audio).lower()
        print("You said:", text)
        return text
    except Exception as e:
        print("STT error:", repr(e))
        return None


# ---------------- WEATHER ----------------
def extract_city(cmd):
    stop_words = [
        "weather", "temperature", "in", "today",
        "now", "what", "is", "the", "tell", "me"
    ]
    words = cmd.lower().split()
    city_words = [w for w in words if w not in stop_words]
    return " ".join(city_words).strip()


def get_weather(cmd):
    city = extract_city(cmd)
    if not city:
        return "Please tell me the city name."

    url = (
        "https://api.openweathermap.org/data/2.5/weather"
        f"?q={city}&appid={WEATHER_API_KEY}&units=metric"
    )

    res = requests.get(url).json()
    if res.get("cod") != 200:
        return f"I couldn't find weather for {city}"

    temp = res["main"]["temp"]
    desc = res["weather"][0]["description"]

    return (
        f"The weather in {city} is {desc}. "
        f"Temperature is {temp} degrees Celsius."
    )


# ---------------- NEWS ----------------
def get_news(cmd):
    topic_match = re.search(r"news.*?(?:about|on)?\s([a-zA-Z ]+)", cmd)
    topic = topic_match.group(1).strip() if topic_match else "latest news"

    url = "https://serpapi.com/search.json"
    params = {
        "engine": "google_news",
        "q": topic,
        "api_key": NEWS_API_KEY
    }

    res = requests.get(url, params=params).json()
    articles = res.get("news_results", [])

    if not articles:
        return f"No news found about {topic}"

    response = f"Here are today's top news about {topic}. "
    for i, a in enumerate(articles[:3], 1):
        response += f"{i}. {a.get('title')}. "

    return response


# ---------------- PRICE ----------------
def get_price(product, city):
    try:
        query = f"{product} price in {city} INR"

        params = {
            "engine": "google",
            "q": query,
            "hl": "en",
            "gl": "in",
            "api_key": NEWS_API_KEY
        }

        search = GoogleSearch(params)
        results = search.get_dict()

        if "organic_results" in results:
            for result in results["organic_results"]:
                snippet = result.get("snippet", "")
                price_match = re.search(r"₹\s?[\d,]+", snippet)
                if price_match:
                    return f"The price of {product} in {city} is around {price_match.group()}."

        return None

    except Exception as e:
        print("Price error:", e)
        return None


def extract_product_and_city(cmd):
    cmd = cmd.lower()
    city_match = re.search(r"in\s+([a-z ]+)$", cmd)
    if not city_match:
        return None, None

    city = city_match.group(1).strip()

    product_part = cmd[:city_match.start()]
    for w in ["price", "of", "what is", "the"]:
        product_part = product_part.replace(w, "")
    product_part = product_part.strip()

    return product_part, city


# ---------------- SYSTEM / UTILITIES ----------------
def get_wifi_name():
    try:
        output = subprocess.check_output(
            "netsh wlan show interfaces",
            shell=True,
            stderr=subprocess.DEVNULL
        ).decode(errors="ignore")

        for line in output.splitlines():
            if "SSID" in line and "BSSID" not in line:
                ssid = line.split(":")[1].strip()
                if ssid:
                    return ssid
    except Exception:
        pass

    return None


def take_screenshot():
    folder = "screenshots"
    os.makedirs(folder, exist_ok=True)

    filename = f"{folder}/screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
    pyautogui.screenshot(filename)

    speak("Screenshot captured")
    return f"Screenshot saved at {filename}"


def start_screen_recording():
    global screen_record_process

    os.makedirs("recordings", exist_ok=True)
    filename = f"recordings/record_{datetime.now().strftime('%Y%m%d_%H%M%S')}.mp4"

    screen_record_process = subprocess.Popen(
        [
            "ffmpeg",
            "-y",
            "-f", "gdigrab",
            "-framerate", "30",
            "-i", "desktop",
            "-pix_fmt", "yuv420p",
            filename
        ],
        stdin=subprocess.PIPE,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )

    return "Screen recording started"


def stop_screen_recording():
    global screen_record_process

    if screen_record_process:
        try:
            screen_record_process.stdin.write(b"q")
            screen_record_process.stdin.flush()
            screen_record_process.wait()
            screen_record_process = None
            return "Screen recording stopped and saved"
        except Exception:
            return "Failed to stop recording properly"
    else:
        return "No active recording"


def open_file_by_name(name):
    for root, dirs, files in os.walk("C:\\"):
        for file in files:
            if name.lower() in file.lower():
                path = os.path.join(root, file)
                os.startfile(path)
                return f"Opened {file}"
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

    if name in folders and os.path.exists(folders[name]):
        os.startfile(folders[name])
        return f"Opened {name} folder"

    return "Folder not found"


def open_application(name):
    apps = {
        "chrome": "chrome",
        "vs code": "code",
        "visual studio code": "code",
        "notepad": "notepad",
        "calculator": "calc",
    }

    if name in apps:
        subprocess.Popen(apps[name], shell=True)
        return f"Opening {name}"

    return None


def create_notepad():
    subprocess.Popen(["notepad.exe"])
    return "Notepad opened"


def create_word():
    try:
        subprocess.Popen("winword", shell=True)
        return "Word opened"
    except Exception:
        return "Microsoft Word not found"


def create_ppt():
    try:
        subprocess.Popen("powerpnt", shell=True)
        return "PowerPoint opened"
    except Exception:
        return "Microsoft PowerPoint not found"


# ---------------- WHATSAPP ----------------
def send_whatsapp_voice_assisted(full_cmd: str):
    # full_cmd example: "send whatsapp message i am here 9940 999567"
    # remove the trigger phrase
    text = full_cmd.replace("send whatsapp message", "").strip()
    if not text:
        speak("Please say your message and the ten digit phone number.")
        return "WhatsApp message failed"

    print("WhatsApp raw text:", repr(text))

    # extract digits
    digits_only = re.sub(r"[^\d]", "", text)
    if len(digits_only) < 10:
        speak("I could not hear a ten digit phone number. Please try again.")
        return "WhatsApp message failed"

    raw_number = digits_only[-10:]
    phone = "+91" + raw_number

    # message = part before the number
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
        now_dt = datetime.now()
        send_time = now_dt + timedelta(minutes=2)
        hour, minute = send_time.hour, send_time.minute

        speak(f"Scheduling your WhatsApp message to {phone} at {hour}:{minute}.")
        pywhatkit.sendwhatmsg(phone, message, hour, minute, wait_time=20, tab_close=True)
        speak("Your WhatsApp message has been scheduled.")
        return "WhatsApp message scheduled"
    except Exception as e:
        print("WhatsApp error:", e)
        speak("Something went wrong while sending the WhatsApp message.")
        return "WhatsApp message failed"

# ---------------- COMMAND ROUTER ----------------
def execute_command(cmd):
    cmd = cmd.lower().strip()

    # WHATSAPP
    if "send whatsapp message" in cmd:
        return send_whatsapp_voice_assisted(cmd)

    # OPEN CHROME
    if "open" in cmd and "chrome" in cmd:
        speak("Opening Chrome")
        paths = [
            r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
            r"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"
        ]
        for p in paths:
            if os.path.exists(p):
                subprocess.Popen(p)
                return "Opening Chrome"
        return "Chrome not found"

    # REFRESH
    if "refresh" in cmd:
        subprocess.Popen(
            "powershell -cmd \"Add-Type -AssemblyName System.Windows.Forms;"
            "[System.Windows.Forms.SendKeys]::SendWait('{F5}')\"",
            shell=True
        )
        speak("Refreshing the window")
        return "Refreshing the current window"

    # TIME
    if "time" in cmd:
        now_str = datetime.now().strftime("%I:%M %p")
        speak(f"The time is {now_str}")
        return f"Time: {now_str}"

    # BATTERY
    if "battery" in cmd or "charge" in cmd:
        b = psutil.sensors_battery()
        status = "charging" if b.power_plugged else "not charging"
        speak(f"Battery is {b.percent} percent and {status}")
        return f"Battery: {b.percent}% ({status})"

    # WIFI
    if "wifi" in cmd:
        wifi = get_wifi_name()
        if wifi:
            response = f"You are connected to Wi-Fi named {wifi}."
        else:
            response = "Wi-Fi is not connected."
        speak(response)
        return response

    # GOOGLE SEARCH
    if "search" in cmd:
        q = cmd.replace("search", "").strip()
        webbrowser.open(f"https://www.google.com/search?q={q}")
        speak(f"Searching for {q}")
        return f"Searching Google for {q}"

    # VOLUME
    if "volume" in cmd:
        if "high" in cmd or "up" in cmd:
            subprocess.Popen("powershell (New-Object -ComObject WScript.Shell).SendKeys([char]175)", shell=True)
            speak("Volume increased")
            return "Volume increased"
        if "low" in cmd or "down" in cmd:
            subprocess.Popen("powershell (New-Object -ComObject WScript.Shell).SendKeys([char]174)", shell=True)
            speak("Volume decreased")
            return "Volume decreased"

    # WEATHER
    if "weather" in cmd:
        result = get_weather(cmd)
        speak(result)
        return result

    # NEWS
    if "news" in cmd:
        result = get_news(cmd)
        speak(result)
        return result

    # PRICE
    if "price" in cmd:
        product, city = extract_product_and_city(cmd)
        if not product or not city:
            response = "Please tell me the product and the city."
            speak(response)
            return response

        price = get_price(product, city)
        if price:
            speak(price)
            return price
        else:
            response = f"I couldn't find the price for {product} in {city}."
            speak(response)
            return response

    # SCREENSHOT
    if any(x in cmd for x in ["take screenshot", "take a screenshot"]):
        return take_screenshot()

    # SCREEN RECORD
    if any(x in cmd for x in ["record screen", "start recording"]):
        return start_screen_recording()

    if any(x in cmd for x in ["stop recording", "stop screen recording"]):
        return stop_screen_recording()

    # CREATE FILES
    if "create notepad" in cmd or "open notepad" in cmd:
        return create_notepad()

    if "open word" in cmd or "create word" in cmd:
        return create_word()

    if "open powerpoint" in cmd or "create ppt" in cmd:
        return create_ppt()

    # OPEN FILE & FOLDER
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

    # WIKIPEDIA
    if "wikipedia" in cmd:
        topic = cmd.replace("wikipedia", "").strip()
        try:
            summary = wikipedia.summary(topic, sentences=2)
            speak(summary)
            return summary
        except Exception:
            return "I couldn't find information on that topic"

    # TRANSLATE
    if "translate" in cmd and "to" in cmd:
        text = cmd.split("translate", 1)[1].split("to")[0].strip()
        lang = cmd.split("to")[-1].strip()
        try:
            translated = translator.translate(text, dest=lang)
            speak(translated.text)
            return translated.text
        except Exception:
            return "Translation failed"

    # EXIT
    if "exit" in cmd or "stop assistant" in cmd:
        speak("Stopping assistant. Goodbye")
        os._exit(0)

    speak("Sorry, I didn't understand")
    return "Sorry, I didn't understand"
