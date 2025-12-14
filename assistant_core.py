# ===================== assistant_core.py =====================
import speech_recognition as sr
import pyttsx3
import subprocess
import platform
import os
import webbrowser
import psutil
from datetime import datetime
import requests
import re
import queue
import threading
import subprocess
from serpapi import GoogleSearch
# ---------------- CONFIG (DUMMY API KEYS) ----------------
WEATHER_API_KEY = "836849a8608be336d996262cd8fd2c40"
NEWS_API_KEY = "5fdb52e91386ee5a9e3d53a6b3e4a62fd78ab61d5dc9f8550e83c985f63944a7"
PRICE_API_KEY = "5fdb52e91386ee5a9e3d53a6b3e4a62fd78ab61d5dc9f8550e83c985f63944a7"

recognizer = sr.Recognizer()
engine = pyttsx3.init()
engine.setProperty('rate', 175)
recognizer.pause_threshold = 0.6
recognizer.energy_threshold = 300
speech_queue = queue.Queue()
engine.setProperty("volume", 1.0)

def speech_worker():
    while True:
        text = speech_queue.get()
        engine.say(text)
        engine.runAndWait()

threading.Thread(target=speech_worker, daemon=True).start()

def speak(text):
    if not text:
        return

    def run():
        engine.stop()
        engine.say(text)
        engine.runAndWait()

    threading.Thread(target=run, daemon=True).start()

def listen():
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source, duration=0.4)
        try:
            audio = recognizer.listen(source, timeout=5, phrase_time_limit=6)
        except sr.WaitTimeoutError:
            return ""
    try:
        return recognizer.recognize_google(audio).lower()
    except:
        return ""
    
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

def get_news(cmd):
    topic_match = re.search(r"news.*?(?:about|on)?\s([a-zA-Z ]+)", cmd)
    topic = topic_match.group(1).strip() if topic_match else "latest news"

    url = "https://serpapi.com/search.json"
    params = {
        "engine": "google_news",
        "q": topic,
        "api_key":  NEWS_API_KEY
    }

    res = requests.get(url, params=params).json()
    articles = res.get("news_results", [])

    if not articles:
        return f"No news found about {topic}"

    response = f"Here are today's top news about {topic}. "
    for i, a in enumerate(articles[:3], 1):
        response += f"{i}. {a.get('title')}. "

    return response

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

        # Try to extract price from snippets
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
    product_part = product_part.replace("price", "")
    product_part = product_part.replace("of", "")
    product_part = product_part.replace("what is", "")
    product_part = product_part.replace("the", "")
    product_part = product_part.strip()

    return product_part, city

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
    except:
        pass

    return None

def execute_command(cmd):
    cmd = cmd.lower()

    # -------- OPEN CHROME --------
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

    # -------- REFRESH --------
    if "refresh" in cmd:
        subprocess.Popen(
            "powershell -command \"Add-Type -AssemblyName System.Windows.Forms;"
            "[System.Windows.Forms.SendKeys]::SendWait('{F5}')\"",
            shell=True
        )
        speak("Refreshing the window")
        return "Refreshing the current window"

    # -------- TIME --------
    if "time" in cmd:
        now = datetime.now().strftime("%I:%M %p")
        speak(f"The time is {now}")
        return f"Time: {now}"

    # -------- BATTERY --------
    if "battery" in cmd or "charge" in cmd:
        b = psutil.sensors_battery()
        status = "charging" if b.power_plugged else "not charging"
        speak(f"Battery is {b.percent} percent and {status}")
        return f"Battery: {b.percent}% ({status})"

    # -------- WIFI --------
    if "wifi" in cmd:
        wifi = get_wifi_name()

        if wifi:
            response = f"You are connected to Wi-Fi named {wifi}."
        else:
            response = "Wi-Fi is not connected."

        speak(response)
        return response

    # -------- GOOGLE SEARCH --------
    if "search" in cmd:
        q = cmd.replace("search", "").strip()
        webbrowser.open(f"https://www.google.com/search?q={q}")
        speak(f"Searching for {q}")
        return f"Searching Google for {q}"

    # -------- VOLUME --------
    if "volume" in cmd:
        if "high" in cmd or "up" in cmd:
            subprocess.Popen("powershell (New-Object -ComObject WScript.Shell).SendKeys([char]175)", shell=True)
            speak("Volume increased")
            return "Volume increased"
        if "low" in cmd or "down" in cmd:
            subprocess.Popen("powershell (New-Object -ComObject WScript.Shell).SendKeys([char]174)", shell=True)
            speak("Volume decreased")
            return "Volume decreased"

    # -------- WEATHER (API PLACEHOLDER) --------
    if "weather" in cmd:
        result = get_weather(cmd)
        speak(result)
        return result

    # -------- NEWS (API PLACEHOLDER) --------
    if "news" in cmd:
        result = get_news(cmd)
        speak(result)
        return result

    # -------- PRICE (API PLACEHOLDER) --------
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
        
    # -------- EXIT --------
    if "exit" in cmd or "stop assistant" in cmd:
        speak("Stopping assistant. Goodbye")
        os._exit(0)

    speak("Sorry, I didn't understand")
    return "Sorry, I didn't understand"


