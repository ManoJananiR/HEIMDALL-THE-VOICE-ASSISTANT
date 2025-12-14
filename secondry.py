import os
import webbrowser
import pywhatkit
import wikipedia
import pyautogui
import datetime
import requests
import re
import subprocess
import platform
import psutil
import time
import numpy as np
from bs4 import BeautifulSoup
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
import json
import logging
from pathlib import Path
import shlex
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import subprocess
import threading

exchange_rates = {
    'USD': 82.0,
    'INR': 1.0,
}

amazon_domains = {
    'india': 'amazon.in',
    'mumbai': 'amazon.in',
    'chennai': 'amazon.in',
    'america': 'amazon.com',
    'usa': 'amazon.com',
    'united states': 'amazon.com',
}

def get_amazon_domain(location: str) -> str:
    for loc in amazon_domains:
        if loc in location:
            return amazon_domains[loc]
    return 'amazon.in'

def get_price_from_amazon(product_name: str, location: str, talk=None) -> float:
    api_key = "bdf3f0ad02a5690191bcae56cf0f4134"
    base_url = "https://api.scraperapi.com"
    amazon_domain = get_amazon_domain(location)
    target_url = f"https://{amazon_domain}/s?k={product_name.replace(' ', '+')}"
    scraper_url = f"{base_url}?api_key={api_key}&url={target_url}"

    try:
        response = requests.get(scraper_url, timeout=10)
        soup = BeautifulSoup(response.content, "html.parser")
    except Exception as e:
        if talk:
            talk(f"HTTP error while getting price: {e}")
        return None

    price = None
    try:
        price = soup.find("span", {"class": "a-offscreen"}).text
    except AttributeError:
        if talk:
            talk("Sorry, I couldn't find the price for that item.")

    if price:
        price = re.sub(r'[^\d.]', '', price)
        currency = 'USD' if 'amazon.com' in target_url else 'INR'
        try:
            price_in_rupees = float(price) * exchange_rates[currency]
            return price_in_rupees
        except Exception:
            return None
    return None

def check_wifi(talk=None):
    wifi_name = None
    if platform.system() == "Windows":
        try:
            result = subprocess.check_output(["netsh", "wlan", "show", "interfaces"], encoding="utf-8")
            for line in result.split("\n"):
                if "SSID" in line and "BSSID" not in line:
                    wifi_name = line.split(":")[1].strip()
                    break
        except Exception as e:
            if talk:
                talk(f"Error checking Wi-Fi: {e}")
            return
    if wifi_name:
        if talk:
            talk(f"Your laptop is connected to {wifi_name}.")
        return wifi_name
    else:
        if talk:
            talk("Your laptop is not connected to any Wi-Fi.")
        return None

def set_volume_by_percentage(percentage: int, talk=None) -> int:
    try:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = interface.QueryInterface(IAudioEndpointVolume)
        scalar = max(0.0, min(percentage / 100.0, 1.0))
        volume.SetMasterVolumeLevelScalar(scalar, None)
        current_volume = round(volume.GetMasterVolumeLevelScalar() * 100)
        if talk:
            talk(f"Volume set to {current_volume} percent.")
        return current_volume
    except Exception as e:
        if talk:
            talk(f"Unable to set volume: {e}")
        return None

def process_volume_command(take_command, talk=None):
    command = take_command() if take_command else None
    if command:
        words = command.split()
        for word in words:
            if word.isdigit():
                volume_level = int(word)
                return set_volume_by_percentage(volume_level, talk=talk)
    if talk:
        talk("I did not catch a valid number. Please try again.")
    return None

def fetch_weather(take_command, talk=None):
    if talk:
        talk("Please tell me the city name.")
    city = take_command() if take_command else None
    if city:
        api_key = "05dd4d28db225f9c4f1b2f7adeb2a774"
        url = f"https://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}&units=metric"
        try:
            response = requests.get(url).json()
            if "main" in response:
                temperature = response["main"]["temp"]
                description = response["weather"][0]["description"]
                if talk:
                    talk(f"The weather in {city} is {description} with a temperature of {temperature} degrees Celsius.")
                return response
            else:
                if talk:
                    talk("I couldn't fetch the weather details. Please try again.")
                return None
        except Exception as e:
            if talk:
                talk("There was an error fetching the weather details.")
            return None
    return None

def search_wikipedia(command: str, talk=None):
    topic = command.replace("wikipedia", "").strip()
    if topic:
        try:
            info = wikipedia.summary(topic, sentences=2)
            if talk:
                talk(info)
            return info
        except wikipedia.exceptions.DisambiguationError:
            if talk:
                talk("The topic is ambiguous. Please be more specific.")
        except wikipedia.exceptions.PageError:
            if talk:
                talk("I couldn't find any page on Wikipedia for that topic.")
    return None

def translate_text(take_command, talk=None):
    try:
        if talk:
            talk("Which word do you want to translate?")
        word = take_command() if take_command else None
        if not word:
            return None
        if talk:
            talk("To which language should I translate?")
        language = take_command() if take_command else None
        if not language:
            return None
        lang_codes = {"bengali": "bn", "chinese": "zh-cn", "english": "en", "french": "fr", "german": "de", "hindi": "hi", "italian": "it", "japanese": "ja", "korean": "ko", "portuguese": "pt", "russian": "ru", "spanish": "es", "tamil": "ta", "turkish": "tr", "hebrew": "iw"}
        lang_code = lang_codes.get(language)
        if not lang_code:
            if talk:
                talk(f"Sorry, I don't support the language '{language}'. Please try again.")
            return None
        from googletrans import Translator
        translator = Translator()
        translated_word = translator.translate(word, src='auto', dest=lang_code).text
        if talk:
            talk(f"The translation for '{word}' in {language} is '{translated_word}'.")
        return translated_word
    except Exception as e:
        if talk:
            talk("An error occurred while translating.")
        return None

def fetch_news(take_command, talk=None):
    try:
        if talk:
            talk("What topic are you interested in?")
        topic = take_command().strip() if take_command else None
        if not topic:
            if talk:
                talk("You didn't provide a topic. Please try again.")
            return None
        api_key = "74ba339c1205483cbf2d4205c51cd690"
        url = f"https://newsapi.org/v2/everything?q={topic}&language=en&pageSize=5&apiKey={api_key}"
        response = requests.get(url).json()
        if response.get("status") == "error":
            if talk:
                talk(f"Error fetching news: {response.get('message')}")
            return None
        articles = response.get("articles", [])
        if articles and talk:
            talk(f"Here are the top news articles on {topic}:")
            for i, article in enumerate(articles[:5], 1):
                title = article.get("title", "No title available")
                description = article.get("description", "No description available")
                talk(f"News {i}: {title}. {description}")
        return articles
    except Exception as e:
        if talk:
            talk("There was an error fetching the news.")
        return None

def get_device_information(talk=None):
    try:
        system_info = {
            "System": platform.system(),
            "Node Name": platform.node(),
            "Release": platform.release(),
            "Version": platform.version(),
            "Machine": platform.machine(),
            "Processor": platform.processor(),
            "RAM": f"{round(psutil.virtual_memory().total / (1024 ** 3), 2)} GB"
        }
        if talk:
            talk("Here is your device information:")
            for key, value in system_info.items():
                talk(f"{key}: {value}")
        return system_info
    except Exception as e:
        if talk:
            talk("An error occurred while retrieving device information.")
        return None

def check_battery_status(talk=None):
    battery = psutil.sensors_battery()
    if battery is None:
        if talk:
            talk("Sorry, I couldn't fetch battery details.")
        return None
    percent = battery.percent
    charging = battery.power_plugged
    status = "charging" if charging else "not charging"
    if talk:
        talk(f"The battery is at {percent} percent and it is currently {status}.")
    return {"percent": percent, "charging": charging}

def refresh_windows(talk=None):
    try:
        pyautogui.press('f5')
        if talk:
            talk("The desktop has been refreshed.")
        return True
    except Exception as e:
        if talk:
            talk("An error occurred while refreshing the desktop.")
        return False

# --- Advanced helpers ---
APP_DIR = Path.cwd()
PAGES_DIR = APP_DIR / "copilot_pages"
PAGES_DIR.mkdir(exist_ok=True)
SETTINGS_FILE = APP_DIR / "copilot_settings.json"
WATCHLIST_FILE = APP_DIR / "copilot_watchlist.json"

def save_settings(settings: dict):
    try:
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(settings, f, indent=2)
        return True
    except Exception:
        logging.exception('Failed to save settings')
        return False

def load_settings():
    if not SETTINGS_FILE.exists():
        return {}
    try:
        with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        logging.exception('Failed to load settings')
        return {}

def search_files(query: str, start_dirs=None, extensions=None, max_results=50):
    """Search filenames under provided start_dirs (defaults to user's home)."""
    if extensions is None:
        extensions = ['.txt', '.docx', '.pdf', '.xlsx', '.pptx', '.jpg', '.png']
    results = []
    if start_dirs is None:
        start_dirs = [str(Path.home())]
    q = query.lower()
    for start in start_dirs:
        for root, dirs, files in os.walk(start):
            for f in files:
                if len(results) >= max_results:
                    return results
                if q in f.lower() and any(f.lower().endswith(ext) for ext in extensions):
                    results.append(os.path.join(root, f))
    return results

def set_timer(seconds: int, message: str, talk=None):
    def _notify():
        if talk:
            talk(f"Timer finished: {message}")
    t = threading.Timer(seconds, _notify)
    t.daemon = True
    t.start()
    return t

def create_page(title: str, content: str):
    filename = PAGES_DIR / f"{title}.md"
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(content)
    return str(filename)

def list_pages():
    return [p.name for p in PAGES_DIR.glob('*.md')]

def edit_page(title: str, new_content: str):
    filename = PAGES_DIR / f"{title}.md"
    if not filename.exists():
        return False
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(new_content)
    return True

def generate_chart(values, labels=None, filename='chart.png'):
    try:
        fig, ax = plt.subplots()
        if labels:
            ax.pie(values, labels=labels, autopct='%1.1f%%')
        else:
            ax.plot(values)
        out = APP_DIR / filename
        fig.savefig(out, bbox_inches='tight')
        plt.close(fig)
        return str(out)
    except Exception:
        logging.exception('Failed to generate chart')
        return None

def add_watchlist_item(item: dict):
    lst = []
    if WATCHLIST_FILE.exists():
        try:
            with open(WATCHLIST_FILE, 'r', encoding='utf-8') as f:
                lst = json.load(f)
        except Exception:
            lst = []
    lst.append(item)
    with open(WATCHLIST_FILE, 'w', encoding='utf-8') as f:
        json.dump(lst, f, indent=2)
    return True

def get_watchlist():
    if not WATCHLIST_FILE.exists():
        return []
    try:
        with open(WATCHLIST_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        logging.exception('Failed to read watchlist')
        return []

def check_watchlist(talk=None):
    items = get_watchlist()
    alerts = []
    for item in items:
        try:
            name = item.get('name') or item.get('url')
            url = item.get('url')
            target = float(item.get('target'))
            price = None
            if url and 'amazon' in url:
                price = get_price_from_amazon(name, 'india', talk)
            if price and price <= target:
                msg = f"Price alert: {name} is now {price} (target {target})"
                alerts.append(msg)
                if talk:
                    talk(msg)
        except Exception:
            continue
    return alerts

def launch_app(app_name: str):
    mapping = {
        'notepad': 'notepad.exe',
        'vscode': r"C:\\Users\\%USERNAME%\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe",
        'chrome': 'chrome.exe',
        'word': 'winword.exe',
        'excel': 'excel.exe',
        'powerpoint': 'powerpnt.exe'
    }
    exe = mapping.get(app_name.lower(), app_name)
    try:
        try:
            os.startfile(exe)
        except Exception:
            subprocess.Popen(shlex.split(exe))
        return True
    except Exception:
        logging.exception('Failed to launch app')
        return False

def send_sms_via_twilio(number: str, msg: str):
    try:
        from twilio.rest import Client
    except Exception:
        return False, 'twilio library not installed'
    sid = os.environ.get('TWILIO_SID')
    token = os.environ.get('TWILIO_AUTH')
    from_num = os.environ.get('TWILIO_FROM')
    if not (sid and token and from_num):
        return False, 'Twilio credentials not configured in environment variables'
    client = Client(sid, token)
    try:
        client.messages.create(body=msg, from_=from_num, to=number)
        return True, 'Sent'
    except Exception as e:
        logging.exception(e)
        return False, str(e)
