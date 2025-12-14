import os 
import webbrowser
import speech_recognition as sr
import pywhatkit
import wikipedia
import pyautogui
import datetime
import requests
import re
import urllib.parse
import time
import pywinauto
import pygetwindow as gw
from pywinauto import Application,Desktop
from pywinauto.findwindows import find_element
import platform
import uiautomation as auto
import psutil
import traceback
from pywinauto.findwindows import ElementNotFoundError
import math
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from urllib.parse import quote
import uiautomation as automation
import pytesseract
import win32con
import win32gui
import pyttsx3
import win32ui
from google_trans_new import google_translator
from pytesseract import image_to_string
import pytesseract
from bs4 import BeautifulSoup
import secondry as sec
from googletrans import Translator
import tkinter as tk
from tkinter import scrolledtext,Label, Entry, Frame,Button,Canvas,Scrollbar,font,PhotoImage, Checkbutton, IntVar, Toplevel, messagebox
import threading
from PIL import Image, ImageTk,ImageFont,ImageGrab
import itertools
import subprocess
import platform
import cv2
import phonenumbers
import numpy as np
import pyaudio
import json

recording = False  # Global flag to track recording status
output_file = None  # Global variable to store filename

# Set the path for Tesseract-OCR (update this path if installed elsewhere)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
# GUI Setup
root = tk.Tk()
root.title("Heimdall - Voice Assistant")
logo = tk.PhotoImage(file=r"C:\Users\manoj\OneDrive\Desktop\Final yr project\Heimdall\Capture.PNG")
root.iconphoto(True, logo)
root.geometry("1000x800+200+10")
root.resizable(True, True)

# Permissions dict: set to False by default; UI allows enabling
permissions = {
    'microphone': False,
    'file_search': False,
    'file_access': False,
    'screen_record': False,
    'read_window_text': False,
    'send_messages': False,
    'voice_activation': False,
    'device_info': False,
    'network_access': True,  # Allow network by default for functionality
}

# UI setting: whether to automatically open Explorer and select first search result
auto_open_on_search = False

# Control for listening thread
assistant_running = False
assistant_thread = None
engine = pyttsx3.init()
engine.setProperty('rate', 150)
engine.setProperty('volume', 1.0)

def show_permissions_modal():
    """Show a modal for initial permissions configuration."""
    modal = Toplevel(root)
    modal.title("Permissions")
    modal.geometry("420x300+400+200")
    modal.resizable(False, False)
    modal.grab_set()

    vars_map = {}
    row = 0
    for key, default in permissions.items():
        var = IntVar(value=1 if default else 0)
        chk = Checkbutton(modal, text=key.replace('_', ' ').title(), variable=var)
        chk.grid(row=row, column=0, sticky='w', padx=10, pady=6)
        vars_map[key] = var
        row += 1

    def apply_and_close():
        for k, v in vars_map.items():
            permissions[k] = True if v.get() == 1 else False
        # persist settings
        try:
            sec.save_settings({'permissions': permissions})
        except Exception:
            pass
        modal.destroy()

    apply_btn = Button(modal, text='Apply', command=apply_and_close)
    apply_btn.grid(row=row, column=0, pady=10)
    modal.focus_force()

def ensure_permission(key: str, description: str) -> bool:
    """Ensure a permission is granted; ask the user if not."""
    if permissions.get(key):
        return True
    answer = messagebox.askyesno("Permission Request", f"{description}\nGrant '{key.replace('_', ' ').title()}' permission?")
    if answer:
        permissions[key] = True
        return True
    return False

# UI Colors (simple all-white theme with green chat bubbles)
light_bg = "white"
root.configure(bg=light_bg)
custom_font = ("Delius", 14, "bold")

# Chat colors (assistant and user both in green hues)
assistant_color = "#d4f7d4"
user_color = "#b8f0b8"
assistant_text_color = "#013220"
user_text_color = "#013220"

def create_chat_bubble(parent, text, bg_color, fg_color, align_right=False):
    """Creates a simple rectangular message bubble with correct padding."""
    text_label = tk.Label(
        parent,
        text=text,
        font=custom_font,
        fg=fg_color,
        bg=bg_color,
        padx=8, pady=8,  # Ensure text padding is consistent
        wraplength=600,  # Ensures text wraps properly within a max width
        # relief="flat",  # No border
        bd=0
    )

    # Correct alignment: right for user, left for assistant
    text_label.pack(anchor="e" if align_right else "w", padx=5, pady=5)

    return text_label

def add_chat_message(text, sender):
    """Display chat messages with correct padding based on its own text."""
    # Choose green-themed colors
    if sender == "assistant":
        bg_color = assistant_color
        fg_color = assistant_text_color
        is_user = False
    else:
        bg_color = user_color
        fg_color = user_text_color
        is_user = True

    bubble = create_chat_bubble(scrollable_frame, text, bg_color, fg_color, align_right=is_user)

    # Make file paths clickable (opens Explorer and selects the file)
    if isinstance(text, str) and os.path.exists(text) and sender == 'assistant':
        def _open(e=None, p=text):
            try:
                subprocess.Popen(['explorer', '/select,', p])
            except Exception:
                try:
                    os.startfile(os.path.dirname(p))
                except Exception:
                    pass
        bubble.bind('<Button-1>', _open)

    # Ensures proper spacing between user & assistant messages
    bubble.pack(fill="x", padx=(400, 10) if is_user else (10, 400), pady=6)

    # Scroll to the latest message
    canvas.update_idletasks()
    canvas.yview_moveto(1.0)

# No animated GIF; keep interface clean and static white

# Toggle Button

# Toggle button removed for simplified white theme


# Assistant control buttons and permission indicator
control_frame = Frame(root, bg=light_bg)
control_frame.pack(pady=4)

start_btn = Button(control_frame, text="Start Assistant", font=custom_font, padx=8, pady=4)
stop_btn = Button(control_frame, text="Stop Assistant", font=custom_font, padx=8, pady=4)
perm_btn = Button(control_frame, text="Permissions", font=custom_font, padx=8, pady=4)
settings_btn = Button(control_frame, text="Settings", font=custom_font, padx=8, pady=4)
voice_var = IntVar(value=1 if permissions.get('voice_activation') else 0)
voice_chk = Checkbutton(control_frame, text='Voice Activation (Hey Copilot)', variable=voice_var)
status_label = Label(control_frame, text="Status: Idle", bg=light_bg, fg=assistant_text_color, font=("Delius", 11, "bold"))
start_btn.grid(row=0, column=0, padx=6)
stop_btn.grid(row=0, column=1, padx=6)
perm_btn.grid(row=0, column=2, padx=6)
voice_chk.grid(row=0, column=4, padx=6)
settings_btn.grid(row=0, column=5, padx=6)
status_label.grid(row=0, column=3, padx=6)

def update_status(text: str):
    status_label.config(text=f"Status: {text}")

def start_assistant_button_cb():
    global assistant_running, assistant_thread
    if assistant_running:
        talk("Assistant already running.")
        return
    if not ensure_permission('microphone', 'Microphone access is required for voice commands'):
        talk("Microphone permission denied. Cannot start assistant.")
        return
    assistant_running = True
    update_status('Listening')
    threading.Thread(target=run_assistant, daemon=True).start()
    # If voice activation is enabled start the wakeword listener too
    if permissions.get('voice_activation'):
        start_wakeword_listener()

def stop_assistant_button_cb():
    global assistant_running
    if not assistant_running:
        talk("Assistant is not running.")
        return
    assistant_running = False
    update_status('Idle')
    talk("Assistant stopped.")
    stop_wakeword_listener()

start_btn.config(command=start_assistant_button_cb)
stop_btn.config(command=stop_assistant_button_cb)
perm_btn.config(command=show_permissions_modal)
settings_btn.config(command=lambda: show_settings_modal())
voice_chk.config(command=lambda: permissions.update({'voice_activation': bool(voice_var.get())}))

# Quick file search box
file_search_var = tk.StringVar()
file_search_entry = Entry(control_frame, textvariable=file_search_var, width=24)
file_search_btn = Button(control_frame, text='Search Files', font=custom_font, command=lambda: start_file_search())
file_search_entry.grid(row=1, column=0, columnspan=2, pady=6)
file_search_btn.grid(row=1, column=2, pady=6)

def start_file_search():
    q = file_search_var.get().strip()
    if not q:
        talk('Please enter a search query in the file search box.')
        return
    if not ensure_permission('file_search', 'Allow searching files on your device'):
        talk('Permission denied for file search.')
        return
    results = sec.search_files(q)
    if results:
        talk(f'Found {len(results)} results. Showing top 5.')
        for i, r in enumerate(results[:5]):
            add_chat_message(r, 'assistant')
            talk(r)
        # Optionally open the location of the first result in Explorer
        if auto_open_on_search:
            try:
                subprocess.Popen(['explorer', '/select,', results[0]])
            except Exception:
                try:
                    os.startfile(os.path.dirname(results[0]))
                except Exception:
                    pass
    else:
        talk('No files found.')

# Wakeword listener controls
wakeword_thread = None
wakeword_running = False

def wakeword_listener():
    global wakeword_running, assistant_running
    r = sr.Recognizer()
    mic = None
    try:
        mic = sr.Microphone()
    except Exception:
        return
    with mic as source:
        r.adjust_for_ambient_noise(source, duration=0.5)
    while wakeword_running:
        try:
            with sr.Microphone() as source:
                audio = r.listen(source, timeout=3, phrase_time_limit=4)
            text = r.recognize_google(audio, language='en-US').lower()
            if 'hey copilot' in text or 'hey, copilot' in text:
                if not assistant_running:
                    talk('Yes Boss, how can I help you?')
                    assistant_running = True
                    update_status('Listening')
                    threading.Thread(target=run_assistant, daemon=True).start()
        except sr.WaitTimeoutError:
            continue
        except sr.UnknownValueError:
            continue
        except Exception:
            continue

def start_wakeword_listener():
    global wakeword_thread, wakeword_running
    if wakeword_running:
        return
    if not ensure_permission('microphone', 'Microphone required for wake word detection'):
        return
    wakeword_running = True
    wakeword_thread = threading.Thread(target=wakeword_listener, daemon=True)
    wakeword_thread.start()

def stop_wakeword_listener():
    global wakeword_running
    wakeword_running = False

def show_settings_modal():
    modal = Toplevel(root)
    modal.title('Settings')
    modal.geometry('420x200+420+220')
    modal.grab_set()
    # Voice selection
    available_voices = engine.getProperty('voices')
    voice_names = [v.name for v in available_voices]
    current_voice_id = engine.getProperty('voice')
    current_voice_name = next((v.name for v in available_voices if v.id == current_voice_id), voice_names[0] if voice_names else '')
    sel_var = tk.StringVar(value=current_voice_name)
    tk.Label(modal, text='Select Voice:').pack(anchor='w', padx=8, pady=6)
    voice_menu = tk.OptionMenu(modal, sel_var, *voice_names)
    voice_menu.pack(anchor='w', padx=8)

    # Auto-open on search option
    auto_open_var = IntVar(value=1 if auto_open_on_search else 0)
    auto_chk = Checkbutton(modal, text='Auto-open first file search result', variable=auto_open_var)
    auto_chk.pack(anchor='w', padx=8, pady=6)

    def apply_and_close():
        selected = sel_var.get()
        for v in available_voices:
            if v.name == selected:
                engine.setProperty('voice', v.id)
                break
        # persist: merge with existing settings
        try:
            saved = sec.load_settings() or {}
            saved.update({'permissions': permissions, 'voice': selected, 'auto_open_on_search': bool(auto_open_var.get())})
            sec.save_settings(saved)
            # update runtime flag
            global auto_open_on_search
            auto_open_on_search = bool(auto_open_var.get())
        except Exception:
            pass
        modal.destroy()

    tk.Button(modal, text='Apply', command=apply_and_close).pack(pady=10)

# Load saved settings if available
try:
    _saved = sec.load_settings()
    if isinstance(_saved, dict):
        if 'permissions' in _saved:
            permissions.update(_saved['permissions'])
            voice_var.set(1 if permissions.get('voice_activation') else 0)
            if permissions.get('voice_activation'):
                start_wakeword_listener()
        if 'voice' in _saved:
            # try to apply saved voice by name
            for v in engine.getProperty('voices'):
                if v.name == _saved['voice']:
                    engine.setProperty('voice', v.id)
                    break
        if 'auto_open_on_search' in _saved:
            auto_open_on_search = bool(_saved.get('auto_open_on_search'))
except Exception:
    pass

# Scrollable Chat Area
chat_frame = Frame(root, bg=light_bg)
chat_frame.pack(fill="both", expand=True, padx=10, pady=5)

canvas = Canvas(chat_frame, bg=light_bg, highlightthickness=0)
scrollbar = Scrollbar(chat_frame, orient="vertical", command=canvas.yview)
canvas.configure(yscrollcommand=scrollbar.set)

scrollable_frame = Frame(canvas, bg=light_bg)
scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# Initialize TTS Engine
engine = pyttsx3.init()
engine.setProperty('rate', 150)
engine.setProperty('volume', 1.0)

def talk(text):
    """Talk the provided text and display it in chat"""
    add_chat_message(text, "assistant")
    root.update()
    engine.say(text)
    engine.runAndWait()

def take_command():
    recognizer = sr.Recognizer()
    recognizer.energy_threshold = 50  # Increased sensitivity
    recognizer.dynamic_energy_threshold = True
    recognizer.pause_threshold = 0.5  # 🔹 Must be >= non_talking_duration
    recognizer.non_talking_duration = 0.2  # 🔹 Keep this lower

    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source, duration=0.5)

        print("Listening...")
        while True:
            try:
                audio = recognizer.listen(source, timeout=None, phrase_time_limit=60)
                print("Processing...")

                command = recognizer.recognize_google(audio, language='en-US')
                command = command.lower()

                add_chat_message(command, "user")
                return command  
            
            except sr.UnknownValueError:
                print("Didn't catch that. Listening...")
                continue  
            
            except sr.RequestError:
                talk("There was an issue with the speech recognition service. Check your internet connection.")
                return "error"  
    
# Amazon utilities and exchange rates are moved to secondry.py

# Volume and Wi-Fi utilities are moved to secondry.py

def get_split_phone_number():
    while True:
        talk("Please say the first 5 digits of the phone number.")
        first_part = take_command()
        first_digits = ''.join(filter(str.isdigit, first_part))

        talk("Now say the last 5 digits of the phone number.")
        second_part = take_command()
        second_digits = ''.join(filter(str.isdigit, second_part))

        combined = first_digits + second_digits

        if len(combined) == 10:
            full_number = f"+91{combined}"
            talk(f"You said {full_number}. Is that correct? Say yes or no.")
            confirm = take_command().lower()

            if "yes" in confirm:
                return full_number
            else:
                talk("Okay, let's try again.")
        else:
            talk("That doesn't seem like a valid 10-digit number. Let's try again.")

def send_whatsapp_message():
    if not ensure_permission('send_messages', 'Allow sending messages on your behalf via WhatsApp?'):
        talk('Permission denied to send messages.')
        return
    phone_number = get_split_phone_number()

    talk("What message should I send?")
    message = take_command()

    now = datetime.datetime.now() + datetime.timedelta(minutes=1)
    hour = now.hour
    minute = now.minute

    talk(f"Sending your message to {phone_number}")
    pywhatkit.sendwhatmsg(phone_number, message, hour, minute, wait_time=10, tab_close=True)

    talk("Your message has been scheduled.")

# Read active window content
def read_active_window_content():
    if not ensure_permission('read_window_text', 'Allow the assistant to read the visible content of the active window'):
        talk('Permission denied for reading active window content.')
        return
    try:
        hwnd = win32gui.GetForegroundWindow()
        window_title = win32gui.GetWindowText(hwnd)
        talk(f"The active window is titled: {window_title}")

        rect = win32gui.GetWindowRect(hwnd)
        x, y, x1, y1 = rect

        # Capture the window area
        window_image = ImageGrab.grab(bbox=(x, y, x1, y1))

        # Use OCR to extract text
        extracted_text = pytesseract.image_to_string(window_image)

        if extracted_text.strip():
            talk("The content of the active window is as follows:")
            talk(extracted_text[:1000])  # Read only the first 1000 characters
        else:
            talk("The active window has no readable content.")
    except Exception as e:
        talk("An error occurred while reading the active window's content.")
        print(f"Error: {e}")

# Function to get device information
# get_device_information moved to secondry.get_device_information(talk)

recording = False  # Global flag to control recording
output_file = None  # File name for saving the recording

def record_screen(file_name):
    """Records the screen and saves it to a file."""
    global recording

    recording = True  # Start recording
    screen_size = pyautogui.size()  # Get screen resolution

    # Use MJPG codec for better compatibility
    fourcc = cv2.VideoWriter_fourcc(*"MJPG")
    out = cv2.VideoWriter(file_name, fourcc, 20.0, screen_size)

    if not out.isOpened():
        talk("Error initializing video writer. Recording failed.")
        return

    talk(f"Recording started. File will be saved as {file_name}")

    try:
        while recording:
            img = pyautogui.screenshot()
            frame = np.array(img)
            frame = cv2.cvtColor(frame, cv2.COLOR_RGB2BGR)
            out.write(frame)  # Write frame to video
            time.sleep(1 / 20)  # 20 FPS

    except Exception as e:
        print(f"⚠️ Error during recording: {e}")
    
    out.release()  # Release video writer
    talk(f"Recording saved as {file_name}")

# Function to take a screenshot and save with user-provided name
def take_screenshot():
    try:
        if not ensure_permission('file_access', 'Allow saving screenshots to your local files'):
            talk('Permission denied for saving screenshots.')
            return
        # Ask the user for the file name to save the screenshot
        talk("What name would you like to save the screenshot as?")
        file_name = take_command().strip()
        
        if file_name:
            screenshot = pyautogui.screenshot()
            screenshot.save(f"{file_name}.png")  # Save screenshot with the user-provided name
            talk(f"Screenshot taken and saved as {file_name}.png.")
            print(f"Screenshot saved as {file_name}.png")
        else:
            talk("No file name provided. Please try again.")

    except Exception as e:
        talk("An error occurred while taking the screenshot.")
        print(f"Error taking screenshot: {e}")

# Function to translate text and talk the result
# Moved `translate_text` to secondry.translate_text(take_command, talk)
# Function to fetch weather information
# Moved `fetch_weather` to secondry.fetch_weather(take_command, talk)
# Function to search Wikipedia
# Moved `search_wikipedia` to secondry.search_wikipedia(command, talk)
            
def take_spelled_filename():
    talk("Please spell the file name, letter by letter.")
    filename = ""
    while True:
        letter = take_command()
        if "stop" in letter:  # Stop input when user says "stop"
            break
        elif len(letter) == 1:  # Ensure the input is a single letter
            filename += letter
            talk(f"Added letter {letter}.")
        else:
            talk("Please say a single letter.")
    return filename

def open_notepad():
    talk("Opening Notepad and starting to write. talk your content, and say 'stop writing' when you want to stop.")
    os.startfile("notepad.exe")  # Open Notepad
    time.sleep(1)  # Allow some time for Notepad to open
    content = ""
    start_time = time.time()
    while True:
        # Listen for content until the user says 'stop writing'
        command = take_command()
        if "stop writing" in command:
            break
        elif command:  # If the command is not empty, add to content
            content += command + "\n"
            pyautogui.write(command)  # Simulate typing in Notepad
        # Break after 20 seconds of no input or wait for the next command
        if time.time() - start_time > 20:
            talk("Listening for your next input...")
            start_time = time.time()  # Reset the timer
        time.sleep(3)  # Add a break of 3 seconds after each listening
    # Ask for file name to save the content
    talk("What name would you like to save the file as? Please say the file name.")
    file_name = take_command().strip()
    # If the recognition of the filename fails or is empty, provide alternative options
    if not file_name:
        talk("I couldn't recognize the filename. Would you like to spell it out instead? Say 'yes' to spell it or 'no' to enter it manually.")
        response = take_command().strip().lower()
        if "yes" in response:
            file_name = take_spelled_filename()
        elif "no" in response:
            talk("Please type the file name manually.")
            file_name = input("Enter file name manually: ").strip()
        else:
            talk("Invalid response. Please try again later.")
            return
    if file_name:
        if not ensure_permission('file_access', 'Allow the assistant to save files to your Documents folder'):
            talk('Permission denied. Not saving the file.')
            return
        # Fixed directory path where the file will be saved
        save_directory = "C:\\Users\\User\\Documents"
        save_path = os.path.join(save_directory, f"{file_name}.txt")
        try:
            # Save the content to the file
            with open(save_path, "w") as f:
                f.write(content)
            talk(f"Content saved as {file_name}.txt in your Documents folder.")
            print(f"Content saved as {file_name}.txt in {save_directory}")
        except Exception as e:
            talk(f"An error occurred while saving the file: {e}")
    else:
        talk("No file name provided. The content was not saved.")
def close_notepad():
    try:
        # Terminate all Notepad instances
        os.system("taskkill /im notepad.exe /f")
        talk("Notepad has been closed.")
        print("Notepad closed successfully.")
    except Exception as e:
        talk("An error occurred while closing Notepad.")
        print(f"Error closing Notepad: {e}")
def open_word():
    talk("Opening Microsoft Word. talk your content, and say 'stop writing' when you want to stop.")
    os.startfile("winword.exe")  # Open Word
    time.sleep(5)  # Allow some time for Word to open
    content = ""
    start_time = time.time()
    while True:
        command = take_command()
        if "stop writing" in command:
            break
        elif command:
            content += command + "\n"
            pyautogui.write(command)
        if time.time() - start_time > 20:
            talk("Listening for your next input...")
            start_time = time.time()
        time.sleep(3)
    talk("What name would you like to save the document as?")
    file_name = take_command().strip()
    if file_name:
        if not ensure_permission('file_access', 'Allow the assistant to save files on your device'):
            talk('Permission denied. Document not saved.')
            return
        pyautogui.hotkey("ctrl", "s")
        time.sleep(2)
        pyautogui.write(file_name)
        pyautogui.press("enter")
        talk(f"Document saved as {file_name}.")
    else:
        talk("No file name provided. The document was not saved.")
def close_word():
    try:
        os.system("taskkill /im winword.exe /f")
        talk("Microsoft Word has been closed.")
        print("Word closed successfully.")
    except Exception as e:
        talk("An error occurred while closing Word.")
        print(f"Error closing Word: {e}")
def open_excel():
    talk("Opening Microsoft Excel. talk your content, and say 'stop writing' when you want to stop.")
    os.startfile("excel.exe")  # Open Excel
    time.sleep(5)  # Allow some time for Excel to open
    content = ""
    start_time = time.time()
    while True:
        command = take_command()
        if "stop writing" in command:
            break
        elif command:
            pyautogui.write(command)
            pyautogui.press("tab")
        if time.time() - start_time > 20:
            talk("Listening for your next input...")
            start_time = time.time()
        time.sleep(3)
    talk("What name would you like to save the spreadsheet as?")
    file_name = take_command().strip()
    if file_name:
        if not ensure_permission('file_access', 'Allow the assistant to save files on your device'):
            talk('Permission denied. Spreadsheet not saved.')
            return
        pyautogui.hotkey("ctrl", "s")
        time.sleep(2)
        pyautogui.write(file_name)
        pyautogui.press("enter")
        talk(f"Spreadsheet saved as {file_name}.")
    else:
        talk("No file name provided. The spreadsheet was not saved.")
def close_excel():
    try:
        os.system("taskkill /im excel.exe /f")
        talk("Microsoft Excel has been closed.")
        print("Excel closed successfully.")
    except Exception as e:
        talk("An error occurred while closing Excel.")
        print(f"Error closing Excel: {e}")
def open_powerpoint():
    talk("Opening Microsoft PowerPoint. talk your content, and say 'stop writing' when you want to stop.")
    os.startfile("powerpnt.exe")  # Open PowerPoint
    time.sleep(5)  # Allow some time for PowerPoint to open
    content = ""
    start_time = time.time()
    while True:
        command = take_command()
        if "stop writing" in command:
            break
        elif command:
            pyautogui.write(command)
            pyautogui.press("enter")
        if time.time() - start_time > 20:
            talk("Listening for your next input...")
            start_time = time.time()
        time.sleep(3)
    talk("What name would you like to save the presentation as?")
    file_name = take_command().strip()
    if file_name:
        if not ensure_permission('file_access', 'Allow the assistant to save files on your device'):
            talk('Permission denied. Presentation not saved.')
            return
        pyautogui.hotkey("ctrl", "s")
        time.sleep(2)
        pyautogui.write(file_name)
        pyautogui.press("enter")
        talk(f"Presentation saved as {file_name}.")
    else:
        talk("No file name provided. The presentation was not saved.")
def close_powerpoint():
    try:
        os.system("taskkill /im powerpnt.exe /f")
        talk("Microsoft PowerPoint has been closed.")
        print("PowerPoint closed successfully.")
    except Exception as e:
        talk("An error occurred while closing PowerPoint.")
        print(f"Error closing PowerPoint: {e}")
def get_active_window():
    """Get the title of the currently active window."""
    try:
        active_window = gw.getActiveWindow()
        return active_window.title if active_window else "Unknown Window"
    except Exception as e:
        print("Error getting active window:", e)
        return "Unknown Window"
def switch_window():
    """Switch to the next open window using Alt + Tab."""
    talk("Switching window...")
    pyautogui.keyDown("alt")
    pyautogui.press("tab")
    pyautogui.keyUp("alt")

def switch_tab():
    """Switch between tabs based on the active window."""
    active_window = get_active_window().lower()

    if "chrome" in active_window:
        talk("Switching tab in Chrome")
        pyautogui.hotkey("ctrl", "tab")  # Next tab in Chrome

    elif "visual studio code" in active_window or "vs code" in active_window:
        talk("Switching tab in VS Code")
        pyautogui.hotkey("ctrl", "pageup")  # Next tab in VS Code

    else:
        talk("This application does not support tab switching.")

# fetch_news moved to secondry.fetch_news(take_command, talk)
# check_battery_status moved to secondry.check_battery_status(talk)

# refresh_windows moved to secondry.refresh_windows(talk)
# Function to perform actions based on commands
def run_assistant():
     global recording, output_file, assistant_running  # Needed to modify recording status and control loop
     while assistant_running:
        command = take_command()
        if not assistant_running:
            break
        if not command:  # Check if command is None or empty
             continue
            # root.after(100, run_assistant)
            # return  # Return to listening state if no command is detected
        if "refresh windows" in command or "refresh desktop" in command:
            sec.refresh_windows(talk)
        elif "charge percentage" in command or "battery status" in command:
            sec.check_battery_status(talk)
        elif "time" in command:
            current_time = datetime.datetime.now().strftime('%I:%M %p')
            talk(f"The current time is {current_time}.")
        elif "search" in command:
            # file search or web search
            if command.startswith('search files for') or command.startswith('find file'):
                if not ensure_permission('file_search', 'Allow searching files on your device'):
                    talk('File search permission denied.')
                    continue
                q = command.replace('search files for', '').replace('find file', '').strip()
                results = sec.search_files(q)
                if results:
                    talk(f'I found {len(results)} files. Showing top results.')
                    for i, r in enumerate(results[:5]):
                        add_chat_message(r, 'assistant')
                        talk(r)
                    # Optionally open Explorer to the first match
                    if auto_open_on_search:
                        try:
                            subprocess.Popen(['explorer', '/select,', results[0]])
                        except Exception:
                            try:
                                os.startfile(os.path.dirname(results[0]))
                            except Exception:
                                pass
                else:
                    talk('No files found matching that query.')
                continue
            if not ensure_permission('network_access', 'Allow internet access for web searches'):
                talk('Network permission denied.')
                continue
            query = command.replace("search", "").strip()
            webbrowser.open(f"https://www.google.com/search?q={query}")
            talk(f"Searching for {query}.")
        elif "who are you" in command:
            talk("I am Heimdall, your personal assistant. How can I help you?")
        elif "how are you" in command:
            talk("I am fine and hope you are too, boss. How can I help you?")
        elif "tell me a joke" in command:
            talk("Why was the math book sad? Because it had too many problems.")
        elif "play" in command:
            song = command.replace("play", "").strip()
            talk(f"Playing {song} on YouTube.")
            pywhatkit.playonyt(song)
        elif "Wi-Fi" in command or "wi-fi" in command:
            if not ensure_permission('device_info', 'Allow reading network interface details'):
                talk('Permission denied for Wi-Fi access.')
                continue
            sec.check_wifi(talk)
        elif "price of" in command:
            parts = command.split(" in ")
            if len(parts) == 2:
                item_name = parts[0].replace("price of", "").strip()
                location = parts[1].strip()
                if not ensure_permission('network_access', 'Allow internet access to fetch price information'):
                    talk('Network permission denied. Unable to fetch price.')
                    continue
                price = sec.get_price_from_amazon(item_name, location, talk)
                if price:
                    talk(f"The price of {item_name} in {location} is {price} rupees.")
                else:
                    talk("Sorry, I couldn't find the price for that item.")
            else:
                talk("Please specify the location for the price.")
        elif "get weather" in command:
            if not ensure_permission('network_access', 'Allow internet access to fetch weather'):
                talk('Network permission denied. Unable to fetch weather.')
                continue
            sec.fetch_weather(take_command, talk)
        elif "set timer" in command or "set a timer" in command:
            # e.g., 'set a timer for 10 minutes'
            m = re.search(r"(\d+)\s*(second|seconds|minute|minutes|hour|hours)", command)
            if m:
                num = int(m.group(1))
                unit = m.group(2)
                seconds = num * 60 if 'minute' in unit else num * 3600 if 'hour' in unit else num
                talk(f'Setting timer for {num} {unit}.')
                sec.set_timer(seconds, 'Your timer is done', talk)
            else:
                talk('For how long should I set the timer?')
                resp = take_command()
                m = re.search(r"(\d+)", resp)
                if m:
                    seconds = int(m.group(1))
                    talk(f'Setting timer for {seconds} seconds.')
                    sec.set_timer(seconds, 'Your timer is done', talk)
                else:
                    talk('I could not understand the duration.')
        elif "create page" in command or "new page" in command:
            talk('What is the title of the page?')
            title = take_command().strip()
            talk('Please say the content of the page now.')
            content = take_command().strip()
            path = sec.create_page(title, content)
            talk(f'Page {title} created at {path}')
        elif "list pages" in command or "show pages" in command:
            pages = sec.list_pages()
            if pages:
                talk(f'I have {len(pages)} pages:')
                for p in pages:
                    talk(p)
            else:
                talk('No pages found.')
        elif "generate chart" in command:
            talk('Please speak comma separated numbers for the chart.')
            nums = take_command()
            try:
                values = [float(x.strip()) for x in re.split('[, ]+', nums) if x.strip()]
                outfile = sec.generate_chart(values, filename='chart.png')
                if outfile:
                    talk('Chart generated.')
                    try:
                        os.startfile(outfile)
                    except Exception:
                        pass
                else:
                    talk('Failed to generate chart.')
            except Exception:
                talk('I could not parse the numbers.')
        elif "add watchlist" in command or "track price" in command:
            talk('Please give the product URL or name.')
            url = take_command().strip()
            talk('What is your target price?')
            target = take_command().strip()
            try:
                sec.add_watchlist_item({'url': url, 'target': target, 'name': url})
                talk('Added to watchlist.')
            except Exception:
                talk('Failed to add to watchlist.')
        elif "check watchlist" in command:
            alerts = sec.check_watchlist(talk)
            if not alerts:
                talk('No price alerts right now.')
        elif "device information" in command:
            sec.get_device_information(talk)
        elif "set volume to" in command:
            sec.process_volume_command(take_command, talk)
        elif "record screen" in command and not recording:
            if not ensure_permission('screen_record', 'Allow the assistant to record your screen'):
                talk('Permission denied to start recording.')
                continue
            talk("Say the file name to save the recording.")
            file_name = take_command().replace(" ", "_") + ".avi"
            output_file = file_name if file_name.strip() else "screen_recording.avi"
            # Start screen recording in a new thread
            threading.Thread(target=record_screen, args=(output_file,), daemon=True).start()
        elif "stop" in command and recording:
            recording = False  # Stop recording
            talk("Recording stopped.")
        elif "open google" in command or "open Google" in command or "Open Google" in command:
            if not ensure_permission('network_access', 'Allow internet access for web searches'):
                talk('Network permission denied.')
                continue
            webbrowser.open("https://www.google.com")
            talk("Opening Google.")
        elif command.startswith('open') or command.startswith('launch'):
            # e.g., 'open vscode' or 'launch notepad'
            parts = command.split(maxsplit=1)
            if len(parts) == 2:
                app = parts[1]
                ok = sec.launch_app(app)
                talk(f'Launching {app}' if ok else f'Unable to launch {app}')
        elif "open w3schools" in command or "open W3Schools" in command or "Open W3Schools" in command:
            if not ensure_permission('network_access', 'Allow internet access for web searches'):
                talk('Network permission denied.')
                continue
            webbrowser.open("https://www.w3schools.com/")
            talk("Opening W3Schools.")
        elif "wikipedia" in command:
            sec.search_wikipedia(command, talk)
        elif "translate" in command:
            sec.translate_text(take_command, talk)
        elif "take a screenshot" in command:
            take_screenshot()
        elif "where is" in command:
            location = command.replace("where is", "").strip()
            webbrowser.open(f"https://www.google.com/maps/place/{location}")
            talk(f"Here is the location of {location}.")
        elif "send message" in command or "send whatsapp message" in command:
            # send via whatsapp or SMS
            if 'sms' in command or 'send sms' in command:
                if not ensure_permission('send_messages', 'Allow sending SMS via configured service'):
                    talk('Permission denied to send messages.')
                    continue
                talk('Please say the number to send SMS to, including country code.')
                number = take_command().strip()
                talk('What is the message?')
                message = take_command().strip()
                ok, info = sec.send_sms_via_twilio(number, message)
                talk(info if not ok else 'SMS scheduled/sent.')
            else:
                send_whatsapp_message()
        elif "read window content" in command:
            read_active_window_content()
        elif "move window" in command or "move" in command:
            switch_window()
        elif "move tab" in command:
            switch_tab()
        # elif "open vscode" in command or "open Vscode" in command
        elif "open notepad" in command:
            open_notepad()
        elif "close notepad" in command:
            close_notepad()
        elif "open word" in command:
            open_word()
        elif "close word" in command:
            close_word()
        elif "open excel" in command:
            open_excel()
        elif "close excel" in command:
            close_excel()
        elif "open powerpoint" in command:
            open_powerpoint()
        elif "close powerpoint" in command:
            close_powerpoint()
        elif "news" in command:
            if not ensure_permission('network_access', 'Allow internet access to fetch news'):
                talk('Network permission denied. Unable to fetch news.')
                continue
            sec.fetch_news(take_command, talk)
        elif "goodbye" in command or "exit" in command:
            talk("Good bye Boss, See you later!")
            assistant_running = False
            root.after(1000, root.destroy)  # Close GUI after message
            break
        else:
            talk("I didn't catch that. Please try again.")
       #root.after(100, run_assistant)  # Schedule next listening

def start_assistant():
    """Start assistant in a separate thread"""
    global assistant_running
    if assistant_running:
        talk("Assistant already running.")
        return
    if not ensure_permission('microphone', 'Microphone access is required for voice commands'):
        talk("Microphone permission denied. Cannot start assistant.")
        return
    talk("Hello Boss, how can I assist you now?")
    assistant_running = True
    update_status('Listening')
    threading.Thread(target=run_assistant, daemon=True).start()

# Start Assistant After GUI Loads
root.after(1000, start_assistant)

# Start Main Event Loop
root.mainloop()