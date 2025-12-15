# ===================== assistant_ui.py =====================
import tkinter as tk
import threading
import queue
from assistant_core import listen, execute_command, speak

command_queue = queue.Queue()
listening = True

ui = tk.Tk()
ui.title("Voice Assistant")
ui.geometry("500x700")
ui.configure(bg="#0f172a")

# -------- Chat container --------
canvas = tk.Canvas(ui, bg="#0f172a", highlightthickness=0)
scroll = tk.Scrollbar(ui, command=canvas.yview)
canvas.configure(yscrollcommand=scroll.set)
scroll.pack(side="right", fill="y")
canvas.pack(side="left", fill="both", expand=True)

messages = tk.Frame(canvas, bg="#0f172a")
canvas.create_window((0, 0), window=messages, anchor="nw")


def add_bubble(text, sender):
    bg = "#2563eb" if sender == "assistant" else "#1e293b"
    anchor = "w" if sender == "assistant" else "e"

    frame = tk.Frame(messages, bg="#0f172a")
    label = tk.Label(
    frame,
    text=text,
    bg=bg,
    fg="white",
    wraplength=460,
    padx=14,
    pady=10,
    font=("Segoe UI", 10),
    relief="flat",
    justify="left"
    )
    frame.pack(fill="x", padx=20, pady=6)

    label.pack(
        anchor=anchor,
        padx=10
    )
    ui.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))
    canvas.yview_moveto(1)

def on_resize(event):
    canvas.config(scrollregion=canvas.bbox("all"))

ui.bind("<Configure>", on_resize)

status = tk.Label(ui, text="🎧 Listening", fg="#22c55e", bg="#0f172a")
status.pack(pady=5)


def voice_listener():
    while True:
        if listening:
            cmd = listen()
            if cmd:
                command_queue.put(cmd)


def process_commands():
    try:
        while not command_queue.empty():
            cmd = command_queue.get()
            add_bubble(cmd, "user")
            response = execute_command(cmd)
            add_bubble(response, "assistant")
    finally:
        ui.after(150, process_commands)


add_bubble("Hello Boss, how can I assist you now?", "assistant")
speak("Hello Boss, how can I assist you now?")

threading.Thread(target=voice_listener, daemon=True).start()
process_commands()
ui.mainloop()
