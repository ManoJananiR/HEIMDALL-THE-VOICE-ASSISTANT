import tkinter as tk
from tkinter import font as tkfont
from PIL import Image, ImageTk
import threading
import time

from assistant_core import execute_command, listen, speak


class HeimdallUI(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("HEIMDALL Assistant")
        self.state("zoomed")
        self.configure(bg="#242022")

        # ---------- mode state ----------
        self.dark_mode = True
        self.chat_bg_dark = "#242022"
        self.chat_bg_light = "#f8d5ed"
        self.bubble_dark = "#464548"
        self.bubble_light = "#ce90bc"

        # ---------- fonts ----------
        self.chat_font = tkfont.Font(family="Merriweather 18pt", size=14)
        self.title_font = tkfont.Font(family="Merriweather SemiBold 18pt", size=16)

        # ---------- top background image ----------
        bg_img = Image.open("assets/bg.png")
        self.bg_original = bg_img
        self.bg_photo = ImageTk.PhotoImage(bg_img)
        self.bg_label = tk.Label(self, image=self.bg_photo)
        self.bg_label.place(x=0, y=0, relwidth=1, relheight=1)
        self.bg_label.lower()
        self.bind("<Configure>", self._resize_bg)

        # ---------- main frame ----------
        self.main_frame = tk.Frame(self, bg="#242022")
        self.main_frame.place(relx=0.5, rely=0.5, anchor="center",
                              relwidth=1.0, relheight=1.0)

        # ---------- title bar ----------
        self.title_bar = tk.Frame(self.main_frame, bg="#81689d", height=40)
        self.title_bar.pack(fill="x")

        logo_img = Image.open("assets/logo.png").resize((32, 32), Image.LANCZOS)
        self.logo_photo = ImageTk.PhotoImage(logo_img)
        tk.Label(self.title_bar, image=self.logo_photo, bg="#81689d").pack(
            side="left", padx=10
        )

        tk.Label(
            self.title_bar, text="HEIMDALL",
            fg="white", bg="#81689d", font=self.title_font
        ).pack(side="left", padx=10)

        self.mode_btn = tk.Button(
            self.title_bar, text="🌙", command=self._toggle_mode,
            bg="#81689d", fg="white", relief="flat",
            font=self.chat_font, padx=10, pady=2
        )
        self.mode_btn.pack(side="right", padx=10)

        # ---------- chat area ----------
        chat_container = tk.Frame(self.main_frame, bg=self.chat_bg_dark)
        chat_container.pack(fill="both", expand=True)

        self.chat_canvas = tk.Canvas(chat_container, bg=self.chat_bg_dark,
                                     highlightthickness=0)
        self.chat_canvas.pack(side="left", fill="both", expand=True)

        scrollbar = tk.Scrollbar(chat_container, command=self.chat_canvas.yview)
        scrollbar.pack(side="right", fill="y")
        self.chat_canvas.configure(yscrollcommand=scrollbar.set)

        self.chat_frame = tk.Frame(self.chat_canvas, bg=self.chat_bg_dark)
        self.chat_window = self.chat_canvas.create_window(
            (0, 0), window=self.chat_frame, anchor="n"
        )

        self.chat_canvas.bind(
            "<Configure>",
            lambda e: self.chat_canvas.itemconfig(self.chat_window, width=e.width),
        )
        self.chat_frame.bind(
            "<Configure>",
            lambda e: self.chat_canvas.configure(
                scrollregion=self.chat_canvas.bbox("all")
            ),
        )

        # mouse wheel + keyboard scrolling
        self.chat_canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.chat_canvas.bind_all("<Button-4>", self._on_mousewheel)   # Linux up
        self.chat_canvas.bind_all("<Button-5>", self._on_mousewheel)   # Linux down
        self.chat_canvas.bind_all("<Up>", self._on_key_scroll)
        self.chat_canvas.bind_all("<Down>", self._on_key_scroll)
        self.chat_canvas.bind_all("<Prior>", self._on_page_scroll)     # PageUp
        self.chat_canvas.bind_all("<Next>", self._on_page_scroll)      # PageDown

        # ---------- input bar ----------
        input_frame = tk.Frame(self.main_frame, bg="#242022")
        input_frame.pack(fill="x")

        self.entry = tk.Entry(
            input_frame, font=self.chat_font,
            bg="#242022", fg="white", insertbackground="white"
        )
        self.entry.pack(side="left", fill="x", expand=True, padx=10, pady=8)
        self.entry.bind("<Return>", self._on_enter)

        send_btn = tk.Button(
            input_frame, text="Send", command=self._on_click,
            bg="#ce90bc", fg="black", font=self.chat_font,
            relief="flat", padx=15, pady=5
        )
        send_btn.pack(side="right", padx=10, pady=8)

        # ---------- initial message ----------
        self.add_assistant_message("Hello Boss, how can I assist you now?")
        speak("Hello Boss, how can I assist you now?")

        # background voice loop (start after greeting)
        self._voice_running = True
        self.after(1500, self._start_voice_loop)

    # ---------- start voice loop ----------
    def _start_voice_loop(self):
        threading.Thread(target=self._voice_loop_thread, daemon=True).start()

    # ---------- background resize ----------
    def _resize_bg(self, event):
        if event.width <= 1 or event.height <= 1:
            return
        img = self.bg_original.resize((event.width, event.height), Image.LANCZOS)
        self.bg_photo = ImageTk.PhotoImage(img)
        self.bg_label.configure(image=self.bg_photo)

    # ---------- mode toggle ----------
    def _toggle_mode(self):
        self.dark_mode = not self.dark_mode
        if self.dark_mode:
            new_bg = self.chat_bg_dark
            bubble_bg = self.bubble_dark
            fg = "white"
            self.mode_btn.config(text="🌙")
        else:
            new_bg = self.chat_bg_light
            bubble_bg = self.bubble_light
            fg = "black"
            self.mode_btn.config(text="☀")

        self.chat_canvas.config(bg=new_bg)
        self.chat_frame.config(bg=new_bg)

        for bubble in self.chat_frame.winfo_children():
            if isinstance(bubble, tk.Frame):
                bubble.config(bg=bubble_bg)
                for label in bubble.winfo_children():
                    if isinstance(label, tk.Label):
                        label.config(bg=bubble_bg, fg=fg)

    # ---------- chat bubbles ----------
    def _current_bubble_colors(self):
        if self.dark_mode:
            return self.bubble_dark, "white"
        else:
            return self.bubble_light, "black"

    def add_user_message(self, text):
        bg, fg = self._current_bubble_colors()
        bubble = tk.Frame(self.chat_frame, bg=bg, padx=20, pady=15)
        tk.Label(
            bubble, text=text, bg=bg, fg=fg,
            font=self.chat_font, wraplength=900, justify="right"
        ).pack()
        bubble.pack(pady=10, anchor="e", padx=25)

    def add_assistant_message(self, text):
        bg, fg = self._current_bubble_colors()
        bubble = tk.Frame(self.chat_frame, bg=bg, padx=20, pady=15)
        tk.Label(
            bubble, text=text, bg=bg, fg=fg,
            font=self.chat_font, wraplength=900, justify="left"
        ).pack()
        bubble.pack(pady=10, anchor="w", padx=20)

    # ---------- scrolling handlers ----------
    def _on_mousewheel(self, event):
        if event.num == 4 or event.delta > 0:
            self.chat_canvas.yview_scroll(-1, "units")
        elif event.num == 5 or event.delta < 0:
            self.chat_canvas.yview_scroll(1, "units")

    def _on_key_scroll(self, event):
        if event.keysym == "Up":
            self.chat_canvas.yview_scroll(-1, "units")
        elif event.keysym == "Down":
            self.chat_canvas.yview_scroll(1, "units")

    def _on_page_scroll(self, event):
        if event.keysym == "Prior":   # PageUp
            self.chat_canvas.yview_scroll(-1, "pages")
        elif event.keysym == "Next":  # PageDown
            self.chat_canvas.yview_scroll(1, "pages")

    # ---------- voice loop (background thread) ----------
    def _voice_loop_thread(self):
        while self._voice_running:
            spoken = listen()
            if spoken:
                self.after(0, self._handle_voice_command, spoken)
                time.sleep(0.5)
            else:
                time.sleep(0.2)

    def _handle_voice_command(self, spoken):
        self.add_user_message(spoken)
        result = execute_command(spoken)
        self.add_assistant_message(result)
        speak(result)

    # ---------- text input ----------
    def _on_enter(self, event):
        self._send_text()

    def _on_click(self):
        self._send_text()

    def _send_text(self):
        text = self.entry.get().strip()
        if not text:
            return
        self.entry.delete(0, "end")
        self.add_user_message(text)
        result = execute_command(text)
        self.add_assistant_message(result)
        speak(result)

    def destroy(self):
        self._voice_running = False
        super().destroy()


if __name__ == "__main__":
    app = HeimdallUI()
    app.mainloop()

