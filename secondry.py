import tkinter as tk
from tkinter import font as tkfont
from PIL import Image, ImageTk  # requires Pillow

from assistant_core import execute_command, listen


class HeimdallUI(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("HEIMDALL Assistant")
        self.state("zoomed")  # full screen / responsive
        self.configure(bg="#020617")

        # ---------- fonts ----------
        self.chat_font = tkfont.Font(family="Merriweather 18pt", size=14)
        self.title_font = tkfont.Font(family="Merriweather SemiBold 18pt", size=16)

        # ---------- main frame centered ----------
        self.main_frame = tk.Frame(self, bg="#020617", bd=0)
        self.main_frame.place(
            relx=0.5,
            rely=0.5,
            anchor="center",
            relwidth=0.9,
            relheight=0.9,
        )

        # ---------- title bar ----------
        self.title_bar = tk.Frame(self.main_frame, bg="#0b1120", height=40)
        self.title_bar.pack(fill="x")

        # left title label
        tk.Label(
            self.title_bar,
            text="HEIMDALL",
            fg="white",
            bg="#0b1120",
            font=self.title_font,
        ).pack(side="left", padx=10, pady=5)

        # right logo
        logo_img = Image.open("assets/logo.png").resize((32, 32), Image.LANCZOS)
        self.logo_photo = ImageTk.PhotoImage(logo_img)
        tk.Label(self.title_bar, image=self.logo_photo, bg="#0b1120").pack(
            side="right", padx=10, pady=5
        )

        # ---------- chat area (centered messages + own scrollbar) ----------
        chat_container = tk.Frame(self.main_frame, bg="#020617")
        chat_container.pack(fill="both", expand=True)

        self.chat_canvas = tk.Canvas(
            chat_container, bg="#020617", highlightthickness=0
        )
        self.chat_canvas.pack(side="left", fill="both", expand=True)

        scrollbar = tk.Scrollbar(chat_container, command=self.chat_canvas.yview)
        scrollbar.pack(side="right", fill="y")
        self.chat_canvas.configure(yscrollcommand=scrollbar.set)

        self.chat_frame = tk.Frame(self.chat_canvas, bg="#020617")
        self.chat_window = self.chat_canvas.create_window(
            (0, 0), window=self.chat_frame, anchor="n"
        )

        self.chat_canvas.bind(
            "<Configure>",
            lambda e: self.chat_canvas.itemconfig(
                self.chat_window, width=e.width
            ),
        )

        self.chat_frame.bind(
            "<Configure>",
            lambda e: self.chat_canvas.configure(
                scrollregion=self.chat_canvas.bbox("all")
            ),
        )

        # ---------- input bar ----------
        input_frame = tk.Frame(self.main_frame, bg="#020617")
        input_frame.pack(fill="x")

        self.entry = tk.Entry(
            input_frame,
            font=self.chat_font,
            bg="#020617",
            fg="white",
            insertbackground="white",
        )
        self.entry.pack(side="left", fill="x", expand=True, padx=10, pady=8)
        self.entry.bind("<Return>", self._on_enter)

        send_btn = tk.Button(
            input_frame,
            text="🚀",
            command=self._on_click,
            bg="#1d4ed8",
            fg="white",
            font=self.chat_font,
            relief="flat",
            padx=15,
            pady=5,
        )
        send_btn.pack(side="right", padx=10, pady=8)

        mic_btn = tk.Button(
            input_frame,
            text="🎤",
            command=self._start_voice_once,
            bg="#1d4ed8",
            fg="white",
            font=self.chat_font,
            relief="flat",
            padx=10,
            pady=5,
        )
        mic_btn.pack(side="right", padx=(0, 10), pady=8)

        # ---------- initial message ----------
        self.add_assistant_message("Hello Boss, how can I assist you now?")

        # start continuous voice loop after short delay
        self.after(800, self._voice_loop)

    # ---------- chat rendering (centered bubbles) ----------
    def add_user_message(self, text):
        bubble = tk.Frame(self.chat_frame, bg="#020617", padx=30, pady=30)
        tk.Label(
            bubble,
            text=text,
            bg="#020617",
            fg="white",
            font=self.chat_font,
            wraplength=900,
            justify="center",
        ).pack()
        bubble.pack(pady=20, anchor="center")

    def add_assistant_message(self, text):
        bubble = tk.Frame(self.chat_frame, bg="#020617", padx=30, pady=30)
        tk.Label(
            bubble,
            text=text,
            bg="#020617",
            fg="white",
            font=self.chat_font,
            wraplength=900,
            justify="center",
        ).pack()
        bubble.pack(pady=20, anchor="center")

    # ---------- continuous voice loop ----------
    def _voice_loop(self):
        """
        Listen once, handle command, then schedule itself again.
        Runs until process exits (execute_command handles 'exit').
        """
        spoken = listen()
        if spoken:
            self.add_user_message(spoken)
            result = execute_command(spoken)
            self.add_assistant_message(result)
        # schedule next listen
        self.after(500, self._voice_loop)

    # manual mic trigger (one-shot)
    def _start_voice_once(self):
        spoken = listen()
        if not spoken:
            return
        self.add_user_message(spoken)
        result = execute_command(spoken)
        self.add_assistant_message(result)

    # ---------- events (text input) ----------
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


if __name__ == "__main__":
    app = HeimdallUI()
    app.mainloop()
