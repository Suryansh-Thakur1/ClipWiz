import tkinter as tk
import pyperclip
import keyboard
import threading
import time
import winshell
import os
import sys
from win32com.client import Dispatch

history = []

def add_to_startup(file_path=None, name="ClipWiz"):
    if file_path is None:
        file_path = os.path.abspath(sys.executable)

    startup = winshell.startup()
    shortcut_path = os.path.join(startup, f"{name}.lnk")

    if not os.path.exists(shortcut_path):
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = file_path
        shortcut.WorkingDirectory = os.path.dirname(file_path)
        shortcut.IconLocation = file_path
        shortcut.save()

def check_clipboard():
    last = ""
    while True:
        try:
            current = pyperclip.paste()
            if current != last and current.strip() != "":
                last = current
                if current not in history:
                    history.insert(0, current)
                    if len(history) > 10:
                        history.pop()
        except:
            pass
        time.sleep(0.5)

def copy_again(text):
    pyperclip.copy(text)

def open_window():
    win = tk.Tk()
    win.title("ClipWiz")
    win.configure(bg="#1e1e1e")
    win.attributes("-topmost", True)
    win.attributes("-alpha", 0.95)

    screen_width = win.winfo_screenwidth()
    width = 450
    x = screen_width - width - 40
    y = 80

    win.geometry(f"{width}x500+{x}+{y}")
    win.resizable(False, False)

    canvas = tk.Canvas(win, bg="#1e1e1e", bd=0, highlightthickness=0)
    canvas.pack(side="left", fill="both", expand=True)

    scrollbar = tk.Scrollbar(win, command=canvas.yview, troughcolor="#2e2e2e", bg="#444", activebackground="#555")
    scrollbar.pack(side="right", fill="y")

    canvas.configure(yscrollcommand=scrollbar.set)

    inner_frame = tk.Frame(canvas, bg="#1e1e1e")
    canvas.create_window((0, 0), window=inner_frame, anchor="nw")

    def on_frame_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def on_mouse_scroll(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    canvas.bind_all("<MouseWheel>", on_mouse_scroll)
    inner_frame.bind("<Configure>", on_frame_configure)

    for item in history:
        frame = tk.Frame(inner_frame, bg="#1e1e1e")
        frame.pack(fill="x", padx=6, pady=4)

        preview = item.replace("\n", " ")[:120]
        if len(item) > len(preview):
            preview += "..."

        label = tk.Label(
            frame,
            text=preview,
            fg="white",
            bg="#1e1e1e",
            anchor="w",
            justify="left",
            wraplength=350,
            font=("Segoe UI", 10)
        )
        label.pack(side="left", fill="x", expand=True)

        btn = tk.Button(
            frame,
            text="📋",
            command=lambda t=item: copy_again(t),
            bg="#3a3a3a",
            fg="white",
            relief="flat",
            width=3,
            height=1,
            font=("Segoe UI", 10)
        )
        btn.pack(side="right", padx=4)

    win.mainloop()

def start_hotkey():
    keyboard.add_hotkey("ctrl+alt+w", open_window)
    keyboard.wait()

# Add to startup silently
add_to_startup()

# Start clipboard monitoring + hotkey listener in background
threading.Thread(target=check_clipboard, daemon=True).start()
threading.Thread(target=start_hotkey, daemon=True).start()

# Keep app alive
while True:
    time.sleep(1)
