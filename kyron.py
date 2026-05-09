"""
Kyron - lightweight desktop auto clicker.

Dependency:
    pip install pynput

Run:
    python kyron.py
"""

import json
import random
import re
import sys
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, simpledialog, ttk

from pynput import keyboard as pynput_keyboard
from pynput import mouse as pynput_mouse


APP_TITLE = "Kyron - AutoCliker"
DEFAULT_HOTKEY = "/"
UNTITLED_SCRIPT = "Script baru"
CLICK_POSITION_JITTER = 10
DELAY_JITTER_MS = 12


def app_base_path():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resource_path(filename):
    base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base_dir / filename


APP_DIR = app_base_path()
SCRIPT_DIR = APP_DIR / "scripts"
LEGACY_SCRIPT_FILE = APP_DIR / "scripts.json"
OLD_SCRIPT_DIR = APP_DIR / "kyron_scripts"
OLD_LEGACY_SCRIPT_FILE = APP_DIR / "kyron_scripts.json"
OLDER_SCRIPT_DIR = APP_DIR / "blueclick_scripts"
OLDER_LEGACY_SCRIPT_FILE = APP_DIR / "blueclick_scripts.json"
LOGO_FILE = APP_DIR / "logo.png"


THEMES = {
    "dark": {
        "bg": "#0B1120",
        "panel": "#111A2E",
        "panel_alt": "#0F172A",
        "input": "#16213B",
        "border": "#1F2A44",
        "grid": "#F2F4FF",
        "text": "#F2F4FF",
        "muted": "#94A3C3",
        "accent": "#4F46FF",
        "accent_alt": "#00F0FF",
        "button": "#16213B",
        "button_active": "#1E2C4C",
        "button_primary": "#4F46FF",
        "button_primary_active": "#655DFF",
        "button_soft": "#132846",
        "button_soft_active": "#19355E",
        "danger": "#A83B66",
        "status": "#0F172A",
        "selection": "#203B73",
    },
    "light": {
        "bg": "#EEF4FF",
        "panel": "#FFFFFF",
        "panel_alt": "#F7FAFF",
        "input": "#F8FBFF",
        "border": "#D7E1F2",
        "grid": "#7A8AA6",
        "text": "#101A2B",
        "muted": "#5E6C84",
        "accent": "#4F46FF",
        "accent_alt": "#0EA5FF",
        "button": "#EEF4FF",
        "button_active": "#E3ECFF",
        "button_primary": "#4F46FF",
        "button_primary_active": "#6159FF",
        "button_soft": "#DFF5FF",
        "button_soft_active": "#C7EEFF",
        "danger": "#D5527B",
        "status": "#EAF0FA",
        "selection": "#CFE4FF",
    },
}


class KyronApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("860x620")
        self.minsize(720, 520)

        self.theme_name = "dark"
        self.colors = THEMES[self.theme_name]
        self.scripts = {}
        self.script_names = []
        self.current_script = tk.StringVar(value=UNTITLED_SCRIPT)
        self.actions = []
        self.is_running = False
        self.stop_event = threading.Event()
        self.worker_thread = None
        self.hotkey = tk.StringVar(value=DEFAULT_HOTKEY)
        self.picking_coordinate = False
        self.app_icon = None

        self.set_window_icon()
        self.mouse_controller = pynput_mouse.Controller()
        self.keyboard_controller = pynput_keyboard.Controller()
        self.hotkey_listener = pynput_keyboard.Listener(on_press=self.on_global_key_press)
        self.hotkey_listener.start()

        self.x_var = tk.IntVar(value=0)
        self.y_var = tk.IntVar(value=0)
        self.delay_var = tk.IntVar(value=600)
        self.repeat_var = tk.IntVar(value=1)
        self.key_var = tk.StringVar()

        self.load_scripts()
        self.ensure_current_script()
        self.build_ui()
        self.center_window()
        self.apply_theme()
        self.refresh_script_list()
        self.show_empty_action_list()
        self.protocol("WM_DELETE_WINDOW", self.close_app)

    def build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.main_frame = tk.Frame(self)
        self.main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10, 0))
        self.main_frame.columnconfigure(0, minsize=260, weight=0)
        self.main_frame.columnconfigure(1, weight=1)
        self.main_frame.rowconfigure(0, weight=1)

        self.left_panel = tk.Frame(self.main_frame, bd=0, highlightthickness=1)
        self.left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.left_panel.columnconfigure(0, weight=1)
        self.left_panel.rowconfigure(1, weight=1)

        self.script_label = tk.Label(self.left_panel, text="Daftar Script", anchor="w", font=("Segoe UI", 11, "bold"))
        self.script_label.grid(row=0, column=0, sticky="ew", padx=12, pady=(14, 8))

        self.script_table = tk.Frame(self.left_panel, bd=0, highlightthickness=1)
        self.script_table.grid(row=1, column=0, sticky="nsew", padx=12)
        self.script_table.columnconfigure(0, weight=1)
        self.script_table.rowconfigure(0, weight=1)

        self.script_list = ttk.Treeview(
            self.script_table,
            columns=("no", "name"),
            show="headings",
            selectmode="browse",
            height=12,
        )
        self.script_list.heading("no", text="No")
        self.script_list.heading("name", text="Script")
        self.script_list.column("no", width=64, minwidth=64, anchor="center", stretch=False)
        self.script_list.column("name", anchor="w", stretch=True)
        self.script_list.grid(row=0, column=0, sticky="nsew")
        self.script_list.bind("<<TreeviewSelect>>", self.on_script_selected)
        self.script_list.bind("<Button-3>", self.show_script_menu)
        self.script_list.bind("<Button-2>", self.show_script_menu)
        self.script_divider = tk.Frame(self.script_table, width=3, bd=0, highlightthickness=0)
        self.script_divider.place(x=63, y=0, width=3, relheight=1)
        self.script_divider.lift()
        self.script_divider_left = tk.Frame(self.script_divider, width=1, bd=0, highlightthickness=0)
        self.script_divider_left.place(x=0, y=0, width=1, relheight=1)
        self.script_divider_right = tk.Frame(self.script_divider, width=1, bd=0, highlightthickness=0)
        self.script_divider_right.place(x=2, y=0, width=1, relheight=1)

        self.control_label = tk.Label(self.left_panel, text="Kontrol", anchor="w", font=("Segoe UI", 11, "bold"))
        self.control_label.grid(row=2, column=0, sticky="ew", padx=12, pady=(14, 8))

        self.start_button = tk.Button(self.left_panel, text="Mulai", command=self.toggle_running, font=("Segoe UI", 9, "bold"))
        self.start_button.grid(row=3, column=0, sticky="ew", padx=12, pady=(0, 5), ipady=5)
        self.start_button._role = "primary"

        self.theme_button = tk.Button(self.left_panel, text="Ganti Tema", command=self.toggle_theme, font=("Segoe UI", 9, "bold"))
        self.theme_button.grid(row=4, column=0, sticky="ew", padx=12, pady=(2, 14), ipady=5)

        self.script_menu = tk.Menu(self, tearoff=0)
        self.script_menu.add_command(label="Rename", command=self.rename_selected_script)
        self.script_menu.add_command(label="Copy", command=self.copy_selected_script)
        self.script_menu.add_command(label="Hapus", command=self.delete_selected_script_file)

        self.right_panel = tk.Frame(self.main_frame, bd=0, highlightthickness=1)
        self.right_panel.grid(row=0, column=1, sticky="nsew")
        self.right_panel.columnconfigure(0, weight=1)
        self.right_panel.rowconfigure(7, weight=1)

        self.build_coordinate_section()
        self.build_general_section()
        self.build_keyboard_section()
        self.build_action_section()
        self.build_hotkey_section()

        self.status_card = tk.Frame(self, bd=0, highlightthickness=1)
        self.status_card.grid(row=1, column=0, sticky="ew", padx=10, pady=(8, 10))
        self.status_card.columnconfigure(0, weight=1)

        self.status = tk.Label(self.status_card, text="Status: Siap", anchor="w", font=("Segoe UI", 9))
        self.status.grid(row=0, column=0, sticky="ew", padx=12, pady=7)

    def build_coordinate_section(self):
        label = tk.Label(self.right_panel, text="Input Koordinat", anchor="w", font=("Segoe UI", 11, "bold"))
        label.grid(row=0, column=0, sticky="ew", padx=12, pady=(14, 8))

        row = tk.Frame(self.right_panel)
        row.grid(row=1, column=0, sticky="ew", padx=10)
        row.columnconfigure(2, weight=1)
        row.columnconfigure(4, weight=1)

        self.pick_button = tk.Button(row, text="Pick", command=self.start_pick_coordinate, font=("Segoe UI", 9, "bold"))
        self.pick_button.grid(row=0, column=0, padx=(0, 6), ipady=4, ipadx=8)
        self.pick_button._role = "soft"
        tk.Label(row, text="X:", font=("Segoe UI", 9)).grid(row=0, column=1, sticky="w")
        self.x_input = tk.Spinbox(row, from_=0, to=99999, textvariable=self.x_var, font=("Segoe UI", 9))
        self.x_input.grid(row=0, column=2, sticky="ew", padx=(5, 6), ipady=3)
        tk.Label(row, text="Y:", font=("Segoe UI", 9)).grid(row=0, column=3, sticky="w")
        self.y_input = tk.Spinbox(row, from_=0, to=99999, textvariable=self.y_var, font=("Segoe UI", 9))
        self.y_input.grid(row=0, column=4, sticky="ew", padx=(5, 6), ipady=3)
        self.add_click_button = tk.Button(row, text="Tambah", command=self.add_click_action, font=("Segoe UI", 9, "bold"))
        self.add_click_button.grid(row=0, column=5, ipady=4, ipadx=8)
        self.add_click_button._role = "primary"

    def build_general_section(self):
        label = tk.Label(self.right_panel, text="Pengaturan Umum", anchor="w", font=("Segoe UI", 11, "bold"))
        label.grid(row=2, column=0, sticky="ew", padx=12, pady=(14, 8))

        row = tk.Frame(self.right_panel)
        row.grid(row=3, column=0, sticky="ew", padx=10)
        row.columnconfigure(1, weight=1)
        row.columnconfigure(3, weight=1)

        tk.Label(row, text="Delay (ms):", font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w")
        self.delay_input = tk.Spinbox(row, from_=1, to=999999, textvariable=self.delay_var, font=("Segoe UI", 9))
        self.delay_input.grid(row=0, column=1, sticky="ew", padx=(6, 10), ipady=3)
        tk.Label(row, text="Repetisi:", font=("Segoe UI", 9)).grid(row=0, column=2, sticky="w")
        self.repeat_input = tk.Spinbox(row, from_=0, to=999999, textvariable=self.repeat_var, font=("Segoe UI", 9))
        self.repeat_input.grid(row=0, column=3, sticky="ew", padx=(6, 0), ipady=3)

    def build_keyboard_section(self):
        label = tk.Label(self.right_panel, text="Input Keyboard", anchor="w", font=("Segoe UI", 11, "bold"))
        label.grid(row=4, column=0, sticky="ew", padx=12, pady=(14, 8))

        row = tk.Frame(self.right_panel)
        row.grid(row=5, column=0, sticky="ew", padx=10)
        row.columnconfigure(1, weight=1)

        tk.Label(row, text="Key:", font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w")
        self.key_entry = tk.Entry(row, textvariable=self.key_var, font=("Segoe UI", 9))
        self.key_entry.grid(row=0, column=1, sticky="ew", padx=(5, 6), ipady=4)
        self.add_key_button = tk.Button(row, text="Tambah Key", command=self.add_key_action, font=("Segoe UI", 9, "bold"))
        self.add_key_button.grid(row=0, column=2, ipady=4, ipadx=8)
        self.add_key_button._role = "primary"

    def build_action_section(self):
        self.action_label = tk.Label(self.right_panel, text="Daftar Aksi", anchor="w", font=("Segoe UI", 11, "bold"))
        self.action_label.grid(row=6, column=0, sticky="ew", padx=12, pady=(14, 8))

        area = tk.Frame(self.right_panel)
        area.grid(row=7, column=0, sticky="nsew", padx=10)
        area.columnconfigure(0, weight=1)
        area.rowconfigure(0, weight=1)

        self.action_table = tk.Frame(area, bd=0, highlightthickness=1)
        self.action_table.grid(row=0, column=0, sticky="nsew")
        self.action_table.columnconfigure(0, weight=1)
        self.action_table.rowconfigure(0, weight=1)

        self.action_list = ttk.Treeview(
            self.action_table,
            columns=("no", "detail"),
            show="headings",
            selectmode="browse",
            height=12,
        )
        self.action_list.heading("no", text="No")
        self.action_list.heading("detail", text="Aksi")
        self.action_list.column("no", width=64, minwidth=64, anchor="center", stretch=False)
        self.action_list.column("detail", anchor="w", stretch=True)
        self.action_list.grid(row=0, column=0, sticky="nsew")
        self.action_list.bind("<Button-3>", self.show_action_menu)
        self.action_list.bind("<Button-2>", self.show_action_menu)
        self.action_divider = tk.Frame(self.action_table, width=3, bd=0, highlightthickness=0)
        self.action_divider.place(x=63, y=0, width=3, relheight=1)
        self.action_divider.lift()
        self.action_divider_left = tk.Frame(self.action_divider, width=1, bd=0, highlightthickness=0)
        self.action_divider_left.place(x=0, y=0, width=1, relheight=1)
        self.action_divider_right = tk.Frame(self.action_divider, width=1, bd=0, highlightthickness=0)
        self.action_divider_right.place(x=2, y=0, width=1, relheight=1)

        self.action_buttons = tk.Frame(self.right_panel)
        self.action_buttons.grid(row=8, column=0, sticky="ew", padx=10, pady=(8, 0))
        self.action_buttons.columnconfigure(0, weight=1)
        self.action_buttons.columnconfigure(1, weight=1)

        self.save_button = tk.Button(
            self.action_buttons,
            text="Simpan Script",
            command=self.save_current_script,
            font=("Segoe UI", 9, "bold"),
        )
        self.save_button.grid(row=0, column=0, sticky="ew", padx=(0, 5), ipady=5)
        self.save_button._role = "soft"

        self.load_button = tk.Button(
            self.action_buttons,
            text="Reload",
            command=self.reload_saved_scripts,
            font=("Segoe UI", 9, "bold"),
        )
        self.load_button.grid(row=0, column=1, sticky="ew", padx=(5, 0), ipady=5)
        self.load_button._role = "soft"

        self.action_menu = tk.Menu(self, tearoff=0)
        self.action_menu.add_command(label="Edit", command=self.edit_selected_action)
        self.action_menu.add_command(label="Hapus", command=self.delete_selected_action)
        self.action_menu.add_separator()
        self.action_menu.add_command(label="Naikkan", command=lambda: self.move_action(-1))
        self.action_menu.add_command(label="Turunkan", command=lambda: self.move_action(1))
        self.action_menu.add_separator()
        self.action_menu.add_command(label="Hapus Semua", command=self.clear_actions)

    def build_hotkey_section(self):
        label = tk.Label(self.right_panel, text="Hotkey Toggle", anchor="w", font=("Segoe UI", 11, "bold"))
        label.grid(row=9, column=0, sticky="ew", padx=12, pady=(14, 8))

        row = tk.Frame(self.right_panel)
        row.grid(row=10, column=0, sticky="ew", padx=10, pady=(0, 14))
        row.columnconfigure(1, weight=1)

        tk.Label(row, text="Toggle saat ini:", font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w")
        self.hotkey_entry = tk.Entry(row, textvariable=self.hotkey, font=("Segoe UI", 9))
        self.hotkey_entry.grid(row=0, column=1, sticky="ew", padx=(6, 6), ipady=4)
        self.set_hotkey_button = tk.Button(row, text="Set", command=self.set_hotkey, font=("Segoe UI", 9, "bold"))
        self.set_hotkey_button.grid(row=0, column=2, ipady=4, ipadx=10)
        self.set_hotkey_button._role = "soft"

    def set_window_icon(self):
        icon_path = resource_path("logo.ico")
        png_path = resource_path("logo.png")
        try:
            if icon_path.exists():
                self.iconbitmap(default=str(icon_path))
        except tk.TclError:
            pass

        try:
            if png_path.exists():
                self.app_icon = tk.PhotoImage(file=str(png_path))
                self.iconphoto(True, self.app_icon)
        except tk.TclError:
            pass

    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = max(0, (screen_width - width) // 2)
        y = max(0, (screen_height - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def apply_theme(self):
        self.colors = THEMES[self.theme_name]
        self.configure(bg=self.colors["bg"])
        self.set_widget_colors(self)
        self.apply_treeview_theme()
        self.script_table.configure(bg=self.colors["grid"], highlightbackground=self.colors["grid"], highlightcolor=self.colors["grid"])
        self.action_table.configure(bg=self.colors["grid"], highlightbackground=self.colors["grid"], highlightcolor=self.colors["grid"])
        self.script_divider.configure(bg=self.colors["input"])
        self.action_divider.configure(bg=self.colors["input"])
        self.script_divider_left.configure(bg=self.colors["grid"])
        self.script_divider_right.configure(bg=self.colors["grid"])
        self.action_divider_left.configure(bg=self.colors["grid"])
        self.action_divider_right.configure(bg=self.colors["grid"])
        self.status.configure(bg=self.colors["status"], fg=self.colors["text"])
        if hasattr(self, "action_menu"):
            self.action_menu.configure(
                bg=self.colors["panel"],
                fg=self.colors["text"],
                activebackground=self.colors["selection"],
                activeforeground=self.colors["text"],
                relief="flat",
                bd=0,
            )
        if hasattr(self, "script_menu"):
            self.script_menu.configure(
                bg=self.colors["panel"],
                fg=self.colors["text"],
                activebackground=self.colors["selection"],
                activeforeground=self.colors["text"],
                relief="flat",
                bd=0,
            )

    def set_widget_colors(self, widget):
        for child in widget.winfo_children():
            class_name = child.winfo_class()
            if class_name in {"Frame", "LabelFrame"}:
                bg_color = child.master.cget("bg")
                if child in {self.left_panel, self.right_panel, self.status_card}:
                    bg_color = self.colors["panel"]
                elif child.master in {self.left_panel, self.right_panel, self.status_card}:
                    bg_color = self.colors["panel"]
                child.configure(bg=bg_color)
                if child in {self.left_panel, self.right_panel, self.status_card, self.script_table, self.action_table}:
                    child.configure(highlightbackground=self.colors["border"], highlightcolor=self.colors["border"])
            elif class_name == "Label":
                child.configure(bg=child.master.cget("bg"), fg=self.colors["text"])
            elif class_name == "Button":
                role = getattr(child, "_role", "default")
                bg = self.colors["button"]
                active_bg = self.colors["button_active"]
                fg = self.colors["text"]
                if role == "primary":
                    bg = self.colors["button_primary"]
                    active_bg = self.colors["button_primary_active"]
                elif role == "soft":
                    bg = self.colors["button_soft"]
                    active_bg = self.colors["button_soft_active"]
                child.configure(
                    bg=bg,
                    fg=fg,
                    activebackground=active_bg,
                    activeforeground=fg,
                    relief="flat",
                    bd=0,
                    cursor="hand2",
                    highlightthickness=0,
                )
            elif class_name in {"Entry", "Spinbox"}:
                child.configure(
                    bg=self.colors["input"],
                    fg=self.colors["text"],
                    insertbackground=self.colors["text"],
                    relief="flat",
                    bd=0,
                    highlightthickness=1,
                    highlightbackground=self.colors["border"],
                    highlightcolor=self.colors["accent_alt"],
                )
            elif class_name == "Listbox":
                child.configure(
                    bg=self.colors["input"],
                    fg=self.colors["text"],
                    selectbackground=self.colors["selection"],
                    selectforeground=self.colors["text"],
                    relief="flat",
                    bd=0,
                    highlightthickness=1,
                    highlightbackground=self.colors["border"],
                    highlightcolor=self.colors["accent"],
                )
            self.set_widget_colors(child)

    def apply_treeview_theme(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(
            "Kyron.Treeview",
            background=self.colors["input"],
            fieldbackground=self.colors["input"],
            foreground=self.colors["text"],
            bordercolor=self.colors["border"],
            lightcolor=self.colors["border"],
            darkcolor=self.colors["border"],
            rowheight=18,
            borderwidth=1,
            relief="solid",
            font=("Segoe UI", 10),
        )
        style.map(
            "Kyron.Treeview",
            background=[("selected", self.colors["selection"])],
            foreground=[("selected", self.colors["text"])],
        )
        style.configure(
            "Kyron.Treeview.Heading",
            background=self.colors["panel_alt"],
            foreground=self.colors["text"],
            bordercolor=self.colors["border"],
            lightcolor=self.colors["border"],
            darkcolor=self.colors["border"],
            borderwidth=1,
            relief="solid",
            font=("Segoe UI", 10, "bold"),
        )
        self.script_list.configure(style="Kyron.Treeview")
        self.action_list.configure(style="Kyron.Treeview")

    def toggle_theme(self):
        self.theme_name = "light" if self.theme_name == "dark" else "dark"
        self.apply_theme()

    def load_scripts(self):
        self.scripts = {}
        self.load_script_files()
        if self.scripts and not any(SCRIPT_DIR.glob("*.json")):
            self.persist_scripts()
        if not self.scripts:
            self.load_legacy_script_file()
            if self.scripts:
                self.persist_scripts()

    def load_script_files(self):
        self.load_script_files_from(SCRIPT_DIR)
        if not self.scripts:
            self.load_script_files_from(OLD_SCRIPT_DIR)
        if not self.scripts:
            self.load_script_files_from(OLDER_SCRIPT_DIR)

    def load_script_files_from(self, directory):
        if not directory.exists():
            return

        for script_path in sorted(directory.glob("*.json")):
            try:
                data = json.loads(script_path.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                continue

            if not isinstance(data, dict):
                continue

            name = str(data.get("name") or script_path.stem).strip()
            if not name:
                continue
            self.scripts[name] = self.normalize_script_data(data)

    def load_legacy_script_file(self):
        self.load_legacy_script_file_from(LEGACY_SCRIPT_FILE)
        if not self.scripts:
            self.load_legacy_script_file_from(OLD_LEGACY_SCRIPT_FILE)
        if not self.scripts:
            self.load_legacy_script_file_from(OLDER_LEGACY_SCRIPT_FILE)

    def load_legacy_script_file_from(self, file_path):
        if not file_path.exists():
            return

        try:
            data = json.loads(file_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return

        if not isinstance(data, dict):
            return

        for name, script in data.items():
            if not isinstance(script, dict):
                continue
            name = str(name).strip()
            if name:
                self.scripts[name] = self.normalize_script_data(script)

    def normalize_script_data(self, script):
        return {
            "hotkey": script.get("hotkey", DEFAULT_HOTKEY),
            "actions": list(script.get("actions", [])),
            "delay_ms": max(1, self.safe_int(script.get("delay_ms", 600), 600)),
            "repeat": max(0, self.safe_int(script.get("repeat", 1), 1)),
        }

    def safe_int(self, value, default):
        try:
            return int(value)
        except (TypeError, ValueError):
            return default

    def ensure_current_script(self):
        if self.current_script.get() in self.scripts:
            return
        first_script = next(iter(sorted(self.scripts)), UNTITLED_SCRIPT)
        self.current_script.set(first_script)

    def persist_scripts(self):
        SCRIPT_DIR.mkdir(exist_ok=True)
        for name, script in self.scripts.items():
            self.write_script_file(name, script)

    def write_script_file(self, name, script):
        SCRIPT_DIR.mkdir(exist_ok=True)
        payload = {"name": name, **self.normalize_script_data(script)}
        self.script_file_path(name).write_text(json.dumps(payload, indent=2), encoding="utf-8")

    def delete_script_storage(self, name):
        safe_name = self.safe_script_filename(name)
        for file_path in (
            SCRIPT_DIR / f"{safe_name}.json",
            OLD_SCRIPT_DIR / f"{safe_name}.json",
            OLDER_SCRIPT_DIR / f"{safe_name}.json",
        ):
            try:
                if file_path.exists():
                    file_path.unlink()
            except OSError:
                continue

    def script_file_path(self, name):
        return SCRIPT_DIR / f"{self.safe_script_filename(name)}.json"

    def safe_script_filename(self, name):
        filename = re.sub(r"[^A-Za-z0-9._ -]+", "_", name).strip()
        filename = re.sub(r"\s+", " ", filename).replace(" ", "-")
        return filename.strip(".-_") or "script"

    def current_script_data(self):
        name = self.current_script.get()
        return self.scripts.get(name)

    def sync_current_script(self):
        script = self.current_script_data()
        if script is None:
            return
        script["hotkey"] = self.hotkey.get().strip() or DEFAULT_HOTKEY
        script["actions"] = list(self.actions)
        script["delay_ms"] = max(1, self.delay_var.get())
        script["repeat"] = max(0, self.repeat_var.get())

    def reload_saved_scripts(self):
        if self.is_running:
            messagebox.showwarning("Script sedang berjalan", "Stop script lebih dulu sebelum memuat ulang.", parent=self)
            return

        previous_script = self.current_script.get()
        self.load_scripts()
        if previous_script not in self.scripts:
            self.ensure_current_script()
        self.refresh_script_list()
        self.show_empty_action_list()

        if self.scripts:
            self.set_status(f"Status: Script dimuat dari folder {SCRIPT_DIR.name}. Pilih script untuk melihat aksi.")
        else:
            self.set_status("Status: Belum ada script tersimpan")

    def refresh_script_list(self):
        self.script_list.delete(*self.script_list.get_children())
        self.script_names = []
        selected_item = None
        names = sorted(self.scripts)
        for index, name in enumerate(names):
            self.script_names.append(name)
            item_id = self.script_list.insert("", "end", values=(index + 1, f"  {name}"))
            if name == self.current_script.get():
                selected_item = item_id
        if names:
            if selected_item is not None:
                self.script_list.selection_set(selected_item)
                self.script_list.focus(selected_item)
                self.script_list.see(selected_item)
        else:
            self.script_list.insert("", "end", values=("", "  (Belum ada script tersimpan)"), tags=("placeholder",))
            self.script_list.tag_configure("placeholder", foreground=self.colors["muted"])

    def on_script_selected(self, _event=None):
        selection = self.script_list.selection()
        if not selection:
            return
        index = self.script_list.index(selection[0])
        if index >= len(self.script_names):
            self.script_list.selection_remove(selection[0])
            return
        selected_name = self.script_names[index]
        if self.current_script.get() != selected_name:
            self.sync_current_script()
        self.current_script.set(selected_name)
        self.load_selected_script()

    def show_script_menu(self, event):
        if not self.scripts:
            return

        item_id = self.script_list.identify_row(event.y)
        if not item_id:
            return

        index = self.script_list.index(item_id)
        if index < 0 or index >= len(self.script_names):
            return

        selected_name = self.script_names[index]
        if selected_name not in self.scripts:
            return

        self.script_list.selection_set(item_id)
        self.script_list.focus(item_id)

        try:
            self.script_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.script_menu.grab_release()
        return "break"

    def load_selected_script(self):
        script = self.scripts.get(self.current_script.get(), {})
        self.actions = list(script.get("actions", []))
        self.delay_var.set(script.get("delay_ms", 600))
        self.repeat_var.set(script.get("repeat", 1))
        self.hotkey.set(script.get("hotkey", DEFAULT_HOTKEY))
        self.refresh_action_list()
        self.set_status(f"Status: Script aktif '{self.current_script.get()}'")

    def show_empty_action_list(self):
        self.actions = []
        self.current_script.set(UNTITLED_SCRIPT)
        self.delay_var.set(600)
        self.repeat_var.set(1)
        self.hotkey.set(DEFAULT_HOTKEY)
        self.script_list.selection_remove(self.script_list.selection())
        self.action_list.delete(*self.action_list.get_children())
        self.action_label.configure(text="Daftar Aksi")
        self.action_list.insert("", "end", values=("", "  (Pilih script untuk menampilkan daftar aksi)"), tags=("placeholder",))
        self.action_list.tag_configure("placeholder", foreground=self.colors["muted"])
        self.update_start_button()

    def save_current_script(self):
        name = simpledialog.askstring("Simpan Script", "Nama script:", initialvalue=self.current_script.get(), parent=self)
        if not name:
            return
        name = name.strip()
        if not name:
            return

        self.current_script.set(name)
        if name not in self.scripts:
            self.scripts[name] = {}
        self.sync_current_script()
        self.write_script_file(name, self.scripts[name])
        self.refresh_script_list()
        self.load_selected_script()
        self.set_status(f"Status: Script '{name}' disimpan")

    def copy_selected_script(self):
        selection = self.script_list.selection()
        if not selection:
            return

        index = self.script_list.index(selection[0])
        if index >= len(self.script_names):
            return

        source_name = self.script_names[index]
        if source_name not in self.scripts:
            return

        new_name = simpledialog.askstring(
            "Copy Script",
            "Nama script baru:",
            initialvalue=f"{source_name} Copy",
            parent=self,
        )
        if not new_name:
            return

        new_name = new_name.strip()
        if not new_name:
            return

        if new_name in self.scripts:
            messagebox.showwarning("Nama sudah ada", "Gunakan nama lain untuk hasil copy script.", parent=self)
            return

        source_script = self.normalize_script_data(self.scripts[source_name])
        self.scripts[new_name] = {
            "hotkey": source_script["hotkey"],
            "actions": list(source_script["actions"]),
            "delay_ms": source_script["delay_ms"],
            "repeat": source_script["repeat"],
        }
        self.write_script_file(new_name, self.scripts[new_name])
        self.current_script.set(new_name)
        self.refresh_script_list()
        self.load_selected_script()
        self.set_status(f"Status: Script '{source_name}' dicopy ke '{new_name}'")

    def rename_selected_script(self):
        selection = self.script_list.selection()
        if not selection:
            return

        index = self.script_list.index(selection[0])
        if index >= len(self.script_names):
            return

        old_name = self.script_names[index]
        if old_name not in self.scripts:
            return

        new_name = simpledialog.askstring(
            "Rename Script",
            "Nama script baru:",
            initialvalue=old_name,
            parent=self,
        )
        if not new_name:
            return

        new_name = new_name.strip()
        if not new_name or new_name == old_name:
            return

        if new_name in self.scripts:
            messagebox.showwarning("Nama sudah ada", "Gunakan nama lain untuk rename script.", parent=self)
            return

        script_data = self.normalize_script_data(self.scripts[old_name])
        self.scripts[new_name] = script_data
        self.write_script_file(new_name, script_data)
        self.delete_script_storage(old_name)
        del self.scripts[old_name]
        self.current_script.set(new_name)
        self.refresh_script_list()
        self.load_selected_script()
        self.set_status(f"Status: Script '{old_name}' diubah menjadi '{new_name}'")

    def delete_selected_script_file(self):
        selection = self.script_list.selection()
        if not selection:
            return

        index = self.script_list.index(selection[0])
        if index >= len(self.script_names):
            return

        name = self.script_names[index]
        if name not in self.scripts:
            return

        confirmed = messagebox.askyesno(
            "Hapus Script",
            f"Hapus script '{name}' dari daftar dan file simpanannya?",
            parent=self,
        )
        if not confirmed:
            return

        self.delete_script_storage(name)
        del self.scripts[name]

        if self.current_script.get() == name:
            self.ensure_current_script()

        self.refresh_script_list()
        if self.scripts:
            self.load_selected_script()
        else:
            self.show_empty_action_list()
        self.set_status(f"Status: Script '{name}' dihapus")

    def add_click_action(self):
        action = {
            "type": "click",
            "x": self.x_var.get(),
            "y": self.y_var.get(),
            "delay_ms": max(1, self.delay_var.get()),
        }
        self.actions.append(action)
        self.sync_current_script()
        self.refresh_action_list()

    def add_key_action(self):
        key = self.key_var.get().strip()
        if not key:
            messagebox.showwarning("Input kosong", "Masukkan tombol keyboard lebih dulu.", parent=self)
            return
        self.actions.append({"type": "key", "key": key, "delay_ms": max(1, self.delay_var.get())})
        self.sync_current_script()
        self.key_var.set("")
        self.refresh_action_list()

    def delete_selected_action(self):
        if not self.actions:
            return
        selection = self.action_list.selection()
        if not selection:
            return
        index = self.action_list.index(selection[0])
        del self.actions[index]
        self.sync_current_script()
        self.refresh_action_list()
        if self.actions:
            next_item = self.action_list.get_children()[min(index, len(self.actions) - 1)]
            self.action_list.selection_set(next_item)

    def edit_selected_action(self):
        if not self.actions:
            return
        selection = self.action_list.selection()
        if not selection:
            return

        index = self.action_list.index(selection[0])
        action = dict(self.actions[index])
        delay = simpledialog.askinteger(
            "Edit Aksi",
            "Delay (ms):",
            initialvalue=action.get("delay_ms", self.delay_var.get()),
            minvalue=1,
            parent=self,
        )
        if delay is None:
            return

        if action["type"] == "click":
            x = simpledialog.askinteger("Edit Aksi", "X:", initialvalue=action.get("x", 0), minvalue=0, parent=self)
            if x is None:
                return
            y = simpledialog.askinteger("Edit Aksi", "Y:", initialvalue=action.get("y", 0), minvalue=0, parent=self)
            if y is None:
                return
            self.actions[index] = {"type": "click", "x": x, "y": y, "delay_ms": delay}
        else:
            key = simpledialog.askstring("Edit Aksi", "Key:", initialvalue=action.get("key", ""), parent=self)
            if key is None:
                return
            key = key.strip()
            if not key:
                messagebox.showwarning("Input kosong", "Key tidak boleh kosong.", parent=self)
                return
            self.actions[index] = {"type": "key", "key": key, "delay_ms": delay}

        self.sync_current_script()
        self.refresh_action_list()
        if self.action_list.get_children():
            self.action_list.selection_set(self.action_list.get_children()[index])

    def move_action(self, direction):
        if not self.actions:
            return
        selection = self.action_list.selection()
        if not selection:
            return
        index = self.action_list.index(selection[0])
        new_index = index + direction
        if new_index < 0 or new_index >= len(self.actions):
            return
        self.actions[index], self.actions[new_index] = self.actions[new_index], self.actions[index]
        self.sync_current_script()
        self.refresh_action_list()
        self.action_list.selection_set(self.action_list.get_children()[new_index])

    def clear_actions(self):
        if not self.actions:
            return
        confirmed = messagebox.askyesno(
            "Hapus Semua Aksi",
            "Hapus semua aksi pada script yang sedang dipakai?",
            parent=self,
        )
        if not confirmed:
            return
        self.actions.clear()
        self.sync_current_script()
        self.refresh_action_list()
        self.set_status("Status: Semua aksi dihapus")

    def show_action_menu(self, event):
        if not self.actions:
            return

        item_id = self.action_list.identify_row(event.y)
        if not item_id:
            return

        index = self.action_list.index(item_id)
        if index < 0 or index >= len(self.actions):
            return

        self.action_list.selection_set(item_id)
        self.action_list.focus(item_id)

        self.action_menu.entryconfig("Naikkan", state=tk.NORMAL if index > 0 else tk.DISABLED)
        self.action_menu.entryconfig("Turunkan", state=tk.NORMAL if index < len(self.actions) - 1 else tk.DISABLED)
        try:
            self.action_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.action_menu.grab_release()
        return "break"

    def refresh_action_list(self):
        self.action_list.delete(*self.action_list.get_children())
        self.action_label.configure(text=f"Daftar Aksi - {self.current_script.get()}")
        for index, action in enumerate(self.actions, start=1):
            delay = action.get("delay_ms", self.delay_var.get())
            if action["type"] == "click":
                text = f"  Click  X:{action['x']}  Y:{action['y']}    Delay {delay} ms"
            else:
                text = f"  Key  '{action['key']}'    Delay {delay} ms"
            self.action_list.insert("", "end", values=(index, text))
        if not self.actions:
            self.action_list.insert("", "end", values=("", "  (Belum ada aksi untuk script ini)"), tags=("placeholder",))
            self.action_list.tag_configure("placeholder", foreground=self.colors["muted"])
        self.update_start_button()

    def start_pick_coordinate(self):
        if self.picking_coordinate:
            return
        self.picking_coordinate = True
        self.set_status("Status: Klik posisi target di layar...")
        self.withdraw()

        def on_click(x, y, _button, pressed):
            if pressed:
                self.x_var.set(int(x))
                self.y_var.set(int(y))
                self.picking_coordinate = False
                self.after(0, self.deiconify)
                self.after(0, lambda: self.set_status(f"Status: Koordinat dipilih X:{int(x)} Y:{int(y)}"))
                return False
            return True

        pynput_mouse.Listener(on_click=on_click).start()

    def set_hotkey(self):
        value = self.hotkey.get().strip()
        if not value:
            self.hotkey.set(DEFAULT_HOTKEY)
        self.sync_current_script()
        self.set_status(f"Status: Hotkey diset ke '{self.hotkey.get()}'")

    def on_global_key_press(self, key):
        try:
            pressed = key.char
        except AttributeError:
            pressed = str(key).replace("Key.", "")
        if pressed == self.hotkey.get().strip():
            self.after(0, self.toggle_running)

    def toggle_running(self):
        if self.is_running:
            self.stop_clicker()
        else:
            self.start_clicker()

    def start_clicker(self):
        if self.is_running:
            return
        if not self.actions:
            messagebox.showwarning("Daftar aksi kosong", "Tambahkan click atau key lebih dulu.", parent=self)
            return
        self.sync_current_script()
        self.is_running = True
        self.stop_event.clear()
        self.worker_thread = threading.Thread(target=self.run_actions, daemon=True)
        self.worker_thread.start()
        self.update_start_button()
        if self.repeat_var.get() == 0:
            self.set_status("Status: Berjalan tanpa batas")
        else:
            self.set_status("Status: Berjalan")

    def stop_clicker(self):
        self.stop_event.set()
        self.is_running = False
        self.update_start_button()
        self.set_status("Status: Berhenti")

    def run_actions(self):
        repeat = max(0, self.repeat_var.get())
        actions = list(self.actions)
        try:
            if repeat == 0:
                while not self.stop_event.is_set():
                    self.run_action_cycle(actions)
            else:
                for _ in range(repeat):
                    if self.stop_event.is_set():
                        break
                    self.run_action_cycle(actions)
        finally:
            self.after(0, self.finish_run)

    def run_action_cycle(self, actions):
        for action in actions:
            if self.stop_event.is_set():
                break
            if action["type"] == "click":
                self.mouse_controller.position = self.randomize_click_position(action["x"], action["y"])
                self.mouse_controller.click(pynput_mouse.Button.left)
            elif action["type"] == "key":
                self.tap_key(action["key"])
            self.sleep_ms(self.randomize_delay_ms(action.get("delay_ms", self.delay_var.get())))

    def tap_key(self, key_name):
        if len(key_name) == 1:
            self.keyboard_controller.press(key_name)
            self.keyboard_controller.release(key_name)
            return

        special_key = getattr(pynput_keyboard.Key, key_name.lower(), None)
        if special_key is None:
            for char in key_name:
                self.keyboard_controller.press(char)
                self.keyboard_controller.release(char)
            return
        self.keyboard_controller.press(special_key)
        self.keyboard_controller.release(special_key)

    def sleep_ms(self, milliseconds):
        end_time = time.monotonic() + (max(1, milliseconds) / 1000)
        while not self.stop_event.is_set() and time.monotonic() < end_time:
            time.sleep(min(0.03, end_time - time.monotonic()))

    def randomize_click_position(self, x, y):
        return (
            x + random.randint(-CLICK_POSITION_JITTER, CLICK_POSITION_JITTER),
            y + random.randint(-CLICK_POSITION_JITTER, CLICK_POSITION_JITTER),
        )

    def randomize_delay_ms(self, base_delay):
        return max(1, base_delay + random.randint(-DELAY_JITTER_MS, DELAY_JITTER_MS))

    def finish_run(self):
        self.is_running = False
        self.update_start_button()
        if self.stop_event.is_set():
            self.set_status("Status: Berhenti")
        else:
            self.set_status("Status: Selesai")

    def update_start_button(self):
        if self.is_running:
            self.start_button.configure(
                text="Stop",
                bg=self.colors["danger"],
                activebackground=self.colors["danger"],
            )
        else:
            self.start_button.configure(
                text="Mulai",
                bg=self.colors["button_primary"],
                activebackground=self.colors["button_primary_active"],
            )

    def set_status(self, text):
        self.status.configure(text=text)

    def close_app(self):
        self.stop_clicker()
        try:
            self.hotkey_listener.stop()
        except RuntimeError:
            pass
        self.destroy()


if __name__ == "__main__":
    app = KyronApp()
    app.mainloop()
