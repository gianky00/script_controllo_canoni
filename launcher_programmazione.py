import tkinter as tk
from tkinter import font, messagebox, scrolledtext, filedialog
from tkinter import ttk
import subprocess
import threading
import queue
import sys
import os
import json
import sqlite3
from datetime import datetime, timedelta
import uuid
import time
from collections import deque

# --- CONTROLLO DIPENDENZE ---
try:
    from PIL import Image, ImageDraw, ImageTk
except ImportError:
    messagebox.showerror("Libreria Mancante", "Pillow non è installata. Esegui: pip install Pillow")
    sys.exit(1)
try:
    import schedule
except ImportError:
    messagebox.showerror("Libreria Mancante", "schedule non è installata. Esegui: pip install schedule")
    sys.exit(1)
try:
    from plyer import notification
except ImportError:
    messagebox.showerror("Libreria Mancante", "plyer non è installata. Esegui: pip install plyer")
    sys.exit(1)
try:
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
except ImportError:
    messagebox.showerror("Libreria Mancante", "matplotlib non è installata. Esegui: pip install matplotlib")
    sys.exit(1)

# --- CLASSI WIDGET PERSONALIZZATI ---
class ModernButton(tk.Frame):
    def __init__(self, parent, text, icon_char, command, base_color, hover_color, press_color, text_color="white", width=230, height=50, radius=25):
        parent_bg = ""
        try:
            parent_bg = parent.cget("bg")
        except tk.TclError:
            try:
                parent_bg = parent.master.cget("bg")
            except (AttributeError, tk.TclError):
                parent_bg = "#F0F0F0"

        super().__init__(parent, bg=parent_bg)
        self.command = command
        self.width = width
        self.height = height
        self.radius = radius
        self.colors = {"normal": (base_color, base_color), "hover": (hover_color, base_color), "press": (press_color, press_color), "disabled": ("#BDBDBD", "#BDBDBD")}
        self.text_color = text_color
        self.disabled_text_color = "#757575"
        self.is_enabled = True
        self.icon_font = font.Font(family="Segoe UI Symbol", size=14)
        self.text_font = font.Font(family="Segoe UI", size=11, weight="bold")
        self.text = text
        self.icon_char = icon_char
        self.canvas = tk.Canvas(self, width=self.width, height=self.height, bg=parent_bg, bd=0, highlightthickness=0)
        self.canvas.pack()
        self.draw_button("normal")
        self.bind_events()

    def bind_events(self):
        self.canvas.bind("<Enter>", self.on_enter)
        self.canvas.bind("<Leave>", self.on_leave)
        self.canvas.bind("<Button-1>", self.on_click)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)

    def draw_button(self, state):
        self.canvas.delete("all")
        fill_color, border_color = self.colors.get(state, self.colors["normal"])
        image = Image.new("RGBA", (self.width * 3, self.height * 3), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        draw.rounded_rectangle((1, 1, self.width * 3 - 1, self.height * 3 - 1), radius=self.radius * 3, fill=fill_color, outline=border_color, width=2)
        self.photo_image = ImageTk.PhotoImage(image.resize((self.width, self.height), Image.Resampling.LANCZOS))
        self.canvas.create_image(0, 0, image=self.photo_image, anchor="nw")

        content_parts = []
        if self.icon_char:
            content_parts.append(self.icon_char)
        if self.text:
            content_parts.append(self.text)

        spacing = 8
        total_width = sum(self.text_font.measure(p) if p == self.text else self.icon_font.measure(p) for p in content_parts)
        if len(content_parts) > 1:
            total_width += spacing * (len(content_parts) - 1)

        start_x = (self.width - total_width) / 2
        current_x = start_x
        current_text_color = self.text_color if self.is_enabled else self.disabled_text_color

        for part in content_parts:
            is_icon = part == self.icon_char
            font_to_use = self.icon_font if is_icon else self.text_font
            self.canvas.create_text(current_x, self.height/2, text=part, font=font_to_use, fill=current_text_color, anchor="w")
            current_x += font_to_use.measure(part) + spacing


    def on_enter(self, event=None):
        # pylint: disable=unused-argument
        if self.is_enabled: self.draw_button("hover"); self.config(cursor="hand2")

    def on_leave(self, event=None):
        # pylint: disable=unused-argument
        if self.is_enabled: self.draw_button("normal"); self.config(cursor="")

    def on_click(self, event=None):
        # pylint: disable=unused-argument
        if self.is_enabled: self.draw_button("press")

    def on_release(self, event):
        if self.is_enabled:
            self.draw_button("hover")
            if 0 <= event.x < self.width and 0 <= event.y < self.height: self.command()

    def set_enabled(self, enabled):
        self.is_enabled = enabled; self.draw_button("normal" if enabled else "disabled"); self.config(cursor="hand2" if enabled else "")

    def update_style(self, base_color, hover_color, press_color):
        self.colors["normal"] = (base_color, base_color); self.colors["hover"] = (hover_color, base_color); self.colors["press"] = (press_color, press_color); self.draw_button("normal")

class ToggleButton(tk.Canvas):
    def __init__(self, parent, text, app_theme, command=None, width=60, height=35):
        super().__init__(parent, width=width, height=height, bg=parent.cget("bg"), bd=0, highlightthickness=0)
        self.app_theme = app_theme; self.text = text; self.is_selected = tk.BooleanVar(value=False)
        self.width=width; self.height=height; self.command = command
        self.colors = {"normal_bg": "#E5E5EA", "normal_fg": "#4A4A4A", "selected_bg": app_theme.COLOR_ACCENT, "selected_fg": "white"}
        self.bind("<Button-1>", self.toggle); self.config(cursor="hand2"); self.draw()

    def draw(self):
        self.delete("all"); bg = self.colors['selected_bg'] if self.is_selected.get() else self.colors['normal_bg']; fg = self.colors['selected_fg'] if self.is_selected.get() else self.colors['normal_fg']
        radius = 18; image = Image.new("RGBA", (self.width * 3, self.height * 3), (0, 0, 0, 0)); draw = ImageDraw.Draw(image)
        draw.rounded_rectangle((0, 0, self.width * 3, self.height * 3), radius=radius, fill=bg); self.photo_image = ImageTk.PhotoImage(image.resize((self.width, self.height), Image.Resampling.LANCZOS))
        self.create_image(0, 0, image=self.photo_image, anchor="nw"); self.create_text(self.width/2, self.height/2, text=self.text, font=("Segoe UI", 10, "bold"), fill=fg)

    def toggle(self, event=None):
        # pylint: disable=unused-argument
        self.is_selected.set(not self.is_selected.get()); self.draw()
        if self.command: self.command()

class TaskDialog(tk.Toplevel):
    def __init__(self, parent_app, task=None):
        super().__init__(parent_app.master)
        self.parent_app = parent_app; self.task = task; self.result = None
        self.schedule_data = task.get("schedule_data", {}).copy() if task else {}
        self.current_day_key = "0"; self.title("Definisci Task" if not task else "Modifica Task"); self.configure(bg=parent_app.COLOR_FRAME)

        self.geometry("850x650")

        self.resizable(False, False); self.transient(parent_app.master); self.grab_set()
        main_notebook = ttk.Notebook(self); main_notebook.pack(expand=True, fill="both", padx=10, pady=10)
        tab1 = ttk.Frame(main_notebook); main_notebook.add(tab1, text="Generale e Pianificazione"); self._create_general_tab(tab1)
        tab2 = ttk.Frame(main_notebook); main_notebook.add(tab2, text="Errori e Notifiche"); self._create_error_handling_tab(tab2)
        button_bar = tk.Frame(self, bg=self.parent_app.COLOR_BG, pady=10); button_bar.pack(side="bottom", fill="x")
        ModernButton(button_bar, "Annulla", "✕", self.destroy, self.parent_app.COLOR_SYSTEM, self.parent_app.COLOR_SYSTEM_HOVER, self.parent_app.COLOR_SYSTEM_PRESS, width=150, height=40, radius=20).pack(side="right", padx=20)
        ModernButton(button_bar, "Salva Task", "✔", self._on_save, self.parent_app.COLOR_SUCCESS, self.parent_app.COLOR_SUCCESS_HOVER, self.parent_app.COLOR_SUCCESS_PRESS, width=150, height=40, radius=20).pack(side="right")
        if self.task: self._populate_fields()
        
        first_day_to_show = next(iter(sorted(self.schedule_data.keys())), "0")
        self._on_day_selected(first_day_to_show)

    def _create_general_tab(self, parent):
        content_frame = tk.Frame(parent, bg=self.parent_app.COLOR_FRAME, padx=15, pady=15); content_frame.pack(expand=True, fill="both"); content_frame.grid_columnconfigure(1, weight=1)
        tk.Label(content_frame, text="Nome Task:", bg=self.parent_app.COLOR_FRAME, font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w", pady=5)
        self.name_entry = ttk.Entry(content_frame, font=("Segoe UI", 10)); self.name_entry.grid(row=0, column=1, columnspan=2, sticky="ew", pady=5, padx=5)
        tk.Label(content_frame, text="Percorso Script:", bg=self.parent_app.COLOR_FRAME, font=("Segoe UI", 10)).grid(row=1, column=0, sticky="w", pady=5)
        script_frame = tk.Frame(content_frame, bg=self.parent_app.COLOR_FRAME); script_frame.grid(row=1, column=1, columnspan=2, sticky="ew", pady=5, padx=5)
        self.script_entry = ttk.Entry(script_frame, font=("Segoe UI", 10)); self.script_entry.pack(side="left", expand=True, fill="x"); ttk.Button(script_frame, text="Sfoglia...", command=self._browse_script).pack(side="left", padx=(5,0))
        ttk.Separator(content_frame, orient="horizontal").grid(row=2, column=0, columnspan=3, sticky="ew", pady=20)
        tk.Label(content_frame, text="Giorni di Esecuzione:", bg=self.parent_app.COLOR_FRAME, font=("Segoe UI", 10, "bold")).grid(row=3, column=0, sticky="w", pady=5)
        self.days_frame = tk.Frame(content_frame, bg=self.parent_app.COLOR_FRAME); self.days_frame.grid(row=3, column=1, columnspan=2, sticky="ew", pady=10, padx=5)
        self.day_buttons = {}; days = {"0":"LUN", "1":"MAR", "2":"MER", "3":"GIO", "4":"VEN", "5":"SAB", "6":"DOM"}

        for day_key, day_text in days.items():
            btn = ToggleButton(self.days_frame, text=day_text, app_theme=self.parent_app, command=lambda k=day_key: self._on_day_selected(k))
            btn.pack(side="left", padx=4)
            self.day_buttons[day_key] = btn

        tk.Label(content_frame, text="Orari:", bg=self.parent_app.COLOR_FRAME, font=("Segoe UI", 10, "bold")).grid(row=4, column=0, sticky="nw", pady=(15,5))
        self.times_frame = tk.Frame(content_frame, bg=self.parent_app.COLOR_BG, bd=1, relief="sunken"); self.times_frame.grid(row=4, column=1, sticky="nsew", rowspan=3, pady=10, padx=5); content_frame.grid_rowconfigure(4, weight=1)
        self.times_listbox = tk.Listbox(self.times_frame, font=("Consolas", 12), relief="flat", borderwidth=0); self.times_listbox.pack(expand=True, fill="both", padx=5, pady=5)

        time_entry_frame = tk.Frame(content_frame, bg=self.parent_app.COLOR_FRAME)
        time_entry_frame.grid(row=5, column=0, sticky="ew", pady=5, padx=5)
        tk.Label(time_entry_frame, text="Imposta e aggiungi un orario (HH:MM):", bg=self.parent_app.COLOR_FRAME, font=("Segoe UI", 9)).pack(anchor="w")
        time_spin_frame = tk.Frame(time_entry_frame, bg=self.parent_app.COLOR_FRAME); time_spin_frame.pack(fill="x", pady=(5,10))
        self.hour_spin = ttk.Spinbox(time_spin_frame, from_=0, to=23, wrap=True, width=3, font=("Segoe UI", 10), format="%02.0f"); self.hour_spin.pack(side="left", padx=(0,5))
        tk.Label(time_spin_frame, text=":", bg=self.parent_app.COLOR_FRAME, font=("Segoe UI", 12, "bold")).pack(side="left")
        self.minute_spin = ttk.Spinbox(time_spin_frame, from_=0, to=59, wrap=True, width=3, font=("Segoe UI", 10), format="%02.0f"); self.minute_spin.pack(side="left", padx=5)
        time_button_frame = tk.Frame(time_entry_frame, bg=self.parent_app.COLOR_FRAME); time_button_frame.pack(fill="x")
        add_time_btn = ttk.Button(time_button_frame, text="Aggiungi Orario", command=self._add_time, style="Accent.TButton"); add_time_btn.pack(fill="x", pady=(0, 5))
        del_time_btn = ttk.Button(time_button_frame, text="Rimuovi Selez.", command=self._remove_time); del_time_btn.pack(fill="x")
        self.hour_spin.bind("<Return>", self._add_time_event)
        self.minute_spin.bind("<Return>", self._add_time_event)

    def _create_error_handling_tab(self, parent):
        content_frame = tk.Frame(parent, bg=self.parent_app.COLOR_FRAME, padx=15, pady=15); content_frame.pack(expand=True, fill="both")
        retry_frame = ttk.LabelFrame(content_frame, text=" Riprova Automatica ", padding=10); retry_frame.pack(fill="x", expand=True, pady=(0,10))
        self.retry_enabled = tk.BooleanVar(); ttk.Checkbutton(retry_frame, text="Abilita riprova", variable=self.retry_enabled, command=self._toggle_retry_options).grid(row=0, column=0, columnspan=2, sticky="w")
        tk.Label(retry_frame, text="Tentativi:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.retry_count_spin = ttk.Spinbox(retry_frame, from_=1, to=10, width=5); self.retry_count_spin.grid(row=1, column=1, sticky="w", padx=5)
        tk.Label(retry_frame, text="Attesa (min):").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.retry_delay_spin = ttk.Spinbox(retry_frame, from_=1, to=60, width=5); self.retry_delay_spin.grid(row=2, column=1, sticky="w", padx=5)
        notify_frame = ttk.LabelFrame(content_frame, text=" Notifiche di Sistema ", padding=10); notify_frame.pack(fill="x", expand=True, pady=10)
        self.notify_on_success = tk.BooleanVar(); self.notify_on_failure = tk.BooleanVar()
        ttk.Checkbutton(notify_frame, text="Notifica in caso di successo", variable=self.notify_on_success).pack(anchor="w", pady=2)
        ttk.Checkbutton(notify_frame, text="Notifica in caso di fallimento", variable=self.notify_on_failure).pack(anchor="w", pady=2)
        self._toggle_retry_options()

    def _toggle_retry_options(self):
        state = "normal" if self.retry_enabled.get() else "disabled"; self.retry_count_spin.config(state=state); self.retry_delay_spin.config(state=state)

    def _populate_fields(self):
        self.name_entry.insert(0, self.task.get("name", "")); self.script_entry.insert(0, self.task.get("script_path", ""))
        for day_key in self.schedule_data:
            if day_key in self.day_buttons: self.day_buttons[day_key].is_selected.set(True); self.day_buttons[day_key].draw()
        retry_count = self.task.get("retry_count", 0)
        if retry_count > 0: self.retry_enabled.set(True); self.retry_count_spin.set(retry_count); self.retry_delay_spin.set(self.task.get("retry_delay", 5))
        self.notify_on_success.set(self.task.get("notify_on_success", False)); self.notify_on_failure.set(self.task.get("notify_on_failure", True)); self._toggle_retry_options()

    def _on_save(self):
        name = self.name_entry.get().strip(); script_path = self.script_entry.get().strip()
        if not name or not script_path: messagebox.showerror("Errore", "Nome e percorso script sono obbligatori.", parent=self); return
        self.schedule_data = {k: v for k, v in self.schedule_data.items() if v}
        if not self.schedule_data: messagebox.showerror("Errore", "Definire almeno un orario per un giorno.", parent=self); return
        self.result = {"id": self.task.get("id", str(uuid.uuid4())), "name": name, "script_path": script_path, "schedule_data": self.schedule_data,
                       "enabled": self.task.get('enabled', True), "retry_count": int(self.retry_count_spin.get()) if self.retry_enabled.get() else 0,
                       "retry_delay": int(self.retry_delay_spin.get()) if self.retry_enabled.get() else 5,
                       "notify_on_success": self.notify_on_success.get(), "notify_on_failure": self.notify_on_failure.get()}
        self.destroy()

    def _on_day_selected(self, day_key):
        is_selected = self.day_buttons[day_key].is_selected.get()
        if is_selected and day_key not in self.schedule_data: self.schedule_data[day_key] = []
        elif not is_selected and day_key in self.schedule_data: del self.schedule_data[day_key]
        self.current_day_key = day_key; self._refresh_times_list()

    def _refresh_times_list(self):
        self.times_listbox.delete(0, tk.END); day_times = sorted(self.schedule_data.get(self.current_day_key, [])); [self.times_listbox.insert(tk.END, t) for t in day_times]

    def _add_time_event(self, event=None):
        # pylint: disable=unused-argument
        self._add_time(); return "break"

    def _add_time(self):
        if not self.day_buttons[self.current_day_key].is_selected.get(): messagebox.showwarning("Giorno non selezionato", "Selezionare un giorno prima di aggiungere un orario.", parent=self); return
        new_time = f"{self.hour_spin.get()}:{self.minute_spin.get()}"
        if self.current_day_key not in self.schedule_data: self.schedule_data[self.current_day_key] = []
        if new_time not in self.schedule_data[self.current_day_key]: self.schedule_data[self.current_day_key].append(new_time); self._refresh_times_list()

    def _remove_time(self):
        if not (selection := self.times_listbox.curselection()): return
        selected_time = self.times_listbox.get(selection[0])
        if self.current_day_key in self.schedule_data and selected_time in self.schedule_data[self.current_day_key]:
            self.schedule_data[self.current_day_key].remove(selected_time)
            if not self.schedule_data[self.current_day_key]: del self.schedule_data[self.current_day_key]; self.day_buttons[self.current_day_key].toggle()
            self._refresh_times_list()

    def _browse_script(self):
        filetypes = [("Script Eseguibili", "*.py *.bat *.vbs"),("Tutti i file", "*.*")]; path = filedialog.askopenfilename(title="Seleziona script", filetypes=filetypes)
        if path: self.script_entry.delete(0, tk.END); self.script_entry.insert(0, path)

class HistoryWindow(tk.Toplevel):
    def __init__(self, parent_app, task):
        super().__init__(parent_app.master); self.parent_app = parent_app; self.task = task; self.title(f"Cronologia - {task['name']}"); self.configure(bg=parent_app.COLOR_FRAME); self.geometry("900x600"); self.transient(parent_app.master); self.grab_set()
        main_frame = tk.Frame(self, bg=parent_app.COLOR_FRAME, padx=10, pady=10); main_frame.pack(expand=True, fill="both")
        tk.Label(main_frame, text="Doppio click su una riga per vedere il log completo.", bg=parent_app.COLOR_FRAME).pack(anchor="w")
        cols = ("status", "start_time", "end_time", "duration", "exit_code"); self.tree = ttk.Treeview(main_frame, columns=cols, show="headings")
        for col, width in {"status": 60, "start_time": 150, "end_time": 150, "duration": 80, "exit_code": 80}.items(): self.tree.column(col, width=width, anchor='center' if col != 'duration' else 'e')
        for col, text in {"status": "Stato", "start_time": "Avvio", "end_time": "Fine", "duration": "Durata (s)", "exit_code": "Cod. Uscita"}.items(): self.tree.heading(col, text=text)
        self.tree.tag_configure("success", foreground="green"); self.tree.tag_configure("failure", foreground="red")
        for run in self.parent_app.db.get_runs_for_task(self.task['id']):
            status_icon = "✔️" if run['exit_code'] == 0 else "❌"; tag = "success" if run['exit_code'] == 0 else "failure"
            start_time = datetime.fromisoformat(run['start_time']).strftime('%d/%m/%Y %H:%M:%S'); end_time = datetime.fromisoformat(run['end_time']).strftime('%d/%m/%Y %H:%M:%S') if run['end_time'] else "N/A"
            duration = f"{run['duration_seconds']:.2f}" if run['duration_seconds'] is not None else "N/A"
            exit_code = run['exit_code'] if run['exit_code'] is not None else "N/A"
            self.tree.insert("", "end", iid=run['id'], values=(status_icon, start_time, end_time, duration, exit_code), tags=(tag,))
        self.tree.pack(expand=True, fill="both", pady=5); self.tree.bind("<Double-1>", self._on_double_click)

    def _on_double_click(self, event=None):
        # pylint: disable=unused-argument
        if selection := self.tree.selection(): self.show_log_details(selection[0])

    def show_log_details(self, run_id):
        log_content = self.parent_app.db.get_log_for_run(run_id); log_window = tk.Toplevel(self); log_window.title(f"Log Esecuzione #{run_id}"); log_window.geometry("800x500")
        log_text = scrolledtext.ScrolledText(log_window, wrap=tk.WORD, font=("Consolas", 10)); log_text.pack(expand=True, fill="both"); log_text.insert(tk.END, log_content); log_text.configure(state='disabled')

class DatabaseManager:
    # ... (nessuna modifica a questa classe, omessa per brevità)
    def __init__(self, db_file="scheduler.db"):
        self.db_file = db_file
        self._initialize_db()

    def _get_connection(self):
        conn = sqlite3.connect(self.db_file, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        return conn

    def _initialize_db(self):
        conn = self._get_connection()
        try:
            conn.execute('''CREATE TABLE IF NOT EXISTS tasks (id TEXT PRIMARY KEY, name TEXT NOT NULL, script_path TEXT NOT NULL, schedule_data TEXT NOT NULL, enabled INTEGER NOT NULL, retry_count INTEGER DEFAULT 0, retry_delay INTEGER DEFAULT 5, notify_on_success INTEGER DEFAULT 0, notify_on_failure INTEGER DEFAULT 1)''')
            conn.execute('''CREATE TABLE IF NOT EXISTS runs (id INTEGER PRIMARY KEY AUTOINCREMENT, task_id TEXT NOT NULL, start_time TEXT NOT NULL, end_time TEXT, duration_seconds REAL, exit_code INTEGER, log_output TEXT, FOREIGN KEY(task_id) REFERENCES tasks(id) ON DELETE CASCADE)''')
            conn.execute('''CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT NOT NULL)''')
            conn.execute("INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)", ('max_concurrent_runs', '3'))
            conn.commit()
        finally:
            conn.close()

    def get_tasks(self):
        conn = self._get_connection()
        try:
            cursor = conn.cursor(); cursor.execute("SELECT * FROM tasks"); tasks_data = cursor.fetchall()
            tasks = []
            for row in tasks_data:
                task = dict(row); task['schedule_data'] = json.loads(task['schedule_data']); task['enabled'] = bool(task['enabled']); tasks.append(task)
            return tasks
        finally:
            conn.close()

    def get_task_by_id(self, task_id):
        conn = self._get_connection()
        try:
            cursor = conn.cursor(); cursor.execute("SELECT * FROM tasks WHERE id = ?", (task_id,)); row = cursor.fetchone()
            if not row: return None
            task = dict(row); task['schedule_data'] = json.loads(task['schedule_data']); task['enabled'] = bool(task['enabled']); return task
        finally:
            conn.close()

    def save_task(self, task):
        conn = self._get_connection()
        try:
            task_data = (task['id'], task['name'], task['script_path'], json.dumps(task['schedule_data']), int(task['enabled']), task.get('retry_count', 0), task.get('retry_delay', 5), int(task.get('notify_on_success', 0)), int(task.get('notify_on_failure', 1)))
            conn.execute("INSERT OR REPLACE INTO tasks (id, name, script_path, schedule_data, enabled, retry_count, retry_delay, notify_on_success, notify_on_failure) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)", task_data)
            conn.commit()
        finally:
            conn.close()

    def delete_task(self, task_id):
        conn = self._get_connection()
        try:
            conn.execute("DELETE FROM tasks WHERE id = ?", (task_id,)); conn.commit()
        finally:
            conn.close()

    def create_run_record(self, task_id):
        conn = self._get_connection()
        try:
            start_time = datetime.now().isoformat(); cursor = conn.cursor(); cursor.execute("INSERT INTO runs (task_id, start_time) VALUES (?, ?)", (task_id, start_time)); conn.commit(); return cursor.lastrowid
        finally:
            conn.close()

    def finalize_run_record(self, run_id, exit_code, log_output):
        conn = self._get_connection()
        try:
            cursor = conn.cursor(); run = cursor.execute("SELECT start_time FROM runs WHERE id = ?", (run_id,)).fetchone()
            if run:
                end_time = datetime.now(); start_time = datetime.fromisoformat(run['start_time']); duration = (end_time - start_time).total_seconds()
                conn.execute("UPDATE runs SET end_time = ?, duration_seconds = ?, exit_code = ?, log_output = ? WHERE id = ?", (end_time.isoformat(), duration, exit_code, log_output, run_id)); conn.commit()
        finally:
            conn.close()

    def get_runs_for_task(self, task_id):
        conn = self._get_connection()
        try:
            cursor = conn.cursor(); cursor.execute("SELECT * FROM runs WHERE task_id = ? ORDER BY start_time DESC", (task_id,)); return [dict(row) for row in cursor.fetchall()]
        finally:
            conn.close()

    def get_log_for_run(self, run_id):
        conn = self._get_connection()
        try:
            cursor = conn.cursor(); cursor.execute("SELECT log_output FROM runs WHERE id = ?", (run_id,)); row = cursor.fetchone(); return row['log_output'] if row else "Log non trovato."
        finally:
            conn.close()

    def get_setting(self, key, default=None):
        conn = self._get_connection()
        try:
            cursor = conn.cursor(); cursor.execute("SELECT value FROM settings WHERE key = ?", (key,)); row = cursor.fetchone(); return row['value'] if row else default
        finally:
            conn.close()

    def save_setting(self, key, value):
        conn = self._get_connection()
        try:
            conn.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (key, str(value))); conn.commit()
        finally:
            conn.close()

    def get_run_stats_last_24h(self):
        conn = self._get_connection()
        try:
            twenty_four_hours_ago = (datetime.now() - timedelta(hours=24)).isoformat()
            cursor = conn.cursor(); cursor.execute("SELECT exit_code, COUNT(*) as count FROM runs WHERE start_time >= ? AND exit_code IS NOT NULL GROUP BY exit_code", (twenty_four_hours_ago,)); return {row['exit_code']: row['count'] for row in cursor.fetchall()}
        finally:
            conn.close()

    def get_hourly_activity(self):
        conn = self._get_connection()
        try:
            twenty_four_hours_ago = (datetime.now() - timedelta(hours=24)).isoformat(); cursor = conn.cursor(); cursor.execute("SELECT strftime('%H', start_time) as hour, COUNT(*) as count FROM runs WHERE start_time >= ? GROUP BY hour", (twenty_four_hours_ago,)); return {int(row['hour']): row['count'] for row in cursor.fetchall()}
        finally:
            conn.close()

    def get_top_long_running_tasks(self, limit=5):
        conn = self._get_connection()
        try:
            cursor = conn.cursor(); cursor.execute("SELECT T.name, AVG(R.duration_seconds) as avg_duration FROM runs R JOIN tasks T ON R.task_id = T.id WHERE R.duration_seconds IS NOT NULL GROUP BY T.name ORDER BY avg_duration DESC LIMIT ?", (limit,)); return cursor.fetchall()
        finally:
            conn.close()
            
class ScriptRunnerApp:
    def __init__(self, master):
        self.master = master; self.process = None; self.running_processes = {}; self.queue: queue.Queue = queue.Queue(); self.pending_queue: deque = deque(); self.base_title = "Gestore Attività SafeWork v10.0"; self.tasks = []
        self.db = DatabaseManager(); self.max_concurrent_runs = int(self.db.get_setting('max_concurrent_runs', 3))
        self.COLOR_BG="#F5F7FA";self.COLOR_FRAME="#FFFFFF";self.COLOR_TEXT="#4A4A4A";self.COLOR_TITLE_TEXT="#263238";self.COLOR_BORDER="#DDE2E7";self.COLOR_ACCENT="#007AFF";self.COLOR_ACCENT_HOVER="#409CFF";self.COLOR_ACCENT_PRESS="#005ECC";self.COLOR_SUCCESS="#34C759";self.COLOR_SUCCESS_HOVER="#5DE27C";self.COLOR_SUCCESS_PRESS="#248A3D";self.COLOR_WARNING="#FF9500";self.COLOR_WARNING_HOVER="#FFAC33";self.COLOR_WARNING_PRESS="#B86B00";self.COLOR_SYSTEM="#8E8E93";self.COLOR_SYSTEM_HOVER="#A4A4A8";self.COLOR_SYSTEM_PRESS="#6D6D72";self.COLOR_DANGER="#D32F2F";self.COLOR_DANGER_HOVER="#E54B4B";self.COLOR_DANGER_PRESS="#C12727"
        self.FONT_TITLE = ("Segoe UI Semibold", 12); self.FONT_CONSOLE = ("Consolas", 11)
        master.overrideredirect(True); width, height = 1200, 800; x_pos = (master.winfo_screenwidth()//2)-(width//2); y_pos = (master.winfo_screenheight()//2)-(height//2)
        master.geometry(f'{width}x{height}+{x_pos}+{y_pos}'); master.configure(bg=self.COLOR_BG); master.protocol("WM_DELETE_WINDOW", self._safe_close)
        main_container = tk.Frame(master, bg=self.COLOR_FRAME, bd=1, relief=tk.SOLID, highlightbackground=self.COLOR_BORDER); main_container.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        self._create_custom_title_bar(main_container); self._create_notebook(main_container); self._load_tasks()

    def _create_custom_title_bar(self, parent):
        self.title_bar = tk.Frame(parent, bg=self.COLOR_BG, height=40); self.title_bar.pack(expand=False, fill='x')
        self.title_label = tk.Label(self.title_bar, text=self.base_title, bg=self.COLOR_BG, fg=self.COLOR_TITLE_TEXT, font=self.FONT_TITLE); self.title_label.pack(side='left', padx=15, pady=10)
        close_button = tk.Label(self.title_bar, text='✕', bg=self.COLOR_BG, fg=self.COLOR_TEXT, font=("Arial", 16), bd=0); close_button.pack(side='right', padx=15)
        close_button.bind("<Button-1>", lambda e: self._safe_close()); close_button.bind("<Enter>", lambda e: e.widget.config(fg=self.COLOR_DANGER)); close_button.bind("<Leave>", lambda e: e.widget.config(fg=self.COLOR_TEXT))
        self.title_bar.bind("<Button-1>", self._start_move); self.title_bar.bind("<B1-Motion>", self._do_move); self.title_label.bind("<Button-1>", self._start_move); self.title_label.bind("<B1-Motion>", self._do_move)

    def _create_notebook(self, parent):
        style = ttk.Style(); style.configure('TNotebook', background=self.COLOR_FRAME, borderwidth=0)
        style.configure('TNotebook.Tab', background=self.COLOR_BG, foreground=self.COLOR_TEXT, font=("Segoe UI", 10, "bold"), padding=[20, 8], borderwidth=0); style.map('TNotebook.Tab', background=[('selected', self.COLOR_FRAME)], foreground=[('selected', self.COLOR_ACCENT)])
        style.configure("Accent.TButton", foreground="white", background=self.COLOR_ACCENT, font=("Segoe UI", 10, "bold"), padding=8); style.map("Accent.TButton", background=[('active', self.COLOR_ACCENT_HOVER), ('pressed', self.COLOR_ACCENT_PRESS)])
        self.notebook = ttk.Notebook(parent, style='TNotebook'); self.notebook.pack(expand=True, fill='both', padx=20, pady=(10, 20))
        launcher_tab, scheduler_tab, dashboard_tab, settings_tab = (tk.Frame(self.notebook, bg=self.COLOR_FRAME) for _ in range(4))
        self.notebook.add(launcher_tab, text=" 🚀 Launcher "); self.notebook.add(scheduler_tab, text=" 🗓️ Scheduler "); self.notebook.add(dashboard_tab, text=" 📊 Dashboard "); self.notebook.add(settings_tab, text=" ⚙️ Impostazioni ")
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_change)
        self._create_launcher_tab(launcher_tab); self._create_scheduler_tab(scheduler_tab); self._create_dashboard_tab(dashboard_tab); self._create_settings_tab(settings_tab)

    def _create_launcher_tab(self, parent_tab):
        button_frame = tk.Frame(parent_tab, bg=self.COLOR_FRAME); button_frame.pack(fill=tk.X, pady=(10, 25))
        self.btn_corrente = ModernButton(button_frame, "Prog. Corrente", "⏵", lambda: self.launch_script("rileva_programmazione_attuale.py"), self.COLOR_SUCCESS, self.COLOR_SUCCESS_HOVER, self.COLOR_SUCCESS_PRESS)
        self.btn_prossima = ModernButton(button_frame, "Prog. Prossima", "⏵", lambda: self.launch_script("rileva_programmazione_prossima.py"), self.COLOR_ACCENT, self.COLOR_ACCENT_HOVER, self.COLOR_ACCENT_PRESS)
        self.btn_flag_a3 = ModernButton(button_frame, "FLAG TCL A3", "⏵", lambda: self.launch_script("SafeWorkFlagA3.py"), self.COLOR_ACCENT, self.COLOR_ACCENT_HOVER, self.COLOR_ACCENT_PRESS)
        self.btn_arresto = ModernButton(button_frame, "Arresto", "⏹", self.arresta_script, self.COLOR_WARNING, self.COLOR_WARNING_HOVER, self.COLOR_WARNING_PRESS)
        self.btn_chiudi = ModernButton(button_frame, "Chiudi", "✕", self._safe_close, self.COLOR_SYSTEM, self.COLOR_SYSTEM_HOVER, self.COLOR_SYSTEM_PRESS)
        self.launch_buttons = [self.btn_corrente, self.btn_prossima, self.btn_flag_a3]
        for btn in self.launch_buttons: btn.pack(side=tk.LEFT, expand=True)
        self.btn_arresto.pack(side=tk.LEFT, expand=True); self.btn_chiudi.pack(side=tk.LEFT, expand=True); self.btn_arresto.set_enabled(False)
        self._create_console(parent_tab)

    def _create_scheduler_tab(self, parent_tab):
        main_scheduler_frame = tk.Frame(parent_tab, bg=self.COLOR_FRAME); main_scheduler_frame.pack(expand=True, fill='both'); tree_frame = tk.Frame(main_scheduler_frame); tree_frame.pack(expand=True, fill='both', pady=(0,10)); columns = ("status", "name", "script", "schedule", "next_run"); self.task_tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        self.task_tree.heading("status", text="Stato"); self.task_tree.column("status", width=60, anchor="center"); self.task_tree.heading("name", text="Nome Task"); self.task_tree.column("name", width=250); self.task_tree.heading("script", text="Script"); self.task_tree.column("script", width=350); self.task_tree.heading("schedule", text="Pianificazione"); self.task_tree.column("schedule", width=300); self.task_tree.heading("next_run", text="Prossima Esecuzione"); self.task_tree.column("next_run", width=180)
        self.task_tree.tag_configure("enabled", foreground="#222"); self.task_tree.tag_configure("disabled", foreground="gray"); tree_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.task_tree.yview); self.task_tree.configure(yscrollcommand=tree_scrollbar.set); tree_scrollbar.pack(side="right", fill="y"); self.task_tree.pack(side="left", expand=True, fill="both")
        self.task_tree.bind('<<TreeviewSelect>>', self._on_task_select); self.task_tree.bind('<Double-1>', self._on_tree_double_click)
        scheduler_buttons_frame = tk.Frame(main_scheduler_frame, bg=self.COLOR_FRAME); scheduler_buttons_frame.pack(fill='x', pady=(5, 5))
        self.btn_add_task = tk.Button(scheduler_buttons_frame, text="➕ Aggiungi Task", command=self._add_task, font=("Segoe UI", 10), relief="flat", bg="#EFEFEF", padx=10, pady=5); self.btn_add_task.pack(side="left", padx=5)
        self.btn_edit_task = tk.Button(scheduler_buttons_frame, text="✏️ Modifica", command=self._edit_task, state="disabled", font=("Segoe UI", 10), relief="flat", bg="#EFEFEF", padx=10, pady=5); self.btn_edit_task.pack(side="left", padx=5)
        self.btn_delete_task = tk.Button(scheduler_buttons_frame, text="🗑️ Elimina", command=self._delete_task, state="disabled", font=("Segoe UI", 10), relief="flat", bg="#EFEFEF", padx=10, pady=5); self.btn_delete_task.pack(side="left", padx=5)
        self.btn_toggle_task = tk.Button(scheduler_buttons_frame, text="⏯️ Pausa/Riprendi", command=self._toggle_task_state, state="disabled", font=("Segoe UI", 10), relief="flat", bg="#EFEFEF", padx=10, pady=5); self.btn_toggle_task.pack(side="left", padx=5)
        self.btn_history = tk.Button(scheduler_buttons_frame, text="📜 Cronologia", command=self._show_history, state="disabled", font=("Segoe UI", 10), relief="flat", bg="#EFEFEF", padx=10, pady=5); self.btn_history.pack(side="left", padx=5)

    def _create_dashboard_tab(self, parent_tab):
        self.dashboard_frame = tk.Frame(parent_tab, bg=self.COLOR_FRAME); self.dashboard_frame.pack(expand=True, fill='both')

    def _create_settings_tab(self, parent_tab):
        content_frame = tk.Frame(parent_tab, bg=self.COLOR_FRAME, padx=20, pady=20); content_frame.pack(fill='x', anchor='n'); settings_frame = ttk.LabelFrame(content_frame, text=" Impostazioni Generali ", padding=20); settings_frame.pack(fill='x')
        tk.Label(settings_frame, text="Esecuzioni simultanee massime:", font=("Segoe UI", 10)).pack(side="left", padx=5); self.concurrency_spin = ttk.Spinbox(settings_frame, from_=1, to=10, width=5, font=("Segoe UI", 10)); self.concurrency_spin.pack(side="left"); self.concurrency_spin.set(self.max_concurrent_runs)
        ModernButton(settings_frame, "Salva", "✔", self._save_settings, self.COLOR_ACCENT, self.COLOR_ACCENT_HOVER, self.COLOR_ACCENT_PRESS, width=120, height=40, radius=20).pack(side="left", padx=20)

    def _create_console(self, parent):
        console_container = tk.Frame(parent, bg=self.COLOR_BG, bd=0); console_container.pack(expand=True, fill=tk.BOTH, pady=(20,0))
        self.console_output = tk.Text(console_container, wrap=tk.WORD, font=self.FONT_CONSOLE, bg=self.COLOR_FRAME, fg=self.COLOR_TEXT, insertbackground=self.COLOR_TEXT, relief=tk.FLAT, borderwidth=0, selectbackground=self.COLOR_ACCENT_HOVER)
        style = ttk.Style(); style.theme_use('clam'); style.configure("Vertical.TScrollbar", gripcount=0, background=self.COLOR_FRAME, darkcolor=self.COLOR_BG, lightcolor=self.COLOR_BG, troughcolor=self.COLOR_BG, bordercolor=self.COLOR_BG, arrowcolor=self.COLOR_SYSTEM)
        scrollbar = ttk.Scrollbar(console_container, command=self.console_output.yview, style="Vertical.TScrollbar"); self.console_output['yscrollcommand'] = scrollbar.set; scrollbar.pack(side=tk.RIGHT, fill=tk.Y); self.console_output.pack(expand=True, fill=tk.BOTH, padx=1, pady=1)
        self.console_output.tag_config('ERROR', foreground="#D32F2F"); self.console_output.tag_config('INFO', foreground=self.COLOR_TEXT); self.console_output.tag_config('SUCCESS', foreground=self.COLOR_SUCCESS); self.console_output.configure(state='disabled')

    def _on_tab_change(self, event=None):
        # pylint: disable=unused-argument
        if self.notebook.index(self.notebook.select()) == 2: self.update_dashboard()

    def _start_move(self, event): self.x, self.y = event.x, event.y
    def _do_move(self, event): self.master.geometry(f"+{self.master.winfo_x() + event.x - self.x}+{self.master.winfo_y() + event.y - self.y}")

    def _stream_reader(self, stream, stream_name, log_accumulator: list[str]):
        for line in iter(stream.readline, ''): log_accumulator.append(line); self.queue.put((stream_name, line))
        stream.close()

    def _check_queue(self):
        while not self.queue.empty():
            stream_name, line = self.queue.get_nowait(); self.console_output.configure(state='normal'); tag = 'ERROR' if stream_name == 'stderr' else 'INFO'; self.console_output.insert(tk.END, line, tag); self.console_output.see(tk.END); self.console_output.configure(state='disabled')
        self.master.after(100, self._check_queue)

    def _check_manual_process(self):
        if self.process and self.process.poll() is not None: self._set_ui_state(idle=True); self.process = None
        self.master.after(500, self._check_manual_process)

    def _set_ui_state(self, idle: bool):
        for btn in self.launch_buttons: btn.set_enabled(idle)
        self.btn_arresto.set_enabled(not idle); self.title_label.config(text=self.base_title if idle else f"{self.running_script_name} - In Esecuzione...")
        if idle: self.btn_arresto.update_style(self.COLOR_WARNING, self.COLOR_WARNING_HOVER, self.COLOR_WARNING_PRESS)
        else: self.btn_arresto.update_style(self.COLOR_DANGER, self.COLOR_DANGER_HOVER, self.COLOR_DANGER_PRESS)

    def arresta_script(self, event=None):
        # pylint: disable=unused-argument
        if self.process and self.process.poll() is None: self.console_output.configure(state='normal'); self.console_output.insert(tk.END, "\n--- Richiesta di arresto inviata... ---\n", 'ERROR'); self.console_output.configure(state='disabled'); self.process.terminate()

    def _safe_close(self):
        if self.running_processes:
            if messagebox.askyesno("Processi in Esecuzione", f"{len(self.running_processes)} processi sono attivi. Chiudere comunque?"): [p['process'].terminate() for p in self.running_processes.values()]; self.master.destroy()
        else: self.master.destroy()

    # --- FIX: Modificato per integrare l'esecuzione manuale con la cronologia ---
    def launch_script(self, script_name):
        if self.process: messagebox.showwarning("Attenzione", "Un processo manuale è già in esecuzione."); return

        # Trova il task associato allo script per registrare la cronologia
        task_to_run = next((t for t in self.db.get_tasks() if t['script_path'].endswith(script_name.replace("\\", "/"))), None)
        if not task_to_run:
            messagebox.showerror("Errore", f"Nessun task trovato per lo script '{script_name}'.\nAssicurati che esista un task nello Scheduler con questo script per poter tracciare la cronologia.")
            return

        threading.Thread(target=self._manual_launch_worker, args=(task_to_run,), daemon=True).start()

    def _manual_launch_worker(self, task):
        self.running_script_name = task['name']
        self.master.after(0, lambda: self._set_ui_state(idle=False))
        self.master.after(0, lambda: self.console_output.config(state='normal'))
        self.master.after(0, lambda: self.console_output.delete('1.0', tk.END))
        self.master.after(0, lambda: self.console_output.config(state='disabled'))

        # Crea un record nel DB per l'esecuzione manuale
        run_id = self.db.create_run_record(task['id'])
        self.master.after(100, self._refresh_task_list)

        proc, log_accumulator = self._execute_script_once(task['script_path'])
        self.process = proc
        self.running_processes[proc.pid] = {"process": proc, "name": task['name'], "id": task['id']}
        proc.wait()
        
        # Finalizza il record nel DB
        log_output = "".join(log_accumulator)
        self.db.finalize_run_record(run_id, proc.returncode, log_output)

        if proc.pid in self.running_processes: del self.running_processes[proc.pid]
        self.master.after(0, lambda: self.console_output.configure(state='normal'))
        self.master.after(0, lambda: self.console_output.insert(tk.END, f"\n--- Processo terminato con codice {proc.returncode} ---\n", "SUCCESS" if proc.returncode == 0 else "ERROR"))
        self.master.after(0, lambda: self.console_output.configure(state='disabled'))
        self.master.after(100, self._refresh_task_list)


    def _task_execution_worker(self, task_id):
        task = self.db.get_task_by_id(task_id)
        if not task: return
        for attempt in range(1 + task.get('retry_count', 0)):
            run_id = self.db.create_run_record(task_id)
            self.master.after(100, self._refresh_task_list)
            proc, log_accumulator = self._execute_script_once(task['script_path'])
            self.running_processes[proc.pid] = {"process": proc, "name": task['name'], "id": task['id']}
            proc.wait()
            if proc.pid in self.running_processes: del self.running_processes[proc.pid]
            exit_code = proc.returncode; log_output = "".join(log_accumulator); self.db.finalize_run_record(run_id, exit_code, log_output)
            if exit_code == 0:
                if task.get('notify_on_success'): self._send_notification(f"✔️ Successo: {task['name']}", "Esecuzione completata."); break
            elif attempt < task.get('retry_count', 0): time.sleep(task.get('retry_delay', 5) * 60)
            else:
                if task.get('notify_on_failure'): self._send_notification(f"❌ Fallimento: {task['name']}", f"Fallito dopo {1 + task.get('retry_count', 0)} tentativi.")
        self.master.after(100, self._refresh_task_list)

    def _execute_script_once(self, script_path):
        _, extension = os.path.splitext(script_path.lower()); command = []
        if extension == '.py': command = [sys.executable, "-u", script_path]
        elif extension in ['.bat', '.cmd']: command = [script_path]
        elif extension == '.vbs': command = ['cscript.exe', script_path]
        startupinfo = None
        if os.name == 'nt': startupinfo = subprocess.STARTUPINFO(); startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        proc = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding='utf-8', errors='replace', startupinfo=startupinfo, bufsize=1, shell=extension in ['.bat', '.cmd'])
        log_accumulator: list[str] = []; threading.Thread(target=self._stream_reader, args=(proc.stdout, 'stdout', log_accumulator), daemon=True).start(); threading.Thread(target=self._stream_reader, args=(proc.stderr, 'stderr', log_accumulator), daemon=True).start()
        return proc, log_accumulator

    def _send_notification(self, title, message):
        try: notification.notify(title=title, message=message, app_name="Gestore Attività", timeout=10)
        except Exception as e: self.queue.put(('stderr', f"Errore notifica: {e}\n"))

    def _on_task_select(self, event=None):
        # pylint: disable=unused-argument
        state = "normal" if self.task_tree.selection() else "disabled"; self.btn_edit_task.config(state=state); self.btn_delete_task.config(state=state); self.btn_toggle_task.config(state=state); self.btn_history.config(state=state)

    def _on_tree_double_click(self, event):
        # pylint: disable=unused-argument
        if item_id := self.task_tree.identify_row(event.y): self.task_tree.selection_set(item_id); self._edit_task()

    def _refresh_task_list(self):
        selections = self.task_tree.selection(); [self.task_tree.delete(i) for i in self.task_tree.get_children()]; self.tasks = self.db.get_tasks(); days_map = ["Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom"]
        for task in self.tasks:
            status_icon = "🟢" if task["enabled"] else "⏸️"; schedule_str = "; ".join([f"{days_map[int(day)]} ({', '.join(sorted(times))})" for day, times in sorted(task.get("schedule_data", {}).items())]) or "N/A"
            job = next((j for j in schedule.get_jobs(task['id'])), None); next_run_str = job.next_run.strftime('%d/%m/%Y %H:%M') if job and job.next_run and task["enabled"] else "---"
            self.task_tree.insert("", "end", iid=task["id"], values=(status_icon, task["name"], task["script_path"], schedule_str, next_run_str), tags=("enabled" if task["enabled"] else "disabled",))
        if selections and self.task_tree.exists(selections[0]): self.task_tree.selection_set(selections[0])
        self._on_task_select()

    def _update_job_in_scheduler(self, task):
        schedule.clear(task["id"]); days_map = {0: schedule.every().monday, 1: schedule.every().tuesday, 2: schedule.every().wednesday, 3: schedule.every().thursday, 4: schedule.every().friday, 5: schedule.every().saturday, 6: schedule.every().sunday}
        if not task["enabled"]: return
        for day_key_str, times in task.get("schedule_data", {}).items():
            if (day_key := int(day_key_str)) in days_map:
                for t in times: days_map[day_key].at(t).do(lambda t_id=task['id']: self._queue_task(t_id)).tag(task["id"])

    def _add_task(self):
        dialog = TaskDialog(self); self.master.wait_window(dialog)
        if dialog.result: task = dialog.result; self.db.save_task(task); self._update_job_in_scheduler(task); self._refresh_task_list()

    def _edit_task(self):
        if not (selections := self.task_tree.selection()): return
        if not (task_to_edit := self.db.get_task_by_id(selections[0])): return
        dialog = TaskDialog(self, task=task_to_edit); self.master.wait_window(dialog)
        if dialog.result: new_task = dialog.result; self.db.save_task(new_task); self._update_job_in_scheduler(new_task); self._refresh_task_list()

    def _delete_task(self):
        if not (selections := self.task_tree.selection()): return
        if messagebox.askyesno("Conferma", "Eliminare il task selezionato e la sua cronologia?"): schedule.clear(selections[0]); self.db.delete_task(selections[0]); self._refresh_task_list()

    def _toggle_task_state(self):
        if not (selections := self.task_tree.selection()): return
        if task := self.db.get_task_by_id(selections[0]): task["enabled"] = not task["enabled"]; self.db.save_task(task); self._update_job_in_scheduler(task); self._refresh_task_list()

    def _show_history(self):
        if not (selections := self.task_tree.selection()): return
        if task := self.db.get_task_by_id(selections[0]): HistoryWindow(self, task)

    def _load_tasks(self):
        self.tasks = self.db.get_tasks(); [self._update_job_in_scheduler(t) for t in self.tasks]; self._refresh_task_list()

    def _run_scheduler(self):
        schedule.run_pending(); self.master.after(1000, self._run_scheduler)

    def _queue_task(self, task_id):
        running_task_ids = [p.get("id") for p in self.running_processes.values() if p.get("id")]
        if task_id not in self.pending_queue and task_id not in running_task_ids: self.pending_queue.append(task_id)

    def _process_pending_queue(self):
        if len(self.running_processes) < self.max_concurrent_runs and self.pending_queue:
            task_id = self.pending_queue.popleft(); threading.Thread(target=self._task_execution_worker, args=(task_id,), daemon=True).start()
        self.master.after(1000, self._process_pending_queue)

    def _save_settings(self):
        try:
            new_limit = int(self.concurrency_spin.get())
            if 1 <= new_limit <= 10: self.max_concurrent_runs = new_limit; self.db.save_setting('max_concurrent_runs', new_limit); messagebox.showinfo("Successo", "Impostazioni salvate.")
            else: messagebox.showerror("Errore", "Il limite deve essere un numero tra 1 e 10.")
        except ValueError: messagebox.showerror("Errore", "Il limite deve essere un numero intero.")
        except Exception as e: messagebox.showerror("Errore", f"Si è verificato un errore: {e}")

    def update_dashboard(self):
        for widget in self.dashboard_frame.winfo_children(): widget.destroy()
        try:
            fig = plt.Figure(figsize=(12, 7), dpi=100, facecolor=self.COLOR_FRAME, constrained_layout=True); gs = fig.add_gridspec(2, 2, hspace=0.5, wspace=0.3); ax1, ax2, ax3 = fig.add_subplot(gs[0, 0]), fig.add_subplot(gs[0, 1]), fig.add_subplot(gs[1, :])
            stats = self.db.get_run_stats_last_24h(); successes = stats.get(0, 0); failures = sum(v for k, v in stats.items() if k != 0)
            if successes > 0 or failures > 0: ax1.pie([successes, failures], labels=['Successi', 'Fallimenti'], autopct='%1.1f%%', colors=[self.COLOR_SUCCESS, self.COLOR_DANGER], startangle=90, textprops={'color': self.COLOR_TEXT})
            else: ax1.pie([1], labels=['Nessun dato'], colors=['#E0E0E0'], textprops={'color': self.COLOR_TEXT})
            ax1.set_title('Esiti ultime 24 ore', color=self.COLOR_TITLE_TEXT); activity = self.db.get_hourly_activity(); hours = list(range(24)); counts = [activity.get(h, 0) for h in hours]; ax2.bar(hours, counts, color=self.COLOR_ACCENT); ax2.set_title('Attività oraria (ultime 24h)', color=self.COLOR_TITLE_TEXT); ax2.set_xlabel('Ora del giorno', color=self.COLOR_TEXT); ax2.set_ylabel('N. Esecuzioni', color=self.COLOR_TEXT)
            top_tasks = self.db.get_top_long_running_tasks()
            if top_tasks: task_names = [row['name'] for row in top_tasks]; avg_durations = [row['avg_duration'] for row in top_tasks]; ax3.barh(task_names, avg_durations, color=self.COLOR_WARNING); ax3.invert_yaxis()
            ax3.set_title('Top 5 Task per Durata Media (s)', color=self.COLOR_TITLE_TEXT)
            for ax in [ax1, ax2, ax3]:
                ax.tick_params(colors=self.COLOR_TEXT)
                for spine in ax.spines.values(): spine.set_edgecolor(self.COLOR_BORDER)
            canvas = FigureCanvasTkAgg(fig, master=self.dashboard_frame); canvas.draw(); canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        except Exception as e: tk.Label(self.dashboard_frame, text=f"Errore dashboard:\n {e}", fg="red", bg=self.COLOR_FRAME, font=("Segoe UI", 12)).pack(expand=True)

    def start(self):
        """Avvia tutti i loop di monitoraggio e il mainloop dell'applicazione."""
        self._run_scheduler(); self._check_queue(); self._check_manual_process(); self._process_pending_queue(); self.master.mainloop()

if __name__ == "__main__":
    root = tk.Tk()
    app = ScriptRunnerApp(root)
    app.start()