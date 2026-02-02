import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
import subprocess
import threading
import queue
import sys
import re
import glob
from datetime import datetime, timedelta
import openpyxl

# Gestione importazione win32com e pythoncom per i thread
try:
    import win32com.client
    import pythoncom 
    PYWIN32_AVAILABLE = True
except ImportError:
    PYWIN32_AVAILABLE = False

CONFIG_FILE = "config_canoni.json"

class SettingsGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Controllo Canoni TS - Smart Config")
        self.root.geometry("1280x800") 
        
        self.config = self.load_config()
        self.log_queue = queue.Queue()
        
        # Variables
        self.selected_month = tk.StringVar()
        self.selected_year = tk.StringVar()
        self.account_var = tk.StringVar()
        self.progress_var = tk.DoubleVar()
        self.status_message = tk.StringVar(value="Pronto")
        self.process = None  # Processo script scaricamento
        
        style = ttk.Style()
        style.theme_use('clam')
        
        self.create_widgets()
        self.populate_fields() 
        self.setup_autosave()
        
        self.root.after(100, self.update_console)
        self.root.after(800, self.startup_sequence)

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        return {
            "account": "Manuale",
            "login_url": "https://portalefornitori.isab.com/Ui/",
            "username": "", "password": "", "download_dir": "C:\\Users\\Coemi\\Downloads",
            "move_dir": "", "giornaliera_path": "",
            "macro_file_path": "\\\\192.168.11.251\\Database_Tecnico_SMI\\MASTER FOGLI DI CALCOLO\\Comparatore_TS-Giornaliera (canoni).xlsm",
            "run_macro": False, "provider": "KK10608 - COEMI S.R.L.", "date_to_insert": "01.01.2025", "orders": [],
            "manual_consuntivi": {"MESSINA": "", "NASELLI": "", "CALDARELLA": "", "CALDARELLA 2": ""}
        }

    def save_config(self, show_msg=False):
        self.config.update({
            "account": self.account_var.get(),
            "login_url": self.login_url_entry.get(),
            "username": self.username_entry.get(),
            "password": self.password_entry.get(),
            "download_dir": self.download_dir_entry.get(),
            "move_dir": self.move_dir_entry.get(),
            "giornaliera_path": self.giornaliera_entry.get(),
            "macro_file_path": self.macro_path_entry.get(),
            "run_macro": self.run_macro_var.get(),
            "provider": self.provider_entry.get(),
            "date_to_insert": self.date_entry.get(),
            "orders": [{"numero": e[0].get().strip(), "posizione": e[1].get().strip(), "nome": e[2].cget("text")} for e in self.order_entries if e[0].get().strip()],
            "manual_consuntivi": {k: v.get().strip() for k, v in self.manual_inputs.items()}
        })
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self.config, f, indent=4)
        if show_msg: messagebox.showinfo("Successo", "Configurazione salvata!")

    def setup_autosave(self):
        widgets = [self.username_entry, self.password_entry, self.download_dir_entry, 
                   self.move_dir_entry, self.giornaliera_entry, self.login_url_entry, self.provider_entry, 
                   self.date_entry, self.macro_path_entry]
        for w in widgets:
            w.bind("<FocusOut>", lambda e: self.save_config())
            if isinstance(w, ttk.Combobox): w.bind("<<ComboboxSelected>>", lambda e: self.save_config())
        
        self.run_macro_var.trace_add("write", lambda *args: self.save_config())
        for num_ent, pos_ent, label_ent in self.order_entries:
            num_ent.bind("<FocusOut>", lambda e: self.save_config())
            pos_ent.bind("<FocusOut>", lambda e: self.save_config())

    def startup_sequence(self):
        self.log(">>> Avvio sequenza automatica...")
        self.import_from_giornaliera()

    def import_from_giornaliera(self):
        threading.Thread(target=self.import_from_giornaliera_thread, args=(False,), daemon=True).start()

    def import_from_giornaliera_thread(self, auto=False):
        try:
            path = self.giornaliera_entry.get()
            if not path or not os.path.exists(path): 
                self.log(f">>> File Giornaliera non trovato al percorso: {path}")
                return
            
            clean_path = os.path.normpath(path)
            self.log(f">>> Apertura Excel (engine=openpyxl): {clean_path}")
            
            # Utilizzo openpyxl in modalità read_only e data_only per leggere i valori calcolati
            wb = openpyxl.load_workbook(clean_path, data_only=True, read_only=True)
            try:
                # Verifica esistenza foglio
                if "RIEPILOGO" not in wb.sheetnames:
                    self.log(">>> Errore: Foglio 'RIEPILOGO' non trovato nel file.")
                    return
                    
                sheet = wb["RIEPILOGO"]
                cols = ["S", "U", "V", "W"]
                data, skipped_names = [], []
                seen_names = {}
                
                for i, col in enumerate(cols):
                    # In openpyxl read_only, l'accesso puntuale cella['A1'] potrebbe non essere ottimale 
                    # ma per poche celle va bene. sheet[f"{col}16"] restituisce la cella.
                    
                    c_name = sheet[f"{col}16"]
                    name_raw = c_name.value
                    base_name = str(name_raw).strip().upper() if name_raw else f"Colonna {col}"
                    
                    if base_name in seen_names:
                        seen_names[base_name] += 1
                        name = f"{base_name} {seen_names[base_name]}"
                    else:
                        seen_names[base_name] = 1
                        name = base_name

                    c_status = sheet[f"{col}19"]
                    status = str(c_status.value if c_status.value is not None else "").strip().upper()
                    
                    c_val = sheet[f"{col}17"]
                    val = c_val.value
                    
                    if status == "ABILITATO" and val:
                        clean = re.sub(r'\D', '', str(val))
                        if clean: 
                            data.append({"val": clean, "nome": name})
                            self.log(f"    - {name}: Trovato OdA {clean}")
                        else:
                            self.log(f"    - {name}: Nessun numero OdA valido")
                    else:
                        if status != "ABILITATO": 
                            skipped_names.append(name)
                            self.log(f"    - {name}: DISABILITATO")
                            
                    self.progress_var.set(20 + (i+1)/len(cols)*70)
                
                self.root.after(0, lambda: self.update_orders_gui(data, skipped_names, auto))
            finally:
                wb.close()
        except Exception as e: 
            self.log(f">>> Errore Lettura Excel: {e}")
            import traceback
            self.log(traceback.format_exc())

    def update_orders_gui(self, data, skipped, auto):
        # Reset campi
        for e in self.order_entries: 
            e[0].delete(0, tk.END); e[1].delete(0, tk.END); e[2].config(text="")
        
        # Popolamento
        for i, item in enumerate(data):
            if i < len(self.order_entries):
                self.order_entries[i][0].insert(0, item["val"])
                self.order_entries[i][1].insert(0, "10")
                self.order_entries[i][2].config(text=item["nome"])
        
        self.save_config()
        if skipped:
            self.status_message.set(f"Attenzione: OdC per {', '.join(skipped)} non abilitato.")
            self.status_label.config(foreground="#d9534f")
        else:
            self.status_message.set(f"Importazione completata: {len(data)} ordini caricati.")
            self.status_label.config(foreground="#5cb85c")

    def calculate_dynamic_paths(self, auto=False, full_update=False):
        months = ["GENNAIO", "FEBBRAIO", "MARZO", "APRILE", "MAGGIO", "GIUGNO", "LUGLIO", "AGOSTO", "SETTEMBRE", "OTTOBRE", "NOVEMBRE", "DICEMBRE"]
        if auto:
            target = datetime.now().replace(day=1) - timedelta(days=1)
            y, m = target.year, target.month
            self.selected_year.set(str(y))
            self.selected_month.set(months[m-1])
        else:
            try: 
                m = months.index(self.selected_month.get()) + 1
                y = int(self.selected_year.get())
            except: return

        m_str, m_name = f"{m:02d}", months[m-1]
        self.move_dir_entry.delete(0, tk.END); self.move_dir_entry.insert(0, f"\\\\192.168.11.251\\Condivisa\\ALLEGRETTI\\{y}\\TS\\CANONI\\{m_str} - {m_name}")
        self.giornaliera_entry.delete(0, tk.END); self.giornaliera_entry.insert(0, f"\\\\192.168.11.251\\Database_Tecnico_SMI\\Giornaliere\\Giornaliere {y}\\Giornaliera {m_str}-{y}.xlsm")
        self.date_entry.delete(0, tk.END); self.date_entry.insert(0, f"01.{m_str}.{y}")
        self.save_config()
        self.log(f">>> Percorsi impostati: {m_name} {y}")

        if full_update:
            self.import_from_giornaliera()
            self.preview_macro_params()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True)
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.settings_tab = ttk.Frame(self.notebook)
        self.log_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_tab, text=" Configurazione ")
        self.notebook.add(self.log_tab, text=" Log Operazioni ")

        canvas = tk.Canvas(self.settings_tab)
        scroll = ttk.Scrollbar(self.settings_tab, orient="vertical", command=canvas.yview)
        self.scroll_frame = ttk.Frame(canvas)
        self.scroll_frame.columnconfigure(0, weight=3)
        self.scroll_frame.columnconfigure(1, weight=2)
        
        canvas_window = canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        
        def configure_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        def configure_window_width(event):
            canvas.itemconfig(canvas_window, width=event.width)

        self.scroll_frame.bind("<Configure>", configure_scroll_region)
        canvas.bind("<Configure>", configure_window_width)

        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        # --- COLONNA SINISTRA ---
        left_col = ttk.Frame(self.scroll_frame)
        left_col.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        g1 = ttk.LabelFrame(left_col, text=" Parametri Base e Percorsi ", padding=10)
        g1.pack(fill=tk.X, pady=5)
        
        f_per = ttk.Frame(g1)
        f_per.pack(fill=tk.X, pady=(0,10))
        self.month_combo = ttk.Combobox(f_per, values=["GENNAIO", "FEBBRAIO", "MARZO", "APRILE", "MAGGIO", "GIUGNO", "LUGLIO", "AGOSTO", "SETTEMBRE", "OTTOBRE", "NOVEMBRE", "DICEMBRE"], textvariable=self.selected_month, width=15, state="readonly")
        self.month_combo.pack(side=tk.LEFT, padx=2)
        self.month_combo.bind("<<ComboboxSelected>>", lambda e: self.calculate_dynamic_paths(False))
        self.year_combo = ttk.Combobox(f_per, values=[str(datetime.now().year+i) for i in range(-1, 2)], textvariable=self.selected_year, width=8, state="readonly")
        self.year_combo.pack(side=tk.LEFT, padx=2)
        self.year_combo.bind("<<ComboboxSelected>>", lambda e: self.calculate_dynamic_paths(False))
        ttk.Button(f_per, text="Ricalcola Percorsi", command=lambda: self.calculate_dynamic_paths(False, full_update=True)).pack(side=tk.LEFT, padx=5)

        def add_grid_row(parent_frame, row_idx, label_text, attr_name, browse=True, show_char=None):
            parent_frame.columnconfigure(1, weight=1)
            ttk.Label(parent_frame, text=label_text, width=15, anchor="w").grid(row=row_idx, column=0, sticky="w", pady=2)
            if attr_name == "account_combo":
                c = ttk.Combobox(parent_frame, textvariable=self.account_var, values=["Manuale", "TRICHINI", "GIGLIUTO"], state="readonly")
                c.grid(row=row_idx, column=1, sticky="ew", padx=2, pady=2)
                c.bind("<<ComboboxSelected>>", self.on_account_change)
                self.account_combo = c
            else:
                ent = ttk.Entry(parent_frame)
                if show_char: ent.config(show=show_char)
                ent.grid(row=row_idx, column=1, sticky="ew", padx=2, pady=2)
                setattr(self, attr_name, ent)
                if browse:
                    cmd = lambda: (self.browse_dir(ent) if "dir" in attr_name else self.browse_file(ent))
                    ttk.Button(parent_frame, text="..", width=3, command=cmd).grid(row=row_idx, column=2, padx=2, pady=2)

        f_fields = ttk.Frame(g1)
        f_fields.pack(fill=tk.X)
        add_grid_row(f_fields, 0, "Account:", "account_combo", False)
        add_grid_row(f_fields, 1, "Username:", "username_entry", False)
        add_grid_row(f_fields, 2, "Password:", "password_entry", False, "*")
        add_grid_row(f_fields, 3, "Download Dir:", "download_dir_entry")
        add_grid_row(f_fields, 4, "Sposta in:", "move_dir_entry")
        add_grid_row(f_fields, 5, "Giornaliera:", "giornaliera_entry")

        f_bar = ttk.Frame(g1); f_bar.pack(fill=tk.X, pady=10)
        ttk.Progressbar(f_bar, variable=self.progress_var, maximum=100).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        # ttk.Button(f_bar, text="Importa OdA", command=self.import_from_giornaliera).pack(side=tk.RIGHT)

        g2 = ttk.LabelFrame(left_col, text=" Parametri Portale ", padding=10)
        g2.pack(fill=tk.X, pady=5)
        f_port = ttk.Frame(g2); f_port.pack(fill=tk.X)
        add_grid_row(f_port, 0, "Login URL:", "login_url_entry", False)
        add_grid_row(f_port, 1, "Fornitore:", "provider_entry", False)
        add_grid_row(f_port, 2, "Data (Inizio):", "date_entry", False)

        g3 = ttk.LabelFrame(left_col, text=" Automazione Macro ", padding=10)
        g3.pack(fill=tk.X, pady=5)
        self.run_macro_var = tk.BooleanVar()
        ttk.Checkbutton(g3, text="Esegui Macro Excel al termine", variable=self.run_macro_var).pack(anchor="w", pady=(0,5))
        f_mac = ttk.Frame(g3); f_mac.pack(fill=tk.X)
        add_grid_row(f_mac, 0, "File Macro:", "macro_path_entry")

        # --- SEZIONE PARAMETRI CONSUNTIVI (MANUALI) ---
        g_cons = ttk.LabelFrame(left_col, text=" Parametri Consuntivi & Anteprima ", padding=10)
        g_cons.pack(fill=tk.X, pady=5)
        
        self.manual_inputs = {}
        consuntivi_keys = ["MESSINA", "NASELLI", "CALDARELLA", "CALDARELLA 2"]
        
        for idx, key in enumerate(consuntivi_keys):
            f_row = ttk.Frame(g_cons)
            f_row.pack(fill=tk.X, pady=2)
            ttk.Label(f_row, text=key, width=15, anchor="w").pack(side=tk.LEFT)
            en = ttk.Entry(f_row, width=10)
            en.pack(side=tk.LEFT, padx=5)
            self.manual_inputs[key] = en
            
            if key == "CALDARELLA 2":
                ttk.Label(f_row, text="(MANUALE)", font=("Arial", 8, "bold"), foreground="red").pack(side=tk.LEFT)

            # Bind salvataggio
            en.bind("<FocusOut>", lambda e: self.save_config())

        # ttk.Button(g_cons, text="ANTEPRIMA PARAMETRI MACRO", command=self.preview_macro_params).pack(fill=tk.X, pady=(10, 0))

        # --- COLONNA DESTRA (OdA con Intestazioni) ---
        right_col = ttk.Frame(self.scroll_frame)
        right_col.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        
        g4 = ttk.LabelFrame(right_col, text=" Elenco Ordini (OdA) ", padding=10)
        g4.pack(fill=tk.BOTH, expand=True)
        
        self.status_label = ttk.Label(g4, textvariable=self.status_message, font=("Arial", 9, "bold"), wraplength=400, justify="center")
        self.status_label.pack(pady=5, fill=tk.X)
        
        f_grid = ttk.Frame(g4); f_grid.pack(fill=tk.BOTH, expand=True)
        for c in range(3): f_grid.columnconfigure(c, weight=1)

        self.order_entries = []
        for i in range(15):
            r, c = i // 3, i % 3
            # Frame principale per ogni ordine
            f_main = ttk.Frame(f_grid, padding=2, borderwidth=1, relief="groove")
            f_main.grid(row=r, column=c, sticky="ew", padx=3, pady=3)
            
            # Intestazione Nome ODC (sopra gli input)
            l_name = ttk.Label(f_main, text="", font=("Arial", 8, "bold"), foreground="#005a9e")
            l_name.pack(side=tk.TOP, fill=tk.X)
            
            f_inputs = ttk.Frame(f_main)
            f_inputs.pack(side=tk.TOP, fill=tk.X)
            
            ttk.Label(f_inputs, text=f"{i+1:02}", font=("Arial", 7)).pack(side=tk.LEFT)
            e_n = ttk.Entry(f_inputs, font=("Arial", 10), width=10)
            e_n.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
            ttk.Label(f_inputs, text="P", font=("Arial", 7)).pack(side=tk.LEFT)
            e_p = ttk.Entry(f_inputs, font=("Arial", 10), width=3)
            e_p.pack(side=tk.LEFT, padx=1)
            
            self.order_entries.append((e_n, e_p, l_name))

        self.console_text = tk.Text(self.log_tab, bg="#1e1e1e", fg="#00ff00", font=("Consolas", 10))
        self.console_text.pack(fill=tk.BOTH, expand=True)
        
        bot = ttk.Frame(main_frame, padding=10)
        bot.pack(fill=tk.X, side=tk.BOTTOM)
        self.btn_run = ttk.Button(bot, text="AVVIA SCARICO", command=self.run_script_threaded)
        self.btn_run.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        self.btn_stop = ttk.Button(bot, text="STOP", command=self.stop_process, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        ttk.Button(bot, text="Esci", command=self.root.destroy).pack(side=tk.RIGHT, padx=5)

    def on_account_change(self, event=None):
        acc = self.account_var.get()
        if acc == "TRICHINI":
            self.username_entry.delete(0, tk.END); self.username_entry.insert(0, "9psaraceno")
            self.password_entry.delete(0, tk.END); self.password_entry.insert(0, "Mascara@13")
            self.password_entry.config(show="")
        elif acc == "GIGLIUTO":
            self.username_entry.delete(0, tk.END); self.username_entry.insert(0, "9mgigliuto")
            self.password_entry.delete(0, tk.END); self.password_entry.insert(0, "Catania9+")
            self.password_entry.config(show="")
        else:
            self.password_entry.config(show="*")
        self.save_config()

    def populate_fields(self):
        self.account_var.set(self.config.get("account", "Manuale"))
        self.username_entry.insert(0, self.config.get("username", ""))
        self.password_entry.insert(0, self.config.get("password", ""))
        self.password_entry.config(show="" if self.account_var.get() != "Manuale" else "*")
        self.download_dir_entry.insert(0, self.config.get("download_dir", ""))
        self.move_dir_entry.insert(0, self.config.get("move_dir", ""))
        self.giornaliera_entry.insert(0, self.config.get("giornaliera_path", ""))
        self.login_url_entry.insert(0, self.config.get("login_url", ""))
        self.provider_entry.insert(0, self.config.get("provider", ""))
        self.date_entry.insert(0, self.config.get("date_to_insert", ""))
        
        self.macro_path_entry.insert(0, self.config.get("macro_file_path", ""))
        self.run_macro_var.set(self.config.get("run_macro", False))
        for i, o in enumerate(self.config.get("orders", [])):
            if i < len(self.order_entries):
                self.order_entries[i][0].insert(0, o.get("numero", ""))
                self.order_entries[i][1].insert(0, o.get("posizione", ""))
                self.order_entries[i][2].config(text=o.get("nome", ""))
        
        manuals = self.config.get("manual_consuntivi", {})
        for k, v in manuals.items():
            if k in self.manual_inputs:
                self.manual_inputs[k].delete(0, tk.END)
                self.manual_inputs[k].insert(0, v)

        # Spostato alla fine per evitare che save_config legga campi vuoti durante l'init
        self.calculate_dynamic_paths(auto=True)

    def browse_dir(self, e):
        p = filedialog.askdirectory()
        if p: e.delete(0, tk.END); e.insert(0, p); self.save_config()

    def browse_file(self, e):
        p = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsm *.xlsx")])
        if p: e.delete(0, tk.END); e.insert(0, p); self.save_config()

    def log(self, m): self.log_queue.put(m)

    def update_console(self):
        while not self.log_queue.empty():
            msg = self.log_queue.get()
            self.console_text.insert(tk.END, msg + "\n")
            self.console_text.see(tk.END)
        self.root.after(100, self.update_console)

    def run_script_threaded(self):
        self.save_config(False); self.notebook.select(self.log_tab); self.console_text.delete("1.0", tk.END)
        self.btn_run.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)
        self.log(">>> AVVIO SCARICO...\n")
        threading.Thread(target=self.execute_workflow, daemon=True).start()

    def _search_network_consuntivo(self, year, month, keyword, check_second=False):
        """Logica pura di ricerca su rete."""
        base_path = f"\\\\192.168.11.251\\Database_Tecnico_SMI\\Contabilita' strumentale\\{year}\\CONSUNTIVI\\{year}"
        if not os.path.exists(base_path):
            self.log(f"    ! Percorso non trovato: {base_path}")
            return ""

        search_pattern = os.path.join(base_path, f"*CANONE*{month}*{keyword}*")
        files = glob.glob(search_pattern)
        
        target_file = None
        
        for f in files:
            fname = os.path.basename(f).upper()
            if check_second:
                if " 2" in fname or "_2" in fname or "-2" in fname:
                    target_file = fname
                    break
            else:
                if " 2" not in fname and "_2" not in fname and "-2" not in fname:
                    target_file = fname
                    break
        
        if not target_file and files:
            target_file = os.path.basename(files[0]).upper()

        if target_file:
            match = re.match(r'^(\d+)', target_file)
            if match:
                num = match.group(1)
                self.log(f"    > Trovato per {keyword}: {num} ({target_file})")
                return num
        
        self.log(f"    ! Nessun file trovato per {keyword} (Mese: {month})")
        return ""

    def find_consuntivo_number(self, year, month, keyword, check_second=False):
        """Cerca il numero del consuntivo nel percorso di rete o usa override manuale."""
        manual_key = keyword
        if keyword == "CALDARELLA" and check_second:
            manual_key = "CALDARELLA 2"
            
        if manual_key in self.manual_inputs:
            val = self.manual_inputs[manual_key].get().strip()
            if val:
                self.log(f"    > Uso valore manuale per {manual_key}: {val}")
                return val

        return self._search_network_consuntivo(year, month, keyword, check_second)

    def preview_macro_params(self):
        """Popola le caselle manuali con i valori trovati (se vuote) e mostra i log."""
        self.log("\n>>> SCANSIONE E POPOLAMENTO CAMPI:")
        
        year = self.selected_year.get()
        month = self.selected_month.get()
        
        configs = [
            ("MESSINA", "MESSINA", False),
            ("NASELLI", "NASELLI", False),
            ("CALDARELLA", "CALDARELLA", False),
            ("CALDARELLA 2", "CALDARELLA", True)
        ]

        for key, keyword, is_second in configs:
            if key == "CALDARELLA 2":
                self.log(f"  [MANUAL] {key} -> Da inserire manualmente.")
                continue

            if key in self.manual_inputs:
                current_val = self.manual_inputs[key].get().strip()
                if not current_val:
                    # Se vuoto, cerca e popola
                    found = self._search_network_consuntivo(year, month, keyword, is_second)
                    if found:
                        self.manual_inputs[key].delete(0, tk.END)
                        self.manual_inputs[key].insert(0, found)
                        self.log(f"  [AUTO-FILL] {key} -> {found}")
                    else:
                        self.log(f"  [AUTO-FAIL] {key} -> Nessun file trovato")
                else:
                    self.log(f"  [MANUAL] {key} -> {current_val} (Mantenuto)")
        
        self.save_config()
        self.log(">>> Scansione completata.\n")

    def update_macro_excel(self):
        if not PYWIN32_AVAILABLE:
            self.log(">>> Modulo win32com non disponibile. Impossibile aggiornare Macro.")
            return

        macro_path = self.macro_path_entry.get()
        if not macro_path or not os.path.exists(macro_path):
            self.log(f">>> File Macro non trovato: {macro_path}")
            return

        self.log("\n>>> Aggiornamento File Macro Excel...")
        
        # Analisi abilitazioni dai config orders
        orders = self.config.get("orders", [])
        
        # Normalizzazione nomi per confronto
        enabled_map = {
            "MESSINA": False,
            "NASELLI": False,
            "CALDARELLA": False,
            "CALDARELLA 2": False
        }

        for o in orders:
            nome = o.get("nome", "").upper()
            if "MESSINA" in nome: enabled_map["MESSINA"] = True
            if "NASELLI" in nome: enabled_map["NASELLI"] = True
            # Logica distinzione Caldarella
            if "CALDARELLA" in nome:
                if " 2" in nome or "2" in nome.split(): # Es "CALDARELLA 2"
                    enabled_map["CALDARELLA 2"] = True
                else:
                    enabled_map["CALDARELLA"] = True

        year = self.selected_year.get()
        month = self.selected_month.get()

        try:
            self.log(f">>> Apertura Excel Macro: {macro_path}")
            excel = win32com.client.Dispatch("Excel.Application")
            # Rendiamo Excel visibile per dare feedback visivo all'utente
            excel.Visible = True 
            excel.DisplayAlerts = True
            wb = None
            try:
                clean_path = os.path.normpath(macro_path)
                wb = excel.Workbooks.Open(clean_path, ReadOnly=False)
                if wb is None: raise Exception(f"Impossibile aprire il file (Excel ha restituito None).")
                sheet = wb.Sheets("rif.VBA")
                
                # ... (rest of the logic for configs) ...

                wb.Save()
                self.log(">>> Parametri aggiornati nel file Excel.")
                self.log(">>> AVVIO MACRO 'elaboraTutto'... (Controlla la finestra di Excel per l'avanzamento)")
                
                # Lancio la macro
                excel.Run("elaboraTutto")
                
                # Selezione foglio STAMPA
                try:
                    wb.Sheets("STAMPA").Activate()
                    self.log(">>> Foglio 'STAMPA' selezionato.")
                except:
                    self.log(">>> Avviso: Foglio 'STAMPA' non trovato.")

                self.log(">>> Elaborazione Excel completata. Il file rimarrà aperto per il controllo.")
                
                # Imposto una variabile per evitare la chiusura nel finally
                success_keep_open = True
                
            finally:
                # Chiudiamo solo se non abbiamo avuto successo (success_keep_open non definita o False)
                if 'success_keep_open' not in locals() or not success_keep_open:
                    if wb: wb.Close(False)
                    excel.Quit()
                else:
                    # Se successo, lasciamo Excel visibile e interattivo
                    excel.Visible = True
                    excel.DisplayAlerts = True
        except Exception as e:
            self.log(f">>> Errore durante l'elaborazione Excel: {e}")
        except Exception as e:
            self.log(f">>> Errore aggiornamento Macro: {e}")

    def execute_workflow(self):
        # Inizializza COM per questo thread
        if PYWIN32_AVAILABLE: pythoncom.CoInitialize()
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f: json.dump(self.config, f, indent=4)
            si = subprocess.STARTUPINFO()
            si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            si.wShowWindow = 0
            self.process = subprocess.Popen(["python", "-u", "scaricaTScanoni.py"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1, creationflags=0x08000000, startupinfo=si)
            
            for line in self.process.stdout: 
                self.log(line.strip())
            
            self.process.wait()
            ret_code = self.process.returncode
            self.log(f"\n>>> TERMINATO ({ret_code})\n")
            
            # Esegui Macro Update se richiesto e se lo script è terminato ok (0)
            # Se è stato killato (-15 o 1 su win), non esegue macro
            if ret_code == 0 and self.run_macro_var.get():
                self.update_macro_excel()
                
        except Exception as e: self.log(f"\n>>> ERRORE: {e}\n")
        finally: 
            self.process = None
            if PYWIN32_AVAILABLE: pythoncom.CoUninitialize()
            self.root.after(0, self.reset_buttons)

    def reset_buttons(self):
        self.btn_run.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.DISABLED)

    def stop_process(self):
        if self.process:
            self.log("\n>>> RICHIESTA DI STOP INVIATA...")
            try:
                # Su Windows usiamo taskkill /T per chiudere l'albero dei processi (incluso chromedriver)
                if sys.platform == "win32":
                    subprocess.run(["taskkill", "/F", "/T", "/PID", str(self.process.pid)], 
                                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                else:
                    self.process.terminate()
            except Exception as e:
                self.log(f">>> Errore durante lo stop: {e}")

if __name__ == "__main__":
    root = tk.Tk(); app = SettingsGUI(root); root.mainloop()