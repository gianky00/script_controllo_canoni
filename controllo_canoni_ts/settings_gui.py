import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
import subprocess
import threading
import queue
import sys

# Prova a importare win32com per leggere Excel
try:
    import win32com.client
    PYWIN32_AVAILABLE = True
except ImportError:
    PYWIN32_AVAILABLE = False

CONFIG_FILE = "config_canoni.json"

class SettingsGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Controllo Canoni TS - Smart Config")
        self.root.geometry("800x850")
        
        self.config = self.load_config()
        self.log_queue = queue.Queue()
        
        self.create_widgets()
        self.populate_fields()
        
        self.root.after(100, self.update_console)

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        return {
            "account": "Manuale",
            "login_url": "https://portalefornitori.isab.com/Ui/",
            "username": "",
            "password": "",
            "download_dir": "C:\\Users\\Coemi\\Downloads",
            "move_dir": "",
            "giornaliera_path": "",
            "macro_file_path": "\\\\192.168.11.251\\Database_Tecnico_SMI\\MASTER FOGLI DI CALCOLO\\Comparatore_TS-Giornaliera (canoni).xlsm",
            "run_macro": False,
            "provider": "KK10608 - COEMI S.R.L.",
            "date_to_insert": "01.01.2025",
            "orders": []
        }

    def save_config(self, show_msg=True):
        self.config["account"] = self.account_var.get()
        self.config["login_url"] = self.login_url_entry.get()
        self.config["username"] = self.username_entry.get()
        self.config["password"] = self.password_entry.get()
        self.config["download_dir"] = self.download_dir_entry.get()
        self.config["move_dir"] = self.move_dir_entry.get()
        self.config["giornaliera_path"] = self.giornaliera_entry.get()
        self.config["macro_file_path"] = self.macro_path_entry.get()
        self.config["run_macro"] = self.run_macro_var.get()
        self.config["provider"] = self.provider_entry.get()
        self.config["date_to_insert"] = self.date_entry.get()
        
        self.config["orders"] = []
        for i in range(len(self.order_entries)):
            num = self.order_entries[i][0].get().strip()
            pos = self.order_entries[i][1].get().strip()
            if num:
                self.config["orders"].append({"numero": num, "posizione": pos})

        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self.config, f, indent=4)
        if show_msg:
            messagebox.showinfo("Successo", "Configurazione salvata!")

    def import_from_giornaliera(self):
        if not PYWIN32_AVAILABLE:
            messagebox.showerror("Errore", "Libreria win32com non disponibile.")
            return

        path = self.giornaliera_entry.get()
        if not path or not os.path.exists(path):
            messagebox.showerror("Errore", f"File Giornaliera non trovato:\n{path}")
            return

        try:
            self.log(">>> Apertura file Giornaliera per estrazione OdA...\n")
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            wb = excel.Workbooks.Open(os.path.abspath(path), ReadOnly=True)
            try:
                sheet = wb.Sheets("RIEPILOGO")
                # Celle: S17, U17, V17, W17
                values = [
                    sheet.Range("S17").Value,
                    sheet.Range("U17").Value,
                    sheet.Range("V17").Value,
                    sheet.Range("W17").Value
                ]
                
                # Pulisci tabella GUI
                for ent_num, ent_pos in self.order_entries:
                    ent_num.delete(0, tk.END)
                    ent_pos.delete(0, tk.END)
                
                # Inserisci i nuovi valori
                idx = 0
                for val in values:
                    if val:
                        # Pulisce eventuali .0 dai numeri
                        s_val = str(val).split('.')[0] if isinstance(val, (float, int)) else str(val)
                        if s_val and s_val != "None":
                            self.order_entries[idx][0].insert(0, s_val)
                            self.order_entries[idx][1].insert(0, "10")
                            idx += 1
                
                self.log(f">>> Estratti {idx} ordini correttamente.\n")
                messagebox.showinfo("Successo", f"Importati {idx} ordini con POS 10.")
                
            finally:
                wb.Close(SaveChanges=False)
                excel.Quit()
        except Exception as e:
            messagebox.showerror("Errore Excel", f"Errore durante la lettura:\n{str(e)}")

    def calculate_dynamic_paths(self):
        from datetime import datetime, timedelta
        today = datetime.now()
        prev_month = today.replace(day=1) - timedelta(days=1)
        y, m = prev_month.year, prev_month.month
        months = {1:"GENNAIO", 2:"FEBBRAIO", 3:"MARZO", 4:"APRILE", 5:"MAGGIO", 6:"GIUGNO", 7:"LUGLIO", 8:"AGOSTO", 9:"SETTEMBRE", 10:"OTTOBRE", 11:"NOVEMBRE", 12:"DICEMBRE"}
        
        m_str = f"{m:02d}"
        path_base = f"\\\\192.168.11.251\\Condivisa\\ALLEGRETTI\\{y}\\TS\\CANONI\\{m_str} - {months[m]}"
        file_g = f"\\\\192.168.11.251\\Database_Tecnico_SMI\\Giornaliere\\Giornaliere {y}\\Giornaliera {m_str}-{y}.xlsm"
        
        self.move_dir_entry.delete(0, tk.END); self.move_dir_entry.insert(0, path_base)
        self.giornaliera_entry.delete(0, tk.END); self.giornaliera_entry.insert(0, file_g)
        self.date_entry.delete(0, tk.END); self.date_entry.insert(0, f"01.{m_str}.{y}")

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.settings_tab = ttk.Frame(self.notebook, padding="10")
        self.log_tab = ttk.Frame(self.notebook, padding="10")
        
        self.notebook.add(self.settings_tab, text=" Configurazione ")
        self.notebook.add(self.log_tab, text=" Log Operazioni ")

        # --- SETTINGS TAB ---
        canvas_settings = tk.Canvas(self.settings_tab)
        v_scroll = ttk.Scrollbar(self.settings_tab, orient="vertical", command=canvas_settings.yview)
        self.scroll_frame = ttk.Frame(canvas_settings)

        self.scroll_frame.bind("<Configure>", lambda e: canvas_settings.configure(scrollregion=canvas_settings.bbox("all")))
        canvas_settings.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        canvas_settings.configure(yscrollcommand=v_scroll.set)
        
        canvas_settings.pack(side="left", fill="both", expand=True)
        v_scroll.pack(side="right", fill="y")

        group_base = ttk.LabelFrame(self.scroll_frame, text=" Credenziali e Percorsi ", padding="10")
        group_base.pack(fill=tk.X, pady=5, padx=5)

        ttk.Label(group_base, text="Account:").grid(row=0, column=0, sticky=tk.W)
        self.account_var = tk.StringVar()
        self.account_combo = ttk.Combobox(group_base, textvariable=self.account_var, values=["Manuale", "TRICHINI", "GIGLIUTO"], state="readonly", width=37)
        self.account_combo.grid(row=0, column=1, pady=2)
        self.account_combo.bind("<<ComboboxSelected>>", self.on_account_change)

        ttk.Label(group_base, text="Username:").grid(row=1, column=0, sticky=tk.W)
        self.username_entry = ttk.Entry(group_base, width=40)
        self.username_entry.grid(row=1, column=1, pady=2)

        ttk.Label(group_base, text="Password:").grid(row=2, column=0, sticky=tk.W)
        self.password_entry = ttk.Entry(group_base, show="*", width=40)
        self.password_entry.grid(row=2, column=1, pady=2)

        ttk.Label(group_base, text="Download Dir:").grid(row=3, column=0, sticky=tk.W)
        self.download_dir_entry = ttk.Entry(group_base, width=40)
        self.download_dir_entry.grid(row=3, column=1, pady=2)
        ttk.Button(group_base, text="...", width=3, command=lambda: self.browse_dir(self.download_dir_entry)).grid(row=3, column=2)

        ttk.Label(group_base, text="Spostamento Dir:").grid(row=4, column=0, sticky=tk.W)
        self.move_dir_entry = ttk.Entry(group_base, width=40)
        self.move_dir_entry.grid(row=4, column=1, pady=2)
        ttk.Button(group_base, text="...", width=3, command=lambda: self.browse_dir(self.move_dir_entry)).grid(row=4, column=2)

        ttk.Label(group_base, text="File Giornaliera:").grid(row=5, column=0, sticky=tk.W)
        self.giornaliera_entry = ttk.Entry(group_base, width=40)
        self.giornaliera_entry.grid(row=5, column=1, pady=2)
        ttk.Button(group_base, text="...", width=3, command=lambda: self.browse_file(self.giornaliera_entry)).grid(row=5, column=2)

        btn_calc_f = ttk.Frame(group_base)
        btn_calc_f.grid(row=6, column=1, pady=5, sticky=tk.E)
        ttk.Button(btn_calc_f, text="Ricalcola Percorsi", command=self.calculate_dynamic_paths).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_calc_f, text="Importa OdA da Giornaliera", command=self.import_from_giornaliera).pack(side=tk.LEFT, padx=2)

        group_site = ttk.LabelFrame(self.scroll_frame, text=" Parametri Portale ", padding="10")
        group_site.pack(fill=tk.X, pady=5, padx=5)

        ttk.Label(group_site, text="Login URL:").grid(row=0, column=0, sticky=tk.W)
        self.login_url_entry = ttk.Entry(group_site, width=40)
        self.login_url_entry.grid(row=0, column=1, pady=2)

        ttk.Label(group_site, text="Fornitore:").grid(row=1, column=0, sticky=tk.W)
        self.provider_entry = ttk.Entry(group_site, width=40)
        self.provider_entry.grid(row=1, column=1, pady=2)

        ttk.Label(group_site, text="Data (DD.MM.YYYY):").grid(row=2, column=0, sticky=tk.W)
        self.date_entry = ttk.Entry(group_site, width=40)
        self.date_entry.grid(row=2, column=1, pady=2)

        group_macro = ttk.LabelFrame(self.scroll_frame, text=" Macro Excel ", padding="10")
        group_macro.pack(fill=tk.X, pady=5, padx=5)

        self.run_macro_var = tk.BooleanVar()
        ttk.Checkbutton(group_macro, text="Esegui Macro al termine", variable=self.run_macro_var).grid(row=0, column=0, columnspan=2, sticky=tk.W)

        ttk.Label(group_macro, text="Percorso Macro:").grid(row=1, column=0, sticky=tk.W)
        self.macro_path_entry = ttk.Entry(group_macro, width=40)
        self.macro_path_entry.grid(row=1, column=1, pady=2)
        ttk.Button(group_macro, text="...", width=3, command=lambda: self.browse_file(self.macro_path_entry)).grid(row=1, column=2)

        group_orders = ttk.LabelFrame(self.scroll_frame, text=" Elenco Ordini (OdA) ", padding="10")
        group_orders.pack(fill=tk.X, pady=5, padx=5)

        self.order_entries = []
        for i in range(1, 16):
            f = ttk.Frame(group_orders)
            f.pack(fill=tk.X)
            ttk.Label(f, text=f"OdA {i:02}:").pack(side=tk.LEFT)
            ent_num = ttk.Entry(f, width=15)
            ent_num.pack(side=tk.LEFT, padx=5)
            ttk.Label(f, text="Pos:").pack(side=tk.LEFT)
            ent_pos = ttk.Entry(f, width=8)
            ent_pos.pack(side=tk.LEFT, padx=5)
            self.order_entries.append((ent_num, ent_pos))

        self.console_text = tk.Text(self.log_tab, wrap=tk.WORD, bg="black", fg="lightgreen", font=("Consolas", 10))
        self.console_text.pack(fill=tk.BOTH, expand=True)
        
        bottom_frame = ttk.Frame(main_frame, padding="5")
        bottom_frame.pack(fill=tk.X, side=tk.BOTTOM)

        ttk.Button(bottom_frame, text="Salva Config", command=self.save_config).pack(side=tk.LEFT, padx=5)
        self.btn_run = ttk.Button(bottom_frame, text="AVVIA SCARICO", command=self.run_script_threaded)
        self.btn_run.pack(side=tk.LEFT, padx=5)
        ttk.Button(bottom_frame, text="Chiudi", command=self.root.destroy).pack(side=tk.RIGHT, padx=5)

    def on_account_change(self, event=None):
        s = self.account_var.get()
        if s == "TRICHINI":
            self.username_entry.delete(0, tk.END); self.username_entry.insert(0, "9psaraceno")
            self.password_entry.delete(0, tk.END); self.password_entry.insert(0, "Mascara@13")
        elif s == "GIGLIUTO":
            self.username_entry.delete(0, tk.END); self.username_entry.insert(0, "9mgigliuto")
            self.password_entry.delete(0, tk.END); self.password_entry.insert(0, "Catania9+")

    def populate_fields(self):
        self.account_var.set(self.config.get("account", "Manuale"))
        self.username_entry.insert(0, self.config.get("username", ""))
        self.password_entry.insert(0, self.config.get("password", ""))
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

    def browse_dir(self, e):
        p = filedialog.askdirectory()
        if p: e.delete(0, tk.END); e.insert(0, p)

    def browse_file(self, e):
        p = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsm *.xlsx")])
        if p: e.delete(0, tk.END); e.insert(0, p)

    def log(self, m): self.log_queue.put(m)

    def update_console(self):
        while not self.log_queue.empty():
            msg = self.log_queue.get()
            self.console_text.insert(tk.END, msg)
            self.console_text.see(tk.END)
        self.root.after(100, self.update_console)

    def run_script_threaded(self):
        self.save_config(False)
        self.notebook.select(self.log_tab)
        self.console_text.delete("1.0", tk.END)
        self.btn_run.config(state=tk.DISABLED)
        self.log(">>> AVVIO PROCESSO DI SCARICO...\n")
        threading.Thread(target=self.execute_workflow, daemon=True).start()

    def execute_workflow(self):
        try:
            p = subprocess.Popen(["python", "-u", "scaricaTScanoni.py"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1, creationflags=0x08000000)
            for line in p.stdout: self.log(line)
            p.wait()
            self.log(f"\n>>> TERMINATO CON CODICE {p.returncode}\n")
        except Exception as e: self.log(f"\n>>> ERRORE: {str(e)}\n")
        finally: self.root.after(0, lambda: self.btn_run.config(state=tk.NORMAL))

if __name__ == "__main__":
    root = tk.Tk()
    app = SettingsGUI(root)
    root.mainloop()
