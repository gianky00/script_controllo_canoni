import sys
import os
import pandas as pd
import numpy as np # Importato per le operazioni vettorizzate
from datetime import datetime, time, date, timedelta
import calendar
import json

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTableView, QLineEdit, QComboBox, QDateEdit, QPushButton,
    QLabel, QFrame, QStatusBar, QMessageBox, QFileDialog, QHeaderView,
    QStyle, QMenuBar, QCheckBox, QDialog, QListWidget, QListWidgetItem,
    QDialogButtonBox, QSpinBox, QGridLayout, QTextBrowser, QTimeEdit,
    QGroupBox
)
from PyQt6.QtCore import (
    QAbstractTableModel, Qt, QDate, QTimer, QSettings, QSortFilterProxyModel, QTime
)
from PyQt6.QtGui import QIcon, QColor, QAction

from reportlab.lib.pagesizes import letter, landscape, portrait
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# --- Stile (invariato) ---
LIGHT_STYLE = """
    QWidget {
        background-color: #F0F2F5; color: #333333; font-family: 'Segoe UI', Arial, sans-serif;
    }
    QMainWindow, QMenuBar, QDialog { background-color: #FFFFFF; }
    QTextBrowser { background-color: #FFFFFF; border: 1px solid #DCDCDC; }
    QTableView {
        background-color: #FFFFFF; gridline-color: #DCDCDC; border: 1px solid #DCDCDC;
        selection-background-color: #0078D7; selection-color: #FFFFFF; font-size: 13px;
    }
    QHeaderView::section {
        background-color: #F0F2F5; padding: 6px; border: none;
        border-bottom: 2px solid #0078D7; font-weight: 600; font-size: 13px;
    }
    QLineEdit, QComboBox, QDateEdit, QSpinBox, QListWidget, QTimeEdit {
        background-color: #FFFFFF; border: 1px solid #C8C8C8; padding: 6px;
        border-radius: 4px; font-size: 13px;
    }
    QPushButton#date_button { font-size: 14px; font-weight: bold; max-width: 25px; }
    QPushButton {
        background-color: #0078D7; color: #FFFFFF; border: none; padding: 8px 16px;
        border-radius: 4px; font-weight: 600; font-size: 13px;
    }
    QPushButton#quick_filter {
        background-color: #E1E1E1; color: #333333; font-weight: normal;
        padding: 6px 10px;
    }
    QPushButton:hover { background-color: #1085E5; }
    QLabel#dashboard_value { font-size: 20px; font-weight: 700; color: #0078D7; }
    QGroupBox { font-weight: bold; margin-top: 10px; }
    QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; padding: 0 3px; }
"""

DEFAULT_CONFIG = {
    "minuti_ravvicinata": 60, "alert_mancanze": True, "alert_invertiti": True,
    "alert_fuori_orario": False, "orario_inizio_std": "07:00", "orario_fine_std": "20:00",
    "alert_turno_breve": True, "min_ore_valide": 1,
    "alert_turno_esteso": False, "max_ore_normali": 10
}
USER_NOTES_FILE = "user_notes.json"

class PandasModel(QAbstractTableModel):
    def __init__(self, data, checked_set, user_notes_dict_ref, app_ref):
        super().__init__()
        self._data = data
        self.checked_set = checked_set
        self.user_notes_dict = user_notes_dict_ref
        self.app = app_ref
        self.highlight_colors = {
            "AVVISO": QColor("#FFF3CD"), "ATTENZIONE": QColor("#FFE5CC"), "ERRORE": QColor("#F8D7DA"),
        }
        self.column_tooltips = {
            "Seleziona": "Spunta per selezionare la riga per l'esportazione.",
            "Sito": "Sito di timbratura.", "Reparto": "Reparto di appartenenza.",
            "Data": "Data timbratura.", "Nome": "Nome.", "Cognome": "Cognome.",
            "Ingresso": "Orario ingresso effettivo.", "Uscita": "Orario uscita effettivo.",
            "Ingresso Contabile": "Ingresso arrotondato.", "Uscita Contabile": "Uscita arrotondata.",
            "Ore Contabili": "Ore calcolate su orari contabili.",
            "Avvisi Sistema": "Avvisi automatici basati sulle regole.",
            "Note Utente": "Doppio click per aggiungere/modificare una nota."
        }

    def rowCount(self, parent=None): return self._data.shape[0]
    def columnCount(self, parent=None): return self._data.shape[1]

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or self._data.empty: return None
        original_df_idx = self._data.iloc[index.row()].get('original_df_index', -1)
        col_name = self._data.columns[index.column()]
        if role == Qt.ItemDataRole.CheckStateRole and col_name == 'Seleziona':
            return Qt.CheckState.Checked if original_df_idx in self.checked_set else Qt.CheckState.Unchecked
        if role == Qt.ItemDataRole.BackgroundRole and 'Highlight' in self._data.columns:
            return self.highlight_colors.get(self._data.iloc[index.row()]['Highlight'])
        if role == Qt.ItemDataRole.DisplayRole or role == Qt.ItemDataRole.EditRole:
            if col_name == 'Seleziona': return ""
            if col_name == 'Note Utente': return self.user_notes_dict.get(original_df_idx, "")
            value = self._data.iloc[index.row(), index.column()]
            if isinstance(value, float): return f"{value:.2f}".replace('.', ',') if role == Qt.ItemDataRole.DisplayRole else value
            return str(value) if pd.notna(value) else ""
        return None

    def setData(self, index, value, role):
        if not index.isValid() or self._data.empty: return False
        original_df_idx = self._data.iloc[index.row()].get('original_df_index', -1)
        if original_df_idx == -1: return False
        col_name = self._data.columns[index.column()]
        if col_name == 'Seleziona' and role == Qt.ItemDataRole.CheckStateRole:
            if value == Qt.CheckState.Checked.value: self.checked_set.add(original_df_idx)
            else: self.checked_set.discard(original_df_idx)
            self.dataChanged.emit(index, index, [role]); return True
        if col_name == 'Note Utente' and role == Qt.ItemDataRole.EditRole:
            clean_value = str(value).strip()
            if clean_value: self.user_notes_dict[original_df_idx] = clean_value
            else: self.user_notes_dict.pop(original_df_idx, None)
            self.app.save_user_notes()
            self.dataChanged.emit(index, index, [role]); return True
        return False

    def flags(self, index):
        base_flags = super().flags(index)
        if not self._data.empty:
            col_name = self._data.columns[index.column()]
            if col_name == 'Seleziona': return base_flags | Qt.ItemFlag.ItemIsUserCheckable
            if col_name == 'Note Utente': return base_flags | Qt.ItemFlag.ItemIsEditable
        return base_flags

    def headerData(self, section, orientation, role):
        if orientation == Qt.Orientation.Horizontal:
            if section < len(self._data.columns):
                col_name = self._data.columns[section]
                if role == Qt.ItemDataRole.DisplayRole: return str(col_name)
                if role == Qt.ItemDataRole.ToolTipRole: return self.column_tooltips.get(col_name, col_name)
        return None

# --- Finestra di Dialogo Impostazioni Avvisi (invariata) ---
class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Impostazioni Regole Avvisi")
        self.settings = QSettings("MyCompany", "TimbratureApp_v9")
        self.setMinimumWidth(500)
        layout = QVBoxLayout(self)

        # Timbratura Ravvicinata
        group_ravvicinata = QGroupBox("Timbratura Ravvicinata")
        rav_layout = QHBoxLayout()
        rav_layout.addWidget(QLabel("Segnala se < di (minuti):"))
        self.spin_min_ravvicinata = QSpinBox(); self.spin_min_ravvicinata.setRange(1, 240); self.spin_min_ravvicinata.setToolTip("Durata in minuti sotto la quale la timbratura è 'ravvicinata'.")
        rav_layout.addWidget(self.spin_min_ravvicinata)
        group_ravvicinata.setLayout(rav_layout)
        layout.addWidget(group_ravvicinata)

        group_enable_alerts = QGroupBox("Abilita Avvisi Specifici")
        enable_layout = QGridLayout()
        self.cb_alert_mancanze = QCheckBox("Mancanze Ingresso/Uscita"); self.cb_alert_mancanze.setToolTip("Segnala se manca l'ingresso o l'uscita.")
        self.cb_alert_invertiti = QCheckBox("Orari Invertiti"); self.cb_alert_invertiti.setToolTip("Segnala se l'uscita precede l'ingresso (errore logico).")
        self.cb_alert_fuori_orario = QCheckBox("Timbrature Fuori Orario Standard"); self.cb_alert_fuori_orario.setToolTip("Usa gli orari standard sotto per segnalare timbrature anomale.")
        self.cb_alert_turno_breve = QCheckBox("Turno Troppo Breve"); self.cb_alert_turno_breve.setToolTip("Segnala turni più corti del minimo definito (esclusi 'ravvicinati').")
        self.cb_alert_turno_esteso = QCheckBox("Turno Esteso (per Revisione)"); self.cb_alert_turno_esteso.setToolTip("Segnala turni più lunghi del massimo definito.")

        enable_layout.addWidget(self.cb_alert_mancanze, 0, 0); enable_layout.addWidget(self.cb_alert_invertiti, 0, 1)
        enable_layout.addWidget(self.cb_alert_fuori_orario, 1, 0); enable_layout.addWidget(self.cb_alert_turno_breve, 1, 1)
        enable_layout.addWidget(self.cb_alert_turno_esteso, 2,0)
        group_enable_alerts.setLayout(enable_layout)
        layout.addWidget(group_enable_alerts)

        self.group_orari_std = QGroupBox("Orario Standard (per avviso 'Fuori Orario')")
        orari_std_layout = QGridLayout()
        orari_std_layout.addWidget(QLabel("Inizio Standard:"),0,0); self.time_inizio_std = QTimeEdit(); self.time_inizio_std.setDisplayFormat("HH:mm"); orari_std_layout.addWidget(self.time_inizio_std,0,1)
        orari_std_layout.addWidget(QLabel("Fine Standard:"),1,0); self.time_fine_std = QTimeEdit(); self.time_fine_std.setDisplayFormat("HH:mm"); orari_std_layout.addWidget(self.time_fine_std,1,1)
        self.group_orari_std.setLayout(orari_std_layout); layout.addWidget(self.group_orari_std)
        self.cb_alert_fuori_orario.toggled.connect(self.group_orari_std.setEnabled)

        self.group_soglie_durata = QGroupBox("Durata Turno (per avvisi 'Breve'/'Esteso')")
        soglie_layout = QGridLayout()
        soglie_layout.addWidget(QLabel("Min Ore Turno 'Breve':"),0,0); self.spin_min_ore_valide = QSpinBox(); self.spin_min_ore_valide.setRange(1,12); self.spin_min_ore_valide.setToolTip("Sotto questa soglia (e sopra 'ravvicinata'), il turno è 'troppo breve'."); soglie_layout.addWidget(self.spin_min_ore_valide,0,1)
        soglie_layout.addWidget(QLabel("Max Ore Turno 'Esteso':"),1,0); self.spin_max_ore_normali = QSpinBox(); self.spin_max_ore_normali.setRange(1,24); self.spin_max_ore_normali.setToolTip("Sopra questa soglia, il turno è 'esteso'."); soglie_layout.addWidget(self.spin_max_ore_normali,1,1)
        self.group_soglie_durata.setLayout(soglie_layout); layout.addWidget(self.group_soglie_durata)
        self.cb_alert_turno_breve.toggled.connect(self.spin_min_ore_valide.setEnabled)
        self.cb_alert_turno_esteso.toggled.connect(self.spin_max_ore_normali.setEnabled)

        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.RestoreDefaults | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.button(QDialogButtonBox.StandardButton.Save).setText("Salva"); self.button_box.button(QDialogButtonBox.StandardButton.RestoreDefaults).setText("Predefiniti")
        self.button_box.accepted.connect(self.save_rules_settings); self.button_box.rejected.connect(self.reject)
        self.button_box.button(QDialogButtonBox.StandardButton.RestoreDefaults).clicked.connect(self.restore_defaults)
        layout.addWidget(self.button_box)
        self.load_rules_settings()
        self.group_orari_std.setEnabled(self.cb_alert_fuori_orario.isChecked())
        self.spin_min_ore_valide.setEnabled(self.cb_alert_turno_breve.isChecked())
        self.spin_max_ore_normali.setEnabled(self.cb_alert_turno_esteso.isChecked())

    def load_rules_settings(self):
        for key, default_value in DEFAULT_CONFIG.items():
            if isinstance(default_value, bool):
                checkbox_name = f"cb_alert_{key.split('_', 1)[1]}" if key.startswith("alert_") else None
                if hasattr(self, checkbox_name): getattr(self, checkbox_name).setChecked(self.settings.value(f"rules/{key}", default_value, type=bool))
            elif isinstance(default_value, int):
                spinbox_name = f"spin_{key}"
                if hasattr(self, spinbox_name): getattr(self, spinbox_name).setValue(int(self.settings.value(f"rules/{key}", default_value)))
            elif ":" in str(default_value):
                timeedit_name = f"time_{key.replace('orario_', '').replace('_std','')}_std"
                if hasattr(self, timeedit_name): getattr(self, timeedit_name).setTime(QTime.fromString(self.settings.value(f"rules/{key}", default_value), "HH:mm"))

    def save_rules_settings(self):
        for key, default_value in DEFAULT_CONFIG.items():
            if isinstance(default_value, bool):
                checkbox_name = f"cb_alert_{key.split('_', 1)[1]}" if key.startswith("alert_") else None
                if hasattr(self, checkbox_name): self.settings.setValue(f"rules/{key}", getattr(self, checkbox_name).isChecked())
            elif isinstance(default_value, int):
                spinbox_name = f"spin_{key}"
                if hasattr(self, spinbox_name): self.settings.setValue(f"rules/{key}", getattr(self, spinbox_name).value())
            elif ":" in str(default_value):
                timeedit_name = f"time_{key.replace('orario_', '').replace('_std','')}_std"
                if hasattr(self, timeedit_name): self.settings.setValue(f"rules/{key}", getattr(self, timeedit_name).time().toString("HH:mm"))
        self.accept()

    def restore_defaults(self):
        for key, default_value in DEFAULT_CONFIG.items():
            if isinstance(default_value, bool):
                checkbox_name = f"cb_alert_{key.split('_', 1)[1]}" if key.startswith("alert_") else None
                if hasattr(self, checkbox_name): getattr(self, checkbox_name).setChecked(default_value)
            elif isinstance(default_value, int):
                spinbox_name = f"spin_{key}"
                if hasattr(self, spinbox_name): getattr(self, spinbox_name).setValue(default_value)
            elif ":" in str(default_value):
                timeedit_name = f"time_{key.replace('orario_', '').replace('_std','')}_std"
                if hasattr(self, timeedit_name): getattr(self, timeedit_name).setTime(QTime.fromString(default_value, "HH:mm"))
        QMessageBox.information(self,"Predefiniti","Impostazioni ripristinate. Clicca 'Salva' per applicare.")


# --- Finestra di Dialogo Guida Utente (invariata) ---
class HelpGuideDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Guida Utente - ISAB Sud Timbrature")
        self.setMinimumSize(650, 500)
        layout = QVBoxLayout(self)
        text_browser = QTextBrowser(self); text_browser.setOpenExternalLinks(True)
        guide_html = """
            <html><head><style>
                body { font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; }
                h2 { color: #0078D7; border-bottom: 1px solid #0078D7; padding-bottom: 5px;}
                h3 { color: #333; } p { margin-bottom: 10px; } ul { margin-left: 20px; }
                code { background-color: #f0f0f0; padding: 2px 4px; border-radius: 3px; }
            </style></head><body>
            <h2>Guida Rapida all'Applicazione Timbrature v9.1 (Ottimizzata)</h2>
            <h3>1. Caricamento Dati e Cache</h3>
            <p>All'avvio, carica <code>database_timbrature_isab.xlsm</code>. La prima volta processa l'intero file. Poi usa una <b>cache</b> (<code>data_cache.pkl</code>) per avvii veloci, a meno che l'Excel non sia modificato. Il file Excel deve contenere un foglio "<b>Reparto</b>" (colonne: <code>Nome, Cognome, Reparto</code>).</p>
            <h3>2. Filtri</h3>
            <ul>
                <li><b>Ricerca Testuale:</b> Su Nome, Cognome, Sito.</li>
                <li><b>Sito/Reparto:</b> Filtri a tendina.</li>
                <li><b>Periodo (Da/A):</b> Con pulsanti <code>+/-</code> e filtri rapidi ("Ieri", "Sett. corr.", "Mese corr.").</li>
                <li><b>Filtro Avvisi:</b> Checkbox per mostrare solo righe con avvisi (es. "Timbr. Ravvicinata" se attiva nelle impostazioni).</li>
            </ul>
            <h3>3. Impostazioni Avvisi (Menu File)</h3>
            <p>Permette di personalizzare le regole per generare gli "Avvisi Sistema": definisci soglie per timbrature ravvicinate, abilita/disabilita avvisi per mancanze, orari invertiti, fuori orario (con orari standard configurabili), turni troppo brevi o estesi.</p>
            <h3>4. Tabella Dati</h3>
            <ul>
                <li><b>Ordinamento:</b> Clicca sull'intestazione di colonna.</li>
                <li><b>"Seleziona":</b> Checkbox per esportare righe specifiche.</li>
                <li><b>"Avvisi Sistema":</b> Segnalazioni automatiche basate sulle regole impostate. Cella vuota se OK.</li>
                <li><b>"Note Utente":</b> Colonna per tue annotazioni. Doppio click su una cella per aggiungere/modificare una nota (salvata in <code>user_notes.json</code>).</li>
                <li><b>Formato Ore:</b> Decimali con virgola.</li>
            </ul>
            <h3>5. Esportazione e Report</h3>
            <p>Esporta <em>solo</em> le righe selezionate in CSV/PDF. "Genera Report Mensile" crea PDF dettagliati per dipendente.</p>
            <h3>6. Impostazioni App</h3>
            <p>Dimensione/posizione finestra salvate. Le note e le regole avvisi sono salvate automaticamente.</p>
            <hr><p><i>Versione Applicazione: 9.1</i></p>
            </body></html>
        """
        text_browser.setHtml(guide_html); layout.addWidget(text_browser)
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        button_box.accepted.connect(self.accept); layout.addWidget(button_box)


# --- Classe Principale dell'Applicazione ---
class TimbratureApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ISAB Sud - Control & Report v9.1 (Ottimizzata)") # VERSIONE AGGIORNATA
        self.setWindowIcon(QIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon)))

        self.df_raw_data = None
        self.df_original = None
        self.checked_indices = set()
        self.user_notes = {}
        self.config_rules = {}

        self.settings = QSettings("MyCompany", "TimbratureApp_v9")
        self.load_app_config()
        self.load_user_notes()

        self.search_timer = QTimer(self)
        self.search_timer.setSingleShot(True); self.search_timer.setInterval(300)
        self.search_timer.timeout.connect(self.apply_filters)

        self.init_ui()
        self.load_window_settings()
        QTimer.singleShot(100, self.load_data_and_process)


    def create_menu_bar(self):
        menu_bar = QMenuBar(self)
        file_menu = menu_bar.addMenu("&File")

        settings_action = QAction("Impostazioni Avvisi...", self)
        settings_action.triggered.connect(self.show_settings_dialog)
        file_menu.addAction(settings_action)
        file_menu.addSeparator()
        exit_action = QAction("Esci", self); exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        help_menu = menu_bar.addMenu("&Aiuto")
        guide_action = QAction("Guida Utente...", self); guide_action.triggered.connect(self.show_help_guide_dialog)
        help_menu.addAction(guide_action)
        about_action = QAction("Informazioni", self); about_action.triggered.connect(self.show_about_dialog)
        help_menu.addAction(about_action)
        return menu_bar

    def show_settings_dialog(self):
        dialog = SettingsDialog(self)
        if dialog.exec():
            self.load_app_config()
            if self.df_raw_data is not None:
                # Usa una copia per la ri-elaborazione
                self.df_original = self._process_loaded_data(self.df_raw_data.copy())
                self.apply_filters()
            QMessageBox.information(self, "Impostazioni", "Impostazioni salvate. La vista dati è stata aggiornata.")

    def load_app_config(self):
        self.config_rules = {}
        for key, default_value in DEFAULT_CONFIG.items():
            if isinstance(default_value, bool):
                self.config_rules[key] = self.settings.value(f"rules/{key}", default_value, type=bool)
            elif isinstance(default_value, int):
                self.config_rules[key] = int(self.settings.value(f"rules/{key}", default_value))
            else:
                self.config_rules[key] = self.settings.value(f"rules/{key}", default_value)


    def load_user_notes(self):
        if os.path.exists(USER_NOTES_FILE):
            try:
                with open(USER_NOTES_FILE, 'r', encoding='utf-8') as f:
                    loaded_notes = json.load(f)
                    self.user_notes = {int(k): v for k, v in loaded_notes.items()}
            except Exception as e:
                print(f"Errore caricamento note: {e}")
                self.user_notes = {}
        else:
            self.user_notes = {}

    def save_user_notes(self):
        try:
            with open(USER_NOTES_FILE, 'w', encoding='utf-8') as f:
                notes_to_save = {str(k): v for k, v in self.user_notes.items()}
                json.dump(notes_to_save, f, indent=4)
        except Exception as e:
            print(f"Errore salvataggio note: {e}")

    def init_ui(self):
        self.setMenuBar(self.create_menu_bar())
        main_widget = QWidget(); self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget); main_layout.setSpacing(10); main_layout.setContentsMargins(15, 15, 15, 15)
        self.setup_controls_and_dashboard(main_layout)

        anomaly_filter_group = QFrame(); anomaly_filter_group.setFrameShape(QFrame.Shape.StyledPanel)
        anomaly_layout = QHBoxLayout(anomaly_filter_group)
        anomaly_layout.addWidget(QLabel("<b>Filtra Avvisi di Sistema:</b>"))
        self.cb_filter_anomalies = QCheckBox("Mostra solo righe con Avvisi")
        self.cb_filter_anomalies.setToolTip("Mostra solo le timbrature che hanno generato un avviso automatico secondo le regole correnti.")
        self.cb_filter_anomalies.stateChanged.connect(self.apply_filters)
        anomaly_layout.addWidget(self.cb_filter_anomalies)
        anomaly_layout.addStretch()
        main_layout.addWidget(anomaly_filter_group)

        self.table_view = QTableView(); self.table_view.setSortingEnabled(True)
        self.table_view.doubleClicked.connect(self.handle_double_click) # Gestione doppio click
        main_layout.addWidget(self.table_view)

        export_layout = QHBoxLayout(); export_layout.addStretch()
        self.export_csv_button = QPushButton(" Esporta Selezionati CSV"); self.export_pdf_button = QPushButton(" Esporta Selezionati PDF")
        self.export_csv_button.setIcon(QIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogSaveButton)))
        self.export_pdf_button.setIcon(QIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon)))
        self.export_csv_button.setToolTip("Esporta le righe selezionate con checkbox in formato CSV.")
        self.export_pdf_button.setToolTip("Esporta le righe selezionate con checkbox in formato PDF.")
        self.export_csv_button.clicked.connect(self.export_to_csv); self.export_pdf_button.clicked.connect(self.export_to_pdf)
        export_layout.addWidget(self.export_csv_button); export_layout.addWidget(self.export_pdf_button)
        main_layout.addLayout(export_layout)
        self.status_bar = QStatusBar(); self.setStatusBar(self.status_bar); self.status_bar.showMessage("Pronto.")

    def setup_controls_and_dashboard(self, parent_layout):
        top_frame = QWidget(); top_layout = QVBoxLayout(top_frame)
        top_layout.setContentsMargins(0,0,0,0); top_layout.setSpacing(10)
        actions_layout = QHBoxLayout()
        self.report_button = QPushButton(" Genera Report Mensile"); self.report_button.setIcon(QIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon)))
        self.report_button.setToolTip("Apre la finestra per generare un report PDF mensile per i dipendenti selezionati.")
        self.report_button.clicked.connect(self.open_report_dialog)
        self.reset_button = QPushButton(" Reset Filtri"); self.reset_button.setIcon(QIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogCancelButton)))
        self.reset_button.setToolTip("Resetta tutti i filtri ai valori predefiniti e ricarica la tabella completa.")
        self.reset_button.clicked.connect(self.reset_all_filters)
        actions_layout.addWidget(self.report_button); actions_layout.addStretch(); actions_layout.addWidget(self.reset_button)
        top_layout.addLayout(actions_layout)
        filter_box = QFrame(); filter_box.setFrameShape(QFrame.Shape.StyledPanel); filter_layout = QVBoxLayout(filter_box)
        row1_layout = QHBoxLayout()
        self.search_bar = QLineEdit(); self.search_bar.setPlaceholderText("Cerca per nome, cognome, sito...");
        self.search_bar.setToolTip("Ricerca testuale istantanea (dopo breve pausa) su Nome, Cognome e Sito.")
        self.search_bar.textChanged.connect(self.on_search_text_changed)
        row1_layout.addWidget(QLabel("Ricerca:")); row1_layout.addWidget(self.search_bar, 3)
        self.sito_combo = QComboBox(); self.sito_combo.setToolTip("Filtra per Sito di timbratura."); self.sito_combo.currentIndexChanged.connect(self.apply_filters)
        row1_layout.addSpacing(20); row1_layout.addWidget(QLabel("Sito:")); row1_layout.addWidget(self.sito_combo, 1)
        self.reparto_combo = QComboBox(); self.reparto_combo.setToolTip("Filtra per Reparto (dati dal foglio 'Reparto' del file Excel)."); self.reparto_combo.currentIndexChanged.connect(self.apply_filters)
        row1_layout.addSpacing(20); row1_layout.addWidget(QLabel("Reparto:")); row1_layout.addWidget(self.reparto_combo, 1)
        filter_layout.addLayout(row1_layout)
        date_filter_layout = QHBoxLayout()
        self.date_from = QDateEdit(calendarPopup=True); self.date_from.setToolTip("Data di inizio del periodo da analizzare."); self.date_from.dateChanged.connect(self.apply_filters)
        btn_from_minus = QPushButton("-"); btn_from_plus = QPushButton("+"); btn_from_minus.setObjectName("date_button"); btn_from_plus.setObjectName("date_button")
        btn_from_minus.setToolTip("Diminuisci la data di inizio di un giorno."); btn_from_plus.setToolTip("Aumenta la data di inizio di un giorno.")
        btn_from_minus.clicked.connect(lambda: self.date_from.setDate(self.date_from.date().addDays(-1))); btn_from_plus.clicked.connect(lambda: self.date_from.setDate(self.date_from.date().addDays(1)))
        date_filter_layout.addWidget(QLabel("Periodo da:")); date_filter_layout.addWidget(btn_from_minus); date_filter_layout.addWidget(self.date_from); date_filter_layout.addWidget(btn_from_plus); date_filter_layout.addSpacing(10)
        self.date_to = QDateEdit(calendarPopup=True); self.date_to.setToolTip("Data di fine del periodo da analizzare."); self.date_to.dateChanged.connect(self.apply_filters)
        btn_to_minus = QPushButton("-"); btn_to_plus = QPushButton("+"); btn_to_minus.setObjectName("date_button"); btn_to_plus.setObjectName("date_button")
        btn_to_minus.setToolTip("Diminuisci la data di fine di un giorno."); btn_to_plus.setToolTip("Aumenta la data di fine di un giorno.")
        btn_to_minus.clicked.connect(lambda: self.date_to.setDate(self.date_to.date().addDays(-1))); btn_to_plus.clicked.connect(lambda: self.date_to.setDate(self.date_to.date().addDays(1)))
        date_filter_layout.addWidget(QLabel("A:")); date_filter_layout.addWidget(btn_to_minus); date_filter_layout.addWidget(self.date_to); date_filter_layout.addWidget(btn_to_plus)
        date_filter_layout.addSpacing(20)
        quick_filters_map = {"Ieri": (self.filter_yesterday, "Imposta il periodo a ieri."), "Sett. corr.": (self.filter_this_week, "Imposta il periodo alla settimana corrente (Lun-Dom)."), "Mese corr.": (self.filter_this_month, "Imposta il periodo al mese corrente.")}
        for text, (func, tooltip_text) in quick_filters_map.items():
            btn = QPushButton(text); btn.setObjectName("quick_filter"); btn.setToolTip(tooltip_text); btn.clicked.connect(func); date_filter_layout.addWidget(btn)
        date_filter_layout.addStretch(); filter_layout.addLayout(date_filter_layout); top_layout.addWidget(filter_box)
        parent_layout.addWidget(top_frame)

    def handle_double_click(self, index):
        """Apre l'editor sulla colonna 'Note Utente' con doppio click."""
        proxy_model = self.table_view.model()
        source_model = proxy_model.sourceModel()
        col_name = source_model._data.columns[index.column()]
        if col_name == 'Note Utente':
            self.table_view.edit(index)

    def on_search_text_changed(self): self.search_timer.start()

    def load_data_and_process(self):
        excel_file = "database_timbrature_isab.xlsm"; cache_file = "data_cache.pkl"
        if not os.path.exists(excel_file): QMessageBox.critical(self, "Errore", f"File timbrature non trovato: {excel_file}"); return

        use_cache = False
        if os.path.exists(cache_file):
            excel_mod_time = os.path.getmtime(excel_file); cache_mod_time = os.path.getmtime(cache_file)
            if cache_mod_time > excel_mod_time:
                try:
                    self.status_bar.showMessage("Verifica cache...");
                    # La cache ora contiene sempre il df processato, quindi non serve controllare le colonne
                    self.df_original = pd.read_pickle(cache_file)
                    use_cache = True
                    self.status_bar.showMessage("Caricamento dati dalla cache (veloce)...")
                    # Ricostruisci una versione approssimativa di df_raw_data se necessario
                    cols_to_drop = ['Avvisi Sistema', 'Highlight', 'Ingresso Contabile_t', 'Uscita Contabile_t', 'Ore Contabili', 'Reparto']
                    self.df_raw_data = self.df_original.drop(columns=cols_to_drop, errors='ignore')
                except Exception as e:
                    self.status_bar.showMessage(f"Errore cache: {e}. Ricarico da Excel...")

        try:
            if not use_cache:
                self.status_bar.showMessage("Caricamento file Excel (può richiedere tempo)...")
                QApplication.processEvents() # Forza aggiornamento UI
                df_raw = pd.read_excel(excel_file, engine='openpyxl', usecols='B,C,D,H,I,P', sheet_name=0)
                df_raw.columns = ['Data', 'Ingresso', 'Uscita', 'Nome', 'Cognome', 'Sito']
                df_raw.dropna(how='all', inplace=True); df_raw.dropna(subset=['Nome', 'Cognome', 'Data'], inplace=True)
                for col in ['Nome', 'Cognome', 'Sito']: df_raw[col] = df_raw[col].astype(str).str.strip()
                df_raw['Nome'] = df_raw['Nome'].str.title(); df_raw['Cognome'] = df_raw['Cognome'].str.title()
                df_raw['Sito'].replace('', "Non Specificato", inplace=True)

                # Conversione date/ore con gestione errori
                df_raw['Data_dt'] = pd.to_datetime(df_raw['Data'], errors='coerce')
                df_raw['Ingresso_t_raw'] = pd.to_datetime(df_raw['Ingresso'], format='%H:%M', errors='coerce').dt.time
                df_raw['Uscita_t_raw'] = pd.to_datetime(df_raw['Uscita'], format='%H:%M', errors='coerce').dt.time
                df_raw.dropna(subset=['Data_dt'], inplace=True) # Rimuove righe con date invalide

                self.df_raw_data = df_raw.copy()
                self.df_original = self._process_loaded_data(df_raw.copy()) # Usa una copia
                self.df_original.to_pickle(cache_file)

            self.status_bar.showMessage(f"Caricate {len(self.df_original)} timbrature.", 5000)
            self.setup_filters(); self.apply_filters()
        except Exception as e:
            QMessageBox.critical(self, "Errore Lettura Dati", f"Impossibile leggere il file.\nErrore: {e}\n\nAssicurarsi che il file non sia corrotto e che le colonne siano corrette.")

    def _process_loaded_data(self, df_to_process):
        self.status_bar.showMessage("Processamento dati (reparti e avvisi)...")
        QApplication.processEvents()
        excel_file = "database_timbrature_isab.xlsm"
        try:
            df_reparti = pd.read_excel(excel_file, sheet_name="Reparto", usecols="A,B,C", engine='openpyxl')
            df_reparti.columns = ['Nome', 'Cognome', 'Reparto']
            for col in ['Nome', 'Cognome', 'Reparto']: df_reparti[col] = df_reparti[col].astype(str).str.strip().str.title()
            df_reparti.dropna(subset=['Nome', 'Cognome'], inplace=True)
            df_to_process = pd.merge(df_to_process, df_reparti, on=['Nome', 'Cognome'], how='left')
            df_to_process['Reparto'].fillna("Non Assegnato", inplace=True)
        except Exception as e:
            df_to_process['Reparto'] = "Non Assegnato"
            self.status_bar.showMessage(f"Foglio 'Reparto' non trovato o errore ({e}).", 7000)

        # >>> OTTIMIZZAZIONE: Chiama la funzione vettorizzata
        df_processed = self._analyze_data_vectorized(df_to_process)
        return df_processed

    def _analyze_data_vectorized(self, df):
        """Versione vettorizzata per l'analisi delle timbrature. Molto più veloce."""
        self.status_bar.showMessage("Analisi vettorizzata in corso...")
        QApplication.processEvents()

        # --- 1. Preparazione Dati ---
        ingresso_raw = df['Ingresso_t_raw']
        uscita_raw = df['Uscita_t_raw']

        # Arrotondamento vettorizzato
        df['Ingresso Contabile_t'] = self.round_time_vectorized(ingresso_raw, 'up')
        df['Uscita Contabile_t'] = self.round_time_vectorized(uscita_raw, 'down')

        # Calcolo ore contabili vettorizzato
        start_dt = pd.to_datetime(df['Data_dt'].dt.date.astype(str) + ' ' + df['Ingresso Contabile_t'].astype(str), errors='coerce')
        end_dt = pd.to_datetime(df['Data_dt'].dt.date.astype(str) + ' ' + df['Uscita Contabile_t'].astype(str), errors='coerce')
        # Gestione turni notturni (uscita il giorno dopo)
        end_dt = np.where(end_dt < start_dt, end_dt + pd.Timedelta(days=1), end_dt)
        df['Ore Contabili'] = (end_dt - start_dt).dt.total_seconds() / 3600
        df['Ore Contabili'].fillna(0, inplace=True)

        # --- 2. Creazione Maschere Booleane per Avvisi ---
        # Maschere di base
        has_both_times = ingresso_raw.notna() & uscita_raw.notna()
        has_error = pd.Series(False, index=df.index)

        # Avviso: Mancanze (priorità alta)
        m_ing_mancante = ingresso_raw.isna() & uscita_raw.notna()
        m_usc_mancante = ingresso_raw.notna() & uscita_raw.isna()
        m_entrambi_mancanti = ingresso_raw.isna() & uscita_raw.isna()
        m_mancanze = (m_ing_mancante | m_usc_mancante | m_entrambi_mancanti) if self.config_rules.get("alert_mancanze") else pd.Series(False, index=df.index)

        # Avviso: Invertiti (priorità massima, errore logico)
        m_invertiti = pd.Series(False, index=df.index)
        if self.config_rules.get("alert_invertiti"):
             m_invertiti = has_both_times & (df['Ore Contabili'] < 0)
             has_error |= m_invertiti
             df.loc[m_invertiti, 'Ore Contabili'] = 0 # Invalida ore se invertito

        # Avviso: Timbratura Ravvicinata
        min_ravv_h = self.config_rules.get("minuti_ravvicinata") / 60.0
        m_ravvicinata = has_both_times & ~has_error & (df['Ore Contabili'] >= 0) & (df['Ore Contabili'] < min_ravv_h)

        # Avviso: Turno Breve
        m_turno_breve = pd.Series(False, index=df.index)
        if self.config_rules.get("alert_turno_breve"):
             min_ore = self.config_rules.get("min_ore_valide")
             m_turno_breve = has_both_times & ~has_error & ~m_ravvicinata & (df['Ore Contabili'] < min_ore)

        # Avviso: Turno Esteso
        m_turno_esteso = pd.Series(False, index=df.index)
        if self.config_rules.get("alert_turno_esteso"):
             max_ore = self.config_rules.get("max_ore_normali")
             m_turno_esteso = has_both_times & ~has_error & (df['Ore Contabili'] > max_ore)

        # Avviso: Fuori Orario
        m_fuori_orario_ing = pd.Series(False, index=df.index)
        m_fuori_orario_usc = pd.Series(False, index=df.index)
        if self.config_rules.get("alert_fuori_orario"):
            inizio_std = pd.to_datetime(self.config_rules.get("orario_inizio_std")).time()
            fine_std = pd.to_datetime(self.config_rules.get("orario_fine_std")).time()
            m_fuori_orario_ing = df['Ingresso Contabile_t'].notna() & (df['Ingresso Contabile_t'] < inizio_std)
            m_fuori_orario_usc = df['Uscita Contabile_t'].notna() & (df['Uscita Contabile_t'] > fine_std)

        # --- 3. Assemblaggio Stringhe Avvisi ---
        # Crea colonne temporanee per ogni avviso, poi le concatena
        alerts = []
        if self.config_rules.get("alert_mancanze"):
            alerts.append(np.select([m_ing_mancante, m_usc_mancante, m_entrambi_mancanti],
                                    ["Ingresso Mancante", "Uscita Mancante", "Ingr./Usc. Mancanti"], default=""))
        if self.config_rules.get("alert_invertiti"):
             alerts.append(np.where(m_invertiti, "Uscita prima di Ingresso", ""))
        alerts.append(np.where(m_ravvicinata, f"Timbr. Ravvicinata (<{self.config_rules.get('minuti_ravvicinata')}min)", ""))
        if self.config_rules.get("alert_turno_breve"):
            alerts.append(np.where(m_turno_breve, f"Turno Troppo Breve (<{self.config_rules.get('min_ore_valide')}h)", ""))
        if self.config_rules.get("alert_turno_esteso"):
            alerts.append(np.where(m_turno_esteso, f"Turno Esteso (>{self.config_rules.get('max_ore_normali')}h)", ""))
        if self.config_rules.get("alert_fuori_orario"):
            alerts.append(np.where(m_fuori_orario_ing, "Ingr. Fuori Orario", ""))
            alerts.append(np.where(m_fuori_orario_usc, "Usc. Fuori Orario", ""))

        # Concatena tutti i messaggi di avviso
        df['Avvisi Sistema'] = pd.DataFrame(alerts).T.apply(lambda x: ', '.join(x[x != '']), axis=1)

        # --- 4. Assegnazione Highlight con np.select (molto veloce) ---
        conditions = [
            m_invertiti,
            m_mancanze,
            m_ravvicinata,
            m_turno_breve,
            m_turno_esteso,
            m_fuori_orario_ing | m_fuori_orario_usc
        ]
        choices = [ "ERRORE", "AVVISO", "AVVISO", "ATTENZIONE", "ATTENZIONE", "ATTENZIONE" ]
        df['Highlight'] = np.select(conditions, choices, default='NONE')

        return df


    def apply_filters(self):
        if self.df_original is None: return
        # Non serve clearare qui, viene fatto nel modello
        # self.checked_indices.clear()
        df_filtered = self.df_original

        # Filtro per avvisi
        if self.cb_filter_anomalies.isChecked():
            df_filtered = df_filtered[df_filtered['Highlight'] != "NONE"]

        # Filtro testuale
        search_term = self.search_bar.text().strip().lower()
        if search_term:
            # OTTIMIZZAZIONE: Cerca su colonne separate e unisci con OR
            mask = (df_filtered['Nome'].str.lower().str.contains(search_term, na=False) |
                    df_filtered['Cognome'].str.lower().str.contains(search_term, na=False) |
                    df_filtered['Sito'].str.lower().str.contains(search_term, na=False))
            df_filtered = df_filtered[mask]

        # Filtri a tendina
        selected_sito = self.sito_combo.currentText()
        if selected_sito != "Tutti i Siti":
            df_filtered = df_filtered[df_filtered['Sito'] == selected_sito]

        selected_reparto = self.reparto_combo.currentText()
        if selected_reparto != "Tutti i Reparti" and 'Reparto' in df_filtered.columns:
            df_filtered = df_filtered[df_filtered['Reparto'] == selected_reparto]

        # Filtro data
        date_from = self.date_from.date().toPyDate()
        date_to = self.date_to.date().toPyDate()
        if not df_filtered.empty:
             df_filtered = df_filtered[df_filtered['Data_dt'].dt.date.between(date_from, date_to)]

        self.update_table_view(df_filtered)


    def update_table_view(self, df):
        df_display = pd.DataFrame()
        display_columns_ordered = ['Sito', 'Reparto', 'Data', 'Nome', 'Cognome', 'Ingresso', 'Uscita',
                                   'Ingresso Contabile', 'Uscita Contabile', 'Ore Contabili',
                                   'Avvisi Sistema', 'Note Utente', 'Highlight', 'original_df_index']

        if not df.empty:
            df_display = df[['Sito', 'Reparto', 'Nome', 'Cognome', 'Ore Contabili', 'Avvisi Sistema', 'Highlight']].copy()
            df_display['Data'] = df['Data_dt'].dt.strftime('%d/%m/%Y')
            df_display['Ingresso'] = df['Ingresso_t_raw'].apply(lambda x: x.strftime('%H:%M') if pd.notna(x) else '')
            df_display['Uscita'] = df['Uscita_t_raw'].apply(lambda x: x.strftime('%H:%M') if pd.notna(x) else '')
            df_display['Ingresso Contabile'] = df['Ingresso Contabile_t'].apply(lambda x: x.strftime('%H:%M') if pd.notna(x) else '')
            df_display['Uscita Contabile'] = df['Uscita Contabile_t'].apply(lambda x: x.strftime('%H:%M') if pd.notna(x) else '')
            df_display['Note Utente'] = '' # Sarà riempito dal modello
            df_display['original_df_index'] = df.index
            df_display = df_display.reindex(columns=[c for c in display_columns_ordered if c in df_display.columns])
        else:
            df_display = pd.DataFrame(columns=[c for c in display_columns_ordered if c != 'original_df_index'])

        df_display.insert(0, 'Seleziona', False)

        pandas_model = PandasModel(df_display, self.checked_indices, self.user_notes, self)
        proxy_model = QSortFilterProxyModel(); proxy_model.setSourceModel(pandas_model)
        self.table_view.setModel(proxy_model)

        header = self.table_view.horizontalHeader()
        # Nascondi colonne tecniche
        for i, col_name in enumerate(df_display.columns):
            if col_name in ['Highlight', 'original_df_index']:
                self.table_view.setColumnHidden(i, True)
            # Adatta larghezza colonne in modo intelligente
            header.setSectionResizeMode(i, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(df_display.columns.get_loc('Avvisi Sistema'), QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(df_display.columns.get_loc('Note Utente'), QHeaderView.ResizeMode.Stretch)


    def open_report_dialog(self):
        if self.df_original is None or self.df_original.empty: QMessageBox.warning(self, "Dati non disponibili", "Nessun dato caricato."); return
        dialog = MonthlyReportDialog(self.df_original, self)
        if dialog.exec():
            month, year, employee_tuples = dialog.get_selection()
            if not employee_tuples: QMessageBox.warning(self, "Selezione Vuota", "Nessun dipendente selezionato."); return
            self.generate_monthly_report_pdf(month, year, employee_tuples)

    def generate_monthly_report_pdf(self, month, year, employees):
        path, _ = QFileDialog.getSaveFileName(self, "Salva Report Mensile", f"Report_{month}-{year}.pdf", "PDF Files (*.pdf)")
        if not path: return
        self.status_bar.showMessage("Generazione del report in corso...")
        df_month = self.df_original[(self.df_original['Data_dt'].dt.month == month) & (self.df_original['Data_dt'].dt.year == year)]
        doc = SimpleDocTemplate(path, pagesize=portrait(letter)); story = []; styles = getSampleStyleSheet()
        for i, (nome, cognome) in enumerate(employees):
            df_employee = df_month[(df_month['Nome'] == nome) & (df_month['Cognome'] == cognome)].sort_values(by='Data_dt')
            if df_employee.empty: continue
            if i > 0: story.append(PageBreak())
            story.append(Paragraph(f"Rapportino Mensile di: <b>{nome} {cognome}</b>", styles['h1']))
            story.append(Paragraph(f"Mese: <b>{month}/{year}</b>", styles['h2']))
            reparto_val = df_employee['Reparto'].iloc[0] if 'Reparto' in df_employee.columns and not df_employee['Reparto'].empty else ""
            story.append(Paragraph(f"Reparto: <b>{reparto_val}</b>", styles['Normal'])); story.append(Spacer(1, 0.2*inch))
            data_for_table = [['Data', 'Ingresso', 'Uscita', 'Ore Contabili', 'Avvisi Sistema', 'Note Utente']]
            for _, row in df_employee.iterrows():
                ore_contabili_str = f"{row['Ore Contabili']:.2f}".replace('.', ',')
                user_note_for_report = self.user_notes.get(row.name, "")
                data_for_table.append([
                    row['Data_dt'].strftime('%d/%m/%Y'),
                    row.get('Ingresso_t_raw', pd.NaT).strftime('%H:%M') if pd.notna(row.get('Ingresso_t_raw')) else '',
                    row.get('Uscita_t_raw', pd.NaT).strftime('%H:%M') if pd.notna(row.get('Uscita_t_raw')) else '',
                    ore_contabili_str,
                    row.get('Avvisi Sistema', ''),
                    user_note_for_report
                ])
            total_hours = df_employee['Ore Contabili'].sum(); total_days = len(df_employee)
            total_hours_str = f"{total_hours:.2f}".replace('.', ',')
            story.append(Paragraph(f"<b>Totale Giorni Lavorati:</b> {total_days}", styles['Normal']))
            story.append(Paragraph(f"<b>Totale Ore Contabili:</b> {total_hours_str}", styles['Normal'])); story.append(Spacer(1, 0.2*inch))
            table = Table(data_for_table, colWidths=[1*inch, 1*inch, 1*inch, 1*inch, 1.5*inch, 1.5*inch])
            table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#0078D7')), ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke), ('ALIGN',(0,0),(-1,-1),'CENTER'), ('GRID',(0,0),(-1,-1),1,colors.darkgrey), ('FONTSIZE', (0,0), (-1,-1), 8)]))
            story.append(table)
        doc.build(story)
        self.status_bar.showMessage(f"Report generato con successo: {path}", 5000)

    def export_to_csv(self): self.export_selected_data('csv')
    def export_to_pdf(self): self.export_selected_data('pdf')

    def export_selected_data(self, format_type):
        if not self.checked_indices:
            QMessageBox.information(self, "Esportazione", "Nessuna riga selezionata."); return

        # Prendi le righe complete da self.df_original usando gli indici selezionati e validi
        df_to_export = self.df_original.loc[list(self.checked_indices)].copy()
        if df_to_export.empty:
            QMessageBox.warning(self, "Esportazione", "Le righe selezionate non sono valide (filtri cambiati?). Riprova la selezione."); return

        # Aggiungi le note utente
        df_to_export['Note Utente'] = df_to_export.index.map(self.user_notes).fillna('')

        # Seleziona, rinomina e ordina le colonne per l'esportazione
        export_cols = {
            'Sito': 'Sito', 'Reparto': 'Reparto', 'Data_dt': 'Data', 'Nome': 'Nome', 'Cognome': 'Cognome',
            'Ingresso_t_raw': 'Ingresso', 'Uscita_t_raw': 'Uscita',
            'Ingresso Contabile_t': 'Ingresso Contabile', 'Uscita Contabile_t': 'Uscita Contabile',
            'Ore Contabili': 'Ore Contabili', 'Avvisi Sistema': 'Avvisi Sistema', 'Note Utente': 'Note Utente'
        }
        df_final_export = df_to_export[export_cols.keys()].rename(columns=export_cols)

        # Formatta i dati per la leggibilità
        df_final_export['Data'] = df_final_export['Data'].dt.strftime('%d/%m/%Y')
        for col in ['Ingresso', 'Uscita', 'Ingresso Contabile', 'Uscita Contabile']:
            df_final_export[col] = df_final_export[col].apply(lambda x: x.strftime('%H:%M') if pd.notna(x) else '')
        df_final_export['Ore Contabili'] = df_final_export['Ore Contabili'].apply(lambda x: f"{x:.2f}".replace('.', ','))

        path, _ = QFileDialog.getSaveFileName(self, f"Salva come {format_type.upper()}", f"export_selezionati.{format_type}", "CSV Files (*.csv)" if format_type == 'csv' else "PDF Files (*.pdf)")
        if not path: return
        try:
            if format_type == 'csv':
                df_final_export.to_csv(path, index=False, sep=';', encoding='utf-8-sig')
            else:
                doc_width, doc_height = landscape(letter); margin = 0.5 * inch; available_width = doc_width - (2 * margin)
                col_widths = [available_width / len(df_final_export.columns)] * len(df_final_export.columns) if df_final_export.columns.size > 0 else []
                doc = SimpleDocTemplate(path, pagesize=landscape(letter), leftMargin=margin, rightMargin=margin)
                data_list = [df_final_export.columns.tolist()] + df_final_export.astype(str).values.tolist()
                table = Table(data_list, colWidths=col_widths)
                table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#0078D7')),('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),('ALIGN',(0,0),(-1,-1),'CENTER'),('GRID',(0,0),(-1,-1),1,colors.darkgrey),('WORDWRAP',(0,0),(-1,-1),'CJK'),('FONTSIZE', (0,0), (-1,-1), 8)]))
                doc.build([table])
            self.status_bar.showMessage(f"Dati esportati con successo in {path}", 5000)
        except Exception as e:
            QMessageBox.critical(self, "Errore Esportazione", f"Impossibile salvare il file.\nErrore: {e}")

    @staticmethod
    def round_time_vectorized(series, direction='up'):
        """Arrotonda una Series di 'time' al quarto d'ora più vicino."""
        # Converte in minuti totali dal giorno per il calcolo
        minutes = series.apply(lambda t: t.hour * 60 + t.minute if pd.notna(t) else np.nan)
        minutes.dropna(inplace=True)

        if direction == 'up':
            rounded_minutes = (minutes // 15 + (minutes % 15 > 0)) * 15
        else: # down
            rounded_minutes = (minutes // 15) * 15

        # Riconverte in 'time'
        new_hour = (rounded_minutes // 60) % 24
        new_minute = rounded_minutes % 60
        
        # Crea una series di time objects
        time_series = pd.Series([time(int(h), int(m)) if pd.notna(h) else pd.NaT for h, m in zip(new_hour, new_minute)], index=new_hour.index)
        
        # Riunisce i valori NaT originali
        return time_series.reindex(series.index)


    def setup_filters(self):
        if self.df_original is None: return
        siti = sorted(self.df_original['Sito'].dropna().unique())
        self.sito_combo.blockSignals(True); self.sito_combo.clear(); self.sito_combo.addItems(["Tutti i Siti"] + siti); self.sito_combo.blockSignals(False)
        if 'Reparto' in self.df_original.columns:
            reparti = sorted(self.df_original['Reparto'].dropna().unique())
            self.reparto_combo.blockSignals(True); self.reparto_combo.clear(); self.reparto_combo.addItems(["Tutti i Reparti"] + reparti); self.reparto_combo.blockSignals(False); self.reparto_combo.setEnabled(True)
        else:
            self.reparto_combo.blockSignals(True); self.reparto_combo.clear(); self.reparto_combo.addItem("N/D"); self.reparto_combo.blockSignals(False); self.reparto_combo.setEnabled(False)

        if self.df_original.empty or self.df_original['Data_dt'].isna().all():
            min_date = max_date = datetime.now().date()
        else:
            min_date = self.df_original['Data_dt'].min().date(); max_date = self.df_original['Data_dt'].max().date()
        min_qdate = QDate(min_date.year, min_date.month, min_date.day); max_qdate = QDate(max_date.year, max_date.month, max_date.day)
        self.date_from.setMinimumDate(min_qdate); self.date_from.setMaximumDate(max_qdate); self.date_to.setMinimumDate(min_qdate); self.date_to.setMaximumDate(max_qdate)
        self.set_date_range(min_date, max_date)

    def set_date_range(self, start, end):
        self.date_from.blockSignals(True); self.date_to.blockSignals(True)
        self.date_from.setDate(QDate(start.year, start.month, start.day)); self.date_to.setDate(QDate(end.year, end.month, end.day))
        self.date_from.blockSignals(False); self.date_to.blockSignals(False); self.apply_filters()

    def filter_yesterday(self): today = datetime.now().date(); self.set_date_range(today - timedelta(days=1), today - timedelta(days=1))
    def filter_this_week(self): today = datetime.now().date(); start_of_week = today - timedelta(days=today.weekday()); self.set_date_range(start_of_week, start_of_week + timedelta(days=6))
    def filter_this_month(self): today = datetime.now().date(); start_of_month = today.replace(day=1); _, num_days = calendar.monthrange(today.year, today.month); self.set_date_range(start_of_month, today.replace(day=num_days))

    def reset_all_filters(self):
        self.search_bar.blockSignals(True); self.search_bar.clear(); self.search_bar.blockSignals(False)
        self.sito_combo.blockSignals(True); self.sito_combo.setCurrentIndex(0); self.sito_combo.blockSignals(False)
        if hasattr(self, 'reparto_combo'): self.reparto_combo.blockSignals(True); self.reparto_combo.setCurrentIndex(0); self.reparto_combo.blockSignals(False)
        self.cb_filter_anomalies.blockSignals(True); self.cb_filter_anomalies.setChecked(False); self.cb_filter_anomalies.blockSignals(False)
        self.setup_filters(); self.status_bar.showMessage("Filtri resettati.", 3000)

    def show_help_guide_dialog(self): dialog = HelpGuideDialog(self); dialog.exec()
    def show_about_dialog(self): QMessageBox.about(self, "Informazioni", "<b>ISAB Sud - Control & Report v9.1 (Ottimizzata)</b><br>Applicazione per l'analisi avanzata delle timbrature.<br><br>Sviluppata con Python e PyQt6.<br>Ottimizzata da un assistente AI di Google.")
    def save_window_settings(self): settings = QSettings("MyCompany", "TimbratureApp_v9"); settings.setValue("geometry", self.saveGeometry()); settings.setValue("windowState", self.saveState())
    def load_window_settings(self):
        settings = QSettings("MyCompany", "TimbratureApp_v9"); geometry = settings.value("geometry")
        if geometry: self.restoreGeometry(geometry)
        state = settings.value("windowState")
        if state: self.restoreState(state)
        else: self.resize(1600, 900)
    def closeEvent(self, event): self.save_window_settings(); self.save_user_notes(); event.accept()

# Classe MonthlyReportDialog (invariata)
class MonthlyReportDialog(QDialog):
    def __init__(self, df, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Genera Report Mensile"); layout = QVBoxLayout(self)
        form_layout = QHBoxLayout()
        current_date = date.today()
        self.month_spin = QSpinBox(); self.month_spin.setRange(1, 12); self.month_spin.setValue(current_date.month)
        self.year_spin = QSpinBox(); self.year_spin.setRange(2020, 2050); self.year_spin.setValue(current_date.year)
        form_layout.addWidget(QLabel("Mese:")); form_layout.addWidget(self.month_spin); form_layout.addWidget(QLabel("Anno:")); form_layout.addWidget(self.year_spin)
        layout.addLayout(form_layout); layout.addWidget(QLabel("Seleziona i dipendenti per il report:"))
        self.employee_list = QListWidget(); self.employee_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        employees = sorted(df[['Nome', 'Cognome']].drop_duplicates().to_records(index=False))
        for nome, cognome in employees:
            self.employee_list.addItem(f"{nome} {cognome}")
        layout.addWidget(self.employee_list)
        select_all_button = QPushButton("Seleziona Tutti"); select_all_button.clicked.connect(self.employee_list.selectAll)
        deselect_all_button = QPushButton("Deseleziona Tutti"); deselect_all_button.clicked.connect(self.employee_list.clearSelection)
        btn_layout = QHBoxLayout(); btn_layout.addWidget(select_all_button); btn_layout.addWidget(deselect_all_button)
        layout.addLayout(btn_layout)
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept); self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

    def get_selection(self):
        month = self.month_spin.value(); year = self.year_spin.value(); selected_items = self.employee_list.selectedItems()
        employee_tuples = []
        for item in selected_items:
            # Gestisce cognomi con spazi
            parts = item.text().split(' ', 1)
            nome = parts[0]
            cognome = parts[1] if len(parts) > 1 else ""
            employee_tuples.append((nome, cognome))
        return month, year, employee_tuples

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(LIGHT_STYLE)
    window = TimbratureApp()
    window.show()
    sys.exit(app.exec())