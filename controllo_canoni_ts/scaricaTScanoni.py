from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
import time
import traceback
import sys
import os
from pathlib import Path
import shutil
import gc
import logging
import json

# --- CONFIGURAZIONE LOGGING (Standard Output per la GUI) ---
logging.basicConfig(
    level=logging.INFO, 
    format="%(asctime)s - %(message)s", 
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger(__name__)

# Forza l'output non bufferizzato per la comunicazione con la GUI
sys.stdout.reconfigure(line_buffering=True)

try:
    import win32com.client
    PYWIN32_AVAILABLE = True
except ImportError:
    PYWIN32_AVAILABLE = False

def attendi_scomparsa_overlay(driver, timeout_secondi=45):
    try:
        overlay_wait = WebDriverWait(driver, timeout_secondi)
        xpath_mask = "//div[contains(@class, 'x-mask-msg') or contains(@class, 'x-mask')][not(contains(@style,'display: none'))]"
        xpath_text = "//div[text()='Caricamento...']"
        overlay_wait.until(EC.invisibility_of_element_located((By.XPATH, f"{xpath_mask} | {xpath_text}")))
        logger.info(" -> Overlay scomparso.")
        return True
    except Exception:
        return False

# --- CARICAMENTO CONFIG ---
SCRIPT_DIR = Path(__file__).resolve().parent
CONFIG_FILE = SCRIPT_DIR / "config_canoni.json"

if not CONFIG_FILE.exists():
    logger.error("ERRORE: config_canoni.json non trovato!")
    sys.exit(1)

with open(CONFIG_FILE, "r", encoding="utf-8") as f:
    config = json.load(f)

# Variabili
USERNAME = config.get("username")
PASSWORD = config.get("password")
DOWNLOAD_DIR = config.get("download_dir")
MOVE_DIR = config.get("move_dir")
LOGIN_URL = config.get("login_url", "https://portalefornitori.isab.com/Ui/")
PROVIDER = config.get("provider", "KK10608 - COEMI S.R.L.")
DATE_TO_INSERT = config.get("date_to_insert", "01.01.2025")
MACRO_PATH = config.get("macro_file_path")
RUN_MACRO = config.get("run_macro", False)
ORDERS = config.get("orders", [])

logger.info("--- AVVIO ROBOT SELENIUM ---")

driver = None
try:
    chrome_options = webdriver.ChromeOptions()
    prefs = {"download.default_directory": DOWNLOAD_DIR, "download.prompt_for_download": False}
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--log-level=3") # Meno rumore nei log

    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 20)

    logger.info(f"Navigazione a: {LOGIN_URL}")
    driver.get(LOGIN_URL)

    logger.info("Login in corso...")
    wait.until(EC.presence_of_element_located((By.NAME, "Username"))).send_keys(USERNAME)
    wait.until(EC.presence_of_element_located((By.NAME, "Password"))).send_keys(PASSWORD)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Accedi']"))).click()
    
    attendi_scomparsa_overlay(driver, 60)

    try:
        wait_ok = WebDriverWait(driver, 5)
        wait_ok.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='OK']"))).click()
        logger.info("Pop-up OK gestito.")
    except:
        pass

    logger.info("Navigazione: Report -> Timesheet")
    wait.until(EC.element_to_be_clickable((By.XPATH, "//*[normalize-space(text())='Report']"))).click()
    attendi_scomparsa_overlay(driver)
    
    btn_ts = "//span[contains(@id, 'generic_menu_button')][.//span[text()='Timesheet']]"
    wait.until(EC.element_to_be_clickable((By.XPATH, btn_ts))).click()
    attendi_scomparsa_overlay(driver)

    logger.info(f"Selezione Fornitore: {PROVIDER}")
    wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'x-form-arrow-trigger')]"))).click()
    opt = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f"//li[normalize-space(text())='{PROVIDER}']")))
    driver.execute_script("arguments[0].scrollIntoView(); arguments[0].click();", opt)
    attendi_scomparsa_overlay(driver)

    logger.info(f"Impostazione Data: {DATE_TO_INSERT}")
    campo_d = wait.until(EC.visibility_of_element_located((By.NAME, "DataTimesheetDa")))
    campo_d.clear()
    campo_d.send_keys(DATE_TO_INSERT)

    js_ev = "var e=new Event('change',{bubbles:true}); arguments[0].dispatchEvent(e);"

    # Loop
    for o in ORDERS:
        n, p = o.get("numero"), o.get("posizione", "")
        if not n: continue

        logger.info(f"Elaborazione OdA {n} (Pos: {p})...")
        c_n = wait.until(EC.presence_of_element_located((By.NAME, "NumeroOda")))
        driver.execute_script(f"arguments[0].value='{n}';", c_n)
        driver.execute_script(js_ev, c_n)

        c_p = wait.until(EC.presence_of_element_located((By.NAME, "PosizioneOda")))
        driver.execute_script(f"arguments[0].value='{p}';", c_p)
        driver.execute_script(js_ev, c_p)

        wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Cerca']"))).click()
        attendi_scomparsa_overlay(driver, 90)

        # Download
        p_dl = Path(DOWNLOAD_DIR)
        start_files = set(p_dl.iterdir())
        
        btn_dl = "//div[contains(@class, 'x-tool')]//div[contains(@style, 'FontAwesome')]"
        wait.until(EC.element_to_be_clickable((By.XPATH, btn_dl))).click()
        
        found = None
        for _ in range(60):
            current = set(p_dl.iterdir())
            diff = current - start_files
            files = [f for f in diff if f.suffix.lower() == '.xlsx' and not f.name.endswith('.tmp')]
            if files:
                found = max(files, key=lambda f: f.stat().st_mtime)
                break
            time.sleep(0.5)

        if found:
            sfx = f"-{p}" if p else ""
            name = f"{n}{sfx}.xlsx"
            dest = Path(MOVE_DIR)
            dest.mkdir(parents=True, exist_ok=True)
            
            f_path = dest / name
            if f_path.exists():
                f_path = dest / f"{n}{sfx}_{int(time.time())}.xlsx"
            
            shutil.move(str(found), str(f_path))
            logger.info(f" -> File salvato: {f_path.name}")
        else:
            logger.warning(f" -> Download non riuscito per OdA {n}")

    logger.info("--- OPERAZIONI WEB COMPLETATE ---")

except Exception:
    logger.error("ERRORE DURANTE L'ESECUZIONE:")
    logger.error(traceback.format_exc())
finally:
    if driver:
        driver.quit()

# --- MACRO ---
if RUN_MACRO and PYWIN32_AVAILABLE and MACRO_PATH:
    try:
        logger.info(f"Avvio Macro Excel: {Path(MACRO_PATH).name}")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True # Visibile per monitorare l'avanzamento
        wb = excel.Workbooks.Open(str(Path(MACRO_PATH).resolve()))
        excel.Run("elaboraTutto")
        wb.Close(SaveChanges=True)
        excel.Quit()
        logger.info("Macro completata con successo.")
    except Exception:
        logger.error(f"Errore Macro: {traceback.format_exc()}")

logger.info("Fine Script.")
