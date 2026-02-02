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
    prefs = {
        "download.default_directory": str(Path(DOWNLOAD_DIR).absolute()),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True, 
        "safebrowsing.disable_download_protection": True,
        "profile.default_content_settings.popups": 0,
        "profile.content_settings.exceptions.automatic_downloads.*.setting": 1
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    # Argomenti per disabilitare le nuove feature di sicurezza di Chrome (Bubble, Warnings)
    chrome_options.add_argument("--disable-features=InsecureDownloadWarnings")
    chrome_options.add_argument("--disable-features=DownloadBubble,DownloadBubbleV2")
    
    # Altri argomenti permissivi
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--log-level=3")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--allow-running-insecure-content")
    chrome_options.add_argument("--disable-web-security")
    chrome_options.add_argument(f"--unsafely-treat-insecure-origin-as-secure={LOGIN_URL}")

    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 20)

    logger.info(f"Navigazione a: {LOGIN_URL}")
    driver.get(LOGIN_URL)

    logger.info("Login in corso...")
    wait.until(EC.presence_of_element_located((By.NAME, "Username"))).send_keys(USERNAME)
    wait.until(EC.presence_of_element_located((By.NAME, "Password"))).send_keys(PASSWORD)
    wait.until(EC.element_to_be_clickable((By.ID, "round_button-1017-btnInnerEl"))).click()
    
    attendi_scomparsa_overlay(driver, 60)

    # Gestione Popup "Sessione attiva"
    try:
        wait_popup = WebDriverWait(driver, 5)
        if len(driver.find_elements(By.ID, "messagebox-1001-msg")) > 0 and driver.find_element(By.ID, "messagebox-1001-msg").is_displayed():
             logger.info("Rilevata sessione attiva. Clicco su 'Si'...")
             wait_popup.until(EC.element_to_be_clickable((By.ID, "button-1006-btnInnerEl"))).click()
             attendi_scomparsa_overlay(driver, 60)
    except Exception as e:
        # Ignora errori se il popup non c'è, è opzionale
        pass

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
    # Utilizzo un XPath più specifico per il trigger del fornitore, simile a scaricaTimbratureIsab
    fornitore_trigger_xpath = "//input[@name='CodiceFornitore']/ancestor::div[contains(@class, 'x-form-trigger-wrap')]//div[contains(@class, 'x-form-arrow-trigger')]"
    
    try:
        # Attendo che il trigger sia visibile e cliccabile
        trigger = wait.until(EC.element_to_be_clickable((By.XPATH, fornitore_trigger_xpath)))
        ActionChains(driver).move_to_element(trigger).click().perform()
        logger.info(" -> Trigger fornitore cliccato.")
    except Exception:
        logger.warning(" -> Trigger specifico non trovato, provo quello generico...")
        wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'x-form-arrow-trigger')]"))).click()

    attendi_scomparsa_overlay(driver)
    
    # Selezione dell'opzione dalla lista
    opt_xpath = f"//li[normalize-space(text())='{PROVIDER}']"
    opt = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, opt_xpath)))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'}); arguments[0].click();", opt)
    
    logger.info(f" -> Fornitore '{PROVIDER}' selezionato.")
    attendi_scomparsa_overlay(driver)

    # Verifica se il campo è stato effettivamente popolato (opzionale ma utile)
    try:
        valore_input = driver.find_element(By.NAME, "CodiceFornitore").get_attribute("value")
        if not valore_input:
            logger.warning(" -> ATTENZIONE: Il campo Fornitore risulta ancora vuoto! Provo inserimento manuale...")
            campo_f = driver.find_element(By.NAME, "CodiceFornitore")
            campo_f.send_keys(PROVIDER)
            time.sleep(1)
            driver.execute_script("arguments[0].dispatchEvent(new Event('change', {bubbles:true}));", campo_f)
    except:
        pass

    logger.info(f"Impostazione Data: {DATE_TO_INSERT}")
    campo_d = wait.until(EC.visibility_of_element_located((By.NAME, "DataTimesheetDa")))
    campo_d.clear()
    campo_d.send_keys(DATE_TO_INSERT)

    js_ev = "var e=new Event('change',{bubbles:true}); arguments[0].dispatchEvent(e);"

    all_downloads_ok = True

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
        
        btn_dl_xpath = "//div[contains(@class, 'x-tool')]//div[contains(@style, 'FontAwesome')]"
        btn_dl_elem = wait.until(EC.element_to_be_clickable((By.XPATH, btn_dl_xpath)))
        
        # Uso JS click per evitare ElementClickInterceptedException se ci sono overlay/pulsanti sopra
        driver.execute_script("arguments[0].click();", btn_dl_elem)
        
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
            # Rinomina solo con ODC (numero OdA) come richiesto
            name = f"{n}.xlsx"
            dest = Path(MOVE_DIR)
            dest.mkdir(parents=True, exist_ok=True)
            
            f_path = dest / name
            
            # Logica robusta di spostamento con retry e sovrascrittura
            moved_ok = False
            for attempt in range(5):
                try:
                    # Attesa iniziale/tra i tentativi per permettere il rilascio del file da parte di Chrome/Antivirus
                    time.sleep(2) 
                    
                    # Se il file esiste già, provo a rimuoverlo per permettere la sovrascrittura
                    if f_path.exists():
                        try:
                            os.remove(f_path)
                            logger.info(f" -> File esistente rimosso per sovrascrittura: {name}")
                        except:
                            # Se fallisce la rimozione (es file aperto), shutil.move proverà comunque a sovrascrivere
                            pass

                    shutil.move(str(found), str(f_path))
                    logger.info(f" -> File salvato: {f_path.name}")
                    moved_ok = True
                    break
                except (PermissionError, OSError) as e:
                    logger.warning(f" -> File bloccato o errore spostamento ({e}). Riprovo ({attempt+1}/5)...")
            
            if not moved_ok:
                logger.error(f" -> ERRORE: Impossibile spostare il file {found.name} dopo 5 tentativi.")
                all_downloads_ok = False
                break
        else:
            logger.error(f" -> ERRORE: Download non riuscito per OdA {n}. Interrompo l'elaborazione.")
            all_downloads_ok = False
            break

    logger.info("--- OPERAZIONI WEB COMPLETATE ---")

except Exception:
    logger.error("ERRORE DURANTE L'ESECUZIONE:")
    logger.error(traceback.format_exc())
    all_downloads_ok = False
finally:
    if driver:
        driver.quit()

logger.info("Fine Script.")
