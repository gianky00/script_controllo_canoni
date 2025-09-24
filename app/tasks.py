import os
import time
import logging
from datetime import datetime, timedelta
from pathlib import Path
import shutil
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains

from app import db, create_app
from app.models import Setting, Timbratura, Oda

logger = logging.getLogger(__name__)


def get_setting(key):
    """Helper function to get a setting value."""
    setting = Setting.query.filter_by(key=key).first()
    if not setting or not setting.value:
        logger.error(f"Required setting '{key}' is not configured.")
        return None
    return setting.value

def attendi_scomparsa_overlay(driver, timeout_secondi=45):
    """Waits for loading overlays to disappear."""
    try:
        overlay_wait = WebDriverWait(driver, timeout_secondi)
        xpath_overlay = "//div[contains(@class, 'x-mask-msg') or contains(@class, 'x-mask')][not(contains(@style,'display: none'))] | //div[text()='Caricamento...']"
        overlay_wait.until(EC.invisibility_of_element_located((By.XPATH, xpath_overlay)))
        logger.info(" -> Loading overlay disappeared.")
        time.sleep(0.3)
        return True
    except TimeoutException:
        logger.warning(f"Timeout ({timeout_secondi}s) waiting for overlay to disappear. Proceeding with caution.")
        return False

def run_scarica_timbrature():
    """
    Refactored task to download 'Timbrature' data.
    Wraps the core logic in an app context to allow DB access.
    """
    app = create_app()
    with app.app_context():
        logger.info("Starting 'run_scarica_timbrature' task...")

        # --- 1. Get Configuration from Database ---
        from app.security import decrypt_password

        required_settings = {
            'LOGIN_URL': None, 'USERNAME': None, 'PASSWORD': None,
            'DOWNLOAD_DIR': None, 'FORNITORE_DA_SELEZIONARE': None
        }

        for key in required_settings:
            required_settings[key] = get_setting(key)
            if required_settings[key] is None:
                logger.critical(f"Task aborted: Missing required setting '{key}'.")
                return

        password = decrypt_password(required_settings['PASSWORD'])
        if not password:
            logger.critical("Task aborted: Password could not be decrypted.")
            return

        # --- 2. Selenium Web Automation ---
        driver = None
        final_downloaded_path = None
        try:
            yesterday = datetime.now() - timedelta(days=1)
            data_da_usare = yesterday.strftime('%d.%m.%Y')
            logger.info(f"Calculated date for search (yesterday): {data_da_usare}")

            chrome_options = webdriver.ChromeOptions()
            prefs = {"download.default_directory": required_settings['DOWNLOAD_DIR'], "download.prompt_for_download": False}
            chrome_options.add_experimental_option("prefs", prefs)
            chrome_options.add_argument("--headless")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--start-maximized")

            logger.info("Initializing Chrome WebDriver...")
            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 20)
            long_wait = WebDriverWait(driver, 30)

            logger.info(f"Navigating to: {required_settings['LOGIN_URL']}")
            driver.get(required_settings['LOGIN_URL'])

            logger.info("Attempting login...")
            wait.until(EC.presence_of_element_located((By.NAME, "Username"))).send_keys(required_settings['USERNAME'])
            wait.until(EC.presence_of_element_located((By.NAME, "Password"))).send_keys(password)
            wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Accedi' and contains(@class, 'x-btn-inner')]"))).click()

            logger.info("Verifying login success...")
            attendi_scomparsa_overlay(driver, 60)

            # --- Real Selenium Logic ---
            wait.until(EC.element_to_be_clickable((By.XPATH, "//*[normalize-space(text())='Report']"))).click()
            attendi_scomparsa_overlay(driver)

            wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Timbrature']"))).click()
            logger.info("'Timbrature' clicked.")
            attendi_scomparsa_overlay(driver)

            fornitore_arrow_xpath = "//div[starts-with(@id, 'generic_refresh_combo_box-') and contains(@id, '-trigger-picker')]"
            wait.until(EC.visibility_of_element_located((By.XPATH, fornitore_arrow_xpath)))

            logger.info(f"  Selecting provider: '{required_settings['FORNITORE_DA_SELEZIONARE']}'...")
            ActionChains(driver).move_to_element(wait.until(EC.element_to_be_clickable((By.XPATH, fornitore_arrow_xpath)))).click().perform()
            attendi_scomparsa_overlay(driver)

            fornitore_option_xpath = f"//li[normalize-space(text())='{required_settings['FORNITORE_DA_SELEZIONARE']}']"
            fornitore_option = long_wait.until(EC.presence_of_element_located((By.XPATH, fornitore_option_xpath)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'}); arguments[0].click();", fornitore_option)
            logger.info(f"  Provider '{required_settings['FORNITORE_DA_SELEZIONARE']}' selected.")
            attendi_scomparsa_overlay(driver)

            logger.info(f"  Entering date From/To: '{data_da_usare}'...")
            wait.until(EC.visibility_of_element_located((By.NAME, "DataTsDa"))).clear()
            driver.find_element(By.NAME, "DataTsDa").send_keys(data_da_usare)
            wait.until(EC.visibility_of_element_located((By.NAME, "DataTsA"))).clear()
            driver.find_element(By.NAME, "DataTsA").send_keys(data_da_usare)
            logger.info("  Dates entered.")

            logger.info("  Clicking 'Search' button...")
            cerca_button_xpath = "//a[contains(@class, 'x-btn') and .//span[normalize-space(text())='Cerca']]"
            wait.until(EC.element_to_be_clickable((By.XPATH, cerca_button_xpath))).click()
            logger.info("  'Search' button clicked. Waiting for results...")
            attendi_scomparsa_overlay(driver, 90)

            logger.info("  Attempting to download the Excel file...")
            path_to_downloads_obj = Path(required_settings['DOWNLOAD_DIR'])
            files_before_download = set(path_to_downloads_obj.iterdir())

            excel_button_xpath = "//div[contains(@class, 'x-tool') and @role='button' and .//div[contains(@style, 'FontAwesome')]]"
            wait.until(EC.element_to_be_clickable((By.XPATH, excel_button_xpath))).click()
            logger.info("  Excel download icon clicked. Waiting for download to complete (max 45s)...")

            download_start_time = time.time()
            while time.time() - download_start_time < 45:
                current_files = set(path_to_downloads_obj.iterdir())
                new_files = current_files - files_before_download
                completed_files = [f for f in new_files if f.suffix.lower() == '.xlsx' and not f.name.endswith(('.crdownload', '.tmp'))]
                if completed_files:
                    final_downloaded_path = max(completed_files, key=lambda f: f.stat().st_mtime)
                    time.sleep(1)
                    if final_downloaded_path.exists() and final_downloaded_path.stat().st_size > 0:
                        logger.info(f"  Download COMPLETE. File detected: {final_downloaded_path.name}")
                        break
                    else:
                        final_downloaded_path = None
                time.sleep(1)

            if not final_downloaded_path:
                logger.critical("CRITICAL ERROR: Download of the timesheet report failed or the file is invalid.")
                raise Exception("Download failed")

            logger.info("--- STARTING EXCEL FILE PROCESSING ---")
            process_timbrature_file(final_downloaded_path)

            logger.info("'run_scarica_timbrature' task finished successfully.")

        except Exception as e:
            logger.critical(f"A general error occurred during Selenium operations: {e}", exc_info=True)
        finally:
            if driver:
                logger.info("Closing WebDriver.")
                driver.quit()
            if final_downloaded_path and final_downloaded_path.exists():
                try:
                    os.remove(final_downloaded_path)
                    logger.info(f"Temporary file '{final_downloaded_path.name}' deleted.")
                except OSError as e_remove:
                    logger.warning(f"Could not delete downloaded file: {e_remove}")


def process_timbrature_file(file_path):
    """Reads an Excel file with timekeeping data and inserts it into the database."""
    try:
        logger.info(f"Opening downloaded file: {file_path.name}")
        df_source = pd.read_excel(file_path)

        column_map = {
            'Sito': 'sito', 'Reparto': 'reparto', 'Data': 'data', 'Nome': 'nome',
            'Cognome': 'cognome', 'Ingresso': 'ingresso', 'Uscita': 'uscita',
            'Ingresso Contabile': 'ingresso_contabile', 'Uscita Contabile': 'uscita_contabile',
            'Ore Contabili': 'ore_contabili', 'Avvisi Sistema': 'avvisi_sistema',
            'Note Utente': 'note_utente'
        }

        df_source.rename(columns=column_map, inplace=True)
        existing_cols = [col for col in column_map.values() if col in df_source.columns]
        df_to_insert = df_source[existing_cols]

        if 'data' in df_to_insert:
            df_to_insert['data'] = pd.to_datetime(df_to_insert['data']).dt.date
        for time_col in ['ingresso', 'uscita', 'ingresso_contabile', 'uscita_contabile']:
            if time_col in df_to_insert:
                df_to_insert[time_col] = pd.to_datetime(df_to_insert[time_col], errors='coerce').dt.time

        if 'data' in df_to_insert and 'nome' in df_to_insert and 'cognome' in df_to_insert:
            df_to_insert.dropna(subset=['data', 'nome', 'cognome'], inplace=True)

            dates_in_df = df_to_insert['data'].unique()
            existing_records_query = db.session.query(Timbratura.data, Timbratura.nome, Timbratura.cognome).filter(Timbratura.data.in_(dates_in_df))
            existing_records = pd.read_sql(existing_records_query.statement, db.engine)
            if not existing_records.empty:
                existing_records['data'] = pd.to_datetime(existing_records['data']).dt.date
                merged = pd.merge(df_to_insert, existing_records, on=['data', 'nome', 'cognome'], how='left', indicator=True)
                df_new = merged[merged['_merge'] == 'left_only'].drop(columns=['_merge'])
            else:
                df_new = df_to_insert
        else:
            df_new = df_to_insert

        rows_to_add = len(df_new)
        if rows_to_add > 0:
            logger.info(f"Adding {rows_to_add} new unique rows to the database.")
            df_new.to_sql('timbrature', db.engine, if_exists='append', index=False)
            logger.info("Database insertion complete.")
        else:
            logger.info("No new rows to add. All downloaded data already exists in the database.")

    except Exception as e:
        logger.critical(f"CRITICAL ERROR during Excel file processing: {e}", exc_info=True)

def run_scarica_canoni():
    """
    Refactored task to download 'Canoni' data.
    Wraps the core logic in an app context to allow DB access.
    """
    app = create_app()
    with app.app_context():
        logger.info("Starting 'run_scarica_canoni' task...")

        # --- 1. Get Configuration from Database ---
        from app.security import decrypt_password

        required_settings = {
            'LOGIN_URL': None, 'USERNAME': None, 'PASSWORD': None, 'DOWNLOAD_DIR': None,
            'DIR_SPOSTAMENTO_TS': None, 'FORNITORE_DA_SELEZIONARE': None, 'DATA_DA_INSERIRE': None
        }

        for key in required_settings:
            required_settings[key] = get_setting(key)
            if required_settings[key] is None:
                logger.critical(f"Task aborted: Missing required setting '{key}'.")
                return

        password = decrypt_password(required_settings['PASSWORD'])
        if not password:
            logger.critical("Task aborted: Password could not be decrypted.")
            return

        odas = Oda.query.all()

        # --- 2. Selenium Web Automation ---
        driver = None
        try:
            chrome_options = webdriver.ChromeOptions()
            prefs = {"download.default_directory": required_settings['DOWNLOAD_DIR'], "download.prompt_for_download": False}
            chrome_options.add_experimental_option("prefs", prefs)
            chrome_options.add_argument("--headless")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--start-maximized")

            logger.info("Initializing Chrome WebDriver...")
            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 20)
            long_wait = WebDriverWait(driver, 30)

            logger.info(f"Navigating to: {required_settings['LOGIN_URL']}")
            driver.get(required_settings['LOGIN_URL'])

            logger.info("Attempting login...")
            wait.until(EC.presence_of_element_located((By.NAME, "Username"))).send_keys(required_settings['USERNAME'])
            wait.until(EC.presence_of_element_located((By.NAME, "Password"))).send_keys(password)
            wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Accedi' and contains(@class, 'x-btn-inner')]"))).click()
            attendi_scomparsa_overlay(driver, 60)

            logger.info("Navigating to Report -> Timesheet...")
            wait.until(EC.element_to_be_clickable((By.XPATH, "//*[normalize-space(text())='Report']"))).click()
            attendi_scomparsa_overlay(driver)

            timesheet_menu_xpath = "//span[contains(@id, 'generic_menu_button-') and contains(@id, '-btnEl')][.//span[text()='Timesheet']]"
            wait.until(EC.element_to_be_clickable((By.XPATH, timesheet_menu_xpath))).click()

            fornitore_arrow_xpath = "//div[starts-with(@id, 'generic_refresh_combo_box-') and contains(@id, '-trigger-picker')]"
            wait.until(EC.visibility_of_element_located((By.XPATH, fornitore_arrow_xpath)))
            attendi_scomparsa_overlay(driver)

            logger.info("Setting fixed parameters (Provider, Date)...")
            ActionChains(driver).move_to_element(wait.until(EC.element_to_be_clickable((By.XPATH, fornitore_arrow_xpath)))).click().perform()
            fornitore_option_xpath = f"//li[normalize-space(text())='{required_settings['FORNITORE_DA_SELEZIONARE']}']"
            fornitore_option = long_wait.until(EC.presence_of_element_located((By.XPATH, fornitore_option_xpath)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'nearest'}); arguments[0].click();", fornitore_option)
            attendi_scomparsa_overlay(driver)

            campo_data_da = wait.until(EC.visibility_of_element_located((By.NAME, "DataTimesheetDa")))
            campo_data_da.clear()
            campo_data_da.send_keys(required_settings['DATA_DA_INSERIRE'])

            # --- Loop through OdA data ---
            for oda in odas:
                numero_oda = oda.numero_oda
                posizione_oda = oda.posizione_oda or ""

                logger.info("-" * 20)
                logger.info(f"Processing OdA: {numero_oda}, Position: {posizione_oda}")

                try:
                    campo_numero_oda = wait.until(EC.presence_of_element_located((By.NAME, "NumeroOda")))
                    driver.execute_script("arguments[0].value = arguments[1];", campo_numero_oda, numero_oda)

                    campo_posizione_oda = wait.until(EC.presence_of_element_located((By.NAME, "PosizioneOda")))
                    driver.execute_script("arguments[0].value = '';", campo_posizione_oda) # Clear first
                    driver.execute_script("arguments[0].value = arguments[1];", campo_posizione_oda, posizione_oda)

                    pulsante_cerca_xpath = "//a[contains(@class, 'x-btn') and .//span[normalize-space(text())='Cerca']]"
                    wait.until(EC.element_to_be_clickable((By.XPATH, pulsante_cerca_xpath))).click()
                    attendi_scomparsa_overlay(driver, 90)

                    logger.info("  Downloading file...")
                    path_to_downloads_obj = Path(required_settings['DOWNLOAD_DIR'])
                    files_before_download = {f for f in path_to_downloads_obj.iterdir() if f.is_file() and f.suffix.lower() == '.xlsx'}

                    excel_button_xpath = "//div[contains(@class, 'x-tool')][.//div[contains(@class, 'x-tool-tool-el')]]"
                    wait.until(EC.element_to_be_clickable((By.XPATH, excel_button_xpath))).click()

                    downloaded_file_path = None
                    download_start_time = time.time()
                    while time.time() - download_start_time < 25:
                        current_files = {f for f in path_to_downloads_obj.iterdir() if f.is_file() and f.suffix.lower() == '.xlsx'}
                        new_files = current_files - files_before_download
                        if new_files:
                            downloaded_file_path = max(list(new_files), key=lambda f: f.stat().st_mtime)
                            logger.info(f"  File downloaded: {downloaded_file_path.name}")
                            break
                        time.sleep(0.5)

                    if downloaded_file_path and downloaded_file_path.exists():
                        nome_base_pos = f"-{posizione_oda}" if posizione_oda else ""
                        nuovo_nome_file = f"{numero_oda}{nome_base_pos}.xlsx"

                        path_destinazione_finale = Path(required_settings['DIR_SPOSTAMENTO_TS'])
                        path_destinazione_finale.mkdir(parents=True, exist_ok=True)
                        file_destinazione = path_destinazione_finale / nuovo_nome_file

                        counter = 1
                        while file_destinazione.exists():
                            timestamp = time.strftime("%Y%m%d-%H%M%S")
                            file_destinazione = path_destinazione_finale / f"{numero_oda}{nome_base_pos}-{timestamp}_{counter}.xlsx"
                            counter += 1

                        shutil.move(str(downloaded_file_path), str(file_destinazione))
                        logger.info(f"  File moved successfully to: {file_destinazione}")
                    else:
                        logger.error("  Download failed or file not found for this item.")

                except Exception as e_loop:
                    logger.error(f"  Error processing item {numero_oda}: {e_loop}", exc_info=True)
                    continue

            logger.info("Finished processing all OdA items.")

        except Exception as e:
            logger.critical(f"A general error occurred during Selenium operations: {e}", exc_info=True)
        finally:
            if driver:
                logger.info("Closing WebDriver.")
                driver.quit()
