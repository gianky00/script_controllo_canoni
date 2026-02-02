# Prompt per Implementazione "Trova Numeri Automaticamente"

**Ruolo:** Sei un esperto sviluppatore Python specializzato in automazione desktop (Tkinter) e gestione file system.

**Obiettivo:** Implementare la logica backend e frontend per il pulsante **"Trova Numeri Automaticamente"** situato nella scheda "Stampa Canoni Mensili" dell'applicazione.

## Contesto
L'applicazione gestisce la stampa automatizzata di documenti Excel e Word per la contabilità mensile. L'utente seleziona un **Anno** e un **Mese** dalla GUI. Ci sono tre campi di input che devono essere popolati con un numero identificativo (progressivo) estratto dai nomi di file specifici salvati in rete.

## Requisiti Funzionali

### 1. Trigger
L'azione parte al click del pulsante `Trova Numeri Automaticamente` nella classe `FeesTab`.

### 2. Logica di Ricerca (Backend)
La logica deve essere incapsulata nella classe `MonthlyFeesProcessor` (file `src/logic/monthly_fees.py`).

**Parametri di Input:**
*   **Anno:** Selezionato dall'utente (es. "2025").
*   **Mese:** Selezionato dall'utente (es. "Gennaio").
*   **Lista TCL:** I tre referenti da cercare sono: "MESSINA", "NASELLI", "CALDARELLA".

**Percorso di Ricerca:**
I file si trovano in un percorso di rete strutturato così:
`\\192.168.11.251\Database_Tecnico_SMI\Contabilita' strumentale\{ANNO}\CONSUNTIVI\{ANNO}`
*(Nota: Recuperare il percorso base da `src.utils.constants.CANONI_CONSUNTIVI_BASE_DIR`)*

**Algoritmo di Matching:**
Per ogni referente (TCL), scansiona la cartella e cerca un file che soddisfi **TUTTE** le seguenti condizioni (case-insensitive):
1.  Il nome file contiene la stringa "CANONE".
2.  Il nome file contiene il **nome del mese** selezionato (es. "GENNAIO").
3.  Il nome file contiene il **cognome del referente** (es. "MESSINA").

**Estrazione del Numero:**
Se un file corrisponde, il nome del file inizia solitamente con un numero progressivo seguito da un trattino o uno spazio (es. `102-CANONE...` o `102 CANONE...`).
*   Devi estrarre questo numero iniziale (regex suggerita: `^(\d+)`).

### 3. Gestione GUI e Threading (Frontend)
Poiché l'operazione richiede accesso alla rete (lento), deve essere eseguita in un **thread separato** per non bloccare la GUI.

*   **File:** `src/gui/tabs/fees_tab.py`
*   **Metodo:** `find_numbers_and_populate` (avvia il thread).
*   **Thread Target:** `_find_numbers_thread`.
*   **Logica UI:**
    1.  Disabilitare i pulsanti per evitare click multipli.
    2.  Mostrare una progress bar o un messaggio di "Ricerca in corso...".
    3.  Passare un evento di cancellazione (`threading.Event`) per permettere all'utente di interrompere la ricerca.
    4.  Per ogni numero trovato, aggiornare la variabile Tkinter corrispondente (`canoni_messina_num`, `canoni_naselli_num`, `canoni_caldarella_num`) usando `self.master.after` per thread-safety.
    5.  Alla fine (o in caso di errore), riabilitare i pulsanti e nascondere la progress bar.

### 4. Logging
Utilizzare il widget di log dedicato (`self.logger` o callback passata) per informare l'utente:
*   "Ricerca avviata per {Mese} {Anno}..."
*   "Trovato file per {TCL}: {NomeFile} -> Numero: {N}" (Verde/Successo)
*   "Nessun file trovato per {TCL}" (Arancione/Warning)
*   "Errore di accesso al percorso di rete: ..." (Rosso/Errore)

## Esempio di Struttura Codice Attesa

**In `src/logic/monthly_fees.py`:**
```python
def find_consuntivo_for_tcl(self, year, month_name, tcl_name, cancel_event):
    # Costruisci path
    # Itera file
    # Check "CANONE" + mese + tcl
    # Regex numero
    # Ritorna (numero, full_path)
```

**In `src/gui/tabs/fees_tab.py`:**
```python
def _find_numbers_thread(self, cancel_event):
    target_tcls = {"MESSINA": self.app_config.canoni_messina_num, ...}
    for tcl, var in target_tcls.items():
        if cancel_event.is_set(): return
        num, path = self.processor.find_consuntivo_for_tcl(...)
        if num:
            self.master.after(0, var.set, num)
    # Cleanup UI
```