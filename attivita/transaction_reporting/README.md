# CON-412 Transaction Reporting - Guida Utente

## Descrizione
Sistema automatico per la validazione e il reporting mensile degli ordini CON-412 tramite controlli ESMA e database Oracle. Il programma elabora file Excel, effettua controlli su ogni ordine e genera un file validato pronto per l'invio.

## Requisiti
- **Windows 10/11**
- **Connessione Internet** (per validazione ESMA)
- **File Excel CON-412** da elaborare
- **Credenziali Oracle** (solo se richiesto per filtraggio database)

## Installazione
1. **Scarica la cartella** `TransactionReporting_Mensile` fornita dal reparto IT.
2. **Non serve installare Python**: l'eseguibile è già pronto all'uso.
3. **Verifica la presenza di questi elementi:**
   - `transaction_reporting_mensile.exe` (eseguibile per il reporting mensile)
   - Cartella `oracle_config` (con file TNS, se serve database)
   - Cartella `output` (verrà creata automaticamente)
   - Cartella `log` (verrà creata automaticamente)

## Avvio del Programma
1. **Doppio clic** su `transaction_reporting_mensile.exe`
2. **Segui le istruzioni a schermo:**
   - Inserisci il percorso completo del file Excel da validare (es: `C:\Downloads\CON-412_SETEMBRE.xlsx`)
   - Inserisci le credenziali Oracle se richiesto (username/password)

## Funzionalità Principali
- **Lettura automatica** del file Excel CON-412
- **Filtraggio ordini** tramite database Oracle (opzionale)
- **4 Controlli ESMA sequenziali**:
  1. ISIN censito (verifica presenza su database ESMA)
  2. Trading Venue (validazione MIC code vs mercato)
  3. Date Approval (verifica date di approvazione)
  4. Maturity Date (controllo date di scadenza)
- **Gestione errori**: log dettagliato e messaggi chiari in caso di problemi
- **File output**: viene generato un file Excel validato nella cartella `output` con la struttura originale e le colonne di validazione

## Output
- Il file validato viene salvato in `output/CON-412_[MESE]_Validated.xlsx`
- Viene mantenuta la struttura originale del file
- Vengono aggiunte colonne di validazione (X per controlli falliti)

## FAQ e Risoluzione Problemi
- **Errore database**: controlla le credenziali e la connessione di rete/VPN
- **Errore ESMA/API**: verifica la connessione internet
- **File non trovato**: assicurati che il percorso inserito sia corretto e che il file non sia aperto in Excel
- **Per assistenza**: contatta il supporto IT fornendo il file di log generato nella cartella `log`

## Note Finali
- Puoi rinominare l’eseguibile come preferisci
- I file di log e output vengono sovrascritti ad ogni esecuzione
- Per aggiornamenti o nuove versioni, richiedi sempre la cartella completa dal reparto IT

---
© 2025 Gruppo Banca Sella - Tutti i diritti riservati
- Mese corrente di elaborazione
- Directory di download e output
- URL SharePoint e API ESMA
- Parametri di validazione

## 📁 Struttura Output

I file generati dal sistema vengono salvati in:

```
attivita/transaction_reporting/
├── output/
│   └── con412_reports/           # File validati con X
│       ├── CON-412_[MESE]_Validated_[TIMESTAMP].xlsx
│       └── CON-412_[MESE]_Validated_[TIMESTAMP]_metadata.json
├── downloads/                    # File originali da SharePoint
│   └── CON-412_[MESE].xlsx      
└── logs/                        # File di log dettagliati
    └── con412_processing_[TIMESTAMP].log
```

**Struttura file di output:**
- Mantiene la struttura originale del file SharePoint
- Aggiunge colonna "PRIMO CONTROLLO" con X per ISIN validati ESMA
- Include riassunto delle validazioni effettuate

**Nota:** Le directory `output/`, `downloads/` e `logs/` sono automaticamente escluse dal controllo versione (.gitignore).

## 🔍 Monitoraggio

Durante l'esecuzione vengono mostrati:
- Stato di ogni fase
- Numero di ISIN elaborati
- Risultati della validazione ESMA
- Statistiche finali

## ⚠️ Requisiti

- Python 3.8+
- Connessione internet per API ESMA
- Accesso a SharePoint (se configurato)
- Librerie: pandas, openpyxl, requests

## 🆘 Risoluzione Problemi

In caso di errori, consultare i file di log nella directory `logs/` per dettagli completi.

---
**Autore:** GBS12162 (Samuele De Giosa)  
**Versione:** 1.0.0  
**Data:** 2025-09-16

## 📋 Descrizione

Questa attività fornisce un sistema completo per:
- Analizzare transazioni rifiutate da file CSV o Excel
- Generare report mensili dettagliati
- Esportare statistiche in vari formati (Excel, CSV, JSON)
- Visualizzare trend e pattern di rifiuto

## 🚀 Caratteristiche Principali

### 📊 Analisi Completa
- **Caricamento dati**: Supporto per file CSV ed Excel
- **Filtraggio intelligente**: Per periodo, status, tipo transazione
- **Statistiche avanzate**: Tassi di rifiuto, importi, trend temporali
- **Categorizzazione**: Per motivo di rifiuto, tipo transazione, merchant

### 📈 Reporting Avanzato
- **Report mensili**: Analisi completa per mese specifico
- **Confronti temporali**: Trend rispetto a periodi precedenti
- **Analisi pattern**: Orari, giorni, importi più critici
- **Statistiche dettagliate**: KPI e metriche operative

### 💾 Esportazione Flessibile
- **Excel**: Report formattati con grafici e tabelle
- **CSV**: Per elaborazioni successive
- **JSON**: Per integrazione con altri sistemi
- **Multi-formato**: Esportazione simultanea in tutti i formati

## 🛠️ Tecnologie Utilizzate

- **Python 3.8+**: Linguaggio principale
- **pandas**: Elaborazione e analisi dati
- **openpyxl/xlsxwriter**: Gestione file Excel
- **decimal**: Precisione numerica per importi
- **datetime**: Gestione date e periodi
- **logging**: Sistema di log completo

## 📁 Struttura del Progetto

```
transaction_reporting/
├── __init__.py                 # Inizializzazione modulo
├── main.py                     # Applicazione principale
├── launcher.py                 # Launcher per eseguibile
├── build_exe.py               # Script creazione eseguibile
├── README.md                  # Documentazione
├── data/                      # File di input (opzionale)
├── output/                    # Report generati
├── services/                  # Logica di business
│   ├── __init__.py
│   ├── transaction_data_service.py     # Gestione dati transazioni
│   ├── report_generation_service.py   # Generazione report
│   └── report_export_service.py       # Esportazione report
└── ui/                        # Interfaccia utente
    ├── __init__.py
    ├── console_ui.py          # Interfaccia console
    └── menu_manager.py        # Gestione menu
```

## 🚀 Come Utilizzare

### Esecuzione Diretta
```bash
cd attivita/transaction_reporting
python main.py
```

### Creazione Eseguibile
```bash
cd attivita/transaction_reporting
python build_exe.py
```

L'eseguibile sarà creato in `dist/Transaction_Reporting_Rejecting_Mensile.exe`

### Workflow Tipico

1. **Avvia l'applicazione**
   - Esegui `main.py` o l'eseguibile
   
2. **Configura il report**
   - Seleziona anno e mese
   - Imposta opzioni di filtraggio
   - Scegli formato di esportazione

3. **Configura fonte dati**
   - Seleziona file CSV o Excel
   - Mappa le colonne se necessario
   
4. **Genera e analizza**
   - L'applicazione elabora i dati
   - Genera statistiche e report
   - Esporta nei formati richiesti

## 📊 Formati Dati Supportati

### File di Input
Il sistema supporta file CSV ed Excel con le seguenti colonne:

#### Campi Obbligatori
- `transaction_id`: ID univoco della transazione
- `account_number`: Numero del conto
- `amount`: Importo della transazione (numerico)
- `currency`: Valuta (es. EUR, USD)
- `transaction_date`: Data della transazione
- `transaction_type`: Tipo di transazione
- `status`: Stato (APPROVED, REJECTED, FAILED, DENIED, ecc.)

#### Campi Opzionali
- `rejection_reason`: Motivo del rifiuto
- `rejection_code`: Codice di rifiuto
- `processing_date`: Data di elaborazione
- `merchant_id`: ID del merchant
- `merchant_name`: Nome del merchant
- `card_number_masked`: Numero carta mascherato

### Esempio CSV
```csv
transaction_id,account_number,amount,currency,transaction_date,transaction_type,status,rejection_reason
TX001,123456789,150.00,EUR,2025-09-15 14:30:00,PAYMENT,REJECTED,INSUFFICIENT_FUNDS
TX002,987654321,75.50,EUR,2025-09-15 15:45:00,PAYMENT,APPROVED,
TX003,555666777,200.00,EUR,2025-09-15 16:20:00,TRANSFER,FAILED,INVALID_ACCOUNT
```

## 📈 Tipi di Report Generati

### 1. Report Riepilogativo
- Statistiche generali del periodo
- Tassi di rifiuto complessivi
- Importi totali e rifiutati
- Confronti con periodi precedenti

### 2. Analisi per Motivo di Rifiuto
- Distribuzione dei motivi di rifiuto
- Top 10 motivi più frequenti
- Trend temporali per motivo
- Impatto economico per categoria

### 3. Analisi Temporale
- Distribuzione oraria dei rifiuti
- Pattern settimanali
- Picchi di attività
- Correlazioni temporali

### 4. Analisi per Importo
- Distribuzione per fasce di importo
- Soglie critiche di rifiuto
- Importi medi per categoria
- Analisi outlier

## ⚙️ Configurazioni Avanzate

### Filtri Disponibili
- **Periodo**: Anno/mese specifico
- **Status**: Include/escludi pending e failed
- **Tipo transazione**: Filtraggio per categoria
- **Importo**: Range di importi
- **Merchant**: Filtraggio per merchant

### Opzioni di Esportazione
- **Excel**: Report formattato con fogli multipli
- **CSV**: File separati per ogni categoria di dati
- **JSON**: Struttura dati completa per API
- **All**: Esportazione in tutti i formati

## 🔧 Configurazione Ambiente

### Dipendenze Richieste
```bash
pip install pandas openpyxl xlsxwriter chardet tqdm
```

### Variabili d'Ambiente (Opzionali)
- `TR_OUTPUT_DIR`: Directory output personalizzata
- `TR_LOG_LEVEL`: Livello di logging (INFO, DEBUG, WARNING)
- `TR_DEFAULT_CURRENCY`: Valuta di default

## 🐛 Risoluzione Problemi

### Errori Comuni

1. **"Nessuna transazione caricata"**
   - Verifica formato del file
   - Controlla encoding (deve essere UTF-8)
   - Assicurati che le colonne siano mappate correttamente

2. **"Errore nel parsing delle date"**
   - Usa formato ISO (YYYY-MM-DD) o standard locale
   - Verifica che non ci siano celle vuote nelle date

3. **"Errore negli importi"**
   - Gli importi devono essere numerici
   - Usa punto (.) come separatore decimale
   - Rimuovi simboli di valuta

### Debug
Esegui con flag di debug:
```bash
python main.py --debug
```

## 📝 Log e Monitoring

I log sono salvati in `logs/transaction_reporting_YYYYMMDD_HHMMSS.log` e includono:
- Operazioni di caricamento dati
- Statistiche di elaborazione
- Errori e warning
- Performance metrics

## 🔒 Sicurezza e Privacy

- I dati sono elaborati solo localmente
- Nessuna trasmissione di dati sensibili
- File di log non contengono dati delle transazioni
- Supporto per dati mascherati

## 📊 Metriche e KPI

Il sistema calcola automaticamente:
- **Rejection Rate**: Percentuale di transazioni rifiutate
- **Volume Impact**: Impatto sul volume delle transazioni
- **Financial Impact**: Impatto economico dei rifiuti
- **Temporal Patterns**: Pattern temporali di rifiuto
- **Category Analysis**: Analisi per categoria di transazione

## 🚀 Roadmap Future

- [ ] Supporto per database diretti
- [ ] API REST per integrazione
- [ ] Dashboard web interattiva
- [ ] Machine learning per predizione rifiuti
- [ ] Alerting automatico per soglie critiche
- [ ] Esportazione in PDF con grafici

## 👥 Contributi

Sviluppato da: **GBS12162 (Samuele De Giosa)**  
Data: **2025-09-16**  
Versione: **1.0.0**

Per segnalazioni o miglioramenti, contatta il team di sviluppo.

---

*Questo sistema fa parte del progetto Python Scripts centralizzato per la gestione di multiple attività aziendali.*