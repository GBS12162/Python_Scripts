# CON-412 TRANSACTION REPORTING - SISTEMA AUTOMATICO
## Versione 1.0 - Eseguibile Standalone

### 📋 DESCRIZIONE
Sistema automatico per la validazione ESMA di ordini CON-412 Transaction Reporting.
Elabora file Excel con validazione ISIN, controlli Trading Venue, Date Approval e Maturity Date.

### 🚀 FUNZIONALITÀ
- ✅ **Lettura automatica** file Excel con struttura CON-412
- ✅ **Validazione individuale** di ogni ordine con il proprio mercato
- ✅ **4 Controlli ESMA sequenziali**:
  1. ISIN censito (verifica presenza su database ESMA)
  2. Trading Venue (validazione MIC code vs mercato)
  3. Date Approval (verifica date di approvazione)
  4. Maturity Date (controllo date di scadenza)
- ✅ **Connessione database Oracle** per filtraggio RF/EE (opzionale)
- ✅ **File output** preserva struttura originale + colonne X per controlli

### 📁 FILE E CARTELLE
```
CON-412_System/
├── main_con412.exe          # Eseguibile principale
├── oracle_config/           # Configurazione database Oracle
│   └── tnsnames.ora         # File configurazione TNS
├── output/                  # File Excel risultato elaborazione
├── work/                    # Cartella temporanea di lavoro
└── log/                     # File di log elaborazione
```

### 🔧 INSTALLAZIONE
1. **Copiare** la cartella completa `CON-412_System` sul computer di destinazione
2. **Non** è necessario installare Python o altre dipendenze
3. L'eseguibile è **standalone** e portabile

### 📋 UTILIZZO
1. **Avviare** `main_con412.exe` facendo doppio clic
2. **Inserire** il path completo al file Excel CON-412:
   - File locale: `C:\Downloads\CON-412_AGOSTO.xlsx`
   - NAS/Server: `\\server\share\reports\CON-412_AGOSTO.xlsx`
3. **Credenziali database** (se richieste):
   - Username: `CONSULTA_IT_RUN_INVESTIMENTI`
   - Password: `[password corrente]`
4. Il sistema elabora **automaticamente** tutti i controlli
5. **File risultato** salvato in cartella `output/`

### 📊 STRUTTURA FILE EXCEL
Il sistema legge file Excel con questa struttura:
```
Riga ISIN: [ISIN] [NUMERO_OCCORRENZE]
Riga Ordine 1: [dati ordine]
Riga Ordine 2: [dati ordine]
...
Riga ISIN: [ISIN] [NUMERO_OCCORRENZE]
Riga Ordine 1: [dati ordine]
```

**IMPORTANTE**: Ogni ordine viene validato **individualmente** con il proprio mercato specifico.

### 🔍 CONTROLLI ESMA
1. **CONTROLLO 1 - ISIN CENSITO**
   - Verifica presenza ISIN su database ESMA
   - Se ISIN non trovato → aggiunge X nella colonna

2. **CONTROLLO 2 - TRADING VENUE**
   - Confronta MIC code ESMA con mercato dell'ordine
   - XOFF (OTC) accetta qualsiasi trading venue
   - Se MIC non trovato → aggiunge X nella colonna

3. **CONTROLLO 3 - DATE APPROVAL**
   - Verifica validità date di approvazione
   - Se date non valide → aggiunge X nella colonna

4. **CONTROLLO 4 - MATURITY DATE**
   - Controlla date di scadenza strumento
   - Se date non valide → aggiunge X nella colonna

### 💾 DATABASE ORACLE
- **Servizio**: PPORAFIN (pre-produzione)
- **Scopo**: Filtraggio ordini RF/EE prima dei controlli ESMA
- **Comportamento**: Se database non disponibile, elabora tutti gli ordini
- **Configurazione**: File `oracle_config/tnsnames.ora`

### 📄 FILE OUTPUT
Il sistema genera un file Excel con:
- **Struttura originale** completamente preservata
- **Colonne X aggiunte** per ogni controllo ESMA fallito
- **Nome file**: `[Nome_Originale]_Validated.xlsx`
- **Posizione**: Cartella `output/`

### 🐛 RISOLUZIONE PROBLEMI

#### Database non si connette
- **Causa**: Credenziali errate o servizio non disponibile
- **Soluzione**: Il sistema continua automaticamente senza database
- **Nota**: Alcuni ordini RF/EE potrebbero essere elaborati

#### File Excel non trovato
- **Causa**: Path errato o file non accessibile
- **Soluzione**: Verificare path completo e accesso al file

#### Errori di elaborazione
- **Log dettagliato**: Cartella `log/`
- **Controllo**: File `main_con412.log` per dettagli errori

### 📈 STATISTICHE OUTPUT
Al termine dell'elaborazione, il sistema mostra:
- **ISIN totali** elaborati
- **Ordini totali** processati
- **Risultati per controllo**:
  - Controllo 1: X/Y ordini passati
  - Controllo 2: X/Y ordini passati
  - Controllo 3: X/Y ordini passati
  - Controllo 4: X/Y ordini passati

### 📞 SUPPORTO TECNICO
- **Sistema**: CON-412 Transaction Reporting v1.0
- **Compatibilità**: Windows 10/11
- **Dipendenze**: Nessuna (standalone)
- **Memoria**: Minimo 4GB RAM consigliati
- **Spazio disco**: 100MB per installazione + spazio file temporanei

### 🔒 SICUREZZA
- Le credenziali database **non** vengono salvate
- I file temporanei vengono rimossi automaticamente
- Connessioni database utilizzano crittografia TLS

---
**Data creazione**: Ottobre 2025  
**Versione sistema**: 1.0  
**Ultimo aggiornamento**: 02/10/2025