# GUIDA RAPIDA - CON-412 TRANSACTION REPORTING

## 🚀 AVVIO VELOCE (3 PASSI)

### 1️⃣ PREPARAZIONE
- Assicurarsi di avere il **file Excel CON-412** da elaborare
- Annotare il **path completo** del file (es: `C:\Downloads\CON-412_SETTEMBRE.xlsx`)

### 2️⃣ AVVIO
- Fare **doppio clic** su `Avvia_CON412.bat` 
- Oppure fare **doppio clic** su `dist/main_con412.exe`

### 3️⃣ CONFIGURAZIONE
Quando richiesto dal sistema:

**Path del file Excel:**
```
Inserisci il path completo al file Excel:
> C:\Downloads\CON-412_SETTEMBRE.xlsx
```

**Credenziali database (se richieste):**
```
Username: CONSULTA_IT_RUN_INVESTIMENTI
Password: [password corrente PPORAFIN]
```

## 📊 COSA FA IL SISTEMA

### ✅ AUTOMATICO
1. **Legge** il file Excel CON-412
2. **Interpreta** la struttura ISIN → Ordini
3. **Tenta connessione** database Oracle per filtrare RF/EE
4. **Esegue 4 controlli ESMA** su ogni ordine:
   - 🔍 ISIN censito su ESMA
   - 🏛️ Trading Venue (MIC vs Mercato)
   - 📅 Date Approval valide
   - ⏰ Maturity Date valide
5. **Genera file finale** con X per controlli falliti

### 📁 RISULTATO
File salvato in: `output/[Nome_File]_Validated.xlsx`

## 🔧 CASI COMUNI

### ✅ Database non disponibile
```
❌ Connessione database fallita
⚠️ Database non disponibile - continuando con controlli ESMA
```
**→ NORMALE**: Il sistema continua automaticamente

### ✅ File non trovato
```
❌ File non trovato: C:\Path\Errato.xlsx
```
**→ SOLUZIONE**: Verificare il path del file

### ✅ Credenziali errate
```
❌ ORA-01017: invalid username/password
```
**→ SOLUZIONE**: Il sistema continua senza database

## 📋 INTERPRETI RISULTATI

### Colonne X nel file finale:
- **Controllo 1**: ISIN non censito su ESMA
- **Controllo 2**: MIC code non trovato/non valido  
- **Controllo 3**: Date approval non valide
- **Controllo 4**: Maturity date non valide

### Statistiche finali:
```
📊 CONTROLLO 1 - ISIN CENSITI: 24/25 ordini
📊 CONTROLLO 2 - MIC CODE PRESENTI: 17/25 ordini
📊 CONTROLLO 3 - DATE APPROVAL OK: 25/25 ordini
📊 CONTROLLO 4 - MATURITY DATE OK: 19/25 ordini
```

## ⚠️ IMPORTANTE

### 🎯 VALIDAZIONE INDIVIDUALE
**Ogni ordine viene controllato singolarmente** con il proprio mercato specifico.

### 📄 STRUTTURA PRESERVATA
Il file originale **rimane identico**, vengono aggiunte solo le colonne X.

### 🔒 SICUREZZA
Le password **non vengono salvate** e le connessioni sono crittografate.

---
**Supporto**: Verificare file `log/main_con412.log` per dettagli errori