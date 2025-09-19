# Gestione Dipendenze - Python Scripts Repository

## 🏗️ Architettura

Questo repository usa una **strategia stratificata** per la gestione delle dipendenze:

```
Python Scripts/
├── .venv/                           # Environment virtuale UNICO
├── requirements.txt                 # Dipendenze COMUNI (pandas, requests, etc.)
├── setup_dependencies.py           # Script automatico installazione
├── attivita/
│   ├── controlli_di_linea/
│   │   └── requirements.txt         # Dipendenze specifiche (py7zr, etc.)
│   └── transaction_reporting/
│       └── requirements.txt         # Dipendenze specifiche (oracledb, etc.)
```

## 🚀 Installazione Rapida

### Opzione 1: Script Automatico (Consigliato)
```bash
# Attiva environment virtuale
.venv\Scripts\Activate.ps1

# Installa tutto automaticamente
python setup_dependencies.py
```

### Opzione 2: Installazione Manuale
```bash
# Attiva environment virtuale
.venv\Scripts\Activate.ps1

# 1. Dipendenze comuni
pip install -r requirements.txt

# 2. Dipendenze specifiche (esempio transaction_reporting)
pip install -r attivita/transaction_reporting/requirements.txt
```

### Opzione 3: Installazione Combinata
```bash
# Tutto in un comando
pip install -r requirements.txt -r attivita/transaction_reporting/requirements.txt
```

## 📦 Dipendenze per Progetto

### Comuni (Root)
- **pandas**: Elaborazione dati
- **openpyxl**: Excel I/O
- **requests**: API calls
- **keyring**: Gestione credenziali
- **tqdm**: Progress bars
- **pytest**: Testing

### Transaction Reporting
- **oracledb**: Database Oracle TNS
- **mypy**: Type checking

### Controlli di Linea  
- **py7zr**: Compressione 7z
- **mypy**: Type checking

## 🔧 Gestione Environment

### Creazione Environment
```bash
# Solo la prima volta
python -m venv .venv
```

### Attivazione
```bash
# Windows PowerShell
.venv\Scripts\Activate.ps1

# Windows Command Prompt
.venv\Scripts\activate.bat

# Linux/Mac
source .venv/bin/activate
```

### Verifica
```bash
python -c "import sys; print(sys.prefix)"
pip list
```

## 📋 Best Practices

1. **Un solo environment**: Condiviso per tutto il repository
2. **Dipendenze comuni**: Nel requirements.txt root
3. **Dipendenze specifiche**: Nei requirements.txt dei progetti
4. **Script automatico**: Usa `setup_dependencies.py` per installazioni pulite
5. **Documentazione**: Aggiorna questo README quando aggiungi progetti

## 🆘 Troubleshooting

### Problemi Environment
```bash
# Ricreare environment da zero
Remove-Item .venv -Recurse -Force
python -m venv .venv
.venv\Scripts\Activate.ps1
python setup_dependencies.py
```

### Problemi Dipendenze
```bash
# Aggiornare pip
python -m pip install --upgrade pip

# Installazione forzata
pip install --force-reinstall -r requirements.txt
```

### Verifica Installazione
```bash
# Test imports principali
python -c "import pandas, openpyxl, requests; print('✅ Dipendenze comuni OK')"
python -c "import oracledb; print('✅ Oracle OK')"  # Solo per transaction_reporting
```