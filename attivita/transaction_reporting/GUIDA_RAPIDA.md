# Guida Rapida CON-412 Transaction Reporting

## 🔧 Configurazione SharePoint

### Opzione 1: Centrico IT Run (Raccomandato)
- **Sito:** https://gruppobancasella.sharepoint.com/teams/CentricoITRun
- **Directory:** Documenti Condivisi
- **Nome file automatico:** `CON-412_[MESE].xlsx`
  - Esempio: `CON-412_SEPTEMBER.xlsx`
  - Esempio: `CON-412_AUGUST.xlsx`
  - Esempio: `CON-412_DECEMBER.xlsx`

**Il sistema usa automaticamente il formato corretto!** Non devi inserire nulla, basta selezionare il mese.

### Opzione 2: SharePoint Personalizzato
- **URL sito:** Inserire URL completo (es: https://company.sharepoint.com/sites/finance)
- **Path file:** Path completo dal sito (es: /Shared Documents/Reports/file.xlsx)

## 📄 Nome File Automatico

### Per Centrico IT Run:
```
📄 FILE DA SCARICARE
Nome file: CON-412_SEPTEMBER.xlsx
Formato fisso: CON-412_[MESE].xlsx

Usare nome file diverso? (n/S per personalizzare): n
```

**Comportamento predefinito:**
- Il sistema usa automaticamente `CON-412_[MESE].xlsx`
- Premi **INVIO** (o 'n') per usare il nome standard
- Premi **'s'** solo se il file ha un nome diverso (es: CON-412_AUGUST_V2.xlsx)

### Casi speciali:
Se il file ha un nome diverso dal formato standard, puoi personalizzarlo:
- `CON-412_SEPTEMBER_V2.xlsx` → Per versioni specifiche
- `CON-412_SEP_Final.xlsx` → Per file con abbreviazioni
- `CON-412_SEPTEMBER_Corrected.xlsx` → Per file corretti

## 📁 File di Output

I report vengono salvati in:
```
attivita/transaction_reporting/output/con412_reports/
├── CON-412_SEPTEMBER_Validated_20250916_154530.xlsx
└── CON-412_SEPTEMBER_Validated_20250916_154530_metadata.json
```

### 📊 Struttura del file validato:
Il file mantiene la **stessa struttura** del file originale scaricato da SharePoint, con l'aggiunta di:

| ISIN | OCCORRENCES | PRIMO CONTROLLO | VALIDAZIONE ESMA | NOTE |
|------|-------------|-----------------|------------------|------|
| IT0001234567 | 150 | **X** | CENSURATO | ISIN italiano validato |
| US1234567890 | 45 | | NON CENSURATO | ISIN estero |

**La X appare solo se l'ISIN supera la validazione ESMA!**

## 🚀 Avvio Rapido

1. **Windows:** Doppio click su `run_con412.bat`
2. **Command Line:** `py main_con412.py`

## ❓ FAQ

**Q: Il nome del file è sempre CON-412_[MESE].xlsx?**
A: Sì! Il sistema usa automaticamente questo formato. Non devi inserire nulla.

**Q: Come faccio se il file ha un nome diverso?**
A: Quando ti chiede "Usare nome file diverso?", premi 's' e inserisci il nome completo.

**Q: Devo includere l'estensione .xlsx?**
A: Sì, se personalizzi il nome, includi sempre l'estensione completa.

**Q: Dove trovo i file di output?**
A: Nella cartella `output/con412_reports/` dell'attività

**Q: Il sistema funziona se il file non esiste?**
A: Ti avviserà se il file CON-412_[MESE].xlsx non viene trovato su SharePoint.

---
*Documento aggiornato: 2025-09-16*