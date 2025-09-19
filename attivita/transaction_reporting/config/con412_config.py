"""
Configurazione specifica per Transaction Reporting - Rejecting Mensile
Gruppo Banca Sella - CON-412
"""

import calendar
from datetime import datetime
from models.transaction_reporting import SharePointConfig, ProcessingConfig


class CON412Config:
    """Configurazione specifica per CON-412."""
    
    # URL e configurazione SharePoint
    SHAREPOINT_SITE_URL = "https://gruppobancasella.sharepoint.com/teams/CentricoITRun"
    SHAREPOINT_LIBRARY = "Documenti condivisi"
    SHAREPOINT_BASE_PATH = "/teams/CentricoITRun/Documenti condivisi/Attività in carico IT Run Italia/IT Run Investimenti_ST0238/CNT IT NEGOZIAZIONE, ACCESSO AI MERCATI E AMMINISTRAZIONE/Controllo Rejected Transaction Reporting_CON412/Rejected Transaction Reporting Mensile/Report Rejected - Transaction Reporting"
    
    # Pattern nome file basato sul mese
    @staticmethod
    def get_file_pattern(month: int = None) -> str:
        """
        Genera il pattern del nome file per il mese specificato.
        
        Args:
            month: Mese (1-12), se None usa il mese corrente
            
        Returns:
            Pattern del nome file (es. "CON-412_SEPTEMBER*")
        """
        if month is None:
            month = datetime.now().month
        
        month_name = calendar.month_name[month].upper()
        return f"CON-412_{month_name}*"
    
    @staticmethod
    def get_sharepoint_config(month: int = None) -> SharePointConfig:
        """
        Crea la configurazione SharePoint per il mese specificato.
        
        Args:
            month: Mese (1-12), se None usa il mese corrente
            
        Returns:
            Configurazione SharePoint
        """
        return SharePointConfig(
            site_url=CON412Config.SHAREPOINT_SITE_URL,
            library_name=CON412Config.SHAREPOINT_LIBRARY,
            file_path=CON412Config.SHAREPOINT_BASE_PATH,
            file_name_pattern=CON412Config.get_file_pattern(month),
            download_latest=True,
            cache_locally=True
        )
    
    @staticmethod
    def get_processing_config() -> ProcessingConfig:
        """
        Crea la configurazione di elaborazione per CON-412.
        
        Returns:
            Configurazione elaborazione
        """
        return ProcessingConfig(
            isin_column="ISIN",  # Campo 1: Codice identificativo dello strumento
            occorrenze_column="OCCORRENZE",
            controllo_columns=[
                "ISIN_NON_CENSITO",           # Controllo 1: ISIN non censito
                "MIC_CODE_NON_PRESENTE",      # Controllo 2: MIC code non presente  
                "ISIN_NON_VALIDO_AMMISSIONE", # Controllo 3: ISIN non valido per data ammissione
                "ISIN_NON_VALIDO_CESSAZIONE"  # Controllo 4: ISIN non valido per data cessazione
            ],
            skip_empty_isin=True,
            validate_occorrenze=True,
            strict_mode=False,
            output_format="excel",
            include_summary=True,
            include_details=True
        )


class ControlliConfig:
    """Configurazione per i controlli di validazione."""
    
    # Descrizioni dei controlli
    CONTROLLI_DESCRIPTIONS = {
        "ISIN_NON_CENSITO": {
            "numero": 1,
            "descrizione": "CASISTICA isin non censito",
            "campo_riferimento": "campo 1 Codice identificativo dello strumento",
            "azione": "Verificare se ISIN è presente nell'API di riferimento",
            "risultato_positivo": "X",  # Mette X se ISIN NON è trovato
            "risultato_negativo": ""   # Lascia vuoto se ISIN è trovato
        },
        "MIC_CODE_NON_PRESENTE": {
            "numero": 2,
            "descrizione": "CASISTICA mic code non presente per quel mercato",
            "campo_riferimento": "campo 6 Sede di negoziazione",
            "azione": "Verificare presenza MIC code per il mercato specificato",
            "risultato_positivo": "X",
            "risultato_negativo": ""
        },
        "ISIN_NON_VALIDO_AMMISSIONE": {
            "numero": 3,
            "descrizione": "CASISTICA isin non valido per data e ora rispetto all'esecuzione",
            "campo_riferimento": "campo 11 data di ammissione alla negoziazione o data di prima negoziazione",
            "azione": "Verificare validità ISIN per data ammissione vs esecuzione",
            "risultato_positivo": "X",
            "risultato_negativo": ""
        },
        "ISIN_NON_VALIDO_CESSAZIONE": {
            "numero": 4,
            "descrizione": "CASISTICA isin non valido per data e ora rispetto all'esecuzione",
            "campo_riferimento": "campo 12 data di cessazione",
            "azione": "Verificare validità ISIN per data cessazione vs esecuzione",
            "risultato_positivo": "X",
            "risultato_negativo": ""
        }
    }
    
    # URL API per controlli (da configurare)
    API_ENDPOINTS = {
        "isin_lookup": None,  # Da configurare con l'URL dell'API ISIN
        "mic_code_lookup": None,
        "market_data": None
    }
    
    @staticmethod
    def set_isin_api_url(url: str):
        """Imposta l'URL dell'API per il controllo ISIN."""
        ControlliConfig.API_ENDPOINTS["isin_lookup"] = url
    
    @staticmethod
    def get_controllo_info(controllo_name: str) -> dict:
        """Ottiene le informazioni su un controllo specifico."""
        return ControlliConfig.CONTROLLI_DESCRIPTIONS.get(controllo_name, {})


class FieldMapping:
    """Mapping dei campi del file Excel."""
    
    # Campi standard attesi nel file Excel
    STANDARD_FIELDS = {
        "campo_1": "ISIN",  # Codice identificativo dello strumento
        "campo_6": "SEDE_NEGOZIAZIONE",  # Sede di negoziazione
        "campo_11": "DATA_AMMISSIONE",   # Data ammissione alla negoziazione
        "campo_12": "DATA_CESSAZIONE",   # Data di cessazione
        "occorrenze": "OCCORRENZE"       # Numero di ordini per ISIN
    }
    
    # Campi di controllo (ultime 4 colonne)
    CONTROL_FIELDS = {
        "controllo_1": "ISIN_NON_CENSITO",
        "controllo_2": "MIC_CODE_NON_PRESENTE", 
        "controllo_3": "ISIN_NON_VALIDO_AMMISSIONE",
        "controllo_4": "ISIN_NON_VALIDO_CESSAZIONE"
    }
    
    @staticmethod
    def get_all_expected_columns() -> list:
        """Ottiene tutti i nomi delle colonne attese."""
        return list(FieldMapping.STANDARD_FIELDS.values()) + list(FieldMapping.CONTROL_FIELDS.values())