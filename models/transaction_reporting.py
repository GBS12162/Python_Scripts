"""
Modelli di dati per Transaction Reporting - Rejecting Mensile
Gestisce ISIN, ordini raggruppati e controlli di qualità.
"""

from dataclasses import dataclass
from datetime import datetime, date
from typing import Optional, List, Dict, Any
from decimal import Decimal


@dataclass
class ISINGroup:
    """Modello per un gruppo ISIN con i suoi ordini."""
    
    isin: str  # Codice identificativo ISIN
    occorrenze: int  # Numero di ordini per questo ISIN
    orders: List['Order']  # Lista degli ordini per questo ISIN
    
    # Risultati dei controlli (ultime 4 colonne)
    controllo_1: Optional[str] = None  # Esito primo controllo
    controllo_2: Optional[str] = None  # Esito secondo controllo
    controllo_3: Optional[str] = None  # Esito terzo controllo
    controllo_4: Optional[str] = None  # Esito quarto controllo
    
    def __post_init__(self):
        """Validazione post-inizializzazione."""
        if not self.isin:
            raise ValueError("ISIN è obbligatorio")
        
        if self.occorrenze < 0:
            raise ValueError("Occorrenze non può essere negativo")
        
        if len(self.orders) != self.occorrenze:
            raise ValueError(f"Numero ordini ({len(self.orders)}) non corrisponde alle occorrenze ({self.occorrenze})")


@dataclass
class Order:
    """Modello per un singolo ordine."""
    
    # Identificativi
    order_id: str
    isin: str  # ISIN di riferimento
    
    # Dati dell'ordine
    quantity: Optional[Decimal] = None
    price: Optional[Decimal] = None
    order_type: Optional[str] = None  # BUY, SELL, etc.
    order_date: Optional[datetime] = None
    settlement_date: Optional[datetime] = None
    
    # Dati cliente/intermediario
    client_id: Optional[str] = None
    broker_id: Optional[str] = None
    account_id: Optional[str] = None
    
    # Metadati
    currency: str = "EUR"
    market: Optional[str] = None
    status: Optional[str] = None
    
    # Campi aggiuntivi dinamici per colonne non standard
    additional_fields: Dict[str, Any] = None
    
    def __post_init__(self):
        """Validazione post-inizializzazione."""
        if self.additional_fields is None:
            self.additional_fields = {}
        
        if not self.order_id:
            raise ValueError("ID ordine è obbligatorio")
        
        if not self.isin:
            raise ValueError("ISIN è obbligatorio")


@dataclass
class SharePointConfig:
    """Configurazione per l'accesso a SharePoint."""
    
    site_url: str  # URL del sito SharePoint
    library_name: str  # Nome della libreria documenti
    file_path: str  # Percorso del file nella libreria
    file_name_pattern: str  # Pattern del nome file (con wildcards se necessario)
    
    # Credenziali di accesso
    username: Optional[str] = None
    password: Optional[str] = None
    client_id: Optional[str] = None
    client_secret: Optional[str] = None
    
    # Opzioni
    download_latest: bool = True  # Scarica sempre l'ultima versione
    cache_locally: bool = True    # Mantieni cache locale
    
    def __post_init__(self):
        """Validazione configurazione."""
        if not self.site_url:
            raise ValueError("URL sito SharePoint è obbligatorio")
        
        if not self.file_name_pattern:
            raise ValueError("Pattern nome file è obbligatorio")


@dataclass
class ProcessingConfig:
    """Configurazione per l'elaborazione dei dati."""
    
    # Configurazione colonne Excel
    isin_column: str = "ISIN"
    occorrenze_column: str = "OCCORRENZE"
    
    # Colonne di controllo (ultime 4)
    controllo_columns: List[str] = None
    
    # Opzioni elaborazione
    skip_empty_isin: bool = True
    validate_occorrenze: bool = True
    strict_mode: bool = False  # Se True, fallisce su errori di validazione
    
    # Configurazione output
    output_format: str = "excel"  # excel, csv, json
    include_summary: bool = True
    include_details: bool = True
    
    def __post_init__(self):
        """Inizializzazione valori di default."""
        if self.controllo_columns is None:
            self.controllo_columns = ["Controllo_1", "Controllo_2", "Controllo_3", "Controllo_4"]


@dataclass
class ProcessingResult:
    """Risultato dell'elaborazione."""
    
    success: bool
    message: str
    
    # Dati elaborati
    isin_groups: List[ISINGroup] = None
    total_isin: int = 0
    total_orders: int = 0
    
    # Statistiche
    processing_stats: Dict[str, Any] = None
    validation_errors: List[str] = None
    warnings: List[str] = None
    
    # Output
    output_files: List[str] = None
    processing_time: Optional[float] = None
    
    def __post_init__(self):
        """Inizializzazione liste."""
        if self.isin_groups is None:
            self.isin_groups = []
        if self.processing_stats is None:
            self.processing_stats = {}
        if self.validation_errors is None:
            self.validation_errors = []
        if self.warnings is None:
            self.warnings = []
        if self.output_files is None:
            self.output_files = []


@dataclass
class QualityControlResult:
    """Risultato dei controlli di qualità."""
    
    isin: str
    total_orders: int
    
    # Risultati controlli
    controlli_passed: int = 0
    controlli_failed: int = 0
    controlli_details: Dict[str, str] = None
    
    # Errori specifici
    validation_errors: List[str] = None
    business_rule_violations: List[str] = None
    
    # Raccomandazioni
    recommendations: List[str] = None
    
    def __post_init__(self):
        """Inizializzazione."""
        if self.controlli_details is None:
            self.controlli_details = {}
        if self.validation_errors is None:
            self.validation_errors = []
        if self.business_rule_violations is None:
            self.business_rule_violations = []
        if self.recommendations is None:
            self.recommendations = []
    
    @property
    def success_rate(self) -> float:
        """Calcola il tasso di successo dei controlli."""
        total = self.controlli_passed + self.controlli_failed
        return (self.controlli_passed / total * 100) if total > 0 else 0.0


# Manteniamo i modelli precedenti per compatibilità
@dataclass
class Transaction:
    """Modello per una singola transazione (mantenuto per compatibilità)."""
    
    transaction_id: str
    account_number: str
    amount: Decimal
    currency: str
    transaction_date: datetime
    transaction_type: str
    status: str
    rejection_reason: Optional[str] = None
    rejection_code: Optional[str] = None
    processing_date: Optional[datetime] = None
    merchant_id: Optional[str] = None
    merchant_name: Optional[str] = None
    card_number_masked: Optional[str] = None
    
    def __post_init__(self):
        """Validazione post-inizializzazione."""
        if self.amount < 0:
            raise ValueError("L'importo non può essere negativo")
        
        if not self.transaction_id:
            raise ValueError("ID transazione è obbligatorio")
            
        if not self.account_number:
            raise ValueError("Numero conto è obbligatorio")


@dataclass 
class RejectionReport:
    """Modello per il report delle transazioni rifiutate (mantenuto per compatibilità)."""
    
    report_id: str
    generation_date: datetime
    period_start: datetime
    period_end: datetime
    total_transactions: int
    rejected_transactions: int
    rejection_rate: float
    total_amount: Decimal
    rejected_amount: Decimal
    rejection_by_type: Dict[str, int]
    rejection_by_reason: Dict[str, int]
    transactions: List[Transaction]
    
    def __post_init__(self):
        """Calcola statistiche automaticamente."""
        if self.total_transactions > 0:
            self.rejection_rate = (self.rejected_transactions / self.total_transactions) * 100
        else:
            self.rejection_rate = 0.0


@dataclass
class MonthlyReportConfig:
    """Configurazione per il report mensile (mantenuto per compatibilità)."""
    
    year: int
    month: int
    include_pending: bool = True
    include_failed: bool = True
    group_by_merchant: bool = False
    group_by_reason: bool = True
    export_format: str = "excel"  # excel, csv, pdf
    output_directory: str = "output"
    filename_template: str = "transaction_report_{year}_{month:02d}"
    
    def get_period_start(self) -> datetime:
        """Ottiene la data di inizio del periodo."""
        return datetime(self.year, self.month, 1)
    
    def get_period_end(self) -> datetime:
        """Ottiene la data di fine del periodo."""
        if self.month == 12:
            return datetime(self.year + 1, 1, 1)
        else:
            return datetime(self.year, self.month + 1, 1)
    
    def get_filename(self) -> str:
        """Genera il nome del file di output."""
        return self.filename_template.format(
            year=self.year,
            month=self.month
        )


@dataclass
class DataSource:
    """Configurazione per la fonte dati (mantenuto per compatibilità)."""
    
    source_type: str  # "database", "csv", "excel", "api", "sharepoint"
    connection_string: Optional[str] = None
    file_path: Optional[str] = None
    table_name: Optional[str] = None
    query: Optional[str] = None
    headers: Optional[Dict[str, str]] = None  # Mapping colonne
    
    def __post_init__(self):
        if self.headers is None:
            self.headers = {}