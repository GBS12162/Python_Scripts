"""
Modelli per le applicazioni Python Scripts.
Contiene le classi dei dati e le strutture per componenti Oracle e transaction reporting.
"""

from .component import Component, LookupResult
from .config import Config
from .transaction_reporting import (
    # Nuovi modelli per ISIN/Ordini
    ISINGroup,
    Order, 
    SharePointConfig,
    ProcessingConfig,
    QualityControlResult,
    # Modelli esistenti per compatibilità
    Transaction, 
    RejectionReport, 
    MonthlyReportConfig, 
    ProcessingResult, 
    DataSource
)

__all__ = [
    'Component', 
    'LookupResult', 
    'Config',
    # Nuovi modelli
    'ISINGroup',
    'Order',
    'SharePointConfig', 
    'ProcessingConfig',
    'QualityControlResult',
    # Modelli esistenti
    'Transaction',
    'RejectionReport', 
    'MonthlyReportConfig', 
    'ProcessingResult', 
    'DataSource'
]
