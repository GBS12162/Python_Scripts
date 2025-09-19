"""
Servizio per la gestione dei dati delle transazioni.
Gestisce caricamento, elaborazione e validazione dei dati transazionali.
"""

import pandas as pd
import logging
from datetime import datetime
from decimal import Decimal
from typing import List, Dict, Optional, Any
from pathlib import Path

from models.transaction_reporting import Transaction, DataSource
from utils.file_utils import detect_encoding


class TransactionDataService:
    """Servizio per la gestione dei dati delle transazioni."""
    
    def __init__(self):
        """Inizializza il servizio."""
        self.logger = logging.getLogger(__name__)
        self._transactions: List[Transaction] = []
        self._raw_data: Optional[pd.DataFrame] = None
    
    def load_from_csv(self, file_path: str, data_source: DataSource) -> List[Transaction]:
        """
        Carica le transazioni da un file CSV.
        
        Args:
            file_path: Percorso del file CSV
            data_source: Configurazione della fonte dati
            
        Returns:
            Lista delle transazioni caricate
            
        Raises:
            FileNotFoundError: Se il file non esiste
            ValueError: Se i dati non sono validi
        """
        try:
            file_path_obj = Path(file_path)
            if not file_path_obj.exists():
                raise FileNotFoundError(f"File non trovato: {file_path}")
            
            # Rileva encoding
            encoding = detect_encoding(file_path)
            self.logger.info(f"Caricamento CSV da: {file_path} (encoding: {encoding})")
            
            # Carica DataFrame
            df = pd.read_csv(file_path, encoding=encoding)
            self.logger.info(f"Caricate {len(df)} righe dal CSV")
            
            # Applica mapping delle colonne se specificato
            if data_source.headers:
                df = df.rename(columns=data_source.headers)
            
            # Converte in transazioni
            transactions = self._dataframe_to_transactions(df)
            self._transactions = transactions
            self._raw_data = df
            
            self.logger.info(f"Processate {len(transactions)} transazioni valide")
            return transactions
            
        except Exception as e:
            self.logger.error(f"Errore nel caricamento CSV: {e}")
            raise
    
    def load_from_excel(self, file_path: str, data_source: DataSource, sheet_name: str = None) -> List[Transaction]:
        """
        Carica le transazioni da un file Excel.
        
        Args:
            file_path: Percorso del file Excel
            data_source: Configurazione della fonte dati
            sheet_name: Nome del foglio (opzionale)
            
        Returns:
            Lista delle transazioni caricate
        """
        try:
            file_path_obj = Path(file_path)
            if not file_path_obj.exists():
                raise FileNotFoundError(f"File non trovato: {file_path}")
            
            self.logger.info(f"Caricamento Excel da: {file_path}")
            
            # Carica DataFrame
            if sheet_name:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(file_path)
            
            self.logger.info(f"Caricate {len(df)} righe dall'Excel")
            
            # Applica mapping delle colonne se specificato
            if data_source.headers:
                df = df.rename(columns=data_source.headers)
            
            # Converte in transazioni
            transactions = self._dataframe_to_transactions(df)
            self._transactions = transactions
            self._raw_data = df
            
            self.logger.info(f"Processate {len(transactions)} transazioni valide")
            return transactions
            
        except Exception as e:
            self.logger.error(f"Errore nel caricamento Excel: {e}")
            raise
    
    def filter_by_period(self, start_date: datetime, end_date: datetime) -> List[Transaction]:
        """
        Filtra le transazioni per periodo.
        
        Args:
            start_date: Data di inizio
            end_date: Data di fine
            
        Returns:
            Lista delle transazioni filtrate
        """
        filtered = [
            t for t in self._transactions
            if start_date <= t.transaction_date < end_date
        ]
        
        self.logger.info(f"Filtrate {len(filtered)} transazioni per periodo {start_date} - {end_date}")
        return filtered
    
    def filter_rejected_only(self) -> List[Transaction]:
        """
        Filtra solo le transazioni rifiutate.
        
        Returns:
            Lista delle transazioni rifiutate
        """
        rejected_statuses = ["REJECTED", "FAILED", "DENIED", "DECLINED"]
        rejected = [
            t for t in self._transactions
            if t.status.upper() in rejected_statuses
        ]
        
        self.logger.info(f"Trovate {len(rejected)} transazioni rifiutate")
        return rejected
    
    def get_statistics(self) -> Dict[str, Any]:
        """
        Calcola statistiche sulle transazioni caricate.
        
        Returns:
            Dizionario con le statistiche
        """
        if not self._transactions:
            return {}
        
        total_count = len(self._transactions)
        rejected = self.filter_rejected_only()
        rejected_count = len(rejected)
        
        total_amount = sum(t.amount for t in self._transactions)
        rejected_amount = sum(t.amount for t in rejected)
        
        # Statistiche per status
        status_counts = {}
        for t in self._transactions:
            status_counts[t.status] = status_counts.get(t.status, 0) + 1
        
        # Statistiche per motivo di rifiuto
        rejection_reasons = {}
        for t in rejected:
            if t.rejection_reason:
                rejection_reasons[t.rejection_reason] = rejection_reasons.get(t.rejection_reason, 0) + 1
        
        return {
            "total_transactions": total_count,
            "rejected_transactions": rejected_count,
            "rejection_rate": (rejected_count / total_count * 100) if total_count > 0 else 0,
            "total_amount": total_amount,
            "rejected_amount": rejected_amount,
            "status_distribution": status_counts,
            "rejection_reasons": rejection_reasons,
            "currencies": list(set(t.currency for t in self._transactions)),
            "date_range": {
                "start": min(t.transaction_date for t in self._transactions) if self._transactions else None,
                "end": max(t.transaction_date for t in self._transactions) if self._transactions else None
            }
        }
    
    def _dataframe_to_transactions(self, df: pd.DataFrame) -> List[Transaction]:
        """
        Converte un DataFrame in lista di transazioni.
        
        Args:
            df: DataFrame con i dati delle transazioni
            
        Returns:
            Lista di oggetti Transaction
        """
        transactions = []
        
        for index, row in df.iterrows():
            try:
                # Parsing dei dati
                transaction = Transaction(
                    transaction_id=str(row.get('transaction_id', f"TX_{index}")),
                    account_number=str(row.get('account_number', '')),
                    amount=Decimal(str(row.get('amount', 0))),
                    currency=str(row.get('currency', 'EUR')),
                    transaction_date=pd.to_datetime(row.get('transaction_date')),
                    transaction_type=str(row.get('transaction_type', 'UNKNOWN')),
                    status=str(row.get('status', 'UNKNOWN')),
                    rejection_reason=row.get('rejection_reason'),
                    rejection_code=row.get('rejection_code'),
                    processing_date=pd.to_datetime(row.get('processing_date')) if row.get('processing_date') else None,
                    merchant_id=row.get('merchant_id'),
                    merchant_name=row.get('merchant_name'),
                    card_number_masked=row.get('card_number_masked')
                )
                
                transactions.append(transaction)
                
            except Exception as e:
                self.logger.warning(f"Errore nel parsing riga {index}: {e}")
                continue
        
        return transactions
    
    def get_transactions(self) -> List[Transaction]:
        """
        Ottiene tutte le transazioni caricate.
        
        Returns:
            Lista delle transazioni
        """
        return self._transactions.copy()
    
    def get_raw_data(self) -> Optional[pd.DataFrame]:
        """
        Ottiene i dati grezzi come DataFrame.
        
        Returns:
            DataFrame con i dati grezzi
        """
        return self._raw_data.copy() if self._raw_data is not None else None