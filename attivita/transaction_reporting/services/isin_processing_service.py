"""
Servizio per l'elaborazione dei dati ISIN e ordini.
Gestisce il parsing dei file Excel e l'elaborazione dei gruppi ISIN-OCCORRENZE.
"""

import logging
import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple
from decimal import Decimal

from models.transaction_reporting import (
    ISINGroup, Order, ProcessingConfig, ProcessingResult, QualityControlResult
)
from utils.file_utils import detect_file_encoding


class ISINProcessingService:
    """Servizio per l'elaborazione dei dati ISIN e ordini."""
    
    def __init__(self):
        """Inizializza il servizio."""
        self.logger = logging.getLogger(__name__)
        self._raw_data: Optional[pd.DataFrame] = None
        self._processed_groups: List[ISINGroup] = []
    
    def process_excel_file(
        self, 
        file_path: str, 
        config: ProcessingConfig
    ) -> ProcessingResult:
        """
        Elabora il file Excel con i dati ISIN e ordini.
        
        Args:
            file_path: Percorso del file Excel
            config: Configurazione di elaborazione
            
        Returns:
            Risultato dell'elaborazione
        """
        try:
            start_time = datetime.now()
            self.logger.info(f"Inizio elaborazione file: {file_path}")
            
            # Carica il file Excel
            df = self._load_excel_file(file_path)
            if df is None or df.empty:
                return ProcessingResult(
                    success=False,
                    message="File Excel vuoto o non caricabile",
                    validation_errors=["File non valido"]
                )
            
            self._raw_data = df
            self.logger.info(f"Caricate {len(df)} righe dal file Excel")
            
            # Valida le colonne richieste
            validation_result = self._validate_columns(df, config)
            if not validation_result[0]:
                return ProcessingResult(
                    success=False,
                    message="Colonne richieste mancanti",
                    validation_errors=validation_result[1]
                )
            
            # Elabora i gruppi ISIN
            isin_groups = self._process_isin_groups(df, config)
            self._processed_groups = isin_groups
            
            # Calcola statistiche
            stats = self._calculate_statistics(isin_groups)
            
            # Esegui controlli di qualità
            quality_results = self._perform_quality_controls(isin_groups, config)
            
            processing_time = (datetime.now() - start_time).total_seconds()
            
            return ProcessingResult(
                success=True,
                message=f"Elaborazione completata: {len(isin_groups)} gruppi ISIN processati",
                isin_groups=isin_groups,
                total_isin=len(isin_groups),
                total_orders=sum(len(group.orders) for group in isin_groups),
                processing_stats=stats,
                processing_time=processing_time,
                warnings=self._collect_warnings(quality_results)
            )
            
        except Exception as e:
            self.logger.error(f"Errore nell'elaborazione: {e}")
            return ProcessingResult(
                success=False,
                message=f"Errore nell'elaborazione: {str(e)}",
                validation_errors=[str(e)]
            )
    
    def _load_excel_file(self, file_path: str) -> Optional[pd.DataFrame]:
        """Carica il file Excel."""
        try:
            file_path_obj = Path(file_path)
            if not file_path_obj.exists():
                raise FileNotFoundError(f"File non trovato: {file_path}")
            
            # Prova a caricare con diversi engine
            try:
                df = pd.read_excel(file_path, engine='openpyxl')
            except Exception:
                try:
                    df = pd.read_excel(file_path, engine='xlrd')
                except Exception:
                    df = pd.read_excel(file_path)
            
            self.logger.info(f"File Excel caricato: {len(df)} righe, {len(df.columns)} colonne")
            return df
            
        except Exception as e:
            self.logger.error(f"Errore nel caricamento file Excel: {e}")
            return None
    
    def _validate_columns(
        self, 
        df: pd.DataFrame, 
        config: ProcessingConfig
    ) -> Tuple[bool, List[str]]:
        """Valida la presenza delle colonne richieste."""
        try:
            errors = []
            required_columns = [config.isin_column, config.occorrenze_column]
            
            # Verifica colonne base
            for col in required_columns:
                if col not in df.columns:
                    errors.append(f"Colonna richiesta mancante: {col}")
            
            # Verifica colonne di controllo
            missing_control_cols = []
            for col in config.controllo_columns:
                if col not in df.columns:
                    missing_control_cols.append(col)
            
            if missing_control_cols:
                errors.append(f"Colonne di controllo mancanti: {', '.join(missing_control_cols)}")
            
            # Log delle colonne disponibili per debug
            self.logger.info(f"Colonne disponibili: {list(df.columns)}")
            
            return len(errors) == 0, errors
            
        except Exception as e:
            self.logger.error(f"Errore nella validazione colonne: {e}")
            return False, [str(e)]
    
    def _process_isin_groups(
        self, 
        df: pd.DataFrame, 
        config: ProcessingConfig
    ) -> List[ISINGroup]:
        """Elabora i gruppi ISIN dal DataFrame."""
        try:
            groups = []
            current_row = 0
            
            while current_row < len(df):
                try:
                    # Leggi ISIN e OCCORRENZE dalla riga corrente
                    isin = df.iloc[current_row][config.isin_column]
                    occorrenze = df.iloc[current_row][config.occorrenze_column]
                    
                    # Salta righe con ISIN vuoto se configurato
                    if config.skip_empty_isin and (pd.isna(isin) or str(isin).strip() == ""):
                        current_row += 1
                        continue
                    
                    # Converti occorrenze in intero
                    if pd.isna(occorrenze):
                        occorrenze = 0
                    else:
                        occorrenze = int(float(occorrenze))
                    
                    # Valida occorrenze
                    if config.validate_occorrenze and occorrenze < 0:
                        self.logger.warning(f"ISIN {isin}: occorrenze negative ({occorrenze})")
                        current_row += 1
                        continue
                    
                    # Leggi colonne di controllo dalla riga ISIN
                    controlli = {}
                    for i, col in enumerate(config.controllo_columns):
                        if col in df.columns:
                            controlli[f"controllo_{i+1}"] = df.iloc[current_row][col]
                    
                    # Elabora gli ordini per questo ISIN
                    orders = []
                    orders_start_row = current_row + 1
                    
                    for order_idx in range(occorrenze):
                        order_row = orders_start_row + order_idx
                        if order_row < len(df):
                            order = self._create_order_from_row(df, order_row, str(isin), config)
                            if order:
                                orders.append(order)
                        else:
                            self.logger.warning(f"ISIN {isin}: ordine {order_idx+1} fuori range")
                    
                    # Crea il gruppo ISIN
                    group = ISINGroup(
                        isin=str(isin),
                        occorrenze=occorrenze,
                        orders=orders,
                        controllo_1=controlli.get("controllo_1"),
                        controllo_2=controlli.get("controllo_2"),
                        controllo_3=controlli.get("controllo_3"),
                        controllo_4=controlli.get("controllo_4")
                    )
                    
                    groups.append(group)
                    self.logger.debug(f"Gruppo ISIN creato: {isin} con {len(orders)} ordini")
                    
                    # Avanza alla prossima riga ISIN
                    current_row = orders_start_row + occorrenze
                    
                except Exception as e:
                    self.logger.error(f"Errore elaborazione riga {current_row}: {e}")
                    current_row += 1
                    continue
            
            self.logger.info(f"Elaborati {len(groups)} gruppi ISIN")
            return groups
            
        except Exception as e:
            self.logger.error(f"Errore nell'elaborazione gruppi ISIN: {e}")
            return []
    
    def _create_order_from_row(
        self, 
        df: pd.DataFrame, 
        row_idx: int, 
        isin: str, 
        config: ProcessingConfig
    ) -> Optional[Order]:
        """Crea un oggetto Order da una riga del DataFrame."""
        try:
            if row_idx >= len(df):
                return None
            
            row = df.iloc[row_idx]
            
            # Genera ID ordine se non presente
            order_id = row.get('ORDER_ID', f"{isin}_{row_idx}")
            
            # Estrai campi standard (adatta secondo la struttura del tuo Excel)
            additional_fields = {}
            for col in df.columns:
                if col not in [config.isin_column, config.occorrenze_column] + config.controllo_columns:
                    value = row[col]
                    if not pd.isna(value):
                        additional_fields[col] = value
            
            # Crea l'ordine
            order = Order(
                order_id=str(order_id),
                isin=isin,
                quantity=self._safe_decimal_conversion(row.get('QUANTITY')),
                price=self._safe_decimal_conversion(row.get('PRICE')),
                order_type=str(row.get('ORDER_TYPE', '')),
                order_date=self._safe_date_conversion(row.get('ORDER_DATE')),
                settlement_date=self._safe_date_conversion(row.get('SETTLEMENT_DATE')),
                client_id=str(row.get('CLIENT_ID', '')),
                broker_id=str(row.get('BROKER_ID', '')),
                account_id=str(row.get('ACCOUNT_ID', '')),
                currency=str(row.get('CURRENCY', 'EUR')),
                market=str(row.get('MARKET', '')),
                status=str(row.get('STATUS', '')),
                additional_fields=additional_fields
            )
            
            return order
            
        except Exception as e:
            self.logger.error(f"Errore creazione ordine riga {row_idx}: {e}")
            return None
    
    def _safe_decimal_conversion(self, value) -> Optional[Decimal]:
        """Conversione sicura in Decimal."""
        try:
            if pd.isna(value) or value == "":
                return None
            return Decimal(str(value))
        except:
            return None
    
    def _safe_date_conversion(self, value) -> Optional[datetime]:
        """Conversione sicura in datetime."""
        try:
            if pd.isna(value) or value == "":
                return None
            if isinstance(value, datetime):
                return value
            return pd.to_datetime(value)
        except:
            return None
    
    def _calculate_statistics(self, groups: List[ISINGroup]) -> Dict[str, Any]:
        """Calcola statistiche sui gruppi elaborati."""
        try:
            total_groups = len(groups)
            total_orders = sum(len(group.orders) for group in groups)
            
            # Statistiche per ISIN
            isin_with_orders = sum(1 for group in groups if len(group.orders) > 0)
            avg_orders_per_isin = total_orders / total_groups if total_groups > 0 else 0
            
            # Statistiche occorrenze
            occorrenze_list = [group.occorrenze for group in groups]
            min_occorrenze = min(occorrenze_list) if occorrenze_list else 0
            max_occorrenze = max(occorrenze_list) if occorrenze_list else 0
            avg_occorrenze = sum(occorrenze_list) / len(occorrenze_list) if occorrenze_list else 0
            
            # Controlli di qualità
            controlli_stats = {}
            for i in range(1, 5):
                col_name = f"controllo_{i}"
                values = [getattr(group, col_name) for group in groups if getattr(group, col_name) is not None]
                controlli_stats[col_name] = {
                    "total": len(values),
                    "unique_values": len(set(str(v) for v in values)) if values else 0
                }
            
            return {
                "total_isin_groups": total_groups,
                "total_orders": total_orders,
                "isin_with_orders": isin_with_orders,
                "avg_orders_per_isin": round(avg_orders_per_isin, 2),
                "occorrenze_stats": {
                    "min": min_occorrenze,
                    "max": max_occorrenze,
                    "avg": round(avg_occorrenze, 2)
                },
                "controlli_stats": controlli_stats
            }
            
        except Exception as e:
            self.logger.error(f"Errore nel calcolo statistiche: {e}")
            return {}
    
    def _perform_quality_controls(
        self, 
        groups: List[ISINGroup], 
        config: ProcessingConfig
    ) -> List[QualityControlResult]:
        """Esegue controlli di qualità sui gruppi."""
        try:
            results = []
            
            for group in groups:
                result = QualityControlResult(
                    isin=group.isin,
                    total_orders=len(group.orders)
                )
                
                # Controllo: numero ordini vs occorrenze
                if len(group.orders) != group.occorrenze:
                    result.validation_errors.append(
                        f"Mismatch ordini: attesi {group.occorrenze}, trovati {len(group.orders)}"
                    )
                    result.controlli_failed += 1
                else:
                    result.controlli_passed += 1
                
                # Controllo: presenza controlli
                controlli_presenti = 0
                for i in range(1, 5):
                    controllo = getattr(group, f"controllo_{i}")
                    if controllo is not None and str(controllo).strip() != "":
                        controlli_presenti += 1
                        result.controlli_details[f"controllo_{i}"] = str(controllo)
                
                if controlli_presenti == 4:
                    result.controlli_passed += 1
                else:
                    result.controlli_failed += 1
                    result.validation_errors.append(
                        f"Controlli incompleti: {controlli_presenti}/4 presenti"
                    )
                
                # Controllo: ordini validi
                valid_orders = sum(1 for order in group.orders if order.order_id and order.isin)
                if valid_orders == len(group.orders):
                    result.controlli_passed += 1
                else:
                    result.controlli_failed += 1
                    result.validation_errors.append(
                        f"Ordini non validi: {len(group.orders) - valid_orders}/{len(group.orders)}"
                    )
                
                results.append(result)
            
            return results
            
        except Exception as e:
            self.logger.error(f"Errore nei controlli di qualità: {e}")
            return []
    
    def _collect_warnings(self, quality_results: List[QualityControlResult]) -> List[str]:
        """Raccoglie avvisi dai risultati dei controlli."""
        warnings = []
        
        for result in quality_results:
            if result.validation_errors:
                warnings.extend([f"ISIN {result.isin}: {error}" for error in result.validation_errors])
        
        return warnings
    
    def get_processed_groups(self) -> List[ISINGroup]:
        """Ottiene i gruppi elaborati."""
        return self._processed_groups.copy()
    
    def get_raw_data(self) -> Optional[pd.DataFrame]:
        """Ottiene i dati grezzi."""
        return self._raw_data.copy() if self._raw_data is not None else None