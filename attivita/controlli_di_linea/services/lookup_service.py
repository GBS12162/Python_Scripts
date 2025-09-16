"""
Servizio per l'elaborazione dei lookup dei componenti Oracle.
Gestisce la ricerca parallela e l'elaborazione di grandi dataset.
"""

import pandas as pd
from typing import List, Iterator, Optional, Callable, Dict, Any
from concurrent.futures import ProcessPoolExecutor, as_completed
import logging
from tqdm import tqdm

from models.component import Component, LookupResult, ProcessingStats
from models.config import Config
from services.component_service import ComponentService


class LookupService:
    """Servizio per l'elaborazione dei lookup dei componenti."""
    
    def __init__(self, config: Config, component_service: ComponentService):
        """
        Inizializza il servizio di lookup.
        
        Args:
            config: Configurazione dell'applicazione
            component_service: Servizio dei componenti
        """
        self.config = config
        self.component_service = component_service
        self.logger = logging.getLogger(__name__)
        self.stats = ProcessingStats()
    
    def process_tables(self, table_names: List[str], 
                      progress_callback: Optional[Callable[[int, int], None]] = None) -> Iterator[LookupResult]:
        """
        Elabora una lista di nomi di tabelle.
        
        Args:
            table_names: Lista dei nomi delle tabelle
            progress_callback: Callback per aggiornamenti di progresso
            
        Yields:
            LookupResult per ogni tabella elaborata
        """
        if not self.component_service.is_loaded:
            self.logger.error("ComponentService non è caricato")
            return
        
        self.stats = ProcessingStats()
        total_tables = len(table_names)
        
        self.logger.info(f"Inizio elaborazione di {total_tables} tabelle")
        
        if self.config.enable_multiprocessing and total_tables > 1000:
            # Elaborazione parallela per dataset grandi
            yield from self._process_parallel(table_names, progress_callback)
        else:
            # Elaborazione sequenziale per dataset piccoli
            yield from self._process_sequential(table_names, progress_callback)
        
        self.stats.finish()
        self.logger.info(f"Elaborazione completata: {self.stats}")
    
    def _process_sequential(self, table_names: List[str], 
                          progress_callback: Optional[Callable[[int, int], None]] = None) -> Iterator[LookupResult]:
        """
        Elaborazione sequenziale delle tabelle.
        
        Args:
            table_names: Lista dei nomi delle tabelle
            progress_callback: Callback per progresso
            
        Yields:
            LookupResult per ogni tabella
        """
        total = len(table_names)
        
        # Usa tqdm se non c'è callback personalizzato
        if progress_callback is None:
            table_iter = tqdm(table_names, desc="Elaborazione tabelle", unit="tabelle")
        else:
            table_iter = table_names
        
        for i, table_name in enumerate(table_iter):
            try:
                result = self._lookup_single_table(table_name)
                self.stats.add_result(result)
                
                if progress_callback:
                    progress_callback(i + 1, total)
                
                yield result
                
            except Exception as e:
                self.logger.error(f"Errore nell'elaborazione di {table_name}: {e}")
                error_result = LookupResult(table_name=table_name, error=str(e))
                self.stats.add_result(error_result)
                yield error_result
    
    def _process_parallel(self, table_names: List[str], 
                         progress_callback: Optional[Callable[[int, int], None]] = None) -> Iterator[LookupResult]:
        """
        Elaborazione parallela delle tabelle.
        
        Args:
            table_names: Lista dei nomi delle tabelle
            progress_callback: Callback per progresso
            
        Yields:
            LookupResult per ogni tabella elaborata
        """
        total = len(table_names)
        chunk_size = min(self.config.chunk_size, max(1, total // self.config.max_workers))
        
        # Suddividi in chunk
        chunks = [table_names[i:i + chunk_size] for i in range(0, total, chunk_size)]
        
        self.logger.info(f"Elaborazione parallela: {len(chunks)} chunk, {self.config.max_workers} worker")
        
        # Prepara i dati dei componenti per i worker
        components_data = self._prepare_components_for_workers()
        
        processed = 0
        
        with ProcessPoolExecutor(max_workers=self.config.max_workers) as executor:
            # Sottometti tutti i job
            future_to_chunk = {
                executor.submit(_process_chunk_worker, chunk, components_data): chunk 
                for chunk in chunks
            }
            
            # Raccogli i risultati man mano che sono pronti
            for future in as_completed(future_to_chunk):
                try:
                    chunk_results = future.result()
                    
                    for result in chunk_results:
                        self.stats.add_result(result)
                        processed += 1
                        
                        if progress_callback:
                            progress_callback(processed, total)
                        
                        yield result
                        
                except Exception as e:
                    chunk = future_to_chunk[future]
                    self.logger.error(f"Errore nell'elaborazione del chunk: {e}")
                    
                    # Genera risultati di errore per il chunk
                    for table_name in chunk:
                        error_result = LookupResult(table_name=table_name, error=str(e))
                        self.stats.add_result(error_result)
                        processed += 1
                        
                        if progress_callback:
                            progress_callback(processed, total)
                        
                        yield error_result
    
    def _prepare_components_for_workers(self) -> Dict[str, List[Dict[str, Any]]]:
        """
        Prepara i dati dei componenti per i worker paralleli.
        
        Returns:
            Dizionario con i dati dei componenti serializzati
        """
        components_data = {}
        
        for prefix, components in self.component_service._prefix_index.items():
            components_data[prefix] = [comp.to_dict() for comp in components]
        
        return components_data
    
    def _lookup_single_table(self, table_name: str) -> LookupResult:
        """
        Esegue il lookup per una singola tabella.
        
        Args:
            table_name: Nome della tabella
            
        Returns:
            Risultato del lookup
        """
        try:
            # Usa la cache se abilitata
            if self.config.cache_components:
                component = self.component_service.cached_lookup(table_name)
            else:
                component = self.component_service.find_component_by_table_name(table_name)
            
            return LookupResult(
                table_name=table_name,
                component=component,
                matched=component is not None
            )
            
        except Exception as e:
            self.logger.error(f"Errore nel lookup per {table_name}: {e}")
            return LookupResult(table_name=table_name, error=str(e))
    
    def process_dataframe(self, df: pd.DataFrame, table_column: str = "TABLE_NAME",
                         progress_callback: Optional[Callable[[int, int], None]] = None) -> pd.DataFrame:
        """
        Elabora un DataFrame aggiungendo le informazioni dei componenti.
        
        Args:
            df: DataFrame con i nomi delle tabelle
            table_column: Nome della colonna con i nomi delle tabelle
            progress_callback: Callback per progresso
            
        Returns:
            DataFrame arricchito con informazioni sui componenti
        """
        if table_column not in df.columns:
            raise ValueError(f"Colonna {table_column} non trovata nel DataFrame")
        
        # Estrai i nomi delle tabelle unici per ottimizzare
        unique_tables = df[table_column].unique().tolist()
        self.logger.info(f"Elaborazione di {len(unique_tables)} tabelle uniche da {len(df)} righe")
        
        # Crea un mapping dei risultati
        results_map = {}
        
        for result in self.process_tables(unique_tables, progress_callback):
            results_map[result.table_name] = result
        
        # Applica i risultati al DataFrame originale
        def apply_result(table_name):
            result = results_map.get(table_name)
            if result:
                return pd.Series({
                    'COMPONENTE': result.component_name,
                    'SCHEMA_DATABASE': result.schema_database,
                    'UFFICIO IT': result.ufficio_it,
                    'RESPONSABILE INFORMATICO': result.responsabile_informatico
                })
            else:
                return pd.Series({
                    'COMPONENTE': 'NON TROVATO',
                    'SCHEMA_DATABASE': '',
                    'UFFICIO IT': '',
                    'RESPONSABILE INFORMATICO': ''
                })
        
        # Applica la funzione usando tqdm per il progresso
        tqdm.pandas(desc="Applicazione risultati")
        result_df = df[table_column].progress_apply(apply_result)
        
        # Unisci con il DataFrame originale
        enriched_df = pd.concat([df, result_df], axis=1)
        
        return enriched_df
    
    def get_statistics(self) -> ProcessingStats:
        """
        Restituisce le statistiche dell'ultima elaborazione.
        
        Returns:
            Statistiche di elaborazione
        """
        return self.stats
    
    def create_summary_report(self) -> Dict[str, Any]:
        """
        Crea un report riassuntivo dell'elaborazione.
        
        Returns:
            Dizionario con il report
        """
        stats = self.stats
        
        report = {
            "elaborazione": {
                "tabelle_totali": stats.total_tables,
                "tabelle_trovate": stats.matched_tables,
                "tabelle_non_trovate": stats.unmatched_tables,
                "errori": stats.errors,
                "percentuale_successo": stats.match_rate,
                "durata_secondi": stats.duration
            },
            "configurazione": {
                "elaborazione_parallela": self.config.enable_multiprocessing,
                "numero_worker": self.config.max_workers,
                "dimensione_chunk": self.config.chunk_size,
                "cache_abilitata": self.config.cache_components
            },
            "componenti": self.component_service.get_statistics()
        }
        
        return report


def _process_chunk_worker(table_names: List[str], components_data: Dict[str, List[Dict[str, Any]]]) -> List[LookupResult]:
    """
    Worker function per l'elaborazione parallela di un chunk di tabelle.
    
    Args:
        table_names: Lista di nomi di tabelle da elaborare
        components_data: Dati dei componenti serializzati
        
    Returns:
        Lista di LookupResult
    """
    results = []
    
    # Ricostruisci l'indice dei prefissi localmente
    prefix_index = {}
    for prefix, comp_dicts in components_data.items():
        components = []
        for comp_dict in comp_dicts:
            component = Component(
                nome_componente=comp_dict['Nome_Componente'],
                prefisso=comp_dict['Prefisso'],
                schema_database=comp_dict['Schema_Database'],
                ufficio_it=comp_dict.get('Ufficio IT'),
                responsabile_informatico=comp_dict.get('Responsabile Informatico')
            )
            components.append(component)
        prefix_index[prefix] = components
    
    # Elabora ogni tabella nel chunk
    for table_name in table_names:
        try:
            component = None
            table_upper = table_name.upper()
            
            # Cerca per prefisso
            for prefix, components in prefix_index.items():
                if table_upper.startswith(prefix):
                    component = components[0]  # Prendi il primo (dovrebbe essere unico)
                    break
            
            result = LookupResult(
                table_name=table_name,
                component=component,
                matched=component is not None
            )
            
        except Exception as e:
            result = LookupResult(table_name=table_name, error=str(e))
        
        results.append(result)
    
    return results
