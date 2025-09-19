"""
Parallel Processing Service per CON-412 - Versione Threading
Gestisce l'elaborazione parallela usando threading invece di multiprocessing per compatibilità Windows
"""

import logging
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Dict, Any, Callable, Optional, Tuple
import time
from functools import partial
import os
import sys
from pathlib import Path


class ParallelProcessingServiceThreaded:
    """Servizio per l'elaborazione parallela del CON-412 usando threading"""
    
    def __init__(self, max_workers: Optional[int] = None):
        """
        Inizializza il servizio di elaborazione parallela
        
        Args:
            max_workers: Numero massimo di thread worker (default: ottimizzato per I/O)
        """
        self.logger = self._setup_logging()
        # Per I/O bound tasks come API calls, più thread possono aiutare
        self.max_workers = max_workers or min(os.cpu_count() * 2, 16)
        
        self.logger.info(f"ParallelProcessingServiceThreaded inizializzato:")
        self.logger.info(f"  - Max workers: {self.max_workers}")
        self.logger.info(f"  - Tipo executor: ThreadPool")
        
    def _setup_logging(self):
        """Configura il sistema di logging"""
        logger = logging.getLogger(__name__)
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
            logger.setLevel(logging.INFO)
        return logger
    
    def process_esma_validations_parallel(self, isin_list: List[str], 
                                        batch_size: int = 5) -> Dict[str, bool]:
        """
        Processa validazioni ESMA in parallelo usando threading
        
        Args:
            isin_list: Lista di codici ISIN da validare
            batch_size: Dimensione batch per processing
            
        Returns:
            Dizionario {isin: is_valid}
        """
        self.logger.info(f"Avvio validazione ESMA parallela per {len(isin_list)} ISIN")
        start_time = time.time()
        
        # Divide in batch per ottimizzare le richieste API
        batches = self._create_batches(isin_list, batch_size)
        
        results = {}
        completed_batches = 0
        
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            # Crea i task per ogni batch
            future_to_batch = {}
            
            for batch_idx, batch in enumerate(batches):
                future = executor.submit(self._validate_esma_batch_threaded, batch, batch_idx)
                future_to_batch[future] = (batch, batch_idx)
            
            # Processa i risultati man mano che arrivano
            for future in as_completed(future_to_batch):
                batch, batch_idx = future_to_batch[future]
                
                try:
                    batch_results = future.result()
                    results.update(batch_results)
                    completed_batches += 1
                    
                    self.logger.info(f"Batch {batch_idx + 1}/{len(batches)} completato "
                                   f"({len(batch)} ISIN) - Progress: {completed_batches}/{len(batches)}")
                    
                except Exception as e:
                    self.logger.error(f"Errore nel batch {batch_idx}: {e}")
                    # In caso di errore, assume tutti gli ISIN del batch come validi
                    for isin in batch:
                        results[isin] = True
        
        elapsed_time = time.time() - start_time
        success_rate = sum(1 for v in results.values() if v) / len(results) * 100
        
        self.logger.info(f"Validazione ESMA parallela completata:")
        self.logger.info(f"  - Tempo: {elapsed_time:.2f}s")
        self.logger.info(f"  - ISIN validati: {len(results)}")
        self.logger.info(f"  - Successo: {success_rate:.1f}%")
        
        return results
    
    def process_database_checks_parallel(self, order_data: List[Dict[str, Any]], 
                                       batch_size: int = 10) -> Dict[str, str]:
        """
        Processa controlli database in parallelo usando threading
        
        Args:
            order_data: Lista di dati ordini da controllare
            batch_size: Dimensione batch per processing
            
        Returns:
            Dizionario {order_id: status} con stato ordini
        """
        self.logger.info(f"Avvio controllo database parallelo per {len(order_data)} ordini")
        start_time = time.time()
        
        # Divide in batch
        batches = self._create_batches(order_data, batch_size)
        
        results = {}
        completed_batches = 0
        
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            # Crea i task per ogni batch
            future_to_batch = {}
            
            for batch_idx, batch in enumerate(batches):
                future = executor.submit(self._check_database_batch_threaded, batch, batch_idx)
                future_to_batch[future] = (batch, batch_idx)
            
            # Processa i risultati man mano che arrivano
            for future in as_completed(future_to_batch):
                batch, batch_idx = future_to_batch[future]
                
                try:
                    batch_results = future.result()
                    results.update(batch_results)
                    completed_batches += 1
                    
                    self.logger.info(f"Batch DB {batch_idx + 1}/{len(batches)} completato "
                                   f"({len(batch)} ordini) - Progress: {completed_batches}/{len(batches)}")
                    
                except Exception as e:
                    self.logger.error(f"Errore nel batch DB {batch_idx}: {e}")
                    # In caso di errore, assume tutti gli ordini come validi (RF)
                    for order in batch:
                        if 'order_id' in order:
                            results[order['order_id']] = 'RF'
        
        elapsed_time = time.time() - start_time
        rf_count = sum(1 for v in results.values() if v == 'RF')
        
        self.logger.info(f"Controllo database parallelo completato:")
        self.logger.info(f"  - Tempo: {elapsed_time:.2f}s")
        self.logger.info(f"  - Ordini controllati: {len(results)}")
        self.logger.info(f"  - Ordini RF (mantenuti): {rf_count}/{len(results)}")
        
        return results
    
    def _validate_esma_batch_threaded(self, isin_batch: List[str], batch_idx: int) -> Dict[str, bool]:
        """
        Funzione worker per validare un batch di ISIN tramite ESMA (versione threading)
        
        Args:
            isin_batch: Lista di ISIN da validare
            batch_idx: Indice del batch
            
        Returns:
            Dizionario {isin: is_valid}
        """
        try:
            # Import del servizio ESMA - già disponibile nello stesso processo
            from pathlib import Path
            import sys
            project_root = Path(__file__).parent.parent.parent
            sys.path.insert(0, str(project_root))
            
            # Import dinamico del servizio
            import importlib.util
            isin_validation_module_path = Path(__file__).parent / "isin_validation_service.py"
            spec = importlib.util.spec_from_file_location("isin_validation_service", isin_validation_module_path)
            isin_validation_module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(isin_validation_module)
            
            validation_service = isin_validation_module.ISINValidationService()
            
            results = {}
            for isin in isin_batch:
                try:
                    # Valida ISIN tramite ESMA
                    is_valid = validation_service.check_single_isin(isin)
                    results[isin] = is_valid
                    
                    # Piccola pausa per evitare rate limiting
                    time.sleep(0.1)
                    
                except Exception as e:
                    # In caso di errore, assume ISIN valido
                    results[isin] = True
                    self.logger.warning(f"Thread {batch_idx}: Errore validazione ISIN {isin}: {e}")
            
            return results
            
        except Exception as e:
            self.logger.error(f"Thread {batch_idx}: Errore critico: {e}")
            # Fallback: tutti gli ISIN sono validi
            return {isin: True for isin in isin_batch}
    
    def _check_database_batch_threaded(self, order_batch: List[Dict[str, Any]], batch_idx: int) -> Dict[str, str]:
        """
        Funzione worker per controllare un batch di ordini nel database (versione threading)
        
        Args:
            order_batch: Lista di ordini da controllare
            batch_idx: Indice del batch
            
        Returns:
            Dizionario {order_id: status}
        """
        try:
            # Import del servizio database - già disponibile nello stesso processo
            from pathlib import Path
            import sys
            project_root = Path(__file__).parent.parent.parent
            sys.path.insert(0, str(project_root))
            
            # Per ora simula il controllo database
            # In produzione qui ci sarebbe la connessione al database Oracle
            results = {}
            for order in order_batch:
                try:
                    if 'order_id' in order:
                        order_id = order['order_id']
                        # TODO: Implementare query database per controllo stato ordine
                        # status = db_service.check_order_status(order_id)
                        status = 'RF'  # Simula che l'ordine sia valido
                        results[order_id] = status
                        
                        # Simula tempo di query database
                        time.sleep(0.05)
                    
                except Exception as e:
                    # In caso di errore, assume ordine valido
                    if 'order_id' in order:
                        results[order['order_id']] = 'RF'
                    self.logger.warning(f"Thread DB {batch_idx}: Errore controllo ordine: {e}")
            
            return results
            
        except Exception as e:
            self.logger.error(f"Thread DB {batch_idx}: Errore critico: {e}")
            # Fallback: tutti gli ordini sono validi
            return {order.get('order_id', f'unknown_{i}'): 'RF' 
                    for i, order in enumerate(order_batch)}
    
    def _create_batches(self, items: List, batch_size: int) -> List[List]:
        """Crea batch da una lista di elementi"""
        return [items[i:i + batch_size] for i in range(0, len(items), batch_size)]
    
    def get_performance_stats(self) -> Dict[str, Any]:
        """Restituisce statistiche sulle performance"""
        return {
            "cpu_count": os.cpu_count(),
            "max_workers": self.max_workers,
            "executor_type": "ThreadPool"
        }


# Funzione di utilità per testare le performance
def benchmark_parallel_processing_threaded():
    """Testa le performance del processing parallelo con threading"""
    service = ParallelProcessingServiceThreaded()
    
    # Test con ISIN fittizi
    test_isins = [f"IT000{i:07d}" for i in range(20)]
    
    print("Benchmark Parallel Processing (Threading):")
    print(f"Testing con {len(test_isins)} ISIN...")
    
    start_time = time.time()
    results = service.process_esma_validations_parallel(test_isins, batch_size=3)
    elapsed_time = time.time() - start_time
    
    print(f"Completato in {elapsed_time:.2f}s")
    print(f"Performance: {len(test_isins)/elapsed_time:.1f} ISIN/s")
    print(f"Risultati: {len(results)} validazioni")
    print(f"Workers utilizzati: {service.max_workers}")


if __name__ == "__main__":
    # Test del servizio
    benchmark_parallel_processing_threaded()