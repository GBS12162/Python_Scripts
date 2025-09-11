"""
Servizio per la gestione dei componenti Oracle.
Gestisce il caricamento, la cache e l'accesso ai dati dei componenti.
"""

import pandas as pd
from typing import Dict, List, Optional, Set
from pathlib import Path
import logging
from functools import lru_cache

from models.component import Component
from models.config import Config


class ComponentService:
    """Servizio per la gestione dei componenti Oracle."""
    
    def __init__(self, config: Config):
        """
        Inizializza il servizio dei componenti.
        
        Args:
            config: Configurazione dell'applicazione
        """
        self.config = config
        self.logger = logging.getLogger(__name__)
        self._components_df: Optional[pd.DataFrame] = None
        self._components_cache: Dict[str, Component] = {}
        self._prefix_index: Dict[str, List[Component]] = {}
        self._loaded = False
    
    def load_components(self, file_path: Optional[str] = None) -> bool:
        """
        Carica i componenti dal file CSV.
        
        Args:
            file_path: Percorso del file CSV (opzionale)
            
        Returns:
            True se il caricamento è riuscito
        """
        try:
            csv_path = file_path or self.config.components_file
            
            if not Path(csv_path).exists():
                self.logger.error(f"File componenti non trovato: {csv_path}")
                return False
            
            # Carica il CSV
            self.logger.info(f"Caricamento componenti da: {csv_path}")
            
            # Prova diversi encoding
            encodings = ['utf-8', 'latin-1', 'cp1252']
            df = None
            
            for encoding in encodings:
                try:
                    # Prova prima con separatore punto e virgola, poi con virgola
                    for sep in [';', ',']:
                        try:
                            df = pd.read_csv(csv_path, encoding=encoding, sep=sep)
                            self.logger.info(f"File caricato con encoding: {encoding}, separatore: '{sep}'")
                            break
                        except pd.errors.ParserError:
                            continue
                    if df is not None:
                        break
                except UnicodeDecodeError:
                    continue
            
            if df is None:
                self.logger.error("Impossibile decodificare il file CSV")
                return False
            
            # Validazione colonne richieste
            required_columns = ['Nome_Componente', 'Prefisso', 'Schema_Database']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                self.logger.error(f"Colonne mancanti nel CSV: {missing_columns}")
                return False
            
            # Pulizia e validazione dati
            df = self._clean_dataframe(df)
            
            self._components_df = df
            self._build_caches()
            self._loaded = True
            
            self.logger.info(f"Caricati {len(df)} componenti con successo")
            return True
            
        except Exception as e:
            self.logger.error(f"Errore nel caricamento componenti: {e}")
            return False
    
    def _clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Pulisce e valida il DataFrame dei componenti.
        
        Args:
            df: DataFrame da pulire
            
        Returns:
            DataFrame pulito
        """
        # Rimuovi righe vuote
        df = df.dropna(subset=['Nome_Componente', 'Prefisso'], how='all')
        
        # Converti in stringa e pulisci spazi
        string_columns = ['Nome_Componente', 'Prefisso', 'Schema_Database']
        for col in string_columns:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()
        
        # Gestisci colonne opzionali
        optional_columns = ['Ufficio IT', 'Responsabile Informatico']
        for col in optional_columns:
            if col not in df.columns:
                df[col] = ''
            else:
                df[col] = df[col].fillna('').astype(str).str.strip()
        
        # Rimuovi duplicati per prefisso (mantieni il primo)
        df = df.drop_duplicates(subset=['Prefisso'], keep='first')
        
        # Ordina per lunghezza prefisso (decrescente) per matching ottimale
        df = df.sort_values('Prefisso', key=lambda x: x.str.len(), ascending=False)
        
        return df.reset_index(drop=True)
    
    def _build_caches(self):
        """Costruisce le cache per accesso rapido ai componenti."""
        if self._components_df is None:
            return
        
        self._components_cache.clear()
        self._prefix_index.clear()
        
        for _, row in self._components_df.iterrows():
            try:
                component = Component.from_dataframe_row(row)
                
                # Cache per nome componente
                self._components_cache[component.nome_componente] = component
                
                # Indice per prefisso
                prefix_upper = component.prefisso.upper()
                if prefix_upper not in self._prefix_index:
                    self._prefix_index[prefix_upper] = []
                self._prefix_index[prefix_upper].append(component)
                
            except Exception as e:
                self.logger.warning(f"Errore nella creazione del componente dalla riga: {e}")
    
    @property
    def is_loaded(self) -> bool:
        """Verifica se i componenti sono stati caricati."""
        return self._loaded
    
    @property
    def component_count(self) -> int:
        """Numero di componenti caricati."""
        return len(self._components_cache) if self._loaded else 0
    
    def get_component_by_name(self, name: str) -> Optional[Component]:
        """
        Ottiene un componente per nome.
        
        Args:
            name: Nome del componente
            
        Returns:
            Componente se trovato, None altrimenti
        """
        if not self._loaded:
            return None
        
        return self._components_cache.get(name)
    
    def find_component_by_table_name(self, table_name: str) -> Optional[Component]:
        """
        Trova un componente basato sul nome della tabella.
        
        Args:
            table_name: Nome della tabella Oracle
            
        Returns:
            Componente corrispondente se trovato
        """
        if not self._loaded or not table_name:
            return None
        
        table_upper = table_name.upper()
        
        # Cerca per prefisso, ordinati per lunghezza decrescente
        for prefix, components in self._prefix_index.items():
            if table_upper.startswith(prefix):
                # Restituisci il primo componente (dovrebbe essere unico per prefisso)
                return components[0]
        
        return None
    
    @lru_cache(maxsize=10000)
    def cached_lookup(self, table_name: str) -> Optional[Component]:
        """
        Lookup con cache per migliorare le performance.
        
        Args:
            table_name: Nome della tabella
            
        Returns:
            Componente se trovato
        """
        return self.find_component_by_table_name(table_name)
    
    def get_all_components(self) -> List[Component]:
        """
        Restituisce tutti i componenti caricati.
        
        Returns:
            Lista di tutti i componenti
        """
        if not self._loaded:
            return []
        
        return list(self._components_cache.values())
    
    def get_prefixes(self) -> Set[str]:
        """
        Restituisce tutti i prefissi disponibili.
        
        Returns:
            Set di tutti i prefissi
        """
        if not self._loaded:
            return set()
        
        return set(self._prefix_index.keys())
    
    def get_statistics(self) -> Dict[str, any]:
        """
        Restituisce statistiche sui componenti caricati.
        
        Returns:
            Dizionario con statistiche
        """
        if not self._loaded:
            return {"loaded": False}
        
        stats = {
            "loaded": True,
            "total_components": len(self._components_cache),
            "unique_prefixes": len(self._prefix_index),
            "components_with_office": sum(1 for c in self._components_cache.values() 
                                        if c.ufficio_it),
            "components_with_responsible": sum(1 for c in self._components_cache.values() 
                                             if c.responsabile_informatico)
        }
        
        if self._components_df is not None:
            stats["dataframe_memory_usage"] = self._components_df.memory_usage(deep=True).sum()
        
        return stats
    
    def reload(self) -> bool:
        """
        Ricarica i componenti dal file.
        
        Returns:
            True se il ricaricamento è riuscito
        """
        self._loaded = False
        self._components_df = None
        self._components_cache.clear()
        self._prefix_index.clear()
        
        # Pulisci la cache LRU
        self.cached_lookup.cache_clear()
        
        return self.load_components()
    
    def export_to_dataframe(self) -> Optional[pd.DataFrame]:
        """
        Esporta i componenti come DataFrame.
        
        Returns:
            DataFrame con i componenti
        """
        if not self._loaded:
            return None
        
        return self._components_df.copy() if self._components_df is not None else None