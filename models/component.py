"""
Modello per i componenti Oracle e i risultati della ricerca.
"""

from dataclasses import dataclass
from typing import Optional, Dict, Any, List
from datetime import datetime
import pandas as pd


@dataclass
class Component:
    """Classe per rappresentare un componente Oracle."""
    
    nome_componente: str
    prefisso: str
    schema_database: str
    ufficio_it: Optional[str] = None
    responsabile_informatico: Optional[str] = None
    
    def __post_init__(self):
        """Validazione e normalizzazione dei dati."""
        # Normalizza i nomi
        self.nome_componente = self.nome_componente.strip() if self.nome_componente else ""
        self.prefisso = self.prefisso.strip() if self.prefisso else ""
        self.schema_database = self.schema_database.strip() if self.schema_database else ""
        
        if self.ufficio_it:
            self.ufficio_it = self.ufficio_it.strip()
        if self.responsabile_informatico:
            self.responsabile_informatico = self.responsabile_informatico.strip()
    
    @classmethod
    def from_dataframe_row(cls, row: pd.Series) -> 'Component':
        """Crea un Component da una riga di DataFrame."""
        return cls(
            nome_componente=str(row.get('Nome_Componente', '')),
            prefisso=str(row.get('Prefisso', '')),
            schema_database=str(row.get('Schema_Database', '')),
            ufficio_it=str(row.get('Ufficio IT', '')) if pd.notna(row.get('Ufficio IT')) else None,
            responsabile_informatico=str(row.get('Responsabile Informatico', '')) if pd.notna(row.get('Responsabile Informatico')) else None
        )
    
    def to_dict(self) -> Dict[str, Any]:
        """Converte il componente in dizionario."""
        return {
            'Nome_Componente': self.nome_componente,
            'Prefisso': self.prefisso,
            'Schema_Database': self.schema_database,
            'Ufficio IT': self.ufficio_it,
            'Responsabile Informatico': self.responsabile_informatico
        }
    
    def matches_prefix(self, table_name: str) -> bool:
        """Verifica se una tabella corrisponde al prefisso di questo componente."""
        if not self.prefisso or not table_name:
            return False
        return table_name.upper().startswith(self.prefisso.upper())


@dataclass
class LookupResult:
    """Risultato di una ricerca di componenti."""
    
    table_name: str
    component: Optional[Component] = None
    matched: bool = False
    error: Optional[str] = None
    lookup_time: Optional[datetime] = None
    
    def __post_init__(self):
        """Inizializzazione post-creazione."""
        if self.lookup_time is None:
            self.lookup_time = datetime.now()
        
        # Se abbiamo un componente, siamo matched
        if self.component is not None:
            self.matched = True
    
    @property
    def component_name(self) -> str:
        """Restituisce il nome del componente se trovato."""
        return self.component.nome_componente if self.component else "NON TROVATO"
    
    @property
    def schema_database(self) -> str:
        """Restituisce lo schema database se trovato."""
        return self.component.schema_database if self.component else ""
    
    @property
    def ufficio_it(self) -> str:
        """Restituisce l'ufficio IT se disponibile."""
        return self.component.ufficio_it if self.component and self.component.ufficio_it else ""
    
    @property
    def responsabile_informatico(self) -> str:
        """Restituisce il responsabile informatico se disponibile."""
        return self.component.responsabile_informatico if self.component and self.component.responsabile_informatico else ""
    
    def to_dict(self) -> Dict[str, Any]:
        """Converte il risultato in dizionario per Excel."""
        base_dict = {
            'TABLE_NAME': self.table_name,
            'COMPONENTE': self.component_name,
            'SCHEMA_DATABASE': self.schema_database,
            'UFFICIO IT': self.ufficio_it,
            'RESPONSABILE INFORMATICO': self.responsabile_informatico
        }
        
        if self.error:
            base_dict['ERRORE'] = self.error
            
        return base_dict


@dataclass
class ProcessingStats:
    """Statistiche di elaborazione."""
    
    total_tables: int = 0
    matched_tables: int = 0
    unmatched_tables: int = 0
    errors: int = 0
    start_time: Optional[datetime] = None
    end_time: Optional[datetime] = None
    
    def __post_init__(self):
        """Inizializzazione."""
        if self.start_time is None:
            self.start_time = datetime.now()
    
    def finish(self):
        """Marca la fine dell'elaborazione."""
        self.end_time = datetime.now()
    
    @property
    def duration(self) -> Optional[float]:
        """Durata dell'elaborazione in secondi."""
        if self.start_time and self.end_time:
            return (self.end_time - self.start_time).total_seconds()
        return None
    
    @property
    def match_rate(self) -> float:
        """Percentuale di match."""
        if self.total_tables > 0:
            return (self.matched_tables / self.total_tables) * 100
        return 0.0
    
    def add_result(self, result: LookupResult):
        """Aggiunge un risultato alle statistiche."""
        self.total_tables += 1
        
        if result.error:
            self.errors += 1
        elif result.matched:
            self.matched_tables += 1
        else:
            self.unmatched_tables += 1
    
    def __str__(self) -> str:
        """Rappresentazione stringa delle statistiche."""
        duration_str = f"{self.duration:.2f}s" if self.duration else "in corso"
        return (f"Statistiche: {self.total_tables} tabelle elaborate, "
                f"{self.matched_tables} trovate ({self.match_rate:.1f}%), "
                f"{self.unmatched_tables} non trovate, "
                f"{self.errors} errori - Tempo: {duration_str}")
