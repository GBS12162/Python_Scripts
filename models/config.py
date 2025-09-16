"""
Configurazioni per l'applicazione Oracle Component Lookup.
"""

from dataclasses import dataclass, field
from typing import Dict, Any, Optional
import os


@dataclass
class Config:
    """Configurazione principale dell'applicazione."""
    
    # Percorsi di default
    components_file: str = "Components.csv"
    output_directory: str = "Output"
    
    # Configurazioni Excel
    max_rows_per_file: int = 1048576  # Limite Excel
    excel_engine: str = "xlsxwriter"
    
    # Configurazioni compressione
    enable_7z_compression: bool = True
    compression_level: int = 5
    compression_threshold_mb: int = 10
    
    # Configurazioni elaborazione
    chunk_size: int = 10000
    max_workers: int = 4
    enable_multiprocessing: bool = True
    
    # Configurazioni output
    include_statistics: bool = True
    create_summary_report: bool = True
    
    # Configurazioni avanzate
    memory_optimization: bool = True
    cache_components: bool = True
    
    def __post_init__(self):
        """Validazione e normalizzazione della configurazione."""
        # Assicura che i percorsi esistano
        if not os.path.exists(self.output_directory):
            os.makedirs(self.output_directory, exist_ok=True)
        
        # Validazione dei limiti
        if self.max_rows_per_file <= 0:
            self.max_rows_per_file = 1048576
        
        if self.chunk_size <= 0:
            self.chunk_size = 10000
        
        if self.max_workers <= 0:
            self.max_workers = max(1, os.cpu_count() or 1)
        
        # Validazione compressione
        if self.compression_level < 0 or self.compression_level > 9:
            self.compression_level = 5
    
    @classmethod
    def from_dict(cls, config_dict: Dict[str, Any]) -> 'Config':
        """Crea una configurazione da un dizionario."""
        return cls(**{k: v for k, v in config_dict.items() if hasattr(cls, k)})
    
    def to_dict(self) -> Dict[str, Any]:
        """Converte la configurazione in dizionario."""
        return {
            'components_file': self.components_file,
            'output_directory': self.output_directory,
            'max_rows_per_file': self.max_rows_per_file,
            'excel_engine': self.excel_engine,
            'enable_7z_compression': self.enable_7z_compression,
            'compression_level': self.compression_level,
            'compression_threshold_mb': self.compression_threshold_mb,
            'chunk_size': self.chunk_size,
            'max_workers': self.max_workers,
            'enable_multiprocessing': self.enable_multiprocessing,
            'include_statistics': self.include_statistics,
            'create_summary_report': self.create_summary_report,
            'memory_optimization': self.memory_optimization,
            'cache_components': self.cache_components
        }
    
    def optimize_for_size(self, estimated_rows: int):
        """Ottimizza la configurazione in base alla dimensione stimata."""
        if estimated_rows > 100000:
            # Dataset grandi: ottimizza per performance
            self.enable_multiprocessing = True
            self.chunk_size = 50000
            self.memory_optimization = True
            self.cache_components = True
        elif estimated_rows > 10000:
            # Dataset medi: bilanciamento
            self.chunk_size = 20000
            self.memory_optimization = True
        else:
            # Dataset piccoli: semplicità
            self.enable_multiprocessing = False
            self.chunk_size = estimated_rows
            self.memory_optimization = False


@dataclass
class FileConfig:
    """Configurazione per un singolo file di input."""
    
    file_path: str
    encoding: str = "utf-8"
    separator: str = ","
    has_header: bool = True
    table_column: str = "TABLE_NAME"
    skip_rows: int = 0
    max_rows: Optional[int] = None
    
    # Configurazioni specifiche del file
    custom_settings: Dict[str, Any] = field(default_factory=dict)
    
    def __post_init__(self):
        """Validazione della configurazione del file."""
        # Verifica che il file esista
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"File non trovato: {self.file_path}")
        
        # Validazione encoding
        valid_encodings = ['utf-8', 'latin-1', 'cp1252', 'ascii']
        if self.encoding not in valid_encodings:
            self.encoding = 'utf-8'
    
    @property
    def file_name(self) -> str:
        """Nome del file senza percorso."""
        return os.path.basename(self.file_path)
    
    @property
    def file_size_mb(self) -> float:
        """Dimensione del file in MB."""
        try:
            return os.path.getsize(self.file_path) / (1024 * 1024)
        except OSError:
            return 0.0


@dataclass 
class OutputConfig:
    """Configurazione per l'output dell'elaborazione."""
    
    base_filename: str
    output_directory: str
    include_timestamp: bool = True
    create_compressed_archive: bool = True
    keep_individual_files: bool = False
    
    # Formati di output
    excel_format: bool = True
    csv_format: bool = False
    
    def __post_init__(self):
        """Preparazione della configurazione output."""
        # Assicura che la directory esista
        os.makedirs(self.output_directory, exist_ok=True)
    
    def get_output_filename(self, part_number: Optional[int] = None, 
                          extension: str = "xlsx") -> str:
        """Genera il nome del file di output."""
        timestamp = ""
        if self.include_timestamp:
            from datetime import datetime
            timestamp = f"_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
        part_suffix = ""
        if part_number is not None:
            part_suffix = f"_part{part_number:03d}"
        
        filename = f"{self.base_filename}{timestamp}{part_suffix}.{extension}"
        return os.path.join(self.output_directory, filename)
    
    def get_archive_filename(self) -> str:
        """Genera il nome del file archivio."""
        return self.get_output_filename(extension="7z")
