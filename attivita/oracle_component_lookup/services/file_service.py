"""
Servizio per la gestione dei file di input e output.
Gestisce il caricamento, la validazione e il salvataggio dei file.
"""

import pandas as pd
from typing import List, Optional, Dict, Any, Tuple, Union
from pathlib import Path
import os
import gzip
import chardet
import logging
from datetime import datetime

from models.config import Config, FileConfig
from utils.file_utils import detect_encoding, get_file_size_mb, is_compressed


class FileService:
    """Servizio per la gestione dei file."""
    
    def __init__(self, config: Config):
        """
        Inizializza il servizio file.
        
        Args:
            config: Configurazione dell'applicazione
        """
        self.config = config
        self.logger = logging.getLogger(__name__)
    
    def load_dataframe(self, file_config: FileConfig) -> Optional[pd.DataFrame]:
        """
        Carica un DataFrame da un file.
        
        Args:
            file_config: Configurazione del file
            
        Returns:
            DataFrame caricato o None se errore
        """
        try:
            file_path = file_config.file_path
            
            if not os.path.exists(file_path):
                self.logger.error(f"File non trovato: {file_path}")
                return None
            
            self.logger.info(f"Caricamento file: {file_path}")
            
            # Determina se il file è compresso
            if is_compressed(file_path):
                return self._load_compressed_file(file_config)
            else:
                return self._load_regular_file(file_config)
                
        except Exception as e:
            self.logger.error(f"Errore nel caricamento del file {file_config.file_path}: {e}")
            return None
    
    def _load_regular_file(self, file_config: FileConfig) -> Optional[pd.DataFrame]:
        """
        Carica un file regolare (non compresso).
        
        Args:
            file_config: Configurazione del file
            
        Returns:
            DataFrame caricato
        """
        # Rileva automaticamente l'encoding se necessario
        encoding = file_config.encoding
        if encoding == 'auto':
            encoding = detect_encoding(file_config.file_path)
            self.logger.info(f"Encoding rilevato automaticamente: {encoding}")
        
        # Parametri di caricamento
        load_params = {
            'sep': file_config.separator,
            'encoding': encoding,
            'skiprows': file_config.skip_rows if file_config.skip_rows > 0 else None,
            'nrows': file_config.max_rows,
            'dtype': str,  # Carica tutto come stringa per evitare problemi di tipo
            'na_filter': False  # Non convertire stringhe vuote in NaN
        }
        
        if not file_config.has_header:
            load_params['header'] = None
        
        return pd.read_csv(file_config.file_path, **load_params)
    
    def _load_compressed_file(self, file_config: FileConfig) -> Optional[pd.DataFrame]:
        """
        Carica un file compresso.
        
        Args:
            file_config: Configurazione del file
            
        Returns:
            DataFrame caricato
        """
        file_path = file_config.file_path
        
        if file_path.endswith('.gz'):
            # File gzip
            with gzip.open(file_path, 'rt', encoding=file_config.encoding) as f:
                load_params = {
                    'sep': file_config.separator,
                    'skiprows': file_config.skip_rows if file_config.skip_rows > 0 else None,
                    'nrows': file_config.max_rows,
                    'dtype': str,
                    'na_filter': False
                }
                
                if not file_config.has_header:
                    load_params['header'] = None
                
                return pd.read_csv(f, **load_params)
        
        else:
            # Altri formati compressi potrebbero essere aggiunti qui
            self.logger.error(f"Formato compresso non supportato: {file_path}")
            return None
    
    def load_multiple_files(self, file_paths: List[str], 
                          table_column: str = "TABLE_NAME") -> Optional[pd.DataFrame]:
        """
        Carica e unisce più file.
        
        Args:
            file_paths: Lista dei percorsi dei file
            table_column: Nome della colonna contenente i nomi delle tabelle
            
        Returns:
            DataFrame unificato
        """
        dataframes = []
        
        for file_path in file_paths:
            self.logger.info(f"Caricamento file: {file_path}")
            
            # Crea configurazione automatica per il file
            file_config = self._create_auto_file_config(file_path, table_column)
            
            df = self.load_dataframe(file_config)
            if df is not None and not df.empty:
                # Aggiungi colonna sorgente
                df['_SOURCE_FILE'] = os.path.basename(file_path)
                dataframes.append(df)
            else:
                self.logger.warning(f"File vuoto o non caricabile: {file_path}")
        
        if not dataframes:
            self.logger.error("Nessun file caricato con successo")
            return None
        
        # Unisci tutti i DataFrame
        combined_df = pd.concat(dataframes, ignore_index=True, sort=False)
        
        self.logger.info(f"Caricati {len(dataframes)} file per un totale di {len(combined_df)} righe")
        
        return combined_df
    
    def _create_auto_file_config(self, file_path: str, table_column: str) -> FileConfig:
        """
        Crea una configurazione automatica per un file.
        
        Args:
            file_path: Percorso del file
            table_column: Nome della colonna delle tabelle
            
        Returns:
            Configurazione del file
        """
        # Rileva encoding automaticamente
        encoding = detect_encoding(file_path)
        
        # Rileva separatore (semplice euristica)
        separator = ','
        if file_path.lower().endswith('.tsv'):
            separator = '\t'
        elif file_path.lower().endswith('.txt'):
            separator = '\t'
        
        return FileConfig(
            file_path=file_path,
            encoding=encoding,
            separator=separator,
            table_column=table_column,
            has_header=True
        )
    
    def validate_file(self, file_config: FileConfig) -> Tuple[bool, List[str]]:
        """
        Valida un file prima del caricamento.
        
        Args:
            file_config: Configurazione del file
            
        Returns:
            Tupla (valido, lista_errori)
        """
        errors = []
        
        # Verifica esistenza file
        if not os.path.exists(file_config.file_path):
            errors.append(f"File non trovato: {file_config.file_path}")
            return False, errors
        
        # Verifica dimensione file
        file_size_mb = file_config.file_size_mb
        if file_size_mb > 1000:  # 1GB
            errors.append(f"File molto grande ({file_size_mb:.1f} MB), potrebbero esserci problemi di memoria")
        
        # Verifica se il file è leggibile
        try:
            with open(file_config.file_path, 'r', encoding=file_config.encoding) as f:
                # Leggi solo le prime righe per test
                for i, line in enumerate(f):
                    if i >= 5:  # Testa solo le prime 5 righe
                        break
                    
                    # Verifica separatore
                    if file_config.separator not in line and i > 0:
                        errors.append(f"Separatore '{file_config.separator}' non trovato nella riga {i+1}")
                        break
                        
        except UnicodeDecodeError:
            errors.append(f"Impossibile decodificare il file con encoding '{file_config.encoding}'")
        except Exception as e:
            errors.append(f"Errore nella lettura del file: {e}")
        
        return len(errors) == 0, errors
    
    def estimate_rows(self, file_path: str) -> int:
        """
        Stima il numero di righe in un file.
        
        Args:
            file_path: Percorso del file
            
        Returns:
            Numero stimato di righe
        """
        try:
            if is_compressed(file_path):
                return self._estimate_rows_compressed(file_path)
            else:
                return self._estimate_rows_regular(file_path)
        except Exception as e:
            self.logger.warning(f"Impossibile stimare le righe per {file_path}: {e}")
            return 0
    
    def _estimate_rows_regular(self, file_path: str) -> int:
        """Stima righe per file regolare."""
        with open(file_path, 'rb') as f:
            # Leggi un campione
            sample_size = min(1024 * 1024, os.path.getsize(file_path))  # 1MB o dimensione file
            sample = f.read(sample_size)
            
            # Conta le righe nel campione
            sample_lines = sample.count(b'\n')
            
            if sample_lines == 0:
                return 0
            
            # Estrapola per l'intero file
            file_size = os.path.getsize(file_path)
            estimated_rows = int((sample_lines * file_size) / sample_size)
            
            return max(0, estimated_rows - 1)  # -1 per l'header
    
    def _estimate_rows_compressed(self, file_path: str) -> int:
        """Stima righe per file compresso."""
        if file_path.endswith('.gz'):
            with gzip.open(file_path, 'rb') as f:
                # Per file compressi, contiamo direttamente (meno efficiente ma più accurato)
                lines = 0
                chunk_size = 8192
                
                while True:
                    chunk = f.read(chunk_size)
                    if not chunk:
                        break
                    lines += chunk.count(b'\n')
                
                return max(0, lines - 1)  # -1 per l'header
        
        return 0
    
    def create_output_directory(self, base_path: str = None) -> str:
        """
        Crea la directory di output.
        
        Args:
            base_path: Percorso base (opzionale)
            
        Returns:
            Percorso della directory creata
        """
        if base_path:
            output_dir = base_path
        else:
            output_dir = self.config.output_directory
        
        # Aggiungi timestamp se necessario
        if not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
            self.logger.info(f"Creata directory di output: {output_dir}")
        
        return output_dir
    
    def cleanup_old_files(self, directory: str, max_age_days: int = 30):
        """
        Pulisce i file vecchi dalla directory.
        
        Args:
            directory: Directory da pulire
            max_age_days: Età massima dei file in giorni
        """
        if not os.path.exists(directory):
            return
        
        cutoff_time = datetime.now().timestamp() - (max_age_days * 24 * 3600)
        
        cleaned_files = 0
        for file_path in Path(directory).glob('*'):
            if file_path.is_file():
                file_time = file_path.stat().st_mtime
                if file_time < cutoff_time:
                    try:
                        file_path.unlink()
                        cleaned_files += 1
                    except Exception as e:
                        self.logger.warning(f"Impossibile eliminare {file_path}: {e}")
        
        if cleaned_files > 0:
            self.logger.info(f"Eliminati {cleaned_files} file vecchi da {directory}")
    
    def get_file_info(self, file_path: str) -> Dict[str, Any]:
        """
        Ottiene informazioni dettagliate su un file.
        
        Args:
            file_path: Percorso del file
            
        Returns:
            Dizionario con informazioni sul file
        """
        if not os.path.exists(file_path):
            return {"exists": False}
        
        stat = os.stat(file_path)
        
        info = {
            "exists": True,
            "size_bytes": stat.st_size,
            "size_mb": stat.st_size / (1024 * 1024),
            "modified_time": datetime.fromtimestamp(stat.st_mtime),
            "is_compressed": is_compressed(file_path),
            "estimated_rows": self.estimate_rows(file_path)
        }
        
        # Tenta di rilevare l'encoding
        try:
            info["detected_encoding"] = detect_encoding(file_path)
        except Exception:
            info["detected_encoding"] = "unknown"
        
        return info
