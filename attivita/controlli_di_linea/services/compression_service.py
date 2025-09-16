"""
Servizio per la compressione di file e archivi.
Gestisce la compressione 7z con ottimizzazioni per grandi dataset.
"""

import os
import py7zr
from typing import List, Optional, Dict, Any, Callable
from pathlib import Path
import logging

from models.config import Config, OutputConfig
from utils.file_utils import get_file_size_mb, ensure_directory


class CompressionService:
    """Servizio per la compressione dei file."""
    
    def __init__(self, config: Config):
        """
        Inizializza il servizio di compressione.
        
        Args:
            config: Configurazione dell'applicazione
        """
        self.config = config
        self.logger = logging.getLogger(__name__)
    
    def compress_files(self, file_paths: List[str], 
                      output_config: OutputConfig,
                      progress_callback: Optional[Callable[[int, int], None]] = None) -> Optional[str]:
        """
        Comprime una lista di file in un archivio 7z.
        
        Args:
            file_paths: Lista dei file da comprimere
            output_config: Configurazione di output
            progress_callback: Callback per il progresso
            
        Returns:
            Percorso dell'archivio creato o None se errore
        """
        if not self.config.enable_7z_compression:
            self.logger.info("Compressione 7z disabilitata")
            return None
        
        if not file_paths:
            self.logger.warning("Nessun file da comprimere")
            return None
        
        try:
            # Filtra solo i file che esistono e superano la soglia
            files_to_compress = self._filter_files_for_compression(file_paths)
            
            if not files_to_compress:
                self.logger.info("Nessun file supera la soglia di compressione")
                return None
            
            # Genera nome archivio
            archive_path = output_config.get_archive_filename()
            
            self.logger.info(f"Inizio compressione di {len(files_to_compress)} file")
            self.logger.info(f"Archivio di destinazione: {archive_path}")
            
            # Crea l'archivio 7z
            archive_created = self._create_7z_archive(
                files_to_compress, 
                archive_path, 
                progress_callback
            )
            
            if archive_created:
                # Verifica e statistiche
                self._log_compression_stats(files_to_compress, archive_path)
                
                # Pulizia file originali se richiesto
                if not output_config.keep_individual_files:
                    self._cleanup_original_files(files_to_compress)
                
                return archive_path
            else:
                return None
                
        except Exception as e:
            self.logger.error(f"Errore nella compressione: {e}")
            return None
    
    def _filter_files_for_compression(self, file_paths: List[str]) -> List[str]:
        """
        Filtra i file che devono essere compressi.
        
        Args:
            file_paths: Lista dei file da valutare
            
        Returns:
            Lista dei file da comprimere
        """
        filtered_files = []
        threshold_mb = self.config.compression_threshold_mb
        
        for file_path in file_paths:
            if not os.path.exists(file_path):
                self.logger.warning(f"File non trovato: {file_path}")
                continue
            
            file_size_mb = get_file_size_mb(file_path)
            
            if file_size_mb >= threshold_mb:
                filtered_files.append(file_path)
                self.logger.debug(f"File selezionato per compressione: {file_path} ({file_size_mb:.2f}MB)")
            else:
                self.logger.debug(f"File sotto soglia: {file_path} ({file_size_mb:.2f}MB)")
        
        return filtered_files
    
    def _create_7z_archive(self, file_paths: List[str], 
                          archive_path: str,
                          progress_callback: Optional[Callable[[int, int], None]] = None) -> bool:
        """
        Crea un archivio 7z con i file specificati.
        
        Args:
            file_paths: Lista dei file da comprimere
            archive_path: Percorso dell'archivio di destinazione
            progress_callback: Callback per il progresso
            
        Returns:
            True se la creazione è riuscita
        """
        try:
            # Assicura che la directory di destinazione esista
            ensure_directory(os.path.dirname(archive_path))
            
            # Configurazione compressione
            compression_level = self.config.compression_level
            
            # Crea l'archivio
            with py7zr.SevenZipFile(archive_path, 'w', 
                                   compression=py7zr.FILTER_LZMA2,
                                   preset=compression_level) as archive:
                
                total_files = len(file_paths)
                
                for i, file_path in enumerate(file_paths):
                    try:
                        # Nome nel archivio (solo il nome del file, non il percorso completo)
                        archive_name = os.path.basename(file_path)
                        
                        # Aggiungi il file all'archivio
                        archive.write(file_path, archive_name)
                        
                        self.logger.debug(f"Aggiunto all'archivio: {archive_name}")
                        
                        # Aggiorna progresso
                        if progress_callback:
                            progress_callback(i + 1, total_files)
                            
                    except Exception as e:
                        self.logger.error(f"Errore nell'aggiunta di {file_path} all'archivio: {e}")
                        continue
            
            # Verifica che l'archivio sia stato creato
            if os.path.exists(archive_path) and os.path.getsize(archive_path) > 0:
                self.logger.info(f"Archivio 7z creato con successo: {archive_path}")
                return True
            else:
                self.logger.error("Archivio 7z non creato correttamente")
                return False
                
        except Exception as e:
            self.logger.error(f"Errore nella creazione dell'archivio 7z: {e}")
            return False
    
    def _log_compression_stats(self, original_files: List[str], archive_path: str):
        """
        Registra le statistiche di compressione.
        
        Args:
            original_files: Lista dei file originali
            archive_path: Percorso dell'archivio creato
        """
        try:
            # Dimensione totale file originali
            original_size_mb = sum(get_file_size_mb(f) for f in original_files if os.path.exists(f))
            
            # Dimensione archivio
            archive_size_mb = get_file_size_mb(archive_path)
            
            # Calcolo ratio di compressione
            if original_size_mb > 0:
                compression_ratio = (1 - archive_size_mb / original_size_mb) * 100
                
                self.logger.info(f"Statistiche compressione:")
                self.logger.info(f"  File originali: {len(original_files)}")
                self.logger.info(f"  Dimensione originale: {original_size_mb:.2f} MB")
                self.logger.info(f"  Dimensione archivio: {archive_size_mb:.2f} MB")
                self.logger.info(f"  Ratio compressione: {compression_ratio:.1f}%")
                self.logger.info(f"  Spazio risparmiato: {original_size_mb - archive_size_mb:.2f} MB")
            
        except Exception as e:
            self.logger.warning(f"Errore nel calcolo delle statistiche di compressione: {e}")
    
    def _cleanup_original_files(self, file_paths: List[str]):
        """
        Elimina i file originali dopo la compressione.
        
        Args:
            file_paths: Lista dei file da eliminare
        """
        cleaned_files = 0
        
        for file_path in file_paths:
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
                    cleaned_files += 1
                    self.logger.debug(f"Eliminato file originale: {file_path}")
            except Exception as e:
                self.logger.warning(f"Impossibile eliminare {file_path}: {e}")
        
        if cleaned_files > 0:
            self.logger.info(f"Eliminati {cleaned_files} file originali dopo compressione")
    
    def extract_archive(self, archive_path: str, 
                       destination_dir: str,
                       progress_callback: Optional[Callable[[int, int], None]] = None) -> bool:
        """
        Estrae un archivio 7z.
        
        Args:
            archive_path: Percorso dell'archivio
            destination_dir: Directory di destinazione
            progress_callback: Callback per il progresso
            
        Returns:
            True se l'estrazione è riuscita
        """
        try:
            if not os.path.exists(archive_path):
                self.logger.error(f"Archivio non trovato: {archive_path}")
                return False
            
            # Assicura che la directory di destinazione esista
            ensure_directory(destination_dir)
            
            self.logger.info(f"Estrazione archivio: {archive_path}")
            self.logger.info(f"Destinazione: {destination_dir}")
            
            with py7zr.SevenZipFile(archive_path, 'r') as archive:
                # Ottieni lista dei file nell'archivio
                file_list = archive.getnames()
                total_files = len(file_list)
                
                self.logger.info(f"File nell'archivio: {total_files}")
                
                # Estrai tutti i file
                archive.extractall(path=destination_dir)
                
                # Simula progresso (py7zr non fornisce callback dettagliati)
                if progress_callback:
                    progress_callback(total_files, total_files)
            
            self.logger.info("Estrazione completata con successo")
            return True
            
        except Exception as e:
            self.logger.error(f"Errore nell'estrazione dell'archivio: {e}")
            return False
    
    def get_archive_info(self, archive_path: str) -> Optional[Dict[str, Any]]:
        """
        Ottiene informazioni su un archivio 7z.
        
        Args:
            archive_path: Percorso dell'archivio
            
        Returns:
            Dizionario con informazioni sull'archivio
        """
        try:
            if not os.path.exists(archive_path):
                return None
            
            with py7zr.SevenZipFile(archive_path, 'r') as archive:
                file_list = archive.list()
                
                total_files = len(file_list)
                total_compressed_size = sum(f.compressed for f in file_list)
                total_uncompressed_size = sum(f.uncompressed for f in file_list)
                
                compression_ratio = 0
                if total_uncompressed_size > 0:
                    compression_ratio = (1 - total_compressed_size / total_uncompressed_size) * 100
                
                return {
                    'file_count': total_files,
                    'compressed_size_mb': total_compressed_size / (1024 * 1024),
                    'uncompressed_size_mb': total_uncompressed_size / (1024 * 1024),
                    'compression_ratio': compression_ratio,
                    'files': [f.filename for f in file_list]
                }
                
        except Exception as e:
            self.logger.error(f"Errore nell'ottenere informazioni archivio: {e}")
            return None
    
    def validate_archive(self, archive_path: str) -> bool:
        """
        Valida l'integrità di un archivio 7z.
        
        Args:
            archive_path: Percorso dell'archivio
            
        Returns:
            True se l'archivio è valido
        """
        try:
            if not os.path.exists(archive_path):
                return False
            
            with py7zr.SevenZipFile(archive_path, 'r') as archive:
                # Tenta di leggere la lista dei file
                file_list = archive.list()
                
                # Se riusciamo a leggere la lista, l'archivio è probabilmente valido
                return len(file_list) >= 0
                
        except Exception as e:
            self.logger.error(f"Archivio corrotto o non valido {archive_path}: {e}")
            return False
    
    def estimate_compression_ratio(self, file_paths: List[str]) -> float:
        """
        Stima il ratio di compressione per una lista di file.
        
        Args:
            file_paths: Lista dei file da valutare
            
        Returns:
            Ratio di compressione stimato (0-1)
        """
        # Per file Excel/CSV, stima conservativa
        # I dati testuali si comprimono tipicamente 60-80%
        
        total_size = 0
        text_size = 0
        
        for file_path in file_paths:
            if not os.path.exists(file_path):
                continue
            
            file_size = get_file_size_mb(file_path)
            total_size += file_size
            
            # Se è un file di testo/Excel, conta come comprimibile
            ext = Path(file_path).suffix.lower()
            if ext in ['.xlsx', '.csv', '.txt', '.xml']:
                text_size += file_size
        
        if total_size == 0:
            return 0.0
        
        # Stima: file di testo si comprimono ~70%, altri ~30%
        text_ratio = 0.70
        binary_ratio = 0.30
        
        text_proportion = text_size / total_size
        binary_proportion = (total_size - text_size) / total_size
        
        estimated_ratio = (text_proportion * text_ratio) + (binary_proportion * binary_ratio)
        
        return min(0.90, max(0.10, estimated_ratio))  # Limita tra 10% e 90%
    
    def compress_single_file(self, file_path: str, 
                           archive_path: Optional[str] = None) -> Optional[str]:
        """
        Comprime un singolo file.
        
        Args:
            file_path: Percorso del file da comprimere
            archive_path: Percorso dell'archivio (opzionale)
            
        Returns:
            Percorso dell'archivio creato
        """
        if not os.path.exists(file_path):
            self.logger.error(f"File non trovato: {file_path}")
            return None
        
        if archive_path is None:
            # Genera nome automatico
            base_name = Path(file_path).stem
            archive_path = str(Path(file_path).parent / f"{base_name}.7z")
        
        return self._create_7z_archive([file_path], archive_path)
