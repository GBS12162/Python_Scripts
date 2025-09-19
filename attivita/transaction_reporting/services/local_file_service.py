"""
Servizio per la gestione di file locali e da NAS.
Sostituisce SharePointService per lettura diretta da filesystem.
"""

import logging
import shutil
from pathlib import Path
from datetime import datetime
from typing import Optional, List, Dict, Any


class LocalFileService:
    """Servizio per la gestione di file da percorsi locali e NAS."""
    
    def __init__(self, base_path: str = None):
        """
        Inizializza il servizio per file locali.
        
        Args:
            base_path: Directory base dove cercare i file (opzionale)
        """
        self.base_path = Path(base_path) if base_path else None
        self.logger = logging.getLogger(__name__)
        
    @classmethod
    def create_for_path(cls, file_path: str):
        """
        Factory method per creare servizio per un path specifico.
        
        Args:
            file_path: Path completo al file o directory
            
        Returns:
            LocalFileService configurato
        """
        path_obj = Path(file_path)
        if path_obj.is_file() or not path_obj.exists():
            # È un file o un path che non esiste ancora
            base_path = str(path_obj.parent)
        else:
            # È una directory
            base_path = str(path_obj)
            
        return cls(base_path=base_path)
    
    def check_access(self) -> bool:
        """
        Verifica l'accesso alla directory base.
        
        Returns:
            True se la directory è accessibile
        """
        try:
            if not self.base_path:
                self.logger.warning("Nessuna directory base specificata")
                return True  # Lasceremo che fallisca sui singoli file
                
            if not self.base_path.exists():
                self.logger.error(f"Directory non trovata: {self.base_path}")
                return False
                
            if not self.base_path.is_dir():
                self.logger.error(f"Il path non è una directory: {self.base_path}")
                return False
                
            # Test di lettura
            try:
                list(self.base_path.iterdir())
                self.logger.info(f"✅ Accesso verificato: {self.base_path}")
                return True
            except PermissionError:
                self.logger.error(f"Permessi insufficienti per: {self.base_path}")
                return False
                
        except Exception as e:
            self.logger.error(f"Errore verifica accesso: {str(e)}")
            return False
    
    def find_file(self, file_name: str, search_subdirs: bool = True) -> Optional[str]:
        """
        Cerca un file nella directory base.
        
        Args:
            file_name: Nome del file da cercare
            search_subdirs: Se True, cerca anche nelle sottodirectory
            
        Returns:
            Path completo del file se trovato, None altrimenti
        """
        try:
            if not self.base_path:
                self.logger.error("Nessuna directory base per la ricerca")
                return None
                
            self.logger.info(f"Cerco file: {file_name} in {self.base_path}")
            
            # Cerca il file esatto
            exact_match = self.base_path / file_name
            if exact_match.exists() and exact_match.is_file():
                self.logger.info(f"✅ File trovato: {exact_match}")
                return str(exact_match)
            
            # Cerca con pattern simili (case-insensitive)
            for file_path in self.base_path.glob('*'):
                if file_path.is_file() and file_path.name.lower() == file_name.lower():
                    self.logger.info(f"✅ File trovato (case-insensitive): {file_path}")
                    return str(file_path)
            
            # Se abilitato, cerca nelle sottodirectory
            if search_subdirs:
                self.logger.info("Cerco nelle sottodirectory...")
                for file_path in self.base_path.rglob(file_name):
                    if file_path.is_file():
                        self.logger.info(f"✅ File trovato in sottodirectory: {file_path}")
                        return str(file_path)
                        
                # Cerca con pattern case-insensitive nelle sottodirectory
                for file_path in self.base_path.rglob('*'):
                    if file_path.is_file() and file_path.name.lower() == file_name.lower():
                        self.logger.info(f"✅ File trovato in sottodirectory (case-insensitive): {file_path}")
                        return str(file_path)
            
            self.logger.warning(f"File non trovato: {file_name}")
            return None
            
        except Exception as e:
            self.logger.error(f"Errore ricerca file: {str(e)}")
            return None
    
    def copy_file(self, source_path: str, destination_path: str) -> bool:
        """
        Copia un file verso una destinazione.
        
        Args:
            source_path: Path del file sorgente
            destination_path: Path di destinazione
            
        Returns:
            True se la copia è riuscita
        """
        try:
            source = Path(source_path)
            destination = Path(destination_path)
            
            if not source.exists():
                self.logger.error(f"File sorgente non trovato: {source}")
                return False
                
            # Crea directory di destinazione se non esiste
            destination.parent.mkdir(parents=True, exist_ok=True)
            
            # Copia il file
            shutil.copy2(source, destination)
            
            # Verifica che la copia sia riuscita
            if destination.exists() and destination.stat().st_size > 0:
                self.logger.info(f"✅ File copiato: {source} → {destination}")
                return True
            else:
                self.logger.error("File copiato ma risulta vuoto o non trovato")
                return False
                
        except Exception as e:
            self.logger.error(f"Errore copia file: {str(e)}")
            return False
    
    def get_file_info(self, file_path: str) -> Optional[Dict[str, Any]]:
        """
        Ottiene informazioni su un file.
        
        Args:
            file_path: Path del file
            
        Returns:
            Dizionario con informazioni del file
        """
        try:
            path = Path(file_path)
            
            if not path.exists():
                return None
                
            stat = path.stat()
            
            return {
                'name': path.name,
                'size': stat.st_size,
                'modified': datetime.fromtimestamp(stat.st_mtime).isoformat(),
                'path': str(path.absolute()),
                'is_file': path.is_file(),
                'extension': path.suffix.lower()
            }
            
        except Exception as e:
            self.logger.error(f"Errore lettura info file: {str(e)}")
            return None
    
    def list_excel_files(self, directory: str = None) -> List[Dict[str, Any]]:
        """
        Lista tutti i file Excel in una directory.
        
        Args:
            directory: Directory da scansionare (usa base_path se None)
            
        Returns:
            Lista di informazioni sui file Excel
        """
        try:
            scan_dir = Path(directory) if directory else self.base_path
            
            if not scan_dir or not scan_dir.exists():
                self.logger.error(f"Directory non valida: {scan_dir}")
                return []
                
            excel_files = []
            excel_extensions = {'.xlsx', '.xls', '.xlsm'}
            
            for file_path in scan_dir.glob('*'):
                if file_path.is_file() and file_path.suffix.lower() in excel_extensions:
                    file_info = self.get_file_info(str(file_path))
                    if file_info:
                        excel_files.append(file_info)
            
            self.logger.info(f"Trovati {len(excel_files)} file Excel in {scan_dir}")
            return excel_files
            
        except Exception as e:
            self.logger.error(f"Errore scansione file Excel: {str(e)}")
            return []
    
    def find_con412_file(self, month: str = None) -> Optional[str]:
        """
        Trova il file CON-412 per il mese specificato.
        
        Args:
            month: Mese nel formato stringa (es: "AUGUST", "AUG", "08")
            
        Returns:
            Path del file se trovato, None altrimenti
        """
        try:
            if not self.base_path:
                self.logger.error("Nessuna directory base per la ricerca CON-412")
                return None
                
            # Pattern di ricerca per CON-412
            patterns = ['CON-412*.xlsx', 'con-412*.xlsx', 'CON412*.xlsx']
            
            found_files = []
            
            for pattern in patterns:
                for file_path in self.base_path.glob(pattern):
                    if file_path.is_file():
                        found_files.append(file_path)
            
            if not found_files:
                self.logger.warning("Nessun file CON-412 trovato")
                return None
                
            # Se specificato un mese, filtra per quello
            if month:
                month_upper = month.upper()
                for file_path in found_files:
                    file_name_upper = file_path.name.upper()
                    if month_upper in file_name_upper:
                        self.logger.info(f"✅ File CON-412 trovato per {month}: {file_path}")
                        return str(file_path)
            
            # Se non trovato per mese specifico o mese non specificato, prendi il primo
            selected_file = found_files[0]
            self.logger.info(f"✅ File CON-412 trovato: {selected_file}")
            
            if len(found_files) > 1:
                self.logger.info(f"Trovati {len(found_files)} file CON-412, usando: {selected_file.name}")
                
            return str(selected_file)
            
        except Exception as e:
            self.logger.error(f"Errore ricerca file CON-412: {str(e)}")
            return None


class FilePathHelper:
    """Helper per gestione path e validazione."""
    
    @staticmethod
    def validate_path(path: str) -> bool:
        """Valida se un path è raggiungibile."""
        try:
            path_obj = Path(path)
            
            # Testa se è un path di rete
            if path.startswith(r'\\'):
                # Path UNC (rete)
                return FilePathHelper._test_network_path(path_obj)
            else:
                # Path locale
                return path_obj.exists()
                
        except Exception:
            return False
    
    @staticmethod
    def _test_network_path(path: Path) -> bool:
        """Testa l'accesso a un path di rete."""
        try:
            # Prova a listare il contenuto
            if path.exists():
                list(path.iterdir())
                return True
            return False
        except Exception:
            return False
    
    @staticmethod
    def normalize_path(path: str) -> str:
        """Normalizza un path per il sistema operativo corrente."""
        return str(Path(path).resolve())
    
    @staticmethod
    def get_available_drives() -> List[str]:
        """Ottiene la lista dei drive disponibili (Windows)."""
        import string
        drives = []
        
        for letter in string.ascii_uppercase:
            drive = f"{letter}:\\"
            if Path(drive).exists():
                drives.append(drive)
                
        return drives