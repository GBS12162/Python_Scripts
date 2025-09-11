"""
Utility per la gestione dei file.
"""

import os
import chardet
from pathlib import Path
from typing import Optional


def detect_encoding(file_path: str, sample_size: int = 10000) -> str:
    """
    Rileva automaticamente l'encoding di un file.
    
    Args:
        file_path: Percorso del file
        sample_size: Dimensione del campione da analizzare
        
    Returns:
        Nome dell'encoding rilevato
    """
    try:
        with open(file_path, 'rb') as f:
            # Leggi un campione del file
            sample = f.read(sample_size)
            
            # Rileva l'encoding
            result = chardet.detect(sample)
            
            if result and result['encoding']:
                encoding = result['encoding'].lower()
                
                # Normalizza alcuni encoding comuni
                encoding_map = {
                    'windows-1252': 'cp1252',
                    'iso-8859-1': 'latin-1',
                    'ascii': 'utf-8'  # ASCII è compatibile con UTF-8
                }
                
                return encoding_map.get(encoding, encoding)
            
    except Exception:
        pass
    
    # Default fallback
    return 'utf-8'


def get_file_size_mb(file_path: str) -> float:
    """
    Ottiene la dimensione di un file in MB.
    
    Args:
        file_path: Percorso del file
        
    Returns:
        Dimensione in MB
    """
    try:
        size_bytes = os.path.getsize(file_path)
        return size_bytes / (1024 * 1024)
    except OSError:
        return 0.0


def is_compressed(file_path: str) -> bool:
    """
    Verifica se un file è compresso.
    
    Args:
        file_path: Percorso del file
        
    Returns:
        True se il file è compresso
    """
    compressed_extensions = {'.gz', '.bz2', '.xz', '.zip', '.7z'}
    
    file_extension = Path(file_path).suffix.lower()
    return file_extension in compressed_extensions


def ensure_directory(directory_path: str) -> str:
    """
    Assicura che una directory esista, creandola se necessario.
    
    Args:
        directory_path: Percorso della directory
        
    Returns:
        Percorso della directory (normalizzato)
    """
    path = Path(directory_path)
    path.mkdir(parents=True, exist_ok=True)
    return str(path.absolute())


def get_safe_filename(filename: str) -> str:
    """
    Converte un nome file in una versione sicura per il filesystem.
    
    Args:
        filename: Nome del file originale
        
    Returns:
        Nome del file sicuro
    """
    # Caratteri non validi per Windows
    invalid_chars = '<>:"/\\|?*'
    
    safe_name = filename
    for char in invalid_chars:
        safe_name = safe_name.replace(char, '_')
    
    # Rimuovi spazi multipli e normalizza
    safe_name = ' '.join(safe_name.split())
    
    # Limita la lunghezza
    if len(safe_name) > 200:
        name_part, ext_part = os.path.splitext(safe_name)
        safe_name = name_part[:200-len(ext_part)] + ext_part
    
    return safe_name


def copy_file_with_progress(source: str, destination: str, 
                          progress_callback: Optional[callable] = None) -> bool:
    """
    Copia un file con callback di progresso.
    
    Args:
        source: File sorgente
        destination: File destinazione
        progress_callback: Funzione di callback per il progresso
        
    Returns:
        True se la copia è riuscita
    """
    try:
        source_size = os.path.getsize(source)
        copied_size = 0
        chunk_size = 1024 * 1024  # 1MB
        
        with open(source, 'rb') as src, open(destination, 'wb') as dst:
            while True:
                chunk = src.read(chunk_size)
                if not chunk:
                    break
                
                dst.write(chunk)
                copied_size += len(chunk)
                
                if progress_callback:
                    progress = (copied_size / source_size) * 100
                    progress_callback(progress)
        
        return True
        
    except Exception:
        return False


def get_available_disk_space(path: str) -> int:
    """
    Ottiene lo spazio libero disponibile su disco.
    
    Args:
        path: Percorso per verificare lo spazio
        
    Returns:
        Spazio disponibile in bytes
    """
    try:
        if os.name == 'nt':  # Windows
            import ctypes
            free_bytes = ctypes.c_ulonglong(0)
            ctypes.windll.kernel32.GetDiskFreeSpaceExW(
                ctypes.c_wchar_p(path),
                ctypes.pointer(free_bytes),
                None,
                None
            )
            return free_bytes.value
        else:  # Unix/Linux
            statvfs = os.statvfs(path)
            return statvfs.f_frsize * statvfs.f_bavail
    except Exception:
        return 0


def cleanup_temp_files(directory: str, pattern: str = "*.tmp"):
    """
    Pulisce i file temporanei da una directory.
    
    Args:
        directory: Directory da pulire
        pattern: Pattern dei file da eliminare
    """
    try:
        temp_files = Path(directory).glob(pattern)
        
        for temp_file in temp_files:
            try:
                temp_file.unlink()
            except Exception:
                pass  # Ignora errori su singoli file
                
    except Exception:
        pass  # Ignora errori generali


def normalize_path(path: str) -> str:
    """
    Normalizza un percorso per la compatibilità cross-platform.
    
    Args:
        path: Percorso da normalizzare
        
    Returns:
        Percorso normalizzato
    """
    return str(Path(path).resolve())
