"""
Utility per la gestione di date e timestamp.
"""

from datetime import datetime, timedelta
from typing import Optional


def get_current_timestamp() -> str:
    """
    Ottiene il timestamp corrente in formato standard.
    
    Returns:
        Timestamp formattato come stringa
    """
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def format_timestamp(dt: Optional[datetime] = None, format_str: str = "%Y-%m-%d %H:%M:%S") -> str:
    """
    Formatta un timestamp.
    
    Args:
        dt: DateTime da formattare (None per ora corrente)
        format_str: Formato del timestamp
        
    Returns:
        Timestamp formattato
    """
    if dt is None:
        dt = datetime.now()
    
    return dt.strftime(format_str)


def format_duration(seconds: float) -> str:
    """
    Formatta una durata in secondi in formato leggibile.
    
    Args:
        seconds: Durata in secondi
        
    Returns:
        Durata formattata (es. "2m 30s")
    """
    if seconds < 60:
        return f"{seconds:.1f}s"
    elif seconds < 3600:
        minutes = int(seconds // 60)
        remaining_seconds = seconds % 60
        return f"{minutes}m {remaining_seconds:.0f}s"
    else:
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        return f"{hours}h {minutes}m"


def get_elapsed_time(start_time: datetime) -> str:
    """
    Calcola il tempo trascorso da un momento specifico.
    
    Args:
        start_time: Momento di inizio
        
    Returns:
        Tempo trascorso formattato
    """
    elapsed = datetime.now() - start_time
    return format_duration(elapsed.total_seconds())


def add_business_days(start_date: datetime, days: int) -> datetime:
    """
    Aggiunge giorni lavorativi a una data.
    
    Args:
        start_date: Data di inizio
        days: Numero di giorni lavorativi da aggiungere
        
    Returns:
        Data risultante
    """
    current_date = start_date
    days_added = 0
    
    while days_added < days:
        current_date += timedelta(days=1)
        # Se non è sabato (5) o domenica (6)
        if current_date.weekday() < 5:
            days_added += 1
    
    return current_date


def is_business_day(date: datetime) -> bool:
    """
    Verifica se una data è un giorno lavorativo.
    
    Args:
        date: Data da verificare
        
    Returns:
        True se è un giorno lavorativo
    """
    return date.weekday() < 5  # Lunedì=0, Domenica=6


def get_file_timestamp() -> str:
    """
    Ottiene un timestamp adatto per nomi di file.
    
    Returns:
        Timestamp per file
    """
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def parse_timestamp(timestamp_str: str, format_str: str = "%Y%m%d_%H%M%S") -> Optional[datetime]:
    """
    Converte una stringa timestamp in datetime.
    
    Args:
        timestamp_str: Stringa del timestamp
        format_str: Formato del timestamp
        
    Returns:
        DateTime parsato o None se errore
    """
    try:
        return datetime.strptime(timestamp_str, format_str)
    except ValueError:
        return None


def get_time_range_string(start_time: datetime, end_time: datetime) -> str:
    """
    Crea una stringa che rappresenta un intervallo di tempo.
    
    Args:
        start_time: Tempo di inizio
        end_time: Tempo di fine
        
    Returns:
        Stringa dell'intervallo
    """
    duration = end_time - start_time
    duration_str = format_duration(duration.total_seconds())
    
    return f"{format_timestamp(start_time)} - {format_timestamp(end_time)} ({duration_str})"
