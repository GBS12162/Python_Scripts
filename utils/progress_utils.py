"""
Utility per il tracking del progresso delle operazioni.
"""

import time
from typing import Optional, Callable, Any
from datetime import datetime, timedelta
from dataclasses import dataclass


@dataclass
class ProgressInfo:
    """Informazioni sul progresso di un'operazione."""
    current: int
    total: int
    start_time: datetime
    current_time: datetime
    
    @property
    def percentage(self) -> float:
        """Percentuale di completamento."""
        if self.total <= 0:
            return 0.0
        return (self.current / self.total) * 100
    
    @property
    def elapsed_time(self) -> timedelta:
        """Tempo trascorso."""
        return self.current_time - self.start_time
    
    @property
    def estimated_total_time(self) -> Optional[timedelta]:
        """Tempo totale stimato."""
        if self.current <= 0:
            return None
        
        elapsed = self.elapsed_time.total_seconds()
        rate = self.current / elapsed
        
        if rate <= 0:
            return None
        
        total_seconds = self.total / rate
        return timedelta(seconds=total_seconds)
    
    @property
    def estimated_remaining_time(self) -> Optional[timedelta]:
        """Tempo rimanente stimato."""
        total_time = self.estimated_total_time
        if total_time is None:
            return None
        
        return total_time - self.elapsed_time
    
    @property
    def rate_per_second(self) -> float:
        """Velocità di elaborazione (elementi per secondo)."""
        elapsed_seconds = self.elapsed_time.total_seconds()
        if elapsed_seconds <= 0:
            return 0.0
        
        return self.current / elapsed_seconds


class ProgressTracker:
    """Tracker per il progresso delle operazioni."""
    
    def __init__(self, total: int, description: str = "Elaborazione"):
        """
        Inizializza il tracker.
        
        Args:
            total: Numero totale di elementi da elaborare
            description: Descrizione dell'operazione
        """
        self.total = total
        self.description = description
        self.current = 0
        self.start_time = datetime.now()
        self.last_update_time = self.start_time
        self.update_interval = 1.0  # Secondi tra aggiornamenti
        
        # Callback per aggiornamenti
        self.progress_callback: Optional[Callable[[ProgressInfo], None]] = None
        self.completion_callback: Optional[Callable[[ProgressInfo], None]] = None
    
    def set_progress_callback(self, callback: Callable[[ProgressInfo], None]):
        """
        Imposta il callback per gli aggiornamenti di progresso.
        
        Args:
            callback: Funzione da chiamare per ogni aggiornamento
        """
        self.progress_callback = callback
    
    def set_completion_callback(self, callback: Callable[[ProgressInfo], None]):
        """
        Imposta il callback per il completamento.
        
        Args:
            callback: Funzione da chiamare al completamento
        """
        self.completion_callback = callback
    
    def update(self, increment: int = 1, force_update: bool = False):
        """
        Aggiorna il progresso.
        
        Args:
            increment: Incremento da aggiungere al progresso
            force_update: Forza l'aggiornamento anche se non è passato abbastanza tempo
        """
        self.current = min(self.current + increment, self.total)
        current_time = datetime.now()
        
        # Verifica se è tempo di fare un aggiornamento
        time_since_update = (current_time - self.last_update_time).total_seconds()
        
        if force_update or time_since_update >= self.update_interval or self.current >= self.total:
            self.last_update_time = current_time
            
            progress_info = ProgressInfo(
                current=self.current,
                total=self.total,
                start_time=self.start_time,
                current_time=current_time
            )
            
            # Chiama il callback di progresso
            if self.progress_callback:
                self.progress_callback(progress_info)
            
            # Se completato, chiama il callback di completamento
            if self.current >= self.total and self.completion_callback:
                self.completion_callback(progress_info)
    
    def set_current(self, value: int, force_update: bool = False):
        """
        Imposta direttamente il valore corrente.
        
        Args:
            value: Nuovo valore corrente
            force_update: Forza l'aggiornamento
        """
        old_current = self.current
        self.current = min(max(value, 0), self.total)
        
        increment = self.current - old_current
        if increment != 0:
            self.update(0, force_update)  # 0 perché abbiamo già aggiornato current
    
    def reset(self, new_total: Optional[int] = None):
        """
        Resetta il tracker.
        
        Args:
            new_total: Nuovo totale (opzionale)
        """
        if new_total is not None:
            self.total = new_total
        
        self.current = 0
        self.start_time = datetime.now()
        self.last_update_time = self.start_time
    
    def get_progress_info(self) -> ProgressInfo:
        """
        Ottiene le informazioni di progresso correnti.
        
        Returns:
            Informazioni di progresso
        """
        return ProgressInfo(
            current=self.current,
            total=self.total,
            start_time=self.start_time,
            current_time=datetime.now()
        )
    
    def is_completed(self) -> bool:
        """
        Verifica se l'operazione è completata.
        
        Returns:
            True se completata
        """
        return self.current >= self.total
    
    def get_status_string(self) -> str:
        """
        Ottiene una stringa di stato leggibile.
        
        Returns:
            Stringa di stato
        """
        info = self.get_progress_info()
        
        status_parts = [
            f"{self.description}: {info.current}/{info.total}",
            f"({info.percentage:.1f}%)"
        ]
        
        if info.rate_per_second > 0:
            status_parts.append(f"{info.rate_per_second:.1f}/s")
        
        remaining_time = info.estimated_remaining_time
        if remaining_time and remaining_time.total_seconds() > 0:
            remaining_str = self._format_timedelta(remaining_time)
            status_parts.append(f"ETA: {remaining_str}")
        
        return " | ".join(status_parts)
    
    def _format_timedelta(self, td: timedelta) -> str:
        """
        Formatta un timedelta in modo leggibile.
        
        Args:
            td: Timedelta da formattare
            
        Returns:
            Stringa formattata
        """
        total_seconds = int(td.total_seconds())
        
        if total_seconds < 60:
            return f"{total_seconds}s"
        elif total_seconds < 3600:
            minutes = total_seconds // 60
            seconds = total_seconds % 60
            return f"{minutes}m {seconds}s"
        else:
            hours = total_seconds // 3600
            minutes = (total_seconds % 3600) // 60
            return f"{hours}h {minutes}m"


class MultiStageProgressTracker:
    """Tracker per operazioni multi-stage."""
    
    def __init__(self, stages: list):
        """
        Inizializza il tracker multi-stage.
        
        Args:
            stages: Lista di tuple (nome_stage, peso_relativo)
        """
        self.stages = stages
        self.current_stage = 0
        self.stage_trackers = {}
        self.overall_progress = 0.0
        
        # Calcola i pesi totali
        total_weight = sum(weight for _, weight in stages)
        self.stage_weights = [(name, weight / total_weight) for name, weight in stages]
        
        self.start_time = datetime.now()
        self.progress_callback: Optional[Callable[[str, float, str], None]] = None
    
    def set_progress_callback(self, callback: Callable[[str, float, str], None]):
        """
        Imposta il callback per gli aggiornamenti.
        
        Args:
            callback: Funzione(stage_name, overall_progress, status_message)
        """
        self.progress_callback = callback
    
    def start_stage(self, stage_index: int, total_items: int) -> ProgressTracker:
        """
        Inizia uno stage specifico.
        
        Args:
            stage_index: Indice dello stage
            total_items: Numero totale di elementi per questo stage
            
        Returns:
            ProgressTracker per questo stage
        """
        if stage_index >= len(self.stages):
            raise IndexError("Stage index out of range")
        
        self.current_stage = stage_index
        stage_name, _ = self.stages[stage_index]
        
        tracker = ProgressTracker(total_items, stage_name)
        tracker.set_progress_callback(self._on_stage_progress)
        
        self.stage_trackers[stage_index] = tracker
        
        return tracker
    
    def _on_stage_progress(self, progress_info: ProgressInfo):
        """Callback interno per aggiornamenti di stage."""
        # Calcola il progresso overall
        overall_progress = 0.0
        
        for i, (_, weight) in enumerate(self.stage_weights):
            if i in self.stage_trackers:
                stage_tracker = self.stage_trackers[i]
                stage_progress = stage_tracker.current / stage_tracker.total if stage_tracker.total > 0 else 0
                overall_progress += stage_progress * weight
            elif i < self.current_stage:
                # Stage completati
                overall_progress += weight
        
        self.overall_progress = overall_progress
        
        # Chiama il callback se presente
        if self.progress_callback:
            current_stage_name = self.stages[self.current_stage][0]
            status = f"{current_stage_name}: {progress_info.current}/{progress_info.total}"
            
            self.progress_callback(current_stage_name, overall_progress * 100, status)
    
    def get_overall_progress(self) -> float:
        """
        Ottiene il progresso overall (0-100).
        
        Returns:
            Progresso percentuale overall
        """
        return self.overall_progress * 100
