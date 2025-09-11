"""
Interfaccia console per l'applicazione Oracle Component Lookup.
Gestisce l'interazione con l'utente tramite console.
"""

import os
import sys
from typing import List, Optional, Dict, Any, Callable
from pathlib import Path
import logging

from models.config import Config, FileConfig, OutputConfig
from utils.progress_utils import ProgressTracker, ProgressInfo
from utils.date_utils import format_duration


class ConsoleUI:
    """Interfaccia utente console."""
    
    def __init__(self):
        """Inizializza l'interfaccia console."""
        self.logger = logging.getLogger(__name__)
        self._setup_console()
    
    def _setup_console(self):
        """Configura la console per l'output ottimale."""
        # Configura encoding per Windows
        if os.name == 'nt':
            try:
                # Imposta UTF-8 per la console Windows
                os.system('chcp 65001 > nul')
            except:
                pass
    
    def print_header(self):
        """Stampa l'header dell'applicazione."""
        print("=" * 60)
        print("    ORACLE COMPONENT LOOKUP - CONTROLLI DI LINEA")
        print("=" * 60)
        print()
    
    def print_separator(self, char: str = "-", length: int = 60):
        """
        Stampa un separatore.
        
        Args:
            char: Carattere del separatore
            length: Lunghezza del separatore
        """
        print(char * length)
    
    def print_info(self, message: str, prefix: str = "[INFO]"):
        """
        Stampa un messaggio informativo.
        
        Args:
            message: Messaggio da stampare
            prefix: Prefisso del messaggio
        """
        print(f"{prefix} {message}")
    
    def print_success(self, message: str):
        """Stampa un messaggio di successo."""
        self.print_info(message, "[✓]")
    
    def print_warning(self, message: str):
        """Stampa un messaggio di avviso."""
        self.print_info(message, "[⚠]")
    
    def print_error(self, message: str):
        """Stampa un messaggio di errore."""
        self.print_info(message, "[✗]")
    
    def ask_input(self, prompt: str, default: Optional[str] = None) -> str:
        """
        Richiede input dall'utente.
        
        Args:
            prompt: Prompt da mostrare
            default: Valore di default
            
        Returns:
            Input dell'utente
        """
        if default:
            full_prompt = f"{prompt} [{default}]: "
        else:
            full_prompt = f"{prompt}: "
        
        try:
            user_input = input(full_prompt).strip()
            return user_input if user_input else (default or "")
        except (KeyboardInterrupt, EOFError):
            print("\nOperazione annullata dall'utente.")
            sys.exit(0)
    
    def ask_yes_no(self, question: str, default: bool = True) -> bool:
        """
        Richiede una risposta sì/no.
        
        Args:
            question: Domanda da porre
            default: Risposta di default
            
        Returns:
            True per sì, False per no
        """
        default_str = "S/n" if default else "s/N"
        response = self.ask_input(f"{question} ({default_str})", 
                                "s" if default else "n").lower()
        
        if response in ['s', 'si', 'sì', 'y', 'yes']:
            return True
        elif response in ['n', 'no']:
            return False
        else:
            return default
    
    def ask_choice(self, question: str, choices: List[str], 
                  default: Optional[int] = None) -> int:
        """
        Richiede una scelta da una lista.
        
        Args:
            question: Domanda da porre
            choices: Lista delle scelte
            default: Indice della scelta di default
            
        Returns:
            Indice della scelta selezionata
        """
        print(f"\n{question}")
        
        for i, choice in enumerate(choices):
            marker = " (default)" if default == i else ""
            print(f"  {i + 1}. {choice}{marker}")
        
        while True:
            try:
                response = self.ask_input("Scelta", 
                                        str(default + 1) if default is not None else None)
                
                if not response and default is not None:
                    return default
                
                choice_num = int(response) - 1
                
                if 0 <= choice_num < len(choices):
                    return choice_num
                else:
                    self.print_error(f"Scelta non valida. Inserire un numero da 1 a {len(choices)}")
                    
            except ValueError:
                self.print_error("Inserire un numero valido")
    
    def ask_file_path(self, prompt: str, must_exist: bool = True) -> str:
        """
        Richiede un percorso di file.
        
        Args:
            prompt: Prompt da mostrare
            must_exist: True se il file deve esistere
            
        Returns:
            Percorso del file
        """
        while True:
            file_path = self.ask_input(prompt)
            
            if not file_path:
                self.print_error("Percorso file richiesto")
                continue
            
            # Espandi percorsi relativi
            file_path = os.path.abspath(file_path)
            
            if must_exist and not os.path.exists(file_path):
                self.print_error(f"File non trovato: {file_path}")
                continue
            
            return file_path
    
    def ask_multiple_files(self, prompt: str) -> List[str]:
        """
        Richiede più file (uno per riga, vuoto per terminare).
        
        Args:
            prompt: Prompt da mostrare
            
        Returns:
            Lista dei percorsi dei file
        """
        files = []
        print(f"\n{prompt}")
        print("(Inserire un file per riga, riga vuota per terminare)")
        
        while True:
            file_path = input(f"File #{len(files) + 1}: ").strip()
            
            if not file_path:
                break
            
            file_path = os.path.abspath(file_path)
            
            if not os.path.exists(file_path):
                self.print_warning(f"File non trovato: {file_path}")
                continue
            
            files.append(file_path)
            self.print_info(f"Aggiunto: {os.path.basename(file_path)}")
        
        return files
    
    def show_file_info(self, file_path: str, info: Dict[str, Any]):
        """
        Mostra informazioni su un file.
        
        Args:
            file_path: Percorso del file
            info: Informazioni sul file
        """
        print(f"\nInformazioni file: {os.path.basename(file_path)}")
        print(f"  Percorso: {file_path}")
        print(f"  Dimensione: {info.get('size_mb', 0):.2f} MB")
        print(f"  Righe stimate: {info.get('estimated_rows', 'N/A')}")
        print(f"  Encoding rilevato: {info.get('detected_encoding', 'N/A')}")
        print(f"  Compresso: {'Sì' if info.get('is_compressed', False) else 'No'}")
    
    def show_processing_stats(self, stats_info: Dict[str, Any]):
        """
        Mostra le statistiche di elaborazione.
        
        Args:
            stats_info: Informazioni statistiche
        """
        elaborazione = stats_info.get('elaborazione', {})
        
        print("\n" + "=" * 50)
        print("           STATISTICHE ELABORAZIONE")
        print("=" * 50)
        
        print(f"Tabelle elaborate:     {elaborazione.get('tabelle_totali', 0):,}")
        print(f"Tabelle trovate:       {elaborazione.get('tabelle_trovate', 0):,}")
        print(f"Tabelle non trovate:   {elaborazione.get('tabelle_non_trovate', 0):,}")
        print(f"Errori:                {elaborazione.get('errori', 0):,}")
        print(f"Percentuale successo:  {elaborazione.get('percentuale_successo', 0):.1f}%")
        
        durata = elaborazione.get('durata_secondi')
        if durata:
            print(f"Tempo elaborazione:    {format_duration(durata)}")
            
            tabelle_totali = elaborazione.get('tabelle_totali', 0)
            if tabelle_totali > 0 and durata > 0:
                velocita = tabelle_totali / durata
                print(f"Velocità media:        {velocita:.1f} tabelle/secondo")
        
        print("=" * 50)
    
    def create_progress_callback(self, description: str = "Elaborazione") -> Callable[[int, int], None]:
        """
        Crea un callback per mostrare il progresso.
        
        Args:
            description: Descrizione dell'operazione
            
        Returns:
            Funzione callback
        """
        tracker = ProgressTracker(100, description)  # Verrà aggiornato con il totale reale
        
        def progress_callback(current: int, total: int):
            if tracker.total != total:
                tracker.total = total
                tracker.reset(total)
            
            tracker.set_current(current, force_update=True)
            
            # Stampa progresso sulla stessa riga
            status = tracker.get_status_string()
            print(f"\r{status}", end="", flush=True)
            
            # Se completato, vai a capo
            if current >= total:
                print()  # Nuova riga alla fine
        
        return progress_callback
    
    def wait_for_enter(self, message: str = "Premere INVIO per continuare..."):
        """
        Aspetta che l'utente prema INVIO.
        
        Args:
            message: Messaggio da mostrare
        """
        try:
            input(f"\n{message}")
        except (KeyboardInterrupt, EOFError):
            print("\nOperazione annullata.")
            sys.exit(0)
    
    def clear_screen(self):
        """Pulisce lo schermo."""
        os.system('cls' if os.name == 'nt' else 'clear')
    
    def show_config_summary(self, config: Config):
        """
        Mostra un riassunto della configurazione.
        
        Args:
            config: Configurazione da mostrare
        """
        print("\n" + "-" * 40)
        print("        CONFIGURAZIONE ATTUALE")
        print("-" * 40)
        
        print(f"File componenti:       {config.components_file}")
        print(f"Directory output:      {config.output_directory}")
        print(f"Righe max per file:    {config.max_rows_per_file:,}")
        print(f"Motore Excel:          {config.excel_engine}")
        print(f"Compressione 7z:       {'Abilitata' if config.enable_7z_compression else 'Disabilitata'}")
        print(f"Elaborazione parallela: {'Abilitata' if config.enable_multiprocessing else 'Disabilitata'}")
        
        if config.enable_multiprocessing:
            print(f"Worker paralleli:      {config.max_workers}")
            print(f"Dimensione chunk:      {config.chunk_size:,}")
        
        print("-" * 40)
    
    def show_warning_list(self, warnings: List[str], title: str = "Avvisi"):
        """
        Mostra una lista di avvisi.
        
        Args:
            warnings: Lista degli avvisi
            title: Titolo della sezione
        """
        if warnings:
            print(f"\n{title}:")
            for warning in warnings:
                self.print_warning(warning)
    
    def confirm_operation(self, operation: str, details: List[str] = None) -> bool:
        """
        Richiede conferma per un'operazione.
        
        Args:
            operation: Descrizione dell'operazione
            details: Dettagli aggiuntivi
            
        Returns:
            True se confermato
        """
        print(f"\nStai per eseguire: {operation}")
        
        if details:
            print("\nDettagli:")
            for detail in details:
                print(f"  • {detail}")
        
        return self.ask_yes_no("\nConfermi l'operazione?", default=False)
    
    def show_completion_message(self, operation: str, results: List[str] = None):
        """
        Mostra un messaggio di completamento.
        
        Args:
            operation: Operazione completata
            results: Lista dei risultati
        """
        print(f"\n✓ {operation} completata con successo!")
        
        if results:
            print("\nFile generati:")
            for result in results:
                print(f"  • {os.path.basename(result)}")
                print(f"    {result}")
        
        print()
    
    def handle_error(self, error: Exception, context: str = ""):
        """
        Gestisce la visualizzazione di un errore.
        
        Args:
            error: Eccezione verificatasi
            context: Contesto dell'errore
        """
        error_msg = f"Errore"
        if context:
            error_msg += f" in {context}"
        error_msg += f": {str(error)}"
        
        self.print_error(error_msg)
        
        # In modalità debug, mostra anche lo stack trace
        if self.logger.isEnabledFor(logging.DEBUG):
            import traceback
            print("\nStack trace:")
            traceback.print_exc()
