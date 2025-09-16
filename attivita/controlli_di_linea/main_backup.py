"""
Oracle Component Lookup - Applicazione Principale
Versione modulare con architettura ottimizzata per performance e manutenibilità.

Autore: Sistema Automatizzato
Data: 2025
"""

import sys
import os
import logging
from pathlib import Path
from typing import List, Optional
import traceback
from datetime import datetime

# Aggiungi il percorso corrente al PYTHONPATH
sys.path.insert(0, str(Path(__file__).parent))

# Import dei moduli dell'applicazione
from models.config import Config, FileConfig, OutputConfig
from services.component_service import ComponentService
from services.lookup_service import LookupService
from services.file_service import FileService
from services.excel_service import ExcelService
from services.compression_service import CompressionService
from ui.console_ui import ConsoleUI
from ui.menu_manager import MenuManager
from utils.date_utils import get_current_timestamp
from utils.file_utils import ensure_directory


class OracleComponentLookupApp:
    """Applicazione principale per il lookup dei componenti Oracle."""
    
    def __init__(self):
        """Inizializza l'applicazione."""
        self.config = Config()
        self.ui = ConsoleUI()
        self.menu_manager = MenuManager()
        
        # Servizi
        self.component_service = ComponentService(self.config)
        self.lookup_service = LookupService(self.config, self.component_service)
        self.file_service = FileService(self.config)
        self.excel_service = ExcelService(self.config)
        self.compression_service = CompressionService(self.config)
        
        # Stato dell'applicazione
        self.logger = self._setup_logging()
        self.is_initialized = False
        
    def _setup_logging(self) -> logging.Logger:
        """Configura il sistema di logging."""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('oracle_lookup.log', encoding='utf-8'),
                logging.StreamHandler(sys.stdout)
            ]
        )
        return logging.getLogger(__name__)
    
    def initialize(self) -> bool:
        """
        Inizializza l'applicazione caricando i componenti.
        
        Returns:
            True se l'inizializzazione è riuscita
        """
        try:
            self.ui.print_header()
            self.ui.print_info("Inizializzazione applicazione...")
            
            # Verifica file componenti
            if not os.path.exists(self.config.components_file):
                self.ui.print_error(f"File componenti non trovato: {self.config.components_file}")
                
                # Chiedi all'utente di specificare il percorso
                components_file = self.ui.ask_file_path(
                    "Inserisci il percorso del file Components.csv"
                )
                self.config.components_file = components_file
            
            # Carica i componenti
            self.ui.print_info("Caricamento componenti Oracle...")
            if not self.component_service.load_components():
                self.ui.print_error("Impossibile caricare i componenti")
                return False
            
            self.ui.print_success(f"Caricati {self.component_service.component_count} componenti")
            
            # Crea directory di output
            ensure_directory(self.config.output_directory)
            
            self.is_initialized = True
            return True
            
        except Exception as e:
            self.ui.handle_error(e, "inizializzazione")
            return False
    
    def run(self):
        """Avvia l'applicazione principale."""
        try:
            if not self.initialize():
                self.ui.print_error("Inizializzazione fallita")
                sys.exit(1)
            
            self._setup_main_menu()
            self._show_welcome_message()
            
            # Loop principale del menu (gestito internamente da MenuManager)
            try:
                self.menu_manager.show_menu()
            except KeyboardInterrupt:
                self.ui.print_info("\nOperazione interrotta dall'utente")
            except Exception as e:
                self.ui.handle_error(e, "menu principale")
            
        except Exception as e:
            self.ui.handle_error(e, "applicazione principale")
            sys.exit(1)
        finally:
            self._cleanup()
    
    def _setup_main_menu(self):
        """Configura il menu principale."""
        self.menu_manager.clear_menu()
        self.menu_manager.add_menu_item("📄 Elabora file Oracle (output Excel)", self._process_oracle_files)
    
    def _show_welcome_message(self):
        """Mostra il messaggio di benvenuto."""
        self.ui.print_separator()
        self.ui.print_success("Applicazione inizializzata correttamente")
        self.ui.show_config_summary(self.config)
        
        stats = self.component_service.get_statistics()
        self.ui.print_info(f"Componenti caricati: {stats['total_components']}")
        self.ui.print_info(f"Prefissi unici: {stats['unique_prefixes']}")
        
        if stats.get('components_with_office', 0) > 0:
            self.ui.print_info(f"Componenti con Ufficio IT: {stats['components_with_office']}")
    
    def _process_oracle_files(self):
        """Elabora file Oracle esattamente come l'originale."""
        last_used_path = getattr(self, '_last_used_path', None)
        
        while True:
            try:
                # Clear console e header come l'originale
                print("\n" + "=" * 60)
                print("📄 ELABORAZIONE FILE ORACLE")
                print("━" * 40)
                
                # Chiedi directory con opzione di riutilizzare l'ultimo path
                if last_used_path:
                    print(f"📁 Ultimo percorso utilizzato: {last_used_path}")
                    print("\n💡 Opzioni:")
                    print("   • Premi INVIO per usare l'ultimo percorso")
                    print("   • Inserisci un nuovo percorso")
                    print("   • Digita 'back' o 'b' per tornare al menu principale")
                else:
                    print("📁 Inserisci il percorso della directory contenente i file:")
                    print("   Esempi:")
                    print("   • C:\\temp")
                    print("   • \\\\nas1be\\Docgs\\folder") 
                    print("   • Z:\\logs")
                    print("\n💡 Opzioni:")
                    print("   • Inserisci il percorso completo")
                    print("   • Digita 'back' o 'b' per tornare al menu principale")

                user_input = input("\n➤ Percorso: ").strip().strip('"').strip("'")
                
                # Controlla se l'utente vuole tornare indietro
                if user_input.lower() in ['back', 'b', 'indietro']:
                    return  # Torna al menu principale
                
                # Se vuole usare l'ultimo path e non ha inserito nulla
                if not user_input and last_used_path:
                    user_input = last_used_path
                    print(f"✅ Utilizzo percorso precedente: {user_input}")
                
                if not user_input:
                    print("❌ Percorso non valido")
                    input("\n📝 Premi INVIO per riprovare...")
                    continue
                    
                dir_path = Path(user_input)
                if not dir_path.exists():
                    print(f"❌ Directory non trovata: {user_input}")
                    print("💡 Verifica il percorso e riprova")
                    input("\n📝 Premi INVIO per riprovare o digita CTRL+C per uscire...")
                    continue
                
                # Salva l'ultimo path utilizzato
                self._last_used_path = str(dir_path)
                
                # Lista file
                files = [f for f in dir_path.iterdir() if f.is_file()]
                
                if not files:
                    print("❌ Nessun file trovato nella directory")
                    input("\n📝 Premi INVIO per riprovare...")
                    continue
                
                # Chiedi se elaborare tutti i file o selezionare specifici
                print(f"\n📂 Trovati {len(files)} file nella directory")
                print("\n🤔 Cosa vuoi fare?")
                print("   1. 📚 Elabora TUTTI i file della cartella")
                print("   2. 🎯 Seleziona file specifici")
                print("   3. 🔙 Torna indietro")
                
                mode_choice = input("\n➤ Scelta (1-3): ").strip()
                
                if mode_choice == "3":
                    continue  # Torna alla selezione directory
                elif mode_choice == "1":
                    self._process_all_files(files)
                    return
                elif mode_choice == "2":
                    self._select_specific_files(files)
                    return
                else:
                    print("❌ Scelta non valida")
                    input("\n📝 Premi INVIO per riprovare...")
                    
            except KeyboardInterrupt:
                print("\n\n👋 Operazione annullata")
                return
            except Exception as e:
                self.ui.handle_error(e, "elaborazione file Oracle")
    
    def _process_all_files(self, files):
        """Elabora tutti i file della directory."""
        print(f"\n🚀 Elaborazione di tutti i {len(files)} file...")
        print("━" * 50)
        
        total_processed = 0
        total_success = 0
        total_failed = 0
        
        for i, file_path in enumerate(files, 1):
            print(f"\n📄 [{i}/{len(files)}] Elaborando: {file_path.name}")
            print("⏳ Elaborazione in corso...")
            
            try:
                # Usa il file service per elaborare
                file_config = FileConfig(file_path=str(file_path))
                results = self.lookup_service.process_lookup(file_config)
                
                # Genera output Excel
                output_file = f"lookup_result_{file_path.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                output_path = Path("Output") / output_file
                
                self.excel_service.create_excel_report(results, str(output_path), self.config)
                
                total_success += 1
                print(f"✅ Completato con successo: {output_file}")
                
            except Exception as e:
                total_failed += 1
                print(f"❌ Elaborazione fallita: {e}")
            
            total_processed += 1
        
        # Riepilogo finale
        print(f"\n🎯 RIEPILOGO ELABORAZIONE MULTIPLA")
        print(f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
        print(f"📊 File processati: {total_processed}")
        print(f"✅ Successi: {total_success}")
        print(f"❌ Fallimenti: {total_failed}")
        print(f"📁 File salvati in: Output/")
        
        input(f"\n📝 Premi INVIO per continuare...")
    
    def _select_specific_files(self, files):
        """Permette di selezionare file specifici."""
        while True:
            print(f"\n📋 FILE DISPONIBILI ({len(files)} totali):")
            print("━" * 50)
            
            # Mostra elenco file numerato
            for i, file_path in enumerate(files, 1):
                file_size = file_path.stat().st_size / (1024 * 1024)  # MB
                print(f"  {i:3}. {file_path.name} ({file_size:.2f} MB)")
            
            print("\n💡 Formati di selezione:")
            print("   • Singoli: 1,3,5")
            print("   • Intervalli: 1-5")
            print("   • Misti: 1,3-6,8")
            print("   • 'back' o 'b' per tornare indietro")
            
            file_choice = input("\n➤ Selezione: ").strip()
            
            # Controlla se l'utente vuole tornare indietro
            if file_choice.lower() in ['back', 'b', 'indietro']:
                return
            
            # Parsa la selezione
            selected_indices = self._parse_file_selection(file_choice, len(files))
            
            if selected_indices is None:
                print("❌ Selezione non valida. Usa il formato: 1,3,5 oppure 1-5 oppure 1,3-6,8")
                input("\n📝 Premi INVIO per riprovare...")
                continue
            
            # Mostra la selezione
            selected_files = [files[i-1] for i in selected_indices]
            print(f"\n✅ File selezionati ({len(selected_files)}):")
            for i, file_path in enumerate(selected_files, 1):
                print(f"   {i}. {file_path.name}")
            
            confirm = input(f"\n🤔 Procedere con l'elaborazione di {len(selected_files)} file? (s/n): ").strip().lower()
            
            if confirm in ['s', 'si', 'y', 'yes']:
                self._process_all_files(selected_files)
                return
            else:
                print("❌ Elaborazione annullata")
    
    def _parse_file_selection(self, selection: str, max_files: int) -> Optional[List[int]]:
        """Parsa la selezione dei file (es. 1,3,5-7)."""
        try:
            indices = []
            parts = selection.split(',')
            
            for part in parts:
                part = part.strip()
                if '-' in part:
                    # Intervallo (es. 3-7)
                    start, end = map(int, part.split('-'))
                    if start < 1 or end > max_files or start > end:
                        return None
                    indices.extend(range(start, end + 1))
                else:
                    # Singolo numero
                    num = int(part)
                    if num < 1 or num > max_files:
                        return None
                    indices.append(num)
            
            # Rimuovi duplicati e ordina
            return sorted(list(set(indices)))
            
        except ValueError:
            return None
    
    def _process_single_file(self):
        """Elabora un singolo file."""
        try:
            self.ui.print_separator()
            self.ui.print_info("ELABORAZIONE FILE SINGOLO")
            
            # Richiedi file di input
            input_file = self.ui.ask_file_path("Inserisci il percorso del file da elaborare")
            
            # Ottieni informazioni sul file
            file_info = self.file_service.get_file_info(input_file)
            self.ui.show_file_info(input_file, file_info)
            
            # Valida il file
            file_config = FileConfig(file_path=input_file)
            is_valid, errors = self.file_service.validate_file(file_config)
            
            if not is_valid:
                self.ui.print_error("File non valido:")
                for error in errors:
                    self.ui.print_error(f"  • {error}")
                return
            
            # Conferma elaborazione
            details = [
                f"File: {os.path.basename(input_file)}",
                f"Dimensione: {file_info['size_mb']:.2f} MB",
                f"Righe stimate: {file_info['estimated_rows']:,}"
            ]
            
            if not self.ui.confirm_operation("Elaborazione file", details):
                self.ui.print_info("Operazione annullata")
                return
            
            # Elabora il file
            self._execute_file_processing([input_file])
            
        except Exception as e:
            self.ui.handle_error(e, "elaborazione file singolo")
    
    def _process_multiple_files(self):
        """Elabora file multipli."""
        try:
            self.ui.print_separator()
            self.ui.print_info("ELABORAZIONE FILE MULTIPLI")
            
            # Richiedi file di input
            input_files = self.ui.ask_multiple_files("Inserisci i file da elaborare")
            
            if not input_files:
                self.ui.print_warning("Nessun file selezionato")
                return
            
            # Mostra informazioni sui file
            total_size = 0
            total_rows = 0
            
            self.ui.print_info(f"\nFile selezionati ({len(input_files)}):")
            for file_path in input_files:
                file_info = self.file_service.get_file_info(file_path)
                total_size += file_info['size_mb']
                total_rows += file_info['estimated_rows']
                
                print(f"  • {os.path.basename(file_path)} "
                      f"({file_info['size_mb']:.1f} MB, "
                      f"~{file_info['estimated_rows']:,} righe)")
            
            # Riepilogo
            details = [
                f"File da elaborare: {len(input_files)}",
                f"Dimensione totale: {total_size:.2f} MB",
                f"Righe stimate totali: {total_rows:,}"
            ]
            
            if not self.ui.confirm_operation("Elaborazione file multipli", details):
                self.ui.print_info("Operazione annullata")
                return
            
            # Elabora i file
            self._execute_file_processing(input_files)
            
        except Exception as e:
            self.ui.handle_error(e, "elaborazione file multipli")
    
    def _execute_file_processing(self, file_paths: List[str]):
        """
        Esegue l'elaborazione dei file.
        
        Args:
            file_paths: Lista dei file da elaborare
        """
        try:
            self.ui.print_separator()
            self.ui.print_info("Inizio elaborazione...")
            
            # Carica i dati da tutti i file
            self.ui.print_info("Caricamento dati...")
            combined_df = self.file_service.load_multiple_files(file_paths)
            
            if combined_df is None or combined_df.empty:
                self.ui.print_error("Impossibile caricare i dati dai file")
                return
            
            self.ui.print_success(f"Caricati {len(combined_df):,} record")
            
            # Ottimizza configurazione per la dimensione del dataset
            self.config.optimize_for_size(len(combined_df))
            
            # Crea configurazione di output
            base_filename = f"Oracle_Components_{get_current_timestamp()}"
            output_config = OutputConfig(
                base_filename=base_filename,
                output_directory=self.config.output_directory,
                create_compressed_archive=self.config.enable_7z_compression
            )
            
            # Elabora con progresso
            progress_callback = self.ui.create_progress_callback("Elaborazione tabelle")
            
            self.ui.print_info("Elaborazione lookup componenti...")
            enriched_df = self.lookup_service.process_dataframe(
                combined_df, 
                progress_callback=progress_callback
            )
            
            # Ottieni statistiche
            stats = self.lookup_service.get_statistics()
            self.ui.show_processing_stats(self.lookup_service.create_summary_report())
            
            # Crea file Excel
            self.ui.print_info("Creazione file Excel...")
            results = [result for result in self.lookup_service.process_tables(
                enriched_df['TABLE_NAME'].tolist()
            )]
            
            created_files = self.excel_service.create_excel_files(
                iter(results), 
                output_config, 
                stats
            )
            
            if not created_files:
                self.ui.print_error("Nessun file Excel creato")
                return
            
            # Compressione (se abilitata)
            compressed_file = None
            if self.config.enable_7z_compression and created_files:
                self.ui.print_info("Compressione file...")
                progress_callback = self.ui.create_progress_callback("Compressione")
                
                compressed_file = self.compression_service.compress_files(
                    created_files, 
                    output_config, 
                    progress_callback
                )
                
                if compressed_file:
                    self.ui.print_success(f"File compresso: {os.path.basename(compressed_file)}")
            
            # Mostra risultati finali
            final_files = [compressed_file] if compressed_file else created_files
            self.ui.show_completion_message("Elaborazione", final_files)
            
        except Exception as e:
            self.ui.handle_error(e, "elaborazione file")
    
    def _advanced_configuration(self):
        """Gestisce la configurazione avanzata."""
        try:
            self.ui.print_separator()
            self.ui.print_info("CONFIGURAZIONE AVANZATA")
            
            # Menu configurazione
            config_choices = [
                "Modifica percorso file componenti",
                "Configura directory di output", 
                "Impostazioni Excel",
                "Impostazioni compressione",
                "Impostazioni performance",
                "Ripristina configurazione di default"
            ]
            
            choice = self.ui.ask_choice("Seleziona configurazione da modificare", config_choices)
            
            if choice == 0:  # Percorso componenti
                new_path = self.ui.ask_file_path("Nuovo percorso file componenti")
                self.config.components_file = new_path
                self.ui.print_success("Percorso aggiornato (ricarica componenti per applicare)")
                
            elif choice == 1:  # Directory output
                new_dir = self.ui.ask_input("Nuova directory di output", self.config.output_directory)
                self.config.output_directory = new_dir
                ensure_directory(new_dir)
                self.ui.print_success(f"Directory di output: {new_dir}")
                
            elif choice == 2:  # Impostazioni Excel
                self._configure_excel_settings()
                
            elif choice == 3:  # Impostazioni compressione
                self._configure_compression_settings()
                
            elif choice == 4:  # Impostazioni performance
                self._configure_performance_settings()
                
            elif choice == 5:  # Reset configurazione
                if self.ui.ask_yes_no("Ripristinare la configurazione di default?"):
                    self.config = Config()
                    self.ui.print_success("Configurazione ripristinata")
            
        except Exception as e:
            self.ui.handle_error(e, "configurazione avanzata")
    
    def _configure_excel_settings(self):
        """Configura le impostazioni Excel."""
        print("\nImpostazioni Excel attuali:")
        print(f"  Righe max per file: {self.config.max_rows_per_file:,}")
        print(f"  Motore Excel: {self.config.excel_engine}")
        
        if self.ui.ask_yes_no("Modificare il numero massimo di righe per file?"):
            new_max = self.ui.ask_input("Nuovo massimo righe", str(self.config.max_rows_per_file))
            try:
                self.config.max_rows_per_file = int(new_max)
                self.ui.print_success(f"Massimo righe aggiornato: {self.config.max_rows_per_file:,}")
            except ValueError:
                self.ui.print_error("Valore non valido")
        
        if self.ui.ask_yes_no("Cambiare motore Excel?"):
            engines = ["xlsxwriter", "openpyxl"]
            choice = self.ui.ask_choice("Seleziona motore Excel", engines)
            self.config.excel_engine = engines[choice]
            self.ui.print_success(f"Motore Excel: {self.config.excel_engine}")
    
    def _configure_compression_settings(self):
        """Configura le impostazioni di compressione."""
        print("\nImpostazioni compressione attuali:")
        print(f"  Compressione 7z: {'Abilitata' if self.config.enable_7z_compression else 'Disabilitata'}")
        print(f"  Livello compressione: {self.config.compression_level}")
        print(f"  Soglia compressione: {self.config.compression_threshold_mb} MB")
        
        self.config.enable_7z_compression = self.ui.ask_yes_no(
            "Abilitare compressione 7z?", 
            self.config.enable_7z_compression
        )
        
        if self.config.enable_7z_compression:
            level_str = self.ui.ask_input(
                "Livello compressione (0-9)", 
                str(self.config.compression_level)
            )
            try:
                level = int(level_str)
                if 0 <= level <= 9:
                    self.config.compression_level = level
                else:
                    self.ui.print_error("Livello deve essere tra 0 e 9")
            except ValueError:
                self.ui.print_error("Valore non valido")
        
        self.ui.print_success("Impostazioni compressione aggiornate")
    
    def _configure_performance_settings(self):
        """Configura le impostazioni di performance."""
        print("\nImpostazioni performance attuali:")
        print(f"  Elaborazione parallela: {'Abilitata' if self.config.enable_multiprocessing else 'Disabilitata'}")
        print(f"  Worker paralleli: {self.config.max_workers}")
        print(f"  Dimensione chunk: {self.config.chunk_size:,}")
        print(f"  Cache componenti: {'Abilitata' if self.config.cache_components else 'Disabilitata'}")
        
        self.config.enable_multiprocessing = self.ui.ask_yes_no(
            "Abilitare elaborazione parallela?",
            self.config.enable_multiprocessing
        )
        
        if self.config.enable_multiprocessing:
            workers_str = self.ui.ask_input(
                "Numero worker paralleli", 
                str(self.config.max_workers)
            )
            try:
                workers = int(workers_str)
                if workers > 0:
                    self.config.max_workers = workers
                else:
                    self.ui.print_error("Numero worker deve essere positivo")
            except ValueError:
                self.ui.print_error("Valore non valido")
        
        self.config.cache_components = self.ui.ask_yes_no(
            "Abilitare cache componenti?",
            self.config.cache_components
        )
        
        self.ui.print_success("Impostazioni performance aggiornate")
    
    def _show_system_info(self):
        """Mostra informazioni sul sistema."""
        try:
            self.ui.print_separator()
            self.ui.print_info("INFORMAZIONI SISTEMA")
            
            # Informazioni Python
            print(f"\nPython: {sys.version}")
            print(f"Piattaforma: {sys.platform}")
            
            # Informazioni memoria (se disponibile)
            try:
                import psutil
                memory = psutil.virtual_memory()
                print(f"Memoria totale: {memory.total / (1024**3):.1f} GB")
                print(f"Memoria disponibile: {memory.available / (1024**3):.1f} GB")
                print(f"CPU cores: {psutil.cpu_count()}")
            except ImportError:
                print("Informazioni sistema estese non disponibili (installare psutil)")
            
            # Informazioni componenti
            if self.component_service.is_loaded:
                stats = self.component_service.get_statistics()
                print(f"\nComponenti Oracle:")
                print(f"  Totale componenti: {stats['total_components']}")
                print(f"  Prefissi unici: {stats['unique_prefixes']}")
                print(f"  Con ufficio IT: {stats.get('components_with_office', 0)}")
                print(f"  Con responsabile: {stats.get('components_with_responsible', 0)}")
            
            # Informazioni file
            if os.path.exists(self.config.components_file):
                file_info = self.file_service.get_file_info(self.config.components_file)
                print(f"\nFile componenti:")
                print(f"  Percorso: {self.config.components_file}")
                print(f"  Dimensione: {file_info['size_mb']:.2f} MB")
                print(f"  Encoding: {file_info.get('detected_encoding', 'N/A')}")
            
            self.ui.wait_for_enter()
            
        except Exception as e:
            self.ui.handle_error(e, "informazioni sistema")
    
    def _reload_components(self):
        """Ricarica i componenti Oracle."""
        try:
            self.ui.print_separator()
            self.ui.print_info("RICARICAMENTO COMPONENTI")
            
            if self.ui.ask_yes_no("Ricaricare i componenti Oracle?"):
                self.ui.print_info("Ricaricamento in corso...")
                
                if self.component_service.reload():
                    stats = self.component_service.get_statistics()
                    self.ui.print_success(f"Ricaricati {stats['total_components']} componenti")
                else:
                    self.ui.print_error("Errore nel ricaricamento dei componenti")
        
        except Exception as e:
            self.ui.handle_error(e, "ricaricamento componenti")
    
    def _analyze_failures(self):
        """Analizza i fallimenti raggruppati per TABLE_NAME."""
        try:
            self.ui.print_separator()
            self.ui.print_info("ANALISI FALLIMENTI")
            
            # Richiedi file di input
            input_file = self.ui.ask_file_path("Inserisci il percorso del file Oracle da analizzare")
            
            # Esegui elaborazione
            file_config = FileConfig(file_path=input_file)
            results = self.lookup_service.process_lookup(file_config)
            
            # Raggruppa fallimenti per TABLE_NAME
            failures_by_table = {}
            for result in results:
                if not result.found:
                    table_name = result.input_data.get('TABLE_NAME', 'UNKNOWN')
                    if table_name not in failures_by_table:
                        failures_by_table[table_name] = []
                    failures_by_table[table_name].append(result)
            
            # Mostra analisi
            if failures_by_table:
                self.ui.print_info(f"Trovati fallimenti per {len(failures_by_table)} tabelle:")
                for table_name, failures in failures_by_table.items():
                    print(f"\n📊 {table_name}: {len(failures)} fallimenti")
                    for failure in failures[:5]:  # Mostra primi 5
                        oracle_name = failure.input_data.get('ORACLE_NAME', 'N/A')
                        print(f"  • {oracle_name}")
                    if len(failures) > 5:
                        print(f"  ... e altri {len(failures) - 5} fallimenti")
            else:
                self.ui.print_success("Nessun fallimento trovato!")
                
        except Exception as e:
            self.ui.handle_error(e, "analisi fallimenti")
    
    def _generate_full_report(self):
        """Genera un report completo con tutte le statistiche."""
        try:
            self.ui.print_separator()
            self.ui.print_info("GENERAZIONE REPORT COMPLETO")
            
            # Richiedi file di input
            input_files = []
            while True:
                file_path = input("\nInserisci percorso file (INVIO per terminare): ").strip()
                if not file_path:
                    break
                if os.path.exists(file_path):
                    input_files.append(file_path)
                    print(f"✅ Aggiunto: {os.path.basename(file_path)}")
                else:
                    print("❌ File non trovato")
            
            if not input_files:
                self.ui.print_warning("Nessun file selezionato")
                return
            
            # Elabora tutti i file
            all_results = []
            for file_path in input_files:
                file_config = FileConfig(file_path=file_path)
                results = self.lookup_service.process_lookup(file_config)
                all_results.extend(results)
            
            # Genera statistiche complete
            total_records = len(all_results)
            found_records = sum(1 for r in all_results if r.found)
            success_rate = (found_records / total_records * 100) if total_records > 0 else 0
            
            print(f"\n📊 REPORT COMPLETO:")
            print(f"  Total records processati: {total_records:,}")
            print(f"  Match trovati: {found_records:,}")
            print(f"  Tasso di successo: {success_rate:.2f}%")
            
            # Genera file Excel con report
            output_file = f"report_completo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            self.excel_service.create_excel_report(all_results, output_file, self.config)
            self.ui.print_success(f"Report salvato in: {output_file}")
            
        except Exception as e:
            self.ui.handle_error(e, "generazione report")
    
    def _cleanup(self):
        """Pulizia finale dell'applicazione."""
        try:
            self.ui.print_info("Chiusura applicazione...")
            # Eventuali operazioni di pulizia
            
        except Exception as e:
            self.logger.error(f"Errore durante la pulizia: {e}")


def main():
    """Funzione principale dell'applicazione."""
    try:
        app = OracleComponentLookupApp()
        app.run()
    except KeyboardInterrupt:
        print("\nApplicazione interrotta dall'utente.")
        sys.exit(0)
    except Exception as e:
        print(f"Errore fatale nell'applicazione: {e}")
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
