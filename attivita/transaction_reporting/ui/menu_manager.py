"""
Menu manager per Transaction Reporting - Rejecting Mensile.
Gestisce la navigazione tra i menu e le operazioni.
"""

import logging
from typing import Optional, Dict, Any
from datetime import datetime

from models.transaction_reporting import MonthlyReportConfig, DataSource, ProcessingResult
from .console_ui import TransactionReportingUI


class TransactionReportingMenuManager:
    """Manager per la gestione dei menu dell'applicazione."""
    
    def __init__(self, ui: TransactionReportingUI):
        """
        Inizializza il menu manager.
        
        Args:
            ui: Interfaccia utente
        """
        self.ui = ui
        self.logger = logging.getLogger(__name__)
        self.running = True
        
        # Servizi (da iniettare dal main)
        self.data_service = None
        self.report_service = None
        self.export_service = None
        
        # Stato dell'applicazione
        self.current_config: Optional[MonthlyReportConfig] = None
        self.current_data_source: Optional[DataSource] = None
        self.last_report: Optional[Any] = None
        self.last_statistics: Optional[Dict[str, Any]] = None
    
    def set_services(self, data_service, report_service, export_service):
        """
        Imposta i servizi necessari.
        
        Args:
            data_service: Servizio per i dati delle transazioni
            report_service: Servizio per la generazione di report
            export_service: Servizio per l'esportazione
        """
        self.data_service = data_service
        self.report_service = report_service
        self.export_service = export_service
    
    def run(self):
        """Avvia il loop principale del menu."""
        try:
            self.ui.show_welcome()
            
            while self.running:
                try:
                    choice = self.ui.show_main_menu()
                    self._handle_menu_choice(choice)
                    
                except KeyboardInterrupt:
                    self.ui.show_info("Operazione interrotta dall'utente")
                    self.running = False
                except Exception as e:
                    self.logger.error(f"Errore nel menu principale: {e}")
                    self.ui.show_error("Errore nel menu principale", str(e))
                    self.ui.wait_for_key()
            
            self.ui.show_info("Applicazione terminata. Arrivederci!")
            
        except Exception as e:
            self.logger.error(f"Errore critico nell'applicazione: {e}")
            self.ui.show_error("Errore critico", str(e))
    
    def _handle_menu_choice(self, choice: str):
        """
        Gestisce la scelta del menu principale.
        
        Args:
            choice: Scelta dell'utente
        """
        if choice == '1':
            self._handle_generate_report()
        elif choice == '2':
            self._handle_analyze_file()
        elif choice == '3':
            self._handle_show_statistics()
        elif choice == '4':
            self._handle_configurations()
        elif choice == '5':
            self._handle_help()
        elif choice == '6':
            self.running = False
        else:
            self.ui.show_error("Opzione non valida")
    
    def _handle_generate_report(self):
        """Gestisce la generazione di un report mensile."""
        try:
            self.ui.show_info("Inizio processo di generazione report...")
            
            # Ottieni configurazione report
            config = self.ui.get_report_configuration()
            if not config:
                self.ui.show_warning("Generazione report annullata")
                return
            
            self.current_config = config
            
            # Ottieni configurazione fonte dati
            data_source = self.ui.get_data_source_configuration()
            if not data_source:
                self.ui.show_warning("Configurazione fonte dati annullata")
                return
            
            self.current_data_source = data_source
            
            # Carica i dati
            self.ui.show_processing_status("Caricamento dati...")
            transactions = self._load_transaction_data(data_source)
            
            if not transactions:
                self.ui.show_error("Nessuna transazione caricata")
                return
            
            # Genera report
            self.ui.show_processing_status("Generazione report...")
            report = self.report_service.generate_monthly_report(transactions, config)
            self.last_report = report
            
            # Genera statistiche
            self.ui.show_processing_status("Calcolo statistiche...")
            statistics = self.report_service.generate_summary_statistics(report)
            self.last_statistics = statistics
            
            # Esporta report
            self.ui.show_processing_status("Esportazione report...")
            export_result = self.export_service.export_report(report, config, statistics)
            
            # Mostra risultati
            result = {
                "success": export_result.success,
                "report": report,
                "output_files": export_result.output_files,
                "processing_time": export_result.processing_time,
                "errors": export_result.errors
            }
            
            self.ui.show_results(result)
            
        except Exception as e:
            self.logger.error(f"Errore nella generazione del report: {e}")
            self.ui.show_error("Errore nella generazione del report", str(e))
        
        finally:
            self.ui.wait_for_key()
    
    def _handle_analyze_file(self):
        """Gestisce l'analisi di un file di transazioni."""
        try:
            self.ui.show_info("Analisi file transazioni...")
            
            # Ottieni configurazione fonte dati
            data_source = self.ui.get_data_source_configuration()
            if not data_source:
                self.ui.show_warning("Analisi annullata")
                return
            
            # Carica i dati
            self.ui.show_processing_status("Caricamento e analisi dati...")
            transactions = self._load_transaction_data(data_source)
            
            if not transactions:
                self.ui.show_error("Nessuna transazione caricata")
                return
            
            # Ottieni statistiche
            statistics = self.data_service.get_statistics()
            self.last_statistics = statistics
            
            # Mostra risultati
            self.ui.show_statistics(statistics)
            
        except Exception as e:
            self.logger.error(f"Errore nell'analisi del file: {e}")
            self.ui.show_error("Errore nell'analisi del file", str(e))
        
        finally:
            self.ui.wait_for_key()
    
    def _handle_show_statistics(self):
        """Gestisce la visualizzazione delle statistiche."""
        if self.last_statistics:
            self.ui.show_statistics(self.last_statistics)
        else:
            self.ui.show_warning("Nessuna statistica disponibile. Genera prima un report o analizza un file.")
        
        self.ui.wait_for_key()
    
    def _handle_configurations(self):
        """Gestisce il menu delle configurazioni."""
        print("\\n" + "=" * 50)
        print("⚙️  CONFIGURAZIONI")
        print("=" * 50)
        
        if self.current_config:
            print("📊 CONFIGURAZIONE REPORT CORRENTE:")
            print(f"   • Periodo: {self.current_config.month:02d}/{self.current_config.year}")
            print(f"   • Formato: {self.current_config.export_format}")
            print(f"   • Output: {self.current_config.output_directory}")
            print(f"   • Include sospese: {'Sì' if self.current_config.include_pending else 'No'}")
            print(f"   • Include fallite: {'Sì' if self.current_config.include_failed else 'No'}")
        else:
            print("📊 Nessuna configurazione report attiva")
        
        print()
        
        if self.current_data_source:
            print("🔌 FONTE DATI CORRENTE:")
            print(f"   • Tipo: {self.current_data_source.source_type}")
            print(f"   • Percorso: {self.current_data_source.file_path or 'N/A'}")
        else:
            print("🔌 Nessuna fonte dati configurata")
        
        self.ui.wait_for_key()
    
    def _handle_help(self):
        """Gestisce la visualizzazione dell'aiuto."""
        print("\\n" + "=" * 60)
        print("❓ AIUTO - TRANSACTION REPORTING")
        print("=" * 60)
        print()
        print("📊 GENERA REPORT MENSILE:")
        print("   Crea un report completo delle transazioni rifiutate per un mese specifico.")
        print("   Supporta esportazione in Excel, CSV e JSON.")
        print()
        print("📈 ANALIZZA FILE TRANSAZIONI:")
        print("   Carica e analizza un file di transazioni per ottenere statistiche immediate.")
        print("   Supporta file CSV ed Excel.")
        print()
        print("📋 VISUALIZZA STATISTICHE:")
        print("   Mostra le statistiche dell'ultima analisi o report generato.")
        print()
        print("⚙️  CONFIGURAZIONI:")
        print("   Visualizza le configurazioni correnti dell'applicazione.")
        print()
        print("🔧 FORMATI FILE SUPPORTATI:")
        print("   • CSV: File di testo separato da virgole")
        print("   • Excel: File .xlsx o .xls")
        print()
        print("📋 CAMPI RICHIESTI PER LE TRANSAZIONI:")
        print("   • transaction_id: ID univoco della transazione")
        print("   • account_number: Numero del conto")
        print("   • amount: Importo della transazione")
        print("   • currency: Valuta (es. EUR, USD)")
        print("   • transaction_date: Data della transazione")
        print("   • transaction_type: Tipo di transazione")
        print("   • status: Stato (APPROVED, REJECTED, FAILED, ecc.)")
        print("   • rejection_reason: Motivo del rifiuto (opzionale)")
        print()
        print("💡 SUGGERIMENTI:")
        print("   • Usa nomi di colonne coerenti nei tuoi file")
        print("   • Le date devono essere in formato riconoscibile (YYYY-MM-DD)")
        print("   • Gli importi devono essere numerici")
        print("   • Verifica l'encoding del file (UTF-8 consigliato)")
        print()
        print("=" * 60)
        
        self.ui.wait_for_key()
    
    def _load_transaction_data(self, data_source: DataSource):
        """
        Carica i dati delle transazioni dalla fonte specificata.
        
        Args:
            data_source: Configurazione della fonte dati
            
        Returns:
            Lista delle transazioni caricate
        """
        try:
            if data_source.source_type == "csv":
                return self.data_service.load_from_csv(data_source.file_path, data_source)
            elif data_source.source_type == "excel":
                return self.data_service.load_from_excel(
                    data_source.file_path, 
                    data_source, 
                    data_source.table_name
                )
            else:
                raise ValueError(f"Tipo di fonte dati non supportato: {data_source.source_type}")
                
        except Exception as e:
            self.logger.error(f"Errore nel caricamento dei dati: {e}")
            raise