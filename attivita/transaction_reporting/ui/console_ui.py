"""
Interfaccia console per Transaction Reporting - Rejecting Mensile.
Gestisce l'interazione con l'utente attraverso menu e prompt.
"""

import logging
from datetime import datetime, date
from typing import List, Optional, Dict, Any
from pathlib import Path

from models.transaction_reporting import MonthlyReportConfig, DataSource


class TransactionReportingUI:
    """Interfaccia utente per il sistema di transaction reporting."""
    
    def __init__(self):
        """Inizializza l'interfaccia utente."""
        self.logger = logging.getLogger(__name__)
    
    def show_welcome(self):
        """Mostra il messaggio di benvenuto."""
        print("=" * 70)
        print("    TRANSACTION REPORTING - REJECTING MENSILE")
        print("=" * 70)
        print()
        print("Sistema per l'analisi e reporting delle transazioni rifiutate")
        print("Versione 1.0.0 - Sviluppato da GBS12162")
        print()
    
    def show_main_menu(self) -> str:
        """
        Mostra il menu principale e ottiene la selezione dell'utente.
        
        Returns:
            Scelta dell'utente
        """
        print("=" * 50)
        print("🔧 TRANSACTION REPORTING - MENU PRINCIPALE")
        print("=" * 50)
        print("1. 📊 Genera Report Mensile")
        print("2. 📈 Analizza File Transazioni")
        print("3. 📋 Visualizza Statistiche")
        print("4. ⚙️  Configurazioni")
        print("5. ❓ Aiuto")
        print("6. ❌ Esci")
        print()
        
        while True:
            try:
                choice = input("➤ Seleziona opzione (1-6): ").strip()
                if choice in ['1', '2', '3', '4', '5', '6']:
                    return choice
                else:
                    print("❌ Opzione non valida. Seleziona un numero da 1 a 6.")
            except KeyboardInterrupt:
                return "6"
            except Exception as e:
                print(f"❌ Errore nell'input: {e}")
    
    def get_report_configuration(self) -> Optional[MonthlyReportConfig]:
        """
        Ottiene la configurazione per il report mensile dall'utente.
        
        Returns:
            Configurazione del report o None se annullato
        """
        try:
            print("=" * 60)
            print("📊 CONFIGURAZIONE REPORT MENSILE")
            print("━" * 60)
            print()
            
            # Anno
            current_year = datetime.now().year
            year_input = input(f"📅 Anno [{current_year}]: ").strip()
            year = int(year_input) if year_input else current_year
            
            # Mese
            current_month = datetime.now().month
            month_input = input(f"📅 Mese (1-12) [{current_month}]: ").strip()
            month = int(month_input) if month_input else current_month
            
            if not (1 <= month <= 12):
                print("❌ Mese non valido. Deve essere tra 1 e 12.")
                return None
            
            # Opzioni avanzate
            print("\n🔧 OPZIONI AVANZATE:")
            print("1. Includi transazioni in sospeso? (s/n) [s]: ", end="")
            include_pending = input().strip().lower() not in ['n', 'no']
            
            print("2. Includi transazioni fallite? (s/n) [s]: ", end="")
            include_failed = input().strip().lower() not in ['n', 'no']
            
            # Formato di esportazione
            print("\n📁 FORMATO DI ESPORTAZIONE:")
            print("1. Excel (.xlsx)")
            print("2. CSV")
            print("3. JSON")
            print("4. Tutti i formati")
            
            format_choice = input("➤ Seleziona formato (1-4) [1]: ").strip()
            format_map = {
                '1': 'excel',
                '2': 'csv', 
                '3': 'json',
                '4': 'all'
            }
            export_format = format_map.get(format_choice, 'excel')
            
            # Directory di output
            default_output = "output_tr_mensile"
            output_dir = input(f"📂 Directory output [{default_output}]: ").strip()
            if not output_dir:
                output_dir = default_output
            
            # Crea la configurazione
            config = MonthlyReportConfig(
                year=year,
                month=month,
                include_pending=include_pending,
                include_failed=include_failed,
                export_format=export_format,
                output_directory=output_dir
            )
            
            # Mostra riepilogo
            print("\n" + "=" * 50)
            print("📋 RIEPILOGO CONFIGURAZIONE")
            print("=" * 50)
            print(f"Periodo: {month:02d}/{year}")
            print(f"Include sospese: {'Sì' if include_pending else 'No'}")
            print(f"Include fallite: {'Sì' if include_failed else 'No'}")
            print(f"Formato export: {export_format.upper()}")
            print(f"Directory output: {output_dir}")
            print()
            
            confirm = input("✅ Confermi la configurazione? (s/n) [s]: ").strip().lower()
            if confirm in ['n', 'no']:
                return None
                
            return config
            
        except KeyboardInterrupt:
            print("\n👋 Operazione annullata")
            return None
        except ValueError as e:
            print(f"❌ Errore nei dati inseriti: {e}")
            return None
        except Exception as e:
            print(f"❌ Errore inaspettato: {e}")
            return None
    
    def get_data_source_configuration(self) -> Optional[DataSource]:
        """
        Ottiene la configurazione della fonte dati dall'utente.
        
        Returns:
            Configurazione della fonte dati o None se annullato
        """
        try:
            print("=" * 60)
            print("📊 CONFIGURAZIONE FONTE DATI")
            print("━" * 60)
            print()
            
            print("🔌 TIPO DI FONTE DATI:")
            print("1. File CSV")
            print("2. File Excel")
            print("3. Database")
            print()
            
            source_choice = input("➤ Seleziona tipo (1-3): ").strip()
            
            if source_choice == '1':
                return self._configure_csv_source()
            elif source_choice == '2':
                return self._configure_excel_source()
            elif source_choice == '3':
                return self._configure_database_source()
            else:
                print("❌ Opzione non valida")
                return None
                
        except KeyboardInterrupt:
            print("\n👋 Operazione annullata")
            return None
        except Exception as e:
            print(f"❌ Errore: {e}")
            return None
    
    def _configure_csv_source(self) -> Optional[DataSource]:
        """Configura una fonte dati CSV."""
        file_path = input("📁 Percorso del file CSV: ").strip()
        if not file_path or not Path(file_path).exists():
            print("❌ File non trovato")
            return None
        
        # Mapping colonne (opzionale)
        print("\n🗂️  MAPPING COLONNE (opzionale - premi INVIO per saltare):")
        headers = {}
        
        standard_fields = [
            "transaction_id", "account_number", "amount", "currency",
            "transaction_date", "transaction_type", "status", 
            "rejection_reason", "rejection_code"
        ]
        
        for field in standard_fields:
            column = input(f"Colonna per '{field}': ").strip()
            if column:
                headers[column] = field
        
        return DataSource(
            source_type="csv",
            file_path=file_path,
            headers=headers if headers else None
        )
    
    def _configure_excel_source(self) -> Optional[DataSource]:
        """Configura una fonte dati Excel."""
        file_path = input("📁 Percorso del file Excel: ").strip()
        if not file_path or not Path(file_path).exists():
            print("❌ File non trovato")
            return None
        
        sheet_name = input("📋 Nome del foglio (opzionale): ").strip()
        
        return DataSource(
            source_type="excel",
            file_path=file_path,
            table_name=sheet_name if sheet_name else None
        )
    
    def _configure_database_source(self) -> Optional[DataSource]:
        """Configura una fonte dati database."""
        print("🔧 Configurazione database non ancora implementata")
        print("Usa file CSV o Excel per ora.")
        return None
    
    def show_processing_status(self, message: str, progress: Optional[float] = None):
        """
        Mostra lo stato di elaborazione.
        
        Args:
            message: Messaggio di stato
            progress: Progresso in percentuale (0-100)
        """
        if progress is not None:
            progress_bar = self._create_progress_bar(progress)
            print(f"\r🔄 {message} {progress_bar} {progress:.1f}%", end="", flush=True)
        else:
            print(f"🔄 {message}")
    
    def show_results(self, result: Dict[str, Any]):
        """
        Mostra i risultati dell'elaborazione.
        
        Args:
            result: Dizionario con i risultati
        """
        print("\n" + "=" * 60)
        print("📊 RISULTATI ELABORAZIONE")
        print("=" * 60)
        
        if result.get("success"):
            print("✅ Elaborazione completata con successo!")
            print()
            
            if "report" in result:
                report = result["report"]
                print(f"📋 Report ID: {report.report_id}")
                print(f"📅 Periodo: {report.period_start.strftime('%Y-%m-%d')} - {report.period_end.strftime('%Y-%m-%d')}")
                print(f"📊 Totale transazioni: {report.total_transactions:,}")
                print(f"❌ Transazioni rifiutate: {report.rejected_transactions:,}")
                print(f"📈 Tasso di rifiuto: {report.rejection_rate:.2f}%")
                print(f"💰 Importo totale: {report.total_amount:,.2f}")
                print(f"💸 Importo rifiutato: {report.rejected_amount:,.2f}")
            
            if "output_files" in result:
                print("\n📁 FILE GENERATI:")
                for file_path in result["output_files"]:
                    print(f"   • {file_path}")
            
            if "processing_time" in result:
                print(f"\n⏱️  Tempo di elaborazione: {result['processing_time']:.2f} secondi")
        
        else:
            print("❌ Elaborazione fallita!")
            if "errors" in result:
                print("\n🔴 ERRORI:")
                for error in result["errors"]:
                    print(f"   • {error}")
        
        print("\n" + "=" * 60)
    
    def show_statistics(self, stats: Dict[str, Any]):
        """
        Mostra le statistiche dettagliate.
        
        Args:
            stats: Dizionario con le statistiche
        """
        print("=" * 60)
        print("📈 STATISTICHE DETTAGLIATE")
        print("=" * 60)
        
        if "transactions" in stats:
            trans_stats = stats["transactions"]
            print("📊 TRANSAZIONI:")
            print(f"   • Totali: {trans_stats.get('total', 0):,}")
            print(f"   • Rifiutate: {trans_stats.get('rejected', 0):,}")
            print(f"   • Approvate: {trans_stats.get('approved', 0):,}")
            print(f"   • Tasso rifiuto: {trans_stats.get('rejection_rate', 0):.2f}%")
        
        if "amounts" in stats:
            amount_stats = stats["amounts"]
            print("\n💰 IMPORTI:")
            print(f"   • Totale: {amount_stats.get('total', 0):,.2f}")
            print(f"   • Rifiutato: {amount_stats.get('rejected', 0):,.2f}")
            print(f"   • Media transazione: {amount_stats.get('average_transaction', 0):.2f}")
            print(f"   • Media rifiuto: {amount_stats.get('average_rejected', 0):.2f}")
        
        if "rejection_analysis" in stats:
            rejection = stats["rejection_analysis"]
            
            if "top_reasons" in rejection:
                print("\n🔴 TOP MOTIVI DI RIFIUTO:")
                for reason in rejection["top_reasons"][:5]:
                    print(f"   • {reason['reason']}: {reason['count']} volte")
        
        print("\n" + "=" * 60)
    
    def _create_progress_bar(self, progress: float, width: int = 30) -> str:
        """Crea una barra di progresso ASCII."""
        filled = int(width * progress / 100)
        bar = "█" * filled + "░" * (width - filled)
        return f"[{bar}]"
    
    def show_error(self, message: str, details: Optional[str] = None):
        """Mostra un messaggio di errore."""
        print(f"\n❌ ERRORE: {message}")
        if details:
            print(f"   Dettagli: {details}")
    
    def show_warning(self, message: str):
        """Mostra un messaggio di avviso."""
        print(f"\n⚠️  AVVISO: {message}")
    
    def show_info(self, message: str):
        """Mostra un messaggio informativo."""
        print(f"\n💡 INFO: {message}")
    
    def wait_for_key(self, message: str = "Premi INVIO per continuare..."):
        """Aspetta che l'utente prema un tasto."""
        try:
            input(f"\n📝 {message}")
        except KeyboardInterrupt:
            pass