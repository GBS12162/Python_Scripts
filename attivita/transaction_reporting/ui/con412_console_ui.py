"""
Interfaccia utente console per Transaction Reporting CON-412 - Rejecting Mensile
"""

import logging
from typing import Optional, List
from pathlib import Path

from models.transaction_reporting import ProcessingConfig, SharePointConfig, QualityControlResult
from services.sharepoint_service import SharePointService
from services.isin_processing_service import ISINProcessingService
from services.report_export_service import ReportExportService
from services.isin_validation_service import ISINValidationService
from config.con412_config import CON412Config
from utils.progress_utils import ProgressBar


class CON412ConsoleUI:
    """Interfaccia console specifica per il processo CON-412."""
    
    def __init__(self):
        """Inizializza l'interfaccia console."""
        self.logger = logging.getLogger(__name__)
        
        # Servizi
        self.sharepoint_service = SharePointService()
        self.isin_service = ISINProcessingService()
        self.export_service = ReportExportService()
        self.validation_service = ISINValidationService()
        
        # Configurazioni
        self.config = CON412Config()
        self.current_month: Optional[str] = None
        self.downloaded_files: List[Path] = []
        self.processing_results = None
    
    def run(self):
        """Esegue l'interfaccia console principale."""
        try:
            self._show_welcome()
            
            while True:
                self._show_main_menu()
                choice = input("\nScelta: ").strip()
                
                if choice == "1":
                    self._configure_api_url()
                elif choice == "2":
                    self._select_month_and_download()
                elif choice == "3":
                    self._process_downloaded_files()
                elif choice == "4":
                    self._run_quality_controls()
                elif choice == "5":
                    self._export_results()
                elif choice == "6":
                    self._show_validation_cache_stats()
                elif choice == "7":
                    self._clear_validation_cache()
                elif choice == "8":
                    self._show_processing_summary()
                elif choice == "0":
                    print("Arrivederci!")
                    break
                else:
                    print("Scelta non valida. Riprovare.")
                    
        except KeyboardInterrupt:
            print("\n\nOperazione interrotta dall'utente.")
        except Exception as e:
            self.logger.error(f"Errore nell'interfaccia console: {e}")
            print(f"Errore: {e}")
    
    def _show_welcome(self):
        """Mostra il messaggio di benvenuto."""
        print("=" * 70)
        print("    TRANSACTION REPORTING - REJECTING MENSILE")
        print("                   CON-412 Processing")
        print("=" * 70)
        print(f"SharePoint: {self.config.sharepoint_config.site_url}")
        print(f"Folder: {self.config.sharepoint_config.folder_path}")
        print(f"Pattern file: {self.config.sharepoint_config.file_pattern}")
        print("=" * 70)
    
    def _show_main_menu(self):
        """Mostra il menu principale."""
        print("\n" + "=" * 50)
        print("MENU PRINCIPALE CON-412")
        print("=" * 50)
        print("1. 🔧 Configura URL API validazione ISIN")
        print("2. 📥 Seleziona mese e scarica file")
        print("3. ⚙️  Elabora file scaricati")
        print("4. 🔍 Esegui controlli di qualità")
        print("5. 📊 Esporta risultati")
        print("6. 💾 Statistiche cache validazione")
        print("7. 🗑️  Pulisci cache validazione")
        print("8. 📋 Mostra riepilogo elaborazione")
        print("0. ❌ Esci")
        print("=" * 50)
        
        # Mostra stato corrente
        print("\n📊 STATO CORRENTE:")
        if self.current_month:
            print(f"  Mese selezionato: {self.current_month}")
        else:
            print("  Mese: Non selezionato")
            
        if self.downloaded_files:
            print(f"  File scaricati: {len(self.downloaded_files)}")
        else:
            print("  File scaricati: Nessuno")
            
        if self.processing_results:
            total_orders = sum(len(g.orders) for g in self.processing_results)
            print(f"  Gruppi ISIN processati: {len(self.processing_results)}")
            print(f"  Ordini totali: {total_orders}")
        else:
            print("  Dati processati: Nessuno")
        
        # Stato API
        api_url = self.config.controlli_config.isin_api_url
        if api_url:
            print(f"  API ISIN: Configurata")
        else:
            print("  API ISIN: ❌ Non configurata")
    
    def _configure_api_url(self):
        """Configura l'URL dell'API per la validazione ISIN."""
        try:
            print("\n" + "=" * 60)
            print("🔧 CONFIGURAZIONE API VALIDAZIONE ISIN")
            print("=" * 60)
            
            current_url = self.config.controlli_config.isin_api_url
            if current_url:
                print(f"URL corrente: {current_url}")
            else:
                print("❌ Nessun URL configurato")
            
            print("\nInserire il nuovo URL dell'API per la validazione ISIN:")
            print("(Lascia vuoto per mantenere quello corrente)")
            print("\nEsempio: https://api.example.com/isin/validate")
            
            new_url = input("\n➤ URL API: ").strip()
            
            if new_url:
                self.validation_service.set_api_url(new_url)
                print(f"✅ URL API configurato: {new_url}")
            else:
                print("⚠️  URL non modificato")
                
        except Exception as e:
            self.logger.error(f"Errore nella configurazione API URL: {e}")
            print(f"❌ Errore: {e}")
    
    def _select_month_and_download(self):
        """Seleziona il mese e scarica i file da SharePoint."""
        try:
            print("\n" + "=" * 60)
            print("📥 SELEZIONE MESE E DOWNLOAD FILE")
            print("=" * 60)
            
            # Selezione mese
            print("📅 Mesi disponibili:")
            months = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE",
                     "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"]
            
            for i, month in enumerate(months, 1):
                marker = "👉" if (self.current_month == month) else "  "
                print(f"{marker} {i:2d}. {month}")
            
            try:
                choice = input("\n➤ Seleziona mese (1-12): ").strip()
                if not choice:  # Mantieni selezione corrente
                    if not self.current_month:
                        print("❌ Nessun mese selezionato")
                        return
                else:
                    choice_num = int(choice)
                    if 1 <= choice_num <= 12:
                        self.current_month = months[choice_num - 1]
                        print(f"✅ Mese selezionato: {self.current_month}")
                    else:
                        print("❌ Scelta non valida")
                        return
            except ValueError:
                print("❌ Inserire un numero valido")
                return
            
            # Download file
            print(f"\n📡 Scaricamento file CON-412_{self.current_month}...")
            print(f"Da SharePoint: {self.config.sharepoint_config.site_url}")
            
            config = ProcessingConfig(
                input_folder=Path("downloads"),
                output_folder=Path("output"),
                month=self.current_month
            )
            
            with ProgressBar(description="Download file") as progress:
                self.downloaded_files = self.sharepoint_service.download_monthly_files(
                    self.current_month, 
                    config,
                    progress_callback=progress.update
                )
            
            if self.downloaded_files:
                print(f"\n✅ Scaricati {len(self.downloaded_files)} file:")
                for file_path in self.downloaded_files:
                    size_mb = file_path.stat().st_size / (1024 * 1024)
                    print(f"  📄 {file_path.name} ({size_mb:.1f} MB)")
            else:
                print("❌ Nessun file trovato per il mese selezionato")
                print("Verificare:")
                print("  - Connessione a SharePoint")
                print("  - Permessi di accesso")
                print("  - Nome del mese corretto")
                
        except Exception as e:
            self.logger.error(f"Errore nel download file: {e}")
            print(f"❌ Errore: {e}")
    
    def _process_downloaded_files(self):
        """Elabora i file scaricati."""
        try:
            if not self.downloaded_files:
                print("❌ Nessun file scaricato. Eseguire prima il download.")
                return
            
            print("\n" + "=" * 60)
            print("⚙️ ELABORAZIONE FILE")
            print("=" * 60)
            
            all_groups = []
            file_stats = {}
            
            with ProgressBar(description="Elaborazione file") as progress:
                for i, file_path in enumerate(self.downloaded_files):
                    print(f"\n📄 Elaborazione: {file_path.name}")
                    
                    try:
                        groups = self.isin_service.process_excel_file(file_path)
                        all_groups.extend(groups)
                        
                        # Statistiche file
                        total_orders = sum(len(g.orders) for g in groups)
                        unique_isins = len(set(g.isin for g in groups))
                        
                        file_stats[file_path.name] = {
                            'groups': len(groups),
                            'orders': total_orders,
                            'unique_isins': unique_isins
                        }
                        
                        print(f"  ✅ Processati {len(groups)} gruppi ISIN")
                        print(f"  📊 {total_orders} ordini totali")
                        print(f"  🔑 {unique_isins} ISIN unici")
                        
                    except Exception as e:
                        self.logger.error(f"Errore elaborazione {file_path.name}: {e}")
                        print(f"  ❌ Errore: {e}")
                        file_stats[file_path.name] = {'error': str(e)}
                    
                    progress.update(i + 1, total=len(self.downloaded_files))
            
            self.processing_results = all_groups
            
            # Riepilogo finale
            print(f"\n✅ ELABORAZIONE COMPLETATA:")
            print(f"  📁 File processati: {len([f for f in file_stats if 'error' not in file_stats[f]])}")
            print(f"  📁 File con errori: {len([f for f in file_stats if 'error' in file_stats[f]])}")
            print(f"  📊 Gruppi ISIN totali: {len(all_groups)}")
            print(f"  🔑 ISIN unici: {len(set(g.isin for g in all_groups))}")
            print(f"  📋 Ordini totali: {sum(len(g.orders) for g in all_groups)}")
            
        except Exception as e:
            self.logger.error(f"Errore nell'elaborazione file: {e}")
            print(f"❌ Errore: {e}")
    
    def _run_quality_controls(self):
        """Esegue i controlli di qualità."""
        try:
            if not self.processing_results:
                print("❌ Nessun dato elaborato. Eseguire prima l'elaborazione.")
                return
            
            print("\n" + "=" * 60)
            print("🔍 CONTROLLI DI QUALITÀ")
            print("=" * 60)
            
            if not self.config.controlli_config.isin_api_url:
                print("⚠️  URL API non configurato - solo controlli base disponibili")
                print("Configurare URL API per controlli completi ISIN")
            
            print(f"🔍 Controllo {len(self.processing_results)} gruppi ISIN...")
            print("📋 Controlli attivi:")
            for i, controllo in enumerate(self.config.controlli_config.controlli_attivi, 1):
                status = "✅" if self.config.controlli_config.isin_api_url else "⚠️"
                print(f"  {status} Controllo {i}: {controllo}")
            
            with ProgressBar(description="Controlli qualità") as progress:
                control_results = self.validation_service.validate_isin_groups(
                    self.processing_results
                )
            
            # Analizza risultati
            total_passed = sum(r.controlli_passed for r in control_results)
            total_failed = sum(r.controlli_failed for r in control_results)
            total_controls = total_passed + total_failed
            
            print(f"\n✅ CONTROLLI COMPLETATI:")
            print(f"  📊 Controlli totali eseguiti: {total_controls}")
            print(f"  ✅ Controlli superati: {total_passed}")
            print(f"  ❌ Controlli falliti: {total_failed}")
            
            if total_controls > 0:
                success_rate = (total_passed / total_controls) * 100
                print(f"  📈 Tasso di successo: {success_rate:.1f}%")
            
            # Mostra dettagli errori se presenti
            failed_results = [r for r in control_results if r.controlli_failed > 0]
            if failed_results:
                print(f"\n⚠️  {len(failed_results)} ISIN con problemi:")
                for i, result in enumerate(failed_results[:10], 1):  # Mostra primi 10
                    print(f"  {i:2d}. {result.isin}: {result.controlli_failed} controlli falliti")
                    if result.business_rule_violations:
                        for violation in result.business_rule_violations[:2]:  # Prime 2 violazioni
                            print(f"      • {violation}")
                
                if len(failed_results) > 10:
                    print(f"  ... e altri {len(failed_results) - 10} ISIN con problemi")
                
                print(f"\n💡 RACCOMANDAZIONI:")
                print(f"  • Verificare anagrafica strumenti per ISIN non censiti")
                print(f"  • Controllare formato dati nei file Excel")
                print(f"  • Validare configurazione API se necessario")
            else:
                print(f"\n🎉 Tutti i controlli sono stati superati con successo!")
            
        except Exception as e:
            self.logger.error(f"Errore nei controlli di qualità: {e}")
            print(f"❌ Errore: {e}")
    
    def _export_results(self):
        """Esporta i risultati."""
        try:
            if not self.processing_results:
                print("❌ Nessun dato da esportare. Eseguire prima l'elaborazione.")
                return
            
            print("\n" + "=" * 60)
            print("📊 ESPORTAZIONE RISULTATI")
            print("=" * 60)
            
            config = ProcessingConfig(
                input_folder=Path("downloads"),
                output_folder=Path("output"),
                month=self.current_month or "UNKNOWN"
            )
            
            print(f"📝 Generazione report Excel per {self.current_month}...")
            print(f"📂 Cartella di output: {config.output_folder}")
            
            with ProgressBar(description="Esportazione") as progress:
                output_file = self.export_service.export_groups_to_excel(
                    self.processing_results,
                    config,
                    progress_callback=progress.update
                )
            
            if output_file and output_file.exists():
                size_kb = output_file.stat().st_size / 1024
                print(f"\n✅ Report esportato con successo:")
                print(f"  📄 File: {output_file}")
                print(f"  📊 Dimensione: {size_kb:.1f} KB")
                print(f"  📅 Mese: {self.current_month}")
                print(f"  📋 Gruppi ISIN: {len(self.processing_results)}")
                
                # Suggerimenti per l'uso
                print(f"\n💡 PROSSIMI PASSI:")
                print(f"  • Aprire il file Excel per revisione")
                print(f"  • Verificare i controlli marcati con 'X'")
                print(f"  • Procedere con il processo aziendale")
                
            else:
                print("❌ Errore nell'esportazione del report")
                
        except Exception as e:
            self.logger.error(f"Errore nell'esportazione: {e}")
            print(f"❌ Errore: {e}")
    
    def _show_validation_cache_stats(self):
        """Mostra statistiche della cache di validazione."""
        try:
            print("\n" + "=" * 60)
            print("💾 STATISTICHE CACHE VALIDAZIONE")
            print("=" * 60)
            
            stats = self.validation_service.get_cache_stats()
            
            print(f"📊 STATISTICHE GENERALI:")
            print(f"  🔑 ISIN in cache: {stats['total_cached_isins']}")
            print(f"  ⏰ Età cache: {stats['cache_age_hours']:.1f} ore")
            print(f"  ✅ Cache valida: {'Sì' if stats['cache_valid'] else 'No'}")
            
            if stats['cached_isins']:
                print(f"\n📋 ISIN IN CACHE (primi 15):")
                for i, isin in enumerate(stats['cached_isins'][:15], 1):
                    print(f"  {i:2d}. {isin}")
                    
                if len(stats['cached_isins']) > 15:
                    remaining = len(stats['cached_isins']) - 15
                    print(f"  ... e altri {remaining} ISIN")
            else:
                print(f"\n📭 Cache vuota")
            
            # Efficienza cache
            if stats['total_cached_isins'] > 0:
                print(f"\n📈 EFFICIENZA:")
                print(f"  • Cache riduce chiamate API ripetitive")
                print(f"  • Validità: 24 ore dall'ultimo aggiornamento")
                print(f"  • Pulizia automatica alla scadenza")
            
        except Exception as e:
            self.logger.error(f"Errore visualizzazione statistiche cache: {e}")
            print(f"❌ Errore: {e}")
    
    def _clear_validation_cache(self):
        """Pulisce la cache di validazione."""
        try:
            print("\n" + "=" * 60)
            print("🗑️ PULIZIA CACHE VALIDAZIONE")
            print("=" * 60)
            
            stats = self.validation_service.get_cache_stats()
            
            if stats['total_cached_isins'] == 0:
                print("📭 Cache già vuota")
                return
            
            print(f"📊 Cache corrente:")
            print(f"  🔑 ISIN memorizzati: {stats['total_cached_isins']}")
            print(f"  ⏰ Età: {stats['cache_age_hours']:.1f} ore")
            
            print(f"\n⚠️  ATTENZIONE:")
            print(f"  La pulizia cancellerà tutti i dati memorizzati")
            print(f"  Le prossime validazioni richiederanno chiamate API")
            
            confirm = input("\n➤ Sei sicuro di voler pulire la cache? (s/N): ").strip().lower()
            
            if confirm == 's':
                self.validation_service.clear_cache()
                print("✅ Cache pulita con successo")
                print("💡 La cache verrà ripopolata alle prossime validazioni")
            else:
                print("❌ Operazione annullata")
                
        except Exception as e:
            self.logger.error(f"Errore pulizia cache: {e}")
            print(f"❌ Errore: {e}")
    
    def _show_processing_summary(self):
        """Mostra un riepilogo dettagliato dell'elaborazione corrente."""
        try:
            print("\n" + "=" * 70)
            print("📋 RIEPILOGO ELABORAZIONE CON-412")
            print("=" * 70)
            
            # Configurazione
            print("🔧 CONFIGURAZIONE:")
            print(f"  📅 Mese selezionato: {self.current_month or 'Nessuno'}")
            print(f"  🌐 SharePoint: {self.config.sharepoint_config.site_url}")
            print(f"  📁 Cartella: {self.config.sharepoint_config.folder_path}")
            api_status = "✅ Configurata" if self.config.controlli_config.isin_api_url else "❌ Non configurata"
            print(f"  🔗 API ISIN: {api_status}")
            
            # File scaricati
            print(f"\n📥 FILE SCARICATI:")
            if self.downloaded_files:
                total_size = sum(f.stat().st_size for f in self.downloaded_files) / (1024 * 1024)
                print(f"  📊 Quantità: {len(self.downloaded_files)} file")
                print(f"  📦 Dimensione totale: {total_size:.1f} MB")
                for file_path in self.downloaded_files:
                    size_mb = file_path.stat().st_size / (1024 * 1024)
                    print(f"    📄 {file_path.name} ({size_mb:.1f} MB)")
            else:
                print("  📭 Nessun file scaricato")
            
            # Dati elaborati
            print(f"\n⚙️ DATI ELABORATI:")
            if self.processing_results:
                total_orders = sum(len(g.orders) for g in self.processing_results)
                unique_isins = len(set(g.isin for g in self.processing_results))
                
                print(f"  📊 Gruppi ISIN: {len(self.processing_results)}")
                print(f"  🔑 ISIN unici: {unique_isins}")
                print(f"  📋 Ordini totali: {total_orders}")
                
                # Statistiche per ISIN
                isin_stats = {}
                for group in self.processing_results:
                    if group.isin not in isin_stats:
                        isin_stats[group.isin] = 0
                    isin_stats[group.isin] += len(group.orders)
                
                # Top 5 ISIN per numero di ordini
                top_isins = sorted(isin_stats.items(), key=lambda x: x[1], reverse=True)[:5]
                if top_isins:
                    print(f"\n  🏆 TOP 5 ISIN per numero di ordini:")
                    for i, (isin, count) in enumerate(top_isins, 1):
                        print(f"    {i}. {isin}: {count} ordini")
            else:
                print("  📭 Nessun dato elaborato")
            
            # Cache validazione
            cache_stats = self.validation_service.get_cache_stats()
            print(f"\n💾 CACHE VALIDAZIONE:")
            print(f"  🔑 ISIN memorizzati: {cache_stats['total_cached_isins']}")
            print(f"  ⏰ Età cache: {cache_stats['cache_age_hours']:.1f} ore")
            print(f"  ✅ Stato: {'Valida' if cache_stats['cache_valid'] else 'Scaduta'}")
            
            # Prossimi passi suggeriti
            print(f"\n💡 PROSSIMI PASSI SUGGERITI:")
            if not self.current_month:
                print("  1. 📅 Selezionare un mese")
            elif not self.downloaded_files:
                print("  1. 📥 Scaricare i file dal SharePoint")
            elif not self.processing_results:
                print("  1. ⚙️ Elaborare i file scaricati")
            elif not self.config.controlli_config.isin_api_url:
                print("  1. 🔧 Configurare URL API per controlli completi")
                print("  2. 🔍 Eseguire controlli di qualità")
            else:
                print("  1. 🔍 Eseguire controlli di qualità")
                print("  2. 📊 Esportare risultati")
            
        except Exception as e:
            self.logger.error(f"Errore visualizzazione riepilogo: {e}")
            print(f"❌ Errore: {e}")


def main():
    """Funzione principale per avviare l'interfaccia CON-412."""
    try:
        # Configurazione logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        
        # Avvia interfaccia
        ui = CON412ConsoleUI()
        ui.run()
        
    except Exception as e:
        print(f"❌ Errore critico: {e}")
        logging.error(f"Errore critico nell'applicazione: {e}")


if __name__ == "__main__":
    main()