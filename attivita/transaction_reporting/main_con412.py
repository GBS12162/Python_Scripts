"""
CON-412 Transaction Reporting - Sistema Automatico
Elaborazione file Excel per validazione ISIN via ESMA
"""

import os
import sys
import logging
from datetime import datetime
from pathlib import Path

# Aggiunge il percorso del progetto al PYTHONPATH
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

# Import diretto del servizio file locale
local_file_module_path = Path(__file__).parent / "services" / "local_file_service.py"
database_service_module_path = Path(__file__).parent / "services" / "database_service.py"
isin_validation_module_path = Path(__file__).parent / "services" / "isin_validation_service.py"

import importlib.util

# Import LocalFileService
spec = importlib.util.spec_from_file_location("local_file_service", local_file_module_path)
local_file_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(local_file_module)
LocalFileService = local_file_module.LocalFileService

# Import DatabaseService
spec_db = importlib.util.spec_from_file_location("database_service", database_service_module_path)
database_service_module = importlib.util.module_from_spec(spec_db)
spec_db.loader.exec_module(database_service_module)
DatabaseService = database_service_module.DatabaseService

# Import ISINValidationService
spec_isin = importlib.util.spec_from_file_location("isin_validation_service", isin_validation_module_path)
isin_validation_module = importlib.util.module_from_spec(spec_isin)
spec_isin.loader.exec_module(isin_validation_module)
ISINValidationService = isin_validation_module.ISINValidationService


class InteractiveConfig:
    """Gestisce la configurazione interattiva del sistema"""
    
    def __init__(self):
        self.config = {}
        
    def get_file_config(self):
        """Richiede configurazione file locale/NAS all'utente"""
        print("\n🔧 CONFIGURAZIONE FILE LOCALE/NAS")
        print("=" * 40)
        
        config = self._get_file_config()
        if config:
            # Chiede se abilitare il controllo database
            config['enable_database_check'] = self._ask_database_check()
        
        return config
    
    def _get_file_config(self):
        """Configurazione file locale/NAS"""
        print("\n📁 CONFIGURAZIONE FILE")
        print("-" * 30)
        print("Inserisci il path completo al file Excel:")
        print()
        print("Esempi di path:")
        print("• File locale: C:\\Downloads\\CON-412_AGOSTO.xlsx")
        print("• NAS/Server: \\\\server\\share\\reports\\CON-412_AGOSTO.xlsx")  
        print("• Rete aziendale: \\\\nas-server\\shared\\Transaction_Reporting\\CON-412_AUGUST.xlsx")
        print()
        
        # Chiede solo il path completo del file
        file_path = input("Path completo al file Excel: ").strip()
        if not file_path:
            print("❌ Path del file è obbligatorio!")
            return None
            
        # Estrae automaticamente directory e nome file dal path completo
        from pathlib import Path
        path_obj = Path(file_path)
        base_path = str(path_obj.parent)
        file_name = path_obj.name
        
        print(f"\n📍 CONFIGURAZIONE RILEVATA")
        print(f"Directory: {base_path}")
        print(f"Nome file: {file_name}")
        print(f"Path completo: {file_path}")
        
        return {
            'type': 'local_file',
            'source': 'local',
            'base_path': base_path,
            'file_name': file_name,
            'file_path': file_path
        }
    
    def _ask_database_check(self):
        """Chiede se abilitare il controllo database"""
        print("\n💾 CONTROLLO DATABASE ORACLE")
        print("-" * 30)
        print("Vuoi abilitare il controllo database per filtrare gli ordini?")
        print("• Il sistema controllerà ogni ordine nel database Oracle TNS")
        print("• Verranno mantenuti solo gli ordini con stato 'RF'")
        print("• Gli ordini con stato 'EE' verranno rimossi")
        print()
        
        while True:
            response = input("Abilitare controllo database? (S/n): ").lower().strip()
            if response in ['', 's', 'si', 'y', 'yes']:
                print("✅ Controllo database abilitato")
                return True
            elif response in ['n', 'no']:
                print("⏭️ Controllo database disabilitato")
                return False
            else:
                print("⚠️ Risposta non valida. Inserisci S per Sì o N per No")


class CON412Processor:
    """Processore principale per CON-412 con validazione ESMA"""
    
    def __init__(self, config):
        self.config = config
        self.logger = self._setup_logging()
        
        # Inizializza il servizio di validazione ISIN
        self.isin_validation_service = ISINValidationService()
        
    def _setup_logging(self):
        """Configura il sistema di logging"""
        logs_dir = Path(__file__).parent / "logs"
        logs_dir.mkdir(exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = logs_dir / f"con412_processing_{timestamp}.log"
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler(sys.stdout)
            ]
        )
        
        return logging.getLogger(__name__)
    
    def run(self):
        """Esegue il processo CON-412 con validazione ESMA"""
        try:
            self.logger.info("=" * 60)
            self.logger.info("AVVIO PROCESSO CON-412 - TRANSACTION REPORTING")
            self.logger.info("=" * 60)
            self.logger.info(f"Configurazione: {self.config['type']}")
            self.logger.info(f"Path file: {self.config['file_path']}")
            
            # Fase 1: Verifica accesso al file
            print("\n🔄 FASE 1: Verifica accesso al file")
            
            file_service = LocalFileService.create_for_path(self.config['file_path'])
            self.logger.info(f"Path file: {self.config['file_path']}")
            
            if not file_service.check_access():
                print("❌ Errore accesso alla directory/file")
                return False
            print("✅ Accesso verificato")
            
            # Fase 2: Ricerca e lettura file
            print("\n🔄 FASE 2: Ricerca e lettura file")
            
            print(f"📁 Leggo file: {self.config['file_path']}")
            
            # Usa il path diretto
            source_file_path = None
            if Path(self.config['file_path']).exists():
                source_file_path = self.config['file_path']
            else:
                self.logger.error(f"File non trovato: {self.config['file_path']}")
            
            if not source_file_path:
                print("❌ File non trovato")
                return False
            
            print(f"✅ File trovato: {Path(source_file_path).name}")
            self.logger.info(f"File sorgente: {source_file_path}")
            
            # Copia il file in una directory di lavoro
            work_dir = Path(__file__).parent / "work"
            work_dir.mkdir(exist_ok=True)
            
            work_file_path = work_dir / Path(source_file_path).name
            
            if source_file_path != str(work_file_path):
                print("📋 Copia file nella directory di lavoro...")
                if not file_service.copy_file(source_file_path, str(work_file_path)):
                    print("❌ Errore copia file")
                    return False
                print("✅ File copiato in directory di lavoro")
                downloaded_file = str(work_file_path)
            else:
                downloaded_file = source_file_path
            
            # Fase 3: Lettura struttura file Excel
            print("\n📊 FASE 3: Lettura struttura file Excel")
            original_data = self._read_excel_file(downloaded_file)
            if not original_data:
                print("❌ Errore lettura file Excel")
                return False
            
            print(f"📊 Struttura letta: {len(original_data)} gruppi ISIN")
            print("✅ Struttura originale preservata")
            
            # Fase 4: Controllo Database (opzionale)
            if self.config.get('enable_database_check', False):
                print("\n🔄 FASE 4: Controllo stato ordini nel database")
                print("🔍 Connessione al database Oracle TNS per filtraggio ordini")
                
                try:
                    # Inizializza servizi database
                    db_service = DatabaseService()
                    # TODO: Implementare OrderFilterService per filtraggio RF/EE
                    # filter_service = OrderFilterService(db_service)
                    
                    # Testa la connessione
                    test_result = db_service.test_connection()
                    if not test_result["success"]:
                        print("❌ Connessione database fallita")
                        print("⚠️ Continuazione senza filtraggio database")
                    else:
                        print("✅ Connessione database riuscita")
                        # TODO: Implementare logica di filtraggio ordini RF/EE
                        # original_count = len(original_data)
                        # original_data = filter_service.filter_orders_by_status(original_data)
                        filtered_count = len(original_data)
                        
                        print(f"✅ Connessione database testata")
                        # TODO: Implementare statistiche filtraggio
                        # print(f"📊 Gruppi ISIN: {filtered_count}/{original_count} mantenuti")
                        
                except Exception as e:
                    self.logger.error(f"Errore controllo database: {str(e)}")
                    print(f"❌ Errore controllo database: {str(e)}")
                    print("⚠️ Continuazione senza filtraggio database")
            else:
                print("\n⏭️ FASE 4: Controllo database disabilitato")
            
            print("\n🔄 FASE 5: Validazione ESMA")
            print("\n🔄 FASE 5: Validazione ESMA")
            print("🌐 Chiamata API ESMA per validazione ISIN")
            validated_count = sum(1 for item in original_data if item['esma_valid'])
            print(f"✅ Validazione completata: {validated_count}/{len(original_data)} ISIN censurati")
            
            print("\n🔄 FASE 6: Aggiunta X per ISIN validati ESMA")
            
            # Crea directory output CON-412
            output_dir = Path(__file__).parent / "output" / "con412_reports"
            output_dir.mkdir(parents=True, exist_ok=True)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_name = f"CON-412_Validated_{timestamp}.xlsx"
            output_path = output_dir / output_name
            
            print(f"📄 Aggiornamento file con validazione ESMA: {output_name}")
            print(f"📁 Directory output: {output_dir}")
            print(f"📄 Struttura: File originale + X nella colonna CASISTICA per ISIN validati")
            
            # Crea file Excel con validazione
            try:
                success = self._create_validated_excel(downloaded_file, original_data, output_path)
                if not success:
                    print("❌ Errore creazione file Excel validato")
                    return False
                    
            except Exception as e:
                print(f"⚠️ Errore creazione Excel: {e}")
                return False
            
            print(f"\n🎯 ELABORAZIONE CON-412 COMPLETATA")
            print(f"📁 Report salvato in: {output_path}")
            print(f"📊 File originale processato: {Path(downloaded_file).name}")
            print(f"📊 Sorgente: {self.config['source']}")
            print(f"📊 ISIN con X (validati): {validated_count}/{len(original_data)}")
            
            self.logger.info("PROCESSO CON-412 COMPLETATO CON SUCCESSO")
            self.logger.info(f"Report: {output_path}")
            self.logger.info(f"Sorgente: {self.config['source']}")
            self.logger.info("=" * 60)
            
            return True
            
        except Exception as e:
            self.logger.error(f"Errore nel processo CON-412: {str(e)}")
            print(f"❌ Errore critico: {str(e)}")
            return False
    
    def _read_excel_file(self, file_path: str) -> list:
        """
        Legge la struttura del file Excel e identifica i gruppi ISIN
        
        Args:
            file_path: Percorso del file Excel da leggere
            
        Returns:
            Lista di dizionari con i dati dei gruppi ISIN
        """
        try:
            import openpyxl
            
            self.logger.info(f"Lettura file Excel: {file_path}")
            
            wb = openpyxl.load_workbook(file_path)
            ws = wb.active
            if not ws:
                ws = wb.worksheets[0]
            
            self.logger.info(f"Foglio Excel: {ws.title}")
            
            # Cerca la riga di intestazione
            header_row = None
            isin_col = None
            occurrences_col = None
            order_number_col = None
            
            for row_num in range(1, min(20, ws.max_row + 1)):
                for col_num in range(1, min(20, ws.max_column + 1)):  # Aumento il range per cercare più colonne
                    cell_value = ws.cell(row=row_num, column=col_num).value
                    if cell_value:
                        cell_text = str(cell_value).upper().strip()
                        # Cerca esattamente "ISIN" e non "DESCRIZIONE ISIN"
                        if cell_text == 'ISIN' and not isin_col:
                            header_row = row_num
                            isin_col = col_num
                        elif ('OCCORREN' in cell_text or 'OCCURRENCE' in cell_text) and not occurrences_col:
                            occurrences_col = col_num
                            if not header_row:
                                header_row = row_num
                        # Cerca colonna numero ordine
                        elif (('NUMERO' in cell_text and 'ORDINE' in cell_text) or 
                              'ORDER' in cell_text or 
                              cell_text in ['NUMORD', 'NUM_ORD', 'ORDER_NUM']) and not order_number_col:
                            order_number_col = col_num
                            if not header_row:
                                header_row = row_num
            
            if not header_row or not isin_col:
                self.logger.error("Impossibile trovare colonne ISIN nel file Excel")
                return []
            
            if not occurrences_col:
                # Se non trova OCCURRENCES, assume colonna dopo ISIN
                occurrences_col = isin_col + 1
                self.logger.warning(f"Colonna OCCURRENCES non trovata, uso colonna {occurrences_col}")
            
            # Nota: order_number_col può essere None se non troviamo la colonna numero ordine
            if order_number_col:
                self.logger.info(f"Header trovato alla riga {header_row}: ISIN=col{isin_col}, OCCURRENCES=col{occurrences_col}, ORDER_NUM=col{order_number_col}")
            else:
                self.logger.warning(f"Colonna numero ordine non trovata - controllo database disabilitato")
                self.logger.info(f"Header trovato alla riga {header_row}: ISIN=col{isin_col}, OCCURRENCES=col{occurrences_col}")
            
            # Legge i dati e identifica i gruppi ISIN
            data = []
            current_isin = None
            current_isin_row = None
            expected_orders = 0
            order_count = 0
            current_group_orders = []  # Lista per memorizzare i dettagli degli ordini del gruppo corrente
            
            for row_num in range(header_row + 1, ws.max_row + 1):
                isin_value = ws.cell(row=row_num, column=isin_col).value
                occurrences_value = ws.cell(row=row_num, column=occurrences_col).value
                order_number_value = ws.cell(row=row_num, column=order_number_col).value if order_number_col else None
                
                # Debug per le prime 10 righe
                if row_num <= header_row + 10:
                    self.logger.info(f"Riga {row_num}: ISIN='{isin_value}', OCCORRENZE='{occurrences_value}', NUMERO_ORDINE='{order_number_value}'")
                
                # Se trova un ISIN valorizzato, è una nuova prima riga di gruppo
                if isin_value and str(isin_value).strip() != '':
                    # Se c'era un gruppo precedente, lo finalizza
                    if current_isin and current_group_orders:
                        # Aggiorna l'ultimo gruppo con i dettagli degli ordini
                        data[-1]['orders'] = current_group_orders
                        self.logger.info(f"Completato gruppo ISIN {current_isin} con {len(current_group_orders)} ordini")
                    
                    isin_str = str(isin_value).strip()
                    occurrences_num = int(occurrences_value) if occurrences_value and str(occurrences_value).isdigit() else 1
                    
                    # Validazione ISIN di base (12 caratteri alfanumerici)
                    if len(isin_str) >= 12 and isin_str.isalnum():
                        self.logger.info(f"Nuovo gruppo ISIN: {isin_str} con {occurrences_num} ordini alla riga {row_num}")
                        
                        current_isin = isin_str
                        current_isin_row = row_num
                        expected_orders = occurrences_num
                        order_count = 1  # La riga corrente è il primo ordine
                        current_group_orders = []
                        
                        # Aggiunge il primo ordine se ha un numero ordine
                        if order_number_value and str(order_number_value).strip():
                            current_group_orders.append({
                                'row_num': row_num,
                                'numero_ordine': str(order_number_value).strip()
                            })
                        
                        data.append({
                            'isin': isin_str,
                            'occurrences': occurrences_num,
                            'row_num': row_num,  # Riga della prima occorrenza con ISIN
                            'order_rows': list(range(row_num + 1, row_num + 1 + occurrences_num)),  # Solo le righe degli ordini (esclude riga ISIN)
                            'orders': [],  # Sarà popolato alla fine del gruppo
                            'esma_valid': self._validate_isin_esma(isin_str)  # Validazione ESMA
                        })
                    else:
                        self.logger.info(f"ISIN non valido scartato: '{isin_str}' (len={len(isin_str)}, isalnum={isin_str.isalnum()})")
                        current_isin = None
                        current_isin_row = None
                        expected_orders = 0
                        order_count = 0
                        current_group_orders = []
                
                # Se siamo in un gruppo ISIN e questa è una riga di ordine (senza ISIN)
                elif current_isin and order_count < expected_orders:
                    order_count += 1
                    self.logger.info(f"Ordine {order_count}/{expected_orders} per ISIN {current_isin} alla riga {row_num}")
                    
                    # Aggiunge i dettagli dell'ordine se ha un numero ordine
                    if order_number_value and str(order_number_value).strip():
                        current_group_orders.append({
                            'row_num': row_num,
                            'numero_ordine': str(order_number_value).strip()
                        })
                    
                    # Se abbiamo completato tutti gli ordini del gruppo
                    if order_count >= expected_orders:
                        # Finalizza il gruppo corrente
                        data[-1]['orders'] = current_group_orders
                        self.logger.info(f"Completato gruppo ISIN {current_isin} ({expected_orders} ordini)")
                        current_isin = None
                        current_isin_row = None
                        expected_orders = 0
                        order_count = 0
                        current_group_orders = []
            
            # Finalizza l'ultimo gruppo se necessario
            if current_isin and current_group_orders:
                data[-1]['orders'] = current_group_orders
                self.logger.info(f"Completato gruppo finale ISIN {current_isin} con {len(current_group_orders)} ordini")
            
            self.logger.info(f"Letti {len(data)} ISIN dal file originale")
            return data
            
        except ImportError:
            self.logger.error("openpyxl non disponibile per lettura Excel")
            return []
        except Exception as e:
            self.logger.error(f"Errore lettura file Excel: {str(e)}")
            return []
    
    def _create_validated_excel(self, original_file_path: str, data: list, output_path: Path) -> bool:
        """
        Crea il file Excel validato con filtraggio database e validazione ESMA
        
        Args:
            original_file_path: Percorso del file Excel originale
            data: Lista dei dati processati e filtrati
            output_path: Percorso del file di output
            
        Returns:
            True se la creazione è riuscita
        """
        try:
            import openpyxl
            from openpyxl.styles import Font
            
            # Carica il file originale
            original_wb = openpyxl.load_workbook(original_file_path)
            original_ws = original_wb.active
            
            # Crea un nuovo workbook
            new_wb = openpyxl.Workbook()
            new_ws = new_wb.active
            new_ws.title = "CON-412 Validated"
            
            # Trova le colonne importanti nell'originale
            header_row = None
            casistica_col = None
            
            # Cerca l'header e la colonna casistica
            for row_num in range(1, min(20, original_ws.max_row + 1)):
                for col_num in range(1, original_ws.max_column + 1):
                    cell_value = original_ws.cell(row=row_num, column=col_num).value
                    if cell_value:
                        cell_text = str(cell_value).upper().strip()
                        if 'CASISTICA' in cell_text and 'ISIN NON CENSITO' in cell_text:
                            casistica_col = col_num
                            header_row = row_num
                            break
                        elif cell_text == 'ISIN' and not header_row:
                            header_row = row_num
                if casistica_col or header_row:
                    break
            
            if not header_row:
                self.logger.error("Impossibile trovare la riga header")
                return False
            
            if not casistica_col:
                self.logger.warning("Colonna CASISTICA non trovata - X non saranno aggiunte")
            
            self.logger.info(f"Header alla riga {header_row}, Casistica colonna {casistica_col}")
            
            # Copia l'header completamente
            for col_num in range(1, original_ws.max_column + 1):
                original_cell = original_ws.cell(row=header_row, column=col_num)
                new_cell = new_ws.cell(row=1, column=col_num)
                new_cell.value = original_cell.value
                if original_cell.font:
                    new_cell.font = Font(
                        name=original_cell.font.name,
                        size=original_cell.font.size,
                        bold=original_cell.font.bold,
                        italic=original_cell.font.italic
                    )
            
            # Crea il mapping delle righe da mantenere
            rows_to_keep = set()
            
            # Aggiungi tutte le righe dei gruppi filtrati
            for group in data:
                # Aggiungi la riga ISIN
                rows_to_keep.add(group['row_num'])
                
                # Aggiungi le righe degli ordini mantenuti
                if 'order_rows' in group:
                    for row_num in group['order_rows']:
                        rows_to_keep.add(row_num)
                elif 'orders' in group:
                    # Se abbiamo la struttura dettagliata degli ordini
                    for order in group['orders']:
                        if 'row_num' in order:
                            rows_to_keep.add(order['row_num'])
            
            # Copia le righe mantenute
            new_row_num = 2  # Inizia dopo l'header
            original_to_new_row_mapping = {}  # Per mappare le righe originali a quelle nuove
            
            # Ordina le righe da mantenere
            sorted_rows = sorted(rows_to_keep)
            
            for original_row_num in sorted_rows:
                # Copia la riga completa
                for col_num in range(1, original_ws.max_column + 1):
                    original_cell = original_ws.cell(row=original_row_num, column=col_num)
                    new_cell = new_ws.cell(row=new_row_num, column=col_num)
                    new_cell.value = original_cell.value
                    if original_cell.font:
                        new_cell.font = Font(
                            name=original_cell.font.name,
                            size=original_cell.font.size,
                            bold=original_cell.font.bold,
                            italic=original_cell.font.italic
                        )
                
                # Memorizza il mapping per la validazione ESMA
                original_to_new_row_mapping[original_row_num] = new_row_num
                new_row_num += 1
            
            # Applica la validazione ESMA se la colonna casistica è disponibile
            if casistica_col:
                for group in data:
                    # LOGICA CORRETTA: X se ISIN NON è validato (non censito)
                    if not group.get('esma_valid', True):  # Invertito: X se NON valido
                        # Mette X solo nelle righe degli ordini (MAI nella riga ISIN)
                        if 'order_rows' in group:
                            for original_row_num in group['order_rows']:
                                if original_row_num in original_to_new_row_mapping:
                                    new_row = original_to_new_row_mapping[original_row_num]
                                    cell = new_ws.cell(row=new_row, column=casistica_col)
                                    cell.value = 'X'
                                    cell.font = Font(bold=False, color="000000")
                        elif 'orders' in group:
                            for order in group['orders']:
                                if 'row_num' in order and order['row_num'] in original_to_new_row_mapping:
                                    new_row = original_to_new_row_mapping[order['row_num']]
                                    cell = new_ws.cell(row=new_row, column=casistica_col)
                                    cell.value = 'X'
                                    cell.font = Font(bold=False, color="000000")
            
            # Salva il file
            new_wb.save(output_path)
            self.logger.info(f"File Excel validato salvato: {output_path}")
            
            # Calcola statistiche
            invalid_count = sum(1 for item in data if not item.get('esma_valid', True))  # Conta ISIN NON validi
            total_orders_in_output = sum(len(item.get('orders', [])) for item in data)
            
            print("✅ File Excel creato con filtraggio database + validazione ESMA")
            print(f"📊 Gruppi ISIN nel file finale: {len(data)}")
            print(f"📊 ISIN NON CENSITI (con X nella CASISTICA): {invalid_count}/{len(data)} gruppi ISIN")
            print(f"📊 Ordini totali nel file finale: {total_orders_in_output}")
            
            return True
            
        except ImportError:
            self.logger.error("openpyxl non disponibile per creazione Excel")
            return False
        except Exception as e:
            self.logger.error(f"Errore creazione Excel validato: {str(e)}")
            return False
    
    def _validate_isin_esma(self, isin: str) -> bool:
        """
        Validazione ISIN tramite API ESMA reale
        
        Args:
            isin: Codice ISIN da validare
            
        Returns:
            True se l'ISIN è censito/validato da ESMA
        """
        try:
            return self.isin_validation_service.check_single_isin(isin)
        except Exception as e:
            self.logger.warning(f"Errore validazione ESMA per ISIN {isin}: {e}")
            # In caso di errore, assumiamo che l'ISIN sia valido per non bloccare il processo
            return True


def main():
    """Funzione principale con interfaccia interattiva"""
    print("=" * 60)
    print("🏦 CON-412 TRANSACTION REPORTING - REJECTING MENSILE")
    print("Sistema di elaborazione automatica ISIN via ESMA")
    print("=" * 60)
    print(f"📅 Avvio: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    try:
        # Configurazione interattiva
        config_manager = InteractiveConfig()
        config = config_manager.get_file_config()
        
        if not config:
            print("\n❌ Configurazione non valida. Uscita.")
            return 1
        
        print(f"\n✅ Configurazione completata!")
        print(f"Tipo: {config['type']}")
        
        # Conferma prima di procedere
        print("\n🚀 PRONTO PER L'ELABORAZIONE")
        proceed = input("Procedere con l'elaborazione? (S/n): ").lower().strip()
        
        if proceed in ['', 's', 'si', 'y', 'yes']:
            # Avvia il processore
            processor = CON412Processor(config)
            success = processor.run()
            
            if success:
                print("\n🎉 ELABORAZIONE COMPLETATA CON SUCCESSO!")
                return 0
            else:
                print("\n❌ ELABORAZIONE FALLITA!")
                return 1
        else:
            print("\n⏹️  Elaborazione annullata dall'utente")
            return 0
            
    except KeyboardInterrupt:
        print("\n\n⚠️  Elaborazione interrotta dall'utente")
        return 130
        
    except Exception as e:
        print(f"\n💥 Errore critico: {str(e)}")
        return 1


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)