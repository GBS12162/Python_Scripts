import pandas as pd
import os
import gzip
import chardet
import py7zr
import re
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from tqdm import tqdm
import sys
import multiprocessing as mp

# Costanti per i limiti di Excel
EXCEL_MAX_ROWS_PER_SHEET = 500000   # Limite conservativo per performance (500k righe)
EXCEL_MAX_ROWS_PER_FILE = 800000    # Ottimizzato per file ~20MB (sicuro per 25MB)
EXCEL_MAX_FILE_SIZE_MB = 25         # Limite dimensione file in MB per email
EXCEL_MAX_FILE_SIZE_BYTES = EXCEL_MAX_FILE_SIZE_MB * 1024 * 1024  # Conversione in bytes
EXCEL_TARGET_FILE_SIZE_MB = 22      # Target ottimale per rimanere nel range 20-25MB

def estimate_row_size(df, sample_size=1000):
    """
    Stima la dimensione media di una riga in bytes basandosi su un campione dei dati reali
    
    Args:
        df: DataFrame di cui stimare le dimensioni
        sample_size: Numero di righe da campionare per la stima
    
    Returns:
        float: Dimensione stimata per riga in bytes
    """
    if len(df) == 0:
        return 50  # Default fallback
    
    # Prendi un campione rappresentativo
    sample_size = min(sample_size, len(df))
    if sample_size < len(df):
        sample_df = df.sample(n=sample_size)
    else:
        sample_df = df
    
    # Calcola la dimensione stimata basandosi sui dati reali
    total_chars = 0
    for col in sample_df.columns:
        # Calcola caratteri per colonna includendo header
        col_chars = len(str(col)) + sample_df[col].astype(str).str.len().sum()
        total_chars += col_chars
    
    # Aggiungi overhead per formattazione Excel (separatori, formattazione, metadati)
    excel_overhead_factor = 1.8  # Excel ha overhead significativo
    avg_row_size = (total_chars / sample_size) * excel_overhead_factor
    
    # Assicurati che non sia troppo basso (minimo 25 bytes/riga)
    return max(avg_row_size, 25)

def calculate_optimal_limits(df, target_file_size_mb=EXCEL_TARGET_FILE_SIZE_MB):
    """
    Calcola dinamicamente i limiti ottimali per righe per sheet e per file
    basandosi sulla dimensione reale dei dati e sui target di dimensione
    
    Args:
        df: DataFrame di cui calcolare i limiti
        target_file_size_mb: Dimensione target del file in MB
    
    Returns:
        tuple: (max_rows_per_sheet, max_rows_per_file, estimated_bytes_per_row)
    """
    print("🧮 Calcolo dinamico dei limiti ottimali...")
    
    # Stima dimensione per riga basandosi sui dati reali
    estimated_bytes_per_row = estimate_row_size(df)
    print(f"📏 Dimensione stimata per riga: {estimated_bytes_per_row:.1f} bytes")
    
    # Calcola numero massimo di righe per raggiungere la dimensione target
    target_bytes = target_file_size_mb * 1024 * 1024
    max_rows_for_size = int(target_bytes / estimated_bytes_per_row)
    
    # Applica limiti di sicurezza Excel
    max_rows_per_sheet_safe = min(EXCEL_MAX_ROWS_PER_SHEET, max_rows_for_size // 2)  # Sheet più piccoli per stabilità
    max_rows_per_file_safe = min(max_rows_for_size, EXCEL_MAX_ROWS_PER_FILE)
    
    # Assicurati che i limiti siano sensati
    max_rows_per_sheet_safe = max(max_rows_per_sheet_safe, 50000)   # Minimo 50k righe/sheet
    max_rows_per_file_safe = max(max_rows_per_file_safe, 100000)     # Minimo 100k righe/file
    
    print(f"📊 Limiti calcolati dinamicamente:")
    print(f"   📄 Righe per sheet: {max_rows_per_sheet_safe:,}")
    print(f"   📚 Righe per file: {max_rows_per_file_safe:,}")
    print(f"   🎯 Dimensione target: {target_file_size_mb} MB")
    print(f"   📏 File stimato con {max_rows_per_file_safe:,} righe: {(max_rows_per_file_safe * estimated_bytes_per_row / 1024 / 1024):.1f} MB")
    
    return max_rows_per_sheet_safe, max_rows_per_file_safe, estimated_bytes_per_row

def set_console_title():
    """Imposta il titolo della finestra del terminale"""
    if os.name == 'nt':  # Windows
        os.system('title Look-up components - CDL')

def clear_console():
    """Pulisce la console"""
    os.system('cls' if os.name == 'nt' else 'clear')

def get_file_size_mb(file_path):
    """Ottiene la dimensione del file in MB"""
    try:
        size_bytes = os.path.getsize(file_path)
        size_mb = size_bytes / (1024 * 1024)
        return size_mb
    except:
        return 0

def check_file_size_limit(file_path, max_size_mb=EXCEL_MAX_FILE_SIZE_MB):
    """Controlla se il file supera il limite di dimensione"""
    size_mb = get_file_size_mb(file_path)
    return size_mb <= max_size_mb, size_mb

def safe_input(prompt, default=""):
    """Input sicuro che gestisce l'assenza di stdin (eseguibili noconsole)"""
    try:
        # Forza sempre l'uso di input standard se disponibile
        return input(prompt)
    except (EOFError, RuntimeError):
        # Se non c'è stdin disponibile (eseguibile noconsole), usa tkinter
        try:
            import tkinter as tk
            from tkinter import simpledialog
            
            root = tk.Tk()
            root.withdraw()  # Nasconde la finestra principale
            
            # Estrae il prompt pulito
            clean_prompt = prompt.replace("\n➤ ", "").replace("➤ ", "").replace("\n", " ").strip()
            
            result = simpledialog.askstring("Look-up Components CDL", clean_prompt)
            root.destroy()
            
            return result.strip() if result else default
        except:
            # Se anche tkinter fallisce, usa default o esce
            return default

def safe_pause(message):
    """Pausa sicura che non richiede input se non c'è console"""
    try:
        input(message)
    except (EOFError, RuntimeError):
        # Se non c'è stdin, salta la pausa
        pass

def show_header():
    """Mostra l'header dell'applicazione"""
    clear_console()
    print("=" * 60)
    print("🔍 LOOK-UP COMPONENTS - CDL")
    print("=" * 60)
    print("Strumento per il lookup automatico dei componenti Oracle")
    print("Versione 2.0 - Aggiornato 2025")
    print("=" * 60)

def rename_and_create_7z_archives(created_files, base_name):
    """
    Rinomina i file e crea archivi 7z raggruppando i file correlati
    
    Args:
        created_files: Lista dei percorsi dei file creati
        base_name: Nome base per gli archivi
    
    Returns:
        Lista degli archivi 7z creati
    """
    if not created_files:
        return []
    
    print(f"\n📦 Creazione archivi 7z...")
    
    # Raggruppa i file per parte
    file_groups = defaultdict(list)
    renamed_files = []
    
    for file_path in created_files:
        file_path = Path(file_path)
        file_name = file_path.name
        
        # Rinomina da "Parte X_partY" a "Parte X_Y"
        if "_part" in file_name:
            # Pattern: *Parte X_partY.xlsx -> *Parte X_Y.xlsx
            new_name = re.sub(r'_part(\d+)', r'_\1', file_name)
            new_path = file_path.parent / new_name
            
            # Rinomina il file
            try:
                file_path.rename(new_path)
                print(f"   📝 Rinominato: {file_name} → {new_name}")
                renamed_files.append(new_path)
                file_path = new_path
                file_name = new_name
            except Exception as e:
                print(f"   ⚠️  Errore rinomina {file_name}: {e}")
                renamed_files.append(file_path)
        else:
            renamed_files.append(file_path)
        
        # Identifica il gruppo (Parte X)
        match = re.search(r'Parte (\d+)', file_name)
        if match:
            parte_num = match.group(1)
            file_groups[f"Parte{parte_num}"].append(file_path)
        elif "Analisi_Fallimenti" in file_name or "Fallimenti" in file_name:
            # File di analisi fallimenti
            file_groups["Altri"].append(file_path)
        else:
            # File senza parti specifiche - raggruppalo come "Main"
            file_groups["Main"].append(file_path)
    
    # Crea archivi 7z per ogni gruppo
    created_archives = []
    base_dir = Path(created_files[0]).parent
    
    for group_name, group_files in file_groups.items():
        # Zippa solo i gruppi "Parte" e "Main" (non zippare i file di analisi fallimenti)
        # Il gruppo "Altri" (analisi fallimenti) non viene zippato
        if group_name == "Altri":
            print(f"   📄 File analisi fallimenti non zippato: {len(group_files)} file")
            for file_path in group_files:
                print(f"      📄 {file_path.name}")
            continue
            
        # Nome dell'archivio
        if group_name == "Main":
            archive_name = f"{base_name}.7z"
        else:
            archive_name = f"{base_name}_{group_name}.7z"
        
        archive_path = base_dir / archive_name
        
        try:
            with py7zr.SevenZipFile(archive_path, 'w') as archive:
                for file_path in group_files:
                    if file_path.exists():
                        # Aggiungi solo il nome del file nell'archivio (senza path)
                        archive.write(file_path, file_path.name)
            
            print(f"   📦 Creato: {archive_name} ({len(group_files)} file)")
            created_archives.append(archive_path)
            
            # Rimuovi i file originali dopo aver creato l'archivio
            for file_path in group_files:
                try:
                    file_path.unlink()
                except Exception as e:
                    print(f"   ⚠️  Errore rimozione {file_path.name}: {e}")
                    
        except Exception as e:
            print(f"   ❌ Errore creazione archivio {archive_name}: {e}")
    
    return created_archives
    print()

def load_components(file_path):
    """Carica i componenti dal file CSV con interfaccia semplificata"""
    print("📂 Caricamento database componenti...")
    
    # Prova diversi percorsi per trovare Components.csv
    possible_paths = [
        file_path,  # Percorso originale
        "Components.csv",  # Directory corrente
        os.path.join(os.path.dirname(sys.executable), "Components.csv"),  # Stessa cartella dell'exe
        os.path.join(os.getcwd(), "Components.csv"),  # Directory di lavoro corrente
    ]
    
    # Se siamo in un eseguibile PyInstaller, prova anche il percorso temporaneo
    if getattr(sys, 'frozen', False):
        # Siamo in un eseguibile PyInstaller
        bundle_dir = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
        possible_paths.insert(0, os.path.join(bundle_dir, "Components.csv"))
    
    actual_file_path = None
    for path in possible_paths:
        if os.path.exists(path):
            actual_file_path = path
            break
    
    if actual_file_path is None:
        print(f"❌ ERRORE: File Components.csv non trovato!")
        print("💡 Assicurati che Components.csv sia nella stessa cartella dell'eseguibile")
        print(f"   Percorsi cercati:")
        for path in possible_paths:
            print(f"   - {path}")
        return None, None
    
    try:
        # Carica il CSV
        df = pd.read_csv(actual_file_path, delimiter=';', encoding='utf-8')
        
        if 'Nome_Componente' not in df.columns or 'Prefisso' not in df.columns:
            print("❌ ERRORE: Il file CSV deve avere le colonne 'Nome_Componente' e 'Prefisso'")
            return None, None
        
        # Crea il dizionario di lookup ottimizzato
        lookup_dict = {}
        prefix_count = 0
        
        for _, row in df.iterrows():
            nome = row['Nome_Componente']
            prefissi = str(row['Prefisso']).split(';')
            
            for prefisso in prefissi:
                prefisso = prefisso.strip()
                if prefisso and prefisso != 'nan':
                    lookup_dict[prefisso] = nome
                    prefix_count += 1
        
        print(f"✅ Database caricato: {len(df)} componenti, {prefix_count} prefissi")
        print(f"📁 Percorso usato: {actual_file_path}")
        return df, lookup_dict
        
    except Exception as e:
        print(f"❌ ERRORE durante il caricamento: {str(e)}")
        return None, None

# Variabile globale per ricordare la scelta sui file di fallimenti per operazioni multiple
failed_file_choice_for_all = None

def ask_for_failed_lookup_file(failed_count, total_count, is_batch_operation=False, current_file=1, total_files=1):
    """
    Chiede all'utente se vuole creare il file dei fallimenti
    
    Args:
        failed_count: Numero di lookup falliti
        total_count: Numero totale di lookup
        is_batch_operation: Se True, è un'operazione su più file
        current_file: Numero del file corrente (per operazioni multiple)
        total_files: Numero totale di file (per operazioni multiple)
    """
    global failed_file_choice_for_all
    
    if failed_count == 0:
        return False
    
    # Se è un'operazione batch e l'utente ha già scelto per tutti
    if is_batch_operation and failed_file_choice_for_all is not None:
        return failed_file_choice_for_all
        
    print(f"\n⚠️  ATTENZIONE: {failed_count} su {total_count} lookup sono falliti")
    print(f"   Percentuale fallimenti: {(failed_count/total_count)*100:.1f}%")
    
    if is_batch_operation and total_files > 1:
        print(f"   📄 File corrente: {current_file}/{total_files}")
    
    print("\n🤔 Vuoi creare un file di analisi dei fallimenti?")
    print("   Questo file ti aiuterà a capire perché alcuni lookup non sono riusciti")
    
    if is_batch_operation and total_files > 1:
        print("\n💡 Opzioni disponibili:")
        print("   s  = Sì per questo file")
        print("   n  = No per questo file")
        print("   st = Sì per TUTTI i file rimanenti")
        print("   nt = No per TUTTI i file rimanenti")
        prompt = "\nCreare file fallimenti? (s/n/st/nt): "
    else:
        prompt = "\nCreare file fallimenti? (s/n): "
    
    while True:
        choice = safe_input(prompt).strip().lower()
        
        if choice in ['s', 'si', 'y', 'yes']:
            return True
        elif choice in ['n', 'no']:
            return False
        elif is_batch_operation and choice in ['st', 'sit', 'si tutti', 'yes all']:
            failed_file_choice_for_all = True
            print("✅ Applicato 'Sì' per tutti i file rimanenti")
            return True
        elif is_batch_operation and choice in ['nt', 'not', 'no tutti', 'no all']:
            failed_file_choice_for_all = False
            print("✅ Applicato 'No' per tutti i file rimanenti")
            return False
        else:
            if is_batch_operation and total_files > 1:
                print("❌ Rispondi con 's' (sì), 'n' (no), 'st' (sì tutti) o 'nt' (no tutti)")
            else:
                print("❌ Rispondi con 's' per sì o 'n' per no")

def reset_failed_file_choice():
    """Reset della scelta globale per i file di fallimenti (da usare all'inizio di una nuova operazione batch)"""
    global failed_file_choice_for_all
    failed_file_choice_for_all = None

def create_excel_with_split_support(df, output_path, failed_table_names=None, lookup_dict=None, table_name_col='TABLE_NAME'):
    """
    Crea file Excel gestendo automaticamente la suddivisione in più sheet o file
    quando si superano i limiti di Excel, con calcolo dinamico dei limiti ottimali
    
    Args:
        df: DataFrame principale
        output_path: Percorso del file di output  
        failed_table_names: Lista dei nomi delle tabelle con lookup fallito
        lookup_dict: Dizionario per il lookup
        table_name_col: Nome della colonna con i nomi delle tabelle
    
    Returns:
        List[Path]: Lista dei file creati
    """
    total_rows = len(df)
    print(f"📊 Righe totali da processare: {total_rows:,}")
    
    # Calcola limiti dinamici basati sui dati reali
    max_rows_per_sheet, max_rows_per_file, estimated_bytes_per_row = calculate_optimal_limits(df)
    
    created_files = []
    
    # Determina se serve suddivisione usando i limiti calcolati dinamicamente
    if total_rows <= max_rows_per_sheet:
        # Caso semplice: tutto in un file, un sheet
        print("📄 Creazione file singolo...")
        result = _create_single_excel_file(df, output_path, failed_table_names, lookup_dict, table_name_col, 
                                         max_rows_per_sheet, max_rows_per_file, estimated_bytes_per_row)
        created_files.extend(result if isinstance(result, list) else [result])
        
    elif total_rows <= max_rows_per_file:
        # Caso medio: un file, più sheet
        print(f"📑 Suddivisione in più sheet (massimo {max_rows_per_sheet:,} righe per sheet)...")
        result = _create_multi_sheet_excel_file(df, output_path, failed_table_names, lookup_dict, table_name_col,
                                               max_rows_per_sheet, max_rows_per_file, estimated_bytes_per_row)
        created_files.extend(result if isinstance(result, list) else [result])
        
    else:
        # Caso complesso: più file
        print(f"📚 Suddivisione in più file (massimo {max_rows_per_file:,} righe per file)...")
        created_files.extend(_create_multi_file_excel(df, output_path, failed_table_names, lookup_dict, table_name_col,
                                                     max_rows_per_sheet, max_rows_per_file, estimated_bytes_per_row))
    
    return created_files

def _create_multi_file_excel_size_limited(df, base_output_path, failed_table_names, lookup_dict, table_name_col, max_rows_per_chunk):
    """Crea più file Excel con limitazione di dimensione"""
    base_name = str(base_output_path).replace('.xlsx', '')
    files_created = []
    
    total_rows = len(df)
    num_chunks = (total_rows + max_rows_per_chunk - 1) // max_rows_per_chunk
    
    print(f"📊 Suddivisione in {num_chunks} file da max {max_rows_per_chunk} righe per rispettare limite di {EXCEL_MAX_FILE_SIZE_MB} MB")
    
    for i in range(num_chunks):
        start_idx = i * max_rows_per_chunk
        end_idx = min((i + 1) * max_rows_per_chunk, total_rows)
        chunk_df = df.iloc[start_idx:end_idx].copy()
        
        if num_chunks > 1:
            chunk_file = f"{base_name}_part{i+1}.xlsx"
        else:
            chunk_file = f"{base_name}.xlsx"
        
        print(f"📄 Creazione parte {i+1}/{num_chunks}: {os.path.basename(chunk_file)} ({len(chunk_df)} righe)")
        
        with pd.ExcelWriter(chunk_file, engine='xlsxwriter') as writer:
            chunk_df.to_excel(writer, sheet_name='Data', index=False)
            _apply_excel_formatting(writer, chunk_df, [])
        
        # Verifica dimensione
        file_ok, file_size = check_file_size_limit(chunk_file)
        if file_ok:
            print(f"✅ File creato con successo: {file_size:.1f} MB")
            files_created.append(chunk_file)
        else:
            print(f"❌ File ancora troppo grande: {file_size:.1f} MB")
            # Se ancora troppo grande, dividi ulteriormente
            os.remove(chunk_file)
            reduced_rows = max_rows_per_chunk // 2
            sub_files = _create_multi_file_excel_size_limited(chunk_df, chunk_file, [], lookup_dict, table_name_col, reduced_rows)
            files_created.extend(sub_files)
    
    # Crea analisi fallimenti separata
    if failed_table_names:
        analysis_files = create_failed_lookup_analysis_with_split(failed_table_names, lookup_dict, base_output_path)
        files_created.extend(analysis_files)
    
    return files_created

def _create_single_excel_file(df, output_path, failed_table_names, lookup_dict, table_name_col, 
                            max_rows_per_sheet, max_rows_per_file, estimated_bytes_per_row):
    """Crea un singolo file Excel con un sheet, controllando la dimensione con limiti dinamici"""
    print(f"💾 Creazione file Excel: {output_path}")
    
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        # Sheet principale
        df.to_excel(writer, sheet_name='Data', index=False)
        
        # Sheet analisi fallimenti solo se poche righe
        if failed_table_names:
            estimated_analysis_rows = len(failed_table_names) * 5
            if estimated_analysis_rows <= max_rows_per_sheet:
                failed_analysis_df = create_failed_lookup_analysis(failed_table_names, lookup_dict)
                failed_analysis_df.to_excel(writer, sheet_name='Lookup_Failed', index=False)
            else:
                print("⚠️  Troppe righe di analisi - sarà creato un file separato")
        
        _apply_excel_formatting(writer, df, failed_table_names)
    
    # Controlla la dimensione del file creato
    file_ok, file_size = check_file_size_limit(output_path)
    print(f"📏 Dimensione file: {file_size:.1f} MB")
    
    if not file_ok:
        print(f"⚠️  File troppo grande ({file_size:.1f} MB > {EXCEL_MAX_FILE_SIZE_MB} MB)")
        print("🔄 Ricreo con suddivisione automatica...")
        
        # Rimuovi il file troppo grande
        os.remove(output_path)
        
        # Calcola quante righe possiamo mettere per rispettare il limite di dimensione
        max_rows_for_size = EXCEL_MAX_FILE_SIZE_BYTES // int(estimated_bytes_per_row)
        max_rows_safe = min(max_rows_for_size, max_rows_per_sheet)
        
        # Ricrea con suddivisione forzata
        return _create_multi_file_excel_size_limited(df, output_path, failed_table_names, lookup_dict, 
                                                   table_name_col, max_rows_safe)
    
    # Crea file analisi separato se necessario
    analysis_files = []
    if failed_table_names:
        estimated_analysis_rows = len(failed_table_names) * 5
        if estimated_analysis_rows > max_rows_per_sheet:
            analysis_files = create_failed_lookup_analysis_with_split(failed_table_names, lookup_dict, output_path)
    
    return [output_path] + analysis_files

def _create_multi_sheet_excel_file(df, output_path, failed_table_names, lookup_dict, table_name_col,
                                  max_rows_per_sheet, max_rows_per_file, estimated_bytes_per_row):
    """Crea un file Excel con più sheet, controllando la dimensione con limiti dinamici"""
    print(f"💾 Creazione file Excel multi-sheet: {output_path}")
    
    # Calcola numero di sheet necessari usando i limiti dinamici
    num_sheets = (len(df) + max_rows_per_sheet - 1) // max_rows_per_sheet
    print(f"📑 Creazione di {num_sheets} sheet...")
    
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        # Suddividi DataFrame in chunk
        for i in range(num_sheets):
            start_idx = i * max_rows_per_sheet
            end_idx = min((i + 1) * max_rows_per_sheet, len(df))
            df_chunk = df.iloc[start_idx:end_idx]
            
            sheet_name = f'Data_Part{i+1}' if num_sheets > 1 else 'Data'
            print(f"   📄 Sheet '{sheet_name}': righe {start_idx+1:,} - {end_idx:,}")
            
            df_chunk.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Sheet analisi fallimenti solo se poche righe
        if failed_table_names:
            estimated_analysis_rows = len(failed_table_names) * 5
            if estimated_analysis_rows <= max_rows_per_sheet:
                failed_analysis_df = create_failed_lookup_analysis(failed_table_names, lookup_dict)
                failed_analysis_df.to_excel(writer, sheet_name='Lookup_Failed', index=False)
            else:
                print("⚠️  Troppe righe di analisi - sarà creato un file separato")
        
        _apply_excel_formatting(writer, df, failed_table_names, num_sheets)
    
    # Controlla la dimensione del file creato
    file_ok, file_size = check_file_size_limit(output_path)
    print(f"📏 Dimensione file: {file_size:.1f} MB")
    
    if not file_ok:
        print(f"⚠️  File troppo grande ({file_size:.1f} MB > {EXCEL_MAX_FILE_SIZE_MB} MB)")
        print("🔄 Ricreo con dimensioni ridotte...")
        
        # Rimuovi il file troppo grande
        os.remove(output_path)
        
        # Calcola quante righe possiamo mettere per rispettare il limite di dimensione
        max_rows_for_size = EXCEL_MAX_FILE_SIZE_BYTES // int(estimated_bytes_per_row)
        max_rows_safe = min(max_rows_for_size, max_rows_per_sheet)
        
        # Usa la funzione size-limited per creare con dimensioni ridotte
        return _create_multi_file_excel_size_limited(df, output_path, failed_table_names, lookup_dict, table_name_col, max_rows_safe)
    
    # Crea file analisi separato se necessario
    analysis_files = []
    if failed_table_names:
        estimated_analysis_rows = len(failed_table_names) * 5
        if estimated_analysis_rows > max_rows_per_sheet:
            analysis_files = create_failed_lookup_analysis_with_split(failed_table_names, lookup_dict, output_path)
    
    return [output_path] + analysis_files

def _create_multi_file_excel(df, output_path, failed_table_names, lookup_dict, table_name_col,
                           max_rows_per_sheet, max_rows_per_file, estimated_bytes_per_row):
    """Crea più file Excel con controllo dimensioni dinamiche"""
    # Calcola numero di file necessari usando i limiti dinamici
    num_files = (len(df) + max_rows_per_file - 1) // max_rows_per_file
    print(f"📚 Creazione di {num_files} file...")
    
    created_files = []
    base_path = Path(output_path)
    base_name = base_path.stem
    extension = base_path.suffix
    parent_dir = base_path.parent
    
    # Suddividi DataFrame in chunk per file
    for i in range(num_files):
        start_idx = i * max_rows_per_file
        end_idx = min((i + 1) * max_rows_per_file, len(df))
        df_chunk = df.iloc[start_idx:end_idx]
        
        # Nome file per questa parte
        if num_files > 1:
            file_name = f"{base_name} - Parte {i+1}{extension}"
        else:
            file_name = f"{base_name}{extension}"
        
        file_path = parent_dir / file_name
        print(f"   📄 File '{file_name}': righe {start_idx+1:,} - {end_idx:,}")
        
        # Per ogni file, usa direttamente la funzione size-limited
        # per evitare loop infiniti con file di grandi dimensioni
        max_rows_for_size = EXCEL_MAX_FILE_SIZE_BYTES // int(estimated_bytes_per_row)
        max_rows_safe = min(max_rows_for_size, max_rows_per_sheet)
        
        if len(df_chunk) > max_rows_safe:
            print(f"   ⚠️  Chunk troppo grande, suddivisione forzata...")
            file_results = _create_multi_file_excel_size_limited(df_chunk, file_path, None, lookup_dict, table_name_col, max_rows_safe)
        elif len(df_chunk) <= max_rows_per_sheet:
            file_results = _create_single_excel_file(df_chunk, file_path, None, lookup_dict, table_name_col,
                                                   max_rows_per_sheet, max_rows_per_file, estimated_bytes_per_row)
        else:
            # Usa la funzione size-limited anche per i multi-sheet
            file_results = _create_multi_file_excel_size_limited(df_chunk, file_path, None, lookup_dict, table_name_col, max_rows_per_sheet)
        
        # Aggiungi tutti i file creati
        if isinstance(file_results, list):
            created_files.extend(file_results)
        else:
            created_files.append(file_results)
    
    # Crea file separato per l'analisi fallimenti
    if failed_table_names:
        print(f"   🔍 Creazione analisi fallimenti separata...")
        analysis_files = create_failed_lookup_analysis_with_split(failed_table_names, lookup_dict, output_path)
        created_files.extend(analysis_files)
    
    return created_files

def _apply_excel_formatting(writer, df, failed_table_names, num_sheets=1):
    """Applica formattazione standard agli sheet Excel"""
    workbook = writer.book
    
    # Formati
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#D7E4BC',
        'border': 1
    })
    
    # Formatta tutti i sheet di dati
    for sheet_name in writer.sheets:
        if sheet_name.startswith('Data'):
            worksheet = writer.sheets[sheet_name]
            
            # Applica formattazione header
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Auto-adatta larghezza colonne
            for col_num, col in enumerate(df.columns):
                try:
                    max_len = max(
                        df[col].astype(str).map(len).max(),
                        len(str(col))
                    )
                    worksheet.set_column(col_num, col_num, min(max_len + 2, 50))
                except:
                    worksheet.set_column(col_num, col_num, 15)
    
    # Formatta sheet lookup falliti
    if 'Lookup_Failed' in writer.sheets:
        worksheet_failed = writer.sheets['Lookup_Failed']
        failed_header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#FFE6E6',
            'border': 1
        })
        
        # Auto-adatta colonne per sheet falliti
        try:
            for col_num in range(6):  # Numero approssimativo di colonne
                worksheet_failed.set_column(col_num, col_num, 20)
        except:
            pass

def detect_encoding(file_path):
    """
    Rileva l'encoding di un file con fallback robusti
    """
    try:
        with open(file_path, 'rb') as f:
            result = chardet.detect(f.read(10000))  # Leggi i primi 10KB per rilevare l'encoding
            
            # Estrai l'encoding con controlli di sicurezza
            detected_encoding = None
            if result and isinstance(result, dict):
                detected_encoding = result.get('encoding')
            
            # Se l'encoding rilevato è None o vuoto, usa fallback
            if not detected_encoding or detected_encoding is None:
                return 'utf-8'
            
            # Verifica che l'encoding sia valido
            try:
                'test'.encode(detected_encoding)
                return detected_encoding
            except (LookupError, TypeError):
                # Se l'encoding non è valido, usa utf-8
                return 'utf-8'
                
    except Exception as e:
        print(f"⚠️  Errore rilevamento encoding: {e}")
        return 'utf-8'  # Default fallback

def read_single_file_content(file_path):
    """
    Legge il contenuto di un singolo file senza logica multi-volume
    """
    # Converti il path in oggetto Path per gestire sia path locali che di rete
    path_obj = Path(file_path)
    
    # Controlla se il file esiste
    if not path_obj.exists():
        raise FileNotFoundError(f"Il file {file_path} non esiste o non è accessibile")
    
    # Ottieni la dimensione del file per la barra di progresso
    file_size = path_obj.stat().st_size
    
    # Leggi i primi bytes per determinare se è compresso
    with open(file_path, 'rb') as f:
        first_bytes = f.read(2)
    
    # Magic bytes per gzip: 0x1f 0x8b
    is_gzip_compressed = (len(first_bytes) >= 2 and 
                         first_bytes[0] == 0x1f and 
                         first_bytes[1] == 0x8b)
    
    print(f"📊 File: {file_size/1024/1024:.1f} MB")
    
    if is_gzip_compressed:
        print("🗜️  Decompressione file...")
        
        # Leggi file compresso con progresso
        with open(file_path, 'rb') as file_obj:
            with tqdm(total=file_size, unit='B', unit_scale=True, desc="Lettura") as pbar:
                # Leggi tutto il contenuto compresso
                compressed_data = b''
                chunk_size = 8192  # 8KB chunks
                
                while True:
                    chunk = file_obj.read(chunk_size)
                    if not chunk:
                        break
                    compressed_data += chunk
                    pbar.update(len(chunk))
        
        # Decomprimi i dati
        print("🔄 Decompressione...")
        with tqdm(desc="Decompressione", unit="bytes") as pbar:
            content = gzip.decompress(compressed_data).decode('utf-8', errors='ignore')
            pbar.update(len(content))
            
    else:
        print("📄 File non compresso rilevato")
        # File normale, rileva encoding
        encoding = detect_encoding(file_path)
        
        # Controllo di sicurezza per encoding
        if not encoding or encoding is None:
            encoding = 'utf-8'
            print("⚠️  Encoding non rilevato, uso UTF-8")
        else:
            print(f"🔤 Encoding rilevato: {encoding}")
        
        # Leggi con progresso
        try:
            with open(file_path, 'r', encoding=encoding, errors='ignore') as file_obj:
                with tqdm(total=file_size, unit='B', unit_scale=True, desc="Lettura") as pbar:
                    content = ""
                    chunk_size = 8192
                    
                    while True:
                        chunk = file_obj.read(chunk_size)
                        if not chunk:
                            break
                        content += chunk
                        # Gestisci encoding sicuro per il progresso
                        try:
                            chunk_bytes = len(chunk.encode(encoding, errors='ignore'))
                        except (TypeError, AttributeError):
                            chunk_bytes = len(chunk.encode('utf-8', errors='ignore'))
                        pbar.update(chunk_bytes)
        except UnicodeDecodeError:
            print("⚠️  Errore decodifica, ritento con UTF-8...")
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as file_obj:
                content = file_obj.read()
    
    print(f"✅ Lettura completata: {len(content)} caratteri, {len(content.splitlines())} righe")
    return content

def read_file_content_with_progress(file_path):
    """
    Legge il contenuto di un file con barra di progresso, gestendo file compressi .gz, encoding e file multi-volume
    Rileva automaticamente se il file è compresso guardando i magic bytes
    Gestisce file multi-volume (.001, .002, etc.) concatenandoli automaticamente
    
    Args:
        file_path (str): Percorso del file (può essere su NAS o locale)
    
    Returns:
        str: Contenuto del file (o file concatenati se multi-volume)
    """
    try:
        # Controlla se è un file multi-volume
        # Nuovo approccio: per i file multi-volume, elabora prima singolarmente
        if is_multivolume_file(file_path):
            print(f"🔗 File multi-volume rilevato: {Path(file_path).name}")
            
            # Prova prima a elaborare il file singolarmente
            print("📄 Tentativo elaborazione singola del file...")
            try:
                return read_single_file_content(file_path)
            except Exception as e:
                print(f"⚠️  Elaborazione singola fallita: {e}")
                print("🔄 Tentativo concatenazione multi-volume...")
                
                # Solo se fallisce, prova la concatenazione
                directory_path = Path(file_path).parent
                multivolume_files = find_multivolume_files(file_path, directory_path)
                
                if len(multivolume_files) > 1:
                    print(f"📚 Trovate {len(multivolume_files)} parti multi-volume")
                    return read_multivolume_content(multivolume_files)
                else:
                    print("📄 Solo una parte trovata, rilancio errore originale")
                    raise e
        
        # Per file normali (non multi-volume), usa la funzione helper
        return read_single_file_content(file_path)
        
    except Exception as e:
        raise Exception(f"Errore nella lettura del file {file_path}: {str(e)}")

def read_file_content(file_path):
    """
    Legge il contenuto di un file, gestendo file compressi .gz e encoding
    Rileva automaticamente se il file è compresso guardando i magic bytes
    
    Args:
        file_path (str): Percorso del file (può essere su NAS o locale)
    
    Returns:
        str: Contenuto del file
    """
    try:
        # Converti il path in oggetto Path per gestire sia path locali che di rete
        path_obj = Path(file_path)
        
        # Controlla se il file esiste
        if not path_obj.exists():
            raise FileNotFoundError(f"Il file {file_path} non esiste o non è accessibile")
        
        # Leggi i primi bytes per determinare se è compresso
        with open(file_path, 'rb') as f:
            first_bytes = f.read(2)
        
        # Magic bytes per gzip: 0x1f 0x8b
        is_gzip_compressed = (len(first_bytes) >= 2 and 
                             first_bytes[0] == 0x1f and 
                             first_bytes[1] == 0x8b)
        
        if is_gzip_compressed:
            print("🗜️  File compresso gzip rilevato (da magic bytes)")
            # File compresso con gzip
            with gzip.open(file_path, 'rt', encoding='utf-8', errors='ignore') as f:
                content = f.read()
        else:
            print("📄 File non compresso rilevato")
            # File normale, rileva encoding
            encoding = detect_encoding(file_path)
            with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
                content = f.read()
        
        return content
        
    except Exception as e:
        raise Exception(f"Errore nella lettura del file {file_path}: {str(e)}")

def process_file_to_excel_with_lookup(file_path):
    """
    Processa un file con formato Oracle specifico e crea un file Excel con lookup componenti
    """
    # Percorso del file Components.csv
    components_csv_path = Path(__file__).parent / "Components.csv"
    
    if not components_csv_path.exists():
        print(f"❌ ERRORE: File {components_csv_path} non trovato")
        return
    
    print(f"Processando file: {file_path}")
    
    # Detecta se è un file di rete (NAS)
    if str(file_path).startswith('\\\\'):
        print(f"Rilevato path di rete (NAS): {file_path}")
    
    # Leggi il contenuto del file con progress bar
    content = read_file_content_with_progress(file_path)
    
    if not content:
        print("❌ Errore: Impossibile leggere il file")
        return
    
    # Dividi il contenuto in righe
    lines = content.split('\n')
    print(f"📊 Totale righe nel file: {len(lines):,}")
    
    # Trova la riga con i nomi delle colonne (cerca righe che contengono "TABLE_NAME")
    header_line_index = None
    data_start_index = None
    
    for i, line in enumerate(lines):
        if 'TABLE_NAME' in line and 'COLUMN_NAME' in line:
            header_line_index = i
            print(f"📝 Intestazione trovata alla riga {i + 1}: {line.strip()}")
            
            # Cerca i separatori nelle righe successive
            for j in range(i + 1, min(i + 5, len(lines))):
                if '----' in lines[j]:
                    data_start_index = j + 1
                    print(f"📊 Inizio dati alla riga {data_start_index + 1}")
                    break
            break
    
    if header_line_index is None:
        print("❌ ERRORE: Impossibile trovare l'intestazione con TABLE_NAME")
        return
    
    if data_start_index is None:
        print("❌ ERRORE: Impossibile trovare l'inizio dei dati")
        return
    
    # Estrai le righe di dati
    data_lines = lines[data_start_index:]
    
    # Filtra le righe vuote
    data_lines = [line.strip() for line in data_lines if line.strip()]
    
    print(f"📊 Righe di dati valide: {len(data_lines):,}")
    
    # Processa i dati
    processed_data = []
    
    print("🔄 Parsing dei dati...")
    with tqdm(desc="Parsing righe", unit="righe", total=len(data_lines)) as pbar:
        for line in data_lines:
            # Dividi per virgola (il formato è OWNER,TABLE_NAME,COLUMN_NAME,DATA_TYPE)
            parts = line.split(',')
            if len(parts) >= 4:
                processed_data.append({
                    'OWNER': parts[0].strip(),
                    'TABLE_NAME': parts[1].strip(),
                    'COLUMN_NAME': parts[2].strip(),
                    'DATA_TYPE': parts[3].strip()
                })
            pbar.update(1)
    
    if not processed_data:
        print("❌ ERRORE: Nessun dato valido trovato")
        return
    
    # Crea DataFrame
    print("📊 Creazione DataFrame...")
    df_input = pd.DataFrame(processed_data)
    
    # Leggi il file Components.csv per il lookup
    print("📚 Caricamento file Components.csv...")
    df_components = pd.read_csv(components_csv_path, sep=';')
    
    print(f"📊 Statistiche file di input:")
    print(f"   - Righe: {len(df_input):,}")
    print(f"   - Colonne: {list(df_input.columns)}")
    print(f"📊 Statistiche file Components.csv:")
    print(f"   - Componenti: {len(df_components):,}")
    print(f"   - Colonne: {list(df_components.columns)}")
    
    # Esegui il lookup basato su TABLE_NAME
    print("🔍 Esecuzione lookup componenti...")
    with tqdm(desc="Lookup componenti", unit="righe") as pbar:
        # Merge usando TABLE_NAME con Prefisso, includendo anche Ufficio IT
        df_result = df_input.merge(
            df_components[['Prefisso', 'Nome_Componente', 'Ufficio IT']], 
            left_on='TABLE_NAME',
            right_on='Prefisso', 
            how='left'
        )
        pbar.update(len(df_result))
    
    # Gestisci i casi con più componenti/uffici per la stessa riga
    print("🔄 Gestione componenti multipli...")
    df_grouped = df_result.groupby(df_result.index).agg({
        **{col: 'first' for col in df_result.columns if col not in ['Nome_Componente', 'Ufficio IT', 'Prefisso']},
        'Nome_Componente': lambda x: ', '.join(x.dropna().astype(str).unique()) if not x.isna().all() else '',
        'Ufficio IT': lambda x: ', '.join(x.dropna().astype(str).unique()) if not x.isna().all() else ''
    })
    
    df_result = df_grouped.reset_index(drop=True)
    
    # Rimuovi la colonna duplicata 'Prefisso' e rinomina
    if 'Prefisso' in df_result.columns:
        df_result = df_result.drop('Prefisso', axis=1)
    
    # Rinomina le colonne per chiarezza
    if 'Nome_Componente' not in df_result.columns:
        df_result['Component_Name'] = ''
    else:
        df_result = df_result.rename(columns={'Nome_Componente': 'Component_Name'})
    
    if 'Ufficio IT' in df_result.columns:
        df_result = df_result.rename(columns={'Ufficio IT': 'Office_IT'})
    
    # Crea il nome del file output nella cartella Downloads
    base_name = file_path.stem
    today = datetime.now().strftime("%Y%m%d")
    output_filename = f"{base_name}_with_components_{today}.xlsx"
    
    # Percorso della cartella Downloads
    downloads_path = Path(os.path.expanduser("~")) / "Downloads" / "Lookup Component CDL"
    downloads_path.mkdir(parents=True, exist_ok=True)
    output_path = downloads_path / output_filename
    
    # Salva come file Excel
    print(f"💾 Salvataggio Excel in: {output_path}")
    with tqdm(desc="Salvataggio Excel", unit="righe") as pbar:
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            # Scrivi i dati nel foglio principale
            df_result.to_excel(writer, sheet_name='Database_Structure', index=False)
            
            # Ottieni il workbook e il foglio per la formattazione
            workbook = writer.book
            worksheet = writer.sheets['Database_Structure']
            
            # Formato per l'intestazione
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1
            })
            
            # Formato per le celle con component name trovato
            found_format = workbook.add_format({
                'fg_color': '#C6EFCE',
                'border': 1
            })
            
            # Formato per le celle con component name non trovato
            not_found_format = workbook.add_format({
                'fg_color': '#FFC7CE',
                'border': 1
            })
            
            # Applica formato all'intestazione
            for col_num, value in enumerate(df_result.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Applica formattazione condizionale per Component_Name
            component_col = df_result.columns.get_loc('Component_Name')
            
            # Colora le righe in base al lookup
            for row_num in range(1, len(df_result) + 1):
                component_value = df_result.iloc[row_num - 1]['Component_Name']
                if pd.notna(component_value) and component_value != '':
                    # Component trovato - verde
                    worksheet.write(row_num, component_col, component_value, found_format)
                else:
                    # Component non trovato - rosso
                    worksheet.write(row_num, component_col, '', not_found_format)
            
            # Adatta la larghezza delle colonne
            for i, col in enumerate(df_result.columns):
                max_length = max(
                    df_result[col].astype(str).map(len).max(),
                    len(col)
                )
                worksheet.set_column(i, i, min(max_length + 2, 50))
        
        pbar.update(len(df_result))
    
    print(f"\n🎉 File Excel creato con successo: {output_path}")
    print(f"📊 Righe processate: {len(df_result):,}")
    
    # Statistiche del lookup
    if 'Component_Name' in df_result.columns:
        lookup_success = df_result['Component_Name'].notna().sum()
        total_rows = len(df_result)
        success_rate = (lookup_success / total_rows * 100) if total_rows > 0 else 0
        print(f"✅ Lookup riusciti: {lookup_success:,} su {total_rows:,} ({success_rate:.1f}%)")
        
        # Mostra alcuni esempi
        successful_lookups = df_result[df_result['Component_Name'].notna()]
        failed_lookups = df_result[df_result['Component_Name'].isna()]
        
        if len(successful_lookups) > 0:
            print(f"\n📝 Primi 3 lookup riusciti:")
            for i in range(min(3, len(successful_lookups))):
                row = successful_lookups.iloc[i]
                print(f"   {row['TABLE_NAME']} → {row['Component_Name']}")
        
        if len(failed_lookups) > 0:
            print(f"\n⚠️  Primi 3 TABLE_NAME non trovati:")
            for i in range(min(3, len(failed_lookups))):
                row = failed_lookups.iloc[i]
                print(f"   {row['TABLE_NAME']} → (non trovato)")
        
        # Aggiungi statistiche per OWNER
        print(f"\n📊 Distribuzione per OWNER:")
        owner_stats = df_result['OWNER'].value_counts().head(10)
        for owner, count in owner_stats.items():
            print(f"   {owner}: {count:,} tabelle")

def lookup_component_info(table_name, df_components):
    """
    Cerca il nome del componente e l'ufficio IT basato sul TABLE_NAME usando la logica:
    1. Estrai i primi elementi separati da '_'
    2. Cerca prima con i primi due elementi (;ELEM1_ELEM2;)
    3. Se non trova, cerca solo il primo elemento (;ELEM1;)
    
    Returns:
        tuple: (component_name, office_it) o (None, None) se non trovato
    """
    if pd.isna(table_name) or not table_name:
        return None, None
    
    # Pulisci il table_name e converti in stringa
    table_name = str(table_name).strip()
    
    # Dividi per underscore e rimuovi elementi vuoti
    parts = [part for part in table_name.split('_') if part]
    
    if not parts:
        return None, None
    
    # Strategia 1: Cerca con i primi due elementi se disponibili
    if len(parts) >= 2:
        search_pattern_1 = f"{parts[0]}_{parts[1]}"
        # Cerca nei prefissi con pattern più specifico
        for _, row in df_components.iterrows():
            prefissi = str(row['Prefisso']).split(',')
            for prefisso in prefissi:
                prefisso = prefisso.strip()
                if prefisso and search_pattern_1.upper() == prefisso.upper():
                    return row['Nome_Componente'], row.get('Ufficio IT', '')
    
    # Strategia 2: Cerca solo con il primo elemento
    search_pattern_2 = parts[0]
    for _, row in df_components.iterrows():
        prefissi = str(row['Prefisso']).split(',')
        for prefisso in prefissi:
            prefisso = prefisso.strip()
            if prefisso and search_pattern_2.upper() == prefisso.upper():
                return row['Nome_Componente'], row.get('Ufficio IT', '')
    
    return None, None

def lookup_component_name(table_name, df_components):
    """
    Versione compatibile che restituisce solo il component name
    """
    component_name, _ = lookup_component_info(table_name, df_components)
    return component_name

def process_file_to_excel_with_lookup(file_path):
    """
    Processa un file con formato specifico (header a riga 15, separatori, dati)
    e crea un file Excel con lookup dei componenti
    """
    try:
        print(f"🔄 Processando file: {file_path}")
        
        # Leggi il file e trova la riga con le colonne
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        # Trova la riga con i nomi delle colonne (cerca quella che contiene ||)
        header_line_idx = None
        for i, line in enumerate(lines):
            if '||' in line and 'OWNER' in line.upper():
                header_line_idx = i
                break
        
        if header_line_idx is None:
            print("❌ Impossibile trovare la riga con i nomi delle colonne")
            return
        
        print(f"📍 Header trovato alla riga {header_line_idx + 1}")
        
        # Estrai i nomi delle colonne dalla riga header
        header_line = lines[header_line_idx].strip()
        # Rimuovi i ||',' per ottenere i nomi puliti
        column_names = []
        parts = header_line.split("||','||")
        for part in parts:
            # Pulisci ogni parte dai caratteri extra
            clean_part = part.replace("||','", "").replace("'", "").replace("||", "").strip()
            if clean_part:
                column_names.append(clean_part)
        
        print(f"📊 Colonne trovate: {column_names}")
        
        # Trova la riga con i separatori (-------)
        separator_line_idx = None
        for i in range(header_line_idx + 1, min(header_line_idx + 5, len(lines))):
            if '---' in lines[i]:
                separator_line_idx = i
                break
        
        if separator_line_idx is None:
            print("❌ Impossibile trovare la riga con i separatori")
            return
        
        print(f"📍 Separatori trovati alla riga {separator_line_idx + 1}")
        
        # Leggi i dati a partire dalla riga dopo i separatori
        data_start_idx = separator_line_idx + 1
        data_lines = lines[data_start_idx:]
        
        # Processa i dati
        data_rows = []
        print(f"📖 Processando {len(data_lines)} righe di dati...")
        
        for line_num, line in enumerate(tqdm(data_lines, desc="Parsing dati"), start=data_start_idx+1):
            line = line.strip()
            if not line:  # Salta righe vuote
                continue
            
            # Dividi per virgola
            row_data = [col.strip() for col in line.split(',')]
            
            # Assicurati che il numero di colonne corrisponda
            while len(row_data) < len(column_names):
                row_data.append('')
            
            # Tronca se ci sono troppe colonne
            row_data = row_data[:len(column_names)]
            
            data_rows.append(row_data)
        
        if not data_rows:
            print("❌ Nessuna riga di dati trovata")
            return
        
        # Crea DataFrame
        print(f"📊 Creazione DataFrame con {len(data_rows)} righe...")
        df = pd.DataFrame(data_rows, columns=column_names)
        
        # Carica il file Components.csv
        components_csv_path = Path(__file__).parent / "Components.csv"
        if not components_csv_path.exists():
            print(f"❌ File Components.csv non trovato in: {components_csv_path}")
            return
        
        print("📚 Caricamento file Components.csv...")
        df_components = pd.read_csv(components_csv_path, sep=';')
        print(f"📊 Componenti caricati: {len(df_components)}")
        
        # Esegui il lookup dei componenti
        print("🔍 Esecuzione lookup componenti...")
        component_names = []
        office_its = []
        
        # Trova la colonna TABLE_NAME
        table_name_col = None
        for col in df.columns:
            if 'TABLE_NAME' in col.upper():
                table_name_col = col
                break
        
        if table_name_col is None:
            print("❌ Colonna TABLE_NAME non trovata")
            return
        
        print(f"✅ Usando colonna '{table_name_col}' per il lookup")
        
        # Applica il lookup per ogni riga
        for table_name in tqdm(df[table_name_col], desc="Lookup componenti"):
            component_name, office_it = lookup_component_info(table_name, df_components)
            component_names.append(component_name if component_name else "")
            office_its.append(office_it if office_it else "")
        
        # Aggiungi le colonne Component Name e Office IT
        df['Component Name'] = component_names
        df['Office IT'] = office_its
        
        # Statistiche del lookup
        successful_lookups = sum(1 for name in component_names if name)
        failed_lookups = sum(1 for name in component_names if not name)
        total_rows = len(component_names)
        success_rate = (successful_lookups / total_rows * 100) if total_rows > 0 else 0
        failure_rate = (failed_lookups / total_rows * 100) if total_rows > 0 else 0
        
        print(f"📊 Lookup completato:")
        print(f"   - Righe totali: {total_rows:,}")
        print(f"   - Lookup riusciti: {successful_lookups:,} ({success_rate:.1f}%)")
        print(f"   - Lookup falliti: {failed_lookups:,} ({failure_rate:.1f}%)")
        
        if failed_lookups > 0:
            print(f"⚠️  ATTENZIONE: {failed_lookups:,} righe ({failure_rate:.1f}%) non hanno trovato corrispondenza!")
        else:
            print(f"🎉 Tutti i lookup sono riusciti al 100%!")
        
        # Crea il file Excel nella cartella Downloads
        downloads_path = Path(os.path.expanduser("~")) / "Downloads" / "Lookup Component CDL"
        downloads_path.mkdir(parents=True, exist_ok=True)
        
        base_name = Path(file_path).stem
        today = datetime.now().strftime("%Y%m%d")
        excel_filename = f"{base_name}_lookup_{today}.xlsx"
        excel_path = downloads_path / excel_filename
        
        print(f"💾 Creazione file Excel: {excel_path}")
        
        # Scrivi il file Excel con formattazione
        with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Data', index=False)
            
            # Ottieni il workbook e worksheet per la formattazione
            workbook = writer.book
            worksheet = writer.sheets['Data']
            
            # Formatta l'header
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1
            })
            
            # Formatta le celle di dati
            cell_format = workbook.add_format({
                'text_wrap': True,
                'valign': 'top',
                'border': 1
            })
            
            # Applica formattazione all'header
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Auto-ajusta la larghezza delle colonne
            for i, col in enumerate(df.columns):
                max_length = max(
                    df[col].astype(str).apply(len).max(),
                    len(str(col))
                )
                # Limita la larghezza massima a 50 caratteri
                worksheet.set_column(i, i, min(max_length + 2, 50))
            
            # Evidenzia la colonna Component Name
            component_col_idx = len(df.columns) - 1
            component_format = workbook.add_format({
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#FFE6CC',
                'border': 1
            })
            
            for row_num in range(1, len(df) + 1):
                worksheet.write(row_num, component_col_idx, 
                              df.iloc[row_num-1]['Component Name'], component_format)
        
        print(f"🎉 File Excel creato con successo!")
        print(f"📍 Percorso: {excel_path}")
        
        # Mostra alcuni esempi di lookup
        if successful_lookups > 0:
            print(f"\n📝 Esempi di lookup riusciti:")
            successful_rows = df[df['Component Name'] != ''].head(3)
            for _, row in successful_rows.iterrows():
                print(f"   {row[table_name_col]} → {row['Component Name']}")
        
        failed_lookups = total_rows - successful_lookups
        if failed_lookups > 0:
            print(f"\n⚠️  Esempi di lookup falliti:")
            failed_rows = df[df['Component Name'] == ''].head(3)
            for _, row in failed_rows.iterrows():
                print(f"   {row[table_name_col]} → (non trovato)")
        
    except Exception as e:
        print(f"❌ Errore durante l'elaborazione: {e}")
        import traceback
        traceback.print_exc()

def is_compressed_log_file(file_path):
    """
    Determina se un file è un file di log compresso, inclusi file multi-parte.
    Riconosce pattern come:
    - .gz
    - .log
    - .log.gz  
    - .log.gz.txt
    - .log.gz.001, .log.gz.002, etc. (file multi-parte)
    
    Args:
        file_path (str): Percorso del file
    
    Returns:
        bool: True se è un file di log compresso
    """
    file_path_lower = str(file_path).lower()
    
    # Controlli base
    if (file_path_lower.endswith('.gz') or 
        file_path_lower.endswith('.log') or 
        file_path_lower.endswith('.log.gz.txt')):
        return True
    
    # Controllo per file multi-parte compressi (.log.gz.001, .log.gz.002, etc.)
    # Pattern per rilevare .log.gz.XXX dove XXX sono cifre
    if re.search(r'\.log\.gz\.\d+$', file_path_lower):
        return True
    
    # Pattern per rilevare .gz.XXX dove XXX sono cifre  
    if re.search(r'\.gz\.\d+$', file_path_lower):
        return True
        
    return False

def is_multivolume_file(file_path):
    """
    Determina se un file è parte di un archivio multi-volume (.001, .002, etc.)
    
    Args:
        file_path (str): Percorso del file
    
    Returns:
        bool: True se è un file multi-volume
    """
    file_path_lower = str(file_path).lower()
    
    # Pattern per rilevare file con estensioni numeriche (.001, .002, etc.)
    if re.search(r'\.\d{3}(\.\w+)?$', file_path_lower):
        return True
    
    return False

def get_multivolume_base_name(file_path):
    """
    Ottiene il nome base di un file multi-volume rimuovendo il numero progressivo
    
    Args:
        file_path (str): Percorso del file (es. file.txt.001)
    
    Returns:
        str: Nome base (es. file.txt)
    """
    # Rimuovi il numero progressivo dalla fine
    # Pattern per .001, .002, etc. con o senza estensione aggiuntiva
    base_name = re.sub(r'\.\d{3}(\.\w+)?$', '', str(file_path))
    return base_name

def find_multivolume_files(base_file_path, directory_path):
    """
    Trova tutti i file multi-volume correlati a un file base
    
    Args:
        base_file_path (str): Percorso del primo file (.001)
        directory_path (str): Directory dove cercare i file
    
    Returns:
        list: Lista ordinata di tutti i file multi-volume trovati
    """
    base_name = get_multivolume_base_name(base_file_path)
    directory = Path(directory_path)
    
    # Trova tutti i file che iniziano con il nome base
    pattern_files = []
    
    try:
        for file_path in directory.iterdir():
            if file_path.is_file():
                file_str = str(file_path)
                # Controlla se il file corrisponde al pattern multi-volume
                if file_str.startswith(base_name) and is_multivolume_file(file_str):
                    pattern_files.append(file_str)
        
        # Ordina i file numericamente per progressivo
        pattern_files.sort()
        
    except Exception as e:
        print(f"⚠️  Errore nella ricerca file multi-volume: {e}")
        return [base_file_path]  # Ritorna almeno il file originale
    
    return pattern_files if pattern_files else [base_file_path]

def read_multivolume_content(file_paths):
    """
    Legge e concatena il contenuto di più file multi-volume
    Gestisce sia file testuali che archivi gzip suddivisi
    
    Args:
        file_paths (list): Lista ordinata dei percorsi dei file multi-volume
    
    Returns:
        str: Contenuto concatenato di tutti i file validi
    """
    print(f"📚 Lettura file multi-volume: {len(file_paths)} parti")
    
    # Controlla se sono file gzip multi-volume
    first_file = file_paths[0] if file_paths else ""
    is_gzip_multivolume = any(ext in Path(first_file).name.lower() for ext in ['.gz.', '.gzip.'])
    
    if is_gzip_multivolume:
        print("🗜️  Rilevati file gzip multi-volume - concatenazione binaria")
        return read_gzip_multivolume_content(file_paths)
    else:
        print("📄 File multi-volume standard - concatenazione testuale")
        return read_text_multivolume_content(file_paths)

def read_gzip_multivolume_content(file_paths):
    """
    Concatena file gzip multi-volume in binario e decomprime
    """
    import tempfile
    
    total_size = 0
    for file_path in file_paths:
        try:
            total_size += Path(file_path).stat().st_size
        except:
            pass
    
    print(f"📊 Dimensione totale: {total_size/1024/1024:.1f} MB")
    
    with tqdm(total=total_size, desc="Concatenazione gzip", unit="B", unit_scale=True) as pbar:
        # Crea un file temporaneo per concatenare i dati binari
        with tempfile.NamedTemporaryFile(delete=False, suffix='.gz') as temp_gz:
            for i, file_path in enumerate(file_paths):
                print(f"  Parte {i+1}/{len(file_paths)}: {Path(file_path).name}")
                
                try:
                    # Legge in binario e concatena
                    with open(file_path, 'rb') as f:
                        chunk_size = 8192
                        while True:
                            chunk = f.read(chunk_size)
                            if not chunk:
                                break
                            temp_gz.write(chunk)
                            pbar.update(len(chunk))
                    
                    print(f"✅ Parte {i+1} concatenata")
                    
                except Exception as e:
                    print(f"⚠️  Errore lettura parte {i+1}: {str(e)}")
                    continue
            
            temp_gz_path = temp_gz.name
        
        # Ora decomprime il file concatenato
        print("🔄 Decompressione file concatenato...")
        try:
            with gzip.open(temp_gz_path, 'rt', encoding='utf-8', errors='replace') as gz_file:
                content = gz_file.read()
                print(f"✅ Decompressione completata: {len(content)} caratteri")
                return content
        except Exception as e:
            # Prova con encoding alternativo
            print(f"⚠️  Tentativo con encoding Latin-1: {str(e)}")
            try:
                with gzip.open(temp_gz_path, 'rt', encoding='latin-1', errors='replace') as gz_file:
                    content = gz_file.read()
                    print(f"✅ Decompressione Latin-1 completata: {len(content)} caratteri")
                    return content
            except Exception as e2:
                print(f"❌ Decompressione fallita: {str(e2)}")
                raise
        finally:
            # Pulisce il file temporaneo
            try:
                os.unlink(temp_gz_path)
            except:
                pass

def read_text_multivolume_content(file_paths):
    """
    Concatena file multi-volume testuali normali
    """
    combined_content = ""
    total_size = 0
    valid_parts = 0
    
    # Calcola dimensione totale per la barra di progresso
    for file_path in file_paths:
        try:
            total_size += Path(file_path).stat().st_size
        except:
            pass
    
    print(f"📊 Dimensione totale: {total_size/1024/1024:.1f} MB")
    
    with tqdm(total=total_size, unit='B', unit_scale=True, desc="Lettura multi-volume") as pbar:
        for i, file_path in enumerate(file_paths):
            print(f"📄 Parte {i+1}/{len(file_paths)}: {Path(file_path).name}")
            
            try:
                # Leggi ogni parte usando la funzione esistente
                part_content = read_file_content(file_path)
                
                # Verifica se il contenuto sembra valido (non binario corrotto)
                if is_content_valid(part_content):
                    combined_content += part_content
                    valid_parts += 1
                    print(f"✅ Parte {i+1} aggiunta: {len(part_content)} caratteri")
                else:
                    print(f"⚠️  Parte {i+1} saltata: contenuto corrotto/binario")
                
                # Aggiorna progresso
                file_size = Path(file_path).stat().st_size
                pbar.update(file_size)
                
            except Exception as e:
                print(f"⚠️  Errore lettura parte {i+1}: {str(e)}")
                # Continua con le altre parti
                continue
    
    if not combined_content:
        raise ValueError("Nessun contenuto valido trovato in tutte le parti")
    
    print(f"✅ Multi-volume letto: {len(combined_content)} caratteri da {valid_parts}/{len(file_paths)} parti valide")
    return combined_content
    
    print(f"✅ Multi-volume letto: {len(combined_content)} caratteri da {valid_parts}/{len(file_paths)} parti valide")
    return combined_content

def is_content_valid(content):
    """
    Verifica se il contenuto sembra essere testo valido e non dati binari corrotti
    
    Args:
        content (str): Contenuto da verificare
    
    Returns:
        bool: True se il contenuto sembra valido
    """
    if not content or len(content) < 100:
        return False
    
    # Prendi un campione del contenuto per l'analisi
    sample = content[:1000]
    
    # Conta caratteri stampabili vs non stampabili
    printable_chars = sum(1 for c in sample if c.isprintable() or c in '\n\r\t')
    total_chars = len(sample)
    
    # Se più del 70% dei caratteri sono stampabili, considera il contenuto valido
    if total_chars > 0:
        printable_ratio = printable_chars / total_chars
        return printable_ratio > 0.7
    
    return False

def is_header_line_valid(line):
    """
    Verifica se una riga sembra essere un header valido (meno restrittivo di is_content_valid)
    
    Args:
        line (str): Riga da verificare
    
    Returns:
        bool: True se la riga sembra un header valido
    """
    if not line or len(line.strip()) < 10:
        return False
    
    # Per gli header, richiediamo solo che almeno l'80% dei caratteri sia stampabile
    printable_chars = sum(1 for c in line if c.isprintable() or c in '\n\r\t')
    total_chars = len(line)
    
    if total_chars > 0:
        printable_ratio = printable_chars / total_chars
        # Più permissivo per le righe header
        return printable_ratio > 0.8
    
    return False

def process_file_to_csv_with_lookup(file_path, components_csv_path="Components.csv"):
    """
    Trasforma un file nel formato CSV di Components.csv e aggiunge la colonna Component Name
    tramite lookup dal file Components.csv. Gestisce file su NAS e file compressi .gz
    
    Args:
        file_path (str): Percorso del file da processare (può essere su NAS o locale)
        components_csv_path (str): Percorso del file Components.csv per il lookup
    """
    try:
        print(f"Processando file: {file_path}")
        
        # Converti il path in oggetto Path per gestire path di rete
        input_path = Path(file_path)
        
        # Controlla se è un path di rete
        if str(input_path).startswith('\\\\'):
            print(f"Rilevato path di rete (NAS): {file_path}")
        
        # Leggi il contenuto del file (gestisce .gz e encoding) - USA LA NUOVA FUNZIONE
        if is_compressed_log_file(file_path):
            # File di log, probabilmente non è un CSV standard
            print("📋 File di log rilevato, tentativo di parsing...")
            content = read_file_content_with_progress(file_path)
            
            # Salva temporaneamente il contenuto come file per pandas
            temp_file = "temp_extracted.csv"
            print("💾 Salvataggio temporaneo...")
            with open(temp_file, 'w', encoding='utf-8') as f:
                # Scrivi con progresso per file grandi
                lines = content.splitlines()
                filtered_lines = []
                
                # Filtra le righe, fermandosi a "X rows selected."
                for line in lines:
                    line_strip = line.strip()
                    if " rows selected." in line_strip.lower() or " row selected." in line_strip.lower():
                        print(f"📋 Trovata riga finale Oracle: '{line_strip}' - Fermando il parsing")
                        break
                    filtered_lines.append(line)
                
                print(f"📊 Righe filtrate: {len(filtered_lines)} su {len(lines)} totali")
                
                with tqdm(total=len(filtered_lines), desc="Scrittura temp", unit="righe") as pbar:
                    for line in filtered_lines:
                        f.write(line + '\n')
                        pbar.update(1)
            
            print("🔍 Analisi formato CSV...")
            try:
                # Prova a leggere come CSV
                df_input = pd.read_csv(temp_file, sep=None, engine='python')
                print("✅ Formato CSV rilevato automaticamente")
            except:
                # Se fallisce, prova con diversi separatori
                print("⚠️  Formato non riconosciuto, provo separatori specifici...")
                try:
                    df_input = pd.read_csv(temp_file, sep=';')
                    print("✅ Formato CSV con ';' rilevato")
                except:
                    try:
                        df_input = pd.read_csv(temp_file, sep='\t')
                        print("✅ Formato CSV con TAB rilevato")
                    except:
                        # Se tutto fallisce, crea un dataframe con il contenuto raw
                        print("⚠️  Creazione dataframe raw...")
                        lines = content.split('\n')
                        df_input = pd.DataFrame({'raw_content': lines})
                        print("✅ Dataframe raw creato")
            finally:
                # Rimuovi il file temporaneo
                if os.path.exists(temp_file):
                    os.remove(temp_file)
        else:
            # File CSV normale
            print("📊 Lettura file CSV...")
            df_input = pd.read_csv(file_path)
            print("✅ CSV caricato")
        
        # Leggi il file Components.csv per il lookup (con il nuovo formato)
        print("📚 Caricamento file Components.csv...")
        df_components = pd.read_csv(components_csv_path, sep=';')
        
        print(f"📊 Statistiche file di input:")
        print(f"   - Righe: {len(df_input):,}")
        print(f"   - Colonne: {len(df_input.columns)}")
        print(f"   - Colonne disponibili: {list(df_input.columns)}")
        print(f"📊 Statistiche file Components.csv:")
        print(f"   - Componenti: {len(df_components):,}")
        print(f"   - Colonne: {list(df_components.columns)}")
        
        # Crea il nome del file output nella cartella Downloads
        base_name = input_path.stem  # Nome senza estensione
        today = datetime.now().strftime("%Y%m%d")
        output_filename = f"{base_name}_vlookup_{today}.csv"
        
        # Percorso della cartella Downloads
        downloads_path = Path(os.path.expanduser("~")) / "Downloads" / "Lookup Component CDL"
        downloads_path.mkdir(parents=True, exist_ok=True)  # Crea la cartella se non esiste
        output_path = downloads_path / output_filename
        
        # Adatta le colonne in base al nuovo formato di Components.csv
        # Cerca colonne che potrebbero contenere prefissi di componenti
        print("🔍 Ricerca colonna prefisso...")
        prefix_column = None
        for col in df_input.columns:
            if any(keyword in col.lower() for keyword in ['prefix', 'prefisso', 'component', 'comp']):
                prefix_column = col
                break
        
        if prefix_column:
            print(f"✅ Trovata colonna '{prefix_column}' per il lookup")
            print("🔄 Esecuzione lookup...")
            
            # Mostra progresso del merge
            with tqdm(desc="Lookup componenti", unit="righe") as pbar:
                # Esegui il lookup
                df_result = df_input.merge(
                    df_components[['Prefisso', 'Nome_Componente']], 
                    left_on=prefix_column,
                    right_on='Prefisso', 
                    how='left'
                )
                pbar.update(len(df_result))
        else:
            print("⚠️  Nessuna colonna di prefisso trovata, aggiungendo colonna vuota")
            df_result = df_input.copy()
            df_result['Nome_Componente'] = ''
        
        # Salva il file risultante
        print(f"💾 Salvataggio risultato in: {output_path}")
        with tqdm(desc="Salvataggio CSV", unit="righe") as pbar:
            df_result.to_csv(output_path, index=False, sep=';')
            pbar.update(len(df_result))
        
        print(f"\n🎉 File creato con successo: {output_path}")
        print(f"📊 Righe processate: {len(df_result):,}")
        if 'Nome_Componente' in df_result.columns:
            lookup_success = df_result['Nome_Componente'].notna().sum()
            lookup_failed = df_result['Nome_Componente'].isna().sum()
            total_rows = len(df_result)
            success_rate = (lookup_success / total_rows * 100) if total_rows > 0 else 0
            failure_rate = (lookup_failed / total_rows * 100) if total_rows > 0 else 0
            
            print(f"✅ Lookup riusciti: {lookup_success:,} su {total_rows:,} ({success_rate:.1f}%)")
            print(f"❌ Lookup falliti: {lookup_failed:,} su {total_rows:,} ({failure_rate:.1f}%)")
            
            if lookup_failed > 0:
                print(f"⚠️  ATTENZIONE: {lookup_failed:,} righe ({failure_rate:.1f}%) non hanno trovato corrispondenza!")
            
            # Mostra alcuni esempi di lookup riusciti e falliti
            successful_lookups = df_result[df_result['Nome_Componente'].notna()]
            failed_lookups = df_result[df_result['Nome_Componente'].isna()]
            
            if len(successful_lookups) > 0:
                print(f"\n📝 Primi 3 lookup riusciti:")
                for i in range(min(3, len(successful_lookups))):
                    row = successful_lookups.iloc[i]
                    if prefix_column in row:
                        print(f"   {row[prefix_column]} → {row['Nome_Componente']}")
            
            if len(failed_lookups) > 0:
                print(f"\n⚠️  Primi 3 prefissi non trovati:")
                for i in range(min(3, len(failed_lookups))):
                    row = failed_lookups.iloc[i]
                    if prefix_column in row:
                        print(f"   {row[prefix_column]} → (non trovato)")
        
        # Riepilogo finale evidenziato
        if 'Nome_Componente' in df_result.columns:
            print(f"\n" + "="*80)
            print(f"📊 RIEPILOGO FINALE LOOKUP CSV")
            print(f"="*80)
            print(f"📋 Righe totali processate: {total_rows:,}")
            print(f"✅ Lookup riusciti: {lookup_success:,} ({success_rate:.1f}%)")
            print(f"❌ Lookup falliti: {lookup_failed:,} ({failure_rate:.1f}%)")
            if lookup_failed > 0:
                print(f"⚠️  NOTA: {lookup_failed:,} righe necessitano di verifica manuale")
            else:
                print(f"🎉 PERFETTO! Tutti i lookup sono riusciti!")
            print(f"="*80)
        
    except FileNotFoundError as e:
        print(f"Errore: File non trovato - {e}")
    except Exception as e:
        print(f"Errore durante l'elaborazione: {e}")
        import traceback
        traceback.print_exc()


def prepare_lookup_dictionary(df_components):
    """
    Prepara un dizionario di lookup ottimizzato per ricerche veloci
    Questo riduce drasticamente il tempo di ricerca da O(n) a O(1)
    """
    print("🔧 Preparazione dizionario di lookup ottimizzato...")
    lookup_dict = {}
    
    with tqdm(desc="Preparazione lookup", unit="componenti") as pbar:
        for _, row in df_components.iterrows():
            component_name = row['Nome_Componente']
            prefissi = str(row['Prefisso']).split(',')
            
            for prefisso in prefissi:
                prefisso = prefisso.strip()
                if prefisso and prefisso != 'nan':
                    # Aggiungi al dizionario - se esiste già, concatena
                    if prefisso in lookup_dict:
                        if component_name not in lookup_dict[prefisso]:
                            lookup_dict[prefisso] += f", {component_name}"
                    else:
                        lookup_dict[prefisso] = component_name
            pbar.update(1)
    
    print(f"📊 Dizionario di lookup creato: {len(lookup_dict)} prefissi mappati")
    return lookup_dict

def prepare_office_dictionary(df_components):
    """
    Prepara un dizionario di lookup per gli uffici IT
    """
    print("🔧 Preparazione dizionario uffici IT...")
    office_dict = {}
    
    with tqdm(desc="Preparazione uffici", unit="componenti") as pbar:
        for _, row in df_components.iterrows():
            office_it = row.get('Ufficio IT', '')
            # Gestisci valori NaN e non stringhe
            if pd.isna(office_it) or office_it == '' or str(office_it).lower() == 'nan':
                office_it = ""
            else:
                office_it = str(office_it).strip()
            
            prefissi = str(row['Prefisso']).split(',')
            
            for prefisso in prefissi:
                prefisso = prefisso.strip()
                if prefisso and prefisso != 'nan':
                    # Aggiungi al dizionario - se esiste già, concatena uffici unici
                    if prefisso in office_dict:
                        existing_offices = [off.strip() for off in office_dict[prefisso].split(', ') if off.strip()] if office_dict[prefisso] else []
                        if office_it and office_it not in existing_offices:
                            existing_offices.append(office_it)
                            # Filtra valori vuoti e ordina solo stringhe valide
                            valid_offices = [off for off in existing_offices if off and str(off).lower() != 'nan']
                            office_dict[prefisso] = ', '.join(sorted(valid_offices)) if valid_offices else ""
                    else:
                        office_dict[prefisso] = office_it if office_it else ""
            pbar.update(1)
    
    print(f"📊 Dizionario uffici IT creato: {len(office_dict)} prefissi mappati")
    return office_dict

def lookup_component_info_optimized(table_name, lookup_dict, office_dict):
    """
    Versione ottimizzata del lookup usando dizionari pre-processati
    Restituisce (component_name, office_it)
    """
    if pd.isna(table_name) or not table_name:
        return "", ""
    
    # Pulisci il table_name e converti in stringa
    table_name = str(table_name).strip()
    
    # Parse corretto per gestire $ e _
    # Se c'è un $, inizia dalla parte dopo $
    if '$' in table_name:
        # Divide per $ e prende la parte dopo
        dollar_parts = table_name.split('$')
        if len(dollar_parts) > 1:
            # Prende tutto dopo il primo $
            after_dollar = dollar_parts[1]
            # Divide per _ e # per ottenere le parti, rimuovendo # ma non $
            after_dollar_clean = after_dollar.replace('#', '')
            parts_after_dollar = [part for part in after_dollar_clean.split('_') if part]
            
            # Cerca ogni parte singolarmente, in ordine
            for part in parts_after_dollar:
                if part in lookup_dict:
                    component_name = lookup_dict[part]
                    office_it = office_dict.get(part, "")
                    return component_name, office_it
            
            # Se nessuna parte singola è stata trovata, prova combinazioni
            if len(parts_after_dollar) >= 2:
                # Prova combinazioni di due elementi consecutivi
                for i in range(len(parts_after_dollar) - 1):
                    combined_key = f"{parts_after_dollar[i]}_{parts_after_dollar[i+1]}"
                    if combined_key in lookup_dict:
                        component_name = lookup_dict[combined_key]
                        office_it = office_dict.get(combined_key, "")
                        return component_name, office_it
    
    else:
        # Parse normale senza $
        parts = [part for part in table_name.split('_') if part]
        
        if not parts:
            return "", ""
        
        # Cerca ogni parte singolarmente, in ordine
        for part in parts:
            if part in lookup_dict:
                component_name = lookup_dict[part]
                office_it = office_dict.get(part, "")
                return component_name, office_it
        
        # Se nessuna parte singola è stata trovata, prova combinazioni
        if len(parts) >= 2:
            # Prova combinazioni di due elementi consecutivi
            for i in range(len(parts) - 1):
                combined_key = f"{parts[i]}_{parts[i+1]}"
                if combined_key in lookup_dict:
                    component_name = lookup_dict[combined_key]
                    office_it = office_dict.get(combined_key, "")
                    return component_name, office_it
    
    return "", ""

def lookup_component_name_optimized(table_name, lookup_dict):
    """
    Versione compatibile che restituisce solo il component name
    """
    component_name, _ = lookup_component_info_optimized(table_name, lookup_dict, {})
    return component_name

def analyze_failed_lookup(table_name, lookup_dict):
    """
    Analizza perché un lookup è fallito e restituisce le stringhe cercate
    """
    if pd.isna(table_name) or not table_name:
        return []
    
    # Pulisci il table_name e converti in stringa
    table_name = str(table_name).strip()
    
    searched_strings = []
    
    # Parse corretto per gestire $ e _
    # Se c'è un $, inizia dalla parte dopo $
    if '$' in table_name:
        # Divide per $ e prende la parte dopo
        dollar_parts = table_name.split('$')
        if len(dollar_parts) > 1:
            # Prende tutto dopo il primo $
            after_dollar = dollar_parts[1]
            # Divide per _ e # per ottenere le parti, rimuovendo # ma non $
            after_dollar_clean = after_dollar.replace('#', '')
            parts_after_dollar = [part for part in after_dollar_clean.split('_') if part]
            
            # Cerca ogni parte singolarmente, in ordine
            for i, part in enumerate(parts_after_dollar):
                searched_strings.append({
                    'search_type': f'Elemento {i+1} dopo $ (singolo)',
                    'search_string': part,
                    'found_in_components': part in lookup_dict,
                    'component_found': lookup_dict.get(part, '')
                })
            
            # Se nessuna parte singola è stata trovata, prova combinazioni
            if len(parts_after_dollar) >= 2:
                # Prova combinazioni di due elementi consecutivi
                for i in range(len(parts_after_dollar) - 1):
                    combination = f"{parts_after_dollar[i]}_{parts_after_dollar[i+1]}"
                    searched_strings.append({
                        'search_type': f'Combinazione elementi {i+1}-{i+2} dopo $',
                        'search_string': combination,
                        'found_in_components': combination in lookup_dict,
                        'component_found': lookup_dict.get(combination, '')
                    })
    
    # Se non c'è $ o non si trova nulla nella parte dopo $, 
    # analizza l'intera stringa normalmente
    clean_name = table_name.replace('$', '').replace('#', '')
    parts = [part for part in clean_name.split('_') if part]
    
    if parts:
        # Cerca ogni parte singolarmente, in ordine
        for i, part in enumerate(parts):
            searched_strings.append({
                'search_type': f'Elemento {i+1} (singolo, clean)',
                'search_string': part,
                'found_in_components': part in lookup_dict,
                'component_found': lookup_dict.get(part, '')
            })
        
        # Se nessuna parte singola è stata trovata, prova combinazioni
        if len(parts) >= 2:
            # Prova combinazioni di due elementi consecutivi
            for i in range(len(parts) - 1):
                combination = f"{parts[i]}_{parts[i+1]}"
                searched_strings.append({
                    'search_type': f'Combinazione elementi {i+1}-{i+2} (clean)',
                    'search_string': combination,
                    'found_in_components': combination in lookup_dict,
                    'component_found': lookup_dict.get(combination, '')
                })
    
    return searched_strings

def create_failed_lookup_analysis(failed_table_names, lookup_dict):
    """
    Crea un DataFrame con l'analisi dettagliata dei lookup falliti
    Raggruppa per TABLE_NAME e concatena tutte le stringhe cercate
    """
    print("🔍 Analisi dei lookup falliti...")
    
    analysis_data = []
    
    with tqdm(desc="Analisi fallimenti", unit="righe") as pbar:
        for table_name in failed_table_names:
            searched_list = analyze_failed_lookup(table_name, lookup_dict)
            
            # Raggruppa tutte le informazioni per questo table_name
            search_strings = []
            search_types = []
            found_components = []
            found_in_components_list = []
            
            for search_info in searched_list:
                search_strings.append(search_info['search_string'])
                search_types.append(search_info['search_type'])
                found_in_components_list.append('YES' if search_info['found_in_components'] else 'NO')
                if search_info['component_found']:
                    found_components.append(search_info['component_found'])
            
            # Crea una singola riga per table_name con tutte le stringhe concatenate
            analysis_data.append({
                'TABLE_NAME': table_name,
                'Search_Strings': ' | '.join(search_strings),
                'Search_Types': ' | '.join(search_types),
                'Found_in_Components': ' | '.join(found_in_components_list),
                'Components_Found': ' | '.join(found_components) if found_components else 'NONE',
                'Total_Searches': len(search_strings),
                'Reason_Failed': 'Nessuna stringa trovata nel file Components.csv' if not any(search_info['found_in_components'] for search_info in searched_list) else 'Alcune stringhe trovate ma non selezionate (priorità)'
            })
            pbar.update(1)
    
    return pd.DataFrame(analysis_data)

def create_failed_lookup_analysis_with_split(failed_table_names, lookup_dict, output_path):
    """
    Crea file Excel con analisi dettagliata dei lookup falliti,
    usando la stessa logica robusta del file principale per la suddivisione
    """
    if not failed_table_names:
        return []
    
    print("🔍 Analisi dei lookup falliti...")
    
    # Crea prima l'analisi completa
    print("📊 Generazione analisi completa...")
    analysis_df = create_failed_lookup_analysis(failed_table_names, lookup_dict)
    
    total_rows = len(analysis_df)
    print(f"� Righe di analisi generate: {total_rows:,}")
    
    base_path = Path(output_path)
    base_name = f"{base_path.stem} - Analisi_Fallimenti"
    
    created_files = []
    
    # Usa la stessa logica di suddivisione del file principale
    if total_rows <= EXCEL_MAX_ROWS_PER_SHEET:
        # Caso semplice: tutto in un singolo sheet
        print("📄 Creazione in singolo sheet...")
        analysis_file = base_path.parent / f"{base_name}{base_path.suffix}"
        
        with pd.ExcelWriter(analysis_file, engine='xlsxwriter') as writer:
            analysis_df.to_excel(writer, sheet_name='Lookup_Failed', index=False)
            _apply_basic_formatting(writer, 'Lookup_Failed', analysis_df)
        
        created_files.append(analysis_file)
        print(f"✅ File creato: {analysis_file.name}")
        
    elif total_rows <= EXCEL_MAX_ROWS_PER_FILE:
        # Caso intermedio: suddividi in più sheet nello stesso file
        print("📑 Suddivisione in più sheet...")
        analysis_file = base_path.parent / f"{base_name}{base_path.suffix}"
        
        with pd.ExcelWriter(analysis_file, engine='xlsxwriter') as writer:
            start_row = 0
            sheet_num = 1
            
            while start_row < total_rows:
                end_row = min(start_row + EXCEL_MAX_ROWS_PER_SHEET, total_rows)
                chunk_df = analysis_df.iloc[start_row:end_row].copy()
                
                sheet_name = f'Lookup_Failed_Part{sheet_num}'
                chunk_df.to_excel(writer, sheet_name=sheet_name, index=False)
                _apply_basic_formatting(writer, sheet_name, chunk_df)
                
                print(f"   📄 Sheet {sheet_num}: righe {start_row+1:,}-{end_row:,}")
                
                start_row = end_row
                sheet_num += 1
        
        created_files.append(analysis_file)
        print(f"✅ File multi-sheet creato: {analysis_file.name}")
        
    else:
        # Caso complesso: suddividi in più file
        print("� Suddivisione in più file...")
        
        start_row = 0
        file_num = 1
        
        while start_row < total_rows:
            # Calcola quante righe mettere in questo file
            remaining_rows = total_rows - start_row
            rows_this_file = min(EXCEL_MAX_ROWS_PER_FILE, remaining_rows)
            
            end_row = start_row + rows_this_file
            file_df = analysis_df.iloc[start_row:end_row].copy()
            
            # Nome del file
            analysis_file = base_path.parent / f"{base_name}_Parte_{file_num}{base_path.suffix}"
            
            # Se il file ha poche righe, usa un singolo sheet
            if len(file_df) <= EXCEL_MAX_ROWS_PER_SHEET:
                with pd.ExcelWriter(analysis_file, engine='xlsxwriter') as writer:
                    file_df.to_excel(writer, sheet_name='Lookup_Failed', index=False)
                    _apply_basic_formatting(writer, 'Lookup_Failed', file_df)
                
                print(f"   📄 File {file_num}: {len(file_df):,} righe")
            else:
                # Il file ha troppe righe, suddividi in più sheet
                with pd.ExcelWriter(analysis_file, engine='xlsxwriter') as writer:
                    file_start = 0
                    sheet_num = 1
                    
                    while file_start < len(file_df):
                        file_end = min(file_start + EXCEL_MAX_ROWS_PER_SHEET, len(file_df))
                        sheet_df = file_df.iloc[file_start:file_end].copy()
                        
                        sheet_name = f'Lookup_Failed_Part{sheet_num}'
                        sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        _apply_basic_formatting(writer, sheet_name, sheet_df)
                        
                        file_start = file_end
                        sheet_num += 1
                
                print(f"   📄 File {file_num}: {len(file_df):,} righe in {sheet_num-1} sheet")
            
            created_files.append(analysis_file)
            start_row = end_row
            file_num += 1
    
    print(f"✅ Analisi fallimenti completata: {len(created_files)} file creati")
    return created_files

def _apply_basic_formatting(writer, sheet_name, df):
    """Applica formattazione di base al sheet Excel"""
    try:
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        # Formato header
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D7E4BC',
            'border': 1
        })
        
        # Applica formato all'header
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Auto-fit colonne
        for i, col in enumerate(df.columns):
            max_length = max(
                df[col].astype(str).str.len().max(),
                len(str(col))
            )
            worksheet.set_column(i, i, min(max_length + 2, 50))
            
    except Exception as e:
        print(f"⚠️  Errore nella formattazione: {e}")
        # Continua senza formattazione se c'è un errore

def process_chunk(chunk_data):
    """
    Processa un chunk di dati per il lookup parallelizzato
    """
    chunk, lookup_dict = chunk_data
    return [lookup_component_name_optimized(table_name, lookup_dict) for table_name in chunk]

def process_chunk_with_office(chunk_data):
    """
    Processa un chunk di dati per il lookup parallelizzato con uffici
    """
    chunk, lookup_dict, office_dict = chunk_data
    results = []
    for table_name in chunk:
        component_name, office_it = lookup_component_info_optimized(table_name, lookup_dict, office_dict)
        results.append((component_name, office_it))
    return results

def parallel_lookup_with_office(table_names, lookup_dict, office_dict, n_processes=None):
    """
    Esegue il lookup in parallelo per componenti e uffici
    """
    if n_processes is None:
        n_processes = min(mp.cpu_count(), 4)
    
    print(f"🚀 Avvio lookup parallelo con uffici usando {n_processes} processi...")
    
    # Dividi i dati in chunks
    chunk_size = max(1000, len(table_names) // (n_processes * 4))
    chunks = [table_names[i:i + chunk_size] for i in range(0, len(table_names), chunk_size)]
    
    print(f"📊 Dati divisi in {len(chunks)} chunks di ~{chunk_size} elementi")
    
    # Prepara i dati per il multiprocessing
    chunk_data = [(chunk, lookup_dict, office_dict) for chunk in chunks]
    
    # Processa in parallelo
    component_results = []
    office_results = []
    with mp.Pool(processes=n_processes) as pool:
        with tqdm(desc="Lookup parallelo", unit="chunks") as pbar:
            for result in pool.imap(process_chunk_with_office, chunk_data):
                for component_name, office_it in result:
                    component_results.append(component_name if component_name else "")
                    office_results.append(office_it if office_it else "")
                pbar.update(1)
    
    return component_results, office_results

def parallel_lookup(table_names, lookup_dict, n_processes=None):
    """
    Esegue il lookup in parallelo per migliorare le performance
    """
    if n_processes is None:
        n_processes = min(mp.cpu_count(), 4)  # Usa max 4 core per evitare overhead
    
    print(f"🚀 Avvio lookup parallelo con {n_processes} processi...")
    
    # Dividi i dati in chunks
    chunk_size = max(1000, len(table_names) // (n_processes * 4))  # Chunks non troppo piccoli
    chunks = [table_names[i:i + chunk_size] for i in range(0, len(table_names), chunk_size)]
    
    print(f"📊 Dati divisi in {len(chunks)} chunks di ~{chunk_size} elementi")
    
    # Prepara i dati per il multiprocessing
    chunk_data = [(chunk, lookup_dict) for chunk in chunks]
    
    # Processa in parallelo
    results = []
    with mp.Pool(processes=n_processes) as pool:
        with tqdm(desc="Lookup parallelo", unit="chunks") as pbar:
            for result in pool.imap(process_chunk, chunk_data):
                results.extend(result)
                pbar.update(1)
    
    return results
    """
    Cerca il nome del componente basato sul TABLE_NAME per la funzionalità Excel.
    Logica intelligente con priorità e concatenazione di risultati multipli.
    """
    if pd.isna(table_name) or not table_name:
        return ""
    
    # Pulisci il table_name e converti in stringa
    table_name = str(table_name).strip()
    
    # Rimuovi caratteri speciali e dividi per underscore
    clean_name = table_name.replace('$', '').replace('#', '')
    parts = [part for part in clean_name.split('_') if part]
    
    if not parts:
        return ""
    
    def search_in_prefisso_column(search_value):
        """Cerca un valore nella colonna Prefisso considerando che i valori sono separati da virgole"""
        all_matches = []
        
        # Controllo 1: cerca come valore esatto tra virgole
        pattern1 = f",{search_value},"
        matches1 = df_components[df_components['Prefisso'].str.contains(pattern1, case=False, na=False, regex=False)]
        if not matches1.empty:
            component_names = matches1['Nome_Componente'].tolist()
            all_matches.extend(component_names)
        
        # Controllo 2: cerca all'inizio seguito da virgola
        pattern2 = f"{search_value},"
        matches2 = df_components[df_components['Prefisso'].str.startswith(pattern2, na=False)]
        if not matches2.empty:
            component_names = matches2['Nome_Componente'].tolist()
            for name in component_names:
                if name not in all_matches:
                    all_matches.append(name)
        
        # Controllo 3: cerca alla fine preceduto da virgola
        pattern3 = f",{search_value}"
        matches3 = df_components[df_components['Prefisso'].str.endswith(pattern3, na=False)]
        if not matches3.empty:
            component_names = matches3['Nome_Componente'].tolist()
            for name in component_names:
                if name not in all_matches:
                    all_matches.append(name)
        
        # Controllo 4: cerca come valore unico (senza virgole)
        matches4 = df_components[df_components['Prefisso'].str.match(f"^{search_value}$", case=False, na=False)]
        if not matches4.empty:
            component_names = matches4['Nome_Componente'].tolist()
            for name in component_names:
                if name not in all_matches:
                    all_matches.append(name)
        
        # Se trovati risultati, concatena con virgole
        if all_matches:
            return ", ".join(all_matches)
        
        return None
    
    # Strategia 1: Cerca con i primi due elementi se disponibili
    if len(parts) >= 2:
        search_value_2 = f"{parts[0]}_{parts[1]}"
        result = search_in_prefisso_column(search_value_2)
        if result:
            return result
        
        # Strategia 1b: Cerca solo il secondo elemento se disponibile
        search_value_second = parts[1]
        result = search_in_prefisso_column(search_value_second)
        if result:
            return result
    
    # Strategia 2: Cerca solo con il primo elemento
    search_value_1 = parts[0]
    result = search_in_prefisso_column(search_value_1)
    if result:
        return result
    
    return ""


def process_file_to_excel_with_lookup(file_path, is_batch_operation=False, current_file=1, total_files=1):
    """
    Processa un file con formato specifico (header a riga variabile, separatori, dati)
    e crea un file Excel con lookup dei componenti
    
    Args:
        file_path: Percorso del file da elaborare
        is_batch_operation: Se True, è parte di un'operazione su più file
        current_file: Numero del file corrente (per operazioni multiple)
        total_files: Numero totale di file (per operazioni multiple)
    """
    try:
        print(f"🔄 Processando file per Excel: {file_path}")
        
        # Usa la funzione che gestisce file compressi
        content = read_file_content_with_progress(file_path)
        
        # Dividi il contenuto in righe
        lines = content.splitlines()
        
        print(f"📄 File letto: {len(lines)} righe totali")
        
        # Debug: mostra le prime righe per capire il formato
        print("\n🔍 Analisi formato file (prime 20 righe):")
        print("-" * 60)
        for i in range(min(20, len(lines))):
            line_preview = lines[i][:100] + "..." if len(lines[i]) > 100 else lines[i]
            print(f"{i+1:2d}: {line_preview}")
        print("-" * 60)
        
        # Verifica se il contenuto iniziale sembra corrotto
        initial_content_valid = is_content_valid('\n'.join(lines[:50]))
        if not initial_content_valid:
            print("⚠️  Le prime righe sembrano corrotte, cerco header più avanti nel file...")
        
        # Cerca diversi pattern di header in tutto il file (non solo all'inizio)
        header_line_idx = None
        header_patterns = [
            ('||', 'OWNER'),          # Pattern originale
            ('OWNER', 'OBJECT_NAME'), # Pattern specifico Oracle
            ('OWNER', 'TABLE_NAME'),  # Pattern semplice
            ('Owner', 'Table'),       # Pattern alternativo
            (',', 'OWNER'),           # CSV con OWNER
        ]
        
        # Se il contenuto iniziale è corrotto, inizia la ricerca più avanti
        search_start = 2000 if not initial_content_valid else 0
        search_end = min(len(lines), search_start + 15000)  # Cerca nelle prime 15k righe dopo il punto di inizio
        
        print(f"🔍 Ricerca header dalle righe {search_start+1} a {search_end}")
        
        for pattern_chars, pattern_word in header_patterns:
            for i in range(search_start, search_end):
                line = lines[i]
                if pattern_chars in line and pattern_word.upper() in line.upper():
                    # Per le righe header, usa una validazione più permissiva
                    if is_header_line_valid(line):
                        header_line_idx = i
                        print(f"✅ Header trovato con pattern '{pattern_chars}' + '{pattern_word}' alla riga {i+1}")
                        break
            if header_line_idx is not None:
                break
        
        # Trova la riga con i nomi delle colonne (cerca quella che contiene ||)
        if header_line_idx is None:
            print("❌ Impossibile trovare la riga con i nomi delle colonne")
            print("🔍 Righe che contengono parole chiave:")
            
            keywords = ['OWNER', 'TABLE', 'COLUMN', '||', '----', ',']
            for keyword in keywords:
                matching_lines = [(i+1, line[:100]) for i, line in enumerate(lines) if keyword.upper() in line.upper()]
                if matching_lines:
                    print(f"  '{keyword}': {len(matching_lines)} righe trovate")
                    for line_num, content in matching_lines[:3]:
                        print(f"    Riga {line_num}: {content}")
            
            return None
        
        print(f"📍 Header trovato alla riga {header_line_idx + 1}")
        
        # Estrai i nomi delle colonne dalla riga header
        header_line = lines[header_line_idx].strip()
        column_names = []
        parts = header_line.split("||','||")
        for part in parts:
            clean_part = part.replace("||','", "").replace("'", "").replace("||", "").strip()
            if clean_part:
                column_names.append(clean_part)
        
        print(f"📊 Colonne trovate: {column_names}")
        
        # Trova la riga con i separatori (-------)
        separator_line_idx = None
        for i in range(header_line_idx + 1, min(header_line_idx + 5, len(lines))):
            if '---' in lines[i]:
                separator_line_idx = i
                break
        
        if separator_line_idx is None:
            print("❌ Impossibile trovare la riga con i separatori")
            return None
        
        print(f"📍 Separatori trovati alla riga {separator_line_idx + 1}")
        
        # Leggi i dati a partire dalla riga dopo i separatori
        data_start_idx = separator_line_idx + 1
        data_lines = lines[data_start_idx:]
        
        # Processa i dati
        data_rows = []
        print(f"📖 Processando {len(data_lines)} righe di dati...")
        
        with tqdm(desc="Parsing dati", unit="righe") as pbar:
            for line in data_lines:
                line = line.strip()
                if not line:  # Salta righe vuote
                    continue
                
                # Controlla se è la riga finale "X rows selected."
                if " rows selected." in line.lower() or " row selected." in line.lower():
                    print(f"📋 Trovata riga finale Oracle: '{line}' - Fermando il parsing")
                    break
                
                # Dividi per virgola
                row_data = [col.strip() for col in line.split(',')]
                
                # Assicurati che il numero di colonne corrisponda
                while len(row_data) < len(column_names):
                    row_data.append('')
                
                # Tronca se ci sono troppe colonne
                row_data = row_data[:len(column_names)]
                
                data_rows.append(row_data)
                pbar.update(1)
        
        if not data_rows:
            print("❌ Nessuna riga di dati trovata")
            return None
        
        # Crea DataFrame
        print(f"📊 Creazione DataFrame con {len(data_rows)} righe...")
        df = pd.DataFrame(data_rows, columns=column_names)
        
        # Carica il file Components.csv con gestione percorsi multipli
        possible_paths = [
            "Components.csv",  # Directory corrente
            os.path.join(os.path.dirname(sys.executable), "Components.csv"),  # Stessa cartella dell'exe
            os.path.join(os.getcwd(), "Components.csv"),  # Directory di lavoro corrente
        ]
        
        # Se siamo in un eseguibile PyInstaller, prova anche il percorso temporaneo
        if getattr(sys, 'frozen', False):
            bundle_dir = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
            possible_paths.insert(0, os.path.join(bundle_dir, "Components.csv"))
        
        components_csv_path = None
        for path in possible_paths:
            if os.path.exists(path):
                components_csv_path = path
                break
        
        if components_csv_path is None:
            print(f"❌ File Components.csv non trovato!")
            print("💡 Assicurati che Components.csv sia nella stessa cartella dell'eseguibile")
            return None
        
        print("📚 Caricamento file Components.csv...")
        df_components = pd.read_csv(components_csv_path, sep=';')
        print(f"📊 Componenti caricati: {len(df_components)}")
        
        # Esegui il lookup dei componenti
        print("🔍 Esecuzione lookup componenti...")
        
        # Trova la colonna TABLE_NAME
        table_name_col = None
        for col in df.columns:
            if 'TABLE_NAME' in col.upper():
                table_name_col = col
                break
        
        if table_name_col is None:
            print("❌ Colonna TABLE_NAME non trovata")
            return None
        
        print(f"✅ Usando colonna '{table_name_col}' per il lookup")
        
        # Prepara i dizionari di lookup ottimizzati
        lookup_dict = prepare_lookup_dictionary(df_components)
        office_dict = prepare_office_dictionary(df_components)
        
        # Estrai i table names come lista
        table_names_list = df[table_name_col].tolist()
        
        # Scegli il metodo di lookup in base alla dimensione
        if len(table_names_list) > 10000:
            print("🚀 Dataset grande rilevato, usando lookup parallelo...")
            component_names, office_its = parallel_lookup_with_office(table_names_list, lookup_dict, office_dict)
        else:
            print("📊 Dataset piccolo, usando lookup sequenziale ottimizzato...")
            component_names = []
            office_its = []
            with tqdm(desc="Lookup componenti", unit="righe") as pbar:
                for table_name in table_names_list:
                    component_name, office_it = lookup_component_info_optimized(table_name, lookup_dict, office_dict)
                    component_names.append(component_name if component_name else "")
                    office_its.append(office_it if office_it else "")
                    pbar.update(1)
        
        # Aggiungi le colonne Component Name e Office IT
        df['Component Name'] = component_names
        df['Office IT'] = office_its
        
        # Ordinamento globale per Component Name (alfabetico)
        print("🔤 Ordinamento dati per Component Name...")
        # Ordina per Component Name, mettendo i valori vuoti alla fine
        df_sorted = df.sort_values(
            by='Component Name', 
            key=lambda x: x.fillna('zzz_vuoto'),  # I valori vuoti vanno alla fine
            ascending=True
        ).reset_index(drop=True)
        
        # Usa il DataFrame ordinato per il resto dell'elaborazione
        df = df_sorted
        
        # Statistiche del lookup
        successful_lookups = sum(1 for name in component_names if name)
        failed_lookups = sum(1 for name in component_names if not name)
        total_rows = len(component_names)
        success_rate = (successful_lookups / total_rows * 100) if total_rows > 0 else 0
        failure_rate = (failed_lookups / total_rows * 100) if total_rows > 0 else 0
        
        print(f"📊 Lookup completato:")
        print(f"   ✅ Righe riuscite: {successful_lookups:,} ({success_rate:.1f}%)")
        print(f"   ❌ Righe fallite: {failed_lookups:,} ({failure_rate:.1f}%)")
        
        # Gestione del file dei fallimenti
        create_failed_file = False
        if failed_lookups > 0:
            create_failed_file = ask_for_failed_lookup_file(failed_lookups, total_rows, is_batch_operation, current_file, total_files)
        
        # Crea il file Excel di output
        base_name = Path(file_path).stem
        today = datetime.now().strftime("%Y%m%d")
        timestamp = datetime.now().strftime("%H%M%S")
        
        # Percorso della cartella Downloads
        downloads_path = Path(os.path.expanduser("~")) / "Downloads" / "Lookup Component CDL"
        downloads_path.mkdir(parents=True, exist_ok=True)
        
        output_filename = f"{base_name}_lookup_{today}_{timestamp}.xlsx"
        output_path = downloads_path / output_filename
        
        # Identifica i lookup falliti per l'analisi solo se richiesto
        failed_table_names = None
        if create_failed_file and failed_lookups > 0:
            failed_table_names = [df.iloc[i][table_name_col] for i, name in enumerate(component_names) if not name]
            print("🔍 Preparazione analisi fallimenti...")
        
        print("📝 Creazione file Excel...")
        
        # Usa la nuova funzione di creazione Excel con supporto per suddivisione
        created_files = create_excel_with_split_support(
            df=df, 
            output_path=output_path, 
            failed_table_names=failed_table_names,
            lookup_dict=lookup_dict,
            table_name_col=table_name_col
        )
        
        # Gestisci output dei file creati
        if len(created_files) == 1:
            output_path = created_files[0]
            file_name = Path(output_path).name if isinstance(output_path, str) else output_path.name
            print(f"✅ File Excel creato: {file_name}")
        else:
            print(f"✅ {len(created_files)} file Excel creati")
            for i, file_path in enumerate(created_files, 1):
                file_name = Path(file_path).name if isinstance(file_path, str) else file_path.name
                print(f"   📄 File {i}: {file_name}")
        
        # Crea sempre archivi 7z per i file principali
        # Esclude solo i file di analisi fallimenti dalla zippatura
        base_archive_name = f"{base_name}_lookup_{today}_{timestamp}"
        created_archives = rename_and_create_7z_archives(created_files, base_archive_name)
        
        if created_archives:
            print(f"\n📦 {len(created_archives)} archivi 7z creati:")
            for i, archive_path in enumerate(created_archives, 1):
                archive_name = Path(archive_path).name
                print(f"   📦 Archivio {i}: {archive_name}")
        
        # Riepilogo finale
        print(f"\n🎯 RIEPILOGO")
        print(f"━━━━━━━━━━━━━━━━━━━━━━")
        print(f"� Righe elaborate: {total_rows:,}")
        print(f"✅ Lookup riusciti: {successful_lookups:,} ({success_rate:.1f}%)")
        if failed_lookups > 0:
            print(f"❌ Lookup falliti: {failed_lookups:,} ({failure_rate:.1f}%)")
            if create_failed_file:
                print(f"🔍 Analisi fallimenti inclusa")
        else:
            print(f"🎉 Tutti i lookup riusciti!")
        print(f"📁 File salvati in: Downloads/Lookup Component CDL/")
        
        # Restituisci il primo file se è uno solo, altrimenti la lista
        return created_files[0] if len(created_files) == 1 else created_files
        
    except Exception as e:
        print(f"❌ Errore durante l'elaborazione Excel: {e}")
        import traceback
        traceback.print_exc()
        return None


def show_main_menu():
    """Mostra il menu principale"""
    clear_console()
    show_header()
    print("📋 MENU PRINCIPALE")
    print("━" * 30)
    print("1. 📄 Elabora file Oracle (output Excel)")
    print("2. ❌ Esci")

def filter_multivolume_files(files):
    """
    Filtra i file multi-volume mostrando solo il primo di ogni serie
    e aggiunge informazioni sui file correlati
    
    Args:
        files (list): Lista di file Path objects
    
    Returns:
        tuple: (filtered_files, multivolume_info)
    """
    filtered_files = []
    multivolume_info = {}  # {index: [list_of_related_files]}
    processed_bases = set()
    
    for file_path in files:
        file_str = str(file_path)
        
        if is_multivolume_file(file_str):
            # Ottieni il nome base
            base_name = get_multivolume_base_name(file_str)
            
            if base_name not in processed_bases:
                # È il primo della serie, aggiungilo
                processed_bases.add(base_name)
                
                # Trova tutte le parti correlate
                related_files = find_multivolume_files(file_str, file_path.parent)
                
                # Aggiungi il file principale alla lista filtrata
                filtered_files.append(file_path)
                
                # Memorizza le informazioni sui file correlati
                if len(related_files) > 1:
                    multivolume_info[len(filtered_files) - 1] = related_files
            
            # Altrimenti salta questo file (è una parte successiva)
        else:
            # File normale, aggiungilo
            filtered_files.append(file_path)
    
    return filtered_files, multivolume_info

def show_file_selection_menu(files, dir_path):
    """Mostra il menu di selezione file con gestione multi-volume"""
    # Filtra i file multi-volume
    filtered_files, multivolume_info = filter_multivolume_files(files)
    
    clear_console()
    show_header()
    print("🎯 SELEZIONE FILE SPECIFICI")
    print("━" * 40)
    print(f"📁 Directory: {dir_path}")
    print(f"\n📂 File disponibili ({len(filtered_files)}):")
    print("─" * 60)
    
    for i, file in enumerate(filtered_files, 1):
        try:
            if i-1 in multivolume_info:
                # File multi-volume
                related_files = multivolume_info[i-1]
                total_size = sum(Path(f).stat().st_size for f in related_files if Path(f).exists())
                
                if total_size > 1024*1024:
                    size_info = f"({total_size/1024/1024:.1f} MB, {len(related_files)} parti)"
                elif total_size > 1024:
                    size_info = f"({total_size/1024:.1f} KB, {len(related_files)} parti)"
                else:
                    size_info = f"({total_size} bytes, {len(related_files)} parti)"
                
                print(f"  {i:2d}. {file.name} {size_info} 🔗")
            else:
                # File singolo
                size = file.stat().st_size
                if size > 1024*1024:
                    size_info = f"({size/1024/1024:.1f} MB)"
                elif size > 1024:
                    size_info = f"({size/1024:.1f} KB)"
                else:
                    size_info = f"({size} bytes)"
                
                print(f"  {i:2d}. {file.name} {size_info}")
        except:
            print(f"  {i:2d}. {file.name}")
    
    print(f"\n� ESEMPI DI SELEZIONE:")
    print(f"   • File singolo: 5")
    print(f"   • File multipli: 1,3,6,9")
    print(f"   • Range: 1-5 (file da 1 a 5)")
    print(f"   • Combinazione: 1,3,5-8,10")
    print(f"   • Tutti: 1-{len(filtered_files)}")
    print(f"   • Indietro: back o b")
    
    # Memorizza le informazioni per l'uso successivo
    return filtered_files, multivolume_info

# Variabile globale per ricordare l'ultimo path
last_used_path = None

def parse_file_selection(selection_str, max_files):
    """
    Parsa la selezione file inserita dall'utente.
    Supporta:
    - Numeri singoli: "1", "3", "7"
    - Lista: "1,3,6,9"
    - Range: "1-5", "3-8"
    - Combinazioni: "1,3,5-8,10"
    
    Returns:
        list: Lista di indici (1-based) selezionati, o None se errore
    """
    try:
        selected = set()
        
        # Dividi per virgola
        parts = [part.strip() for part in selection_str.split(',')]
        
        for part in parts:
            if '-' in part:
                # È un range
                range_parts = part.split('-')
                if len(range_parts) != 2:
                    return None
                    
                start, end = int(range_parts[0]), int(range_parts[1])
                if start < 1 or end > max_files or start > end:
                    return None
                    
                selected.update(range(start, end + 1))
            else:
                # È un numero singolo
                num = int(part)
                if num < 1 or num > max_files:
                    return None
                selected.add(num)
        
        return sorted(list(selected))
        
    except ValueError:
        return None

def handle_file_processing():
    """Gestisce la sezione di elaborazione file con navigazione migliorata"""
    global last_used_path
    
    while True:
        clear_console()
        show_header()
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

        user_input = safe_input("\n➤ Percorso: ").strip().strip('"').strip("'")
        
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
        last_used_path = str(dir_path)
        
        try:
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
            
            mode_choice = safe_input("\n➤ Scelta (1-3): ").strip()
            
            if mode_choice == "3":
                continue  # Torna alla selezione directory
            elif mode_choice == "1":
                # Elabora tutti i file
                print(f"\n🚀 Elaborazione di tutti i {len(files)} file...")
                print("━" * 50)
                
                # Reset della scelta globale per i file di fallimenti
                reset_failed_file_choice()
                
                total_processed = 0
                total_success = 0
                total_failed = 0
                
                for i, file_path in enumerate(files, 1):
                    print(f"\n📄 [{i}/{len(files)}] Elaborando: {file_path.name}")
                    print("⏳ Elaborazione in corso...")
                    
                    result = process_file_to_excel_with_lookup(file_path, is_batch_operation=True, current_file=i, total_files=len(files))
                    total_processed += 1
                    
                    if result:
                        total_success += 1
                        print(f"✅ Completato con successo")
                    else:
                        total_failed += 1
                        print(f"❌ Elaborazione fallita")
                
                # Riepilogo finale
                print(f"\n🎯 RIEPILOGO ELABORAZIONE MULTIPLA")
                print(f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
                print(f"📊 File processati: {total_processed}")
                print(f"✅ Successi: {total_success}")
                print(f"❌ Fallimenti: {total_failed}")
                print(f"📁 File salvati in: Downloads/Lookup Component CDL/")
                
                input(f"\n📝 Premi INVIO per continuare...")
                
            elif mode_choice == "2":
                # Selezione file specifici
                while True:
                    # Ottieni file filtrati e info multi-volume
                    filtered_files, multivolume_info = show_file_selection_menu_enhanced(files, dir_path)
                    
                    file_choice = safe_input("\n➤ Selezione: ").strip()
                    
                    # Controlla se l'utente vuole tornare indietro
                    if file_choice.lower() in ['back', 'b', 'indietro']:
                        break  # Torna alla selezione modalità
                    
                    # Parsa la selezione usando i file filtrati
                    selected_indices = parse_file_selection(file_choice, len(filtered_files))
                    
                    if selected_indices is None:
                        print("❌ Selezione non valida. Usa il formato: 1,3,5 oppure 1-5 oppure 1,3-6,8")
                        input("\n📝 Premi INVIO per riprovare...")
                        continue
                    
                    # Mostra la selezione usando i file filtrati
                    selected_files = [filtered_files[i-1] for i in selected_indices]
                    print(f"\n✅ File selezionati ({len(selected_files)}):")
                    for i, file_path in enumerate(selected_files, 1):
                        # Mostra informazioni multi-volume se applicabile
                        if (i-1) in [filtered_files.index(f) for f in filtered_files if filtered_files.index(f) in multivolume_info]:
                            mv_idx = filtered_files.index(file_path)
                            related_count = len(multivolume_info[mv_idx])
                            print(f"   {i}. {file_path.name} 🔗 ({related_count} parti)")
                        else:
                            print(f"   {i}. {file_path.name}")
                    
                    confirm = safe_input(f"\n🤔 Procedere con l'elaborazione di {len(selected_files)} file? (s/n): ").strip().lower()
                    
                    if confirm in ['s', 'si', 'y', 'yes']:
                        # Elabora i file selezionati
                        print(f"\n🚀 Elaborazione di {len(selected_files)} file selezionati...")
                        print("━" * 50)
                        
                        # Reset della scelta globale per i file di fallimenti
                        reset_failed_file_choice()
                        
                        total_processed = 0
                        total_success = 0
                        total_failed = 0
                        
                        for i, file_path in enumerate(selected_files, 1):
                            print(f"\n📄 [{i}/{len(selected_files)}] Elaborando: {file_path.name}")
                            print("⏳ Elaborazione in corso...")
                            
                            result = process_file_to_excel_with_lookup(file_path, is_batch_operation=True, current_file=i, total_files=len(selected_files))
                            total_processed += 1
                            
                            if result:
                                total_success += 1
                                print(f"✅ Completato con successo")
                            else:
                                total_failed += 1
                                print(f"❌ Elaborazione fallita")
                        
                        # Riepilogo finale
                        print(f"\n🎯 RIEPILOGO ELABORAZIONE MULTIPLA")
                        print(f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
                        print(f"📊 File processati: {total_processed}")
                        print(f"✅ Successi: {total_success}")
                        print(f"❌ Fallimenti: {total_failed}")
                        print(f"📁 File salvati in: Downloads/Lookup Component CDL/")
                        
                        input(f"\n📝 Premi INVIO per continuare...")
                        break  # Esce dal loop di selezione file
                    else:
                        print("❌ Operazione annullata")
                        input("\n📝 Premi INVIO per riprovare...")
            else:
                print("❌ Opzione non valida. Scegli 1, 2 o 3")
                input("\n📝 Premi INVIO per riprovare...")
                    
        except Exception as e:
            print(f"❌ Errore: {str(e)}")
            input("\n📝 Premi INVIO per riprovare...")

def show_file_selection_menu_enhanced(files, dir_path):
    """Mostra il menu di selezione file migliorato con supporto per selezioni multiple"""
    clear_console()
    show_header()
    print("🎯 SELEZIONE FILE SPECIFICI")
    print("━" * 40)
    print(f"📁 Directory: {dir_path}")
    print(f"\n� File disponibili ({len(files)}):")
    print("─" * 60)
    
    for i, file in enumerate(files, 1):
        # Calcola dimensione file
        try:
            size = file.stat().st_size
            if size > 1024*1024:
                size_info = f"({size/1024/1024:.1f} MB)"
            elif size > 1024:
                size_info = f"({size/1024:.1f} KB)"
            else:
                size_info = f"({size} bytes)"
        except:
            size_info = ""
        
        print(f"  {i:2d}. {file.name} {size_info}")
    
    print("\n💡 ESEMPI DI SELEZIONE:")
    print("   • File singolo: 5")
    print("   • File multipli: 1,3,6,9")
    print("   • Range: 1-5 (file da 1 a 5)")
    print("   • Combinazione: 1,3,5-8,10")
    print("   • Tutti: 1-" + str(len(files)))
    print("   • Indietro: back o b")
    
    # Per ora restituisce i file originali e un dizionario vuoto per evitare l'errore
    return files, {}

# Esempio di utilizzo
if __name__ == "__main__":
    # Protezione per multiprocessing su Windows
    mp.freeze_support()
    
    # Imposta il titolo della finestra
    set_console_title()
    
    # Mostra header
    show_header()
    
    # Carica il database componenti
    # Determina il percorso del file Components.csv
    if getattr(sys, 'frozen', False):
        # Siamo in un eseguibile PyInstaller - cerca nella directory dell'eseguibile
        components_csv_path = "Components.csv"  # Directory corrente
    else:
        # Siamo in modalità script - usa il percorso relativo
        components_csv_path = Path(__file__).parent / "Components.csv"
    
    df_components, lookup_dict = load_components(components_csv_path)
    
    if df_components is None:
        print("\n❌ Impossibile continuare senza il database componenti")
        input("\n📝 Premi INVIO per uscire...")
        sys.exit(1)
    
    # Loop principale del menu
    while True:
        show_main_menu()
        
        choice = safe_input("\n➤ Seleziona opzione (1-2): ").strip()
        
        if choice == "1":
            handle_file_processing()
            
        elif choice == "2":
            clear_console()
            print("👋 Arrivederci!")
            print("Grazie per aver usato Look-up Components CDL")
            break
            
        else:
            print("❌ Opzione non valida. Scegli 1 o 2")
            input("\n📝 Premi INVIO per continuare...")