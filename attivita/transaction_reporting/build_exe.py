"""
Script per creare l'eseguibile di Transaction Reporting - Rejecting Mensile.
Utilizza PyInstaller per generare un file .exe standalone.
"""

import os
import sys
import subprocess
from pathlib import Path
import shutil

def create_executable():
    """Crea l'eseguibile dell'applicazione."""
    
    print("=" * 70)
    print("    CREAZIONE ESEGUIBILE TRANSACTION REPORTING")
    print("=" * 70)
    print()
    
    # Verifica che PyInstaller sia installato
    try:
        import PyInstaller
        print(f"✓ PyInstaller trovato: {PyInstaller.__version__}")
    except ImportError:
        print("✗ PyInstaller non trovato!")
        print("Installare con: pip install pyinstaller")
        return False
    
    # Directory corrente
    current_dir = Path(__file__).parent
    launcher_script = current_dir / "launcher.py"
    
    if not launcher_script.exists():
        print(f"✗ Script launcher non trovato: {launcher_script}")
        return False
    
    print(f"✓ Script launcher: {launcher_script}")
    
    # Nome eseguibile
    exe_name = "Transaction_Reporting_Rejecting_Mensile"
    
    # Opzioni PyInstaller
    pyinstaller_options = [
        "pyinstaller",
        "--onefile",                    # File singolo
        "--console",                    # Mantieni console
        "--name", exe_name,             # Nome eseguibile
        "--clean",                      # Pulisci cache
        "--noconfirm",                  # Non chiedere conferma
        "--distpath", "dist",           # Directory di output
        "--workpath", "build",          # Directory di lavoro temporanea
        "--paths", str(current_dir),    # Aggiungi percorso corrente
        "--paths", str(current_dir.parent.parent),  # Root del progetto
        "--collect-all", "models",      # Includi tutti i moduli
        "--collect-all", "services",
        "--collect-all", "utils", 
        "--collect-all", "ui",
        "--collect-all", "pandas",      # Includi pandas
        "--collect-all", "openpyxl",    # Includi openpyxl
        "--collect-all", "xlsxwriter",  # Includi xlsxwriter
        "--hidden-import", "pandas",
        "--hidden-import", "openpyxl",
        "--hidden-import", "xlsxwriter",
        "--hidden-import", "decimal",
        "--hidden-import", "uuid",
        "--hidden-import", "json",
        str(launcher_script)            # Script launcher
    ]
    
    # Aggiungi file di esempio se esistono
    sample_files = [
        current_dir / "data" / "sample_transactions.csv",
        current_dir / "README.md"
    ]
    
    for sample_file in sample_files:
        if sample_file.exists():
            pyinstaller_options.extend([
                "--add-data", f"{sample_file};data" if sample_file.name.endswith('.csv') else f"{sample_file};."
            ])
            print(f"✓ Aggiunto file: {sample_file.name}")
    
    # Esegui PyInstaller
    print("\\n" + "-" * 50)
    print("Avvio PyInstaller...")
    print("-" * 50)
    
    try:
        # Cambia directory per evitare problemi di percorsi
        original_cwd = os.getcwd()
        os.chdir(current_dir)
        
        # Esegui PyInstaller
        result = subprocess.run(pyinstaller_options, capture_output=True, text=True)
        
        if result.returncode == 0:
            print("✓ PyInstaller completato con successo!")
            
            # Verifica che l'eseguibile sia stato creato
            exe_path = current_dir / "dist" / f"{exe_name}.exe"
            if exe_path.exists():
                exe_size_mb = exe_path.stat().st_size / (1024 * 1024)
                print(f"✓ Eseguibile creato: {exe_path}")
                print(f"✓ Dimensione: {exe_size_mb:.1f} MB")
                
                # Crea directory necessarie nella dist
                dist_dirs = ["output", "data", "logs"]
                for dir_name in dist_dirs:
                    dist_dir = current_dir / "dist" / dir_name
                    dist_dir.mkdir(exist_ok=True)
                    print(f"✓ Directory creata: {dir_name}/")
                
                # Crea file README per la distribuzione
                create_distribution_readme(current_dir / "dist")
                
                print("\\n" + "=" * 60)
                print("ESEGUIBILE CREATO CON SUCCESSO!")
                print("=" * 60)
                print(f"Percorso: {exe_path}")
                print(f"Per utilizzare l'eseguibile:")
                print(f"1. Vai nella cartella: {current_dir / 'dist'}")
                print(f"2. Esegui: {exe_name}.exe")
                print(f"3. L'applicazione creerà automaticamente le directory necessarie")
                print("=" * 60)
                
                return True
            else:
                print("✗ Eseguibile non trovato nella directory dist")
                return False
        else:
            print("✗ Errore durante l'esecuzione di PyInstaller:")
            print(result.stderr)
            return False
            
    except Exception as e:
        print(f"✗ Errore: {e}")
        return False
    finally:
        # Ripristina directory originale
        os.chdir(original_cwd)

def create_distribution_readme(dist_dir: Path):
    """Crea un file README per la distribuzione."""
    readme_content = '''# Transaction Reporting - Rejecting Mensile

## Descrizione
Sistema per l'analisi e reporting delle transazioni rifiutate su base mensile.

## Come utilizzare
1. Esegui Transaction_Reporting_Rejecting_Mensile.exe
2. Segui le istruzioni nell'interfaccia interattiva
3. I report generati saranno salvati nella cartella 'output/'

## Formati supportati
- File di input: CSV, Excel (.xlsx, .xls)
- File di output: Excel, CSV, JSON

## Directory
- output/: Report generati
- data/: File di dati di input (opzionale)
- logs/: File di log dell'applicazione

## Requisiti dati
Il file di transazioni deve contenere almeno:
- transaction_id: ID univoco della transazione
- account_number: Numero del conto
- amount: Importo
- currency: Valuta
- transaction_date: Data transazione
- transaction_type: Tipo transazione
- status: Stato (APPROVED, REJECTED, FAILED, ecc.)

## Supporto
Sviluppato da: GBS12162 (Samuele De Giosa)
Versione: 1.0.0
Data: 2025-09-16
'''
    
    readme_file = dist_dir / "README.txt"
    with open(readme_file, 'w', encoding='utf-8') as f:
        f.write(readme_content)
    
    print("✓ README.txt creato per la distribuzione")

def cleanup_build_files():
    """Pulisce i file di build temporanei."""
    current_dir = Path(__file__).parent
    
    # Directory da pulire
    dirs_to_clean = ["build", "__pycache__"]
    files_to_clean = ["*.spec"]
    
    for dir_name in dirs_to_clean:
        dir_path = current_dir / dir_name
        if dir_path.exists():
            try:
                shutil.rmtree(dir_path)
                print(f"✓ Rimossa directory: {dir_path}")
            except Exception as e:
                print(f"⚠ Impossibile rimuovere {dir_path}: {e}")
    
    # Rimuovi file .spec
    for spec_file in current_dir.glob("*.spec"):
        try:
            spec_file.unlink()
            print(f"✓ Rimosso file: {spec_file}")
        except Exception as e:
            print(f"⚠ Impossibile rimuovere {spec_file}: {e}")

def main():
    """Funzione principale."""
    try:
        # Crea eseguibile
        success = create_executable()
        
        if success:
            # Chiedi se pulire i file temporanei
            print("\\nPulire i file di build temporanei? (s/n): ", end="")
            response = input().lower().strip()
            
            if response in ['s', 'si', 'sì', 'y', 'yes']:
                print("\\nPulizia file temporanei...")
                cleanup_build_files()
        
        print(f"\\nOperazione {'completata' if success else 'fallita'}.")
        
    except KeyboardInterrupt:
        print("\\nOperazione interrotta dall'utente.")
    except Exception as e:
        print(f"Errore: {e}")
    
    input("\\nPremi INVIO per uscire...")

if __name__ == "__main__":
    main()