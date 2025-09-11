"""
Script per creare l'eseguibile di Oracle Component Lookup.
Utilizza PyInstaller per generare un file .exe standalone.
"""

import os
import sys
import subprocess
from pathlib import Path
import shutil

def create_executable():
    """Crea l'eseguibile dell'applicazione."""
    
    print("=" * 60)
    print("    CREAZIONE ESEGUIBILE ORACLE COMPONENT LOOKUP")
    print("=" * 60)
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
    exe_name = "Oracle_Component_Lookup"
    
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
        "--collect-all", "models",      # Includi tutti i moduli
        "--collect-all", "services",
        "--collect-all", "utils", 
        "--collect-all", "ui",
        str(launcher_script)            # Script launcher
    ]
    
    # Aggiungi dati necessari
    components_file = current_dir / "Components.csv"
    if components_file.exists():
        pyinstaller_options.extend([
            "--add-data", f"{components_file};."
        ])
        print(f"✓ Aggiunto Components.csv")
    
    # Esegui PyInstaller
    print("\n" + "-" * 40)
    print("Avvio PyInstaller...")
    print("-" * 40)
    
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
                
                # Copia Components.csv nella directory dist se necessario
                dist_components = current_dir / "dist" / "Components.csv"
                if components_file.exists() and not dist_components.exists():
                    shutil.copy2(components_file, dist_components)
                    print("✓ Components.csv copiato in dist/")
                
                # Crea directory Output nella dist
                dist_output_dir = current_dir / "dist" / "Output"
                dist_output_dir.mkdir(exist_ok=True)
                print("✓ Directory Output creata in dist/")
                
                print("\n" + "=" * 50)
                print("ESEGUIBILE CREATO CON SUCCESSO!")
                print("=" * 50)
                print(f"Percorso: {exe_path}")
                print(f"Per utilizzare l'eseguibile:")
                print(f"1. Vai nella cartella: {current_dir / 'dist'}")
                print(f"2. Esegui: {exe_name}.exe")
                print(f"3. Assicurati che Components.csv sia presente")
                print("=" * 50)
                
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
            print("\nPulire i file di build temporanei? (s/n): ", end="")
            response = input().lower().strip()
            
            if response in ['s', 'si', 'sì', 'y', 'yes']:
                print("\nPulizia file temporanei...")
                cleanup_build_files()
        
        print(f"\nOperazione {'completata' if success else 'fallita'}.")
        
    except KeyboardInterrupt:
        print("\nOperazione interrotta dall'utente.")
    except Exception as e:
        print(f"Errore: {e}")
    
    input("\nPremi INVIO per uscire...")

if __name__ == "__main__":
    main()
