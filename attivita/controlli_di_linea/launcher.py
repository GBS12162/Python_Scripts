"""
Launcher per l'applicazione Oracle Component Lookup.
Questo file risolve i problemi di import con PyInstaller.
"""

import sys
import os
from pathlib import Path

def setup_paths():
    """Configura i percorsi per i moduli."""
    # Se siamo in un eseguibile PyInstaller
    if getattr(sys, 'frozen', False):
        # Percorso dell'eseguibile
        base_path = Path(sys.executable).parent
    else:
        # Percorso dello script
        base_path = Path(__file__).parent
    
    # Aggiungi il percorso base al PYTHONPATH
    sys.path.insert(0, str(base_path))
    
    return base_path

def main():
    """Funzione principale del launcher."""
    try:
        # Configura percorsi
        base_path = setup_paths()
        
        print("=" * 60)
        print("    ORACLE COMPONENT LOOKUP - VERSIONE MODULARE")
        print("=" * 60)
        print()
        
        # Import dinamici dopo aver configurato i percorsi
        try:
            # Importa tutti i moduli necessari
            import pandas as pd
            import py7zr
            import tqdm
            import chardet
            
            # Verifica che i moduli dell'applicazione siano importabili
            from models.config import Config
            from services.component_service import ComponentService
            from ui.console_ui import ConsoleUI
            
            print("✓ Tutti i moduli caricati correttamente")
            
        except ImportError as e:
            print(f"✗ Errore nell'importazione dei moduli: {e}")
            print(f"Percorso base: {base_path}")
            print(f"Percorsi Python: {sys.path[:3]}...")
            input("\nPremi INVIO per uscire...")
            return
        
        # Avvia l'applicazione principale
        from main import OracleComponentLookupApp
        
        app = OracleComponentLookupApp()
        app.run()
        
    except KeyboardInterrupt:
        print("\nApplicazione interrotta dall'utente.")
    except Exception as e:
        print(f"\nERRORE: {e}")
        import traceback
        traceback.print_exc()
        input("\nPremi INVIO per uscire...")

if __name__ == "__main__":
    main()
