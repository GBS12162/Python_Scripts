"""
Launcher per Transaction Reporting - Rejecting Mensile
Script di avvio ottimizzato per l'eseguibile e l'uso diretto.

Autore: GBS12162 (Samuele De Giosa)
Data: 2025-09-16
"""

import sys
import os
from pathlib import Path

def main():
    """Launcher principale per l'applicazione."""
    try:
        # Aggiungi il percorso corrente al PYTHONPATH
        current_dir = Path(__file__).parent
        if str(current_dir) not in sys.path:
            sys.path.insert(0, str(current_dir))
        
        # Importa e avvia l'applicazione principale
        from main import main as app_main
        app_main()
        
    except ImportError as e:
        print(f"❌ Errore nell'importazione: {e}")
        print("Assicurati che tutti i moduli necessari siano installati.")
        input("Premi INVIO per uscire...")
        sys.exit(1)
        
    except Exception as e:
        print(f"❌ Errore nell'avvio dell'applicazione: {e}")
        input("Premi INVIO per uscire...")
        sys.exit(1)

if __name__ == "__main__":
    main()