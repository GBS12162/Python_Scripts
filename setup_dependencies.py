"""
Script di installazione dipendenze per Python Scripts Repository
Gestisce l'installazione stratificata di dipendenze comuni + specifiche
"""

import subprocess
import sys
import os
from pathlib import Path

def run_command(command, description):
    """Esegue un comando e gestisce gli errori"""
    print(f"🔄 {description}...")
    try:
        result = subprocess.run(command, shell=True, capture_output=True, text=True)
        if result.returncode == 0:
            print(f"✅ {description} completato")
            return True
        else:
            print(f"❌ Errore {description}: {result.stderr}")
            return False
    except Exception as e:
        print(f"❌ Errore {description}: {e}")
        return False

def main():
    """Installazione intelligente delle dipendenze"""
    print("=" * 60)
    print("🐍 PYTHON SCRIPTS REPOSITORY - SETUP DIPENDENZE")
    print("=" * 60)
    
    # Verifica environment virtuale
    if not hasattr(sys, 'real_prefix') and not (hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix):
        print("⚠️ ATTENZIONE: Non sei in un environment virtuale!")
        response = input("Continuare comunque? (y/N): ")
        if response.lower() != 'y':
            print("❌ Installazione annullata")
            return
    
    print("\n1️⃣ Installazione dipendenze COMUNI...")
    if not run_command("pip install -r requirements.txt", "Dipendenze comuni"):
        return
    
    print("\n🔍 Ricerca progetti con dipendenze specifiche...")
    
    # Cerca tutti i requirements.txt nelle sottocartelle
    project_requirements = []
    for root, dirs, files in os.walk("attivita"):
        if "requirements.txt" in files:
            project_path = Path(root)
            project_name = project_path.name
            req_file = project_path / "requirements.txt"
            project_requirements.append((project_name, str(req_file)))
    
    if project_requirements:
        print(f"\n📦 Trovati {len(project_requirements)} progetti con dipendenze specifiche:")
        for i, (name, path) in enumerate(project_requirements, 1):
            print(f"   {i}. {name} ({path})")
        
        print("\n2️⃣ Installazione dipendenze SPECIFICHE...")
        for name, req_file in project_requirements:
            run_command(f"pip install -r \"{req_file}\"", f"Dipendenze {name}")
    
    print("\n✅ Setup completato!")
    print("\n📊 Packages installati:")
    run_command("pip list", "Lista packages")

if __name__ == "__main__":
    main()