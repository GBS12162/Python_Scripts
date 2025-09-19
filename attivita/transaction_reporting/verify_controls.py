#!/usr/bin/env python3
"""
Script per verificare i controlli di qualità nel file Excel generato
"""

import pandas as pd
import sys
from pathlib import Path

def verify_excel_controls(file_path):
    """Verifica i controlli nel file Excel"""
    print(f"🔍 Verifica controlli nel file: {file_path}")
    
    try:
        # Leggi il file Excel
        df = pd.read_excel(file_path)
        print(f"📊 Righe totali: {len(df)}")
        print(f"📊 Colonne: {list(df.columns)}")
        
        # Verifica presenza colonne controlli
        casistica_cols = [col for col in df.columns if 'CASISTICA' in str(col).upper()]
        print(f"\n🎯 Colonne CASISTICA trovate: {casistica_cols}")
        
        # Verifica controllo MIC_CODE_NON_PRESENTE
        mic_code_cols = [col for col in df.columns if 'MIC_CODE_NON_PRESENTE' in str(col).upper() or 'MIC CODE NON PRESENTE' in str(col).upper()]
        print(f"🎯 Colonne MIC_CODE_NON_PRESENTE: {mic_code_cols}")
        
        # Mostra le prime righe per verifica
        print(f"\n📋 Prime 10 righe:")
        print(df.head(10).to_string())
        
        # Controlla se ci sono valori nei controlli
        for col in casistica_cols + mic_code_cols:
            non_null_count = df[col].notna().sum()
            print(f"\n📊 Colonna '{col}': {non_null_count} valori non nulli")
            if non_null_count > 0:
                unique_values = df[col].dropna().unique()
                print(f"   Valori presenti: {unique_values}")
        
        return True
        
    except Exception as e:
        print(f"❌ Errore lettura file: {e}")
        return False

if __name__ == "__main__":
    # Cerca l'ultimo file generato
    output_dir = Path("C:/Dev Projects/Python Scripts/attivita/transaction_reporting/output/con412_reports")
    
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        # Trova l'ultimo file CON-412_Validated
        excel_files = list(output_dir.glob("CON-412_Validated_*.xlsx"))
        if not excel_files:
            print("❌ Nessun file CON-412_Validated trovato")
            sys.exit(1)
        
        # Ordina per data di modifica e prendi l'ultimo
        file_path = max(excel_files, key=lambda x: x.stat().st_mtime)
    
    print(f"🎯 File da verificare: {file_path}")
    success = verify_excel_controls(file_path)
    
    if success:
        print("\n✅ Verifica completata")
    else:
        print("\n❌ Verifica fallita")
        sys.exit(1)