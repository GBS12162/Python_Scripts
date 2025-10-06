#!/usr/bin/env python3
"""
Test script per verificare la connessione al database Oracle
"""

import sys
import os
from pathlib import Path

# Aggiunge il percorso del progetto al PYTHONPATH
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

def test_oracledb_import():
    """Test dell'import del modulo oracledb"""
    print("=" * 60)
    print("🧪 TEST 1: Import modulo oracledb")
    print("=" * 60)
    
    try:
        import oracledb
        print(f"✅ Modulo oracledb importato con successo")
        print(f"📦 Versione: {getattr(oracledb, '__version__', 'Sconosciuta')}")
        print(f"📍 Percorso: {oracledb.__file__}")
        return True
    except ImportError as e:
        print(f"❌ Errore import oracledb: {e}")
        print(f"💡 Suggerimento: pip install oracledb")
        return False
    except Exception as e:
        print(f"❌ Errore generico oracledb: {e}")
        return False

def test_tns_configuration():
    """Test della configurazione TNS"""
    print("\n" + "=" * 60)
    print("🧪 TEST 2: Configurazione TNS")
    print("=" * 60)
    
    # Configura TNS_ADMIN
    project_root = Path(__file__).parent
    oracle_config_dir = project_root / "attivita" / "transaction_reporting" / "oracle_config"
    tnsnames_path = oracle_config_dir / "tnsnames.ora"
    
    print(f"📁 Directory oracle_config: {oracle_config_dir}")
    print(f"📄 File tnsnames.ora: {tnsnames_path}")
    
    if tnsnames_path.exists():
        print(f"✅ File tnsnames.ora trovato")
        os.environ["TNS_ADMIN"] = str(oracle_config_dir)
        print(f"🔧 TNS_ADMIN configurato: {os.environ.get('TNS_ADMIN')}")
        
        # Mostra contenuto del file
        try:
            with open(tnsnames_path, 'r') as f:
                content = f.read()
                print(f"\n📋 Contenuto tnsnames.ora:")
                print("-" * 40)
                print(content)
                print("-" * 40)
        except Exception as e:
            print(f"⚠️ Errore lettura tnsnames.ora: {e}")
        
        return True
    else:
        print(f"❌ File tnsnames.ora non trovato")
        print(f"💡 Crea il file con la configurazione TNS appropriata")
        return False

def test_credentials():
    """Test dell'acquisizione credenziali"""
    print("\n" + "=" * 60)
    print("🧪 TEST 3: Acquisizione credenziali")
    print("=" * 60)
    
    try:
        username = input("👤 Username Oracle: ").strip()
        if not username:
            print("❌ Username vuoto")
            return None
        
        import getpass
        try:
            password = getpass.getpass("🔑 Password Oracle: ")
        except:
            password = input("🔑 Password Oracle (visibile): ")
        
        if not password:
            print("❌ Password vuota")
            return None
        
        print(f"✅ Credenziali acquisite - Username: {username}")
        return (username, password)
        
    except Exception as e:
        print(f"❌ Errore acquisizione credenziali: {e}")
        return None

def test_database_connection(credentials):
    """Test della connessione al database"""
    print("\n" + "=" * 60)
    print("🧪 TEST 4: Connessione database")
    print("=" * 60)
    
    if not credentials:
        print("❌ Nessuna credenziale disponibile")
        return False
    
    username, password = credentials
    
    try:
        import oracledb
        
        # Test 1: Connessione TNS
        print(f"\n🔍 Test connessione TNS...")
        try:
            tns_alias = "pporafin"
            print(f"   Alias TNS: {tns_alias}")
            print(f"   Username: {username}")
            
            connection = oracledb.connect(
                user=username,
                password=password,
                dsn=tns_alias
            )
            
            # Test query semplice
            cursor = connection.cursor()
            cursor.execute("SELECT 1 FROM DUAL")
            result = cursor.fetchone()
            cursor.close()
            connection.close()
            
            print(f"✅ Connessione TNS riuscita!")
            print(f"✅ Query test: {result}")
            return True
            
        except Exception as e:
            print(f"❌ Connessione TNS fallita: {e}")
        
        # Test 2: Connessioni dirette
        print(f"\n🔍 Test connessioni dirette...")
        fallback_hosts = ["172.17.23.61", "172.17.23.62", "172.17.23.63"]
        
        for host in fallback_hosts:
            try:
                dsn = f"{host}:1521/OTH_ORAFIN.bsella.it"
                print(f"   Tentativo: {dsn}")
                
                connection = oracledb.connect(
                    user=username,
                    password=password,
                    dsn=dsn
                )
                
                # Test query semplice
                cursor = connection.cursor()
                cursor.execute("SELECT 1 FROM DUAL")
                result = cursor.fetchone()
                cursor.close()
                connection.close()
                
                print(f"✅ Connessione diretta riuscita con {host}!")
                print(f"✅ Query test: {result}")
                return True
                
            except Exception as e:
                print(f"❌ Connessione fallita con {host}: {e}")
        
        print(f"❌ Tutti i tentativi di connessione falliti")
        return False
        
    except Exception as e:
        print(f"❌ Errore generale test connessione: {e}")
        return False

def main():
    """Funzione principale"""
    print("🔧 TEST CONNESSIONE DATABASE ORACLE")
    print("Questo script testa tutti i componenti della connessione database")
    
    # Test 1: Import oracledb
    if not test_oracledb_import():
        print("\n❌ FALLIMENTO: Modulo oracledb non disponibile")
        return 1
    
    # Test 2: Configurazione TNS
    tns_ok = test_tns_configuration()
    if not tns_ok:
        print("\n⚠️ WARNING: Configurazione TNS non trovata, si userà solo connessione diretta")
    
    # Test 3: Acquisizione credenziali
    credentials = test_credentials()
    if not credentials:
        print("\n❌ FALLIMENTO: Impossibile acquisire credenziali")
        return 1
    
    # Test 4: Connessione database
    if test_database_connection(credentials):
        print("\n" + "=" * 60)
        print("🎉 TUTTI I TEST SUPERATI!")
        print("✅ La connessione database funziona correttamente")
        print("=" * 60)
        return 0
    else:
        print("\n" + "=" * 60)
        print("❌ FALLIMENTO: Connessione database non riuscita")
        print("🔍 Verifica:")
        print("   • Credenziali corrette")
        print("   • Rete accessibile")
        print("   • Server Oracle attivo")
        print("   • Firewall configurato")
        print("=" * 60)
        return 1

if __name__ == "__main__":
    sys.exit(main())