"""
Database Service - Connessione Oracle TNS per Transaction Reporting
"""

import logging
import os
from typing import Optional, List, Dict, Any
from contextlib import contextmanager
from pathlib import Path

try:
    import oracledb
    ORACLEDB_AVAILABLE = True
    print(f"✅ DEBUG STARTUP: Modulo oracledb importato con successo - Versione: {getattr(oracledb, '__version__', 'Sconosciuta')}")
except ImportError as e:
    ORACLEDB_AVAILABLE = False
    print(f"❌ DEBUG STARTUP: Errore import oracledb: {e}")
except Exception as e:
    ORACLEDB_AVAILABLE = False
    print(f"❌ DEBUG STARTUP: Errore generico oracledb: {e}")


class DatabaseService:
    """Servizio per gestione connessioni Oracle TNS"""
    
    def __init__(self, tns_alias: str = "pporafin"):
        self.tns_alias = tns_alias
        self.logger = logging.getLogger(__name__)
        self._setup_tns_environment()
        
        # Host di fallback appropriati per l'ambiente
        if tns_alias == "pporafin":
            # Pre-produzione
            self.fallback_hosts = [
                "orafinprex-scan.sg.gbs.pro", 
                "orafinprey-scan.sg.gbs.pro", 
                "orafinprez-scan.sg.gbs.pro"
            ]
        elif tns_alias == "orafin":
            # Produzione  
            self.fallback_hosts = [
                "orafinpro-scan.sg.gbs.pro",
                "orafinbc-scan.sg.gbs.pro", 
                "orafingdr-scan.sg.gbs.pro"
            ]
        else:
            # Default per altri ambienti
            self.fallback_hosts = ["172.17.23.61", "172.17.23.62", "172.17.23.63"]
            
        self.successful_host = None
    
    def _setup_tns_environment(self):
        """Configura l'ambiente TNS Oracle con controllo prioritario del progetto"""
        # Percorso alla configurazione TNS locale al progetto
        project_root = Path(__file__).parent.parent
        tns_config_path = project_root / "oracle_config"
        
        # Salva il TNS_ADMIN originale per debug
        original_tns = os.environ.get('TNS_ADMIN')
        
        if tns_config_path.exists():
            tns_path = str(tns_config_path.absolute())
            os.environ['TNS_ADMIN'] = tns_path
            self.logger.info(f"TNS_ADMIN configurato: {tns_path}")
            if original_tns and original_tns != tns_path:
                self.logger.info(f"TNS_ADMIN precedente sovrascitto: {original_tns}")
        else:
            self.logger.warning(f"Configurazione TNS non trovata: {tns_config_path}")
            if original_tns:
                self.logger.warning(f"Utilizzo configurazione TNS di sistema: {original_tns}")
            else:
                self.logger.error("Nessuna configurazione TNS disponibile!")
    
    def test_connection(self) -> Dict[str, Any]:
        """Testa la connessione Oracle con fallback multi-host"""
        print(f"🔍 DEBUG: Verifica disponibilità modulo oracledb...")
        if not ORACLEDB_AVAILABLE:
            print(f"❌ DEBUG: Modulo oracledb non disponibile!")
            return {
                "success": False,
                "error": "Modulo oracledb non disponibile",
                "suggestion": "pip install oracledb"
            }
        else:
            print(f"✅ DEBUG: Modulo oracledb disponibile")
        
        # Debug configurazione TNS
        print(f"🔍 DEBUG: Verifica configurazione TNS...")
        tns_admin = os.environ.get("TNS_ADMIN", "Non configurato")
        print(f"🔍 DEBUG: TNS_ADMIN = {tns_admin}")
        
        if tns_admin != "Non configurato":
            try:
                tnsnames_path = Path(tns_admin) / "tnsnames.ora"
                if tnsnames_path.exists():
                    print(f"✅ DEBUG: File tnsnames.ora trovato: {tnsnames_path}")
                    # Leggi e mostra contenuto del tnsnames.ora per debug
                    try:
                        with open(tnsnames_path, 'r') as f:
                            content = f.read()
                            print(f"🔍 DEBUG: Contenuto tnsnames.ora:")
                            print(f"--- INIZIO TNSNAMES.ORA ---")
                            print(content[:500] + "..." if len(content) > 500 else content)
                            print(f"--- FINE TNSNAMES.ORA ---")
                    except Exception as e:
                        print(f"⚠️ DEBUG: Errore lettura tnsnames.ora: {e}")
                else:
                    print(f"❌ DEBUG: File tnsnames.ora non trovato: {tnsnames_path}")
            except Exception as e:
                print(f"⚠️ DEBUG: Errore accesso tnsnames.ora: {e}")
        else:
            print(f"⚠️ DEBUG: TNS_ADMIN non configurato")
        
        # Acquisizione credenziali diretta
        try:
            print("\n🔐 AUTENTICAZIONE DATABASE")
            print("Servizio: PPORAFIN")
            print("=" * 40)
            
            # Gestione encoding per Windows
            import sys
            
            username = input("👤 Username: ")
            # Pulizia rigorosa dell'input
            username = username.strip()
            # Rimuovi caratteri non ASCII che potrebbero causare problemi
            username = ''.join(char for char in username if ord(char) < 128)
            
            print(f"🔍 DEBUG: Username pulito: '{username}' (lunghezza: {len(username)})")
            
            if not username:
                return {
                    "success": False,
                    "error": "Username richiesto per connessione database"
                }
            
            import getpass
            try:
                password = getpass.getpass("🔑 Password: ")
            except:
                # Fallback se getpass non funziona nell'eseguibile
                password = input("🔑 Password: ")
            
            # Pulizia della password (ma mantieni caratteri speciali)
            password = password.strip()
            print(f"🔍 DEBUG: Password ricevuta (lunghezza: {len(password)})")
            
            if not password:
                return {
                    "success": False,
                    "error": "Password richiesta per connessione database"
                }
            
            credentials = (username, password)
            print("✅ Credenziali acquisite")
        except Exception as e:
            self.logger.error(f"Errore acquisizione credenziali: {e}")
            return {
                "success": False,
                "error": f"Errore acquisizione credenziali: {e}"
            }
        
        if not credentials:
            return {
                "success": False,
                "error": "Credenziali Oracle non configurate"
            }
        
        username, password = credentials
        last_error = None
        
        print(f"🔍 DEBUG: Credenziali ottenute - Username: {username}")
        print(f"🔍 DEBUG: Tentativo connessione con TNS alias: {self.tns_alias}")
        
        # Test TNS
        try:
            print(f"🔍 DEBUG: Provo connessione TNS...")
            connection = oracledb.connect(
                user=username,
                password=password,
                dsn=self.tns_alias
            )
            connection.close()
            self.successful_host = "TNS"
            print(f"✅ DEBUG: Connessione TNS riuscita!")
            return {
                "success": True,
                "method": "TNS",
                "alias": self.tns_alias,
                "message": f"Connessione TNS riuscita con alias '{self.tns_alias}'"
            }
        except Exception as e:
            last_error = str(e)
            print(f"❌ DEBUG: Connessione TNS fallita: {e}")
            self.logger.warning(f"Connessione TNS fallita: {e}")
        
        print(f"🔍 DEBUG: Provo connessioni dirette ai server...")
        # Test connessioni dirette
        for host in self.fallback_hosts:
            try:
                dsn = f"{host}:1521/OTH_ORAFIN.bsella.it"
                print(f"🔍 DEBUG: Tentativo connessione diretta a {dsn}")
                connection = oracledb.connect(
                    user=username,
                    password=password,
                    dsn=dsn
                )
                connection.close()
                self.successful_host = host
                print(f"✅ DEBUG: Connessione diretta riuscita con {host}!")
                return {
                    "success": True,
                    "method": "Direct",
                    "host": host,
                    "dsn": dsn,
                    "message": f"Connessione diretta riuscita con {host}"
                }
            except Exception as e:
                last_error = str(e)
                print(f"❌ DEBUG: Connessione fallita con {host}: {e}")
                self.logger.warning(f"Connessione fallita con {host}: {e}")
                continue
        
        return {
            "success": False,
            "error": f"Tutti i metodi falliti. Ultimo errore: {last_error}"
        }
    
    @contextmanager
    def get_connection(self):
        """Context manager per connessione Oracle"""
        if not ORACLEDB_AVAILABLE:
            raise Exception("Modulo oracledb non disponibile")
        
        # Acquisizione credenziali diretta
        try:
            print("\n🔐 AUTENTICAZIONE DATABASE")
            print("Servizio: PPORAFIN")
            print("=" * 40)
            
            # Gestione encoding per Windows
            import sys
            
            username = input("👤 Username: ")
            # Pulizia rigorosa dell'input
            username = username.strip()
            # Rimuovi caratteri non ASCII che potrebbero causare problemi
            username = ''.join(char for char in username if ord(char) < 128)
            
            print(f"🔍 DEBUG: Username pulito: '{username}' (lunghezza: {len(username)})")
            
            if not username:
                raise Exception("Username richiesto per connessione database")
            
            import getpass
            try:
                password = getpass.getpass("🔑 Password: ")
            except:
                # Fallback se getpass non funziona nell'eseguibile
                password = input("🔑 Password: ")
            
            # Pulizia della password (ma mantieni caratteri speciali)
            password = password.strip()
            print(f"🔍 DEBUG: Password ricevuta (lunghezza: {len(password)})")
            
            if not password:
                raise Exception("Password richiesta per connessione database")
            
            credentials = (username, password)
            print("✅ Credenziali acquisite")
        except Exception as e:
            self.logger.error(f"Errore acquisizione credenziali: {e}")
            raise Exception(f"Errore acquisizione credenziali: {e}")
        
        if not credentials:
            raise Exception("Credenziali Oracle non configurate")
        
        username, password = credentials
        connection = None
        
        try:
            # Prova TNS
            if not self.successful_host or self.successful_host == "TNS":
                try:
                    connection = oracledb.connect(
                        user=username,
                        password=password,
                        dsn=self.tns_alias
                    )
                    self.successful_host = "TNS"
                    self.logger.info(f"Connessione TNS aperta: {self.tns_alias}")
                    yield connection
                    return
                except Exception as e:
                    self.logger.warning(f"Connessione TNS fallita: {e}")
            
            # Prova connessioni dirette
            hosts_to_try = self.fallback_hosts
            if self.successful_host and self.successful_host in self.fallback_hosts:
                hosts_to_try = [self.successful_host] + [h for h in self.fallback_hosts if h != self.successful_host]
            
            for host in hosts_to_try:
                try:
                    dsn = f"{host}:1521/OTH_ORAFIN.bsella.it"
                    connection = oracledb.connect(
                        user=username,
                        password=password,
                        dsn=dsn
                    )
                    self.successful_host = host
                    self.logger.info(f"Connessione diretta aperta: {dsn}")
                    yield connection
                    return
                except Exception as e:
                    self.logger.warning(f"Connessione fallita con {host}: {e}")
                    continue
            
            raise Exception("Nessun metodo di connessione disponibile")
            
        finally:
            if connection:
                try:
                    connection.close()
                    self.logger.info("Connessione Oracle chiusa")
                except Exception as e:
                    self.logger.error(f"Errore chiusura connessione: {e}")
    
    def get_order_status(self, order_numbers: List[str]) -> Dict[str, str]:
        """Ottiene lo status degli ordini dal database"""
        if not order_numbers:
            return {}
        
        placeholders = ', '.join([f":order_{i}" for i in range(len(order_numbers))])
        
        query = f"""
        SELECT 
            NUMERO_ORDINE,
            STATUS
        FROM ORDINI 
        WHERE NUMERO_ORDINE IN ({placeholders})
        """
        
        params = {f"order_{i}": order_num for i, order_num in enumerate(order_numbers)}
        
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query, params)
                
                status_map = {}
                for row in cursor.fetchall():
                    status_map[row[0]] = row[1]
                
                self.logger.info(f"Status recuperati per {len(status_map)}/{len(order_numbers)} ordini")
                return status_map
                
        except Exception as e:
            self.logger.error(f"Errore recupero status ordini: {e}")
            raise


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    
    db_service = DatabaseService()
    
    print("Test connessione Oracle...")
    test_result = db_service.test_connection()
    
    if test_result["success"]:
        print(f" {test_result['message']}")
    else:
        print(f" {test_result['error']}")
        if "suggestion" in test_result:
            print(f" {test_result['suggestion']}")
