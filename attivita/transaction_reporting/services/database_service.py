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
except ImportError:
    ORACLEDB_AVAILABLE = False


class DatabaseService:
    """Servizio per gestione connessioni Oracle TNS"""
    
    def __init__(self, tns_alias: str = "pporafin"):
        self.tns_alias = tns_alias
        self.logger = logging.getLogger(__name__)
        self._setup_tns_environment()
        self.fallback_hosts = ["localhost", "pporafin", "pporafin.bsella.it"]
        self.successful_host = None
    
    def _setup_tns_environment(self):
        """Configura l'ambiente TNS con file locale"""
        try:
            project_root = Path(__file__).parent.parent
            oracle_config_dir = project_root / "oracle_config"
            tnsnames_path = oracle_config_dir / "tnsnames.ora"
            
            if tnsnames_path.exists():
                os.environ["TNS_ADMIN"] = str(oracle_config_dir)
                self.logger.info(f"TNS_ADMIN configurato: {oracle_config_dir}")
            else:
                self.logger.warning(f"File tnsnames.ora non trovato: {tnsnames_path}")
                
        except Exception as e:
            self.logger.error(f"Errore configurazione TNS: {e}")
    
    def test_connection(self) -> Dict[str, Any]:
        """Testa la connessione Oracle con fallback multi-host"""
        if not ORACLEDB_AVAILABLE:
            return {
                "success": False,
                "error": "Modulo oracledb non disponibile",
                "suggestion": "pip install oracledb"
            }
        
        try:
            from .credentials_manager import get_database_credentials
            credentials = get_database_credentials("pporafin")
        except ImportError:
            # Import assoluto per esecuzione standalone
            from credentials_manager import get_database_credentials
            credentials = get_database_credentials("pporafin")
        
        if not credentials:
            return {
                "success": False,
                "error": "Credenziali Oracle non configurate"
            }
        
        username, password = credentials
        last_error = None
        
        # Test TNS
        try:
            connection = oracledb.connect(
                user=username,
                password=password,
                dsn=self.tns_alias
            )
            connection.close()
            self.successful_host = "TNS"
            return {
                "success": True,
                "method": "TNS",
                "alias": self.tns_alias,
                "message": f"Connessione TNS riuscita con alias '{self.tns_alias}'"
            }
        except Exception as e:
            last_error = str(e)
            self.logger.warning(f"Connessione TNS fallita: {e}")
        
        # Test connessioni dirette
        for host in self.fallback_hosts:
            try:
                dsn = f"{host}:1521/pporafin"
                connection = oracledb.connect(
                    user=username,
                    password=password,
                    dsn=dsn
                )
                connection.close()
                self.successful_host = host
                return {
                    "success": True,
                    "method": "Direct",
                    "host": host,
                    "dsn": dsn,
                    "message": f"Connessione diretta riuscita con {host}"
                }
            except Exception as e:
                last_error = str(e)
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
        
        try:
            from .credentials_manager import get_database_credentials
            credentials = get_database_credentials("pporafin")
        except ImportError:
            # Import assoluto per esecuzione standalone
            from credentials_manager import get_database_credentials
            credentials = get_database_credentials("pporafin")
        
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
                    dsn = f"{host}:1521/pporafin"
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
