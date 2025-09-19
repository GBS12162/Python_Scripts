"""
Credentials Manager - Gestione sicura credenziali database
Sistema per acquisire e gestire credenziali utente in modo sicuro
"""

import getpass
import logging
from typing import Tuple, Optional
from pathlib import Path


class CredentialsManager:
    """Gestore sicuro per credenziali database"""
    
    def __init__(self):
        """Inizializza il gestore credenziali"""
        self.logger = self._setup_logging()
        self._cached_credentials = {}
    
    def _setup_logging(self):
        """Configura il sistema di logging"""
        logger = logging.getLogger(__name__)
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
            logger.setLevel(logging.INFO)
        return logger
    
    def get_credentials(self, service_name: str, force_new: bool = False) -> Tuple[str, str]:
        """
        Acquisisce credenziali per un servizio specifico
        
        Args:
            service_name: Nome del servizio (es. 'pporafin')
            force_new: Se True, richiede nuove credenziali anche se già presenti in cache
            
        Returns:
            Tupla (username, password)
        """
        # Controlla cache se non forzato
        if not force_new and service_name in self._cached_credentials:
            self.logger.debug(f"Utilizzo credenziali cached per {service_name}")
            return self._cached_credentials[service_name]
        
        print(f"\n🔐 AUTENTICAZIONE DATABASE")
        print(f"Servizio: {service_name.upper()}")
        print("=" * 40)
        
        try:
            # Richiede username
            username = self._get_username()
            
            # Richiede password
            password = self._get_password()
            
            # Valida le credenziali
            self._validate_credentials(username, password)
            
            # Memorizza in cache per la sessione
            credentials = (username, password)
            self._cached_credentials[service_name] = credentials
            
            print("✅ Credenziali acquisite e memorizzate")
            self.logger.info(f"Credenziali acquisite per {service_name} (utente: {username})")
            
            return credentials
            
        except KeyboardInterrupt:
            print("\n❌ Acquisizione credenziali annullata")
            raise Exception("Acquisizione credenziali annullata dall'utente")
        except Exception as e:
            self.logger.error(f"Errore acquisizione credenziali: {str(e)}")
            raise
    
    def _get_username(self) -> str:
        """
        Acquisisce username con validazione
        
        Returns:
            Username validato
        """
        while True:
            try:
                username = input("👤 Username: ").strip()
                
                if not username:
                    print("⚠️ Username non può essere vuoto")
                    continue
                
                if len(username) < 2:
                    print("⚠️ Username troppo corto (minimo 2 caratteri)")
                    continue
                
                # Validazione caratteri username (alfanumerici e alcuni simboli)
                if not all(c.isalnum() or c in ['_', '.', '-'] for c in username):
                    print("⚠️ Username contiene caratteri non validi")
                    continue
                
                return username
                
            except KeyboardInterrupt:
                raise
            except Exception as e:
                print(f"⚠️ Errore inserimento username: {str(e)}")
                continue
    
    def _get_password(self) -> str:
        """
        Acquisisce password in modo sicuro
        
        Returns:
            Password validata
        """
        while True:
            try:
                password = getpass.getpass("🔑 Password: ")
                
                if not password:
                    print("⚠️ Password non può essere vuota")
                    continue
                
                if len(password) < 3:
                    print("⚠️ Password troppo corta (minimo 3 caratteri)")
                    continue
                
                return password
                
            except KeyboardInterrupt:
                raise
            except Exception as e:
                print(f"⚠️ Errore inserimento password: {str(e)}")
                continue
    
    def _validate_credentials(self, username: str, password: str):
        """
        Valida le credenziali acquisite
        
        Args:
            username: Username da validare
            password: Password da validare
        """
        if not username or not password:
            raise ValueError("Username e password sono obbligatori")
        
        if len(username.strip()) == 0 or len(password.strip()) == 0:
            raise ValueError("Username e password non possono essere vuoti")
    
    def clear_cache(self, service_name: Optional[str] = None):
        """
        Pulisce la cache delle credenziali
        
        Args:
            service_name: Nome servizio specifico, o None per pulire tutto
        """
        if service_name:
            if service_name in self._cached_credentials:
                del self._cached_credentials[service_name]
                self.logger.info(f"Cache credenziali pulita per {service_name}")
        else:
            self._cached_credentials.clear()
            self.logger.info("Cache credenziali completamente pulita")
    
    def has_cached_credentials(self, service_name: str) -> bool:
        """
        Verifica se ci sono credenziali in cache per un servizio
        
        Args:
            service_name: Nome del servizio
            
        Returns:
            True se ci sono credenziali in cache
        """
        return service_name in self._cached_credentials
    
    def get_cached_username(self, service_name: str) -> Optional[str]:
        """
        Ottiene l'username cached per un servizio (senza la password)
        
        Args:
            service_name: Nome del servizio
            
        Returns:
            Username se presente in cache, altrimenti None
        """
        if service_name in self._cached_credentials:
            return self._cached_credentials[service_name][0]
        return None


# Istanza globale del gestore credenziali
credentials_manager = CredentialsManager()


def get_database_credentials(service_name: str = "pporafin") -> Tuple[str, str]:
    """
    Funzione di convenienza per ottenere credenziali database
    
    Args:
        service_name: Nome del servizio database
        
    Returns:
        Tupla (username, password)
    """
    return credentials_manager.get_credentials(service_name)


def clear_database_credentials(service_name: str = "pporafin"):
    """
    Funzione di convenienza per pulire credenziali database
    
    Args:
        service_name: Nome del servizio database
    """
    credentials_manager.clear_cache(service_name)