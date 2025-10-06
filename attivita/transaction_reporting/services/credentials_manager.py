"""
Credentials Manager - Gestione sicura credenziali database
Sistema per acquisire e gestire credenziali utente in modo sicuro con salvataggio persistente
"""

import getpass
import logging
import json
import os
import sys
from typing import Tuple, Optional, Dict
from pathlib import Path


class CredentialsManager:
    """Gestore sicuro per credenziali database con salvataggio persistente"""
    
    def __init__(self):
        """Inizializza il gestore credenziali"""
        self.logger = self._setup_logging()
        self._cached_credentials = {}
        
        # Setup directory per credenziali persistenti
        self.config_dir = Path(__file__).parent.parent / "config"
        self.config_dir.mkdir(exist_ok=True)
        self.credentials_file = self.config_dir / "saved_credentials.json"
        
        # Carica credenziali salvate all'avvio
        self._load_saved_credentials()
    
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
    
    def _load_saved_credentials(self):
        """Carica credenziali salvate da file"""
        try:
            if self.credentials_file.exists():
                with open(self.credentials_file, 'r') as f:
                    saved_data = json.load(f)
                    self._cached_credentials.update(saved_data)
                self.logger.info("Credenziali salvate caricate")
        except Exception as e:
            self.logger.warning(f"Errore caricamento credenziali salvate: {e}")
    
    def _save_credentials(self, service_name: str, username: str, password: str):
        """Salva credenziali su file (solo username per sicurezza)"""
        try:
            # Salva solo username, non la password per sicurezza
            saved_data = {}
            if self.credentials_file.exists():
                with open(self.credentials_file, 'r') as f:
                    saved_data = json.load(f)
            
            saved_data[service_name] = {
                'username': username,
                'saved_at': self._get_timestamp()
            }
            
            with open(self.credentials_file, 'w') as f:
                json.dump(saved_data, f, indent=2)
                
            self.logger.info(f"Username salvato per {service_name}")
        except Exception as e:
            self.logger.warning(f"Errore salvataggio credenziali: {e}")
    
    def _get_timestamp(self):
        """Ottiene timestamp corrente"""
        from datetime import datetime
        return datetime.now().isoformat()
    
    def _get_saved_username(self, service_name: str) -> Optional[str]:
        """Ottiene username salvato per un servizio"""
        try:
            if self.credentials_file.exists():
                with open(self.credentials_file, 'r') as f:
                    saved_data = json.load(f)
                    if service_name in saved_data:
                        return saved_data[service_name].get('username')
        except Exception:
            pass
        return None
    
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
            cached = self._cached_credentials[service_name]
            if isinstance(cached, tuple) and len(cached) == 2:
                self.logger.debug(f"Utilizzo credenziali cached per {service_name}")
                return cached
        
        print(f"\n🔐 AUTENTICAZIONE DATABASE")
        print(f"Servizio: {service_name.upper()}")
        print("=" * 40)
        
        # SUGGERIMENTI per username comune
        print("💡 SUGGERIMENTI CREDENZIALI:")
        print("   - Username tipico: CONSULTA_IT_RUN_INVESTIMENTI")
        print("   - Se hai dubbi, verifica con l'amministratore DB")
        print("")
        
        # Controlla se c'è un username salvato
        saved_username = self._get_saved_username(service_name)
        
        try:
            # Richiede username (con default se salvato)
            username = self._get_username(saved_username)
            
            # Richiede password
            password = self._get_password()
            
            # Valida le credenziali
            self._validate_credentials(username, password)
            
            # Memorizza in cache per la sessione
            credentials = (username, password)
            self._cached_credentials[service_name] = credentials
            
            # Salva username per prossime volte (non la password)
            self._save_credentials(service_name, username, password)
            
            print("✅ Credenziali acquisite e memorizzate")
            print(f"🔍 DEBUG: Username = '{username}' (lunghezza: {len(username)})")
            print(f"🔍 DEBUG: Password lunghezza = {len(password)}")
            self.logger.info(f"Credenziali acquisite per {service_name} (utente: {username})")
            
            return credentials
            
        except KeyboardInterrupt:
            print("\n❌ Acquisizione credenziali annullata")
            raise Exception("Acquisizione credenziali annullata dall'utente")
        except Exception as e:
            self.logger.error(f"Errore acquisizione credenziali: {str(e)}")
            raise
    
    def _get_username(self, default_username: Optional[str] = None) -> str:
        """
        Acquisisce username con validazione
        
        Args:
            default_username: Username di default se disponibile
            
        Returns:
            Username validato
        """
        while True:
            try:
                # Verifica se siamo in un eseguibile PyInstaller
                is_exe = getattr(sys, 'frozen', False)
                
                if default_username:
                    prompt = f"👤 Username [{default_username}]: "
                    if is_exe:
                        # Fallback per eseguibili PyInstaller
                        import sys
                        sys.stdout.write(prompt)
                        sys.stdout.flush()
                        username = sys.stdin.readline().strip()
                    else:
                        username = input(prompt).strip()
                    
                    if not username:
                        username = default_username
                        print(f"✅ Utilizzo username salvato: {username}")
                else:
                    if is_exe:
                        # Fallback per eseguibili PyInstaller
                        import sys
                        sys.stdout.write("👤 Username: ")
                        sys.stdout.flush()
                        username = sys.stdin.readline().strip()
                    else:
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
                # Verifica se siamo in un eseguibile PyInstaller
                is_exe = getattr(sys, 'frozen', False)
                
                if is_exe:
                    # Fallback per eseguibili PyInstaller
                    import sys
                    sys.stdout.write("🔑 Password: ")
                    sys.stdout.flush()
                    password = sys.stdin.readline().strip()
                    print("✅ Password inserita")
                else:
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