"""
Servizio per i controlli di validazione ISIN e altri controlli di qualità.
Implementa il controllo 1: ISIN non censito tramite API esterna.
"""

import logging
import requests
from typing import Dict, List, Optional, Set, Any, Tuple
from datetime import datetime
import time

from models.transaction_reporting import ISINGroup, QualityControlResult
from config.con412_config import ControlliConfig


class ISINValidationService:
    """Servizio per la validazione ISIN tramite API esterna."""
    
    def __init__(self):
        """
        Inizializza il servizio di validazione ESMA.
        """
        self.logger = logging.getLogger(__name__)
        # URL fisso dell'API ESMA
        self.api_url = "https://registers.esma.europa.eu/publication/searchRegister/doMainSearch"
        self.session = requests.Session()
        
        # Cache per evitare richieste duplicate
        self._isin_cache: Dict[str, bool] = {}  # ISIN -> è_censito
        self._esma_data_cache: Dict[str, Dict] = {}  # ISIN -> dati completi ESMA
        self._cache_timestamp = datetime.now()
        self._cache_ttl_hours = 24  # Cache valida per 24 ore
        
        # Configurazione richieste per API ESMA
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Accept': 'application/json, text/plain, */*',
            'Content-Type': 'application/json',
            'Accept-Language': 'en-US,en;q=0.9',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
            'Origin': 'https://registers.esma.europa.eu',
            'Referer': 'https://registers.esma.europa.eu/publication/'
        })
        
        # Rate limiting - ESMA ha limiti più restrittivi
        self._last_request_time = 0
        self._min_request_interval = 0.5  # 500ms tra richieste per essere rispettosi
    
    def validate_isin_groups(self, isin_groups: List[ISINGroup]) -> List[QualityControlResult]:
        """
        Esegue il controllo di validazione su tutti i gruppi ISIN.
        
        Args:
            isin_groups: Lista dei gruppi ISIN da validare
            
        Returns:
            Lista dei risultati di controllo
        """
        try:
            self.logger.info(f"Inizio validazione {len(isin_groups)} gruppi ISIN con API ESMA")
            
            results = []
            unique_isins = self._extract_unique_isins(isin_groups)
            
            # Valida ISIN unici (per ridurre chiamate API)
            self.logger.info(f"Validazione {len(unique_isins)} ISIN unici tramite ESMA")
            isin_validation_results = self._validate_unique_isins(unique_isins)
            
            # Applica risultati ai gruppi
            for group in isin_groups:
                result = self._create_quality_control_result(group, isin_validation_results)
                results.append(result)
            
            self.logger.info(f"Validazione completata: {len(results)} risultati generati")
            return results
            
        except Exception as e:
            self.logger.error(f"Errore nella validazione gruppi ISIN: {e}")
            return self._create_error_results(isin_groups, str(e))
    
    def check_single_isin(self, isin: str) -> bool:
        """
        Controlla se un singolo ISIN è censito nell'API.
        
        Args:
            isin: Codice ISIN da controllare
            
        Returns:
            True se l'ISIN è censito, False altrimenti
        """
        try:
            # Controlla cache
            if self._is_cache_valid() and isin in self._isin_cache:
                return self._isin_cache[isin]
            
            # Rate limiting
            self._apply_rate_limiting()
            
            # Effettua richiesta API
            response = self._make_api_request(isin)
            is_valid, esma_data = self._parse_api_response_with_data(response, isin)
            
            # Aggiorna cache
            self._isin_cache[isin] = is_valid
            self._esma_data_cache[isin] = esma_data
            
            return is_valid
            
        except Exception as e:
            # Gestione degli errori in modo conservativo
            if "JSON" in str(e) or "html" in str(e).lower():
                return False  # NON censito se API restituisce HTML
            if "parsing" in str(e).lower():
                return False  # NON censito se parsing fallisce
            # In caso di altri errori, assumiamo che l'ISIN sia valido (approccio conservativo)
            return True
    
    def apply_validation_results_to_groups(
        self, 
        isin_groups: List[ISINGroup], 
        validation_results: Dict[str, bool]
    ):
        """
        Applica i risultati della validazione ai gruppi ISIN.
        
        Args:
            isin_groups: Lista dei gruppi ISIN
            validation_results: Dizionario ISIN -> è_censito
        """
        try:
            for group in isin_groups:
                is_censito = validation_results.get(group.isin, True)  # Default: assume valido
                
                self.logger.info(f"ISIN {group.isin}: validation_result={is_censito}")
                
                # Se ISIN NON è censito, metti "X" nel controllo 1
                if not is_censito:
                    group.controllo_1 = "X"
                    self.logger.info(f"ISIN {group.isin}: NON censito - marcato con X")
                else:
                    # Se è censito, lascia vuoto (o mantieni valore esistente se diverso da X)
                    if group.controllo_1 == "X":
                        group.controllo_1 = ""
                    self.logger.info(f"ISIN {group.isin}: censito correttamente")
            
            self.logger.info("Risultati validazione applicati ai gruppi ISIN")
            
        except Exception as e:
            self.logger.error(f"Errore nell'applicazione risultati validazione: {e}")
    
    def _extract_unique_isins(self, isin_groups: List[ISINGroup]) -> Set[str]:
        """Estrae gli ISIN unici da validare."""
        return {group.isin for group in isin_groups if group.isin and group.isin.strip()}
    
    def _validate_unique_isins(self, isins: Set[str]) -> Dict[str, bool]:
        """Valida un set di ISIN unici."""
        results = {}
        total = len(isins)
        
        for i, isin in enumerate(isins, 1):
            try:
                is_censito = self.check_single_isin(isin)
                results[isin] = is_censito
                
                if i % 10 == 0:  # Log progresso ogni 10 ISIN
                    self.logger.info(f"Validazione progresso: {i}/{total} ISIN processati")
                    
            except Exception as e:
                self.logger.error(f"Errore validazione ISIN {isin}: {e}")
                results[isin] = True  # Assume valido in caso di errore
        
        return results
    
    def _make_api_request(self, isin: str) -> requests.Response:
        """Effettua la richiesta API ESMA per un ISIN."""
        try:
            # Payload corretto per API ESMA (formato completo richiesto)
            payload = {
                "core": "esma_registers_firds",
                "pagingSize": "10",
                "start": 0,
                "keyword": "",
                "sortField": "isin asc",
                "criteria": [
                    {
                        "name": "isin",
                        "value": isin,
                        "type": "text",
                        "isParent": True
                    },
                    {
                        "name": "firdsPublicationDateCustomSearchInputField",
                        "value": "(latest_received_flag:1)",
                        "type": "customSearchInputFieldQuery",
                        "isParent": True
                    }
                ],
                "wt": "json"
            }
            
            response = self.session.post(
                self.api_url, 
                json=payload,
                timeout=30
            )
            
            response.raise_for_status()
            
            return response
            
        except requests.exceptions.RequestException as e:
            self.logger.error(f"Errore richiesta ESMA per ISIN {isin}: {e}")
            raise
    
    def _parse_api_response(self, response: requests.Response, isin: str) -> bool:
        """
        Parsa la risposta dell'API ESMA per determinare se l'ISIN è censito.
        
        Args:
            response: Risposta HTTP dell'API ESMA
            isin: ISIN richiesto
            
        Returns:
            True se l'ISIN è censito in ESMA, False altrimenti
        """
        try:
            if response.status_code == 200:
                data = response.json()
                
                # Struttura risposta ESMA
                if isinstance(data, dict):
                    # Controlla se ci sono risultati nella risposta ESMA
                    if "response" in data:
                        response_data = data["response"]
                        if "docs" in response_data:
                            docs = response_data["docs"]
                            # Se ci sono documenti, l'ISIN è censito
                            is_found = len(docs) > 0
                            return is_found
                        
                        elif "numFound" in response_data:
                            # Alcuni endpoint ESMA usano numFound
                            num_found = response_data["numFound"]
                            is_found = num_found > 0
                            
                            self.logger.debug(f"ISIN {isin} - numFound: {num_found}")
                            return is_found
                    
                    # Fallback: se la risposta contiene dati, assume censito
                    self.logger.debug(f"ISIN {isin} - risposta ESMA struttura non standard, assume censito")
                    return bool(data)
                
                # Se non è dict, assume non censito
                return False
                
            elif response.status_code == 404:
                # ISIN non trovato
                self.logger.debug(f"ISIN {isin} - 404 da ESMA")
                return False
            else:
                # Altri codici di errore - logga warning ma assume censito per sicurezza
                self.logger.warning(f"Codice risposta ESMA inaspettato {response.status_code} per ISIN {isin}")
                return True
                
        except Exception as e:
            self.logger.error(f"Errore nel parsing risposta ESMA per ISIN {isin}: {e}")
            # In caso di errore di parsing, assume censito per sicurezza
            return True
    
    def _create_quality_control_result(
        self, 
        group: ISINGroup, 
        validation_results: Dict[str, bool]
    ) -> QualityControlResult:
        """Crea il risultato del controllo di qualità per un gruppo."""
        try:
            result = QualityControlResult(
                isin=group.isin,
                total_orders=len(group.orders)
            )
            
            is_censito = validation_results.get(group.isin, True)
            
            # Controllo 1: ISIN non censito
            if not is_censito:
                result.controlli_failed += 1
                result.controlli_details["ISIN_NON_CENSITO"] = "X"
                result.business_rule_violations.append("ISIN non presente nell'anagrafica strumenti")
            else:
                result.controlli_passed += 1
                result.controlli_details["ISIN_NON_CENSITO"] = ""
            
            # Controllo 2: Trading Venue vs MERCATO (se disponibile)
            if isinstance(group, dict) and group.get('mercato'):
                venue_match = self.check_trading_venue(group['isin'], group['mercato'])
                if not venue_match:
                    result.controlli_failed += 1
                    result.controlli_details["MIC_CODE_NON_PRESENTE"] = "X"
                    result.business_rule_violations.append("MIC code non presente per quel mercato - Trading venue ESMA non trovato")
                else:
                    result.controlli_passed += 1
                    result.controlli_details["MIC_CODE_NON_PRESENTE"] = ""
            
            # Aggiungi raccomandazioni se necessario
            if result.controlli_failed > 0:
                result.recommendations.append("Verificare codice ISIN nell'anagrafica strumenti")
            
            return result
            
        except Exception as e:
            self.logger.error(f"Errore creazione risultato controllo per ISIN {group.isin}: {e}")
            return QualityControlResult(isin=group.isin, total_orders=len(group.orders))
    
    def _create_empty_results(self, isin_groups: List[ISINGroup]) -> List[QualityControlResult]:
        """Crea risultati vuoti quando l'API non è configurata."""
        return [
            QualityControlResult(isin=group.isin, total_orders=len(group.orders))
            for group in isin_groups
        ]
    
    def _create_error_results(self, isin_groups: List[ISINGroup], error_msg: str) -> List[QualityControlResult]:
        """Crea risultati di errore."""
        results = []
        for group in isin_groups:
            result = QualityControlResult(isin=group.isin, total_orders=len(group.orders))
            result.validation_errors.append(f"Errore validazione: {error_msg}")
            results.append(result)
        return results
    
    def _is_cache_valid(self) -> bool:
        """Controlla se la cache è ancora valida."""
        if not self._cache_timestamp:
            return False
        
        elapsed_hours = (datetime.now() - self._cache_timestamp).total_seconds() / 3600
        return elapsed_hours < self._cache_ttl_hours
    
    def _apply_rate_limiting(self):
        """Applica rate limiting tra le richieste."""
        elapsed = time.time() - self._last_request_time
        if elapsed < self._min_request_interval:
            sleep_time = self._min_request_interval - elapsed
            time.sleep(sleep_time)
        
        self._last_request_time = time.time()
    
    def clear_cache(self):
        """Pulisce la cache ISIN."""
        self._isin_cache.clear()
        self._esma_data_cache.clear()
        self._cache_timestamp = datetime.now()
        self.logger.info("Cache ISIN pulita")
    
    def get_esma_data(self, isin: str) -> Optional[Dict]:
        """Ottiene i dati ESMA completi per un ISIN."""
        return self._esma_data_cache.get(isin)
    
    def check_trading_venue(self, isin: str, mercato_value: str) -> bool:
        """
        Controllo 2: Verifica se almeno una riga ESMA ha il Trading Venue corrispondente al MERCATO.
        
        Args:
            isin: Codice ISIN
            mercato_value: Valore della colonna MERCATO (formato: "CODICE(DESCRIZIONE)")
            
        Returns:
            True se almeno una riga ESMA ha il trading venue corrispondente, False altrimenti
        """
        try:
            if not mercato_value:
                self.logger.debug(f"MERCATO value vuoto per ISIN {isin}")
                return False
            
            # Eccezione speciale: se MERCATO contiene "XOFF", sempre valido
            if "XOFF" in mercato_value.upper():
                self.logger.debug(f"MERCATO contiene XOFF per ISIN {isin}: {mercato_value} - sempre valido")
                return True
            
            # Estrae il codice trading venue dal MERCATO (es: "MTAA(MTA)" -> "MTAA")
            mercato_code = mercato_value.split('(')[0].strip() if '(' in mercato_value else mercato_value.strip()
            
            if not mercato_code:
                self.logger.debug(f"Impossibile estrarre codice MERCATO da '{mercato_value}' per ISIN {isin}")
                return False
            
            # Verifica se abbiamo già i dati ESMA in cache
            if isin not in self._esma_data_cache:
                # Chiama l'API per ottenere i dati completi
                is_valid = self.check_single_isin(isin)
                if not is_valid or isin not in self._esma_data_cache:
                    self.logger.debug(f"Nessun dato ESMA disponibile per ISIN {isin}")
                    return False
            
            esma_data = self._esma_data_cache[isin]
            
            # Ottieni tutti i trading venues da ESMA
            trading_venues = esma_data.get('trading_venues', [])
            
            if not trading_venues:
                self.logger.debug(f"Nessun trading venue trovato nei dati ESMA per ISIN {isin}")
                return False
            
            # Verifica se almeno uno dei trading venues corrisponde al MERCATO
            mercato_code_upper = mercato_code.upper()
            
            for venue in trading_venues:
                venue_upper = str(venue).upper()
                
                # Controllo diretto
                if venue_upper == mercato_code_upper:
                    self.logger.debug(f"Trading venue match diretto per ISIN {isin}: {mercato_code} = {venue}")
                    return True
                
                # Controllo se il venue contiene il codice MERCATO
                if mercato_code_upper in venue_upper:
                    self.logger.debug(f"Trading venue match parziale per ISIN {isin}: {mercato_code} in {venue}")
                    return True
            
            # Nessun match trovato
            self.logger.debug(f"Trading venue mismatch per ISIN {isin}: {mercato_code} non trovato in {trading_venues}")
            return False
            
        except Exception as e:
            self.logger.error(f"Errore nel controllo trading venue per ISIN {isin}: {e}")
            return False
    
    def _parse_api_response_with_data(self, response: requests.Response, isin: str) -> Tuple[bool, Dict]:
        """
        Parsa la risposta dell'API ESMA e restituisce validità + dati completi.
        """
        try:
            if response.status_code == 200:
                # Controlla se la risposta è HTML invece di JSON
                content_type = response.headers.get('content-type', '').lower()
                
                if 'html' in content_type:
                    # Controlla se è un errore generale dell'applicazione ESMA
                    if 'General application error' in response.text:
                        self.logger.error(f"API ESMA ERROR - General application error per ISIN {isin}")
                        self.logger.error(f"URL richiesta: {response.url}")
                        self.logger.error(f"Status code: {response.status_code}")
                        self.logger.error(f"Response HTML (primi 500 char): {response.text[:500]}...")
                        self.logger.warning(f"API ESMA non disponibile per ISIN {isin} - assumo CENSITO (conservativo)")
                        return True, {}  # Censito per sicurezza
                    else:
                        self.logger.warning(f"API ESMA ha restituito HTML per ISIN {isin} - probabilmente NON censito")
                        return False, {}  # NON censito
                
                data = response.json()
                
                if isinstance(data, dict) and "response" in data:
                    response_data = data["response"]
                    num_found = response_data.get("numFound", 0)
                    docs = response_data.get("docs", [])
                    
                    is_found = num_found > 0 and len(docs) > 0
                    
                    # Logging dettagliato per il debug
                    self.logger.info(f"ISIN {isin}: API ESMA response - numFound={num_found}, docs_length={len(docs)}")
                    if num_found == 0:
                        self.logger.info(f"ISIN {isin}: NON CENSITO su ESMA (numFound=0) - riceverà X")
                    else:
                        self.logger.info(f"ISIN {isin}: CENSITO su ESMA (numFound={num_found}) - nessuna X")
                    
                    self.logger.debug(f"ISIN {isin}: numFound={num_found}, docs={len(docs)}, is_found={is_found}")
                    
                    # Estrai dati da TUTTI i documenti
                    esma_data = {
                        'all_docs': docs,  # Conserva tutti i documenti
                        'doc_count': len(docs),
                        'num_found': num_found,
                        'trading_venues': []  # Lista di tutti i trading venues
                    }
                    
                    # Estrai tutti i trading venues da tutti i documenti
                    for doc in docs:
                        # Possibili campi per trading venue
                        venue_fields = [
                            'full_name_of_the_trading_venue',
                            'trading_venue',
                            'trading_venue_of_the_product',
                            'venue_of_the_product',
                            'mic_code_of_the_most_relevant_market'
                        ]
                        
                        for field in venue_fields:
                            if field in doc and doc[field]:
                                venue_value = str(doc[field]).strip()
                                if venue_value and venue_value not in esma_data['trading_venues']:
                                    esma_data['trading_venues'].append(venue_value)
                    
                    # Aggiungi anche i dati del primo documento per compatibilità
                    if docs:
                        doc = docs[0]
                        esma_data.update({
                            'trading_venue': doc.get('full_name_of_the_trading_venue'),
                            'instrument_name': doc.get('instrument_name'),
                            'cfii': doc.get('cfii'),
                            'notional_currency': doc.get('notional_currency')
                        })
                    
                    return is_found, esma_data
                        
            return False, {}
            
        except Exception as e:
            self.logger.error(f"Errore parsing risposta ESMA per {isin}: {e}")
            # Se errore di parsing JSON, probabilmente ISIN non censito (API restituisce HTML)
            if "JSON" in str(e) or "Expecting value" in str(e):
                self.logger.warning(f"Errore JSON per ISIN {isin} - probabilmente NON censito")
                return False, {}
            # Per altri errori, approccio conservativo
            return True, {}
    
    def get_cache_stats(self) -> Dict[str, Any]:
        """Ottiene statistiche sulla cache."""
        return {
            "total_cached_isins": len(self._isin_cache),
            "cache_age_hours": (datetime.now() - self._cache_timestamp).total_seconds() / 3600,
            "cache_valid": self._is_cache_valid(),
            "cached_isins": list(self._isin_cache.keys())
        }