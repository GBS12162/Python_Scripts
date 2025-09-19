"""
Transaction Reporting - CON-412 Rejecting Mensile
Sistema di elaborazione automatica per la validazione ISIN tramite ESMA
"""

import os
import sys
import logging
from datetime import datetime

# Add project root to path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from attivita.transaction_reporting.processor import CON412Processor
from attivita.transaction_reporting.config.con412_config import config
from utils.progress_utils import ProgressUtils


def setup_logging():
    """Configura il logging per il sistema"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_filename = f"logs/con412_processing_{timestamp}.log"
    
    # Assicura che la directory logs esista
    os.makedirs(os.path.dirname(log_filename), exist_ok=True)
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_filename),
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    return logging.getLogger(__name__)


def main():
    """Funzione principale per l'elaborazione CON-412"""
    logger = setup_logging()
    
    try:
        logger.info("=" * 60)
        logger.info("AVVIO SISTEMA CON-412 - TRANSACTION REPORTING")
        logger.info("=" * 60)
        logger.info(f"Data/ora avvio: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        logger.info(f"Mese di elaborazione: {config['current_month']}")
        logger.info("")
        
        # Inizializza il processore
        processor = CON412Processor()
        
        # Esegui l'elaborazione completa
        success = processor.run()
        
        if success:
            logger.info("=" * 60)
            logger.info("ELABORAZIONE CON-412 COMPLETATA CON SUCCESSO")
            logger.info("=" * 60)
            logger.info(f"Data/ora completamento: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            return 0
        else:
            logger.error("=" * 60)
            logger.error("ELABORAZIONE CON-412 FALLITA")
            logger.error("=" * 60)
            return 1
            
    except KeyboardInterrupt:
        logger.warning("\n\nElaborazione interrotta dall'utente")
        return 130
        
    except Exception as e:
        logger.error(f"Errore critico durante l'elaborazione: {str(e)}")
        logger.error("Consultare i log per maggiori dettagli")
        return 1


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)