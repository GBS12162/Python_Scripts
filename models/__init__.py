"""
Modelli per l'applicazione Oracle Component Lookup.
Contiene le classi dei dati e le strutture per i componenti Oracle.
"""

from .component import Component, LookupResult
from .config import Config

__all__ = ['Component', 'LookupResult', 'Config']
