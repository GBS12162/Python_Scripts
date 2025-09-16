"""
Gestione del menu principale per l'applicazione Oracle Component Lookup.
"""

from typing import List, Callable


class MenuManager:
    """Gestore del menu principale."""
    
    def __init__(self):
        self.menu_items: List[tuple[str, Callable[[], None]]] = []
    
    def add_menu_item(self, label: str, action: Callable[[], None]):
        """Aggiunge una voce di menu."""
        self.menu_items.append((label, action))
    
    def show_menu(self, title: str = "MENU PRINCIPALE"):
        """Mostra il menu e gestisce la selezione dell'utente con loop persistente come l'originale."""
        while True:
            # Pulisce lo schermo (simulato)
            print("\n" + "=" * 60)
            print("🔧 Look-up Components CDL - Versione Modulare")  
            print("=" * 60)
            
            # Menu con formato identico all'originale
            for i, (label, _) in enumerate(self.menu_items):
                print(f"{i + 1}. {label}")
            print(f"{len(self.menu_items) + 1}. ❌ Esci")
            
            try:
                choice = input(f"\n➤ Seleziona opzione (1-{len(self.menu_items) + 1}): ").strip()
                
                # Gestione uscita (ultima opzione)
                if choice == str(len(self.menu_items) + 1):
                    print("\n👋 Arrivederci!")
                    print("Grazie per aver usato Look-up Components CDL")
                    return
                
                # Gestione opzioni valide
                choice_int = int(choice)
                if 1 <= choice_int <= len(self.menu_items):
                    _, action = self.menu_items[choice_int - 1]
                    action()
                    # Dopo l'azione, mostra un messaggio e torna al menu
                    input("\n📝 Premi INVIO per continuare...")
                else:
                    print(f"❌ Opzione non valida. Scegli 1-{len(self.menu_items) + 1}")
                    input("\n📝 Premi INVIO per continuare...")
                    
            except ValueError:
                print("❌ Inserire un numero valido")
                input("\n📝 Premi INVIO per continuare...")
            except KeyboardInterrupt:
                print("\n\n👋 Arrivederci!")
                return
    
    def clear_menu(self):
        """Rimuove tutte le voci di menu."""
        self.menu_items.clear()
