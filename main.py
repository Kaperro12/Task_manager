# main.py
# Importujemy funkcje z pliku functions.py
from functions import dodaj_zadanie, rozpocznij_zadanie, zakoncz_zadanie, wyswietl_zadania, menu


# Definicja głównej funkcji aplikacji
def main():
    menu()

# Uruchomienie funkcji main, jeśli skrypt jest uruchamiany bezpośrednio
if __name__ == "__main__":
    main()