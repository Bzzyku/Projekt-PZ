import pandas as pd
import yfinance as yf

# Ścieżka do pliku Excel, który będzie przechowywał dane
excel_file_path = 'spolki_dywidendowe.xlsx'

# Funkcja do pobierania bieżącej ceny akcji oraz historii cen z ostatnich 7 dni
def pobierz_dane_spolki(ticker):
    spolka = yf.Ticker(ticker)
    # Pobieranie bieżącej ceny
    cena_aktualna = spolka.history(period="1d")['Close'].iloc[-1]
    # Pobieranie historii cen z ostatniego tygodnia
    historia_cen = spolka.history(period="5d", interval="1d")['Close'].tolist()

    # Upewnienie się, że historia_cen ma dokładnie 7 elementów
    if len(historia_cen) < 7:
        # Dodajemy None, jeśli brakuje dni do pełnych 7
        historia_cen = [None] * (7 - len(historia_cen)) + historia_cen
    elif len(historia_cen) > 7:
        # Wybieramy tylko ostatnie 7 dni, jeśli jest ich więcej
        historia_cen = historia_cen[-7:]
    
    return cena_aktualna, historia_cen

# Funkcja do aktualizacji pliku Excel z cenami i historią cen
def aktualizuj_excel(spolki):
    # Tworzenie nowego DataFrame, jeśli plik nie istnieje
    kolumny = ['Nazwa Spółki', 'Ticker', 'Cena Aktualna'] + [f'Cena D-{i}' for i in range(1, 8)]
    dane = pd.DataFrame(columns=kolumny)

    for nazwa, ticker in spolki.items():
        cena_aktualna, historia_cen = pobierz_dane_spolki(ticker)
        # Przygotowanie wpisu do DataFrame
        wpis = [nazwa, ticker, cena_aktualna] + historia_cen
        dane.loc[len(dane)] = wpis

    # Zapis do pliku Excel
    dane.to_excel(excel_file_path, index=False)

# Lista spółek dywidendowych (nazwa spółki: ticker)
spolki = {
    'PKN Orlen': 'PKN.WA',
    'KGHM': 'KGH.WA',
    'PZU': 'PZU.WA',
    # Dodaj więcej spółek według potrzeb
}

# Aktualizacja arkusza Excel z bieżącymi cenami i historią cen
aktualizuj_excel(spolki)
print("Dane zostały zapisane w pliku Excel.")
