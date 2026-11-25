# Aplikacja FOREMAN

## Lokalizacja na Dysku Google

Skrypt Apps Script musi być powiązany z arkuszem źródłowym aplikacji. Arkusz musi znajdować się bezpośrednio w folderze `Tabela`:

```text
Dowolna nazwa projektu/
  Tabela/
    arkusz źródłowy main lub dev
  TIME SHEET/
  protokoły/
  Raporty/
  Rachunki/
  ZLEC/
  PZO/
```

- `TIME SHEET`: eksport XLSX zakładki timesheet.
- `protokoły`: karty pracy PDF.
- `Raporty`: raporty zbiorcze PDF i XLSX.
- `Rachunki`: załączniki kosztów; ich linki są zapisywane w zakładce `Koszty` bieżącego arkusza.
- `ZLEC`: zlecenia PDF i powiązane arkusze Google z komponentu Zlecenia i protokoły.
- `PZO`: protokoły zdawczo-odbiorcze PDF z tego komponentu.

Foldery docelowe muszą być bezpośrednimi podfolderami folderu nadrzędnego `Tabela`. Wielkość liter i zewnętrzne spacje nazw są ignorowane, ale polskie znaki i wewnętrzne spacje są istotne. Foldery nie są tworzone automatycznie; brak lub powielona nazwa wymaganego folderu powoduje błąd. Należy przenieść rzeczywisty arkusz, a nie skrót. Nazwa folderu projektu nie ma znaczenia.

Main i dev mogą korzystać ze wspólnej struktury lub osobnych struktur. Istniejące pliki nie są automatycznie przenoszone; nowe zapisy korzystają z aktualnego położenia arkusza. Zmiana folderów może również spowodować ponowne utworzenie dokumentów zleceń i PZO, których istniejące pliki nie znajdują się w nowych folderach.

PDF-y z komponentu Raporty są zapisywane na Dysku przed rozpoczęciem pobierania. Błąd lub brak potwierdzenia zapisu zatrzymuje pobieranie.

## Timesheet i raport tygodniowy

Karty pracy korzystają z zakładki `Spis wykonanych prac` bieżącego arkusza. Foreman zapisuje identyfikator zadania w kolumnie G, a identyfikator pokoju w kolumnie I. Generowanie kart synchronizuje zakładkę `REV OCEAN - timesheet`; wywołanie tej samej funkcji przez komponent Zlecenia i protokoły również uruchamia synchronizację. Synchronizacja czyści wybrany zakres dat dla wszystkich istniejących techników, także nieobecnych w pobranych danych. Błąd synchronizacji nie blokuje generowania kart. Eksport XLSX kopiuje bieżący stan timesheetu i sam nie uruchamia synchronizacji.

Raport „Czas w tygodniach” obejmuje całą historię wykonanych prac i sumuje numery tygodni ISO bez rozróżniania lat. Lista zadań powstaje z wykonanych wpisów; prace dodatkowe są prezentowane osobno. Opcja płatnych przerw obejmuje nazwy zawierające `przerwa`, `break` lub `brake` w zadaniu albo pokładzie.

## Wdrożenie i testy

1. Przygotuj foldery i umieść rzeczywisty arkusz w `Tabela`.
2. Zaktualizuj zmienione pliki w powiązanym projekcie Apps Script każdej używanej wersji i opublikuj nowe wersje wdrożeń. Sama zmiana repozytorium nie aktualizuje wdrożenia.
3. Zapewnij użytkownikowi wykonującemu skrypt dostęp do arkusza i folderów oraz prawo zapisu.
4. Sprawdź zapis i pobieranie obu PDF-ów, raportu XLSX, eksportu timesheetu, rachunku oraz dokumentów `ZLEC` i `PZO`; sprawdź też przycisk folderu rachunków.

Testy lokalne: `node --test tests/*.test.js`. Używają symulowanych usług Apps Script i nie zastępują testów uprawnień, rzeczywistego Dysku ani renderowania PDF po wdrożeniu.