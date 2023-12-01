## Ćwiczenie: Analiza Filmów z Dużą Liczbą Recenzji

### Opis zadania

Otwórz skoroszyt w powyższym folderze. Zawiera on listę dziesięciu najlepszych filmów wszech czasów według RottenTomatoes: zad01.xlsx

Twoim zadaniem jest napisanie makra, które będzie przeszukiwało wybrane powyżej tytuły filmów, kolorując te, które uzyskały więcej niż 100 recenzji.

### Propozycja rozwiązania 

```vba
Option Explicit

Sub KolorujFilmyZDuzaLiczbaRecenzji()

    Dim KomorkaFilmu As Range
    
    ' ustawienie zmiennej do odniesienia się do kolumny z filmami
    Dim ZakresFilmow As Range
    
    ' pobranie referencji do bloku filmów
    Set ZakresFilmow = Range( _
        Range("B2"), _
        Range("B2").End(xlDown))
        
    ' pętla przechodząca przez każdy film
    For Each KomorkaFilmu In ZakresFilmow
    
        ' jeśli film otrzymał więcej niż 100 recenzji, pokoloruj go
        If KomorkaFilmu.Offset(0, 2).Value > 100 Then
            
            KomorkaFilmu.Interior.Color = RGB(200, 200, 255)
            Debug.Print KomorkaFilmu.Value, KomorkaFilmu.Offset(0, 2).Value
        
        End If
    
    Next KomorkaFilmu
    
End Sub
```