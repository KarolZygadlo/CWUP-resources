## Ćwiczenie: Makro do Przetwarzania Informacji o Świętach

### Opis zadania

Napisz makro, które będzie przeszukiwało wszystkie komórki w zakresie nazwanym "Kraje", znajdując informacje o każdym święcie i wypisując je w oknie natychmiastowym.

### Propozycja rozwiązania 

```vba
Option Explicit

Sub PokazSwieta()

    ' używane do przeglądania świąt
    Dim KomorkaSwieta As Range
    Dim ZakresSwiat As Range
    
    ' 3 pola dla każdego święta
    Dim CzasTrwaniaSwieta As Integer
    Dim KurortSwieta As String
    Dim CenaSwieta As Currency
    
    ' utworzenie zmiennej obiektu odnoszącej się do zestawu świąt
    Set ZakresSwiat = Range("Countries")
    
    ' pętla przechodząca przez wszystkie te komórki
    For Each KomorkaSwieta In ZakresSwiat
    
        ' dla każdego pobierz cenę, czas trwania i kurort
        CzasTrwaniaSwieta = KomorkaSwieta.Offset(0, 2).Value
        CenaSwieta = KomorkaSwieta.Offset(0, 4).Value
        KurortSwieta = KomorkaSwieta.Offset(0, 1).Value
        
        Debug.Print CStr(CzasTrwaniaSwieta) & " dni w " & KurortSwieta & " za " & Format(CenaSwieta, "Ł#,##0")
        
    Next KomorkaSwieta
    
End Sub
```