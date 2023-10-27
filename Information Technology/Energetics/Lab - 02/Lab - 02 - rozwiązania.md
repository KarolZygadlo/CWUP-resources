#### 1. Zadania dotyczące instrukcji warunkowej `IF`

**Zadanie 1.1**: Utwórz procedurę, która sprawdzi, czy liczba jest dodatnia, ujemna czy równa zero. Użyj instrukcji `If...ElseIf...Else` i wyświetl odpowiedni komunikat za pomocą `MsgBox`.

```vba
Sub SprawdzLiczbe()
    Dim liczba As Integer
    liczba = InputBox("Wprowadź liczbę:")

    If liczba > 0 Then
        MsgBox "Liczba jest dodatnia."
    ElseIf liczba < 0 Then
        MsgBox "Liczba jest ujemna."
    Else
        MsgBox "Liczba jest równa zero."
    End If
End Sub
```

**Zadanie 1.2**: Utwórz procedurę, która przyjmuje dwa argumenty typu `Integer` i za pomocą `MsgBox` informuje, który z nich jest większy lub czy są równe.

```vba
Sub PorownajLiczby(a As Integer, b As Integer)
    If a > b Then
        MsgBox "Pierwsza liczba (" & a & ") jest większa niż druga (" & b & ")."
    ElseIf a < b Then
        MsgBox "Druga liczba (" & b & ") jest większa niż pierwsza (" & a & ")."
    Else
        MsgBox "Liczby są równe."
    End If
End Sub
```

#### 2. Zadania dotyczące operacji na datach i czasie

**Zadanie 2.1**: Stwórz procedurę, która wyświetli informację o tym, ile dni pozostało do końca roku. Użyj funkcji `DateDiff` i aktualnej daty z funkcji `Date`.

```vba
Sub DniDoKoncaRoku()
    Dim dniPozostale As Integer
    dniPozostale = DateDiff("d", Date, DateSerial(Year(Date), 12, 31))
    MsgBox "Do końca roku pozostało " & dniPozostale & " dni."
End Sub
```

**Zadanie 2.2**: Utwórz procedurę, która dodaje pewną liczbę dni (np. 15) do aktualnej daty i wyświetla wynik w formie komunikatu.

```vba
Sub DodajDniDoDaty()
    Dim nowaData As Date
    nowaData = DateAdd("d", 15, Date)
    MsgBox "Data po dodaniu 15 dni: " & nowaData
End Sub
```

#### 3. Zadania dotyczące ikon i przycisków w oknach komunikatu

**Zadanie 3.1**: Zaprojektuj procedurę, która wyświetli okno komunikatu z pytaniem „Czy jesteś pewien?”. Okno powinno zawierać przyciski Tak/Nie oraz ikonę pytania.

```vba
Sub CzyJestesPewien()
    Dim odpowiedz As VbMsgBoxResult
    odpowiedz = MsgBox("Czy jesteś pewien?", vbQuestion + vbYesNo, "Pytanie")
    If odpowiedz = vbYes Then
        ' Użytkownik wybrał "Tak"
    Else
        ' Użytkownik wybrał "Nie"
    End If
End Sub
```

**Zadanie 3.2**: Utwórz procedurę, która przy użyciu `MsgBox` wyświetli okno komunikatu z ikoną błędu i przyciskami Anuluj, Spróbuj ponownie, Kontynuuj. W zależności od wyboru użytkownika, pokaż odpowiadający mu komunikat.

```vba
Sub KomunikatZBledem()
    Dim odpowiedz As VbMsgBoxResult
    odpowiedz = MsgBox("Wystąpił błąd!", vbCritical + vbAbortRetryIgnore, "Błąd")

    Select Case odpowiedz
        Case vbAbort
            MsgBox "Wybrano Anuluj."
        Case vbRetry
            MsgBox "Wybrano Spróbuj ponownie."
        Case vbIgnore
            MsgBox "Wybrano Kontynuuj."
    End Select
End Sub
```