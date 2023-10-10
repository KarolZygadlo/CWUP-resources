## 1. Instrukcja Warunkowa IF

#### Opis
Instrukcja `IF` jest fundamentalnym budulcem logiki warunkowej w językach programowania, takich jak VBA, służącym do sterowania przepływem programu. Działa na bazie oceny wyrażenia logicznego (warunku) - jeśli jest ono prawdziwe, pewien blok kodu zostaje wykonany; w przeciwnym razie, wykonywany jest inny blok kodu lub też nic się nie dzieje.

```vba
If [warunek] Then
    [instrukcje do wykonania, jeśli warunek jest prawdziwy]
[ElseIf [następny warunek] Then
    [instrukcje do wykonania, jeśli następny warunek jest prawdziwy]]
[Else
    [instrukcje do wykonania, jeśli żaden z warunków nie jest prawdziwy]]
End If
```

Elementy w nawiasach kwadratowych są opcjonalne, co oznacza, że instrukcja `IF` może przybrać różne formy, od bardzo prostej do bardziej złożonej i rozbudowanej.

#### Głębsza Analiza

- **[warunek]**: Jest to wyrażenie, które jest oceniane jako `TRUE` lub `FALSE`. Jeśli wyrażenie jest prawdziwe, instrukcje zawarte między `Then` a `End If` (lub ewentualnie `Else`, jeśli jest używane) są wykonywane. Jeśli wyrażenie jest fałszywe, te instrukcje są pomijane.
- **[instrukcje do wykonania]**: To są linie kodu, które zostaną wykonane w zależności od tego, czy warunek jest prawdziwy czy fałszywy.
- **ElseIf**: Opcjonalny. Służy do sprawdzenia dodatkowego warunku, jeśli poprzedni był fałszywy.
- **Else**: Opcjonalny. Odpowiada za wykonanie instrukcji, gdy żaden z poprzednich warunków nie jest prawdziwy.
#### Przykłady

1. **Podstawowy przykład `If`**:

```vba
Sub ProstyIf()
    Dim liczba As Integer
    liczba = 7
    
    If liczba > 5 Then
        MsgBox "Liczba jest większa niż 5."
    End If
End Sub
```

2. **Użycie `ElseIf`**:

```vba
Sub IfElse()
    Dim liczba As Integer
    liczba = 4
    
    If liczba > 5 Then
        MsgBox "Liczba jest większa niż 5."
    Else
        MsgBox "Liczba jest mniejsza lub równa 5."
    End If
End Sub
```

3. **Zagnieżdżone `If`**:

```vba
Sub ZagniezdzoneIf()
    Dim liczba As Integer
    liczba = 5
    
    If liczba > 5 Then
        MsgBox "Liczba jest większa niż 5."
    ElseIf liczba < 5 Then
        MsgBox "Liczba jest mniejsza niż 5."
    Else
        MsgBox "Liczba jest równa 5."
    End If
End Sub
```

#### Warte uwagi

- **Unikanie zagnieżdżenia**: Zbyt wiele poziomów zagnieżdżenia (IF w IFie) może uczynić kod trudnym do czytania i utrzymania. Dlatego warto rozważyć użycie `Select Case` lub rozbić kod na mniejsze procedury i funkcje.
- **Optymalizacja wydajności**: Gdy mamy wiele warunków, warto upewnić się, że są one sprawdzane w optymalnej kolejności, aby uniknąć niepotrzebnych sprawdzeń i zwiększyć wydajność kodu.
- **Czytelność**: Dbaj o czytelność, formatując kod i używając komentarzy w miejscach, które mogą być niejasne dla innych programistów lub dla Ciebie w przyszłości.

Instrukcje warunkowe IF są istotnym elementem kontroli przepływu programu w VBA, umożliwiającym tworzenie skryptów i makr, które mogą podejmować decyzje i wykonywać różne działania w zależności od spełnienia określonych warunków.

## 2. Działania na Wartościach Daty i Czasu w VBA

### Opis

Operacje na datach i czasie są krytycznym elementem wielu aplikacji bazodanowych, aplikacji finansowych oraz automatyzacji zadań w Excelu przy użyciu VBA (Visual Basic for Applications). VBA oferuje bogaty zestaw funkcji do manipulowania datami i czasem, co obejmuje obliczenia, formatowanie, a także wyodrębnianie konkretnych jednostek czasu, takich jak dni, miesiące czy lata.

### Kluczowe Funkcje

#### Now()

Funkcja `Now()` zwraca bieżącą datę i czas systemowy. Jest używana, gdy chcemy dokładny czas do sekundy.

#### Date()

`Date()` zwraca bieżącą datę systemową, bez informacji o czasie.

#### Time()

`Time()` dostarcza bieżący czas systemowy, bez informacji o dacie.

#### DateAdd()

`DateAdd()` pozwala dodawać lub odejmować jednostki czasu (sekundy, minuty, dni itp.) od określonej daty.

#### DateDiff()

`DateDiff()` oblicza różnicę między dwiema datami w wybranej jednostce czasu (np. w dniach).

### Formatowanie Dat i Czasu

VBA dostarcza także funkcji, takich jak `Format()`, do zmiany sposobu prezentowania dat i czasu, co pozwala dostosować ich wygląd do różnych kontekstów i standardów regionalnych.

### Przykłady

#### 1. Użycie Funkcji `Now()`, `Date()` i `Time()`

```vba
Sub ShowCurrentDateTime()
    MsgBox "Bieżąca data i czas: " & Now()
    MsgBox "Bieżąca data: " & Date()
    MsgBox "Bieżący czas: " & Time()
End Sub
```
Uwaga: Użycie tych funkcji bezpośrednio lub w kombinacji z innymi może dostarczyć potrzebnych danych czasowych.

#### 2. Operacje na Dacie za Pomocą `DateAdd()`

```vba
Sub AddDays()
    Dim futureDate As Date
    futureDate = DateAdd("d", 10, Date())
    MsgBox "Data za 10 dni to: " & futureDate
End Sub
```
W tym przypadku, "d" oznacza, że dodajemy dni. Możemy użyć różnych jednostek, np. "m" dla miesięcy czy "yyyy" dla lat.

#### 3. Formatowanie Daty

```vba
Sub FormatDateExample()
    Dim formattedDate As String
    formattedDate = Format(Date(), "dd-mm-yyyy")
    MsgBox "Sformatowana data: " & formattedDate
End Sub
```
Funkcja `Format()` umożliwia prezentowanie daty w różnych stylach i jest niezwykle użyteczna, gdy chcemy wyświetlić datę w czytelnym formacie lub zgodnie ze standardami regionalnymi.

#### 4. Różnica między Dwoma Datami

```vba
Sub DateDifference()
    Dim startDate As Date
    Dim endDate As Date
    Dim diff As Integer
    
    startDate = #1/1/2023#
    endDate = #12/31/2023#
    diff = DateDiff("d", startDate, endDate)
    
    MsgBox "Różnica dni między datami: " & diff
End Sub
```

`DateDiff()` jest wykorzystywany, gdy chcemy obliczyć różnicę między dwoma punktami czasowymi.

## 3. Ikony i Przyciski w Oknach Komunikatu w VBA

### Opis

Funkcja `MsgBox` jest niezbędnikiem programistycznym w VBA do interakcji z użytkownikami, dostarczając informacji, ostrzeżeń, błędów lub zbierając od nich proste odpowiedzi. Właściwości i funkcjonalność `MsgBox` mogą być dostosowane przez różne argumenty i flagi, umożliwiające prezentowanie różnych typów ikon oraz zestawów przycisków.

### Rodzaje Przycisków i Ikon w `MsgBox`

#### Typy Przycisków:
- `vbOkOnly`: Okno z jednym przyciskiem "OK".
- `vbOkCancel`: Okno z przyciskami "OK" i "Anuluj".
- `vbYesNo`: Okno z przyciskami "Tak" i "Nie".
- `vbYesNoCancel`: Okno z przyciskami "Tak", "Nie" i "Anuluj".
- `vbRetryCancel`: Okno z przyciskami "Ponów próbę" i "Anuluj".
- `vbAbortRetryIgnore`: Okno z przyciskami "Przerwij", "Ponów próbę" i "Ignoruj".

#### Typy Ikon:
- `vbCritical`: Ikona błędu.
- `vbQuestion`: Ikona pytania.
- `vbExclamation`: Ikona wykrzyknika (ostrzeżenie).
- `vbInformation`: Ikona informacyjna.

### Przykłady

#### 1. **Proste Okno Komunikatu**

Podstawowy przykład użycia `MsgBox` do przekazania informacji użytkownikowi.

```vba
Sub SimpleMsgBox()
    MsgBox "Operacja zakończona pomyślnie!", vbInformation, "Sukces"
End Sub
```
W tym przypadku, okno komunikatu będzie zawierało informacyjną ikonę i tytuł "Sukces".

#### 2. **Decyzja Użytkownika: Tak czy Nie**

Użycie `MsgBox` do uzyskania odpowiedzi od użytkownika i podjęcie działania na podstawie tej odpowiedzi.

```vba
Sub UserDecision()
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Czy chcesz kontynuować?", vbYesNo + vbQuestion, "Decyzja")

    If answer = vbYes Then
        MsgBox "Użytkownik wybrał TAK."
    Else
        MsgBox "Użytkownik wybrał NIE."
    End If
End Sub
```

#### 3. **Błąd z Opcją Ponowienia**

Zaprezentowanie błędu i danie użytkownikowi możliwości ponowienia akcji.

```vba
Sub ErrorWithRetry()
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Coś poszło nie tak. Chcesz spróbować ponownie?", vbRetryCancel + vbCritical, "Błąd")

    If answer = vbRetry Then
        MsgBox "Ponowienie akcji..."
    Else
        MsgBox "Anulowano przez użytkownika."
    End If
End Sub
```

#### 4. **Złożona Logika Decyzyjna**

Zastosowanie `MsgBox` w celu wykonania bardziej złożonej logiki decyzyjnej.

```vba
Sub ComplexDecisionLogic()
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Czy chcesz zapisać plik?", vbYesNoCancel + vbQuestion, "Zapisz plik")

    Select Case answer
        Case vbYes
            MsgBox "Plik został zapisany."
        Case vbNo
            MsgBox "Plik nie został zapisany."
        Case vbCancel
            MsgBox "Operacja anulowana przez użytkownika."
    End Select
End Sub
```

## 4. Lista zadań do samodzielnego wykonania 

#### 1. Zadania dotyczące instrukcji warunkowej `IF`

**Zadanie 1.1**: Utwórz procedurę, która sprawdzi, czy liczba jest dodatnia, ujemna czy równa zero. Użyj instrukcji `If...ElseIf...Else` i wyświetl odpowiedni komunikat za pomocą `MsgBox`.

**Zadanie 1.2**: Utwórz procedurę, która przyjmuje dwa argumenty typu `Integer` i za pomocą `MsgBox` informuje, który z nich jest większy lub czy są równe.

#### 2. Zadania dotyczące operacji na datach i czasie

**Zadanie 2.1**: Stwórz procedurę, która wyświetli informację o tym, ile dni pozostało do końca roku. Użyj funkcji `DateDiff` i aktualnej daty z funkcji `Date`.

**Zadanie 2.2**: Utwórz procedurę, która dodaje pewną liczbę dni (np. 15) do aktualnej daty i wyświetla wynik w formie komunikatu.

#### 3. Zadania dotyczące ikon i przycisków w oknach komunikatu

**Zadanie 3.1**: Zaprojektuj procedurę, która wyświetli okno komunikatu z pytaniem „Czy jesteś pewien?”. Okno powinno zawierać przyciski Tak/Nie oraz ikonę pytania.

**Zadanie 3.2**: Utwórz procedurę, która przy użyciu `MsgBox` wyświetli okno komunikatu z ikoną błędu i przyciskami Anuluj, Spróbuj ponownie, Kontynuuj. W zależności od wyboru użytkownika, pokaż odpowiadający mu komunikat.

