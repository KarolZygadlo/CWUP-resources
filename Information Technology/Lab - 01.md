### 1. Uruchomienie edytora VBA, Procedura i Instrukcja Wyjścia, Sposoby Uruchamiania Makr

#### Uruchomienie Edytora VBA

Edytor VBA (Visual Basic for Applications) można uruchomić bezpośrednio z programów pakietu Microsoft Office, takich jak Excel. Aby to zrobić:

- Otwórz Excel.
- Naciśnij klawisze `Alt` + `F11`, aby otworzyć edytor VBA.

#### Procedura i Instrukcja Wyjścia

**Procedura** to zbiór instrukcji, które można wykonać jako jedną jednostkę pracy. W VBA mamy procedury **Sub** oraz **Function**.

- **Sub** - używane do wykonywania czynności, które nie zwracają wartości.
- **Function** - używane do wykonywania czynności, które zwracają wartość.

```vba
Sub PrzykladowaProcedura()
    MsgBox "Cześć, to jest przykładowa procedura!"
End Sub
```

```vba
Function Dodaj(a As Integer, b As Integer) As Integer
    Dodaj = a + b
End Function
```

**Instrukcja wyjścia**: `Exit Sub` lub `Exit Function` używane są do natychmiastowego wyjścia z procedury.

```vba
Sub PrzykladWyjscia()
    On Error GoTo ErrorHandler
    MsgBox 1 / 0
    
ErrorHandler:
    MsgBox "Wystąpił błąd!"
    Exit Sub
End Sub
```

#### Sposoby Uruchamiania Makr

- Bezpośrednio z edytora VBA: Umieść kursor wewnątrz procedury i naciśnij `F5`.
- Przypisanie do przycisku lub innego obiektu na arkuszu Excela.
- Używając skrótów klawiaturowych, które można przypisać makru.

### 2. Operacje Arytmetyczne, Zmienne, Ich Deklarowanie i Wymuszanie Deklarowania

#### Operacje Arytmetyczne

Podstawowe operacje to `+`, `-`, `*`, `/`.

```vba
Sub OperacjeArytmetyczne()
    Dim a As Integer, b As Integer, suma As Integer
    a = 10
    b = 20
    suma = a + b
    MsgBox suma
End Sub
```

Makra, które przeprowadza różne operacje arytmetyczne:

```vba
Sub RozneOperacje()
    Dim a As Integer: a = 10
    Dim b As Integer: b = 20
    
    MsgBox "Dodawanie: " & a + b
    MsgBox "Odejmowanie: " & a - b
    MsgBox "Mnożenie: " & a * b
    MsgBox "Dzielenie: " & a / b
End Sub
```

#### Zmienne i Ich Deklarowanie

Zmienną deklarujemy, używając słowa `Dim`, a następnie nazwy zmiennej i typu.

```vba
Dim liczba As Integer
```

```vba
Sub DeklaracjaZmiennych()
    Dim imie As String, wiek As Integer, wzrost As Double
    
    imie = "Jan"
    wiek = 30
    wzrost = 1.8
    
    MsgBox "Imię: " & imie & ", Wiek: " & wiek & ", Wzrost: " & wzrost
End Sub
```
#### Wymuszanie Deklarowania Zmiennych

Aby wymusić deklarowanie zmiennych, na początku modułu należy umieścić `Option Explicit`.

```vba
Option Explicit

Sub Przyklad()
    Dim a As Integer
    a = 10
    MsgBox a
End Sub
```

### 3. Typy Zmiennych, Zmienne Tekstowe, Wprowadzenie do Funkcji i Sposobu Ich Edycji

#### Typy Zmiennych

Podstawowe typy zmiennych w VBA to:

- Integer
- String
- Boolean
- Double
- Variant (domyślny typ, jeśli nie określono inaczej)

#### Zmienne Tekstowe

Używane do przechowywania ciągów znaków.

```vba
Dim napis As String
napis = "To jest tekst"
MsgBox napis
```

#### Concatenation (Łączenie Tekstów)

Zmienne tekstowe możemy łączyć za pomocą operatora `&`.

```vba
Sub LaczenieTekstu()
    Dim czesc1 As String, czesc2 As String, calosc As String
    
    czesc1 = "Cześć "
    czesc2 = "Świecie!"
    calosc = czesc1 & czesc2
    
    MsgBox calosc
End Sub
```
#### Funkcje i Ich Edycja

Funkcje są procedurami, które zwracają wartość.

```vba
Function Kwadrat(liczba As Double) As Double
    Kwadrat = liczba ^ 2
End Function
```

Funkcja, która zwraca długość ciągu znaków:

```vba
Function DlugoscTekstu(tekst As String) As Integer
    DlugoscTekstu = Len(tekst)
End Function
```

Funkcja, która zwraca mniejszą z dwóch liczb:

```vba
Function MniejszaLiczba(a As Double, b As Double) As Double
    If a < b Then
        MniejszaLiczba = a
    Else
        MniejszaLiczba = b
    End If
End Function
```

Edycję funkcji przeprowadza się w edytorze VBA, zmieniając kod źródłowy i testując zmiany poprzez uruchomienie procedury testowej lub korzystając z Debugowania (F8).