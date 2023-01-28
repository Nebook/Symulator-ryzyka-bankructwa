Sub symulator_ryzyka_bankructwa()

'DEFINIOWANIE ZMIENNYCH
Const MaxTransakcji = 10001 'definicja stałej
Dim WynikTransakcji(MaxTransakcji) As Long '(górna granica zmiennej tablicowej)
Dim LiniaKapitalu(MaxTransakcji) As Long '(górna granica zmiennej tablicowej)
Dim Trafnosc As Variant
Dim PrzewidywanyZysk_wspolczynnik As Variant
Dim Strategia_Zarzadzania_Pieniedzmi As String
Dim Staly_Procent_Ryzyka As Variant
Dim Punkt_Bankructwa As Variant
Dim Konto_Start As Variant
Dim Konto_Balans As Variant
Dim Konto_Nowy_Rekord As Variant
Dim Konto_Bankructwo As Variant
Dim Konto_Bankructwo_Procent As Variant
Dim Zysk_czy_Strata As Variant
Dim Szansa_na_Bankructwo As Variant
Dim NumerWiersza As Variant
Dim Jednostka_Pieniedzy As Integer 'ile jest zaangażowane w transakcje
Dim Staly_Dolar_Ryzyka As Variant
Dim Liczba_Tradow As Long
Dim Liczba_Tradow_przed_Bankructwem As Long
Dim Liczba_Tradow_od_maxKonta As Long
Dim i As Long
Dim j As Long
Dim x As Long
Application.DisplayAlerts = False 'WYŁĄCZENIE KOMUNIKATÓW OSTRZEGAWCZYCH
Application.ScreenUpdating = False 'WYŁĄCZENIE AKTUALIZACJI EKRANU

'WCZYTANIE ZMIENNYCH Z ARKUSZA
Sheets("Symulator").Select
Range("Trafnosc").Select
Trafnosc = Selection 'wczytanie trafności
Range("PrzewZysk").Select
PrzewidywanyZysk_wspolczynnik = Selection 'wczytanie średniego zysku/średniej straty

Range("Strategia_Zarzadzania_Pieniedzmi").Select
If ActiveCell = 1 Then
    Strategia_Zarzadzania_Pieniedzmi = "Stale ryzyko procentowe"
    Else
    Strategia_Zarzadzania_Pieniedzmi = "Stale ryzyko dolarowe"
End If                                  'wczytanie strategii zarządzania pieniędzmi

Range("Start_Kapital").Select
Konto_Start = Selection 'wczytanie wielkości kapitału
Range("Staly_Procent").Select
Staly_Procent_Ryzyka = Selection 'wczytanie stałej wielkości procentowej ryzyka każdej transakcji
Range("Bankructwo").Select
Punkt_Bankructwa = Selection 'wczytanie wielkości procentowej obsunięcia zdefiniowanego jako bankructwo
Range("Jednostka_Pieniedzy").Select
Jednostka_Pieniedzy = Selection 'wczytanie liczby jednostek kapitału na naszym rachunku

'WYCZYSZCZENIE TABLICY
For i = 1 To MaxTransakcji
    WynikTransakcji(i) = Empty
    LiniaKapitalu(i) = 0
Next i

'SYMULACJA PRAWDOPODOBIEŃSTWA BANKRUCTWA
Liczba_Tradow = 1
Konto_Balans = Konto_Start
Konto_Nowy_Rekord = Konto_Start
Konto_Bankructwo_Procent = 0
Liczba_Tradow_przed_Bankructwem = 0
Staly_Dolar_Ryzyka = Konto_Start / Jednostka_Pieniedzy
i = 1
j = 1
x = 0
Do Until Konto_Bankructwo_Procent >= Punkt_Bankructwa Or LiniaKapitalu(i - 1) > 200000000 Or x >= 10000

'Sprawdzenie czy kapitał osiągnął nowe maksimum i wyczyść liczbę stratnych transakcji do 0
If Konto_Balans > Konto_Nowy_Rekord Then
    Konto_Nowy_Rekord = Konto_Balans
    Liczba_Tradow_przed_Bankructwem = 0
    Liczba_Tradow_od_maxKonta = 0
End If

'Generowanie losowych liczb w celu sprawdzenia czy transakcje zyskują, czy tracą
Zysk_czy_Strata = Rnd

'Sprawdzenie czy transakcja zyskała
If Zysk_czy_Strata >= (1 - Trafnosc) Then
'ZYSK!
    'Obliczenie zysku
    If Strategia_Zarzadzania_Pieniedzmi = "Stale ryzyko procentowe" Then
    WynikTransakcji(j) = ((Staly_Procent_Ryzyka * Konto_Balans) * PrzewidywanyZysk_wspolczynnik)
    End If

    If Strategia_Zarzadzania_Pieniedzmi = "Stale ryzyko dolarowe" Then
    WynikTransakcji(j) = Staly_Dolar_Ryzyka * PrzewidywanyZysk_wspolczynnik
    End If

    'Dopisanie do linii kapitału
    If i = 1 Then
        LiniaKapitalu(i) = Konto_Start
        i = i + 1
        LiniaKapitalu(i) = LiniaKapitalu(i - 1) + WynikTransakcji(j)
        Else
        LiniaKapitalu(i) = LiniaKapitalu(i - 1) + WynikTransakcji(j)
        End If

    'Dopisanie do stanu rachunku
    Konto_Balans = Konto_Balans + WynikTransakcji(j)
    Else

    'STRATA!
        'Obliczenie straty
        If Strategia_Zarzadzania_Pieniedzmi = "Stale ryzyko procentowe" Then
            WynikTransakcji(j) = -(Staly_Procent_Ryzyka * Konto_Balans)
        End If

        If Strategia_Zarzadzania_Pieniedzmi = "Stale ryzyko dolarowe" Then
            WynikTransakcji(j) = -Staly_Dolar_Ryzyka
        End If

        'Dopisanie do linii kapitału
        If i = 1 Then
            LiniaKapitalu(i) = Konto_Start
            i = i + 1
            LiniaKapitalu(i) = LiniaKapitalu(i - 1) + WynikTransakcji(j)
            Else
            LiniaKapitalu(i) = LiniaKapitalu(i - 1) + WynikTransakcji(j)
        End If

        'Dopisanie do stanu rachunku
        Konto_Balans = Konto_Balans + WynikTransakcji(j)

        'Obliczenie aktualnego obsunięcia nominalnie i procentowo
        Konto_Bankructwo = Konto_Nowy_Rekord - Konto_Balans
        Konto_Bankructwo_Procent = Konto_Bankructwo / Konto_Nowy_Rekord

        'Obliczenie liczby strat zanim dojdzie do bankructwa
        Liczba_Tradow_przed_Bankructwem = Liczba_Tradow_przed_Bankructwem + 1
        
        End If

        'Obliczenie liczby transakcji
        Liczba_Tradow = Liczba_Tradow + 1
        Liczba_Tradow_od_maxKonta = Liczba_Tradow_od_maxKonta + 1

        'Kolejne transakcje
        x = x + 1
        j = j + 1
        i = i + 1
        Loop

    'Obliczenie prawdopodobieństwa bankructwa
        Szansa_na_Bankructwo = Liczba_Tradow_przed_Bankructwem / Liczba_Tradow_od_maxKonta
    
    'Jeśli linia kapitału przekracza 200 milionów, lub udało się zasymulować 10 000 transakcji zakładamy, że bankructwa dało się uniknąć.
    If LiniaKapitalu(i - 1) > 200000000 Or x >= 10000 Then
    Szansa_na_Bankructwo = 0
    End If

    'Policz prawdopodobieństwo bankructwa w arkuszu
    Sheets("Symulator").Select
    Range("Prawdopodobienstwo").Select
    ActiveCell = Szansa_na_Bankructwo
    Selection.Style = "Percent"

    'Rysuj krzywą kapitału

    'Wyczyść wcześniejszą linię kapitału
    Columns("AA:AA").Select
    Selection.Clear

    'Pokaż linię kapitału w arkuszku - Kolumna AA
    i = 1
    Do Until i >= Liczba_Tradow + 1
        Sheets(1).Cells(i, 27).Value = LiniaKapitalu(i)
    i = i + 1
    Loop

    'Zmiana zakresu wykresu
    Range("AA1").Select
    Selection.End(xlDown).Select
    NumerWiersza = ActiveCell.Row

    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.PlotArea.Select
    ActiveChart.SeriesCollection(1).Values = "=Symulator!R1C27:R" & NumerWiersza & "C27"
    ActiveWindow.Visible = False
    Windows("Kozicki_SymulatorRyzykaBankructwa.xlsm").Activate

    'Ustaw kursor w komórce B22
    Range("B22").Select

    'Odśwież ekran
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

