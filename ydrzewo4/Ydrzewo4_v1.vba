Sub Makro12_ydrzewo4_v1()
    ' Makro do nadawania priorytetów na Northvolta
    ' Płynne kopiowanie danych z ydrzewa4 do PRIO, pokazuje kroki

    Dim wbZrod As Workbook
    Dim wbPRIO As Workbook
    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim ws3 As Worksheet
    Dim wsArkusz1 As Worksheet
    Dim wb As Workbook
    Dim lastRowSrc As Long, lastRowDest As Long
    Dim idx As Long
    Dim nazwaNowegoArkusza As String
    Dim today As Date
    Dim ydrzewoPath As String
    
    today = Date
    application.ScreenUpdating = False
    application.Calculation = xlCalculationManual
    application.StatusBar = "Rozpoczynam makro..."

    ' --- Szukanie otwartych plików ---
    For Each wb In application.Workbooks
        If InStr(1, wb.Name, "ydrzewo 4", vbTextCompare) > 0 Then Set wbZrod = wb
        If InStr(1, LCase(wb.Name), "prio ", vbTextCompare) > 0 Then Set wbPRIO = wb
    Next wb

    ' --- Otwarcie ydrzewa4 jeśli nie jest otwarte ---
    If wbZrod Is Nothing Then
        ydrzewoPath = "C:\Users\robert.cwenar\Documents\SAP\SAP GUI\ydrzewo 4 z d " & Format(today, "dd.mm.yy") & ".xls"
        On Error Resume Next
        Set wbZrod = Workbooks.Open(ydrzewoPath)
        On Error GoTo 0
        If wbZrod Is Nothing Then
            MsgBox "Nie znaleziono pliku: " & ydrzewoPath, vbCritical
            GoTo Cleanup
        Else
            application.StatusBar = "Plik '" & wbZrod.Name & "' został otwarty automatycznie..."
            DoEvents
        End If
    End If

    ' --- Sprawdzenie PRIO ---
    If wbPRIO Is Nothing Then
        MsgBox "Nie znaleziono otwartego pliku PRIO!", vbCritical
        GoTo Cleanup
    End If

    ' --- Arkusz źródłowy ---
    Set wsSrc = wbZrod.Sheets(1)
    application.StatusBar = "Ustawiono arkusz źródłowy..."
    DoEvents

    ' --- Arkusz1 w PRIO ---
    On Error Resume Next
    Set wsArkusz1 = wbPRIO.Sheets("Arkusz1")
    On Error GoTo 0
    If wsArkusz1 Is Nothing Then
        MsgBox "Brak arkusza 'Arkusz1' w pliku PRIO!", vbCritical
        GoTo Cleanup
    End If
    application.StatusBar = "Arkusz1 w PRIO znaleziony..."
    DoEvents

    ' --- Ustalenie arkusza docelowego ---
    idx = wsArkusz1.Index
    If idx < wbPRIO.Sheets.Count Then
        Set wsDest = wbPRIO.Sheets(idx + 1)
    Else
        nazwaNowegoArkusza = "Arkusz" & wbPRIO.Sheets.Count + 1
        Set wsDest = wbPRIO.Sheets.Add(After:=wbPRIO.Sheets(wbPRIO.Sheets.Count))
        wsDest.Name = nazwaNowegoArkusza
    End If
    wsDest.Cells.Clear
    application.StatusBar = "Utworzono arkusz docelowy..."
    DoEvents

    ' --- Kopiowanie danych bez migotania ---
    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, "B").End(xlUp).Row
    wsDest.Range("A1").Resize(lastRowSrc - 5, 10).Value = wsSrc.Range("B6:K" & lastRowSrc).Value
    application.StatusBar = "Dane skopiowane do PRIO..."
    DoEvents

    ' --- Formuły WYSZUKAJ.PIONOWO ---
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row
    Dim i As Long
    For i = 1 To lastRowDest
        wsDest.Cells(i, "K").FormulaLocal = "=WYSZUKAJ.PIONOWO(A" & i & ";Arkusz1!A:B;2;0)"
    Next i
    application.StatusBar = "Formuły wstawione..."
    DoEvents

    ' --- Arkusz3 ---
    On Error Resume Next
    Set ws3 = wbPRIO.Sheets("Arkusz3")
    On Error GoTo 0
    If ws3 Is Nothing Then
        Set ws3 = wbPRIO.Sheets.Add(After:=wbPRIO.Sheets(wbPRIO.Sheets.Count))
        ws3.Name = "Arkusz3"
    Else
        ws3.Cells.Clear
    End If
    application.StatusBar = "Arkusz3 przygotowany..."
    DoEvents

    ' --- Kopiowanie wartości J:K do Arkusz3 ---
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "J").End(xlUp).Row
    ws3.Range("A2").Resize(lastRowDest, 2).Value = wsDest.Range("J1:K" & lastRowDest).Value
    application.StatusBar = "Dane skopiowane do Arkusz3..."
    DoEvents

    ' --- Nagłówki i filtrowanie ---
    ws3.Range("A1").Value = "a"
    ws3.Range("B1").Value = "b"
    ws3.Range("A1:B1").AutoFilter
    application.StatusBar = "Nagłówki i filtr ustawione..."
    DoEvents

    ' --- Sortowanie i usuwanie duplikatów ---
    ws3.Range("A:B").Sort Key1:=ws3.Range("B1"), Order1:=xlAscending, Header:=xlYes
    ws3.Range("A:B").RemoveDuplicates Columns:=1, Header:=xlYes
    application.StatusBar = "Dane posortowane i duplikaty usunięte..."
    DoEvents

    application.CutCopyMode = False
    application.StatusBar = "Gotowe!"
    
Cleanup:
    application.ScreenUpdating = True
    application.Calculation = xlCalculationAutomatic
    application.StatusBar = False
    
    MsgBox "Gotowe! Wszystkie operacje wykonane poprawnie w pliku PRIO.", vbInformation
End Sub
