Sub Makro_Prio_ydrzewo4()

    Dim wbZrod As Workbook      ' źródłowy (ydrzewo)
    Dim wbPRIO As Workbook      ' docelowy (prio)
    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim ws3 As Worksheet
    Dim wsArkusz1 As Worksheet
    Dim wb As Workbook
    Dim lastRowSrc As Long, lastRowDest As Long
    Dim idx As Long
    Dim nazwaNowegoArkusza As String
    
    MsgBox "Upewnij się, że masz otwarty plik: 'ydrzewo 4 ...' oraz plik 'prio'."
    
    ' === SZUKAMY OTWARTYCH PLIKÓW ===
    For Each wb In application.Workbooks
        If InStr(1, wb.Name, "ydrzewo 4", vbTextCompare) > 0 Then Set wbZrod = wb
        If InStr(1, wb.Name, "prio", vbTextCompare) > 0 Then Set wbPRIO = wb
    Next wb
    
    If wbZrod Is Nothing Or wbPRIO Is Nothing Then
        MsgBox "Nie znaleziono wymaganych plików!", vbCritical
        Exit Sub
    End If
    
    ' === ARKUSZ ŹRÓDŁOWY ===
    Set wsSrc = wbZrod.Sheets(1)
    
    ' === ARKUSZ1 W PRIO ===
    On Error Resume Next
    Set wsArkusz1 = wbPRIO.Sheets("Arkusz1")
    On Error GoTo 0
    If wsArkusz1 Is Nothing Then
        MsgBox "Brak arkusza 'Arkusz1' w pliku Prio!", vbCritical
        Exit Sub
    End If
    
    ' === USTALENIE/UTWORZENIE ARKUSZA DO KOPIOWANIA ===
    idx = wsArkusz1.Index
    If idx < wbPRIO.Sheets.Count Then
        Set wsDest = wbPRIO.Sheets(idx + 1)
    Else
        ' Tworzymy nowy arkusz z unikalną nazwą
        nazwaNowegoArkusza = "Arkusz" & wbPRIO.Sheets.Count + 1
        Set wsDest = wbPRIO.Sheets.Add(After:=wbPRIO.Sheets(wbPRIO.Sheets.Count))
        wsDest.Name = nazwaNowegoArkusza
    End If
    wsDest.Cells.Clear
    
    ' === KOPIOWANIE DANYCH Z YDRZEWO 4 ===
    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, "B").End(xlUp).Row
    wsSrc.Range("B6:K" & lastRowSrc).Copy
    wsDest.Range("A1").PasteSpecial xlPasteValues
    
    ' === FORMUŁY WYSZUKAJ.PIONOWO W PRIO ===
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row
    wsDest.Range("K1:K" & lastRowDest).FormulaLocal = "=WYSZUKAJ.PIONOWO(A1;Arkusz1!A:B;2;0)"
    
    ' === TWORZENIE / PRZYGOTOWANIE ARKUSZA 3 ===
    On Error Resume Next
    Set ws3 = wbPRIO.Sheets("Arkusz3")
    On Error GoTo 0
    If ws3 Is Nothing Then
        Set ws3 = wbPRIO.Sheets.Add(After:=wbPRIO.Sheets(wbPRIO.Sheets.Count))
        ws3.Name = "Arkusz3"
    Else
        ws3.Cells.Clear
    End If
    
    ' === KOPIOWANIE WARTOŚCI KOLUMN J:K DO ARKUSZA 3 ===
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "J").End(xlUp).Row
    wsDest.Range("J1:K" & lastRowDest).Copy
    ws3.Range("A2").PasteSpecial Paste:=xlPasteValues
    
    ' === NAGŁÓWKI I FILTROWANIE ===
    ws3.Range("A1").Value = "a"
    ws3.Range("B1").Value = "b"
    ws3.Range("A1:B1").AutoFilter
    
    ' === SORTOWANIE PO KOLUMNIE B ===
    ws3.Range("A:B").Sort Key1:=ws3.Range("B1"), Order1:=xlAscending, Header:=xlYes
    
    ' === USUWANIE DUPLIKATÓW W KOLUMNIE A ===
    ws3.Range("A:B").RemoveDuplicates Columns:=1, Header:=xlYes
    
    ' === CZYSZCZENIE SCHOWKA ===
    application.CutCopyMode = False
    
    MsgBox "Gotowe! Wszystkie operacje wykonane poprawnie w pliku Prio.", vbInformation

End Sub

