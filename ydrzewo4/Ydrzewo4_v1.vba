Sub Makro12_ydrzewo4_v1()
    ' Macro for assigning priorities to Northvolta
    ' Smooth data copying from ydrzewo4 to PRIO, shows steps

    'Definition variables in macro
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

    'Look for required workbooks
    For Each wb In application.Workbooks
        If InStr(1, wb.Name, "ydrzewo 4", vbTextCompare) > 0 Then Set wbZrod = wb
        If InStr(1, LCase(wb.Name), "prio ", vbTextCompare) > 0 Then Set wbPRIO = wb
    Next wb

    'Open ydrzewo4 if not found
    If wbZrod Is Nothing Then
        ydrzewoPath = "C:\Users\robert.cwenar\Documents\ydrzewo 4 z d " & Format(today, "dd.mm.yy") & ".xls" '<- Define your path here
        On Error Resume Next
        Set wbZrod = Workbooks.Open(ydrzewoPath)
        On Error GoTo 0
        If wbZrod Is Nothing Then
            MsgBox "File not found: " & ydrzewoPath, vbCritical
            GoTo Cleanup
        Else
            application.StatusBar = "File '" & wbZrod.Name & "' was opened automatically..."
            DoEvents
        End If
    End If

    'Check that PRIO is open
    If wbPRIO Is Nothing Then
        MsgBox "Not found the open file Prio!", vbCritical
        GoTo Cleanup
    End If

    ' Worksheet source
    Set wsSrc = wbZrod.Sheets(1)
    application.StatusBar = "Ustawiono arkusz źródłowy..."
    DoEvents

    'Arkusz1 in PRIO
    On Error Resume Next
    Set wsArkusz1 = wbPRIO.Sheets("Arkusz1")
    On Error GoTo 0
    If wsArkusz1 Is Nothing Then
        MsgBox "Missing sheet 'Arkusz1' in PRIO file!", vbCritical
        GoTo Cleanup
    End If
    application.StatusBar = "Arkusz1 found in PRIO..."
    DoEvents

    'Find or create target sheet in PRIO
    idx = wsArkusz1.Index
    If idx < wbPRIO.Sheets.Count Then
        Set wsDest = wbPRIO.Sheets(idx + 1)
    Else
        nazwaNowegoArkusza = "Arkusz" & wbPRIO.Sheets.Count + 1
        Set wsDest = wbPRIO.Sheets.Add(After:=wbPRIO.Sheets(wbPRIO.Sheets.Count))
        wsDest.Name = nazwaNowegoArkusza
    End If
    wsDest.Cells.Clear
    application.StatusBar = "The target sheet has been created..."
    DoEvents

    'Copy data from ydrzewo4 to PRIO
    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, "B").End(xlUp).Row
    wsDest.Range("A1").Resize(lastRowSrc - 5, 10).Value = wsSrc.Range("B6:K" & lastRowSrc).Value
    application.StatusBar = "Dane skopiowane do PRIO..."
    DoEvents

    'Formula vlookup in column K
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row
    Dim i As Long
    For i = 1 To lastRowDest
        wsDest.Cells(i, "K").FormulaLocal = "=WYSZUKAJ.PIONOWO(A" & i & ";Arkusz1!A:B;2;0)"
    Next i
    application.StatusBar = "Formulas inserted..."
    DoEvents

    'Arkusz3 in PRIO
    On Error Resume Next
    Set ws3 = wbPRIO.Sheets("Arkusz3")
    On Error GoTo 0
    If ws3 Is Nothing Then
        Set ws3 = wbPRIO.Sheets.Add(After:=wbPRIO.Sheets(wbPRIO.Sheets.Count))
        ws3.Name = "Arkusz3"
    Else
        ws3.Cells.Clear
    End If
    application.StatusBar = "Arkusz3 prepared..."
    DoEvents

    'Copy columns J and K to Arkusz3
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "J").End(xlUp).Row
    ws3.Range("A2").Resize(lastRowDest, 2).Value = wsDest.Range("J1:K" & lastRowDest).Value
    application.StatusBar = "Data copied to Arkusz3..."
    DoEvents

    ' Adding headers and autfilter in Arkusz3
    ws3.Range("A1").Value = "a"
    ws3.Range("B1").Value = "b"
    ws3.Range("A1:B1").AutoFilter
    application.StatusBar = "Headlines and filter set..."
    DoEvents

    'Sort and remove duplicates in Arkusz3
    ws3.Range("A:B").Sort Key1:=ws3.Range("B1"), Order1:=xlAscending, Header:=xlYes
    ws3.Range("A:B").RemoveDuplicates Columns:=1, Header:=xlYes
    application.StatusBar = "Data sorted and deleted duplicates.."
    DoEvents

    application.CutCopyMode = False
    application.StatusBar = "Gotowe!"
    
    'Clean up and finish
Cleanup:
    application.ScreenUpdating = True
    application.Calculation = xlCalculationAutomatic
    application.StatusBar = False
    
    MsgBox "Done! Every operations do in file Prio.", vbInformation
End Sub
