Sub Makro_Prio_ydrzewo4()

    'Definition variables in macro
    Dim wbZrod As Workbook      'source (ydrzewo)
    Dim wbPRIO As Workbook      'target (prio)
    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim ws3 As Worksheet
    Dim wsArkusz1 As Worksheet
    Dim wb As Workbook
    Dim lastRowSrc As Long, lastRowDest As Long
    Dim idx As Long
    Dim nazwaNowegoArkusza As String

    'Added initial message
    MsgBox "Make sure that the file 'ydrzewo 4...' and the 'prio' file are open.'."
    
    'Look for required workbooks
    For Each wb In application.Workbooks
        If InStr(1, wb.Name, "ydrzewo 4", vbTextCompare) > 0 Then Set wbZrod = wb
        If InStr(1, wb.Name, "prio", vbTextCompare) > 0 Then Set wbPRIO = wb
    Next wb
    
    If wbZrod Is Nothing Or wbPRIO Is Nothing Then
        MsgBox "Nie znaleziono wymaganych plików!", vbCritical
        Exit Sub
    End If
    
    'source worksheet
    Set wsSrc = wbZrod.Sheets(1)
    
    'worksheet "Arkusz1" in prio
    On Error Resume Next
    Set wsArkusz1 = wbPRIO.Sheets("Arkusz1")
    On Error GoTo 0
    If wsArkusz1 Is Nothing Then
        MsgBox "Worksheet 'Arkusz1' not found in the Prio file.", vbCritical
        Exit Sub
    End If
    
    'Addressing destination worksheet
    idx = wsArkusz1.Index
    If idx < wbPRIO.Sheets.Count Then
        Set wsDest = wbPRIO.Sheets(idx + 1)
    Else
        'Create new sheet if Arkusz1 is the last sheet
        nazwaNowegoArkusza = "Arkusz" & wbPRIO.Sheets.Count + 1
        Set wsDest = wbPRIO.Sheets.Add(After:=wbPRIO.Sheets(wbPRIO.Sheets.Count))
        wsDest.Name = nazwaNowegoArkusza
    End If
    wsDest.Cells.Clear
    
    'Copying data from source to destination
    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, "B").End(xlUp).Row
    wsSrc.Range("B6:K" & lastRowSrc).Copy
    wsDest.Range("A1").PasteSpecial xlPasteValues
    
    'Adding VLOOKUP formula in column K in destination sheet
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row
    wsDest.Range("K1:K" & lastRowDest).FormulaLocal = "=WYSZUKAJ.PIONOWO(A1;Arkusz1!A:B;2;0)"
    
    'Creating or clearing "Arkusz3" in prio
    On Error Resume Next
    Set ws3 = wbPRIO.Sheets("Arkusz3")
    On Error GoTo 0
    If ws3 Is Nothing Then
        Set ws3 = wbPRIO.Sheets.Add(After:=wbPRIO.Sheets(wbPRIO.Sheets.Count))
        ws3.Name = "Arkusz3"
    Else
        ws3.Cells.Clear
    End If
    
    'Copying columns J and K from destination to Arkusz3
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "J").End(xlUp).Row
    wsDest.Range("J1:K" & lastRowDest).Copy
    ws3.Range("A2").PasteSpecial Paste:=xlPasteValues
    
    'Adding headers in Arkusz3
    ws3.Range("A1").Value = "a"
    ws3.Range("B1").Value = "b"
    ws3.Range("A1:B1").AutoFilter
    
    'Sorting Arkusz3 by column B
    ws3.Range("A:B").Sort Key1:=ws3.Range("B1"), Order1:=xlAscending, Header:=xlYes
    
    'Deleting duplicates in Arkusz3 based on column A
    ws3.Range("A:B").RemoveDuplicates Columns:=1, Header:=xlYes
    
    'Clean up
    application.CutCopyMode = False
    
    'Added final message
    MsgBox "Done! All operations have been completed successfully in the Prio file.", vbInformation

End Sub

