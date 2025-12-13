Sub Makro_Prio_5_wyciagniecie()
    'Macro to extract and process priority data from the "WYNIK" sheet
    'and save it to a new workbook.

    'Define variables in the macro
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim wbPriority As Workbook
    Dim lastRowSource As Long
    Dim lastRowDest As Long
    Dim sciezka As String
    Dim vlookupFormula As String
    
    'Turn off screen updating for performance
    application.ScreenUpdating = False
    
    'Workbook and worksheet source
    Set wbSource = ThisWorkbook
    Set wsSource = ActiveWorkbook.Sheets("WYNIK")
    
    'Instruction message
    MsgBox "Filter the planning column by finding priorities.", vbInformation, "Instrukcja"
    
    'Create new workbook and worksheet
    Dim wbDest As Workbook
    Set wbDest = Workbooks.Add
    Set wsDest = wbDest.Sheets(1)
    
    'The last row in source worksheet
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "T").End(xlUp).Row
    
    'Copy date from J:T to new workbook
    wsSource.Range("J3:T" & lastRowSource).Copy
    wsDest.Range("A1").PasteSpecial Paste:=xlPasteValues
    application.CutCopyMode = False
    
    'The last rows in target worksheet 
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row
    
    'Delete column B:J
    wsDest.Columns("B:J").Delete Shift:=xlToLeft
    
    'Debug
    ' Kopiuj wartości do nowego skoroszytu
    'copyRange.Copy
    'wsNew.Range("A1").PasteSpecial Paste:=xlPasteValues
    'application.CutCopyMode = False
    
    'Rozdziel tekst w kolumnie B w nowym skoroszycie
    With wsDest.Columns("B:B")
        .TextToColumns Destination:=.Cells(1, 1), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=True, Other:=True, OtherChar:="/", _
            FieldInfo:=Array(Array(1, 9), Array(2, 1)), TrailingMinusNumbers:=True
    End With
    
    'Save path 
    filePath = "C:\Users\robert.cwenar\Documents\" & _  ' <- Change to your correct path
           "prio zlecenia " & Format(Date, "dd.mm.yyyy") & ".xlsx"
    
    'Save new file
    application.DisplayAlerts = False
    wbDest.SaveAs filePath, FileFormat:=xlOpenXMLWorkbook
    application.DisplayAlerts = True
    MsgBox "Plik zapisany: " & filePath, vbInformation

    application.ScreenUpdating = True
End Sub



