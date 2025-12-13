'Current date and unique file name
Dim dzien, miesiac, rok, dzis
dzien = Right("0" & Day(Date), 2)
miesiac = Right("0" & Month(Date), 2)
rok = Right(Year(Date), 2)
dzis = dzien & "." & miesiac & "." & rok

Dim fso, folderPath, numer, fileName, f, re, maxNum, matches
Set fso = CreateObject("Scripting.FileSystemObject")
folderPath = "C:\Users\robert.cwenar\" <- ' add your path where it will be save!
If Right(folderPath,1) <> "\" Then folderPath = folderPath & "\"

'Add another number
Set re = New RegExp
re.Pattern = "prio (\d+) z dnia " & dzis & "\.xls"
re.IgnoreCase = True
re.Global = False

maxNum = 0
For Each f In fso.GetFolder(folderPath).Files
    If re.Test(f.Name) Then
        Set matches = re.Execute(f.Name)
        If matches.Count > 0 Then
            If CInt(matches(0).SubMatches(0)) > maxNum Then
                maxNum = CInt(matches(0).SubMatches(0))
            End If
        End If
    End If
Next
'add file name 
numer = maxNum + 1
fileName = "prio " & numer & " z dnia " & dzis & ".xls"


'Start SAP GUI scripting
Dim SapGuiAuto, application, connection, potentialSession, sessionsList(), i, choice, session, sessionFound
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine

'Collect all active sessions
ReDim sessionsList(0)
i = 0
For Each connection In application.Children
    For Each potentialSession In connection.Children
        ReDim Preserve sessionsList(i)
        Set sessionsList(i) = potentialSession
        i = i + 1
    Next
Next

If i = 0 Then
    MsgBox "No active SAP session found."
    WScript.Quit
End If

'Display list of sessions and selection 
Dim msg
msg = "Select the SAP sessions to use:" & vbCrLf
For i = 0 To UBound(sessionsList)
    msg = msg & i+1 & ". Sessions #" & i+1 & vbCrLf
Next

'Get user choice
choice = InputBox(msg, "Choose sessions", "1")
If choice = "" Then WScript.Quit
choice = CInt(choice) - 1

'Attempt to connect to the selected session or next active one
sessionFound = False
For i = choice To UBound(sessionsList)
    On Error Resume Next
    Set session = sessionsList(i)
    session.findById("wnd[0]").maximize
    If Err.Number = 0 Then
        sessionFound = True
        Exit For
    Else
        Err.Clear
    End If
Next

'If no session found, exit
If Not sessionFound Then
    MsgBox "Failed to connect to the selected or next SAP sessionP."
    WScript.Quit
End If

If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject application, "on"
End If

MsgBox "Connected to SAP session #" & i+1

'Export data from SAP to Excel
On Error Resume Next

'Maximize the main SAP window
session.findById("wnd[0]").maximize

'Set the current cell in the grid to the column "DELKZ"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1,"DELKZ"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "DELKZ"
session.findById("wnd[0]/tbar[1]/btn[29]").press

'In the filter dialog, set the filter value to "FE"
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "FE"
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 2
session.findById("wnd[1]").sendVKey 0

'Press the export button
session.findById("wnd[0]/tbar[1]/btn[45]").press

session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press

'Specify file path and name for export
session.findById("wnd[1]/usr/ctxtDY_PATH").text = folderPath
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fileName
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = Len(fileName)
session.findById("wnd[1]/tbar[0]/btn[0]").press

On Error GoTo 0
