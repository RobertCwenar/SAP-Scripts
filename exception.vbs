'Dynamic name and generation for exception report export
Dim dzien, miesiac, rok, dzis
dzien = Right("0" & Day(Date), 2)
miesiac = Right("0" & Month(Date), 2)
rok = Right(Year(Date), 2)
dzis = dzien & "." & miesiac & "." & rok  ' dd.mm.yy

'Display the folder path and generate unique file name
Dim fso, folderPath, fileName, numer
Set fso = CreateObject("Scripting.FileSystemObject")
folderPath = "C:\Users\robert.cwenar\" 'Change to your target folder path
If Right(folderPath,1) <> "\" Then folderPath = folderPath & "\"

'Generate unique file name with today date
numer = 0
fileName = "wyjatki " & dzis & ".xls"
Do While fso.FileExists(folderPath & fileName)
    numer = numer + 1
    fileName = "wyjatki " & dzis & " (" & numer & ").xls"
Loop

'Incoming connection to SAP GUI and session
Dim SapGuiAuto, application, connection, potentialSession, sessionsList, i, choice, session, sessionFound
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine

sessionsList = Array()
i = 0
For Each connection In application.Children
    For Each potentialSession In connection.Children
        ReDim Preserve sessionsList(i)
        Set sessionsList(i) = potentialSession
        i = i + 1
    Next
Next

If i = 0 Then
    MsgBox "No active SAP session found in SAP."
    WScript.Quit
End If

'Display list of sessions and selection
Dim msg
msg = "Choose the SAP session to use:" & vbCrLf
For i = 0 To UBound(sessionsList)
    msg = msg & i+1 & ". Session #" & i+1 & vbCrLf
Next

choice = InputBox(msg, "Choose session", "1")
If choice = "" Then WScript.Quit
choice = CInt(choice) - 1

'Try to connect to the selected session or next available
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

If Not sessionFound Then
    MsgBox "Failed to connect to the selected or next SAP session."
    WScript.Quit
End If

If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject application, "on"
End If

'Information about successful connection
MsgBox "Connected with SAP GUI session #" & i+1

'Start exception report export process
On Error Resume Next
session.findById("wnd[0]/tbar[0]/okcd").text = "SE16N"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/lbl[2,4]").setFocus
session.findById("wnd[1]/usr/lbl[2,4]").caretPosition = 8
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/txtGD-NUMBER").setFocus
session.findById("wnd[0]/usr/txtGD-NUMBER").caretPosition = 0
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = ""
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,14]").text = "01"
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-HIGH[3,14]").text = "99"
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-HIGH[3,14]").setFocus
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-HIGH[3,14]").caretPosition = 2
session.findById("wnd[0]/tbar[1]/btn[8]").press

'Export report data to Excel with dynamic file name
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem "&PC"
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press

'Save exported file with dynamic name
session.findById("wnd[1]/usr/ctxtDY_PATH").text = folderPath
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fileName
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = Len(fileName)
session.findById("wnd[1]/tbar[0]/btn[0]").press
On Error GoTo 0
