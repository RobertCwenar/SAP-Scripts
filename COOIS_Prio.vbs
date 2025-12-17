'Automatically generated name and date
Dim day, month, year, today
day = Right("0" & Day(Date), 2)
month = Right("0" & Month(Date), 2)
year = Right(Year(Date), 2)
today = day & "." & month & "." & year

Dim fso, folderPath, number, fileName
Set fso = CreateObject("Scripting.FileSystemObject")
folderPath = "C:\Users\robert.cwenar" '<- choose the correct folder path
If Right(folderPath,1) <> "\" Then folderPath = folderPath & "\"

'Generate unique file name
number = 0
fileName = "COOIS prio z dnia " & today & ".xls"
Do While fso.FileExists(folderPath & fileName)
    number = number + 1
    fileName = "COOIS prio z dnia " & today & " (" & number & ").xls"
Loop

'Choose and connect to SAP GUI session
Dim SapGuiAuto, application, connection, potentialSession, sessionsList(), i, choice, session, sessionFound
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine

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
msg = "Select SAP session to use:" & vbCrLf
For i = 0 To UBound(sessionsList)
    msg = msg & i+1 & ". Session #" & i+1 & vbCrLf
Next
' Choice of session
choice = InputBox(msg, "Choose session", "1")
If choice = "" Then WScript.Quit
choice = CInt(choice) - 1

'Try to connect to the selected session or next active one
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
    MsgBox "Couldn't connect to the selected or subsequent SAP session."
    WScript.Quit
End If

If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject application, "on"
End If

MsgBox "Connect with SAP session #" & i+1

'COOIS Report Execution
On Error Resume Next
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "coois"
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/txtS_APRIO-LOW").text = "1"
session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/txtS_APRIO-HIGH").text = "9"
session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST1").text = "ZTCH"
session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST2").text = "ZAK"
session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E1").selected = True
session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E2").selected = True

'Start report
session.findById("wnd[0]/tbar[1]/btn[8]").press

'Export file from SAP COOIS
Dim shell
Set shell = session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell")

shell.pressToolbarButton "&NAVIGATION_PROFILE_TOOLBAR_EXPAND"
shell.pressToolbarContextButton "&MB_EXPORT"
shell.selectContextMenuItem "&PC"

session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press

'Save
session.findById("wnd[1]/usr/ctxtDY_PATH").text = folderPath
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fileName
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = Len(fileName)
session.findById("wnd[1]/tbar[0]/btn[0]").press

On Error GoTo 0