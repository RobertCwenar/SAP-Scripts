'Check and connect to SAP GUI Scripting engine
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

' Maximize the main SAP window
session.findById("wnd[0]").maximize

' Select all rows in the grid and open details
session.findById("wnd[0]/usr/cntlGRID_HEAD/shellcont/shell").setCurrentCell -1,""
session.findById("wnd[0]/usr/cntlGRID_HEAD/shellcont/shell").selectAll
session.findById("wnd[0]/usr/cntlGRID_HEAD/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/tbar[1]/btn[9]").press

' Export the grid data
session.findById("wnd[0]/usr/cntlGRID_COMP/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlGRID_COMP/shellcont/shell").selectContextMenuItem "&PC"

' Choose export options
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press

' Specify the file name and save
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Braki.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 6
session.findById("wnd[1]/tbar[0]/btn[0]").press