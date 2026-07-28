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
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00023"
session.findById("wnd[0]/usr/ctxtS_EQUIPE-LOW").text = "csc300"
session.findById("wnd[0]/usr/ctxtS_EQUIPE-LOW").caretPosition = 6
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlCONTAINER_ORDEM/shellcont/shell").setCurrentCell 6,"DESC_STATUS"
session.findById("wnd[0]/usr/cntlCONTAINER_ORDEM/shellcont/shell").selectedRows = "6"
session.findById("wnd[0]/usr/cntlCONTAINER_ORDEM/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlCONTAINER_ORDEM/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
