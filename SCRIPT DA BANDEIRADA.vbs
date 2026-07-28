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
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00017"
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00017"
session.findById("wnd[0]/usr/cmbDY_PARVW").key = "Z2"
session.findById("wnd[0]/usr/ctxtDY_PARNR").text = "csc354"
session.findById("wnd[0]/usr/ctxtDY_PARNR").setFocus
session.findById("wnd[0]/usr/ctxtDY_PARNR").caretPosition = 6
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1,"GSTRP"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "GSTRP"
session.findById("wnd[0]/tbar[1]/btn[40]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00022"
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00022"
session.findById("wnd[0]/usr/txtP_AUFNR").text = "100020542024"
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
