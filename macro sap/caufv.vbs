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
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse16"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "caufv"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
Dim tanggal
If WScript.Arguments.Count > 0 Then
    tanggal = WScript.Arguments(0)
Else
    tanggal = Format(Now(), "DD.MM.YYYY")
End If
session.findById("wnd[0]/usr/ctxtI2-LOW").text = "FCH*"
session.findById("wnd[0]/usr/ctxtI8-LOW").text = tanggal
session.findById("wnd[0]/usr/ctxtI8-LOW").setFocus
session.findById("wnd[0]/usr/ctxtI8-LOW").caretPosition = 2
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/sub:SAPLSPO5:0101/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/sub:SAPLSPO5:0101/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/usr/ctxtRLGRAP-FILENAME").text = "C:\Users\CPIFeedmill\Documents\Program Panji\dashboard-compare-sap\data upload\caufv.xls"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
