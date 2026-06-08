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
session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "aufm"
session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 2
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "2"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/btn%_I11_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[23]").press
session.findById("wnd[2]/tbar[0]/btn[37]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[45]").press
On Error Resume Next
Dim chkExport : Set chkExport = session.findById("wnd[1]/usr/sub:SAPLSPO5:0101/radSPOPLI-SELFLAG[1,0]")
Dim exportOk  : exportOk     = (Err.Number = 0)
Err.Clear : On Error GoTo 0
If Not exportOk Then
    On Error Resume Next : session.findById("wnd[1]/tbar[0]/btn[0]").press : Err.Clear : On Error GoTo 0
    WScript.Quit 3
End If
session.findById("wnd[1]/usr/sub:SAPLSPO5:0101/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/sub:SAPLSPO5:0101/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/usr/ctxtRLGRAP-FILENAME").text = "C:\Users\CPIFeedmill\Documents\Program Panji\dashboard-compare-sap\data upload\aufm.xls"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press