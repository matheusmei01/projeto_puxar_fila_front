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
session.findById("wnd[0]/usr/subSC_SUB_3:Y02VSI_FAX_MONITOR:0300/tabsSC3_TABSTRIP/tabpSC3_TAB_6/ssubSC3_SUB1:Y02VSI_FAX_MONITOR:0310/cntlSC_HTML/shellcont/shell").sapEvent "","","sapevent:S_?SUBM=Busca"
session.findById("wnd[1]/tbar[0]/btn[7]").press
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_ANG").select
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_OPEN").select
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "2201"
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[0]").press
