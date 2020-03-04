# EMBRAER
# Picking FLOW EAF

Sub PICKING YABORA()



If Not IsObject(SAPApp) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set SAPApp = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
   Set Connection = SAPApp.Children(0)
End If
If Not IsObject(session) Then
   Set session = Connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject SAPApp, "on"
End If

Dim v1, v2, v3 As String

v1 = Now()
v2 = Format(v1, "dd.mm.yyyy")
v3 = v1 - 30
v4 = Format(v3, "dd.mm.yyyy")



session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZLORSD015"
session.findById("wnd[0]").sendVKey 0
'session.findById("wnd[0]").maximize

session.findById("wnd[0]/usr/chkP_LOCAL").Selected = True
session.findById("wnd[0]/usr/ctxtS_VKORG-LOW").Text = "lbg1"
session.findById("wnd[0]/usr/ctxtS_LFART-LOW").Text = "*"
session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = "LBG*"





'session.findById("wnd[0]/usr/ctxtS_DATA-LOW").Text = "111111111"

session.findById("wnd[0]/usr/ctxtS_DATA-HIGH").Text = v2

session.findById("wnd[0]/usr/ctxtS_DATA-LOW").Text = v4




session.findById("wnd[0]/usr/chkP_LOCAL").SetFocus
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[6]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, "PROCESSO"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "PROCESSO"
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "V1"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "V2"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "V3"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").SetFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").caretPosition = 2
session.findById("wnd[2]/tbar[0]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, "I_MIN_P"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "I_MIN_P"
session.findById("wnd[0]/tbar[1]/btn[40]").press
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "S:\EAI\03- Picking\EAI PICKING\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "EAI PICKING.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 11
session.findById("wnd[1]/tbar[0]/btn[11]").press

End Sub
