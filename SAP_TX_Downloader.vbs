Sub sap_Report_Excel_Download()

    '
    ' Example of Macro that Download a Report from a Transaccion of SAP
    '


    ' Date manipulation
    ano = DatePart("YYYY", Date)

    mes_texto = MonthName(DatePart("M", Date), False)
    mes_num = DatePart("M", Date)
    myDateText = Format(Date - 1, "ddmm")
    mytime = CDate(Round(CDate(Time) * 1440 / 10, 0) * 10 / 1440)
    myHourText = Format(mytime, "hhmm")
    fname = "RF AGRUP " & myDateText & " 3.xls"
    pathname = "c:\Trabalho\Transporte\" & ano & "\0.-EF Diaria\" & mes_num & ".-" & mes_texto


    'Variables

    Dim usuario
    Dim pass
    Dim ambiente

    Windows("Menu - Estatus de Carga.xlsm").Activate

    Sheets("Menu").Select
    usuario = Range("g1").Value
    pass = Range("g2").Value
    ambiente = Range("i2").Value                        'Ambience of your SAP Server 

    origem_dados = "Eficiencia Diaria.xlsx"             'Source

    'Copy of data to use on SAP as filter
    Windows(origem_dados).Activate
    Sheets("BASE").Select
    Set sht = ActiveSheet
    sht.Select
    lastrow = sht.Cells(sht.Rows.Count, "c").End(xlUp).Row
    Range("C2:C" & lastrow).Select
    Selection.Copy

    'Creation of SAP instance

    Application.DisplayAlerts = True

    Set app = CreateObject("Sapgui.ScriptingCtrl.1")
    Set Connection = app.OpenConnection(ambiente, True)
    Set session = Connection.Children(0)


    ' This was used to connect at SAP with your account, for Single LogOn isn't necessary
    session.findById("wnd[0]").maximize
    'session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = "010"
    'session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = usuario
    'session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = pass
    'session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = "ES"
    'session.findById("wnd[0]/usr/pwdRSYST-BCODE").SetFocus
    'session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 2
    'session.findById("wnd[0]").sendVKey 0


    If session.Children.Count > 1 Then

        session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").Select
        session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").SetFocus
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    End If


    On Error Resume Next
    session.findById("wnd[0]").sendVKey 0
    On Error Resume Next


    'Selection of TX
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "Z0SWT30001TRN"  'TX of Example

    'Selection of Variant

    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 17          
    session.findById("wnd[1]").sendVKey 8
    session.findById("wnd[0]").sendVKey 8


    'Filters
    mes_ant_2 = DateSerial(Year(Date), Month(Date) - 2, 1)
    mes_ant_2 = Format(mes_ant_2, "dd.mm.yyyy")


    fim_mes_atual = DateSerial(Year(Date), Month(Date) + 1, 0)
    fim_mes_atual = Format(fim_mes_atual, "dd.mm.yyyy")



    session.findById("wnd[0]/usr/btnDISPLAY").press
    session.findById("wnd[0]/usr/subG_SUBSCREEN:SAPMZ0SWT30002:9005/ctxtS_WERKA-LOW").Text = "*"
    session.findById("wnd[0]/usr/subG_SUBSCREEN:SAPMZ0SWT30002:9005/ctxtS_ERDAT-LOW").Text = mes_ant_2
    session.findById("wnd[0]/usr/subG_SUBSCREEN:SAPMZ0SWT30002:9005/ctxtS_ERDAT-HIGH").Text = fim_mes_atual
    session.findById("wnd[0]/usr/subG_SUBSCREEN:SAPMZ0SWT30002:9005/ctxtS_ERDAT-HIGH").SetFocus
    session.findById("wnd[0]/usr/subG_SUBSCREEN:SAPMZ0SWT30002:9005/ctxtS_ERDAT-HIGH").caretPosition = 10
    session.findById("wnd[0]/usr/subG_SUBSCREEN:SAPMZ0SWT30002:9005/btn%_S_TKNUM_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/tbar[0]/btn[24]").press

    'Use of Data from Excel to filter

    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press


    session.findById("wnd[0]/usr/cntlCUSTOM_CONTROL_100/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"

    'gravar com com nome RF AGRUP 'ddmm' 2.xlsx

    '    Downloading the Report
    session.findById("wnd[0]/usr/cntlCUSTOM_CONTROL_100/shellcont/shell").selectContextMenuItem "&PC"
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = pathname & "\"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = fname
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
    session.findById("wnd[1]/tbar[0]/btn[11]").press

    On Error Resume Next
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press


    Windows(origem_dados).Activate
        

End Sub
