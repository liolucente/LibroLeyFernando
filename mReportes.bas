Attribute VB_Name = "mReportes"
Option Explicit

Dim mRutaReportes As String

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private iEscala As Integer

Const SETTINGS_PROGID = "biopdf.PDFSettings"
Const UTIL_PROGID = "biopdf.PDFUtil"

Public Property Let RutaDeReportes(ByVal sRuta As String)
    mRutaReportes = sRuta
End Property

Public Property Get RutaDeReportesGet() As String
    RutaDeReportesGet = mRutaReportes
End Property

Public Sub MostrarReporte(ByVal r As CRAXDRT.Report, ByVal rec As Recordset, Optional ByVal sTitulo As String = "NN", Optional ByVal sRuta As String = "NN")
'Este procedimiento muestra un reporte (obvio, verdad?), de Crystal Reports
'r es el Reporte que se va a utilizar
'REC es el recordset de donde sacar los datos
'sTitulo es el titulo de la ventana,
'       si no se pone se utiliza el titulo del reporte.
'sRuta es la ruta donde esta el reporte que se va a usar, es opcional.

Dim sWhere As String
Dim i As Integer


Load r
If Not sRuta = "NN" Then
    r.ReportTemplate = sRuta
End If
'r.ExportOptions.PromptForExportOptions
'r.Export "Desea exportar el reporte a otro formato?"

r.DiscardSavedData
r.VerifyOnEveryPrint = True
r.Database.Verify
r.Database.SetDataSource rec
For i = 1 To r.Database.Tables.Count
    r.Database.Tables(i).SetDataSource rec, 3
Next i


On Error Resume Next
r.ReadRecords
r.Database.Verify

On Error GoTo 0
Load frmREPORTES
With frmREPORTES
    .caption = IIf(sTitulo = "NN", r.ReportTitle, sTitulo)
    .crv.ReportSource = r
    .crv.ViewReport
    .Show
    .WindowState = 2
End With

Set r = Nothing
Set rec = Nothing

End Sub

Public Sub MostrarReporteEX(ByVal r As String, ByVal rec As Recordset, Optional ByVal sTitulo As String = "NN", Optional ByVal sRuta As String = "NN", Optional ByVal sReportTitle As String = "NN", Optional ByVal param1 As String = "NN")
'Este procedimiento muestra un reporte (obvio, verdad?), de Crystal Reports
'r es el Reporte que se va a utilizar
'REC es el recordset de donde sacar los datos
'sTitulo es el titulo de la ventana,
'       si no se pone se utiliza el titulo del reporte.
'sRuta es la ruta donde esta el reporte que se va a usar, es opcional.

Dim i As Integer
Dim crapp As CRAXDRT.Application
Dim rep As CRAXDRT.Report

If sRuta = "NN" Then
    sRuta = mRutaReportes
End If

Set crapp = New CRAXDRT.Application

If sRuta = "NN" Then
    Set rep = crapp.OpenReport(App.Path & "\Reportes\" & r)
Else
    Set rep = crapp.OpenReport(sRuta & "\" & r)
End If


''''rep.ParameterFields(1) = "?cuantos;" & param1 & ";TRUE"
''crptToPrinter.ParameterFields(1) = "Página;" & CStr(nHoja) & "/" & CStr(nHojas) & ";TRUE"

rep.DiscardSavedData
rep.VerifyOnEveryPrint = True
rep.Database.SetDataSource rec, 3
If sReportTitle <> "NN" Then
    rep.ReportTitle = sReportTitle
End If
If param1 <> "NN" Then
    rep.ReportAuthor = param1
End If
For i = 1 To rep.Database.Tables.Count
    rep.Database.Tables(i).SetDataSource rec, 3
Next i


rep.ReadRecords
rep.Database.Verify


On Error GoTo 0
Load frmREPORTES
With frmREPORTES
    .caption = IIf(sTitulo = "NN", rep.ReportTitle, sTitulo)
    .crv.ReportSource = rep
    Load frmConfigReportes
    Set .FormularioDeImpresion = frmConfigReportes
    Set frmConfigReportes.Reporte = rep
    frmConfigReportes.CargarDatos
    .crv.ViewReport
    .WindowState = 2
    Screen.MousePointer = vbNormal
    .Show vbModal

'    Unload frmREPORTES
End With

On Error Resume Next
Set rec.ActiveConnection = Nothing
Set rep = Nothing
'
'If Not crapp.CanClose And Not SERVIDOR = "NULL" Then
'    crapp.LogOffServer CRDLL, SERVIDOR
'End If

Set crapp = Nothing
Set rec = Nothing

End Sub
Public Sub MostrarReportePDF(ByVal r As String, ByVal rec As Recordset, Optional ByVal sTitulo As String = "NN", Optional ByVal sRuta As String = "NN", Optional ByVal sReportTitle As String = "NN", Optional ByVal param1 As String = "NN", Optional ByVal param2 As String = "NN")
'Este procedimiento muestra un reporte (obvio, verdad?), de Crystal Reports
'r es el Reporte que se va a utilizar
'REC es el recordset de donde sacar los datos
'sTitulo es el titulo de la ventana,
'       si no se pone se utiliza el titulo del reporte.
'sRuta es la ruta donde esta el reporte que se va a usar, es opcional.

    Dim i As Integer
    Dim crapp As CRAXDRT.Application
    Dim rep As CRAXDRT.Report

    Dim prtidx As Integer
    Dim sPrinterName As String
    Dim settings As Object
    Dim util As Object

    If sRuta = "NN" Then
        sRuta = mRutaReportes
    End If

    Set crapp = New CRAXDRT.Application

    Set rep = crapp.OpenReport(App.Path & "\Reportes\" & r)

    rep.DiscardSavedData
    rep.VerifyOnEveryPrint = True
    rep.Database.SetDataSource rec, 3
    
    If sReportTitle <> "NN" Then
        rep.ReportTitle = sReportTitle
    End If

    If param1 <> "NN" Then
        rep.ReportAuthor = param1
    End If

    For i = 1 To rep.Database.Tables.Count
        rep.Database.Tables(i).SetDataSource rec, 3
    Next i


    rep.ReadRecords
    rep.Database.Verify

'--------------------------
    
    Set util = CreateObject(UTIL_PROGID)
    sPrinterName = util.defaultprintername
    
        Rem -- Configure the PDF print job
    Set settings = CreateObject(SETTINGS_PROGID)
    settings.printerName = sPrinterName
    settings.setvalue "Output", sRuta & "\" & sTitulo
    settings.setvalue "ConfirmOverwrite", "no"
    settings.setvalue "ShowSaveAS", "never"
    settings.setvalue "ShowSettings", "never"
    settings.setvalue "ShowPDF", "no"
    settings.setvalue "RememberLastFileName", "no"
    settings.setvalue "RememberLastFolderName", "no"
    settings.setvalue "superimpose", App.Path & "\Reportes\Fondo recibo.pdf"
    If Len(Trim(param1)) > 0 Then
        settings.setvalue "OwnerPassword", param1
        settings.setvalue "UserPassword", param2
    End If
    settings.WriteSettings True
    
    Rem -- Find the index of the printer
    prtidx = PrinterIndex(sPrinterName)
    If prtidx < 0 Then Err.Raise 1000, , "No printer was found by the name of '" & sPrinterName & "'."
        
    Rem -- Set the current printer
    Set Printer = Printers(prtidx)
    Printer.Orientation = PrinterObjectConstants.vbPRORPortrait

    Call rep.SelectPrinter(Printer.DriverName, Printer.DeviceName, Printer.Port)
    
    Printer.Zoom = IIf(iEscala = 0, 100, iEscala)
    rep.DiscardSavedData
    rep.ReadRecords

    rep.PrintOut False, CLng(1), True
    
    Set rec.ActiveConnection = Nothing
    Set rep = Nothing

    
'    rep.PrintOut False, CLng(1), True

    Rem -- Wait for runonce settings file to disappear
    Dim runonce As String
    runonce = settings.GetSettingsFilePath(True)
    While Dir(runonce, vbNormal) <> ""
        Sleep 100
    Wend

End Sub


Public Sub MostrarReporte_RutaCompleta(ByVal r As String, ByVal rec As Recordset, Optional ByVal sTitulo As String = "NN")
'Este procedimiento muestra un reporte (obvio, verdad?), de Crystal Reports
'r es el Reporte que se va a utilizar
'REC es el recordset de donde sacar los datos
'sTitulo es el titulo de la ventana,
'       si no se pone se utiliza el titulo del reporte.

Dim i As Integer
Dim crapp As CRAXDRT.Application
Dim rep As CRAXDRT.Report

Set crapp = New CRAXDRT.Application

Set rep = crapp.OpenReport(r)

rep.DiscardSavedData
rep.VerifyOnEveryPrint = True
rep.Database.SetDataSource rec, 3
For i = 1 To rep.Database.Tables.Count
    rep.Database.Tables(i).SetDataSource rec, 3
Next i
rep.ReadRecords
rep.Database.Verify


On Error GoTo 0
Load frmREPORTES
With frmREPORTES
    .caption = IIf(sTitulo = "NN", rep.ReportTitle, sTitulo)
    .crv.ReportSource = rep
    Load frmConfigReportes
    Set .FormularioDeImpresion = frmConfigReportes
    Set frmConfigReportes.Reporte = rep
    frmConfigReportes.CargarDatos
    .crv.ViewReport
    .WindowState = 2
    .Show vbModal
End With

On Error Resume Next
Set rec.ActiveConnection = Nothing
Set rep = Nothing
Set crapp = Nothing
Set rec = Nothing

End Sub

Public Sub MostrarReporteEXX(ByVal r As String, Optional ByVal sTitulo As String = "NN", Optional ByVal sRuta As String = "NN")
'Este procedimiento muestra un reporte (obvio, verdad?), de Crystal Reports
'r es el Reporte que se va a utilizar
'sTitulo es el titulo de la ventana,
'       si no se pone se utiliza el titulo del reporte.
'sRuta es la ruta donde esta el reporte que se va a usar, es opcional.

Dim i As Integer
Dim crapp As CRAXDRT.Application
Dim rep As CRAXDRT.Report

If sRuta = "NN" Then
    sRuta = mRutaReportes
End If

Set crapp = New CRAXDRT.Application

If sRuta = "NN" Then
    Set rep = crapp.OpenReport(App.Path & "\Reportes\" & r)
Else
    Set rep = crapp.OpenReport(sRuta & "\" & r)
End If

rep.DiscardSavedData
rep.VerifyOnEveryPrint = True

On Error GoTo 0
Load frmREPORTES
With frmREPORTES
    .caption = IIf(sTitulo = "NN", rep.ReportTitle, sTitulo)
    .crv.ReportSource = rep
    Load frmConfigReportes
    Set .FormularioDeImpresion = frmConfigReportes
    Set frmConfigReportes.Reporte = rep
    frmConfigReportes.CargarDatos
    .crv.ViewReport
    .WindowState = 2
    .Show vbModal
End With

On Error Resume Next
'Set rec.ActiveConnection = Nothing
Set rep = Nothing

If Not crapp.CanClose And Not SERVIDOR = "NULL" Then
    crapp.LogOffServer CRDLL, SERVIDOR
End If

Set crapp = Nothing
'Set rec = Nothing

End Sub
Private Function PrinterIndex(ByVal printerName As String) As Integer
    Dim i As Integer
    
    For i = 0 To Printers.Count - 1
        If LCase(Printers(i).DeviceName) Like LCase(printerName) Then
            PrinterIndex = i
            Exit Function
        End If
    Next
    PrinterIndex = -1
End Function


Public Sub ImprimirReportePDF(ByVal r As String, ByVal rec As Recordset, Optional ByVal sTitulo As String = "NN", Optional ByVal sRuta As String = "NN", Optional ByVal sReportTitle As String = "NN", Optional ByVal param1 As String = "NN", Optional ByVal param2 As String = "NN")
'Este procedimiento muestra un reporte (obvio, verdad?), de Crystal Reports
'r es el Reporte que se va a utilizar
'REC es el recordset de donde sacar los datos
'sTitulo es el titulo de la ventana,
'       si no se pone se utiliza el titulo del reporte.
'sRuta es la ruta donde esta el reporte que se va a usar, es opcional.

    Dim i As Integer
    Dim crapp As CRAXDRT.Application
    Dim rep As CRAXDRT.Report

    Dim prtidx As Integer
    Dim sPrinterName As String
    Dim settings As Object
    Dim util As Object

    If sRuta = "NN" Then
        sRuta = mRutaReportes
    End If

    Set crapp = New CRAXDRT.Application

    Set rep = crapp.OpenReport(App.Path & "\Reportes\" & r)

    rep.DiscardSavedData
    rep.VerifyOnEveryPrint = True
    rep.Database.SetDataSource rec, 3
    
    If sReportTitle <> "NN" Then
        rep.ReportTitle = sReportTitle
    End If

    If param1 <> "NN" Then
        rep.ReportAuthor = param1
    End If

    For i = 1 To rep.Database.Tables.Count
        rep.Database.Tables(i).SetDataSource rec, 3
    Next i


    rep.ReadRecords
    rep.Database.Verify

'--------------------------
    
    Set util = CreateObject(UTIL_PROGID)
    sPrinterName = util.defaultprintername
    
        Rem -- Configure the PDF print job
    Set settings = CreateObject(SETTINGS_PROGID)
    settings.printerName = sPrinterName
    settings.setvalue "Output", "c:\temp\" & sTitulo
    settings.setvalue "ConfirmOverwrite", "no"
    settings.setvalue "ShowSaveAS", "never"
    settings.setvalue "ShowSettings", "never"
    settings.setvalue "ShowPDF", "no"
    settings.setvalue "RememberLastFileName", "no"
    settings.setvalue "RememberLastFolderName", "no"
    settings.setvalue "superimpose", App.Path & "\Reportes\Fondo recibo.pdf"
    
    settings.setvalue "AfterPrintProgram", "c:\temp\imprimirPDF.bat"
    
    settings.WriteSettings True
    
    Rem -- Find the index of the printer
    prtidx = PrinterIndex(sPrinterName)
    If prtidx < 0 Then Err.Raise 1000, , "No printer was found by the name of '" & sPrinterName & "'."
        
    Rem -- Set the current printer
    Set Printer = Printers(prtidx)
    Printer.Orientation = PrinterObjectConstants.vbPRORPortrait

    Call rep.SelectPrinter(Printer.DriverName, Printer.DeviceName, Printer.Port)
    
    Printer.Zoom = IIf(iEscala = 0, 100, iEscala)
    rep.DiscardSavedData
    rep.ReadRecords

    rep.PrintOut False, CLng(1), True
    
    Set rec.ActiveConnection = Nothing
    Set rep = Nothing

    
'    rep.PrintOut False, CLng(1), True

    Rem -- Wait for runonce settings file to disappear
    Dim runonce As String
    runonce = settings.GetSettingsFilePath(True)
    While Dir(runonce, vbNormal) <> ""
        Sleep 100
    Wend
    Sleep 30000
    
End Sub

