VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmConfigReportes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuracion de la impresora"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3450
   Icon            =   "frmConfigReportes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   3450
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPDF 
      Caption         =   "PDF"
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Width           =   495
   End
   Begin MSComDlg.CommonDialog Archivo 
      Left            =   2820
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   6148
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   4200
      Top             =   2220
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfigReportes.frx":5C12
            Key             =   "IMPRESORA"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfigReportes.frx":B836
            Key             =   "HTML"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfigReportes.frx":11AD2
            Key             =   "OFFICE"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfigReportes.frx":11DEE
            Key             =   "HOJAS"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione la impresora"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3435
      Begin MSComctlLib.ImageCombo cboImpresoras 
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         ImageList       =   "IMG"
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Desde - Hasta"
      Height          =   975
      Left            =   1680
      TabIndex        =   11
      Top             =   1980
      Width           =   1755
      Begin VB.TextBox txtPagFinal 
         Height          =   285
         Left            =   1140
         TabIndex        =   15
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtPagInicio 
         Height          =   285
         Left            =   1140
         TabIndex        =   14
         Text            =   "1"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Pagina Final"
         Height          =   195
         Left            =   60
         TabIndex        =   13
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Pagina inicial"
         Height          =   195
         Left            =   60
         TabIndex        =   12
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cantidad de copias"
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   1980
      Width           =   1635
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   870
         TabIndex        =   8
         Top             =   360
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtCantidad"
         BuddyDispid     =   196617
         OrigLeft        =   900
         OrigTop         =   300
         OrigRight       =   1125
         OrigBottom      =   615
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtCantidad 
         Height          =   285
         Left            =   480
         TabIndex        =   7
         Text            =   "1"
         Top             =   360
         Width           =   390
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   735
      Left            =   720
      Picture         =   "frmConfigReportes.frx":16F26
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3060
      Width           =   915
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   735
      Left            =   1680
      MaskColor       =   &H00FF00FF&
      Picture         =   "frmConfigReportes.frx":1CB38
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3060
      UseMaskColor    =   -1  'True
      Width           =   915
   End
   Begin VB.Frame Frame5 
      Caption         =   "Tipo de hoja"
      Height          =   615
      Left            =   0
      TabIndex        =   16
      Top             =   720
      Width           =   3435
      Begin MSComctlLib.ImageCombo cboHoja 
         Height          =   330
         Left            =   180
         TabIndex        =   17
         Top             =   240
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Orientacion"
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   3435
      Begin VB.OptionButton optDefault 
         Caption         =   "Definida"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2460
         TabIndex        =   6
         Top             =   300
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optHorizontal 
         Caption         =   "Horizontal"
         Height          =   195
         Left            =   1260
         TabIndex        =   5
         Top             =   300
         Width           =   1095
      End
      Begin VB.OptionButton optVertical 
         Caption         =   "Vertical"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmConfigReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Reporte As Report
Private iEscala As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const SETTINGS_PROGID = "biopdf.PDFSettings"
Const UTIL_PROGID = "biopdf.PDFUtil"

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

Public Property Let Escala(ByVal vEscala As Integer)
    iEscala = vEscala
End Property

Private Sub cmdImprimir_Click()
Dim PRT As Printer

If Left(cboImpresoras.SelectedItem.key, 1) = "A" Then
    
    Reporte.ExportOptions.DestinationType = crEDTDiskFile
    Reporte.ExportOptions.NumberOfLinesPerPage = 0
    
    Select Case cboImpresoras.SelectedItem.key
    Case Is = "AXLS"
        Archivo.DialogTitle = "Seleccione el nombre del archivo..."
        Archivo.filter = "Archivos de Excel (*.xls)|*.xls"
        Archivo.ShowSave
        If Archivo.FileName = "" Then Exit Sub
        Reporte.ExportOptions.DiskFileName = Archivo.FileName
        If Right(Reporte.ExportOptions.DiskFileName, 3) <> "xls" And Reporte.ExportOptions.DiskFileName = "" Then
            Reporte.ExportOptions.DiskFileName = App.Path & "Exportacion a Excel.xls"
        Else
            Reporte.ExportOptions.DiskFileName = Reporte.ExportOptions.DiskFileName & ".xls"
        End If
        Reporte.ExportOptions.FormatType = crEFTExcel80Tabular
        
        On Error Resume Next
        Reporte.Export False
   Case Is = "AHTML"
        Archivo.DialogTitle = "Seleccione el nombre del archivo..."
        Archivo.filter = "Paginas HTML (*.htm,*.html)|*.htm;*.html"
        Archivo.ShowSave
        If Archivo.FileName = "" Then Exit Sub
        Reporte.ExportOptions.DiskFileName = Archivo.FileName
        Reporte.ExportOptions.FormatType = crEFTHTML32Standard
        If Not Right(Reporte.ExportOptions.DiskFileName, 3) = "html" Then
            Reporte.ExportOptions.DiskFileName = Reporte.ExportOptions.DiskFileName & ".html"
        End If
        On Error Resume Next
        Reporte.Export False
    Case Is = "ADOC"
        Archivo.DialogTitle = "Seleccione el nombre del archivo..."
        Archivo.filter = "Documentos de Word (*.doc)|*.doc"
        Archivo.ShowSave
        If Archivo.FileName = "" Then Exit Sub
        Reporte.ExportOptions.DiskFileName = Archivo.FileName
        If Not Right(Reporte.ExportOptions.DiskFileName, 4) = ".doc" Then
            Reporte.ExportOptions.DiskFileName = Reporte.ExportOptions.DiskFileName & ".doc"
        End If
        Reporte.ExportOptions.FormatType = crEFTRichText
        On Error Resume Next
        Reporte.Export False
    End Select

Else

    Call Reporte.SelectPrinter(Left(cboImpresoras.SelectedItem.key, Len(cboImpresoras.SelectedItem.key) - 1), cboImpresoras.SelectedItem.Text, cboImpresoras.Tag)
    
'    For Each PRT In Printers
'        If PRT.DeviceName = cboImpresoras.SelectedItem.Text Then
'            PRT.Zoom = iEscala
'        End If
'    Next
    Printer.Zoom = IIf(iEscala = 0, 100, iEscala)
 
    If optHorizontal = True Then
        Reporte.PaperOrientation = crLandscape
    ElseIf optVertical = True Then
        Reporte.PaperOrientation = crPortrait
    End If
    
    Select Case cboHoja.SelectedItem.key
    Case "A4"
        Reporte.PaperSize = crPaperA4
    Case "OFICIO"
        Reporte.PaperSize = crPaperLegal
    Case "CARTA"
        Reporte.PaperSize = crPaperLetter
    Case Else
        Reporte.PaperSize = crPaperA4
    End Select
    Reporte.DiscardSavedData
    Reporte.ReadRecords
    
    If Not txtPagFinal = "" And Not txtPagInicio = "" Then
        Reporte.PrintOut False, IIf(txtCantidad.Text = "", CLng(1), CLng(txtCantidad.Text)), , CLng(txtPagInicio.Text), CLng(txtPagFinal.Text)
    Else
        Reporte.PrintOut False, IIf(txtCantidad.Text = "", CLng(1), CLng(txtCantidad.Text))
    End If
End If

Me.Hide
End Sub

Private Sub cmdSalir_Click()
Me.Hide
End Sub

Public Sub CargarDatos()
On Error Resume Next

    Select Case Reporte.PaperOrientation
    Case crLandscape
        optHorizontal.value = True
    Case crPortrait
        optVertical.value = True
    Case crDefaultPaperOrientation
        optDefault.value = True
    End Select

Dim mPrinter As Printer

For Each mPrinter In Printers
    cboImpresoras.ComboItems.Add , mPrinter.DriverName & CStr(cboImpresoras.ComboItems.Count), mPrinter.DeviceName, "IMPRESORA"
    cboImpresoras.ComboItems(cboImpresoras.ComboItems.Count).Tag = mPrinter.Port
Next
cboImpresoras.ComboItems.Add , "AXLS", "Exportar a Excel"
cboImpresoras.ComboItems.Add , "ADOC", "Exportar a Word "
cboImpresoras.ComboItems.Add , "AHTML", "Exportar a HTML"
cboImpresoras.ComboItems(1).Selected = True

Set cboHoja.ImageList = IMG
cboHoja.ComboItems.Add , "A4", "Hoja ""A4""", "HOJAS"
cboHoja.ComboItems.Add , "CARTA", "Hoja ""Carta""", "HOJAS"
cboHoja.ComboItems.Add , "OFICIO", "Hoja ""Oficio""", "HOJAS"
If Reporte.PaperSize = crPaperLegal Then
    cboHoja.ComboItems(3).Selected = True
Else
    cboHoja.ComboItems(1).Selected = True
End If
End Sub

Private Sub ComPDF_Click()

    Dim prtidx As Integer
    Dim sPrinterName As String
    Dim settings As Object
    Dim util As Object
    
    Set util = CreateObject(UTIL_PROGID)
    sPrinterName = util.defaultprintername
    
    Rem -- Configure the PDF print job
    Set settings = CreateObject(SETTINGS_PROGID)
    settings.printerName = sPrinterName
    settings.setvalue "Output", "<desktop>\myfile.pdf"
    settings.setvalue "ConfirmOverwrite", "no"
    settings.setvalue "ShowSaveAS", "never"
    settings.setvalue "ShowSettings", "never"
    settings.setvalue "ShowPDF", "yes"
    settings.setvalue "RememberLastFileName", "no"
    settings.setvalue "RememberLastFolderName", "no"
    settings.setvalue "superimpose", "N:\Usuarios\Hospital\Recibos sueldo\Fondo recibo.pdf"
    settings.WriteSettings True
    
    Rem -- Find the index of the printer
    prtidx = PrinterIndex(sPrinterName)
    If prtidx < 0 Then Err.Raise 1000, , "No printer was found by the name of '" & sPrinterName & "'."
        
    Rem -- Set the current printer
    Set Printer = Printers(prtidx)
    Printer.Orientation = PrinterObjectConstants.vbPRORPortrait

    Call Reporte.SelectPrinter(Printer.DriverName, Printer.DeviceName, Printer.Port)
    
    Printer.Zoom = IIf(iEscala = 0, 100, iEscala)
    Reporte.DiscardSavedData
    Reporte.ReadRecords
    
    If Not txtPagFinal = "" And Not txtPagInicio = "" Then
        Reporte.PrintOut False, IIf(txtCantidad.Text = "", CLng(1), CLng(txtCantidad.Text)), , CLng(txtPagInicio.Text), CLng(txtPagFinal.Text)
    Else
        Reporte.PrintOut False, IIf(txtCantidad.Text = "", CLng(1), CLng(txtCantidad.Text))
    End If

    Rem -- Wait for runonce settings file to disappear
    Dim runonce As String
    runonce = settings.GetSettingsFilePath(True)
    While Dir(runonce, vbNormal) <> ""
        Sleep 100
    Wend

    Me.Hide
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmConfigReportes.Reporte = Nothing
End Sub

Private Sub txtCantidad_Validate(Cancel As Boolean)
DejarNumeros txtCantidad
If Not IsNumeric(txtCantidad.Text) Then
    MsgBox "Por favor ingrese un numero.", vbOKOnly + vbExclamation, "Cantidad..."
    Cancel = True
End If
End Sub
