VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExportarExcel 
   Caption         =   "Exportar a Excel"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6150
   HelpContextID   =   1870
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   3360
      Picture         =   "frmExportarExcel.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   615
      Left            =   1680
      Picture         =   "frmExportarExcel.frx":040D
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdExaminar 
      Height          =   495
      Left            =   5520
      Picture         =   "frmExportarExcel.frx":0691
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   5520
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtNombre 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese el nombre del archivo Excel a generar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmExportarExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TIPO As String

Private Sub cmdAceptar_Click()
    Dim sXLSFile As String
   
    sXLSFile = Trim(txtNombre.Text)
    If UCase(Right(sXLSFile, 4)) <> ".XLS" Then
        sXLSFile = sXLSFile & ".xls"
    End If
    
    Screen.MousePointer = vbHourglass
    Select Case TIPO
        Case Is = "Bloqueados"
            frmMaileador.sgBloqueados.ExportData sXLSFile, sgFormatExcel, sgExportOverwrite + sgExportFieldNames
            ShellExecute frmMaileador.hWnd, "", sXLSFile, "", "", SW_NORMAL
            
        Case Is = "LogEnvio"
            frmMaileador.sgGrilla.ExportData sXLSFile, sgFormatExcel, sgExportOverwrite + sgExportFieldNames
            ShellExecute sgGrilla.hWnd, "", sXLSFile, "", "", SW_NORMAL
            
    End Select
    Screen.MousePointer = vbNormal
   
    Unload Me
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdExaminar_Click()
    
    Dialogo.ShowOpen
    txtNombre.Text = Dialogo.FileName

End Sub

