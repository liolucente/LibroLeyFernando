VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Verificando los datos del Sistema..."
   ClientHeight    =   4155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7110
   HelpContextID   =   240
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4155
      Left            =   45
      TabIndex        =   0
      Top             =   -30
      Width           =   7035
      Begin VB.PictureBox PicFirma 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   375
         Picture         =   "frmSplash.frx":0000
         ScaleHeight     =   555
         ScaleWidth      =   1695
         TabIndex        =   8
         Top             =   2820
         Width           =   1695
      End
      Begin VB.PictureBox PicGali 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2280
         Left            =   375
         Picture         =   "frmSplash.frx":3166
         ScaleHeight     =   2280
         ScaleWidth      =   1605
         TabIndex        =   7
         Top             =   360
         Width           =   1605
      End
      Begin VB.Timer tmrVerifica 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   2430
         Top             =   240
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         Caption         =   "Version: 0.00.0001"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3720
         TabIndex        =   3
         Top             =   3465
         Width           =   3165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Galileo Sistemas S.R.L."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3090
         TabIndex        =   5
         Top             =   3750
         Width           =   3840
      End
      Begin VB.Label lblEstado 
         Caption         =   "Verificando el registro del sistema..."
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   3825
         Width           =   3930
      End
      Begin VB.Label Label2 
         Caption         =   "Galileo XXI"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   2925
         TabIndex        =   4
         Top             =   1080
         Width           =   3915
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Plataformas: Windows 9x, Me, NT, 2000, XP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2550
         TabIndex        =   2
         Top             =   2820
         Width           =   4335
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Uso autorizado "
         Height          =   240
         Left            =   2115
         TabIndex        =   1
         Top             =   135
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3D3873E102C6"
 Option Explicit

'***********************************************'
'                                               '
'  Pantalla de inicio del sistema, que verifica '
'  la consistencia de los datos del registro    '
'  y si es posible correr el sistema            '
'                                               '
'  Hecho por: Fernando (Chicho)                 '
'  Fecha de creacion: 03/12/2001                '
'                                               '
'***********************************************'

Const sTituloNoInicia As String = "No se puede iniciar el sistema en esta terminal."
Const sCopyright As String = "Galileo Sistemas XXI"
Dim sVersion As String
Dim sFechaUltimo As String
Dim sFechaTope As String

Dim adoCon As New ADODB.Connection
Dim adoRec As New ADODB.Recordset

Private Sub Form_Load()

tmrVerifica.Enabled = True
sVersion = Format(App.Major, "0") & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "0000")
lblVersion.caption = "Version: " & sVersion
lblVersion.Refresh
Me.Refresh
' PonerVentanaPrimero Me.hwnd
End Sub

Private Sub tmrVerifica_Timer()

'***********************************************'
'                                               '
'  Luego de 1 segundo se ejecuta este proc.     '
'  que verifica si los datos del registro       '
'  son validos. Los datos que verifica son:     '
'  el Copyright, el Nombre del Sistema Instalado'
'  la version (que la saco del EXE con el       '
'  formato 0.00.0000) el Tope de Usuarios       '
'  Ultimo Acceso y Tope Vigencia                '
'                                               '
'  Hecho por: Fernando (Chicho)                 '
'  Fecha de creacion: 03/12/2001                '
'                                               '
'***********************************************'

tmrVerifica.Enabled = False

Dim sMensaje As String
Dim ini As sgINI
Dim sUltimoAcceso As String
Dim fs As FileSystemObject

Set fs = New FileSystemObject

lblEstado = "Verificando la configuracion de los datos..."

If Not fs.FileExists(App.Path & "\config.ini") Then

    MsgBox "No existe el archivo de configuracion de datos." & vbCrLf, "Error en la configuracion de datos..."
    Set fs = Nothing
    End

Else

    Set ini = New sgINI
    ini.FullPath = App.Path & "\config.ini"
    ini.SectionName = "CONFIGURACION"

    SGINST = ini.GetString("BASESGI", SGINST)
    On Error GoTo ErrorConexion

    Set adoCon = New Connection
    adoCon.Open SGINST
    adoCon.Close
    Set adoCon = Nothing

    SGEPOB = ini.GetString("BASEPOB", SGEPOB)
    SGINTE = ini.GetString("INTERCAMBIO", SGINTE)
    PDFS_DEST = ini.GetString("PDFDEST", PDFS_DEST)
    PDFS_ORIG = ini.GetString("PDFORIG", PDFS_ORIG)
    FOX = ini.GetString("DIRFOX", FOX)
    MI_EMAIL = ini.GetString("MAIL", MI_EMAIL)
    
    ASUNTO1 = ini.GetString("ASUNTO", ASUNTO1)
    LINEA1 = ini.GetString("LIN01", LINEA1)
    LINEA2 = ini.GetString("LIN02", LINEA2)
    LINEA3 = ini.GetString("LIN03", LINEA3)
    LINEA4 = ini.GetString("LIN04", LINEA4)
    LINEA5 = ini.GetString("LIN05", LINEA5)
    LINEA6 = ini.GetString("LIN06", LINEA6)
    LINEA7 = ini.GetString("LIN07", LINEA7)
    LINEA8 = ini.GetString("LIN08", LINEA8)
    LINEA9 = ini.GetString("LIN09", LINEA9)
    LINEA10 = ini.GetString("LIN10", LINEA10)

    CRDLL = ini.GetString("CRDLL", "P2lsql.dll")

End If
Set ini = Nothing
Set fs = Nothing

sVersion = Format(App.Major, "0") & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "0000")
lblVersion.caption = "Version: " & sVersion
lblVersion.Refresh
Me.Refresh

Screen.MousePointer = vbHourglass

Unload Me
'frmFondo.Show

Screen.MousePointer = vbNormal
frmInicioSesion.Show vbModal
Screen.MousePointer = vbNormal

Exit Sub

ErrorConexion:
    sMensaje = "Ha ocurrido un error al intentar "
    sMensaje = sMensaje & "conectarse a la base de instalacion"
    sMensaje = sMensaje & vbCrLf & "No podra continuar cargando el sistema."
    sMensaje = sMensaje & vbCrLf & "Por favor contactese con Galileo Sistemas S.R.L."
    MsgBox sMensaje, vbOKOnly + vbCritical, sTituloNoInicia
    GuardarLOGError_SU Err, App.Path & "\ErrInicio.txt"
    
    If adoCon.State = adStateOpen Then adoCon.Close
    Set adoCon = Nothing
    If adoRec.State = adStateOpen Then adoRec.Close
    Set adoRec = Nothing
    End

Screen.MousePointer = vbNormal
frmInicioSesion.Show vbModal
Screen.MousePointer = vbNormal

End Sub
