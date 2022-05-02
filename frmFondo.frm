VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.MDIForm frmFondo 
   BackColor       =   &H8000000C&
   Caption         =   "Galileo Sistemas XXI"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7500
   Icon            =   "frmFondo.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar Estado 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   5370
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuRecibos 
      Caption         =   "&Recibos digitales"
      Begin VB.Menu mnuActualizarDatos 
         Caption         =   "Actualizar datos"
      End
   End
   Begin VB.Menu mnuLibroLey 
      Caption         =   "&Libro ley digital"
      Begin VB.Menu mnuGenerarLibro 
         Caption         =   "Generar Libro ley"
      End
   End
   Begin VB.Menu mnuSAlir 
      Caption         =   "&Salir"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuCambiarUsuario 
         Caption         =   "Cambiar &usuario"
      End
      Begin VB.Menu mnuCambiarEmpresa 
         Caption         =   "Cambiar de &empresa"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReiniciar 
         Caption         =   "&Reiniciar sesion"
      End
      Begin VB.Menu mnuCerrarSesion 
         Caption         =   "&Cerrar sesion"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalirSistema 
         Caption         =   "Cerrar &sistema"
      End
   End
End
Attribute VB_Name = "frmFondo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3D3873E90122"
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
    ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" ( _
   ByVal hLibModule As Long) As Long
   
Private m_hMod As Long

Private Sub Barra_ItemClick(Item As vbalLbar6.cListBarItem, Bar As vbalLbar6.cListBar)
'On Error GoTo GENERICO_ERRORES
'
'CargarSistema Item.key
'
'Me.AutoShowChildren = True
'
'If Item.key = "CONFIG" Then
'    On Error Resume Next
'    frmCONFIG.Show
'ElseIf Item.key = "PADRONES" Then
'    'On Error Resume Next
'    'IniciarTratamiento
'    'frmABMGenerico.Show
'    frmPadrones.Show
'ElseIf Item.key = "OSPIN" Then
'    frmObraSocial.Show
'ElseIf Item.key = "SUTNA" Then
'    frmSindicato.Show
'End If
'Exit Sub
'
'GENERICO_ERRORES:
'    MsgBox "Error inesperado.", vbCritical, "Errores"
'
End Sub

Private Sub MDIForm_Initialize()
    m_hMod = LoadLibrary("shell32.dll")
    InitCommonControls
    
End Sub

Private Sub MDIForm_Load()
Me.caption = sNombreSistema

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If CerrarSistema = False Then
'    Cancel = CInt(True)
'End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    FreeLibrary m_hMod
End Sub

Private Sub mnuActualizarDatos_Click()
    frmActualizador.Show
End Sub

Private Sub mnuCambiarEmpresa_Click()
'Me.Enabled = False
'On Error GoTo ErrorES
'frmCambioEmpresa.Show
'
'ErrorES:
'    Exit Sub
End Sub

Private Sub mnuCambiarUsuario_Click()
'frmCambioSesion.Show vbModal
End Sub

Private Sub mnuCerrarSesion_Click()
ReiniciarSistema True
End Sub

Private Sub mnuGenerarLibro_Click()
    frmLibroLeyDigital.Show
End Sub

Private Sub mnuMaileador_Click()
'    IniciarMaileadorRecibos
End Sub

Private Sub mnuReiniciar_Click()
    ReiniciarSistema False
End Sub

Private Sub mnuSalirSistema_Click()
    Unload Me
    End
End Sub
