VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmInicioSesion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingreso al Sistema"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4980
   HelpContextID   =   390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameEmpresas 
      Caption         =   "Empresas"
      Height          =   2355
      Left            =   0
      TabIndex        =   11
      Top             =   2400
      Width           =   4905
      Begin VB.CommandButton cmdEntrar 
         Caption         =   "Entrar al sistema"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1778
         TabIndex        =   14
         Top             =   1845
         Width           =   1680
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3548
         TabIndex        =   13
         Top             =   1845
         Width           =   960
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar ventana"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   293
         TabIndex        =   12
         Top             =   1845
         Width           =   1410
      End
      Begin MSComctlLib.ListView lvwEmpresas 
         Height          =   1095
         Left            =   90
         TabIndex        =   15
         Top             =   225
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   1931
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nombre de la Empresa"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descripcion"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Conexion"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "EMPRCODIGO"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblDesc 
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   4515
      End
   End
   Begin VB.Frame frameUsuario 
      Caption         =   "Iniciar sesión en el sistema"
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton cmdIniciar 
         Caption         =   "Iniciar sesion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1725
         TabIndex        =   9
         Top             =   1500
         Width           =   1545
      End
      Begin VB.TextBox txtClave 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   2655
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   8
         Text            =   "0123456789"
         Top             =   1095
         Width           =   1635
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2655
         MaxLength       =   10
         TabIndex        =   7
         Top             =   645
         Width           =   1635
      End
      Begin VB.TextBox txtNombre 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   1635
      End
      Begin VB.PictureBox picOlvidoClave 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   4440
         Picture         =   "frmInicioSesion.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   240
      End
      Begin VB.Timer tmrImagen 
         Interval        =   1000
         Left            =   120
         Top             =   1920
      End
      Begin VB.PictureBox picLlave2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   4350
         Picture         =   "frmInicioSesion.frx":2072
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   4
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Chequee las mayusculas y minusculas, ya que el sistema es sensible a estas."
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1905
         Width           =   4560
      End
      Begin VB.Label lblClave 
         Caption         =   "Clave:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   3
         Top             =   1170
         Width           =   1725
      End
      Begin VB.Label lblNombre 
         Caption         =   "Nombre de usuario:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   2
         Top             =   360
         Width           =   1725
      End
      Begin VB.Label lblCodigo 
         Caption         =   "Codigo de Usuario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   1
         Top             =   765
         Width           =   1725
      End
   End
End
Attribute VB_Name = "frmInicioSesion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3D387415037A"
Option Explicit

Dim lUSUA_ID As Long
Dim adoClaveUsuario As Recordset
Dim iErrores As Integer

Dim bSinEmpresa As Boolean
Dim tmpsgUser As New sgUsuario
Const m_def_RutaApp = ""
Dim m_RutaApp As String

Private Sub Form_Load()
    
    txtNombre.Text = ""
    txtCodigo.Text = ""
    txtClave.Text = ""
    
    m_RutaApp = App.Path
    
End Sub

'##ModelId=3D387416006E
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        Me.Hide
        frmFondo.mnuCerrarSesion.Enabled = False
        frmFondo.mnuReiniciar.Enabled = False
        frmFondo.mnuCambiarUsuario.caption = "&Iniciar sesion"
        Me.Show
    End If

End Sub

Public Function VerificarUsuario(ByVal sCodigo As String) As Boolean

'***********************************************'
'                                               '
'  Funcion para verificar la existencia del     '
'  usuario en el sistema.                       '
'                                               '
'  Hecho por: Fernando (Chicho)                 '
'  Fecha de creacion: 04/12/2001                '
'                                               '
'***********************************************'

Dim adoCon As New ADODB.Connection
Dim adoRec As New Recordset
Dim sMensaje As String

On Error GoTo ErrorConexion
adoCon.CursorLocation = adUseClient
adoCon.Open SGINST

On Error GoTo ErrorRecordSet
adoRec.ActiveConnection = adoCon
adoRec.CursorLocation = adUseClient
adoRec.CursorType = adOpenStatic
adoRec.LockType = adLockReadOnly
adoRec.source = "Select * From SGUSUARIO Where USUASENIAL = '" & Encripta(CInt(True)) & "'"
adoRec.source = adoRec.source & " AND USUACODIGO = '" & sCodigo & "'"
adoRec.Open , , , , adCmdText

adoRec.ActiveConnection = Nothing
adoCon.Close
Set adoCon = Nothing

If adoRec.RecordCount = 0 Then
    sMensaje = "El usuario ingresado no existe en el sistema o no esta habilitado."
    sMensaje = sMensaje & vbCrLf & "Por favor comuniquese con su supervisor"
    MsgBox sMensaje, vbOKOnly + vbExclamation, "Usuario: " & txtCodigo.Text & ", incorrecto."
    VerificarUsuario = False
Else
    VerificarUsuario = True
    lUSUA_ID = adoRec!USUA_ID
End If

If adoCon.State = adStateOpen Then adoCon.Close
If adoRec.State = adStateOpen Then adoRec.Close
Set adoCon = Nothing
Set adoRec = Nothing

Exit Function

ErrorConexion:
    sMensaje = "Ha ocurrido el siguiente error al intetar conectarse con la base."
    sMensaje = sMensaje & vbCrLf & Err.Description & vbCrLf
    sMensaje = sMensaje & "Por favor comuniquese con Galileo Sistemas S.R.L."
    MsgBox sMensaje, vbOKOnly + vbCritical, "Error al abrir la Base. (Usuario)"
    If adoCon.State = adStateOpen Then adoCon.Close
    If adoRec.State = adStateOpen Then adoRec.Close
    Set adoCon = Nothing
    Set adoRec = Nothing
    VerificarUsuario = False
    Exit Function

ErrorRecordSet:
    sMensaje = "Ha ocurrido un error al intentar abrir la tabla de usuarios."
    sMensaje = sMensaje & vbCrLf & Err.Description & vbCrLf
    sMensaje = sMensaje & "Por favor comuniquese con Galileo Sistemas S.R.L."
    MsgBox sMensaje, vbOKOnly + vbCritical, "Error al abrir la Base. (Usuario)"
    If adoCon.State = adStateOpen Then adoCon.Close
    If adoRec.State = adStateOpen Then adoRec.Close
    Set adoCon = Nothing
    Set adoRec = Nothing
    VerificarUsuario = False
    Exit Function

End Function

Public Function VerificarClave(ByVal sCodigo As String, ByVal sClave As String) As Boolean
'***********************************************'
'                                               '
'  Funcion para verificar la clave del          '
'  usuario.                                     '
'                                               '
'  Hecho por: Fernando (Chicho)                 '
'  Fecha de creacion: 06/02/2002                '
'                                               '
'***********************************************'

Dim adoCon As New ADODB.Connection
Dim adoRec As New Recordset
Dim sMensaje As String
Static iError As Integer

sClave = Encripta(sClave)

On Error GoTo ErrorConexion
adoCon.CursorLocation = adUseClient
adoCon.Open SGINST

On Error GoTo ErrorRecordSet
adoRec.ActiveConnection = adoCon
adoRec.CursorLocation = adUseClient
adoRec.CursorType = adOpenStatic
adoRec.LockType = adLockReadOnly
adoRec.source = "Select * From SGUSUARIO Where USUASENIAL = '" & Encripta(CInt(True)) & "'"
adoRec.source = adoRec.source & " AND USUACODIGO = '" & sCodigo & "'"
adoRec.Open , , , , adCmdText

adoRec.ActiveConnection = Nothing
adoCon.Close
Set adoCon = Nothing

If adoRec.RecordCount = 0 Then
    sMensaje = "El usuario ingresado no existe en el sistema o no esta habilitado."
    sMensaje = sMensaje & vbCrLf & "Por favor comuniquese con su supervisor"
    MsgBox sMensaje, vbOKOnly + vbExclamation, "Usuario: " & txtCodigo.Text & ", incorrecto."
    VerificarClave = False
Else
    If adoRec!USUACLAVE = sClave Or (IsNull(adoRec!USUACLAVE) And sClave = "") Then
        VerificarClave = True
        Set tmpsgUser = New sgUsuario
        tmpsgUser.USUA_ID = adoRec!USUA_ID
        tmpsgUser.USUACLAVE = IIf(IsNull(adoRec!USUACLAVE), "", adoRec!USUACLAVE)
        tmpsgUser.USUANOMBRE = adoRec!USUANOMBRE
        tmpsgUser.USUAPREGUNTA = IIf(IsNull(adoRec!USUAPREGUNTA), "", adoRec!USUAPREGUNTA)
        tmpsgUser.USUARESPUESTA = IIf(IsNull(adoRec!USUARESPUESTA), "", adoRec!USUARESPUESTA)
        tmpsgUser.USUACODIGO = adoRec!USUACODIGO
        lUSUA_ID = adoRec!USUA_ID
        txtNombre.Text = adoRec!USUANOMBRE
    Else
        sMensaje = "La clave ingresada es invalida." & vbCrLf
        sMensaje = sMensaje & "Por favor verifiquela."
        MsgBox sMensaje, vbOKOnly + vbExclamation, "Clave incorrecta..."
        VerificarClave = False
        If iError = 2 Then
            MostrarPregunta
            iError = 0
        Else
            iError = iError + 1
        End If
    End If
End If

If adoCon.State = adStateOpen Then adoCon.Close
If adoRec.State = adStateOpen Then adoRec.Close
Set adoCon = Nothing
Set adoRec = Nothing

Exit Function

ErrorConexion:
    sMensaje = "Ha ocurrido el siguiente error al intetar conectarse con la base."
    sMensaje = sMensaje & vbCrLf & Err.Description & vbCrLf
    sMensaje = sMensaje & "Por favor comuniquese con Galileo Sistemas S.R.L."
    MsgBox sMensaje, vbOKOnly + vbCritical, "Error al abrir la Base. (Usuario)"
    If adoCon.State = adStateOpen Then adoCon.Close
    If adoRec.State = adStateOpen Then adoRec.Close
    Set adoCon = Nothing
    Set adoRec = Nothing
    VerificarClave = False
    Exit Function

ErrorRecordSet:
    sMensaje = "Ha ocurrido un error al intentar abrir la tabla de usuarios."
    sMensaje = sMensaje & vbCrLf & Err.Description & vbCrLf
    sMensaje = sMensaje & "Por favor comuniquese con Galileo Sistemas S.R.L."
    MsgBox sMensaje, vbOKOnly + vbCritical, "Error al abrir la Base. (Usuario)"
    If adoCon.State = adStateOpen Then adoCon.Close
    If adoRec.State = adStateOpen Then adoRec.Close
    Set adoCon = Nothing
    Set adoRec = Nothing
    VerificarClave = False
    Exit Function
End Function
Private Sub MostrarPregunta()
Dim sMEN As String
Dim adoRec As ADODB.Recordset
Dim adoCon As New Connection
Set adoRec = New Recordset

adoCon.Open SGINST

With adoRec
    .ActiveConnection = adoCon
    .CursorLocation = adUseClient
    .LockType = adLockReadOnly
    .CursorType = adOpenStatic
    .source = "Select * From SGUSUARIO Where USUASENIAL ='" & Encripta(CInt(True)) & "'"
    .source = .source & " AND USUACODIGO ='" & txtCodigo.Text & "'"
    .Open
    Set .ActiveConnection = Nothing
    adoCon.Close
    Set adoCon = Nothing
End With

If IsNull(adoRec!USUAPREGUNTA) Or IsNull(adoRec!USUARESPUESTA) Then
    MsgBox "No se ha ingresado la Pregunta/Respuesta, por lo tanto no tiene forma de recuperar su clave.", vbOKCancel + vbCritical, "Pregunta/Respuesta..."
    adoRec.Close
    Set adoRec = Nothing
    Exit Sub
End If

    iErrores = 0
    Load frmRecuperacionClave
    frmRecuperacionClave.Show vbModal
    If frmRecuperacionClave.Correcto Then
        txtClave.SetFocus
        txtClave.Text = DesEncripta(adoRec!USUACLAVE)
        SendKeys "{TAB}"
        Unload frmRecuperacionClave
    Else
        sMEN = "La respuesta es incorrecta."
        MsgBox sMEN, vbOKOnly + vbInformation, "Respuesta..."
        Unload frmRecuperacionClave
    Exit Sub
End If
adoRec.Close
Set adoRec = Nothing
End Sub
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get RutaApp() As String
    RutaApp = m_RutaApp
End Property
Private Function BuscarsgEmpresa(ByVal sCodigo As String) As sgEmpresaCLS
    
    For Each BuscarsgEmpresa In tmpsgUser.sgEmpresas
        If BuscarsgEmpresa.EMPRCODIGO = sCodigo Then Exit Function
Next
End Function

Private Sub cmdCerrar_Click()
'Aqui se cierra la ventana donde esta el control

Dim sgEMPR As sgEmpresaCLS

    If bSinEmpresa Then
        
        Set sgEMPR = New sgEmpresaCLS
        sgEMPR.EMPRCODIGO = "CONFIGURACION"
    
'         MostrarSistemas user, sgEMPR.EMPRNOMBRE
    
         frmFondo.mnuReiniciar.Enabled = True
         frmFondo.mnuCerrarSesion.Enabled = True
         frmFondo.mnuCambiarEmpresa.Enabled = True
         frmFondo.mnuCambiarUsuario.Enabled = True
         frmFondo.mnuCambiarUsuario.caption = "&Cambiar usuario"
    
        Shell App.Path & "\Recordatorios.exe " & user.USUACODIGO
    
         Screen.MousePointer = vbNormal
         Set sgEMPR = Nothing
         
    End If
    
    Unload Me
End Sub

Private Sub cmdEntrar_Click()
    
    If lvwEmpresas.ListItems.Count = 0 Then
        Exit Sub
    Else
        user.sgEmpresas.Seleccionada.EMPRCONEXION = user.sgEmpresas.Seleccionada.EMPRAUXILIAR
'        MostrarSistemas user, user.sgEmpresas.Seleccionada.EMPRNOMBRE

        frmFondo.mnuReiniciar.Enabled = True
        frmFondo.mnuCerrarSesion.Enabled = True
        frmFondo.mnuCambiarEmpresa.Enabled = True
        frmFondo.mnuCambiarUsuario.Enabled = True
        frmFondo.mnuCambiarUsuario.caption = "&Cambiar usuario"

       Shell App.Path & "\Recordatorios.exe " & user.USUACODIGO

        Screen.MousePointer = vbNormal
        Unload Me
    End If
    
End Sub

Private Sub cmdIniciar_Click()

'***************************************************'
'                                                   '
'  Aqui es donde se hace toda la verificacion del   '
'  usuario. Y si es correcta se devuelven los datos '
'  a traves del Evento DevolverDatos, como se       '
'  describe en el comentario gral.                  '
'                                                   '
'  Hecho por: Fernando (Chicho)                     '
'  Fecha de creacion: 15/12/2001                    '
'                                                   '
'***************************************************'
    
    Dim sMensaje As String
    Dim cn As Connection
    Dim rs As Recordset
    
    If VerificarClave(txtCodigo.Text, txtClave.Text) Then
        Set user = tmpsgUser
        CargarEmpresas
        
        frmFondo.mnuReiniciar.Enabled = True
        frmFondo.mnuCerrarSesion.Enabled = True
        frmFondo.mnuCambiarUsuario.Enabled = True
        frmFondo.mnuCambiarUsuario.caption = "&Cambiar usuario"
        Screen.MousePointer = vbNormal

        Set cn = New Connection
        cn.Open SGINST
    
        Set rs = New Recordset
        rs.Open "Select Empr_Id from SGRXEMPRUSUA where Empr_Id > 100 and Usua_Id = " & user.USUA_ID, cn, adOpenDynamic, adLockBatchOptimistic, adCmdText
            
        sEmprUsua = "("
        While Not rs.EOF
            sEmprUsua = sEmprUsua & Trim(Str(rs!EMPR_ID)) & ","
            rs.MoveNext
        Wend
        sEmprUsua = Mid(sEmprUsua, 1, Len(sEmprUsua) - 1) & ")"
        
        Unload Me
        frmFondo.Show

    End If
End Sub
Private Sub CargarEmpresas()
    
    Dim cn As Connection
    Dim rs As Recordset
    Dim sList As String
    
    Set cn = New Connection
    Set rs = New Recordset
    cn.Open SGINST
    
    sList = ""
    
    rs.Open "SELECT * FROM SGRXEMPRUSUA WHERE USUA_ID = " & user.USUA_ID, cn, adOpenDynamic, adLockBatchOptimistic, adCmdText
    While Not rs.EOF
        sList = sList & rs!EMPR_ID & ", "
        rs.MoveNext
    Wend
    
    sList = Mid(sList, 1, Len(sList) - 2)
    user.USUARESPUESTA = sList
    
    rs.Close
    cn.Close
    
    Set rs = Nothing
    Set cn = Nothing
    
End Sub
Private Sub cmdSalir_Click()
'En este lugar se debe terminar la aplicacion
'como no puedo poner un End aqui
'salta un evento.
'    CerrarSistema
    
End Sub

Private Sub picLlave2_Click()
    picOlvidoClave_Click
End Sub

Private Sub picOlvidoClave_Click()
Dim sMEN As String
Dim sUSUACODIGO As String

If txtCodigo.Text = "" Then
    sMEN = "Ingrese en la caja de texto de abajo "
    sMEN = sMEN & "su codigo de usuario para poder mostrar la pregunta."
    sUSUACODIGO = InputBox(sMEN, "Por favor ingrese el codigo de usuario...")
    If VerificarUsuario(sUSUACODIGO) Then
        txtCodigo.Text = sUSUACODIGO
        MostrarPregunta
    Else
        sMEN = "El usuario: " & sUSUACODIGO & " es incorrecto."
        MsgBox sMEN, vbOKOnly + vbExclamation, "Usuario incorrecto..."
    End If
Else
    If VerificarUsuario(txtCodigo.Text) Then
        MostrarPregunta
    Else
        sMEN = "El usuario: " & txtCodigo.Text & " es incorrecto."
        MsgBox sMEN, vbOKOnly + vbExclamation, "Usuario incorrecto..."
    End If
End If
  
End Sub

Private Sub tmrImagen_Timer()
    picOlvidoClave.Visible = Not picOlvidoClave.Visible
    picLlave2.Visible = Not picLlave2.Visible
End Sub

