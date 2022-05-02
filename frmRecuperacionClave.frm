VERSION 5.00
Begin VB.Form frmRecuperacionClave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recuperacion de Claves"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4620
   Icon            =   "frmRecuperacionClave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   735
      Left            =   2348
      Picture         =   "frmRecuperacionClave.frx":2072
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1710
      Width           =   1545
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   735
      Left            =   728
      MaskColor       =   &H00FF00FF&
      Picture         =   "frmRecuperacionClave.frx":40E4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1710
      UseMaskColor    =   -1  'True
      Width           =   1545
   End
   Begin VB.Frame Frame1 
      Caption         =   "Responda la pregunta para obtener el acceso"
      Height          =   1635
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4605
      Begin VB.TextBox txtRespuesta 
         Height          =   375
         Left            =   135
         TabIndex        =   3
         Top             =   1170
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Su respuesta"
         Height          =   240
         Left            =   135
         TabIndex        =   2
         Top             =   945
         Width           =   1050
      End
      Begin VB.Label lblPregunta 
         Caption         =   "Label1"
         Height          =   510
         Left            =   135
         TabIndex        =   1
         Top             =   315
         Width           =   4425
      End
   End
End
Attribute VB_Name = "frmRecuperacionClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3D387418001E"
Option Explicit

'##ModelId=3D3874180136
Dim lID As Long
'##ModelId=3D38741801D8
Dim adoCon As Connection
'##ModelId=3D3874180214
Dim adoRec As Recordset
'##ModelId=3D3874180215
Dim mvarCorrecto As Boolean

'##ModelId=3D38741802EE
Public Property Let Usuario(ByVal sUser As String)
With adoRec
    If .State = adStateOpen Then .Close
    .Source = Left(.Source, InStr(1, .Source, "=") + 1)
    .Source = .Source & ObtenerUsuario(sUser)
    If .State = adStateOpen Then .Close
    .Open
    txtRespuesta.Tag = !USUARESPUESTA
    lblPregunta.Caption = DesEncripta(!USUAPREGUNTA)
    .Close
End With

adoCon.Close

Set adoCon = Nothing
Set adoRec = Nothing

End Property

'##ModelId=3D38741900BE
Private Sub cmdAceptar_Click()
If Encripta(txtRespuesta) = txtRespuesta.Tag Then
    mvarCorrecto = True
Else
    mvarCorrecto = False
End If
Me.Hide
End Sub

'##ModelId=3D387419015E
Private Sub cmdCancelar_Click()
mvarCorrecto = False
Me.Hide
End Sub

'##ModelId=3D38741901CC
Private Sub Form_Load()
Set adoCon = New ADODB.Connection
Set adoRec = New ADODB.Recordset

With adoCon
    .CommandTimeout = 15
    .ConnectionString = SGINST
    .CursorLocation = adUseClient
    .Open
End With

With adoRec
    .ActiveConnection = adoCon
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
    .Source = "Select USUAPREGUNTA, USUARESPUESTA From SGUSUARIO Where USUA_ID =  "
End With

End Sub

'##ModelId=3D3874190276
Private Sub txtRespuesta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub

'##ModelId=3D38741A000A
Public Property Get Correcto()
Correcto = mvarCorrecto
End Property
