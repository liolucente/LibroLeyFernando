VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sgTRCriterio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private mvarCampo As String
Private mvarTitulo As String
Private mvarTexto As String
Private mvarTexto2 As String
Private mvarTipo As sgTRCRIT_TIPO
Private mvarKey As String
Private mvarUsaCombo As Boolean
Private mvarSQLCombo As String
Private mvarSigno As String
Private mvarUsaTexto2 As Boolean

Public Property Let Signo(ByVal vData As String)
    mvarSigno = vData
End Property

Public Property Get Signo() As String
    Signo = mvarSigno
End Property

Public Property Let UsaCombo(ByVal vData As Boolean)
    mvarUsaCombo = vData
End Property

Public Property Get UsaCombo() As Boolean
    UsaCombo = mvarUsaCombo
End Property

Public Property Let UsaTexto2(ByVal vData As Boolean)
    mvarUsaTexto2 = vData
End Property

Public Property Get UsaTexto2() As Boolean
    UsaTexto2 = mvarUsaTexto2
End Property

Public Property Let SQLCombo(ByVal vData As String)
    mvarSQLCombo = vData
End Property

Public Property Get SQLCombo() As String
    SQLCombo = mvarSQLCombo
End Property

Public Property Let Key(ByVal vData As String)
    mvarKey = vData
End Property

Public Property Get Key() As String
    Key = mvarKey
End Property

Public Property Let Tipo(ByVal vData As sgTRCRIT_TIPO)
    mvarTipo = vData
End Property

Public Property Get Tipo() As sgTRCRIT_TIPO
    Tipo = mvarTipo
End Property

Public Property Let Texto(ByVal vData As String)
    mvarTexto = vData
End Property

Public Property Get Texto() As String
    Texto = mvarTexto
End Property

Public Property Let Texto2(ByVal vData As String)
    mvarTexto2 = vData
End Property

Public Property Get Texto2() As String
    Texto2 = mvarTexto2
End Property

Public Property Let Titulo(ByVal vData As String)
    mvarTitulo = vData
End Property

Public Property Get Titulo() As String
    Titulo = mvarTitulo
End Property

Public Property Let Campo(ByVal vData As String)
    mvarCampo = vData
End Property

Public Property Get Campo() As String
    Campo = mvarCampo
End Property

Private Sub Class_Initialize()
Signo = "="
End Sub
