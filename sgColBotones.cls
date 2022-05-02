VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sgColBotones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mCol As Collection

Public Function AddBoton(ByVal Boton As sgBoton)
    mCol.Add Boton
End Function

Public Function Add(Imagen As String, _
                    sTag As String, _
                    Tipo As Integer, _
                    Accion As String, _
                    Optional sKey As String) As sgBoton
                    
    Dim objNewMember As sgBoton
    Set objNewMember = New sgBoton

    objNewMember.sImagen = Imagen
    objNewMember.sTag = sTag
    objNewMember.Tipo = Tipo
    objNewMember.Accion = Accion
    
    objNewMember.key = sKey
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(i As Variant) As sgBoton
Dim ii  As Integer

    If IsNumeric(i) Then
        Set Item = mCol(i)
    Else
        For ii = 1 To mCol.Count
            If mCol(ii).key = i Then
                Set Item = mCol(ii)
                Exit For
            End If
        Next
    End If
    
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(i As Variant)
Dim ii  As Integer

    If IsNumeric(i) Then
        mCol.Remove i
    Else
        For ii = 1 To mCol.Count
            If mCol(ii).key = i Then
                mCol.Remove ii
                Exit For
            End If
        Next
    End If
End Sub

Public Property Get NewEnum() As IUnknown
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub



