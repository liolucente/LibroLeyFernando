VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sgColAdjuntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As Collection

Public Function Add(NombreArchivo As String, NombreRuta As String) As sgDAdjuntos


    Dim objNewMember As sgDAdjuntos
    Set objNewMember = New sgDAdjuntos

    objNewMember.sNombreArchivo = NombreArchivo
    objNewMember.sNombreRuta = NombreRuta
    
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(i As Variant) As sgDAdjuntos

Dim ii  As Integer

    If IsNumeric(i) Then
        If mCol.Count = 0 Then
            Set Item = Nothing
            Exit Property
        End If
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

Public Function AddSGDAdjuntos(ByVal mAdj As sgDAdjuntos)
    mCol.Add mAdj
End Function

Public Sub Limpiar()
Dim ii  As Integer

    For ii = mCol.Count To 1 Step -1
        mCol.Remove ii
    Next
End Sub

Public Function ExiteEste(ByVal idreg As Long) As Boolean
Dim ii  As Integer

    For ii = mCol.Count To 1 Step -1
        If mCol(ii).idreg = idreg Then
           ExiteEste = True
           Exit For
        End If
    Next
End Function



