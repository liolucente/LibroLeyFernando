VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sgTRCriterios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"sgTRCriterio"
Attribute VB_Ext_KEY = "Member0" ,"sgTRCriterio"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection
Public Enum sgTRCRIT_TIPO
    sgnumerico = 0
    sgTexto = 1
    sgFecha = 2
    sgMoneda = 3
    sgHora = 4
End Enum


Public Function AddCriterio(ByVal Crit As sgTRCriterio)
mCol.Add Crit
End Function

Public Function Add(Campo As String, Titulo As String, TIPO As sgTRCRIT_TIPO, Optional sKey As String) As sgTRCriterio
    Dim objNewMember As sgTRCriterio
    Set objNewMember = New sgTRCriterio

    objNewMember.Campo = Campo
    objNewMember.Titulo = Titulo
    objNewMember.TIPO = TIPO
    objNewMember.key = sKey
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(i As Variant) As sgTRCriterio
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
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Public Function ObtenerSQL() As String
Dim Crit As sgTRCriterio
Dim sTMP As String
Dim iCuantos As Integer
Dim lFecha1 As Long
Dim lFecha2 As Long
Dim Valor1 As String
Dim Valor2 As String
Dim cadena1 As String
Dim cadena2 As String

For Each Crit In mCol
    If Not Crit.Texto = "" Then
'        sTMP = sTMP & "( "
        Select Case Crit.TIPO
        Case sgTexto
            cadena1 = Crit.Texto
            If Crit.UsaTexto2 And (Not Crit.UsaTexto2 = 0) And (Crit.Texto2 <> "") Then
                cadena2 = Crit.Texto2
                sTMP = sTMP & Crit.Campo & " BETWEEN " & "'" & cadena1 & "'" & " AND " & "'" & cadena2 & "'"
            Else
                sTMP = sTMP & " UPPER(" & Crit.Campo & ") like UPPER('%" & Crit.Texto & "%')"
            End If
            
        Case sgnumerico
            Valor1 = Crit.Texto
            If Crit.UsaTexto2 And (Not Len(Trim(Crit.UsaTexto2)) = 0) And (Crit.Texto2 <> "") Then
                Valor2 = Crit.Texto2
                sTMP = sTMP & Crit.Campo & " BETWEEN " & Valor1 & " AND " & Valor2 & ""
            Else
                sTMP = sTMP & Crit.Campo & Crit.Signo & Valor1 & ""
            End If
            
        Case sgMoneda
            Valor1 = " Convert(Money,'" & Crit.Texto & "')"
            If Crit.UsaTexto2 And (Not Len(Trim(Crit.UsaTexto2)) = 0) Then
                Valor2 = " Convert(Money,'" & Crit.Texto2 & "')"
                sTMP = sTMP & Crit.Campo & " BETWEEN " & Valor1 & " AND " & Valor2 & ""
            Else
                sTMP = sTMP & Crit.Campo & Crit.Signo & Valor1 & ""
            End If
            
        Case sgFecha
            lFecha1 = ConvertirFecha(CDate(Format(Crit.Texto, "dd/MM/yyyy")))
            If Crit.UsaTexto2 And (Not Len(Trim(Crit.UsaTexto2)) = 0) And (Crit.Texto2 <> "") Then
                lFecha2 = ConvertirFecha(CDate(Format(Crit.Texto2, "dd/MM/yyyy")))
                sTMP = sTMP & Crit.Campo & " BETWEEN " & lFecha1 & " AND " & lFecha2 & ""
            Else
                sTMP = sTMP & Crit.Campo & Crit.Signo & lFecha1 & ""
            End If
            
        Case sgHora
            Valor1 = "Convert(DateTime,'" & Format(Crit.Texto, "17/04/2000 HH:mm:ss") & "')"
            If Crit.UsaTexto2 And (Not Len(Trim(Crit.UsaTexto2)) = 0) Then
                Valor2 = "Convert(DateTime,'" & Format(Crit.Texto2, "17/04/2000 HH:mm:ss") & "')"
                sTMP = sTMP & Crit.Campo & " BETWEEN " & Valor1 & " AND " & Valor2 & ""
            Else
                sTMP = sTMP & Crit.Campo & Crit.Signo & Valor1 & ""
            End If
        End Select
        sTMP = sTMP & " AND "
    End If
Next

sTMP = Trim(sTMP)
If sTMP = "" Then
    ObtenerSQL = ""
Else
    sTMP = Left(sTMP, Len(sTMP) - 4)
End If

ObtenerSQL = sTMP
sTMP = Empty
End Function
