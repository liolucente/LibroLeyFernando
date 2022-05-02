Attribute VB_Name = "mSGXXIGenerales"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3D38740E03CA"
Option Explicit

'Esta variable se usa en el programa de recuperacion
'de clave para pasar al que lo llamo si se respondio
'bien la respuesta
'##ModelId=3D38740F00BE
Public bRecuperacionClaveCorrecta As Boolean

'colFormularios sirve para saver que
'formularios hay cargados
'##ModelId=3D38740F0168
Public colFormularios As New Collection

Public iImpresiones As Integer
Public bCambioOtrosDatos As Boolean
Public bCambioGenerales As Boolean

Public SGINST As String
Public SGEPOB As String
Public SGINTE As String

Public PDFS_DEST As String
Public PDFS_ORIG As String
Public FOX As String
Public MI_EMAIL As String
Public Asunto As String
Public Texto As String
Public Enviar As Boolean
Public Dir_Erroneas As String
Public iError As Boolean

Public ASUNTO1 As String
Public LINEA1 As String
Public LINEA2 As String
Public LINEA3 As String
Public LINEA4 As String
Public LINEA5 As String
Public LINEA6 As String
Public LINEA7 As String
Public LINEA8 As String
Public LINEA9 As String
Public LINEA10 As String

Public TIPOSERVIDOR As String
Public SERVIDOR As String
Public CRDLL As String

Public sEmprUsua As String
Public sCarpeta As String


Public Sub MostrarEstatus(ByVal sMensaje As String)
frmFondo.Estado.Panels(2).Text = sMensaje
frmFondo.Estado.Refresh
Screen.MousePointer = vbHourglass
End Sub

Public Sub CargarGaliABMForm(ByVal frm As Form, ByVal bAgregar As Boolean, ByVal bModificar As Boolean, ByVal bEliminar As Boolean, Optional ByVal Modal As Boolean = False)
Load frm
If frm.Tag = "NOCARGAR" Then
    Unload frm
    Set frm = Nothing
    Exit Sub
End If
frm.Agregar = bAgregar
frm.Modificar = bModificar
frm.Eliminar = bEliminar
frm.ABM.AbrirRec adOpenKeyset, adLockBatchOptimistic, adCmdText
CallByName frm, "Linkear", VbMethod
Screen.MousePointer = vbNormal
frm.Show IIf(Modal, vbModal, vbModeless)
End Sub

Public Sub FormularioCargado(frm As Form, Col As Collection, Optional Agregar As Boolean = True)
On Error Resume Next

If Agregar Then
    Col.Add frm
Else
    Col.Remove IndiceFormulario(frm.name, Col)
End If

End Sub

'##ModelId=3D38740F0208
Public Function ObtenerColPosition(ByVal Cols As SGColumns, ByVal sKey As String) As Long

'**************************************************'
'                                                  '
'  Devuelve la posicion de la columna sKey.        '
'                                                  '
'  Hecho por: Fernando (Chicho)                    '
'  Fecha de creacion: 17/07/2002                   '
'                                                  '
'**************************************************'

Dim l As Long

On Error GoTo NADA
For l = 1 To Cols.Count
    If Cols(l).key = sKey Then
        ObtenerColPosition = l
        Set Cols = Nothing
        Exit Function
    End If
Next l

NADA:
ObtenerColPosition = 0

Set Cols = Nothing

End Function

Public Function IndiceFormulario(ByVal sNombre As String, ByVal Col As Collection)

Dim l As Long
Dim frm As Form

For l = 1 To Col.Count
    Set frm = Col.Item(l)
    If UCase(frm.name) = UCase(sNombre) Then
        IndiceFormulario = l
        Set frm = Nothing
        Exit Function
    End If
Next l

IndiceFormulario = -1


End Function

Public Function Indice(ByVal sKey As String, ByVal Col As Collection)

Dim l As Long
Dim frm As Form

For l = 1 To Col.Count
    If (UCase(Trim(Col(l).key)) = UCase(Trim(sKey))) Or (Val(Col(l).key) = Val(sKey)) Then
        Indice = l
        Exit Function
    End If
Next l

Indice = -1

End Function

Public Function ObtenerProximoIDSGI(ByVal sTabla As String, ByVal sCampo As String) As Long
Screen.MousePointer = vbHourglass
'**************************************************'
'                                                  '
'  Devuelve el ultimo sCampo en sTabla.            '
'                                                  '
'  Nota: lo devuelve como Long asi que no usarla   '
'  para devolver indices alfanumericos.            '
'                                                  '
'  Hecho por: Fernando (Chicho)                    '
'  Fecha de creacion: 19/12/2001                   '
'                                                  '
'**************************************************'


Dim adoRec As New Recordset
Dim adoCon As New Connection
Dim sSQL As String
Dim lUltimo As Long

'adoCon.CursorLocation = adUseClient
adoCon.Open SGINST

With adoRec
    
    .ActiveConnection = adoCon
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly

    sSQL = "Select Max(" & sCampo & ") + 1 as Ultimo From " & sTabla
'    .source = sSQL & " WITH (READUNCOMMITTED) "
    .Open sSQL
    
    ObtenerProximoIDSGI = IIf(IsNull(adoRec!Ultimo), 1, adoRec!Ultimo)
    
    .Close
End With

adoCon.Close

Set adoRec = Nothing
Set adoCon = Nothing

Screen.MousePointer = vbNormal

End Function

Public Function ObtenerProximoID(ByVal sConexion As String, ByVal sTabla As String, ByVal sCampo As String, Optional sWhere As String = "") As Long
Screen.MousePointer = vbHourglass
'**************************************************'
'                                                  '
'  Devuelve el ultimo sCampo en sTabla.            '
'                                                  '
'  Nota: lo devuelve como Long asi que no usarla   '
'  para devolver indices alfanumericos.            '
'                                                  '
'  Hecho por: Fernando (Chicho)                    '
'  Fecha de creacion: 19/12/2001                   '
'                                                  '
'**************************************************'


Dim adoRec As New Recordset
Dim adoCon As New Connection
Dim sSQL As String
Dim lUltimo As Long

'adoCon.CursorLocation = adUseClient
adoCon.Open sConexion
adoCon.CommandTimeout = 20000

With adoRec
    
    .ActiveConnection = adoCon
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly

'    sSQL = "Select Max(" & sCampo & ") + 1 as Ultimo From " & sTabla & " WITH (READUNCOMMITTED) " & sWhere
    sSQL = "Select Max(" & sCampo & ") + 1 as Ultimo From " & sTabla & " " & sWhere
    .source = sSQL
    .Open
    
    ObtenerProximoID = IIf(IsNull(adoRec!Ultimo), 1, adoRec!Ultimo)
    
    .Close
End With

adoCon.Close

Set adoRec = Nothing
Set adoCon = Nothing

Screen.MousePointer = vbNormal

End Function

Public Function ExisteValor(ByVal sConexion As String, ByVal sTabla As String, ByVal sCampo As String, SValor) As Boolean
Screen.MousePointer = vbHourglass

'**************************************************'
'                                                  '
'  Hecho por: Fernando (Chicho)                    '
'  Fecha de creacion: 19/12/2001                   '
'                                                  '
'**************************************************'


Dim adoRec As New Recordset
Dim adoCon As New Connection
Dim sSQL As String

adoCon.CursorLocation = adUseClient
adoCon.Open sConexion

With adoRec
    
    .ActiveConnection = adoCon
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly

    sSQL = "Select " & sCampo & " From " & sTabla & " Where " & sCampo & "=" & Replace(ConvertirDato(SValor), "''", "'")
    .source = sSQL
    .Open
    
    ExisteValor = IIf(.RecordCount = 0, False, True)
    
    .Close
End With

adoCon.Close

Set adoRec = Nothing
Set adoCon = Nothing

Screen.MousePointer = vbNormal

End Function

Public Function ExisteValorEX(ByVal sConexion As String, ByVal sTabla As String, ByVal sWhere As String) As Boolean
Screen.MousePointer = vbHourglass

'**************************************************'
'                                                  '
'  Hecho por: Fernando (Chicho)                    '
'  Fecha de creacion: 19/12/2001                   '
'                                                  '
'**************************************************'


Dim adoRec As New Recordset
Dim adoCon As New Connection
Dim sSQL As String

adoCon.CursorLocation = adUseClient
adoCon.Open sConexion

With adoRec
    
    .ActiveConnection = adoCon
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly

    sSQL = "Select 1 From " & sTabla & " Where " & sWhere
    .source = sSQL
    .Open
    
    ExisteValorEX = IIf(.RecordCount = 0, False, True)
    
    .Close
End With

adoCon.Close

Set adoRec = Nothing
Set adoCon = Nothing

Screen.MousePointer = vbNormal

End Function

Public Function SeCambio(ByVal adoRec As Recordset) As Boolean

'**************************************************'
'                                                  '
'  Esta funcion devuleve True si se modifico algun '
'  campo en algun registro de adoRec.              '
'                                                  '
'**************************************************'

adoRec.MoveFirst
While Not adoRec.EOF
    If adoRec.Status = adRecModified Then
        SeCambio = True
        adoRec.MoveFirst
        Exit Function
    End If
    adoRec.MoveNext
Wend
adoRec.MoveFirst
SeCambio = False

End Function


Public Function ObtenerValor_EX(ByVal Con As Connection, ByVal sTabla As String, ByVal sCampo As String, Optional Tipo As DataTypeEnum = adIUnknown, Optional sWhere As String = "") As Variant
Screen.MousePointer = vbHourglass
'**************************************************'
'                                                  '
'  Devuelve el valor de sCampo en sTabla.          '
'  Esta version usa una conexion que se le pasa    '
'                                                  '
'  Hecho por: Fernando (Chicho)                    '
'  Fecha de creacion: 27/10/2004                   '
'                                                  '
'**************************************************'


Dim adoRec As New Recordset
Dim sSQL As String

If Not Con.State = adStateOpen Then
    On Error GoTo NoTieneConString
    Con.Open
    On Error GoTo 0
End If

With adoRec

    .ActiveConnection = Con
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly

'    sSQL = "Select " & sCampo & " From " & sTabla & " WITH (READUNCOMMITTED) " & IIf(InStr(1, sWhere, "WHERE", vbTextCompare) = 0, " WHERE " & sWhere, sWhere)
    sSQL = "Select " & sCampo & " From " & sTabla & " " & IIf(InStr(1, sWhere, "WHERE", vbTextCompare) = 0, " WHERE " & sWhere, sWhere)
    .source = sSQL
    .Open

    On Error GoTo NoHayValor

    If InStr(1, sCampo, " as ", vbTextCompare) Then
        sCampo = Right(sCampo, Len(sCampo) - InStr(1, sCampo, "as", vbTextCompare) - 2)
    End If

    ObtenerValor_EX = IIf(IsNull(adoRec.Fields(0).value), "", adoRec.Fields(0).value)

    .Close

End With

Con.Close

Screen.MousePointer = vbNormal

If Tipo = adDate Then
    ObtenerValor_EX = IIf(ObtenerValor_EX = -1, Date, CDate(ObtenerValor_EX))
ElseIf Tipo = adCurrency Then
    ObtenerValor_EX = IIf(ObtenerValor_EX = -1, 0, CCur(ObtenerValor_EX))
ElseIf Tipo = adNumeric Then
    ObtenerValor_EX = IIf(ObtenerValor_EX = -1, 1, CLng(ObtenerValor_EX))
End If

Set adoRec = Nothing
Set Con = Nothing

Screen.MousePointer = vbNormal

Exit Function
NoHayValor:
    If Tipo = adDate Then
        ObtenerValor_EX = Format(Date, "dd/mm/yyyy")
    ElseIf Tipo = adCurrency Then
        ObtenerValor_EX = 0
    ElseIf Tipo = adNumeric Then
        ObtenerValor_EX = 0
    End If
    
    If Not Con Is Nothing Then
        If Con.State = adStateOpen Then Con.Close
        Set Con = Nothing
    End If
        If Not Con Is Nothing Then
        If Con.State = adStateOpen Then Con.Close
        Set Con = Nothing
    End If
    Screen.MousePointer = vbNormal
    
NoTieneConString:
    ObtenerValor_EX = -999
    Screen.MousePointer = vbNormal


End Function

Public Function ObtenerValor(ByVal sConexion As String, ByVal sTabla As String, ByVal sCampo As String, Optional Tipo As DataTypeEnum = adIUnknown, Optional sWhere As String = "") As Variant
Screen.MousePointer = vbHourglass
'**************************************************'
'                                                  '
'  Devuelve el valor de sCampo en sTabla.          '
'                                                  '
'  Hecho por: Fernando (Chicho)                    '
'  Fecha de creacion: 19/12/2001                   '
'                                                  '
'**************************************************'


Dim adoRec As New Recordset
Dim adoCon As New Connection
Dim sSQL As String

'adoCon.CursorLocation = adUseClient
adoCon.Open sConexion

With adoRec

    .ActiveConnection = adoCon
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly

'    sSQL = "Select " & sCampo & " From " & sTabla & " WITH (READUNCOMMITTED) " & IIf(InStr(1, sWhere, "WHERE", vbTextCompare) = 0, " WHERE " & sWhere, sWhere)
    sSQL = "Select " & sCampo & " From " & sTabla & " " & IIf(InStr(1, sWhere, "WHERE", vbTextCompare) = 0, " WHERE " & sWhere, sWhere)
    
    .source = sSQL
    .Open

    On Error GoTo NoHayValor

    If InStr(1, sCampo, " as ", vbTextCompare) Then
        sCampo = Right(sCampo, Len(sCampo) - InStr(1, sCampo, "as", vbTextCompare) - 2)
    End If

    ObtenerValor = IIf(IsNull(adoRec.Fields(0).value), "", adoRec.Fields(0).value)

    .Close

End With

adoCon.Close

Screen.MousePointer = vbNormal

If Tipo = adDate Then
    ObtenerValor = IIf(ObtenerValor = -1, Date, CDate(ObtenerValor))
ElseIf Tipo = adCurrency Then
    ObtenerValor = IIf(ObtenerValor = -1, 0, CCur(ObtenerValor))
ElseIf Tipo = adNumeric Then
    ObtenerValor = IIf(ObtenerValor = -1, 1, CLng(ObtenerValor))
End If

Set adoRec = Nothing
Set adoCon = Nothing

Screen.MousePointer = vbNormal

Exit Function
NoHayValor:
    If Tipo = adDate Then
        ObtenerValor = Format(Date, "dd/mm/yyyy")
    ElseIf Tipo = adCurrency Then
        ObtenerValor = 0
    ElseIf Tipo = adNumeric Then
        ObtenerValor = 0
    End If
    
    If Not adoCon Is Nothing Then
        If adoCon.State = adStateOpen Then adoCon.Close
        Set adoCon = Nothing
    End If
        If Not adoCon Is Nothing Then
        If adoCon.State = adStateOpen Then adoCon.Close
        Set adoCon = Nothing
    End If
    Screen.MousePointer = vbNormal


End Function

Public Sub BuscarParecidoCBOIMG(ByRef cbo As ImageCombo, ByVal sTexto As String)
Dim ct As ComboItem

For Each ct In cbo.ComboItems
    If UCase(ct.Text) Like (UCase(sTexto) & "*") Then
        ct.Selected = True
        SendKeys "{DOWN}"
        SendKeys "{UP}"
        'SendKeys "{DOWN}"
        Exit Sub
    End If
Next

End Sub

Public Sub CheckearLista(ByRef lst As ListItems, ByRef sCodigos() As String)
'Dim lstITM As ListItem
'Dim iCod As Integer
'
'For Each lstITM In lst
'    lstITM.Checked = False
'Next
'Set lstITM = Nothing
'
'For iCod = LBound(sCodigos) To UBound(sCodigos)
'    If Not sCodigos(iCod) = "" Then
'        For Each lstITM In lst
'            If UCase(lstITM.Tag) = UCase(Trim(sCodigos(iCod))) Then
'                lstITM.Checked = True
'                Exit For
'            End If
'        Next
'    End If
'Next
'
'Set lstITM = Nothing

End Sub

Public Function ConvertirFecha(vFecha As Variant) As Variant
Dim lFecha As Long
Dim dFecha As Date
Dim sDia As String
Dim sMes As String
Dim sAnio As String

If VarType(vFecha) = vbNull Then
    ConvertirFecha = Format(Date, "dd/MM/yyyy")
    Exit Function
End If

If Val(vFecha) = 0 Then
    ConvertirFecha = Val(vFecha)
Else

    If VarType(vFecha) = vbDate Then
        lFecha = CLng(Format(vFecha, "yyyymmdd"))
        ConvertirFecha = lFecha
    ElseIf Not VarType(vFecha) = vbString Then
        sAnio = Left(CStr(vFecha), 4)
        sMes = Mid(CStr(vFecha), 5, 2)
        sDia = Right(CStr(vFecha), 2)
        dFecha = DateSerial(CInt(sAnio), CInt(sMes), CInt(sDia))
        ConvertirFecha = CDate(Format(dFecha, "dd/MM/yyyy"))
    Else
        ConvertirFecha = vFecha
    End If
End If

End Function

Public Function CrearTMPBackUp() As Boolean
Dim cn As Connection
Dim sSQL As String
Dim FSO As FileSystemObject
Dim sBaseEmpresa As String
Dim Emp As sgEmpresaCLS

CrearTMPBackUp = True


Select Case TIPOSERVIDOR
Case "SQL"

    Set FSO = New FileSystemObject
    If FSO.FileExists(App.Path & "\BACKSGI.TMP") Then
        FSO.DeleteFile App.Path & "\BACKSGI.TMP", True
    End If
    
    If FSO.FileExists(App.Path & "\BACKSGE.TMP") Then
        FSO.DeleteFile App.Path & "\BACKSGE.TMP", True
    End If
    Set FSO = Nothing

    Set cn = New Connection
    cn.Open user.sgEmpresas.Seleccionada.EMPRAUXILIAR
    cn.Execute "USE MASTER"
    
    sBaseEmpresa = SGINST
    sBaseEmpresa = Right(sBaseEmpresa, Len(sBaseEmpresa) - InStr(1, sBaseEmpresa, "Initial Catalog=", vbTextCompare) + 1)
    sBaseEmpresa = Replace(sBaseEmpresa, "Initial Catalog=", "", , , vbTextCompare)
    sBaseEmpresa = Left(sBaseEmpresa, InStr(1, sBaseEmpresa, ";", vbTextCompare) - 1)
    
    cn.Execute " BACKUP DATABASE " & sBaseEmpresa & " TO DISK='" & App.Path & "\BACKSGI.TMP' WITH INIT"
    
    Dim iEmp As Integer
    iEmp = 1
    For Each Emp In user.sgEmpresas
    
        sBaseEmpresa = Emp.EMPRAUXILIAR
        sBaseEmpresa = Right(sBaseEmpresa, Len(sBaseEmpresa) - InStr(1, sBaseEmpresa, "Initial Catalog=", vbTextCompare) + 1)
        sBaseEmpresa = Replace(sBaseEmpresa, "Initial Catalog=", "", , , vbTextCompare)
        sBaseEmpresa = Left(sBaseEmpresa, InStr(1, sBaseEmpresa, ";", vbTextCompare) - 1)

        cn.Execute "USE MASTER"
        cn.Execute " BACKUP DATABASE " & sBaseEmpresa & " TO DISK='" & App.Path & "\BACKSGE" & Format(iEmp, "0000") & ".TMP'  WITH INIT"
        
        iEmp = iEmp + 1
    Next

    cn.Close
    Set cn = Nothing
Case "ACCESS"
    MsgBox "No se puede hacer el back-up en bases de datos Access." & vbCrLf & "Disculpe las molestias ocasionadas.", vbOKOnly + vbCritical, "No implementado aun."
    CrearTMPBackUp = False
End Select
End Function


Public Function UnicoEnSgrid(ByRef Grilla As SGGrid, RowKey As Long, cell As Long) As Boolean
    Dim i As Integer
    UnicoEnSgrid = True
    For i = 1 To Grilla.Rows.Count - 1
        If Grilla.Rows.At(i).key <> RowKey And Grilla.Rows.At(i).Cells(cell).value = Grilla.Rows(RowKey).Cells(cell).value Then
            UnicoEnSgrid = False
        End If
    Next
End Function
Public Function ConvertirPeriodo(per As String) As String
    ConvertirPeriodo = Mid(per, 7, 4) & Mid(per, 4, 2) & "01"
End Function

