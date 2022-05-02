Attribute VB_Name = "mSGXXIUsuarios"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3D38740D01AE"
Option Explicit

'Variable para guardar los datos del usuario
'recibidos en el evento DevolverDAtos del SGInicio
Public user As sgUsuario

Public Function ExisteUsuario(ByVal sCodigo As String) As Boolean
Screen.MousePointer = vbHourglass
'**************************************************'
'                                                  '
'  Esta funcion devuelve True si se encontro al    '
'  usuario con el codigo indicado.                 '
'                                                  '
'  Hecho por: Fernando (Chicho)                    '
'  Fecha de creacion: 18/12/2001                   '
'                                                  '
'**************************************************'


Dim adoCon As New ADODB.Connection
Dim adoRec As New ADODB.Recordset
Dim sSQL As String

adoCon.CursorLocation = adUseClient
adoCon.Open SGINST

With adoRec
    .ActiveConnection = adoCon
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
'    sSQL = "Select USUA_ID FRom SGUSUARIO " & " WITH (READUNCOMMITTED) Where "
    sSQL = "Select USUA_ID FRom SGUSUARIO " & " Where "
    sSQL = sSQL & " USUACODIGO='" & sCodigo & "'"
    .source = sSQL
    .Open
    
    If .RecordCount = 0 Then
        ExisteUsuario = False
    Else
        ExisteUsuario = True
    End If
    
    .Close

End With

Set adoRec = Nothing

adoCon.Close
Set adoCon = Nothing

Screen.MousePointer = vbNormal

End Function

Public Function ObtenerUsuario(Optional ByVal sCodigo As String, Optional ByVal lID As Long) As Variant
Screen.MousePointer = vbHourglass
'**************************************************'
'                                                  '
'  Devuelve el Id o el Codigo, segun lo que falte. '
'                                                  '
'  Hecho por: Fernando (Chicho)                    '
'  Fecha de creacion: 19/12/2001                   '
'                                                  '
'**************************************************'

Dim adoRec As New Recordset
Dim adoCon As New Connection
Dim sSQL As String

adoCon.CursorLocation = adUseClient
adoCon.Open SGINST


With adoRec

    .ActiveConnection = adoCon
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly

    sSQL = "SGUSUARIO"
'    .source = sSQL & " WITH (READUNCOMMITTED) "
    .source = sSQL & " "
    .Open
    
    If Not IsMissing(sCodigo) Then
        
        adoRec.MoveFirst
        adoRec.Find "USUACODIGO='" & sCodigo & "'"
        If adoRec.EOF Then
            ObtenerUsuario = "NOEXISTE"
        Else
            ObtenerUsuario = adoRec!USUA_ID
        End If
    
    Else
        
        adoRec.MoveFirst
        adoRec.Find "USUA_ID=" & lID
        'If adoRec.EOF Then
        '    ObtenerUsuario = "NOEXISTE"
        'Else
            ObtenerUsuario = adoRec!USUACODIGO
        'End If
        
    End If
 
    .Close
    
End With

adoCon.Close

Set adoRec = Nothing
Set adoCon = Nothing

Screen.MousePointer = vbNormal

End Function

Public Sub BorrarUsuario(ByVal sCodigo As String)
Dim lID As Long
Dim adoCon As Connection
Dim sSQL As String

lID = ObtenerUsuario(sCodigo)

Set adoCon = New Connection
With adoCon
    .ConnectionString = SGINST
    .CursorLocation = adUseServer
    .Open
    .BeginTrans
    
    sSQL = "Delete From SGRXEMPRUSUA where USUA_ID = " & lID
    .Execute sSQL
    
    sSQL = "Delete FRom SGRXSISTUSUA Where USUA_ID = " & lID
    .Execute sSQL
    
    sSQL = "Delete FRom SGRXUSUALLAV Where USUA_ID =" & lID
    .Execute sSQL
    
    sSQL = "Delete FRom SGACCESO Where USUA_ID = " & lID
    .Execute sSQL
    
    .CommitTrans
    .Close
    
End With
    
Set adoCon = Nothing

End Sub

