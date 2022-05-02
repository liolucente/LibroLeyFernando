Attribute VB_Name = "mSGXXISistema"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3D3874090212"
Option Explicit

Public Const sNombreSistema As String = "Galileo XXI ©"
Public Const sNombreSistemaVerificar As String = "XXI_LT"

'Usado para linkear correctamente al ComCtl32.dll
'Para que se puedan ver bien los

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function

Public Sub Main()
    InitCommonControlsVB
    
    frmSplash.Show
End Sub

Public Sub CrearDatosEmpresa(ByVal DSN As String, ByVal sNombre As String, ByVal sComenta As String)
Dim conTMP As Connection
Dim recTMP As Recordset

Set conTMP = New Connection

conTMP.CursorLocation = adUseServer
conTMP.Open DSN
Set recTMP = New Recordset

recTMP.CursorLocation = adUseServer
recTMP.CursorType = adOpenKeyset
recTMP.LockType = adLockBatchOptimistic
recTMP.Open "RAZON_SOCIAL", conTMP, , , adCmdTable
On Error Resume Next
recTMP.MoveFirst
recTMP.Delete
recTMP.UpdateBatch adAffectAllChapters
recTMP.AddNew
recTMP!RAZS_ID = 1
recTMP!RAZSNOMBRE = sNombre
recTMP!RAZSCOMENTA = Left(sComenta, 40)

recTMP.UpdateBatch
recTMP.Close
Set recTMP = Nothing
conTMP.Close
Set conTMP = Nothing

End Sub

Public Sub CargarSistema(ByVal sSistema As String) ' , ByVal Boton As JSGroupItem)

'***************************************************'
'                                                   '
'  Esta funcion verifica que el usuario pueda       '
'  ejecutar el sistema ( pasado como string en      '
'  sSistema) y si lo tiene lo ejecuta               '
'                                                   '
'  Hecho por: Fernando (Chicho)                     '
'  Fecha de creacion: 13/12/2001                    '
'                                                   '
'***************************************************'

Dim adoSis As New ADODB.Recordset
Dim adoCon As New ADODB.Connection
Dim sSIST_ID As String
Dim sSQL As String
Dim lDatos As Long
Dim sMensaje As String

sSIST_ID = ""
For lDatos = 1 To user.sgSistemas.Count
    sSIST_ID = sSIST_ID & ConvertirDato(user.sgSistemas(lDatos).SIST_ID) & ", "
Next lDatos

If user.sgSistemas.Count = 0 Then
    sSIST_ID = "=" & Left(sSIST_ID, Len(sSIST_ID) - 2)
Else
    sSIST_ID = "IN (" & Left(sSIST_ID, Len(sSIST_ID) - 2) & ")"
End If

On Error GoTo ErrorConexion
adoCon.CursorLocation = adUseClient
adoCon.Open SGINST

adoSis.ActiveConnection = adoCon
adoSis.CursorLocation = adUseClient
adoSis.CursorType = adOpenStatic
adoSis.LockType = adLockReadOnly
sSQL = "Select SISTCODIGO From SGSISTEMA WHERE "
sSQL = sSQL & " SIST_ID " & sSIST_ID & " AND "
sSQL = sSQL & " SISTCODIGO ='" & sSistema & "'"
On Error GoTo ErrorRecordSet
adoSis.Open sSQL, , , , adCmdText

If adoSis.RecordCount = 0 Then
    sMensaje = "Usted no esta habilitado para ejecutar este sistema."
    sMensaje = sMensaje & vbCrLf & "Por favor contacte a su supervisor."
    MsgBox sMensaje, vbOKOnly + vbInformation, "No esta habilitado..."
End If

adoSis.Close
adoCon.Close
Set adoSis = Nothing
Set adoCon = Nothing

On Error GoTo ErrorEjecutando
'Ahora cargo el Child Form con los menues de este sistema
'CallByName frmFondo, sSistema, VbMethod

Screen.MousePointer = vbNormal

Exit Sub
ErrorEjecutando:
    sMensaje = "Ha ocurrido un error al ejecutar el modulo: " & sSistema & "." & vbCrLf
    sMensaje = sMensaje & "Comuniquese con mantenimiento." & vbCrLf
    sMensaje = sMensaje & "Descripcion: " & vbCrLf
    sMensaje = sMensaje & "No se ha encontrado la funcion que carga el sistema."
    MsgBox sMensaje, vbOKOnly + vbCritical, "Error al iniciar el sistema: " & sSistema
    On Error Resume Next
    If adoCon.State = adStateOpen Then adoCon.Close
    Set adoCon = Nothing
    If adoSis.State = adStateOpen Then adoSis.Close
    Set adoSis = Nothing
    Screen.MousePointer = vbNormal
    Exit Sub

ErrorConexion:
    sMensaje = "Ha ocurrido un error al intentar "
    sMensaje = sMensaje & "conectarse a la base de instalacion"
    'sMensaje = sMensaje & vbCrLf & "No podra continuar cargando el sistema."
    sMensaje = sMensaje & vbCrLf & "Por favor contactese con Galileo Sistemas S.R.L."
    MsgBox sMensaje, vbOKOnly + vbCritical, "Error al iniciar los sistemas: " & sSistema
    If adoCon.State = adStateOpen Then adoCon.Close
    Set adoCon = Nothing
    If adoSis.State = adStateOpen Then adoSis.Close
    Set adoSis = Nothing
    Screen.MousePointer = vbNormal
    Exit Sub

ErrorRecordSet:
    sMensaje = "Ha ocurrido un error al intentar "
    sMensaje = sMensaje & "abrir la base de instalacion"
    'sMensaje = sMensaje & vbCrLf & "No podra continuar cargando el sistema."
    sMensaje = sMensaje & vbCrLf & "Por favor contactese con Galileo Sistemas S.R.L."
    MsgBox sMensaje, vbOKOnly + vbCritical, "Error al iniciar el sistema: " & sSistema
    If adoCon.State = adStateOpen Then adoCon.Close
    Set adoCon = Nothing
    If adoSis.State = adStateOpen Then adoSis.Close
    Set adoSis = Nothing
    Screen.MousePointer = vbNormal
    Exit Sub
    
End Sub

'##ModelId=3D38740A012C
Public Sub ReiniciarSistema(ByVal bCompleto As Boolean)

'**************************************************'
'                                                  '
'  Este proceso reinicia el sistema completamente  '
'  (vuelve a la pantalla inicial pidiendo usuario) '
'  o solo reinicia para actualizar los parametros  '
'  (con el mismo usuario)                          '
'                                                  '
'  Hecho por: Fernando (Chicho)                    '
'  Fecha de creacion: 18/12/2001                   '
'                                                  '
'**************************************************'

Dim sMensaje As String

If bCompleto Then
    sMensaje = "Esto reiniciara el sistema por completo"
    sMensaje = sMensaje & ", dejandolo en la pantalla de "
    sMensaje = sMensaje & """Inicio de Sesion""."
Else
    sMensaje = "Esto reinicia el sistema, pero "
    sMensaje = sMensaje & "se conserva el usuario y la empresa."
End If

sMensaje = sMensaje & vbCrLf & "¿Desea continuar?"

If MsgBox(sMensaje, vbYesNo + vbQuestion, "Reiniciar el sistema...") = vbNo Then Exit Sub

If bCompleto Then
    CerrarProcesos
'    frmFondo.Barra.Bars.Clear
'    frmFondo.picBarra.Refresh
    frmFondo.mnuReiniciar.Enabled = False
    frmFondo.mnuCambiarEmpresa.Enabled = False
    frmFondo.mnuCambiarUsuario.Enabled = False
    frmFondo.Estado.Panels(1) = ""
    frmFondo.Estado.Panels(3) = ""
    frmInicioSesion.Show vbModal
Else
    CerrarProcesos
'    MostrarSistemas user, frmFondo.Estado.Panels(1)
    frmFondo.SetFocus
End If

Screen.MousePointer = vbNormal

End Sub


'##ModelId=3D38740B00B4
Public Sub CerrarProcesos()
Screen.MousePointer = vbHourglass
'***************************************************'
'                                                   '
'  Procedimiento que llama a la funcion cerrame de  '
'  cada Form para que se cierre correctamente.      '
'                                                   '
'  Hecho por: Fernando (Chicho)                     '
'  Fecha de creacion: en creacion                   '
'                                                   '
'***************************************************'

Dim frm As Form
Dim lItem As Long

For lItem = colFormularios.Count To 1 Step -1
    Set frm = colFormularios.Item(lItem)
    Unload frm
    Set frm = Nothing
Next

Screen.MousePointer = vbNormal

End Sub


Public Function UsuarioYFecha(Optional ByVal SP As Boolean = True) As String
Dim sTMP As String

If SP Then
    sTMP = "@USUARI = '" & user.USUACODIGO & "'"
    sTMP = sTMP & ", @FECPRO = '" & Format(Now, "dd/MM/yyyy HH:mm") & "'"
Else
    sTMP = "'" & user.USUACODIGO & "', Convert(Datetime,'" & Format(Now, "dd/MM/yyyy HH:mm") & "')"
End If
UsuarioYFecha = sTMP
End Function
