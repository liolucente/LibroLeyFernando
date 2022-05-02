Attribute VB_Name = "mManejoDeErrores"
Option Explicit
Public Const ErrorNoDefinido = 9999

Private Function ObtenerTextoDeError(NumError As Long) As String
      On Error GoTo ObtenerTextoDeErrorError
      
      ObtenerTextoDeError = LoadResString(NumError)

      Exit Function
ObtenerTextoDeErrorError:
      If Err.Number <> 0 Then
            ObtenerTextoDeError = "Ha ocurrido un error inesperado!" & vbCrLf & "Comuniquese con mantenimiento."
      End If
End Function

Public Sub MostrarError(NumeroError As Long, Origen As String)
      Dim sTextoError As String

      sTextoError = ObtenerTextoDeError(NumeroError)

      Err.Raise NumeroError, Origen, sTextoError
End Sub

Public Sub MostrarErrorDeTransacciones(ers As ADODB.Errors, Optional sAccion As String)
Dim e As ADODB.error
Dim sMSG As String
Dim FSO As FileSystemObject
Dim tsLOG As Scripting.TextStream

sMSG = "Ha" & IIf(ers.Count > 1, "n ", " ") & "ocurrido" & IIf(ers.Count = 1, " un error ", " errores ") & " " & IIf(sAccion = "", " al intentar acceder a la base", sAccion) & "."

For Each e In ers
    sMSG = sMSG & vbCrLf & "Error numero: " & e.Number & " (en:" & e.source & ")" & vbCrLf & "Descripcion: " & e.Description & " (Estado de SQL:" & e.SQLState & ")"
Next

Set FSO = New FileSystemObject
If FSO.FileExists(App.Path & "\Galileo.err") Then
    Set tsLOG = FSO.OpenTextFile(App.Path & "\Galileo.err", ForAppending, True)
Else
    Set tsLOG = FSO.OpenTextFile(App.Path & "\Galileo.err", ForWriting, True)
End If

tsLOG.WriteBlankLines 1
tsLOG.WriteLine "[" & Format(Date, "dd/mm/yyyy") & "-" & Format(Time, "hh:mm:ss") & "]"
tsLOG.WriteLine "USUARIO=" & user.USUACODIGO & " (" & user.USUANOMBRE & ")"
tsLOG.WriteLine "MENSAJE=" & Replace(sMSG, vbCrLf, " - ", 1, , vbTextCompare)

tsLOG.Close
Set tsLOG = Nothing
Set FSO = Nothing

MsgBox sMSG, vbOKOnly, "Errores " & IIf(sAccion = "", " al acceder a la base", sAccion) & "."

End Sub

Public Sub GuardarLOG(ByVal sTexto As String, Optional ByVal sArchivo As String = "NONE,JA!")
Dim FSO As FileSystemObject
Dim tsLOG As Scripting.TextStream


Set FSO = New FileSystemObject
If FSO.FileExists(App.Path & "\Galileo.Err") Then
    Set tsLOG = FSO.OpenTextFile(App.Path & "\Galileo.Err", ForAppending, True)
Else
    Set tsLOG = FSO.OpenTextFile(App.Path & "\Galileo.Err", ForWriting, True)
End If

tsLOG.WriteBlankLines 1
tsLOG.WriteLine "[" & Format(Date, "dd/mm/yyyy") & "-" & Format(Time, "hh:mm:ss") & "]"
tsLOG.WriteLine "USUARIO=" & user.USUACODIGO & " (" & user.USUANOMBRE & ")"
tsLOG.WriteLine "MENSAJE=" & Replace(sTexto, vbCrLf, " - ", 1, , vbTextCompare)

tsLOG.Close
Set tsLOG = Nothing
Set FSO = Nothing

End Sub

Public Sub GuardarLOGError(ByVal e As VBA.ErrObject, Optional ByVal sArchivo As String = "NONE,JA!")

    Dim FSO As FileSystemObject
    Dim tsLOG As Scripting.TextStream

    Set FSO = New FileSystemObject
    If FSO.FileExists(App.Path & "\Galileo.Err") Then
        Set tsLOG = FSO.OpenTextFile(App.Path & "\Galileo.Err", ForAppending, True)
    Else
        Set tsLOG = FSO.OpenTextFile(App.Path & "\Galileo.Err", ForWriting, True)
    End If

    tsLOG.WriteBlankLines 1
    tsLOG.WriteLine "[" & Format(Date, "dd/mm/yyyy") & "-" & Format(Time, "hh:mm:ss") & "]"
    tsLOG.WriteLine "USUARIO=" & user.USUACODIGO & " (" & user.USUANOMBRE & ")"
    tsLOG.WriteLine "MENSAJE=" & Err.Description & " (" & Err.Number & "-" & Err.source & ")"

    tsLOG.Close
    Set tsLOG = Nothing
    Set FSO = Nothing

End Sub

Public Sub GuardarLOGError_SU(ByVal e As VBA.ErrObject, Optional ByVal sArchivo As String = "NONE,JA!")

    Dim FSO As FileSystemObject
    Dim tsLOG As Scripting.TextStream

    Set FSO = New FileSystemObject
    If FSO.FileExists(App.Path & "\Galileo.Err") Then
        Set tsLOG = FSO.OpenTextFile(App.Path & "\Galileo.Err", ForAppending, True)
    Else
        Set tsLOG = FSO.OpenTextFile(App.Path & "\Galileo.Err", ForWriting, True)
    End If

    tsLOG.WriteBlankLines 1
    tsLOG.WriteLine "[" & Format(Date, "dd/mm/yyyy") & "-" & Format(Time, "hh:mm:ss") & "]"
    tsLOG.WriteLine "USUARIO=SIN USUARIO"
    tsLOG.WriteLine "MENSAJE=" & Err.Description & " (" & Err.Number & "-" & Err.source & ")"

    tsLOG.Close
    Set tsLOG = Nothing
    Set FSO = Nothing

End Sub

Public Sub GuardarLOGErrorInicio(ByVal e As VBA.ErrObject, Optional ByVal sArchivo As String = "NONE,JA!")
Dim FSO As FileSystemObject
Dim tsLOG As Scripting.TextStream


Set FSO = New FileSystemObject
If FSO.FileExists(App.Path & "\Galileo.Err") Then
    Set tsLOG = FSO.OpenTextFile(App.Path & "\Galileo.Err", ForAppending, True)
Else
    Set tsLOG = FSO.OpenTextFile(App.Path & "\Galileo.Err", ForWriting, True)
End If

tsLOG.WriteBlankLines 1
tsLOG.WriteLine "[" & Format(Date, "dd/mm/yyyy") & "-" & Format(Time, "hh:mm:ss") & "]"
tsLOG.WriteLine "USUARIO=SIN USUARIO"
tsLOG.WriteLine "MENSAJE=" & Err.Description & " (" & Err.Number & "-" & Err.source & ")"

tsLOG.Close
Set tsLOG = Nothing
Set FSO = Nothing
End Sub

Public Sub GuardarErrorTransacciones(ers As ADODB.Errors, Optional sAccion As String)

    Dim e As ADODB.error
    Dim sMSG As String
    Dim FSO As FileSystemObject
    Dim tsLOG As Scripting.TextStream

    sMSG = "Ha" & IIf(ers.Count > 1, "n ", " ") & "ocurrido" & IIf(ers.Count = 1, " un error ", " errores ") & " " & IIf(sAccion = "", " al intentar acceder a la base", sAccion) & "."

    For Each e In ers
        sMSG = sMSG & vbCrLf & "Error numero: " & e.Number & " (en:" & e.source & ")" & vbCrLf & "Descripcion: " & e.Description & " (Estado de SQL:" & e.SQLState & ")"
    Next

    Set FSO = New FileSystemObject
    If FSO.FileExists(App.Path & "\Galileo.err") Then
        Set tsLOG = FSO.OpenTextFile(App.Path & "\Galileo.err", ForAppending, True)
    Else
        Set tsLOG = FSO.OpenTextFile(App.Path & "\Galileo.err", ForWriting, True)
    End If

    tsLOG.WriteBlankLines 1
    tsLOG.WriteLine "[" & Format(Date, "dd/mm/yyyy") & "-" & Format(Time, "hh:mm:ss") & "]"
    tsLOG.WriteLine "USUARIO=" & user.USUACODIGO & " (" & user.USUANOMBRE & ")"
    tsLOG.WriteLine "MENSAJE=" & Replace(sMSG, vbCrLf, " - ", 1, , vbTextCompare)

    tsLOG.Close
    Set tsLOG = Nothing
    Set FSO = Nothing

'    MsgBox sMSG, vbOKOnly, "Errores " & IIf(sAccion = "", " al acceder a la base", sAccion) & "."

End Sub
