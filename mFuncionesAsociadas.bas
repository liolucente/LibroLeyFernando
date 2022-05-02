Attribute VB_Name = "mFuncionesAsociadas"
Public Sub EnviarMail(Optional sPara As String = "", Optional sAsunto As String = "", Optional sMensaje As String = "", Optional sRuta As String = "")
    Load frmMail
    With frmMail
        .txtPara = sPara
        .txtAdjunto = sRuta
        .txtMensaje = sMensaje
        .txtAsunto = sAsunto
        .Show vbModal
    End With
End Sub

