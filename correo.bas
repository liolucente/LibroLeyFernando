Attribute VB_Name = "Correo"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Enum ConstantesCorreoLibre
    ConstanteCDOPuerto = 2
    ' Usar autentificacion básica.
    ConstanteCDOA_Basica = 1
    'Usar autentificación NTLM
    ConstanteCDO_NTLM = 2    'NTLM
End Enum
'Public Function EnviaCorreo(strPara As String, StrDe As String, _
'                          strAsunto As String, StrCuerpo As String, Optional StrRutaAdjunto As String, _
'                          Optional IntFormato As Integer, Optional Adjunto2 As String, Optional Adjunto3 As String)
'    On Error GoTo MAL_DIR
'
''    Const StrServer As String = "smtp.office365.com"
''    Const StrUsuario As String = "controlnovit@controlseguridad.com.ar" '"facturas@controlseguridad.com.ar"
''    Const StrPassword As String = "Guba3880" ' "Folu6485"
'
''    Const StrServer As String = "mail.controlweb.com.ar" '"smtp.office365.com"
''    Const StrUsuario As String = "facturas@controlweb.com.ar"
''    Const StrPassword As String = "galileo" ' "Folu6485"
'
'    Const StrServer As String = "smtp.office365.com" '"novit-exchange"
'    Const StrUsuario As String = "administracion@mol.com.ar" '"facturas"
'    Const StrPassword As String = "Mercosur820" ' "galileo"
'
'
'    Dim oSmtp As New EASendMailObjLib.Mail
'
'    oSmtp.LicenseCode = "ES-C1407722592-00317-9F7CF1AA2F5FE11F-55319EF76128354U"
'
'    ' Set your sender email address
'    oSmtp.FromAddr = StrDe
'
'    ' Add recipient email address
'    oSmtp.AddRecipientEx strPara, 0
'
''    Para enviar una copia
''    oSmtp.AddRecipientEx strPara, 1
''    Para enviar una copia oculta
''    oSmtp.AddRecipientEx strPara, 2
'
'    ' Set email subject
'    oSmtp.Subject = strAsunto
'
'    ' Your SMTP server address
'    oSmtp.ServerAddr = StrServer
'
'    ' Set TLS connection
'    oSmtp.SSL_init
'    ' Set 25 port, if your server uses 587 port, please change 25 to 587
'    oSmtp.ServerPort = 587
'    oSmtp.SSL_starttls = 1
'
'  'If your server requires SSL connection
'  'oSmtp.SSL_init
'  'oSmtp.SSL_starttls = 0
'  'oSmtp.ServerPort = 465
'
'    ' User and password for ESMTP authentication, if your server doesn't require
'    ' User authentication, please remove the following codes.
'    oSmtp.userName = StrUsuario
'    oSmtp.password = StrPassword
'
'    Select Case IntFormato
'        Case 1
'            'Aqui va el cuerpo sin formato
'            oSmtp.BodyText = StrCuerpo
'        Case 2
'            ' Si quieres formato HTML:
'            ' Formato HTML en el cuerpo del mensaje.
'    '       strHTML = "<html>"
'            oSmtp.BodyFormat = strHTML
'        Case Else
'           oSmtp.BodyText = StrCuerpo
'    End Select
'
'' Para archivo adjunto: No se envia el mail al cliente si la factura no puede ser adjuntada
'    If Len(Trim(StrRutaAdjunto)) <> 0 Then
'        If oSmtp.AddAttachment(StrRutaAdjunto) <> 0 Then
'            frmTrMaileadorRecibos.sgGrilla.Rows.Current.Style.BackColor = vbYellow '&HFF '&H80FF&
'            frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(11).value = "No"
'            frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(7).value = strPara
'            frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(8).value = "-2147024894"
'            frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(9).value = "El sistema no puede hayar el archivo especificado."
'            frmTrMaileadorRecibos.sgRechazados.DataRowCount = frmTrMaileadorRecibos.sgRechazados.DataRowCount + 1
'            With frmTrMaileadorRecibos.sgRechazados.Rows.Last
'                .Cells(1).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(5).value 'legajo
'                .Cells(2).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(7).value 'nombre
'                .Cells(3).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(8).value 'fecha
'                .Cells(4).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(9).value 'comprobante
'                .Cells(5).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(10).value 'mail
'                .Cells(6).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(12).value 'Error
'                .Cells(7).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(13).value 'Numero
'                .Cells(8).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(14).value 'Descripcion
'            End With
'            iError = True
'            DoEvents
'            'Salimos
'            GoTo EnviaCorreoLibre_Exit
'        End If
'    End If
''Esto de acá por si queremos mas adjuntos
'    If Len(Trim(Adjunto2)) <> 0 Then
'        oSmtp.AddAttachments Trim(Adjunto2)
'    End If
'    If Len(Trim(Adjunto3)) <> 0 Then
'        oSmtp.AddAttachments Trim(Adjunto3)
'    End If
'
''Enviamos!
'    If oSmtp.AttachmentCount > 0 Then
'        If oSmtp.SendMail() <> 0 Then GoTo ERR_ENVIANDO
'    End If
'
''Marcamos el renglon como enviado
'    DoEvents
'    frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(11).value = "Si"
'    frmTrMaileadorRecibos.sgGrilla.Rows.Current.Style.BackColor = vbGreen
'    frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(10).value = 15
'    DoEvents
'
''Cargamos el renglon en la grilla que informa los enviados
'    frmTrMaileadorRecibos.sgEnviados.DataRowCount = frmTrMaileadorRecibos.sgEnviados.DataRowCount + 1
'    With frmTrMaileadorRecibos.sgEnviados.Rows.Last
'        .Cells(1).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(5).value 'legajo
'        .Cells(2).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(7).value 'nombre
'        .Cells(3).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(8).value 'fecha
'        .Cells(4).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(9).value 'comprobante
'        .Cells(5).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(10).value 'mail
'        .Cells(6).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(12).value 'Error
'        .Cells(7).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(13).value 'Numero
'        .Cells(8).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(14).value 'Descripcion
'    End With
'    DoEvents
'
'EnviaCorreoLibre_Exit:
'    Exit Function
'
''Manejo de errores MAL_DIR
'MAL_DIR:
'    Dir_Erroneas = Dir_Erroneas & "Cliente: " & frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(5).value & "-" & frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(7).value & vbCrLf & strPara & vbCrLf & "Error: " & Err.Number & vbCrLf & Err.Description & Chr(10) & Chr(13)
'    frmTrMaileadorRecibos.sgGrilla.Rows.Current.Style.BackColor = vbYellow '&HFF '&H80FF&
'    frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(11).value = "No"
'    frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(12).value = strPara
'    frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(13).value = Err.Number
'    frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(14).value = Err.Description
'
''Cargamos el renglon en la grilla que informa los rechazados
'    frmTrMaileadorRecibos.sgRechazados.DataRowCount = frmTrMaileadorRecibos.sgRechazados.DataRowCount + 1
'    With frmTrMaileadorRecibos.sgRechazados.Rows.Last
'        .Cells(1).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(5).value 'legajo
'        .Cells(2).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(7).value 'nombre
'        .Cells(3).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(8).value 'fecha
'        .Cells(4).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(9).value 'comprobante
'        .Cells(5).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(10).value 'mail
'        .Cells(6).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(12).value 'Error
'        .Cells(7).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(13).value 'Numero
'        .Cells(8).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(14).value 'Descripcion
'    End With
'    iError = True
'    Resume EnviaCorreoLibre_Exit
'
''Manejo de Errores ERR_ENVIANDO
'ERR_ENVIANDO:
'
'    Dir_Erroneas = Dir_Erroneas & "Cliente: " & frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(1).value & "-" & frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(2).value & vbCrLf & strPara & vbCrLf & "Error: " & Err.Number & vbCrLf & Err.Description & Chr(10) & Chr(13)
'    frmTrMaileadorRecibos.sgGrilla.Rows.Current.Style.BackColor = vbYellow '&HFF '&H80FF&
'    frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(11).value = "No"
'    frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(12).value = strPara
'    frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(13).value = oSmtp.GetLastError
'    frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(14).value = oSmtp.GetLastErrDescription()
'
''Cargamos el renglon en la grilla que informa los rechazados
'    frmTrMaileadorRecibos.sgRechazados.DataRowCount = frmTrMaileadorRecibos.sgRechazados.DataRowCount + 1
'    With frmTrMaileadorRecibos.sgRechazados.Rows.Last
'        .Cells(1).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(5).value 'legajo
'        .Cells(2).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(7).value 'nombre
'        .Cells(3).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(8).value 'fecha
'        .Cells(4).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(9).value 'comprobante
'        .Cells(5).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(10).value 'mail
'        .Cells(6).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(12).value 'Error
'        .Cells(7).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(13).value 'Numero
'        .Cells(8).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(14).value 'Descripcion
'    End With
'
'    iError = True
'    GoTo EnviaCorreoLibre_Exit
'
'End Function

Public Function EnviarLog(strPara As String, StrDe As String, _
                          strAsunto As String, StrCuerpo As String, Optional StrRutaAdjunto As String, _
                          Optional IntFormato As Integer, Optional Adjunto2 As String, Optional Adjunto3 As String)
    On Error GoTo MAL_DIR
    
    Dim ObjetoMensajeLibre As Object, strHTML As String
    Dim i As Double

    Const StrServer As String = "novit-exchange"
    Const StrUsuario As String = "galileo" '"facturas@controlseguridad.com.ar"
    Const StrPassword As String = "galileo" ' "Folu6485"

'    Const StrServer As String = "smtp.office365.com"
'    Const StrUsuario As String = "controlnovit@controlseguridad.com.ar" '"facturas@controlseguridad.com.ar"
'    Const StrPassword As String = "Guba8038" ' "Folu6485"
 
    Dim oSmtp As New EASendMailObjLib.Mail
 
    oSmtp.LicenseCode = "ES-C1407722592-00317-9F7CF1AA2F5FE11F-55319EF76128354U"
    
    ' Set your sender email address
    oSmtp.FromAddr = StrDe

    ' Add recipient email address
    oSmtp.AddRecipientEx strPara, 0
    
    ' Set email subject
    oSmtp.Subject = strAsunto
    
    ' Your SMTP server address
    oSmtp.ServerAddr = StrServer
   
    ' Set 25 port, if your server uses 587 port, please change 25 to 587
    oSmtp.ServerPort = 587

    ' Set TLS connection

    oSmtp.SSL_starttls = 1

    oSmtp.SSL_init

    ' User and password for ESMTP authentication, if your server doesn't require
    ' User authentication, please remove the following codes.

    oSmtp.userName = StrUsuario

    oSmtp.password = StrPassword

'Ojo con lo del formato.. como cambian los nombres de las propiedades no supe qué ponerle
'salvo en el case select y el case 1
    Select Case IntFormato
        Case 1
            'Aqui va el cuerpo sin formato
            oSmtp.BodyText = StrCuerpo
        Case 2
            ' Si quieres formato HTML:
            ' Formato HTML en el cuerpo del mensaje.
    '       strHTML = "<html>"
            oSmtp.BodyFormat = strHTML
        Case Else
           oSmtp.BodyText = StrCuerpo
    End Select
    
    ' Para archivo adjunto:
    If Len(StrRutaAdjunto) <> 0 Then
        oSmtp.AddAttachment Trim(StrRutaAdjunto)
    End If
    If Len(Trim(Adjunto2)) <> 0 Then
        oSmtp.AddAttachment Trim(Adjunto2)
    End If
    If Len(Trim(Adjunto3)) <> 0 Then
        oSmtp.AddAttachment Trim(Adjunto3)
    End If
    DoEvents
    If oSmtp.SendMail() <> 0 Then GoTo MAL_DIR
    DoEvents

EnviaCorreoLibre_Exit:
    Exit Function
  
MAL_DIR:
    Dir_Erroneas = Dir_Erroneas & "Hubo un problema enviando el log. Mail: " & strPara & vbCrLf & vbCrLf & "Error: " & Err.Number & vbCrLf & Err.Description & Chr(10) & Chr(13)
    iError = True
    Resume EnviaCorreoLibre_Exit
End Function

Public Function EnviaCorreo_SSL(strPara As String, StrDe As String, _
                          strAsunto As String, StrCuerpo As String, Optional StrRutaAdjunto As String, _
                          Optional IntFormato As Integer, Optional Adjunto2 As String, Optional Adjunto3 As String)
    On Error GoTo MAL_DIR
    
    Dim ObjetoMensajeLibre As Object, strHTML As String
    Dim i As Double
    
'Esto se cambió el 15/03/2016 (cambia bastante seguido)
    Const StrServer As String = "190.224.160.81"
    Const StrUsuario As String = "soporte@sistemasgalileo.com.ar"
    Const StrPassword As String = "gali1953"
    
    'Enviando Correo utilizando autentificación remota del server
    ' Creamos el objeto.
   ' On Error GoTo EnviaCorreoLibre_Err
    Set ObjetoMensajeLibre = CreateObject("CDO.Message")

    ' A quien va el mensaje. Para multimples destinos, separa las direcciones de correo por «;»
    ObjetoMensajeLibre.To = "info@sistemasgalileo.com.ar" 'strPara

    'Agregamos copia oculta
    'ObjetoMensajeLibre.Bcc = "carlosfoti@sistemasgalileo.com.ar"
    
    ' Aquí va el asunto
    ObjetoMensajeLibre.Subject = strAsunto

    'Quien envía el correo. ! ojo ! solo funciona con esta sitaxis:
    ' Nombre <tucuenta@tuservidor.com>
    'es decir, por ejemplo: Francisco Javier garcía Aguado<paco@nemo.es>
    ObjetoMensajeLibre.From = "soporte@sistemasgalileo.com.ar" '"Soporte técnico <soporte@sistemasgalileo.com.ar>"
    'Esto se cambió el 26/05/2014
'    ObjetoMensajeLibre.From = "Administracion <sistemasgalileo@novitsa.com.ar>"
    
    'se rechaza si no se recibe una cadena tipo <algo@algo.com>

    Select Case IntFormato

    Case 1
        'Aqui va el cuerpo sin formato
        ObjetoMensajeLibre.TextBody = StrCuerpo
    Case 2
        ' Si quieres formato HTML:
        ' Formato HTML en el cuerpo del mensaje.
'       strHTML = "<html>"
        ObjetoMensajeLibre.HTMLBody = strHTML
    Case Else
        ObjetoMensajeLibre.TextBody = StrCuerpo
    End Select
    ' Para archivo adjunto:
'    If Len(StrRutaAdjunto) <> 0 Then
'        If Dir$(StrRutaAdjunto) <> "" Then
'            ObjetoMensajeLibre.AddAttachment ("file://" & StrRutaAdjunto)
'        Else
'            frmTrMaileadorRecibos.sgGrilla.Rows.Current.Style.BackColor = vbYellow '&HFF '&H80FF&
'            frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(11).value = "No"
'            frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(12).value = strPara
'            frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(13).value = "-2147024894"
'            frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(14).value = "El sistema no puede hayar el archivo especificado."
'            frmTrMaileadorRecibos.sgRechazados.DataRowCount = frmTrMaileadorRecibos.sgRechazados.DataRowCount + 1
'            With frmTrMaileadorRecibos.sgRechazados.Rows.Last
'                .Cells(1).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(5).value 'legajo
'                .Cells(2).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(7).value 'nombre
'                .Cells(3).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(8).value 'fecha
'                .Cells(4).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(9).value 'comprobante
'                .Cells(5).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(10).value 'mail
'                .Cells(6).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(12).value 'Error
'                .Cells(7).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(13).value 'Numero
'                .Cells(8).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(14).value 'Descripcion
'            End With
'            iError = True
'            DoEvents
'            'Salimos
'            GoTo EnviaCorreoLibre_Exit
'        End If
'    End If
    
    If Len(Trim(Adjunto2)) <> 0 Then
        ObjetoMensajeLibre.AddAttachment ("file://" & Trim(Adjunto2))
    End If
    
    If Len(Trim(Adjunto3)) <> 0 Then
        ObjetoMensajeLibre.AddAttachment ("file://" & Trim(Adjunto3))
    End If

    
    '****************************************************
    'Aqui comienzan en si las configuraciones del server de correo remoto
    ' ****************************************************
    ObjetoMensajeLibre.Configuration.Fields.Item( _
            "http://schemas.microsoft.com/cdo/configuration/sendusing") = _
            ConstanteCDOPuerto

    '  Aquí podes poner, bien el literal, bien la IP
    'de la maquina remota del server.
    ObjetoMensajeLibre.Configuration.Fields.Item( _
            "http://schemas.microsoft.com/cdo/configuration/smtpserver") = _
            StrServer

    'Ahora especificamos el tipo de autentificacion que el server espera:
    ObjetoMensajeLibre.Configuration.Fields.Item( _
            "http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = _
            ConstanteCDOA_Basica

    ' Aquí el nombre de usuario que el sever espera recibir
    ObjetoMensajeLibre.Configuration.Fields.Item( _
            "http://schemas.microsoft.com/cdo/configuration/sendusername") = _
            StrUsuario

    ' Aquí el password  del usuario
    ObjetoMensajeLibre.Configuration.Fields.Item( _
            "http://schemas.microsoft.com/cdo/configuration/sendpassword") = _
            StrPassword

    'Aqui el puerto. Casi siempre el 25.... asi que lo dejo a piñon fijo
    ObjetoMensajeLibre.Configuration.Fields.Item( _
            "http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587

    'Si estas usando SSL para la conexión. Valores (False or True)
    ObjetoMensajeLibre.Configuration.Fields.Item( _
            "http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False

    ' Tiempo en segundos para que se establezca correctamente el protocolo SMTP.
    ObjetoMensajeLibre.Configuration.Fields.Item( _
            "http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

    ' Guardamos....
    ObjetoMensajeLibre.Configuration.Fields.Update

    ' Y ! POR FIN !  enviamos
'    MsgBox "1"
    ObjetoMensajeLibre.Send
'    MsgBox "2"
    
    If Err.Number <> 0 Then GoTo MAL_DIR
    
    frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(11).value = "Si"
    frmTrMaileadorRecibos.sgGrilla.Rows.Current.Style.BackColor = vbGreen
    frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(10).value = 15
    
    'Cargamos el renglon en la grilla que informa los enviados
    frmTrMaileadorRecibos.sgEnviados.DataRowCount = frmTrMaileadorRecibos.sgEnviados.DataRowCount + 1
    With frmTrMaileadorRecibos.sgEnviados.Rows.Last
        .Cells(1).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(5).value 'legajo
        .Cells(2).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(7).value 'nombre
        .Cells(3).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(8).value 'fecha
        .Cells(4).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(9).value 'comprobante
        .Cells(5).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(10).value 'mail
        .Cells(6).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(12).value 'Error
        .Cells(7).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(13).value 'Numero
        .Cells(8).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(14).value 'Descripcion
    End With
    
    'esto lo pongo para evitar agregar un timer
'    For I = 1 To 300000000
'    Next
EnviaCorreoLibre_Exit:
    Exit Function
EnviaCorreoLibre_Err:
   MsgBox "Error nº " & Err.Number & vbCrLf & Err.Description & vbCrLf & _
           "en procedimiento EnviaCorreoLibre de Módulo ModuloCdo", vbCritical, "Aviso de error"
    Resume EnviaCorreoLibre_Exit
    
MAL_DIR:

    Dir_Erroneas = Dir_Erroneas & "Cliente: " & frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(5).value & "-" & frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(7).value & vbCrLf & strPara & vbCrLf & "Error: " & Err.Number & vbCrLf & Err.Description & Chr(10) & Chr(13)
    frmTrMaileadorRecibos.sgGrilla.Rows.Current.Style.BackColor = vbYellow '&HFF '&H80FF&
    frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(11).value = "No"
    frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(12).value = strPara
    frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(13).value = Err.Number
    frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(14).value = Err.Description
    iError = True
    
    'Cargamos el renglon en la grilla que informa los rechazados
    frmTrMaileadorRecibos.sgRechazados.DataRowCount = frmTrMaileadorRecibos.sgRechazados.DataRowCount + 1
    With frmTrMaileadorRecibos.sgRechazados.Rows.Last
        .Cells(1).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(5).value 'legajo
        .Cells(2).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(7).value 'nombre
        .Cells(3).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(8).value 'fecha
        .Cells(4).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(9).value 'comprobante
        .Cells(5).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(10).value 'mail
        .Cells(6).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(12).value 'Error
        .Cells(7).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(13).value 'Numero
        .Cells(8).value = frmTrMaileadorRecibos.sgGrilla.Rows.Current.Cells(14).value 'Descripcion
    End With

End Function

Public Function EnviarLog_SSL(strPara As String, StrDe As String, _
                          strAsunto As String, StrCuerpo As String, Optional StrRutaAdjunto As String, _
                          Optional IntFormato As Integer, Optional Adjunto2 As String, Optional Adjunto3 As String)
    On Error GoTo MAL_DIR
    
    Dim ObjetoMensajeLibre As Object, strHTML As String
    Dim i As Double

    Const StrServer As String = "smtp.sistemasgalileo.com.ar"
    Const StrUsuario As String = "soporte@sistemasgalileo.com.ar"
    Const StrPassword As String = "gali1953"

'    Const StrServer As String = "192.168.0.32" '"mail.controlweb.com.ar" '"novit-exchange"
'    Const StrUsuario As String = "galileo@controlweb.com.ar" '"facturas@controlseguridad.com.ar"
'    Const StrPassword As String = "galileo" ' "Folu6485"
    
    
'    Const StrServer As String = "ca3.toservers.com" '"smtp.office365.com"
'    Const StrUsuario As String = "facturas@novitsa.com.ar" '"facturas@seguridadcontrol.com.ar"
'    Const StrPassword As String = "control4970" '"Control.4970"

'Esto se cambió el 26/05/2014
'Const StrUsuario As String = "sistemasgalileo"
'Const StrPassword As String = "4970"
    
    'Enviando Correo utilizando autentificación remota del server
    ' Creamos el objeto.
   ' On Error GoTo EnviaCorreoLibre_Err
    Set ObjetoMensajeLibre = CreateObject("CDO.Message")

    ' A quien va el mensaje. Para multimples destinos, separa las direcciones de correo por «;»
    ObjetoMensajeLibre.To = strPara

    ' Aquí va el asunto
    ObjetoMensajeLibre.Subject = strAsunto

    'Quien envía el correo. ! ojo ! solo funciona con esta sitaxis:
    ' Nombre <tucuenta@tuservidor.com>
    'es decir, por ejemplo: Francisco Javier garcía Aguado<paco@nemo.es>
    ObjetoMensajeLibre.From = "soporte@sistemasgalileo.com.ar"
    'Esto se cambió el 26/05/2014
'    ObjetoMensajeLibre.From = "Administracion <sistemasgalileo@novitsa.com.ar>"
    
    'se rechaza si no se recibe una cadena tipo <algo@algo.com>

    Select Case IntFormato

    Case 1
        'Aqui va el cuerpo sin formato
        ObjetoMensajeLibre.TextBody = StrCuerpo
    Case 2
        ' Si quieres formato HTML:
        ' Formato HTML en el cuerpo del mensaje.
'       strHTML = "<html>"
        ObjetoMensajeLibre.HTMLBody = strHTML
    Case Else
        ObjetoMensajeLibre.TextBody = StrCuerpo
    End Select
    ' Para archivo adjunto:
    If Len(StrRutaAdjunto) <> 0 Then
        ObjetoMensajeLibre.AddAttachment ("file://" & StrRutaAdjunto)
    End If
    If Len(Trim(Adjunto2)) <> 0 Then
        ObjetoMensajeLibre.AddAttachment ("file://" & Trim(Adjunto2))
    End If
    If Len(Trim(Adjunto3)) <> 0 Then
        ObjetoMensajeLibre.AddAttachment ("file://" & Trim(Adjunto3))
    End If

    
    '****************************************************
    'Aqui comienzan en si las configuraciones del server de correo remoto
    ' ****************************************************
    ObjetoMensajeLibre.Configuration.Fields.Item( _
            "http://schemas.microsoft.com/cdo/configuration/sendusing") = _
            ConstanteCDOPuerto

    '  Aquí podes poner, bien el literal, bien la IP
    'de la maquina remota del server.

    ObjetoMensajeLibre.Configuration.Fields.Item( _
            "http://schemas.microsoft.com/cdo/configuration/smtpserver") = _
            StrServer

    'Ahora especificamos el tipo de autentificacion que el server espera:

    ObjetoMensajeLibre.Configuration.Fields.Item( _
            "http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = _
            ConstanteCDOA_Basica

    ' Aquí el nombre de usuario que el sever espera recibir

    ObjetoMensajeLibre.Configuration.Fields.Item( _
            "http://schemas.microsoft.com/cdo/configuration/sendusername") = _
            StrUsuario

    ' Aquí el password  del usuario

    ObjetoMensajeLibre.Configuration.Fields.Item( _
            "http://schemas.microsoft.com/cdo/configuration/sendpassword") = _
            StrPassword

    'Aqui el puerto. Casi siempre el 25.... asi que lo dejo a piñon fijo
    ObjetoMensajeLibre.Configuration.Fields.Item( _
            "http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587

    'Si estas usando SSL para la conexión. Valores (False or True)
    ObjetoMensajeLibre.Configuration.Fields.Item( _
            "http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False

    ' Tiempo en segundos para que se establezca correctamente el protocolo SMTP.
    ObjetoMensajeLibre.Configuration.Fields.Item( _
            "http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

    ' Guardamos....
    ObjetoMensajeLibre.Configuration.Fields.Update

    ' Y ! POR FIN !  enviamos
'    MsgBox "3"
    ObjetoMensajeLibre.Send
'    MsgBox "4"

    
    If Err.Number <> 0 Then GoTo MAL_DIR
    
EnviaCorreoLibre_Exit:
    Exit Function
EnviaCorreoLibre_Err:
   MsgBox "Error nº " & Err.Number & vbCrLf & Err.Description & vbCrLf & _
           "en procedimiento EnviaCorreoLibre de Módulo ModuloCdo", vbCritical, "Aviso de error"
    Resume EnviaCorreoLibre_Exit
    
MAL_DIR:
    Dir_Erroneas = Dir_Erroneas & "Hubo un problema enviando el log. Mail: " & strPara & vbCrLf & vbCrLf & "Error: " & Err.Number & vbCrLf & Err.Description & Chr(10) & Chr(13)
    iError = True
    Resume EnviaCorreoLibre_Exit
End Function

