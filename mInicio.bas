Attribute VB_Name = "mInicio"
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

Public Sub main_anterior()

    Dim ini As sgINI
    
    Set ini = New sgINI
    
    ini.FullPath = App.Path & "\Config.ini"
    ini.SectionName = "CONFIGURACION"
    
    SGEPOB = ini.GetString("BASEPOB", SGEPOB)
    PDFS_DEST = ini.GetString("PDFDEST", PDFS_DEST)
    PDFS_ORIG = ini.GetString("PDFORIG", PDFS_ORIG)
    FOX = ini.GetString("DIRFOX", FOX)
    MI_EMAIL = ini.GetString("MAIL", MI_EMAIL)
    
    ASUNTO1 = ini.GetString("ASUNTO", ASUNTO1)
    LINEA1 = ini.GetString("LIN01", LINEA1)
    LINEA2 = ini.GetString("LIN02", LINEA2)
    LINEA3 = ini.GetString("LIN03", LINEA3)
    LINEA4 = ini.GetString("LIN04", LINEA4)
    LINEA5 = ini.GetString("LIN05", LINEA5)
    LINEA6 = ini.GetString("LIN06", LINEA6)
    LINEA7 = ini.GetString("LIN07", LINEA7)
    LINEA8 = ini.GetString("LIN08", LINEA8)
    LINEA9 = ini.GetString("LIN09", LINEA9)
    LINEA10 = ini.GetString("LIN10", LINEA10)
    
    frmMaileador.Show
    
End Sub

