Attribute VB_Name = "mGenerales"

Option Explicit


'Declaracione para la funcon PonerVentanaPrimero
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
'                                               '

'Declaracione para la funcion SoloNumeros       '
Const ES_NUMBER = &H2000
Const GWL_STYLE = -16
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'                                               '

Public Const Gris = &HE0E0E0
Public Const Rojo = &HFF&
Public Const blanco = &H80000005

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function OemToChar Lib "user32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Const WS_EX_TRANSPARENT = &H20&
Declare Function PrinterProperties Lib "winspool.drv" (ByVal hWnd As Long, ByVal hPrinter As Long) As Long

'Colecciones de formularios de los sistemas
Public colCONFIG As Collection
Public colREPARA As Collection
Public colPADRONES As Collection
Public colOSPIN As Collection
Public colSUTNA As Collection
'                                         '

Public sTextoBusqueda As Variant
Public sCompaBusqueda As String

Public Enum t_ABMmodo
    Alta = 1
    Baja = 2
    Consulta = 3
    Modificacion = 4
End Enum

Public Sub SoloNumeros(ByVal hWnd As Long)
Dim V As Long
V = GetWindowLong(hWnd, GWL_STYLE)
SetWindowLong hWnd, GWL_STYLE, V + ES_NUMBER
End Sub

Public Sub PonerVentanaPrimero(ByVal hWnd As Long, Optional ByVal QueTodas As Boolean = True)
If QueTodas Then
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_SHOWWINDOW
Else
    BringWindowToTop hWnd
End If
End Sub

'##ModelId=3D3873F6038E
Public Sub Transparente(ByVal frmhWnd As Long)

'*************************************************'
'                                                 '
' Procedimiento que hace transparente un Form     '
'                                                 '
'  Hecho por: Fernando (Chicho)                   '
'  Fecha de Creacion: 26/12/2001                  '
'                                                 '
'*************************************************'


SetWindowLong frmhWnd, GWL_STYLE, WS_EX_TRANSPARENT

End Sub

'##ModelId=3D3873F7019A
Public Function Encripta(ByVal sTexto As String) As String

'***********************************************'
'                                               '
'  Funcion para encriptar datos pasados         '
'  a la funcion por medio de sTexto.            '
'                                               '
'  lPosicion es la variable para el For que     '
'  recorre el string a encriptar                '
'                                               '
'  CharToOem es una API para pasar caracteres   '
'  Windows a DOS.                               '
'                                               '
'  lPosicion se vuelve a uno cuando llega a 15  '
'  para que no genere errores.                  '
'                                               '
'  Hecho por: Fernando (Chicho)                 '
'  (Sacado del sistema de Fox, se cambio        '
'  el 17 que se suma en la formula por un 5)    '
'  Fecha de creacion: 30/11/2001                '
'                                               '
'***********************************************'

Dim lPos As Long
Dim lPosicion As Long
Dim sCaracter As String
lPosicion = 0

For lPos = 1 To Len(sTexto)
    
    If lPosicion = 15 Then
        lPosicion = 1
    Else
        lPosicion = lPosicion + 1
    End If
    
    sCaracter = Space(1)
    CharToOem Mid(sTexto, lPos, 1), sCaracter
    
    Encripta = Encripta & Chr(Asc(sCaracter) - 1 * (lPosicion * (1 + lPosicion Mod 2) + 5))

Next lPos

End Function

'##ModelId=3D3873F70352
Public Function DesEncripta(ByVal sTexto As String) As String

'***********************************************'
'                                               '
'  Funcion para desencriptar datos pasados      '
'  a la funcion por medio de sTexto.            '
'                                               '
'  lPos es la variable para el For que          '
'  recorre el string a encriptar                '
'                                               '
'  CharToOem es una API para pasar caracteres   '
'  Windows a DOS.                               '
'                                               '
'  lPosicion se vuelve a uno cuando llega a 15  '
'  para que no genere errores.                  '
'                                               '
'  Hecho por: Fernando (Chicho)                 '
'  (Sacado del sistema de Fox, se cambio        '
'  el 17 que se suma en la formula por un 5)    '
'  Fecha de creacion: 30/11/2001                '
'                                               '
'***********************************************'

Dim lPos As Long
Dim lPosicion As Long
Dim sCaracter As String

lPosicion = 0

For lPos = 1 To Len(sTexto)
    
    If lPosicion = 15 Then
        lPosicion = 1
    Else
        lPosicion = lPosicion + 1
    End If
    
    sCaracter = Space(1)
    OemToChar Mid(sTexto, lPos, 1), sCaracter

    DesEncripta = DesEncripta & Chr(Asc(sCaracter) + 1 * (lPosicion * (1 + lPosicion Mod 2) + 5))

Next lPos

End Function

'##ModelId=3D3873F80122
Public Function PasarControl()

'*************************************************'
'                                                 '
' Funcion que simula un TAB.                      '
' Sirve para pasar de control al apretar una tecla'
' Ej. al apretar enter en una caja de texto.      '
'                                                 '
'  Hecho por: Fernando (Chicho)                   '
'  Fecha de Creacion: 03/12/2001                  '
'                                                 '
'*************************************************'

SendKeys "{TAB}"

End Function

Public Function LlenarIN(ByRef adoRec As Recordset, ByVal sCampoValor As String, Optional ByVal Formato As String = "NN") As String
Dim i As Long
Dim sTMP As String

'*************************************************'
'                                                 '
'  Funcion que devuelve un string con un "IN()" o '
'  un "= X", para una sentecia de SQL             '
'                                                 '
'  Hecho por: Fernando (Chicho)                   '
'  Fecha de Creacion: 04/12/2001                  '
'                                                 '
'*************************************************'

If adoRec.RecordCount = 0 Then
    LlenarIN = ""
    Exit Function
End If

adoRec.MoveFirst
If adoRec.RecordCount = 1 Then
    sTMP = ConvertirDato(adoRec.Fields(sCampoValor).value)
    LlenarIN = "=" & IIf(Formato = "NN", sTMP, Format(sTMP, Formato))
Else
    LlenarIN = " IN("
    While Not adoRec.EOF
        sTMP = ConvertirDato(adoRec.Fields(sCampoValor).value)
        LlenarIN = LlenarIN & IIf(Formato = "NN", sTMP, Format(sTMP, Formato)) & ", "
        adoRec.MoveNext
    Wend
    LlenarIN = Left(LlenarIN, Len(LlenarIN) - 2) & ")"
End If
adoRec.MoveFirst

If Right(LlenarIN, 2) = ", " Then
    LlenarIN = Left(LlenarIN, Len(LlenarIN) - 2) & ")"
End If

End Function

Public Function LlenarINCR(ByRef adoRec As Recordset, ByVal sCampoValor As String) As String
Dim i As Long

'*************************************************'
'                                                 '
'  Funcion que devuelve un string con un "IN()" o '
'  un "= X", para una sentecia de SQL             '
'                                                 '
'  Hecho por: Fernando (Chicho)                   '
'  Fecha de Creacion: 04/12/2001                  '
'                                                 '
'*************************************************'

If adoRec.RecordCount = 0 Then
    LlenarINCR = ""
    Exit Function
End If

adoRec.MoveFirst
If adoRec.RecordCount = 1 Then
    LlenarINCR = "{" & adoRec.Fields(sCampoValor).Properties(1).value & "." & adoRec.Fields(sCampoValor).name & "} =" & ConvertirDatoCR(adoRec.Fields(sCampoValor).value)
Else
    LlenarINCR = "{" & adoRec.Fields(sCampoValor).Properties(1).value & "." & adoRec.Fields(sCampoValor).name & "} IN["
    For i = 1 To adoRec.RecordCount
        If i = adoRec.RecordCount Then
            LlenarINCR = LlenarINCR & ConvertirDatoCR(adoRec.Fields(sCampoValor).value) & "]"
        Else
            LlenarINCR = LlenarINCR & ConvertirDatoCR(adoRec.Fields(sCampoValor).value) & ", "
        End If
        adoRec.MoveNext
    Next i
End If
adoRec.MoveFirst

If Right(LlenarINCR, 2) = ", " Then
    LlenarINCR = Left(LlenarINCR, Len(LlenarINCR) - 2) & "]"
End If

End Function

Public Function ConvertirDatoCR(ByVal sDato As String) As String

'*************************************************'
'                                                 '
'  Esta funcion acomoda el string segun el tipo   '
'  de dato que recibe. Solo para Crystal Reports  '
'                                                 '
'  Hecho por: Fernando (Chicho)                   '
'  Fecha de Creacion: 29/07/2002                  '
'                                                 '
'*************************************************'

If IsDate(sDato) Then
    If EsHora(sDato) Then
        ConvertirDatoCR = HMScr(sDato)
        Exit Function
    Else
        ConvertirDatoCR = DMAcr(sDato)
        Exit Function
    End If
ElseIf IsNumeric(sDato) Then
    ConvertirDatoCR = sDato
Else
    ConvertirDatoCR = " """ & sDato & """"
End If

End Function

Public Function HMScr(ByVal sH As String) As String

'*************************************************'
'                                                 '
'  Esta funcion acomoda el sH como una hora.      '
'                                                 '
'  Hecho por: Fernando (Chicho)                   '
'  Fecha de Creacion: 29/07/2002                  '
'                                                 '
'*************************************************'

sH = Left(sH, 2) & "," & Mid(sH, 4, 2) & "," & Right(sH, 2)

HMScr = " TimeSeial(" & sH & ") "

End Function

Public Function DMAcr(ByVal sDMA As String) As String

'*************************************************'
'                                                 '
'  Esta funcion acomoda el sDMA como un dia.      '
'                                                 '
'  Hecho por: Fernando (Chicho)                   '
'  Fecha de Creacion: 29/07/2002                  '
'                                                 '
'*************************************************'

sDMA = Left(sDMA, 2) & "," & Mid(sDMA, 4, 2) & "," & Right(sDMA, 4)

DMAcr = "DateSerial(" & sDMA & ")"

End Function

Public Function ConvertirDato(ByVal sDato As String) As String

'*************************************************'
'                                                 '
'  Esta funcion acomoda el string segun el tipo   '
'  de dato que recibe.                            '
'                                                 '
'  Hecho por: Fernando (Chicho)                   '
'  Fecha de Creacion: 04/12/2001                  '
'                                                 '
'*************************************************'

If IsDate(sDato) Then
    If EsHora(sDato) Then
        ConvertirDato = " #" & HMS(sDato) & "#"
        Exit Function
    Else
        ConvertirDato = " '" & DMA(sDato) & "'"
        Exit Function
    End If
ElseIf IsNumeric(sDato) Then
    ConvertirDato = sDato
Else
    ConvertirDato = " '" & sDato & "'"
End If

End Function

Public Function UltimoDiaDe(ByVal mes As String, ByVal Anio As String) As Integer

mes = Trim(mes)
Anio = Trim(Anio)

Dim Fecha As Date
Fecha = "01/" & mes & "/" & Anio
While Month(Fecha) = Val(mes)
    Fecha = Fecha + 1
Wend
UltimoDiaDe = Day((Fecha - 1))

End Function

'##ModelId=3D3873F902C6
Public Function EsHora(ByVal sH As String) As Boolean

'*************************************************'
'                                                 '
'  Esta funcion verifica si sH es una hora valda  '
'                                                 '
'  Hecho por: Fernando (Chicho)                   '
'  Fecha de Creacion: 04/12/2001                  '
'                                                 '
'*************************************************'

EsHora = IIf(InStr(1, sH, ":") = 0, False, True)

End Function

Public Function HMS(ByVal sH As String) As String

'*************************************************'
'                                                 '
'  Esta funcion acomoda el sH como una hora.      '
'                                                 '
'  Hecho por: Fernando (Chicho)                   '
'  Fecha de Creacion: 04/12/2001                  '
'                                                 '
'*************************************************'

sH = TimeValue(sH)

'Lo comentado sirve para timeSerial (Crytal Reports)
'HMS = Left(sH, 2) & "," & Mid(sH, 4, 2) & "," & Right(sH, 2)

HMS = sH

End Function

Public Function DMA(ByVal sDMA As String) As String

'*************************************************'
'                                                 '
'  Esta funcion acomoda el sDMA como un dia.      '
'                                                 '
'  Hecho por: Fernando (Chicho)                   '
'  Fecha de Creacion: 04/12/2001                  '
'                                                 '
'*************************************************'

sDMA = CDate(sDMA)
sDMA = Left(sDMA, 10)

'Lo comentado sirve para DateSerial (Crytal Reports)
'DMA = Left(sDMA, 2) & "," & Mid(sDMA, 4, 2) & "," & Right(sDMA, 4)

DMA = sDMA

End Function

'##ModelId=3D3873FB005A
Public Sub Oscurecer(ByRef txt As TextBox)

txt.SelStart = 0
txt.SelLength = Len(txt.Text)

End Sub

'##ModelId=3D3873FB0244
Public Function EstaAbierta(ByVal WindowHWND As Long) As Boolean
'If IsWindowEnabled(WindowHWND) Then
    If IsWindowVisible(WindowHWND) Then
        EstaAbierta = True
        Exit Function
    End If
'End If
EstaAbierta = False
End Function

'##ModelId=3D3873FC0050
Public Function ObtenerMenu(ByVal frm As Form, ByVal sMenu As String) As Menu

'**************************************************'
'                                                  '
'  Funcion que devuelve un menu con el nombre      '
'  indicado en sMenu.                              '
'                                                  '
'  Hecho por: Fernando (Chicho)                    '
'  Fecha de creacion: 17/12/2001                   '
'                                                  '
'**************************************************'
Dim lCTRL As Long
On Error Resume Next

For lCTRL = 0 To frm.Controls.Count - 1
    If UCase(frm.Controls(lCTRL).name) = UCase(sMenu) And _
        TypeOf frm.Controls(lCTRL) Is Menu Then
            Set ObtenerMenu = frm.Controls(lCTRL)
        Exit Function
    End If
Next

Set ObtenerMenu = Nothing

End Function

'##ModelId=3D3873FC0384
Public Sub CargarIMGCombo(ByVal cbo As ImageCombo, ByVal vIMG As Variant, ByVal adoRec As Recordset, ByVal sCampoMostrar As String, Optional ByVal sCampoTag As String = "", Optional ByVal sKey As String)

'**************************************************'
'                                                  '
'  Este procedimiento llena un ImageCombo con adoRec
'  mostrando sCampoMostrar y con la imagen iIMG.   '
'                                                  '
'  Agrege (13/03/2002) el sCampoTag.               '
'                                                  '
'  Agrege el sKey (03/04/2002)                     '
'                                                  '
'**************************************************'

Dim cboITM As ComboItem

Screen.MousePointer = vbHourglass

cbo.ComboItems.Clear
On Error GoTo errLOLO
If Not adoRec.RecordCount = 0 Then
    adoRec.MoveFirst
End If
OtroERR:
On Error Resume Next
While Not adoRec.EOF
    Set cboITM = cbo.ComboItems.Add(, , Trim(adoRec.Fields(sCampoMostrar).value))
    If Not vIMG = "-1" Then cboITM.Image = vIMG
    If Not sCampoTag = "" Then
        cboITM.Tag = adoRec.Fields(sCampoTag).value
    End If
    If Not sKey = "" Then
        If IsNumeric(adoRec.Fields(sKey).value) Then
            cboITM.key = adoRec.Fields(sKey).value & " n"
        Else
            cboITM.key = adoRec.Fields(sKey).value
        End If
    End If
    adoRec.MoveNext
Wend

On Error Resume Next
adoRec.MoveFirst
Set cboITM = Nothing

If Not cbo.ComboItems.Count = 0 Then
    cbo.ComboItems(1).Selected = True
End If

Screen.MousePointer = vbNormal

Exit Sub
errLOLO:
    If Err.Number = 91 Or Err.Number = 3704 Then
        Screen.MousePointer = vbNormal
        Exit Sub
    Else
        Resume OtroERR
    End If

End Sub

Public Sub CargarSGCombo(ByVal cbo As Object, ByVal vIMG As Variant, ByVal adoRec As Recordset, ByVal sCampoMostrar As String, Optional ByVal sCampoTag As String = "", Optional ByVal sKey As String)

'**************************************************'
'                                                  '
'  Este procedimiento llena un ImageCombo con adoRec
'  mostrando sCampoMostrar y con la imagen iIMG.   '
'                                                  '
'  Agrege (13/03/2002) el sCampoTag.               '
'                                                  '
'  Agrege el sKey (03/04/2002)                     '
'                                                  '
'**************************************************'

Dim cboITM As ComboItem

Screen.MousePointer = vbHourglass

cbo.Clear
On Error Resume Next
adoRec.MoveFirst
While Not adoRec.EOF
    Set cboITM = cbo.Add(, , adoRec.Fields(sCampoMostrar).value)
    If Not vIMG = "-1" Then cboITM.Image = vIMG
    If Not sCampoTag = "" Then
        cboITM.Tag = adoRec.Fields(sCampoTag).value
    End If
    If Not sKey = "" Then
        If IsNumeric(adoRec.Fields(sKey).value) Then
            cboITM.key = adoRec.Fields(sKey).value & " n"
        Else
            cboITM.key = adoRec.Fields(sKey).value
        End If
    End If
    adoRec.MoveNext
Wend

On Error Resume Next
adoRec.MoveFirst
Set cboITM = Nothing

If Not cbo.ComboItems.Count = 0 Then
    cbo(1).Selected = True
End If

Screen.MousePointer = vbNormal

End Sub

'##ModelId=3D3873FE0334
Public Sub CargarListView(ByVal lvw As ListView, ByVal adoRec As Recordset, Optional ByVal lvwSinEstos As ListView, Optional ByVal iIMG As Integer = -1)

'**************************************************'
'                                                  '
'  Este procedimiento carga los datos de adoRec    '
'  en un listView segun el Tag sus columnas. Si se '
'  especifica un ListView en lvwsinEstos, se cargan'
'  en lvw solo los datos que no estan en este y si '
'  iIMG es diferente a -1 se carga la imagen       '
'  especificada.                                   '
'                                                  '
'**************************************************'

Dim lvwITM As ListItem
Dim lCols As Long

lvw.ListItems.Clear
If adoRec.RecordCount = 0 Then Exit Sub
If adoRec.EOF And adoRec.BOF Then Exit Sub

While Not adoRec.EOF
On Error GoTo NoEsta
    If lvwSinEstos Is Nothing Then
SeGUiR:
        lCols = 1
        If InStr(lvw.ColumnHeaders(lCols).Tag, ",") > 0 Then
            If Not IsNull(adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, InStr(lvw.ColumnHeaders(lCols).Tag, ",") - 1))) Then
                Set lvwITM = lvw.ListItems.Add(, , adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, InStr(lvw.ColumnHeaders(lCols).Tag, ",") - 1)))
            Else
                Set lvwITM = lvw.ListItems.Add(, , IIf(IsNull(adoRec.Fields(Mid(lvw.ColumnHeaders(lCols).Tag, InStr(lvw.ColumnHeaders(lCols).Tag, ",") + 1))), "", adoRec.Fields(Mid(lvw.ColumnHeaders(lCols).Tag, InStr(lvw.ColumnHeaders(lCols).Tag, ",") + 1))))
            End If
        Else
            If Right(lvw.ColumnHeaders(lCols).Tag, 1) = "F" Then
                If IsNumeric(adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, Len(lvw.ColumnHeaders(lCols).Tag) - 2)).value) Then
                    If adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, Len(lvw.ColumnHeaders(lCols).Tag) - 2)).value > 0 Then
                        Set lvwITM = lvw.ListItems.Add(, , CDate(ConvertirFecha(CLng(adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, Len(lvw.ColumnHeaders(lCols).Tag) - 2))))))
                    Else
                        Set lvwITM = lvw.ListItems.Add(, , "")
                    End If
                Else
                    Set lvwITM = lvw.ListItems.Add(, , adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag & "", Len(lvw.ColumnHeaders(lCols).Tag & "") - 2)))
                End If
            Else
                If InStr(lvw.ColumnHeaders(lCols).Tag, ".N") > 0 Then
                    Set lvwITM = lvw.ListItems.Add(, , Right(String(Val(Right(lvw.ColumnHeaders(lCols).Tag, 2)), "0") & adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, Len(lvw.ColumnHeaders(lCols).Tag) - 4)), Val(Right(lvw.ColumnHeaders(lCols).Tag, 2))))
                Else
                    If InStr(lvw.ColumnHeaders(lCols).Tag, ".$") > 0 Then
                        Set lvwITM = lvw.ListItems.Add(, , Format(adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, Len(lvw.ColumnHeaders(lCols).Tag) - 2)), "Currency"))
                    Else
                        Set lvwITM = lvw.ListItems.Add(, , adoRec.Fields(lvw.ColumnHeaders(lCols).Tag) & "")
                    End If
                End If
            End If
        End If
        
        For lCols = 2 To lvw.ColumnHeaders.Count
            If InStr(lvw.ColumnHeaders(lCols).Tag, ",") > 0 Then
                If Not IsNull(adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, InStr(lvw.ColumnHeaders(lCols).Tag, ",") - 1))) Then
                    lvwITM.SubItems(lCols - 1) = adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, InStr(lvw.ColumnHeaders(lCols).Tag, ",") - 1))
                Else
                    lvwITM.SubItems(lCols - 1) = IIf(IsNull(adoRec.Fields(Mid(lvw.ColumnHeaders(lCols).Tag, InStr(lvw.ColumnHeaders(lCols).Tag, ",") + 1))), "", adoRec.Fields(Mid(lvw.ColumnHeaders(lCols).Tag, InStr(lvw.ColumnHeaders(lCols).Tag, ",") + 1)))
                End If
            Else
                If Right(lvw.ColumnHeaders(lCols).Tag, 1) = "F" Then
                    If IsNumeric(adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, Len(lvw.ColumnHeaders(lCols).Tag) - 2)).value) Then
                        If adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, Len(lvw.ColumnHeaders(lCols).Tag) - 2)).value > 0 Then
                            lvwITM.SubItems(lCols - 1) = CDate(ConvertirFecha(CLng(adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, Len(lvw.ColumnHeaders(lCols).Tag) - 2)))))
                        Else
                            lvwITM.SubItems(lCols - 1) = ""
                        End If
                    Else
                        lvwITM.SubItems(lCols - 1) = adoRec.Fields(Replace(lvw.ColumnHeaders(lCols).Tag, ".F", vbNullString))
                    End If
                Else
                    If InStr(lvw.ColumnHeaders(lCols).Tag, ".N") > 0 Then
                        lvwITM.SubItems(lCols - 1) = Right(String(Val(Right(lvw.ColumnHeaders(lCols).Tag, 2)), "0") & adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, Len(lvw.ColumnHeaders(lCols).Tag) - 4)), Val(Right(lvw.ColumnHeaders(lCols).Tag, 2)))
                    Else
                        If InStr(lvw.ColumnHeaders(lCols).Tag, ".$") > 0 Then
                            lvwITM.SubItems(lCols - 1) = Format(adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, Len(lvw.ColumnHeaders(lCols).Tag) - 2)), "Currency")
                        Else
                            lvwITM.SubItems(lCols - 1) = adoRec.Fields(lvw.ColumnHeaders(lCols).Tag) & ""
                        End If
                    End If
                End If
            End If
        Next lCols
        On Error Resume Next
        If Not iIMG = -1 Then
            lvwITM.Icon = iIMG
            lvwITM.SmallIcon = iIMG
        End If
    Else
        If Not EstaEnLaLista(lvwSinEstos, adoRec.Fields(lvw.ColumnHeaders(1).Tag) & "") Then
            lCols = 1
            If InStr(lvw.ColumnHeaders(lCols).Tag, ",") > 0 Then
                If Not IsNull(adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, InStr(lvw.ColumnHeaders(lCols).Tag, ",") - 1))) Then
                    Set lvwITM = lvw.ListItems.Add(, , adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, InStr(lvw.ColumnHeaders(lCols).Tag, ",") - 1)))
                Else
                    Set lvwITM = lvw.ListItems.Add(, , IIf(IsNull(adoRec.Fields(Mid(lvw.ColumnHeaders(lCols).Tag, InStr(lvw.ColumnHeaders(lCols).Tag, ",") + 1))), "", adoRec.Fields(Mid(lvw.ColumnHeaders(lCols).Tag, InStr(lvw.ColumnHeaders(lCols).Tag, ",") + 1))))
                End If
            Else
                If Right(lvw.ColumnHeaders(lCols).Tag, 1) = "F" Then
                    If IsNumeric(adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, Len(lvw.ColumnHeaders(lCols).Tag) - 2)).value) Then
                       If adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, Len(lvw.ColumnHeaders(lCols).Tag) - 2)).value > 0 Then
                            Set lvwITM = lvw.ListItems.Add(, , ConvertirFecha(CLng(adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag & "", Len(lvw.ColumnHeaders(lCols).Tag & "") - 2)))))
                        Else
                            Set lvwITM = lvw.ListItems.Add(, , "")
                        End If
                    Else
                        Set lvwITM = lvw.ListItems.Add(, , adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag & "", Len(lvw.ColumnHeaders(lCols).Tag & "") - 2)))
                    End If
                Else
                    If InStr(lvw.ColumnHeaders(lCols).Tag, ".$") > 0 Then
                        Set lvwITM = lvw.ListItems.Add(, , Format(adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, Len(lvw.ColumnHeaders(lCols).Tag) - 2)), "Currency"))
                    Else
                        Set lvwITM = lvw.ListItems.Add(, , adoRec.Fields(lvw.ColumnHeaders(lCols).Tag) & "")
                    End If
                End If
            End If
            
            For lCols = 2 To lvw.ColumnHeaders.Count
                If InStr(lvw.ColumnHeaders(lCols).Tag, ",") > 0 Then
                    If Not IsNull(adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, InStr(lvw.ColumnHeaders(lCols).Tag, ",") - 1))) Then
                        lvwITM.SubItems(lCols - 1) = adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, InStr(lvw.ColumnHeaders(lCols).Tag, ",") - 1))
                    Else
                        lvwITM.SubItems(lCols - 1) = IIf(IsNull(adoRec.Fields(Mid(lvw.ColumnHeaders(lCols).Tag, InStr(lvw.ColumnHeaders(lCols).Tag, ",") + 1))), "", adoRec.Fields(Mid(lvw.ColumnHeaders(lCols).Tag, InStr(lvw.ColumnHeaders(lCols).Tag, ",") + 1)))
                    End If
                Else
                    If Right(lvw.ColumnHeaders(lCols).Tag, 1) = "F" Then
                        If IsNumeric(adoRec.Fields(Replace(lvw.ColumnHeaders(lCols).Tag, "_F", vbNullString)).value) Then
                           If adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, Len(lvw.ColumnHeaders(lCols).Tag) - 2)).value > 0 Then
                                lvwITM.SubItems(lCols - 1) = ConvertirFecha(CLng(adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, Len(lvw.ColumnHeaders(lCols).Tag) - 2))))
                            Else
                                lvwITM.SubItems(lCols - 1) = ""
                            End If
                        Else
                            lvwITM.SubItems(lCols - 1) = adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag & "", Len(lvw.ColumnHeaders(lCols).Tag & "") - 2))
                        End If
                    Else
                        If InStr(lvw.ColumnHeaders(lCols).Tag, ".$") > 0 Then
                            lvwITM.SubItems(lCols - 1) = Format(adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, Len(lvw.ColumnHeaders(lCols).Tag) - 2)), "Currency")
                        Else
                            lvwITM.SubItems(lCols - 1) = adoRec.Fields(lvw.ColumnHeaders(lCols).Tag & "")
                        End If
                    End If
                End If
            Next lCols
            On Error Resume Next
            If Not iIMG = -1 Then
                lvwITM.Icon = iIMG
                lvwITM.SmallIcon = iIMG
            End If
        End If
    End If
    adoRec.MoveNext
Wend

Exit Sub
NoEsta:
    If Err.Number = 91 Or Err.Number = 35600 Then Resume SeGUiR

End Sub

'##ModelId=3D38740000FA
Public Function EstaEnLaLista(ByVal lvw As ListView, ByVal sQue As String) As Boolean

'**************************************************'
'                                                  '
'  Esta funcion devuelve True si se encuentra sQue '
'  en lvw.                                         '
'                                                  '
'**************************************************'

Dim lIndx As Long

For lIndx = 1 To lvw.ListItems.Count
    If lvw.ListItems(lIndx).Text = sQue Then
        EstaEnLaLista = True
        Exit Function
    End If
Next
EstaEnLaLista = False

End Function


Public Sub PasardeListView(ByRef lvwDesde As ListView, ByVal lvwA As ListView, Optional bBorrarOriginal As Boolean = True)

Dim lvwITM As ListItem
Dim lCols As Long

Set lvwITM = lvwA.ListItems.Add(, lvwDesde.SelectedItem.key, lvwDesde.SelectedItem.Text)
lvwITM.Tag = lvwDesde.SelectedItem.Tag
For lCols = 1 To lvwDesde.ColumnHeaders.Count - 1
    lvwITM.SubItems(lCols) = lvwDesde.SelectedItem.SubItems(lCols)
Next

If bBorrarOriginal Then
    lvwDesde.ListItems.Remove lvwDesde.SelectedItem.Index
End If

End Sub

Public Sub ObtenerCBOITEM(ByRef cbo As ImageCombo, Optional sTexto As String, Optional sKey As String, Optional sTag As String)
Dim ct As ComboItem

If Not sKey = "" Then
    For Each ct In cbo.ComboItems
        If ct.key = sKey Then
            ct.Selected = True
            Exit For
        End If
    Next
ElseIf Not sTexto = "" Then
    For Each ct In cbo.ComboItems
    
        If ct.Text = sTexto Then
            ct.Selected = True
            Exit For
        End If
    Next
ElseIf Not sTag = "" Then
    For Each ct In cbo.ComboItems
            If ct.Tag = sTag Then
                ct.Selected = True
                Exit For
            End If
    Next
End If

Set ct = Nothing

End Sub

Public Function ExisteEnCombo(cbo As ImageCombo, Optional sTexto As String, Optional sKey As String, Optional sTag As String) As Boolean
Dim ct As ComboItem

ExisteEnCombo = False

For Each ct In cbo.ComboItems
    If Not sKey = "" Then
        If ct.key = sKey Then
            ExisteEnCombo = True
        End If
    ElseIf Not sTexto = "" Then
        If ct.Text = sTexto Then
            ExisteEnCombo = True
        End If
    ElseIf Not sTag = "" Then
        If ct.Tag = sTag Then
            ExisteEnCombo = True
        End If
    End If
Next

Set ct = Nothing

End Function

Public Sub DejarNumeros(ByRef txt As TextBox)
Dim i As Integer
Dim sNumero As String

If txt.Text = "" Then
    txt.Text = "0"
Else
    For i = 1 To Len(txt.Text)
        If IsNumeric(Mid(txt.Text, i, 1)) Then
            sNumero = sNumero & Mid(txt.Text, i, 1)
        End If
    Next
End If
    
End Sub

Public Sub ObtenerLSTITEM(ByRef lst As ListView, Optional sTexto As String, Optional sKey As String, Optional sTag As String)
Dim lt As ListItem

If Not sKey = "" Then
    For Each lt In lst.ListItems
        If lt.key = sKey Then
            lt.Selected = True
            lt.EnsureVisible
            Exit For
        End If
    Next
ElseIf Not sTexto = "" Then
    For Each lt In lst.ListItems
        If lt.Text = sTexto Then
            lt.Selected = True
            lt.EnsureVisible
            Exit For
        End If
    Next
ElseIf Not sTag = "" Then
    For Each lt In lst.ListItems
        If lt.Tag = sTag Then
            lt.Selected = True
            lt.EnsureVisible
            Exit For
        End If
    Next
End If

Set lt = Nothing

End Sub
'##ModelId=3D3873FE0334
Public Sub CargarListView1(ByVal lvw As ListView, ByVal adoRec As Recordset, Optional ByVal lvwSinEstos As ListView, Optional ByVal iIMG As Integer = -1)

'**************************************************'
'                                                  '
'  Este procedimiento carga los datos de adoRec    '
'  en un listView segun el Tag sus columnas. Si se '
'  especifica un ListView en lvwsinEstos, se cargan'
'  en lvw solo los datos que no estan en este y si '
'  iIMG es diferente a -1 se carga la imagen       '
'  especificada.                                   '
'                                                  '
'**************************************************'

Dim lvwITM As ListItem
Dim lCols As Long

lvw.ListItems.Clear
If adoRec.RecordCount = 0 Then Exit Sub
If adoRec.EOF And adoRec.BOF Then Exit Sub

adoRec.MoveFirst
While Not adoRec.EOF
On Error GoTo NoEsta
    If lvwSinEstos Is Nothing Then
SeGUiR:
        lCols = 1
        If Right(lvw.ColumnHeaders(lCols).Tag, 1) = "F" Then
            If IsNumeric(adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, Len(lvw.ColumnHeaders(lCols).Tag) - 2)).value) Then
                Set lvwITM = lvw.ListItems.Add(, , ConvertirFecha(CLng(adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, Len(lvw.ColumnHeaders(lCols).Tag) - 2)))))
            Else
                Set lvwITM = lvw.ListItems.Add(, , adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, Len(lvw.ColumnHeaders(lCols).Tag) - 2)))
            End If
        Else
            Set lvwITM = lvw.ListItems.Add(, , adoRec.Fields(lvw.ColumnHeaders(lCols).Tag))
        End If

        For lCols = 2 To lvw.ColumnHeaders.Count
            If Right(lvw.ColumnHeaders(lCols).Tag, 1) = "F" Then
                If IsNumeric(adoRec.Fields(Replace(lvw.ColumnHeaders(lCols).Tag, ".F", vbNullString)).value) Then
                    lvwITM.SubItems(lCols - 1) = ConvertirFecha(CLng(adoRec.Fields(Replace(lvw.ColumnHeaders(lCols).Tag, ".F", vbNullString))))
                Else
                    lvwITM.SubItems(lCols - 1) = adoRec.Fields(Replace(lvw.ColumnHeaders(lCols).Tag, ".F", vbNullString))
                End If
            Else
                lvwITM.SubItems(lCols - 1) = adoRec.Fields(lvw.ColumnHeaders(lCols).Tag)
            End If
        Next lCols
        On Error Resume Next
        If Not iIMG = -1 Then
            lvwITM.Icon = iIMG
            lvwITM.SmallIcon = iIMG
        End If
    Else
        If Not EstaEnLaLista(lvwSinEstos, adoRec.Fields(lvw.ColumnHeaders(1).Tag)) Then
            lCols = 1
            If Right(lvw.ColumnHeaders(lCols).Tag, 1) = "F" Then
                If IsNumeric(adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, Len(lvw.ColumnHeaders(lCols).Tag) - 2)).value) Then
                    Set lvwITM = lvw.ListItems.Add(, , ConvertirFecha(CLng(adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, Len(lvw.ColumnHeaders(lCols).Tag) - 2)))))
                Else
                    Set lvwITM = lvw.ListItems.Add(, , adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, Len(lvw.ColumnHeaders(lCols).Tag) - 2)))
                End If
            Else
                Set lvwITM = lvw.ListItems.Add(, , adoRec.Fields(lvw.ColumnHeaders(lCols).Tag))
            End If
        For lCols = 2 To lvw.ColumnHeaders.Count
                If Right(lvw.ColumnHeaders(lCols).Tag, 1) = "F" Then
                    If IsNumeric(adoRec.Fields(Replace(lvw.ColumnHeaders(lCols).Tag, "_F", vbNullString)).value) Then
                        lvwITM.SubItems(lCols - 1) = ConvertirFecha(CLng(adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, Len(lvw.ColumnHeaders(lCols).Tag) - 2))))
                    Else
                        lvwITM.SubItems(lCols - 1) = adoRec.Fields(Left(lvw.ColumnHeaders(lCols).Tag, Len(lvw.ColumnHeaders(lCols).Tag) - 2))
                    End If
                Else
                    lvwITM.SubItems(lCols - 1) = adoRec.Fields(lvw.ColumnHeaders(lCols).Tag)
                End If
            Next lCols
            On Error Resume Next
            If Not iIMG = -1 Then
                lvwITM.Icon = iIMG
                lvwITM.SmallIcon = iIMG
            End If
        End If
    End If
    adoRec.MoveNext
Wend

Exit Sub
NoEsta:
    If Err.Number = 91 Or Err.Number = 35600 Then Resume SeGUiR

End Sub
Public Sub ObtenerCBOITEMLike(ByRef cbo As ImageCombo, Optional sTexto As String)
Dim ct As ComboItem

If Not sTexto = "" Then
    For Each ct In cbo.ComboItems
        If UCase(ct.Text) Like (UCase(sTexto) & "*") Then
            cbo.Text = ct.Text
            cbo.Tag = ct.Tag
            cbo.SelLength = Len(cbo.Text)
            cbo.SelStart = Len(sTexto)
            ct.Selected = True
            Exit For
        End If
    Next
End If

Set ct = Nothing

End Sub


Public Sub CargarIMGCombo_PriFila(ByVal cbo As ImageCombo, ByVal vIMG As Variant, ByVal adoRec As Recordset, ByVal sCampoMostrar As String, Optional ByVal sCampoTag As String = "", Optional ByVal sKey As String, Optional StringPriFila As String)

'**************************************************'
'                                                  '
'  Es como CargarIMGCombo                          '
'  pero también permite agregar una primer fila    '
'  con el texto StringPriFila                      '
'                                                  '
'  ********************************************    '
'  Ej: StringPriFila="Seleccionar"                 '
'      entonces, la primera opción que se          '
'      mostrará dirá "Seleccionar" y obligará      '
'      al usuario a seleccionar alguna de las      '
'      de las otras opciones del combo             '
'                                                  '
'**************************************************'

Dim cboITM As ComboItem

Screen.MousePointer = vbHourglass

cbo.ComboItems.Clear

'agrego la primera fila
If StringPriFila <> "" Then

    Set cboITM = cbo.ComboItems.Add(, , StringPriFila)


End If




On Error GoTo errLOLO
If Not adoRec.RecordCount = 0 Then
    adoRec.MoveFirst
End If
OtroERR:
On Error Resume Next
While Not adoRec.EOF
    Set cboITM = cbo.ComboItems.Add(, , adoRec.Fields(sCampoMostrar).value)
    If Not vIMG = "-1" Then cboITM.Image = vIMG
    If Not sCampoTag = "" Then
        cboITM.Tag = adoRec.Fields(sCampoTag).value
    End If
    If Not sKey = "" Then
        If IsNumeric(adoRec.Fields(sKey).value) Then
            cboITM.key = adoRec.Fields(sKey).value & " n"
        Else
            cboITM.key = adoRec.Fields(sKey).value
        End If
    End If
    adoRec.MoveNext
Wend

On Error Resume Next
adoRec.MoveFirst
Set cboITM = Nothing

If Not cbo.ComboItems.Count = 0 Then
    cbo.ComboItems(1).Selected = True
End If

Screen.MousePointer = vbNormal

Exit Sub
errLOLO:
    If Err.Number = 91 Or Err.Number = 3704 Then
        Screen.MousePointer = vbNormal
        Exit Sub
    Else
        Resume OtroERR
    End If

End Sub




Public Sub ObtenerCBOITEM_val(ByRef cbo As ImageCombo, Optional sTexto As String, Optional sKey As String, Optional sTag As String)
Dim ct As ComboItem

If Not sKey = "" Then
    For Each ct In cbo.ComboItems
        If ct.key = sKey Then
            ct.Selected = True
            Exit For
        End If
    Next
ElseIf Not sTexto = "" Then
    For Each ct In cbo.ComboItems
        If ct.Text = sTexto Then
            ct.Selected = True
            Exit For
        End If
    Next
ElseIf Not sTag = "" Then
    For Each ct In cbo.ComboItems
'            If ct.Tag = sTag Then
'                ct.Selected = True
'                Exit For
'            End If
        If ct.Tag <> "" Then
            If Val(ct.Tag) = Val(sTag) Then
                ct.Selected = True
                Exit For
            End If
        Else
            If ct.Tag = sTag Then
                ct.Selected = True
                Exit For
            End If
        End If
    Next
End If

Set ct = Nothing

End Sub

Public Sub CargarIMGCombo_DescCod(ByVal cbo As ImageCombo, ByVal vIMG As Variant, ByVal adoRec As Recordset, ByVal sCampoMostrar As String, Optional ByVal sCampoTag As String = "", Optional ByVal sKey As String, Optional ByVal sOtroCampo As String = "")

'CARGA EL IMGCOMBO
'PONIENDO EN EL TEXTO => "sCampoMostrar"&-&"sOtroCampo"
'ej: "abcdef"&-&"012346" : abcdef-0123456

Dim cboITM As ComboItem

Screen.MousePointer = vbHourglass

cbo.ComboItems.Clear
On Error GoTo errLOLO
If Not adoRec.RecordCount = 0 Then
    adoRec.MoveFirst
End If
OtroERR:
On Error Resume Next
While Not adoRec.EOF
    If Not sOtroCampo = "" Then
        Set cboITM = cbo.ComboItems.Add(, , (Trim(adoRec.Fields(sCampoMostrar).value) & "-" & Trim(adoRec.Fields(sOtroCampo).value)))
    Else
        Set cboITM = cbo.ComboItems.Add(, , Trim(adoRec.Fields(sCampoMostrar).value))
    End If
    If Not vIMG = "-1" Then cboITM.Image = vIMG
    If Not sCampoTag = "" Then
        cboITM.Tag = adoRec.Fields(sCampoTag).value
    End If
    If Not sKey = "" Then
        If IsNumeric(adoRec.Fields(sKey).value) Then
            cboITM.key = adoRec.Fields(sKey).value & " n"
        Else
            cboITM.key = adoRec.Fields(sKey).value
        End If
    End If
    adoRec.MoveNext
Wend

On Error Resume Next
adoRec.MoveFirst
Set cboITM = Nothing

If Not cbo.ComboItems.Count = 0 Then
    cbo.ComboItems(1).Selected = True
End If

Screen.MousePointer = vbNormal

Exit Sub
errLOLO:
    If Err.Number = 91 Or Err.Number = 3704 Then
        Screen.MousePointer = vbNormal
        Exit Sub
    Else
        Resume OtroERR
    End If

End Sub


Public Function LlenarIN_2(ByRef adoRec As Recordset, ByVal sCampoValor As String, Optional ByVal Formato As String = "NN") As String
Dim i As Long
Dim sTMP As String

'*************************************************'
'                                                 '
'  Funcion que devuelve un string con un "IN()" o '
'  un "= X", para una sentecia de SQL             '
'                                                 '
'  Hecho por: Fernando (Chicho)                   '
'  Fecha de Creacion: 04/12/2001                  '
'                                                 '
'*************************************************'

'20071128
Dim sCampoValor2 As String


If adoRec.RecordCount = 0 Then
    LlenarIN_2 = ""
    Exit Function
End If

adoRec.MoveFirst
If adoRec.RecordCount = 1 Then
    sTMP = ConvertirDato(adoRec.Fields(sCampoValor).value)
    LlenarIN_2 = "=" & IIf(Formato = "NN", sTMP, Format(sTMP, Formato))
Else
    LlenarIN_2 = " IN("
    While Not adoRec.EOF
        '20071128
        If Not IsNull(adoRec.Fields(sCampoValor).value) Then
            sTMP = ConvertirDato(adoRec.Fields(sCampoValor).value)
            If Formato = "NN" And Not IsNull(adoRec.Fields(sCampoValor).value) Then
                LlenarIN_2 = LlenarIN_2 & sTMP & ", "
            Else
    '        LlenarIN_2 = LlenarIN_2 & IIf(Formato = "NN", sTMP, Format(sTMP, Formato)) & ", "
                LlenarIN_2 = LlenarIN_2 & Format(sTMP, Formato) & ", "
            End If
        Else
            sCampoValor2 = " or " & sCampoValor & " is NULL"
        End If
        
        
        adoRec.MoveNext
    Wend
    LlenarIN_2 = Left(LlenarIN_2, Len(LlenarIN_2) - 2) & ")"
End If
adoRec.MoveFirst

If Right(LlenarIN_2, 2) = ", " Then
    LlenarIN_2 = Left(LlenarIN_2, Len(LlenarIN_2) - 2) & ")"
End If

If sCampoValor2 <> "" Then
    LlenarIN_2 = LlenarIN_2 & sCampoValor2 & ")"
Else
    LlenarIN_2 = LlenarIN_2 & ")"
End If


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




Public Function UnicoID(conexion As String, id As Long, tabla As String, Campo As String, Optional sWhere As String = "") As Boolean
    Dim sSQL As String
    Dim rs As Recordset
    Dim cn As Connection
    
    sSQL = "SELECT " & Campo & " FROM " & tabla & " WHERE " & Campo & " = " & id & " " & Replace(sWhere, "WHERE", "AND")
    
    Set cn = New Connection
    cn.Open conexion
    Set rs = New Recordset
    rs.Open sSQL, cn, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.EOF And rs.BOF Then UnicoID = True Else UnicoID = False
    
    rs.Close
    cn.Close
    
    Set cn = Nothing
    Set rs = Nothing
    
End Function

Public Function UnicoStrID(conexion As String, id As String, tabla As String, Campo As String) As Boolean
    Dim sSQL As String
    Dim cn As Connection
    Dim rs As Recordset
    
    sSQL = "SELECT " & Campo & " FROM " & tabla & " WHERE " & Campo & " = '" & id & "'"
    
    Set cn = New Connection
    cn.Open conexion
    Set rs = New Recordset
    rs.Open sSQL, cn, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.EOF And rs.BOF Then UnicoStrID = True Else UnicoStrID = False
    
    rs.Close
    cn.Close
    
    Set cn = Nothing
    Set rs = Nothing
    
End Function

Public Sub LlenarCBOInte(cbo As ImageCombo, INTE_CODIGO As String, sCon As String)
'    Dim rs As Recordset
'    Dim cn As Connection
'    Dim sSQL As String
'
'    sSQL = "SELECT INTE_TEXTO AS MOSTRAR, INTE_ITEM AS TAG FROM REFER_INTERNAS WHERE INTE_CODIGO = " & INTE_CODIGO
'
'    Set cn = New Connection
'    Set rs = New Recordset
'
'    cn.Open sCon
'    rs.Open sSQL, cn, adOpenDynamic, adLockReadOnly, adCmdText
'
'    CargarComboTabla cbo, rs, False, False
'
'    rs.Close
'    cn.Close
'
'    Set rs = Nothing
'    Set cn = Nothing
    
End Sub

Public Sub LlenarCBOGral(cbo As ImageCombo, GRAL_TABLA As String, sCon As String, Optional NoAsigna As Boolean = True)
'    Dim rs As Recordset
'    Dim cn As Connection
'    Dim sSQL As String
'
'    sSQL = "SELECT GRAL_DESCRIPCION AS MOSTRAR, GRAL_IDREG AS TAG FROM TABLA_GENERAL WHERE GRAL_TABLA = " & GRAL_TABLA
'
'    Set cn = New Connection
'    Set rs = New Recordset
'
'    cn.Open sCon
'    rs.Open sSQL, cn, adOpenDynamic, adLockReadOnly, adCmdText
'
'    CargarComboTabla cbo, rs, True, NoAsigna
'
'    rs.Close
'    cn.Close
'
'    Set rs = Nothing
'    Set cn = Nothing
    
End Sub

Public Sub LlenarCBOPBInte(cbo As ImageCombo, INTE_CODIGO As String, sCon As String)
'    Dim rs As Recordset
'    Dim cn As Connection
'    Dim sSQL As String
'
'    sSQL = "SELECT PBINTE_TEXTO AS MOSTRAR, PBINTE_ITEM AS TAG FROM PUB_REFER_INTERNAS WHERE PBINTE_CODIGO = " & INTE_CODIGO
'
'    Set cn = New Connection
'    Set rs = New Recordset
'
'    cn.Open sCon
'    rs.Open sSQL, cn, adOpenDynamic, adLockReadOnly, adCmdText
'
'    CargarComboTabla cbo, rs, False, False
'
'    rs.Close
'    cn.Close
'
'    Set rs = Nothing
'    Set cn = Nothing
'
End Sub

Public Sub LlenarCBOPBGral(cbo As ImageCombo, GRAL_TABLA As String, sCon As String, Optional NoAsigna As Boolean = True)
'    Dim rs As Recordset
'    Dim cn As Connection
'    Dim sSQL As String
'
'    sSQL = "SELECT PBGRAL_DESCRIPCION AS MOSTRAR, PBGRAL_IDREG AS TAG FROM PUB_TABLA_GENERAL WHERE PBGRAL_TABLA = " & GRAL_TABLA
'
'    Set cn = New Connection
'    Set rs = New Recordset
'
'    cn.Open sCon
'    rs.Open sSQL, cn, adOpenDynamic, adLockReadOnly, adCmdText
'
'    CargarComboTabla cbo, rs, True, NoAsigna
'
'    rs.Close
'    cn.Close
'
'    Set rs = Nothing
'    Set cn = Nothing
    
End Sub

Public Function ConvertirCuit(Cuit As String) As String
    Cuit = Replace(Replace(Cuit, "-", ""), "_", "")
    If Len(Cuit) = 0 Then
        Cuit = 0
    End If
    ConvertirCuit = Cuit
    
End Function

Public Sub ObtenerCBOITEMSplit(ByRef cbo As ImageCombo, sTag As String, Index As Integer, Optional separador As String = "|")
Dim ct As ComboItem


    For Each ct In cbo.ComboItems
            If Split(ct.Tag, separador)(Index) = sTag Then
                ct.Selected = True
                Exit For
            End If
    Next


Set ct = Nothing

End Sub

Public Function ArchivoExiste(ByVal ruta As String) As Boolean
    'By Germán  02/07/2008
    On Local Error Resume Next

    ArchivoExiste = Len(Dir$(ruta))
    If ruta = "" Then ArchivoExiste = False
    If Err Then
        ArchivoExiste = False
    End If
    Err = 0
    On Local Error GoTo 0
End Function

Public Function ExportarSGaExcel(ByVal hWnd As Long, ByRef Grilla As SGGrid)
    'By Germán 07/07/2008
'    frmExportarExcelGenerico.Mostrar hWnd, Grilla
End Function

Public Function RellenarCadena(cadena As String, longitud As Integer, Optional Relleno As String = " ", Optional Derecha As Boolean = True)
    If Not IsNull(cadena) Then
        While Len(cadena) < longitud
            If Derecha Then
                cadena = cadena & Relleno
            Else
                cadena = Relleno & cadena
            End If
        Wend
        RellenarCadena = cadena
    Else
        RellenarCadena = Space(longitud)
    End If
    
End Function

 Public Function Dias_Del_Mes(Optional ByVal Fecha As Variant) As Integer


     Dim mes As Integer, y  As Integer

     If IsMissing(Fecha) Then Fecha = Date

If IsDate(Fecha) Then
         y = Year(Fecha)
         mes = Month(Fecha)
ElseIf IsNumeric(Fecha) Then
         y = Year(Date)
         mes = IIf(Fecha > 0 And Fecha < 13, CInt(Fecha), 0)
ElseIf VarType(Fecha) = vbString Then
         y = Year(Date)
         Select Case UCase(Left$(Fecha, 3))
             Case "FEB":                                             mes = 2
             Case "JAN", "MAR", "MAY", "JUL", "AUG", "OCT", "DEC":   mes = 1
             Case "APR", "JUN", "SEP", "NOV":                        mes = 4
         End Select
End If

     Select Case mes
         Case 2:                     Dias_Del_Mes = IIf(saltarYear(Fecha), 29, 28)
         Case 1, 3, 5, 7, 8, 10, 12: Dias_Del_Mes = 31
         Case 4, 6, 9, 11:           Dias_Del_Mes = 30
     End Select

 End Function
 

 Public Function saltarYear(ByVal valor As Variant) As Boolean


     On Error GoTo LocalError

     Dim iYear As Integer

     If IsDate(valor) Then iYear = Year(valor) Else iYear = CInt(valor)

If TypeName(iYear) = "Integer" Then
         saltarYear = Day(DateSerial(iYear, 3, 0)) = 29
End If
 Exit Function

LocalError:
 End Function

Public Function ValidarFecha_txt(ByVal sFecha As String, iAnioMin As Integer, iAnioMax As Integer) As Boolean
'Estamos asumiendo que la fecha es ddmmmaaaa

    Dim iFecha As Long
    Dim iDia As Integer, iMes As Integer, iAnio As Integer
    Dim iTopeFeb As Integer
    
    iFecha = Val(sFecha)
    iDia = Val(Mid(sFecha, 1, 2))
    iMes = Val(Mid(sFecha, 3, 2))
    iAnio = Val(Mid(sFecha, 5, 4))
    
    If (iAnio Mod 4) = 0 And (iAnio Mod 100) <> 0 Then
        iTopeFeb = 29
    Else
        iTopeFeb = 28
    End If
    
    ValidarFecha_txt = (iFecha < 0)
    ValidarFecha_txt = ValidarFecha_txt Or ((iDia < 1 Or iDia > 31) Or (iMes < 1 Or iMes > 12) Or (iAnio < iAnioMin Or iAnio > iAnioMax))
    ValidarFecha_txt = ValidarFecha_txt Or ((InStr(Mid(sFecha, 3, 2), "04060911") > 0 And iDia > 30) Or (iMes = 2 And iDia > iTopeFeb))
    ValidarFecha_txt = Not ValidarFecha_txt
    
End Function

