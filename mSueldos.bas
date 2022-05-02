Attribute VB_Name = "mSueldos"
Public Function CargarDatos(ByVal sPath As String, ByVal iCodigo As Integer) As Boolean
    
    Dim mRegistro As String * 492
    Dim mRegAnex As String * 249
    Dim mRecibo As String * 442
    
    Dim cn As Connection
    Dim cmd As Command
    
    Set cn = New Connection
    Set cmd = New Command
   
    FrmProceso.Show
    
    cn.Open SGEPOB
    
    cn.Execute "Delete from PersSue where Sucursal = " & iCodigo
    cn.Execute "Delete from PersAnex where Sucursal = " & iCodigo
    
    cmd.ActiveConnection = cn
    cmd.CommandText = "PERSSUE_INSERT"
    cmd.CommandType = adCmdStoredProc
    
    Open sPath & "\Sue101.Arc" For Input As #1
    
    While Not EOF(1)
    
        Line Input #1, mRegistro
        
        If Mid(mRegistro, 299, 1) <> "*" Then
                   
            FrmProceso.pbar.value = IIf(FrmProceso.pbar.value = FrmProceso.pbar.Max, 1, FrmProceso.pbar.value + 1)
            
            cmd.Parameters.Refresh
                
            cmd.Parameters.Item(1).value = Mid(mRegistro, 1, 5)
            cmd.Parameters.Item(2).value = Mid(mRegistro, 6, 24)
            cmd.Parameters.Item(3).value = Mid(mRegistro, 30, 24)
            cmd.Parameters.Item(4).value = Mid(mRegistro, 54, 4)
            cmd.Parameters.Item(5).value = Mid(mRegistro, 58, 8)
            cmd.Parameters.Item(6).value = Mid(mRegistro, 66, 1)
            cmd.Parameters.Item(7).value = Mid(mRegistro, 67, 1)
            cmd.Parameters.Item(8).value = Mid(mRegistro, 68, 6)
            cmd.Parameters.Item(9).value = Mid(mRegistro, 74, 6)
            cmd.Parameters.Item(10).value = Mid(mRegistro, 80, 2)
            cmd.Parameters.Item(11).value = Mid(mRegistro, 82, 8)
            cmd.Parameters.Item(12).value = Mid(mRegistro, 90, 1)
            cmd.Parameters.Item(13).value = Mid(mRegistro, 91, 1)
            cmd.Parameters.Item(14).value = Mid(mRegistro, 92, 10)
            cmd.Parameters.Item(15).value = Mid(mRegistro, 102, 1)
            cmd.Parameters.Item(16).value = Mid(mRegistro, 103, 1)
            cmd.Parameters.Item(17).value = Mid(mRegistro, 104, 10)
            cmd.Parameters.Item(18).value = Mid(mRegistro, 114, 5)
            cmd.Parameters.Item(19).value = Mid(mRegistro, 119, 1)
            cmd.Parameters.Item(20).value = Mid(mRegistro, 120, 1)
            cmd.Parameters.Item(21).value = Mid(mRegistro, 121, 10)
            cmd.Parameters.Item(22).value = Mid(mRegistro, 131, 5)
            cmd.Parameters.Item(23).value = Mid(mRegistro, 136, 6)
            cmd.Parameters.Item(24).value = Mid(mRegistro, 142, 2)
            cmd.Parameters.Item(25).value = Mid(mRegistro, 144, 2)
            cmd.Parameters.Item(26).value = Mid(mRegistro, 146, 2)
            cmd.Parameters.Item(27).value = Mid(mRegistro, 148, 2)
            cmd.Parameters.Item(28).value = Mid(mRegistro, 150, 1)
            cmd.Parameters.Item(29).value = Mid(mRegistro, 151, 1)
            cmd.Parameters.Item(30).value = Mid(mRegistro, 152, 6)
            cmd.Parameters.Item(31).value = Mid(mRegistro, 158, 2)
            cmd.Parameters.Item(32).value = Mid(mRegistro, 160, 1)
            cmd.Parameters.Item(33).value = Mid(mRegistro, 161, 1)
            cmd.Parameters.Item(34).value = Mid(mRegistro, 162, 2)
            cmd.Parameters.Item(35).value = Mid(mRegistro, 164, 10)
            cmd.Parameters.Item(36).value = Mid(mRegistro, 174, 1) 'estado del personal: 1=bloqueado
            cmd.Parameters.Item(37).value = Mid(mRegistro, 175, 2)
            cmd.Parameters.Item(38).value = Mid(mRegistro, 177, 6)
            cmd.Parameters.Item(39).value = Mid(mRegistro, 183, 1)
            cmd.Parameters.Item(40).value = Mid(mRegistro, 184, 1)
            cmd.Parameters.Item(41).value = Mid(mRegistro, 185, 1)
            cmd.Parameters.Item(42).value = Mid(mRegistro, 186, 1)
            cmd.Parameters.Item(43).value = Mid(mRegistro, 187, 2)
            cmd.Parameters.Item(44).value = Mid(mRegistro, 189, 5)
            cmd.Parameters.Item(45).value = Mid(mRegistro, 194, 5)
            cmd.Parameters.Item(46).value = Mid(mRegistro, 199, 5)
            cmd.Parameters.Item(47).value = Mid(mRegistro, 204, 1)
            cmd.Parameters.Item(48).value = Mid(mRegistro, 205, 1)
            cmd.Parameters.Item(49).value = Mid(mRegistro, 206, 1)
            cmd.Parameters.Item(50).value = Mid(mRegistro, 207, 1)
            cmd.Parameters.Item(51).value = Mid(mRegistro, 208, 1)
            cmd.Parameters.Item(52).value = Mid(mRegistro, 209, 1)
            cmd.Parameters.Item(53).value = Mid(mRegistro, 210, 1)
            cmd.Parameters.Item(54).value = Mid(mRegistro, 211, 1)
            cmd.Parameters.Item(55).value = Mid(mRegistro, 212, 1)
            cmd.Parameters.Item(56).value = Mid(mRegistro, 213, 1)
            cmd.Parameters.Item(57).value = Mid(mRegistro, 214, 1)
            cmd.Parameters.Item(58).value = Mid(mRegistro, 215, 1)
            cmd.Parameters.Item(59).value = Mid(mRegistro, 216, 1)
            cmd.Parameters.Item(60).value = Mid(mRegistro, 217, 1)
            cmd.Parameters.Item(61).value = Mid(mRegistro, 218, 8)
            cmd.Parameters.Item(62).value = Mid(mRegistro, 226, 8)
            cmd.Parameters.Item(63).value = Mid(mRegistro, 234, 9)
            cmd.Parameters.Item(64).value = Mid(mRegistro, 243, 16)
            cmd.Parameters.Item(65).value = Mid(mRegistro, 259, 1)
            cmd.Parameters.Item(66).value = Mid(mRegistro, 260, 1)
            cmd.Parameters.Item(67).value = Mid(mRegistro, 261, 4)
            cmd.Parameters.Item(68).value = Mid(mRegistro, 265, 2)
            cmd.Parameters.Item(69).value = Mid(mRegistro, 267, 2)
            cmd.Parameters.Item(70).value = Mid(mRegistro, 269, 5)
            cmd.Parameters.Item(71).value = Mid(mRegistro, 274, 6)
            cmd.Parameters.Item(72).value = Mid(mRegistro, 280, 8)
            cmd.Parameters.Item(73).value = Mid(mRegistro, 288, 8)
            cmd.Parameters.Item(74).value = Mid(mRegistro, 296, 8)
            cmd.Parameters.Item(75).value = Mid(mRegistro, 304, 8)
            cmd.Parameters.Item(76).value = Mid(mRegistro, 312, 8)
            cmd.Parameters.Item(77).value = Mid(mRegistro, 320, 8)
            cmd.Parameters.Item(78).value = Mid(mRegistro, 328, 8)
            cmd.Parameters.Item(79).value = Mid(mRegistro, 336, 3)
            cmd.Parameters.Item(80).value = Mid(mRegistro, 339, 2)
            cmd.Parameters.Item(81).value = Mid(mRegistro, 341, 5)
            cmd.Parameters.Item(82).value = Mid(mRegistro, 346, 5)
            cmd.Parameters.Item(83).value = Mid(mRegistro, 351, 5)
            cmd.Parameters.Item(84).value = Mid(mRegistro, 356, 5)
            cmd.Parameters.Item(85).value = Mid(mRegistro, 361, 8)
            cmd.Parameters.Item(86).value = Mid(mRegistro, 369, 8)
            cmd.Parameters.Item(87).value = Mid(mRegistro, 377, 8)
            cmd.Parameters.Item(88).value = Mid(mRegistro, 385, 8)
            cmd.Parameters.Item(89).value = Mid(mRegistro, 393, 16)
            cmd.Parameters.Item(90).value = Mid(mRegistro, 409, 8)
            cmd.Parameters.Item(91).value = Mid(mRegistro, 417, 2)
            cmd.Parameters.Item(92).value = Mid(mRegistro, 419, 2)
            cmd.Parameters.Item(93).value = Mid(mRegistro, 421, 3)
            cmd.Parameters.Item(94).value = Mid(mRegistro, 424, 8)
            cmd.Parameters.Item(95).value = Mid(mRegistro, 432, 16)
            cmd.Parameters.Item(96).value = Mid(mRegistro, 448, 16)
            cmd.Parameters.Item(97).value = Mid(mRegistro, 464, 9)
            cmd.Parameters.Item(98).value = Mid(mRegistro, 473, 9)
            cmd.Parameters.Item(99).value = Mid(mRegistro, 482, 9)
            cmd.Parameters.Item(100).value = iCodigo
            cmd.Execute
        
        End If
    
    Wend

    Close #1
    
    Open sPath & "\SuelAnex.Arc" For Input As #1
    
    cmd.CommandText = "PERSANEX_INSERT"
    cmd.CommandType = adCmdStoredProc

    While Not EOF(1)
        
        Line Input #1, mRegAnex
        
        If Mid(mRegistro, 1, 1) <> "*" Then
                   
            FrmProceso.pbar.value = IIf(FrmProceso.pbar.value = FrmProceso.pbar.Max, 1, FrmProceso.pbar.value + 1)
            
            cmd.Parameters.Refresh
                
            cmd.Parameters.Item(1).value = Mid(mRegAnex, 1, 1)
            cmd.Parameters.Item(2).value = Mid(mRegAnex, 2, 5)
            cmd.Parameters.Item(3).value = Mid(mRegAnex, 7, 22)
            cmd.Parameters.Item(4).value = Mid(mRegAnex, 29, 3)
            cmd.Parameters.Item(5).value = Mid(mRegAnex, 32, 40)
            cmd.Parameters.Item(6).value = Mid(mRegAnex, 72, 110)
            cmd.Parameters.Item(7).value = Mid(mRegAnex, 182, 60)
            cmd.Parameters.Item(8).value = Mid(mRegAnex, 242, 6)
            cmd.Parameters.Item(9).value = iCodigo
            cmd.Execute
        
        End If
        
    Wend
    
    Close #1
    
    cn.Close
    
    Unload FrmProceso
    
    Set cn = Nothing
    Set cmd = Nothing
    
    
End Function

Public Function CargarDatosConRecibos(ByVal sPath As String, ByVal iCodigo As Integer) As Boolean

    Dim mRegistro As String * 492
    Dim mRegAnex As String * 249
    Dim mRecibo As String * 442

    Dim cn As Connection
    Dim cmd As Command
    Dim cmd01 As Command
    Dim cmd31 As Command
    Dim cmd41 As Command
    Dim cmd51 As Command
    Dim cmd52 As Command
    
    Set cn = New Connection
    Set cmd = New Command

    Set cmd01 = New Command
    Set cmd31 = New Command
    Set cmd41 = New Command
    Set cmd51 = New Command
    Set cmd52 = New Command

    cn.Open SGEPOB
    
    FrmProceso.Show
    
    cn.Execute "Delete from PersSue where Sucursal = " & iCodigo
    cn.Execute "Delete from PersAnex where Sucursal = " & iCodigo
    
    cmd.ActiveConnection = cn
    cmd.CommandText = "PERSSUE_INSERT"
    cmd.CommandType = adCmdStoredProc

    Open sPath & "\Sue101.Arc" For Input As #1
    
    While Not EOF(1)
    
        Line Input #1, mRegistro
        
        If Mid(mRegistro, 299, 1) <> "*" Then

            cmd.Parameters.Refresh
            
            FrmProceso.pbar.value = IIf(FrmProceso.pbar.value = FrmProceso.pbar.Max, 1, FrmProceso.pbar.value + 1)

            cmd.Parameters.Item(1).value = Mid(mRegistro, 1, 5)
            cmd.Parameters.Item(2).value = Mid(mRegistro, 6, 24)
            cmd.Parameters.Item(3).value = Mid(mRegistro, 30, 24)
            cmd.Parameters.Item(4).value = Mid(mRegistro, 54, 4)
            cmd.Parameters.Item(5).value = Mid(mRegistro, 58, 8)
            cmd.Parameters.Item(6).value = Mid(mRegistro, 66, 1)
            cmd.Parameters.Item(7).value = Mid(mRegistro, 67, 1)
            cmd.Parameters.Item(8).value = Mid(mRegistro, 68, 6)
            cmd.Parameters.Item(9).value = Mid(mRegistro, 74, 6)
            cmd.Parameters.Item(10).value = Mid(mRegistro, 80, 2)
            cmd.Parameters.Item(11).value = Mid(mRegistro, 82, 8)
            cmd.Parameters.Item(12).value = Mid(mRegistro, 90, 1)
            cmd.Parameters.Item(13).value = Mid(mRegistro, 91, 1)
            cmd.Parameters.Item(14).value = Mid(mRegistro, 92, 10)
            cmd.Parameters.Item(15).value = Mid(mRegistro, 102, 1)
            cmd.Parameters.Item(16).value = Mid(mRegistro, 103, 1)
            cmd.Parameters.Item(17).value = Mid(mRegistro, 104, 10)
            cmd.Parameters.Item(18).value = Mid(mRegistro, 114, 5)
            cmd.Parameters.Item(19).value = Mid(mRegistro, 119, 1)
            cmd.Parameters.Item(20).value = Mid(mRegistro, 120, 1)
            cmd.Parameters.Item(21).value = Mid(mRegistro, 121, 10)
            cmd.Parameters.Item(22).value = Mid(mRegistro, 131, 5)
            cmd.Parameters.Item(23).value = Mid(mRegistro, 136, 6)
            cmd.Parameters.Item(24).value = Mid(mRegistro, 142, 2)
            cmd.Parameters.Item(25).value = Mid(mRegistro, 144, 2)
            cmd.Parameters.Item(26).value = Mid(mRegistro, 146, 2)
            cmd.Parameters.Item(27).value = Mid(mRegistro, 148, 2)
            cmd.Parameters.Item(28).value = Mid(mRegistro, 150, 1)
            cmd.Parameters.Item(29).value = Mid(mRegistro, 151, 1)
            cmd.Parameters.Item(30).value = Mid(mRegistro, 152, 6)
            cmd.Parameters.Item(31).value = Mid(mRegistro, 158, 2)
            cmd.Parameters.Item(32).value = Mid(mRegistro, 160, 1)
            cmd.Parameters.Item(33).value = Mid(mRegistro, 161, 1)
            cmd.Parameters.Item(34).value = Mid(mRegistro, 162, 2)
            cmd.Parameters.Item(35).value = Mid(mRegistro, 164, 10)
            cmd.Parameters.Item(36).value = Mid(mRegistro, 174, 1) 'estado del personal: 1=bloqueado
            cmd.Parameters.Item(37).value = Mid(mRegistro, 175, 2)
            cmd.Parameters.Item(38).value = Mid(mRegistro, 177, 6)
            cmd.Parameters.Item(39).value = Mid(mRegistro, 183, 1)
            cmd.Parameters.Item(40).value = Mid(mRegistro, 184, 1)
            cmd.Parameters.Item(41).value = Mid(mRegistro, 185, 1)
            cmd.Parameters.Item(42).value = Mid(mRegistro, 186, 1)
            cmd.Parameters.Item(43).value = Mid(mRegistro, 187, 2)
            cmd.Parameters.Item(44).value = Mid(mRegistro, 189, 5)
            cmd.Parameters.Item(45).value = Mid(mRegistro, 194, 5)
            cmd.Parameters.Item(46).value = Mid(mRegistro, 199, 5)
            cmd.Parameters.Item(47).value = Mid(mRegistro, 204, 1)
            cmd.Parameters.Item(48).value = Mid(mRegistro, 205, 1)
            cmd.Parameters.Item(49).value = Mid(mRegistro, 206, 1)
            cmd.Parameters.Item(50).value = Mid(mRegistro, 207, 1)
            cmd.Parameters.Item(51).value = Mid(mRegistro, 208, 1)
            cmd.Parameters.Item(52).value = Mid(mRegistro, 209, 1)
            cmd.Parameters.Item(53).value = Mid(mRegistro, 210, 1)
            cmd.Parameters.Item(54).value = Mid(mRegistro, 211, 1)
            cmd.Parameters.Item(55).value = Mid(mRegistro, 212, 1)
            cmd.Parameters.Item(56).value = Mid(mRegistro, 213, 1)
            cmd.Parameters.Item(57).value = Mid(mRegistro, 214, 1)
            cmd.Parameters.Item(58).value = Mid(mRegistro, 215, 1)
            cmd.Parameters.Item(59).value = Mid(mRegistro, 216, 1)
            cmd.Parameters.Item(60).value = Mid(mRegistro, 217, 1)
            cmd.Parameters.Item(61).value = Mid(mRegistro, 218, 8)
            cmd.Parameters.Item(62).value = Mid(mRegistro, 226, 8)
            cmd.Parameters.Item(63).value = Mid(mRegistro, 234, 9)
            cmd.Parameters.Item(64).value = Mid(mRegistro, 243, 16)
            cmd.Parameters.Item(65).value = Mid(mRegistro, 259, 1)
            cmd.Parameters.Item(66).value = Mid(mRegistro, 260, 1)
            cmd.Parameters.Item(67).value = Mid(mRegistro, 261, 4)
            cmd.Parameters.Item(68).value = Mid(mRegistro, 265, 2)
            cmd.Parameters.Item(69).value = Mid(mRegistro, 267, 2)
            cmd.Parameters.Item(70).value = Mid(mRegistro, 269, 5)
            cmd.Parameters.Item(71).value = Mid(mRegistro, 274, 6)
            cmd.Parameters.Item(72).value = Mid(mRegistro, 280, 8)
            cmd.Parameters.Item(73).value = Mid(mRegistro, 288, 8)
            cmd.Parameters.Item(74).value = Mid(mRegistro, 296, 8)
            cmd.Parameters.Item(75).value = Mid(mRegistro, 304, 8)
            cmd.Parameters.Item(76).value = Mid(mRegistro, 312, 8)
            cmd.Parameters.Item(77).value = Mid(mRegistro, 320, 8)
            cmd.Parameters.Item(78).value = Mid(mRegistro, 328, 8)
            cmd.Parameters.Item(79).value = Mid(mRegistro, 336, 3)
            cmd.Parameters.Item(80).value = Mid(mRegistro, 339, 2)
            cmd.Parameters.Item(81).value = Mid(mRegistro, 341, 5)
            cmd.Parameters.Item(82).value = Mid(mRegistro, 346, 5)
            cmd.Parameters.Item(83).value = Mid(mRegistro, 351, 5)
            cmd.Parameters.Item(84).value = Mid(mRegistro, 356, 5)
            cmd.Parameters.Item(85).value = Mid(mRegistro, 361, 8)
            cmd.Parameters.Item(86).value = Mid(mRegistro, 369, 8)
            cmd.Parameters.Item(87).value = Mid(mRegistro, 377, 8)
            cmd.Parameters.Item(88).value = Mid(mRegistro, 385, 8)
            cmd.Parameters.Item(89).value = Mid(mRegistro, 393, 16)
            cmd.Parameters.Item(90).value = Mid(mRegistro, 409, 8)
            cmd.Parameters.Item(91).value = Mid(mRegistro, 417, 2)
            cmd.Parameters.Item(92).value = Mid(mRegistro, 419, 2)
            cmd.Parameters.Item(93).value = Mid(mRegistro, 421, 3)
            cmd.Parameters.Item(94).value = Mid(mRegistro, 424, 8)
            cmd.Parameters.Item(95).value = Mid(mRegistro, 432, 16)
            cmd.Parameters.Item(96).value = Mid(mRegistro, 448, 16)
            cmd.Parameters.Item(97).value = Mid(mRegistro, 464, 9)
            cmd.Parameters.Item(98).value = Mid(mRegistro, 473, 9)
            cmd.Parameters.Item(99).value = Mid(mRegistro, 482, 9)
            cmd.Parameters.Item(100).value = iCodigo
            cmd.Execute

        End If

    Wend

    Close #1
    
    Open sPath & "\SuelAnex.Arc" For Input As #1

    cmd.CommandText = "PERSANEX_INSERT"
    cmd.CommandType = adCmdStoredProc
    
    While Not EOF(1)
    
        Line Input #1, mRegAnex
        
        If Mid(mRegistro, 1, 1) <> "*" Then

            FrmProceso.pbar.value = IIf(FrmProceso.pbar.value = FrmProceso.pbar.Max, 1, FrmProceso.pbar.value + 1)

            cmd.Parameters.Refresh

            cmd.Parameters.Item(1).value = Mid(mRegAnex, 1, 1)
            cmd.Parameters.Item(2).value = Mid(mRegAnex, 2, 5)
            cmd.Parameters.Item(3).value = Mid(mRegAnex, 7, 22)
            cmd.Parameters.Item(4).value = Mid(mRegAnex, 29, 3)
            cmd.Parameters.Item(5).value = Mid(mRegAnex, 32, 40)
            cmd.Parameters.Item(6).value = Mid(mRegAnex, 72, 110)
            cmd.Parameters.Item(7).value = Mid(mRegAnex, 182, 60)
            cmd.Parameters.Item(8).value = Mid(mRegAnex, 242, 6)
            cmd.Parameters.Item(9).value = iCodigo
            cmd.Execute

        End If

    Wend

    Close #1
    
    Open sPath & "\IfRecibo.txt" For Input As #1
    Line Input #1, mRecibo
    Close #1
    
    sSQL = "Delete from Recibos_reg01 where Reg01_Anio = " & Mid(mRecibo, 4, 4)
    sSQL = sSQL & " and Reg01_Mes = " & Mid(mRecibo, 8, 2)
    sSQL = sSQL & " and Reg01_TipoLiq = " & Mid(mRecibo, 10, 2)
    sSQL = sSQL & " and Reg01_NroLiq = " & Mid(mRecibo, 12, 6)
    cn.Execute sSQL
    
    sSQL = "Delete from Recibos_reg31 where Reg31_Anio = " & Mid(mRecibo, 4, 4)
    sSQL = sSQL & " and Reg31_Mes = " & Mid(mRecibo, 8, 2)
    sSQL = sSQL & " and Reg31_TipoLiq = " & Mid(mRecibo, 10, 2)
    sSQL = sSQL & " and Reg31_NroLiq = " & Mid(mRecibo, 12, 6)
    cn.Execute sSQL
    
    sSQL = "Delete from Recibos_reg41 where Reg41_Anio = " & Mid(mRecibo, 4, 4)
    sSQL = sSQL & " and Reg41_Mes = " & Mid(mRecibo, 8, 2)
    sSQL = sSQL & " and Reg41_TipoLiq = " & Mid(mRecibo, 10, 2)
    sSQL = sSQL & " and Reg41_NroLiq = " & Mid(mRecibo, 12, 6)
    cn.Execute sSQL
    
    sSQL = "Delete from Recibos_reg51 where Reg51_Anio = " & Mid(mRecibo, 4, 4)
    sSQL = sSQL & " and Reg51_Mes = " & Mid(mRecibo, 8, 2)
    sSQL = sSQL & " and Reg51_TipoLiq = " & Mid(mRecibo, 10, 2)
    sSQL = sSQL & " and Reg51_NroLiq = " & Mid(mRecibo, 12, 6)
    cn.Execute sSQL
    
    sSQL = "Delete from Recibos_reg52 where Reg52_Anio = " & Mid(mRecibo, 4, 4)
    sSQL = sSQL & " and Reg52_Mes = " & Mid(mRecibo, 8, 2)
    sSQL = sSQL & " and Reg52_TipoLiq = " & Mid(mRecibo, 10, 2)
    sSQL = sSQL & " and Reg52_NroLiq = " & Mid(mRecibo, 12, 6)
    cn.Execute sSQL
    
    cmd01.ActiveConnection = cn
    cmd01.CommandText = "RECIBOS_REG01_INSERT"
    cmd01.CommandType = adCmdStoredProc

    cmd31.ActiveConnection = cn
    cmd31.CommandText = "RECIBOS_REG31_INSERT"
    cmd31.CommandType = adCmdStoredProc

    cmd41.ActiveConnection = cn
    cmd41.CommandText = "RECIBOS_REG41_INSERT"
    cmd41.CommandType = adCmdStoredProc

    cmd51.ActiveConnection = cn
    cmd51.CommandText = "RECIBOS_REG51_INSERT"
    cmd51.CommandType = adCmdStoredProc

    cmd52.ActiveConnection = cn
    cmd52.CommandText = "RECIBOS_REG52_INSERT"
    cmd52.CommandType = adCmdStoredProc

    Open sPath & "\IfRecibo.txt" For Input As #1

    While Not EOF(1)

        Line Input #1, mRecibo
        
        FrmProceso.pbar.value = IIf(FrmProceso.pbar.value = FrmProceso.pbar.Max, 1, FrmProceso.pbar.value + 1)

        If Mid(mRecibo, 1, 3) = "001" Then
            cmd01.Parameters.Refresh

            cmd01.Parameters.Item(1).value = Mid(mRecibo, 4, 4)
            cmd01.Parameters.Item(2).value = Mid(mRecibo, 8, 2)
            cmd01.Parameters.Item(3).value = Mid(mRecibo, 10, 2)
            cmd01.Parameters.Item(4).value = Mid(mRecibo, 12, 6)
            cmd01.Parameters.Item(5).value = Val(Mid(mRecibo, 18, 5))
            cmd01.Parameters.Item(6).value = Val(Mid(mRecibo, 31, 3))
            cmd01.Parameters.Item(7).value = CDate(Mid(mRecibo, 34, 10))
            cmd01.Parameters.Item(8).value = Mid(mRecibo, 44, 2)
            cmd01.Parameters.Item(9).value = Mid(mRecibo, 46, 3)
            cmd01.Parameters.Item(10).value = Mid(mRecibo, 49, 30)
            cmd01.Parameters.Item(11).value = iCodigo
            cmd01.Parameters.Item(12).value = "N"
            cmd01.Execute
        End If

        If Mid(mRecibo, 1, 3) = "031" Then
            cmd31.Parameters.Refresh

            cmd31.Parameters.Item(1).value = Mid(mRecibo, 4, 4)
            cmd31.Parameters.Item(2).value = Mid(mRecibo, 8, 2)
            cmd31.Parameters.Item(3).value = Mid(mRecibo, 10, 2)
            cmd31.Parameters.Item(4).value = Mid(mRecibo, 12, 6)
            cmd31.Parameters.Item(5).value = Mid(mRecibo, 18, 5)
            cmd31.Parameters.Item(6).value = Mid(mRecibo, 31, 3)
            cmd31.Parameters.Item(7).value = Mid(mRecibo, 34, 12)
            cmd31.Parameters.Item(8).value = CDate(Mid(mRecibo, 46, 8))
            cmd31.Parameters.Item(9).value = Mid(mRecibo, 54, 40)
            cmd31.Parameters.Item(10).value = Mid(mRecibo, 94, 12)
            cmd31.Parameters.Item(11).value = Mid(mRecibo, 106, 13)
            cmd31.Parameters.Item(12).value = CDate(Mid(mRecibo, 119, 8))
            cmd31.Parameters.Item(13).value = Replace(Mid(mRecibo, 127, 12), ".", ",")
            cmd31.Parameters.Item(14).value = Mid(mRecibo, 144, 8)
            cmd31.Parameters.Item(15).value = Mid(mRecibo, 152, 40)
            cmd31.Parameters.Item(16).value = Mid(mRecibo, 192, 30)
            cmd31.Parameters.Item(17).value = Mid(mRecibo, 222, 30)
            cmd31.Parameters.Item(18).value = CDate(Mid(mRecibo, 252, 8))
            cmd31.Parameters.Item(19).value = Mid(mRecibo, 260, 5)
            cmd31.Parameters.Item(20).value = Mid(mRecibo, 265, 30)
            cmd31.Parameters.Item(21).value = Replace(Mid(mRecibo, 295, 12), ".", ",")
            cmd31.Parameters.Item(22).value = Mid(mRecibo, 307, 30)
            cmd31.Parameters.Item(23).value = Mid(mRecibo, 337, 35)
            cmd31.Parameters.Item(24).value = Mid(mRecibo, 372, 29)
            cmd31.Parameters.Item(25).value = Mid(mRecibo, 401, 32)
            cmd31.Parameters.Item(26).value = Mid(mRecibo, 433, 8)
            cmd31.Parameters.Item(27).value = iCodigo
            cmd31.Parameters.Item(28).value = "N"
            cmd31.Execute
        End If

        If Mid(mRecibo, 1, 3) = "041" Then
            cmd41.Parameters.Refresh

            cmd41.Parameters.Item(1).value = Mid(mRecibo, 4, 4)
            cmd41.Parameters.Item(2).value = Mid(mRecibo, 8, 2)
            cmd41.Parameters.Item(3).value = Mid(mRecibo, 10, 2)
            cmd41.Parameters.Item(4).value = Mid(mRecibo, 12, 6)
            cmd41.Parameters.Item(5).value = Mid(mRecibo, 18, 5)
            cmd41.Parameters.Item(6).value = Mid(mRecibo, 31, 3)
            cmd41.Parameters.Item(7).value = Mid(mRecibo, 34, 6)
            cmd41.Parameters.Item(8).value = Mid(mRecibo, 40, 3)
            cmd41.Parameters.Item(9).value = Mid(mRecibo, 43, 60)
            cmd41.Parameters.Item(10).value = Replace(Mid(mRecibo, 103, 10), ".", ",")
            cmd41.Parameters.Item(11).value = Mid(mRecibo, 113, 2)
            cmd41.Parameters.Item(12).value = Replace(Mid(mRecibo, 115, 10), ".", ",")
            cmd41.Parameters.Item(13).value = Replace(Mid(mRecibo, 125, 12), ".", ",")
            cmd41.Parameters.Item(14).value = Replace(Mid(mRecibo, 137, 12), ".", ",")
            cmd41.Parameters.Item(15).value = Replace(Mid(mRecibo, 149, 12), ".", ",")
            cmd41.Parameters.Item(16).value = Mid(mRecibo, 161, 1)
            cmd41.Parameters.Item(17).value = Mid(mRecibo, 162, 1)
            cmd41.Execute
        End If

        If Mid(mRecibo, 1, 3) = "051" Then
            cmd51.Parameters.Refresh

            cmd51.Parameters.Item(1).value = Mid(mRecibo, 4, 4)
            cmd51.Parameters.Item(2).value = Mid(mRecibo, 8, 2)
            cmd51.Parameters.Item(3).value = Mid(mRecibo, 10, 2)
            cmd51.Parameters.Item(4).value = Mid(mRecibo, 12, 6)
            cmd51.Parameters.Item(5).value = Mid(mRecibo, 18, 5)
            cmd51.Parameters.Item(6).value = Mid(mRecibo, 31, 3)
            cmd51.Parameters.Item(7).value = Mid(mRecibo, 34, 60)
            cmd51.Parameters.Item(8).value = Mid(mRecibo, 94, 60)
            cmd51.Parameters.Item(9).value = Mid(mRecibo, 154, 60)
            cmd51.Execute
        End If

        If Mid(mRecibo, 1, 3) = "052" Then
            cmd52.Parameters.Refresh

            cmd52.Parameters.Item(1).value = Mid(mRecibo, 4, 4)
            cmd52.Parameters.Item(2).value = Mid(mRecibo, 8, 2)
            cmd52.Parameters.Item(3).value = Mid(mRecibo, 10, 2)
            cmd52.Parameters.Item(4).value = Mid(mRecibo, 12, 6)
            cmd52.Parameters.Item(5).value = Mid(mRecibo, 18, 5)
            cmd52.Parameters.Item(6).value = Mid(mRecibo, 31, 3)
            cmd52.Parameters.Item(7).value = Mid(mRecibo, 34, 50)
            cmd52.Parameters.Item(8).value = Mid(mRecibo, 84, 50)
            cmd52.Parameters.Item(9).value = Mid(mRecibo, 134, 50)
            cmd52.Parameters.Item(10).value = Replace(Mid(mRecibo, 184, 12), ".", ",")
            cmd52.Parameters.Item(11).value = Replace(Mid(mRecibo, 196, 12), ".", ",")
            cmd52.Parameters.Item(12).value = Replace(Mid(mRecibo, 208, 12), ".", ",")
            cmd52.Parameters.Item(13).value = Replace(Mid(mRecibo, 222, 12), ".", ",")
            cmd52.Execute
        End If
        
        
    Wend

    Close #1

    cn.Close

    Set cn = Nothing
    Set cmd = Nothing
    Set cmd01 = Nothing
    Set cmd11 = Nothing
    Set cmd31 = Nothing
    Set cmd41 = Nothing
    Set cmd51 = Nothing
    Set cmd52 = Nothing

    Unload FrmProceso
    
End Function

Public Function GenerarArchivoSimplifReg(ByVal sPath As String, ByVal iCodigo As Integer)
    
    Dim sReg As String
    Dim Arch As Integer
    
    Dim cn As Connection
    Dim rs As Recordset
    
    Set cn = New Connection
    Set rs = New Recordset
    
    Arch = FreeFile
    Open sPath & "\SimplifReg.txt" For Output As Arch
    
    cn.Open SGEPOB
    rs.Open "Select * from Personal_Sueldos where Sucursal = " & iCodigo, cn, adOpenDynamic, adLockBatchOptimistic, adCmdText
    
    
    While Not rs.EOF
    
        sReg = "02MR"
        sReg = sReg & Right("00000000000" & Trim(Replace(rs!Cuil, "-", "")), 11)
        sReg = sReg & Mid(rs!Nacimiento, 5, 4) & "/" & Mid(rs!Nacimiento, 3, 2) & "/" & Mid(rs!Nacimiento, 1, 2)
        sReg = sReg & Space(3)                                      'Nivel de formación
        sReg = sReg & "N"                                           'Incapacidad S/N
        sReg = sReg & Space(4)                                      'codigo pais telefono
        sReg = sReg & Space(4)                                      'codigo area telefono
        sReg = sReg & Space(15)                                     'telefono
        sReg = sReg & Space(1)                                      'tipo de telefono
        sReg = sReg & Right(Space(60) & rs!Mail, 60)
        sReg = sReg & "2"                                           'tipo de mail (2=personal)
        sReg = sReg & rs!Cbu
        sReg = sReg & Space(57)                                     'datos domicilio
        sReg = sReg & Space(35)                                     'cod postal localidad provincia
    
        Print #Arch, sReg
        rs.MoveNext
        
    Wend
    
    Close Arch
    
    rs.Close
    cn.Close
    
    Set rs = Nothing
    Set cn = Nothing

End Function
