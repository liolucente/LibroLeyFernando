VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form frmActualizarDatos 
   Caption         =   "Generar recibos PDFs"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImpresion 
      Caption         =   "Generar Recibos e imprimirlos"
      Height          =   615
      Left            =   4800
      TabIndex        =   3
      Top             =   2880
      Width           =   1335
   End
   Begin ComctlLib.ProgressBar pbar 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Generar Recibos y PDFs"
      Height          =   615
      Left            =   2760
      TabIndex        =   0
      Top             =   2880
      Width           =   1335
   End
   Begin DDSharpGridOLEDB2.SGGrid sgEmpresas 
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9135
      _cx             =   16113
      _cy             =   3836
      DataMember      =   ""
      DataMode        =   1
      AutoFields      =   -1  'True
      Enabled         =   -1  'True
      GridBorderStyle =   1
      ScrollBars      =   3
      FlatScrollBars  =   0
      ScrollBarTrack  =   0   'False
      DataRowCount    =   2
      BeginProperty HeadingFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataColCount    =   6
      HeadingRowCount =   1
      HeadingColCount =   1
      TextAlignment   =   0
      WordWrap        =   0   'False
      Ellipsis        =   1
      HeadingBackColor=   -2147483633
      HeadingForeColor=   -2147483630
      HeadingTextAlignment=   0
      HeadingWordWrap =   0   'False
      HeadingEllipsis =   1
      GridLines       =   1
      HeadingGridLines=   2
      GridLinesColor  =   -2147483633
      HeadingGridLinesColor=   -2147483632
      EvenOddStyle    =   0
      ColorEven       =   -2147483628
      ColorOdd        =   -2147483624
      UserResizeAnimate=   1
      UserResizing    =   3
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      UserDragging    =   2
      UserHiding      =   2
      CellPadding     =   15
      CellBkgStyle    =   1
      CellBackColor   =   -2147483643
      CellForeColor   =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   1
      FocusRectColor  =   0
      FocusRectLineWidth=   1
      TabKeyBehavior  =   0
      EnterKeyBehavior=   0
      NavigationWrapMode=   1
      SkipReadOnly    =   0   'False
      DefaultColWidth =   1200
      DefaultRowHeight=   255
      CellsBorderColor=   0
      CellsBorderVisible=   -1  'True
      RowNumbering    =   0   'False
      EqualRowHeight  =   0   'False
      EqualColWidth   =   0   'False
      HScrollHeight   =   0
      VScrollWidth    =   0
      Format          =   "General"
      Appearance      =   2
      FitLastColumn   =   -1  'True
      SelectionMode   =   0
      MultiSelect     =   0
      AllowAddNew     =   0   'False
      AllowDelete     =   0   'False
      AllowEdit       =   -1  'True
      ScrollBarTips   =   0
      CellTips        =   0
      CellTipsDelay   =   1000
      SpecialMode     =   0
      OutlineLines    =   1
      CacheAllRecords =   -1  'True
      ColumnClickSort =   0   'False
      PreviewPaneColumn=   ""
      PreviewPaneType =   0
      PreviewPanePosition=   2
      PreviewPaneSize =   2000
      GroupIndentation=   225
      InactiveSelection=   1
      AutoScroll      =   -1  'True
      AutoResize      =   0
      AutoResizeHeadings=   0   'False
      OLEDragMode     =   0
      OLEDropMode     =   0
      Caption         =   ""
      ScrollTipColumn =   ""
      MaxRows         =   4194304
      MaxColumns      =   8192
      NewRowPos       =   1
      CustomBkgDraw   =   0
      AutoGroup       =   0   'False
      GroupByBoxVisible=   0   'False
      GroupByBoxText  =   "Drag a column header here to group by that column"
      AlphaBlendEnabled=   0   'False
      DragAlphaLevel  =   206
      AutoSearch      =   0
      AutoSearchDelay =   2000
      StylesCollection=   $"frmActualizarDatos.frx":0000
      ColumnsCollection=   $"frmActualizarDatos.frx":1DA5
      ValueItems      =   $"frmActualizarDatos.frx":3D7C
   End
End
Attribute VB_Name = "frmActualizarDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub LlenarEmpresas()

    Dim cn As Connection
    Dim cn2 As Connection
    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim sSQL As String
    
    Set cn = New Connection
    Set cn2 = New Connection
    
    Set rs = New Recordset
    Set rs2 = New Recordset
    
    rs.CursorLocation = adUseClient
    rs2.CursorLocation = adUseClient
    
    cn.Open SGINST
    rs.Open "Select * from sgEmpresa where EmprOrden > 100 order by EmprOrden", cn, adOpenDynamic, adLockBatchOptimistic, adCmdText
    
    sgEmpresas.DataRowCount = 0
    
    While Not rs.EOF
        
        sgEmpresas.DataRowCount = sgEmpresas.DataRowCount + 1
        
        With sgEmpresas.Rows.At(sgEmpresas.DataRowCount)
            .Cells(1).value = False
            .Cells(2).value = rs!EMPRNOMBRE
            .Cells(3).value = rs!EMPR_ID
            .Cells(4).value = rs!EmprCarpeta
            .Cells(5).value = rs!EmprOrden
            .Cells(6).value = rs!EmprBaseDato
        End With
        
        If Not ArchivoExiste(Trim(rs!EmprCarpeta) & "\IfRecibo.txt") Then
            sgEmpresas.Rows.At(sgEmpresas.DataRowCount).Style.BackColor = &HC0C0C0
            sgEmpresas.Rows.At(sgEmpresas.DataRowCount).Cells(1).ReadOnly = True
        Else
            sgEmpresas.Rows.At(sgEmpresas.DataRowCount).Style.BackColor = &H80000005
            sgEmpresas.Rows.At(sgEmpresas.DataRowCount).Cells(1).ReadOnly = False
        End If
        
        rs.MoveNext
    
    Wend
    
    rs.Close
    cn.Close
 
    cn2.Open SGINST
    rs2.Open "Select * from SGRXEMPRUSUA where usua_id = " & user.USUA_ID, cn2, adOpenDynamic, adLockBatchOptimistic, adCmdText
                    
    bHabilitado = True
    bApto = True
    
    For i = 1 To sgEmpresas.DataRowCount
        besta = False
        rs2.MoveFirst
        While Not rs2.EOF
            If rs2!EMPR_ID = Val(sgEmpresas.Rows.At(i).Cells(3).value) And sgEmpresas.Rows.At(i).Style.BackColor <> &HC0C0C0 Then
                besta = True
                sgEmpresas.Rows.At(i).Style.BackColor = &H80000005
                sgEmpresas.Rows.At(i).Cells(1).ReadOnly = False
            End If
            rs2.MoveNext
        Wend
        If besta = False Then
            bHabilitado = False
            sgEmpresas.Rows.At(i).Style.BackColor = &HC0C0C0
            sgEmpresas.Rows.At(i).Cells(1).ReadOnly = True
        End If
    Next
    
    rs2.Close
    cn2.Close
    
    Set rs = Nothing
    Set cn = Nothing
    
    Set rs2 = Nothing
    Set cn2 = Nothing
    
End Sub


Private Sub cmdActualizar_Click()
    
    Dim cn As Connection
    Dim rs As Recordset
    
    Set cn = New Connection
    Set rs = New Recordset
    
    cn.Open SGEPOB
    
    For i = 1 To sgEmpresas.DataRowCount
        
        If sgEmpresas.Rows.At(i).Cells(1).value = True Then
            
            CargarDatosConRecibos sgEmpresas.Rows.At(i).Cells(4).value, sgEmpresas.Rows.At(i).Cells(5).value
            
            rs.Open "Select top 1 * from Legajos_Claves where rtrim(legajo)='E" & Right("00" & Trim(Str(sgEmpresas.Rows.At(i).Cells(3).value - 100)), 2) & "' order by FechaDesde desc", cn, adOpenDynamic, adLockBatchOptimistic, adCmdText
            If Not rs.EOF Then
                GenerarPDFS sgEmpresas.Rows.At(i).Cells(4).value, sgEmpresas.Rows.At(i).Cells(5).value, sgEmpresas.Rows.At(i).Cells(6).value, rs!Clave
            Else
                GenerarPDFS sgEmpresas.Rows.At(i).Cells(4).value, sgEmpresas.Rows.At(i).Cells(5).value, sgEmpresas.Rows.At(i).Cells(6).value, ""
            End If
            rs.Close
            
            cn.Execute "Update Recibos_Reg01 set Reg01_Impreso = 'S' where Reg01_Impreso = 'N' and Reg01_Empresa = " & sgEmpresas.Rows.At(i).Cells(5).value
        End If
    Next
    
    cn.Close
    
    MsgBox "Los recibos PDFs fueron generados con èxito", vbInformation, "Información"
    
    
End Sub

Private Sub GenerarPDFS(ByVal sRuta As String, ByVal idEmpr As Integer, ByVal sRutaDest As String, ByVal sClave As String)
    
    Dim sSQL As String
    Dim repo As Report
    Dim crapp As CRAXDRT.Application
    Dim cn As Connection
    Dim rs As Recordset
    Dim rsLega As Recordset
    Dim rsClave As Recordset
    
    Dim cnInte As Connection
    Dim cmd As Command
    
    Dim i As Integer
    
    Set cn = New Connection
    Set rs = New Recordset
    Set rsLega = New Recordset
    Set rsClave = New Recordset
    
    Set cnInte = New Connection
    Set cmd = New Command
    Set cmd2 = New Command 'Este lo usamos para guardar los datos en las tablas sgedemo
    
    cnInte.Open SGINTE
    cn.Open SGEPOB
    
    cmd.ActiveConnection = cnInte
    'Original INTERFAZ_RECIBOS_DIGITALES_INSERT
    cmd.CommandText = "INTERFAZ_RECIBOS_DIGITALES_INSERT"
    cmd.CommandType = adCmdStoredProc
    
    cmd2.ActiveConnection = cn
    cmd2.CommandText = "RECIBOS_DIGITALES_INSERT"
    cmd2.CommandType = adCmdStoredProc
    
    rs.CursorLocation = adUseClient
    rsLega.CursorLocation = adUseClient
    
    sCarpeta = sRuta

    sSQL = "Select * from Recibos_Reg31 where Reg31_Empresa = " & idEmpr & " and Reg31_Impreso = 'N' order by Reg31_legajo"
    rs.Open sSQL, cn, adOpenDynamic, adLockBatchOptimistic, adCmdText
    
    sSQL = "Delete from Interfaz_Recibos_Digitales where Tipo_Liquidacion = " & rs!Reg31_TipoLiq
    sSQL = sSQL & " and Anio = " & rs!Reg31_Anio
    sSQL = sSQL & " and Mes = " & rs!Reg31_Mes
    sSQL = sSQL & " and Lote = " & rs!Reg31_NroLiq
    cnInte.Execute sSQL

    While Not rs.EOF
        
        sSQL = "Select * from Recibos_Tratamiento where convert(varchar(5),Reg31_legajo) = '" & rs!Reg31_Legajo & "'"
        sSQL = sSQL & " and Reg31_Anio = " & rs!Reg31_Anio
        sSQL = sSQL & " and Reg31_Mes = " & rs!Reg31_Mes
        sSQL = sSQL & " and Reg31_TipoLiq = " & rs!Reg31_TipoLiq
        sSQL = sSQL & " and Reg31_NroLiq = " & rs!Reg31_NroLiq
        sSQL = sSQL & " and Reg31_Orden = " & rs!Reg31_Orden
        sSQL = sSQL & " order by Reg41_Orden"
        
        rsLega.Open sSQL, cn, adOpenDynamic, adLockBatchOptimistic, adCmdText
                
        bImprimo = False
        
        If Len(Trim(sClave)) > 0 Then
            rsClave.Open "Select top 1 * from Legajos_Claves where convert(varchar(5),Legajo) = '" & Right("00000" & rs!Reg31_Legajo, 5) & "' order by FechaDesde desc", cn, adOpenDynamic, adLockBatchOptimistic, adCmdText
            If Not rsClave.EOF Then
                MostrarReportePDF "crReciboHospital.rpt", rsLega, Trim(rsLega!Reg31_NombrePDF) & ".pdf", sRutaDest, "Recibo Hospital", DesEncripta(sClave), DesEncripta(rsClave!Clave)
                bImprimo = True
            End If
            rsClave.Close
        Else
            MostrarReportePDF "crReciboHospital.rpt", rsLega, Trim(rsLega!Reg31_NombrePDF) & ".pdf", sRutaDest, "Recibo Hospital", "", ""
            bImprimo = True
        End If
        
        If bImprimo = True Then
            sSQL = "Update Recibos_Reg31 set Reg31_Impreso = 'S' where Reg31_legajo = " & rs!Reg31_Legajo
            sSQL = sSQL & " and Reg31_Anio = " & rs!Reg31_Anio
            sSQL = sSQL & " and Reg31_Mes = " & rs!Reg31_Mes
            sSQL = sSQL & " and Reg31_TipoLiq = " & rs!Reg31_TipoLiq
            sSQL = sSQL & " and Reg31_NroLiq = " & rs!Reg31_NroLiq
            sSQL = sSQL & " and Reg31_Orden = " & rs!Reg31_Orden
            cn.Execute sSQL
            
            rsLega.MoveFirst
            
            cmd.Parameters.Refresh
            cmd.Parameters.Item(1).value = 0
            cmd.Parameters.Item(2).value = rsLega!Reg31_Legajo
            cmd.Parameters.Item(3).value = Replace(rsLega!Reg31_Cuil, "-", "")
            cmd.Parameters.Item(4).value = rsLega!Reg31_ApeNom
            cmd.Parameters.Item(5).value = rsLega!Reg31_Anio
            cmd.Parameters.Item(6).value = rsLega!Reg31_Mes
            cmd.Parameters.Item(7).value = rsLega!Reg01_TipoLiqCodigo
            cmd.Parameters.Item(8).value = rsLega!Reg01_TipoLiqDescripcion
            cmd.Parameters.Item(9).value = rsLega!Reg01_TipoLiqAbreviatura
            cmd.Parameters.Item(10).value = rsLega!Reg31_NroLiq
            cmd.Parameters.Item(11).value = CDate(rsLega!Reg01_FechaGen)
            cmd.Parameters.Item(12).value = Date
            cmd.Parameters.Item(13).value = sRutaDest
            cmd.Parameters.Item(14).value = Trim(rsLega!Reg31_NombrePDF) & ".pdf"
            cmd.Parameters.Item(15).value = Date
            cmd.Parameters.Item(16).value = Date
            cmd.Execute
            
        End If
        
        rsLega.Close
        rs.MoveNext
    
    Wend
        
    rs.Close
    cn.Close
    cnInte.Close
    
    Set rs = Nothing
    Set cn = Nothing
    Set cnInte = Nothing
    Set cmd = Nothing
    
End Sub
Private Sub GenerarPDFS_Temporales(ByVal sRuta As String, ByVal idEmpr As Integer, ByVal sRutaDest As String, ByVal sClave As String)
    
    Dim sSQL As String
    Dim repo As Report
    Dim crapp As CRAXDRT.Application
    Dim cn As Connection
    Dim rs As Recordset
    Dim rsLega As Recordset
    Dim rsClave As Recordset
    
    Dim cnInte As Connection
    Dim cmd As Command
    
    Dim i As Integer
    
    Set cn = New Connection
    Set rs = New Recordset
    Set rsLega = New Recordset
    Set rsClave = New Recordset
    
    Set cnInte = New Connection
    Set cmd = New Command
    Set cmd2 = New Command 'Este lo usamos para guardar los datos en las tablas sgedemo
    
    cnInte.Open SGINTE
    cn.Open SGEPOB
    
    cmd.ActiveConnection = cnInte
    'Original INTERFAZ_RECIBOS_DIGITALES_INSERT
    cmd.CommandText = "INTERFAZ_RECIBOS_DIGITALES_INSERT"
    cmd.CommandType = adCmdStoredProc
    
    cmd2.ActiveConnection = cn
    cmd2.CommandText = "RECIBOS_DIGITALES_INSERT"
    cmd2.CommandType = adCmdStoredProc
    
    rs.CursorLocation = adUseClient
    rsLega.CursorLocation = adUseClient
    
    sCarpeta = sRuta

    sSQL = "Select * from Recibos_Reg31 where Reg31_Empresa = " & idEmpr & " and Reg31_Impreso = 'N' order by Reg31_legajo"
    rs.Open sSQL, cn, adOpenDynamic, adLockBatchOptimistic, adCmdText
    
    sSQL = "Delete from Interfaz_Recibos_Digitales where Tipo_Liquidacion = " & rs!Reg31_TipoLiq
    sSQL = sSQL & " and Anio = " & rs!Reg31_Anio
    sSQL = sSQL & " and Mes = " & rs!Reg31_Mes
    sSQL = sSQL & " and Lote = " & rs!Reg31_NroLiq
    cnInte.Execute sSQL

    While Not rs.EOF
        
        sSQL = "Select * from Recibos_Tratamiento where convert(varchar(5),Reg31_legajo) = '" & rs!Reg31_Legajo & "'"
        sSQL = sSQL & " and Reg31_Anio = " & rs!Reg31_Anio
        sSQL = sSQL & " and Reg31_Mes = " & rs!Reg31_Mes
        sSQL = sSQL & " and Reg31_TipoLiq = " & rs!Reg31_TipoLiq
        sSQL = sSQL & " and Reg31_NroLiq = " & rs!Reg31_NroLiq
        sSQL = sSQL & " and Reg31_Orden = " & rs!Reg31_Orden
        sSQL = sSQL & " order by Reg41_Orden"
        
        rsLega.Open sSQL, cn, adOpenDynamic, adLockBatchOptimistic, adCmdText
                
        bImprimo = False
        
        ImprimirReportePDF "crReciboHospital.rpt", rsLega, "temporal.pdf", sRutaDest, "Recibo Hospital", "", ""
        bImprimo = True
        
        If bImprimo = True Then
            sSQL = "Update Recibos_Reg31 set Reg31_Impreso = 'S' where Reg31_legajo = " & rs!Reg31_Legajo
            sSQL = sSQL & " and Reg31_Anio = " & rs!Reg31_Anio
            sSQL = sSQL & " and Reg31_Mes = " & rs!Reg31_Mes
            sSQL = sSQL & " and Reg31_TipoLiq = " & rs!Reg31_TipoLiq
            sSQL = sSQL & " and Reg31_NroLiq = " & rs!Reg31_NroLiq
            sSQL = sSQL & " and Reg31_Orden = " & rs!Reg31_Orden
            cn.Execute sSQL
            
            rsLega.MoveFirst
            
            cmd.Parameters.Refresh
            cmd.Parameters.Item(1).value = 0
            cmd.Parameters.Item(2).value = rsLega!Reg31_Legajo
            cmd.Parameters.Item(3).value = Replace(rsLega!Reg31_Cuil, "-", "")
            cmd.Parameters.Item(4).value = rsLega!Reg31_ApeNom
            cmd.Parameters.Item(5).value = rsLega!Reg31_Anio
            cmd.Parameters.Item(6).value = rsLega!Reg31_Mes
            cmd.Parameters.Item(7).value = rsLega!Reg01_TipoLiqCodigo
            cmd.Parameters.Item(8).value = rsLega!Reg01_TipoLiqDescripcion
            cmd.Parameters.Item(9).value = rsLega!Reg01_TipoLiqAbreviatura
            cmd.Parameters.Item(10).value = rsLega!Reg31_NroLiq
            cmd.Parameters.Item(11).value = CDate(rsLega!Reg01_FechaGen)
            cmd.Parameters.Item(12).value = Date
            cmd.Parameters.Item(13).value = sRutaDest
            cmd.Parameters.Item(14).value = Trim(rsLega!Reg31_NombrePDF) & ".pdf"
            cmd.Parameters.Item(15).value = Date
            cmd.Parameters.Item(16).value = Date
            cmd.Execute
            
        End If
        
        rsLega.Close
        rs.MoveNext
    
    Wend
        
    rs.Close
    cn.Close
    cnInte.Close
    
    Set rs = Nothing
    Set cn = Nothing
    Set cnInte = Nothing
    Set cmd = Nothing
    
End Sub

Private Sub cmdImpresion_Click()

    Dim cn As Connection
    Dim sCmdLine As String
    Dim idprog As Long, iExit As Long
    
    Set cn = New Connection
    
    cn.Open SGEPOB
    
    For i = 1 To sgEmpresas.DataRowCount
        
        If sgEmpresas.Rows.At(i).Cells(1).value = True Then
            
            CargarDatosConRecibos sgEmpresas.Rows.At(i).Cells(4).value, sgEmpresas.Rows.At(i).Cells(5).value
            GenerarPDFS_Temporales sgEmpresas.Rows.At(i).Cells(4).value, sgEmpresas.Rows.At(i).Cells(5).value, sgEmpresas.Rows.At(i).Cells(6).value, ""
            cn.Execute "Update Recibos_Reg01 set Reg01_Impreso = 'S' where Reg01_Impreso = 'N' and Reg01_Empresa = " & sgEmpresas.Rows.At(i).Cells(5).value
            
        End If
    Next
    
    cn.Close
    
    
    sCmdLine = "DELETE C:\TEMP\TEMPORAL.PDF"
    idprog = Shell(sCmdLine)
'    iExit = fWait(idprog)
    
    

End Sub

Private Sub Form_Load()
    
    sgEmpresas.Columns(1).Control.Type = sgCellCheckBox
    sgEmpresas.Columns(1).Style.PictureAlignment = sgPicAlignCenterCenter
    
    LlenarEmpresas
    
End Sub
