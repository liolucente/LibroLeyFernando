VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form frmActualizador 
   Caption         =   "Actualizar datos"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSimplifReg 
      Caption         =   "Datos complementarios Simplificación Registral"
      Height          =   615
      Left            =   7440
      TabIndex        =   3
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar datos desde Sueldos"
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   2880
      Width           =   1335
   End
   Begin ComctlLib.ProgressBar pbar 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
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
      DataColCount    =   5
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
      StylesCollection=   $"frmActualizador.frx":0000
      ColumnsCollection=   $"frmActualizador.frx":1DA5
      ValueItems      =   $"frmActualizador.frx":38F7
   End
End
Attribute VB_Name = "frmActualizador"
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
        End With
        
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
    
    
    For i = 1 To sgEmpresas.DataRowCount
        
        If sgEmpresas.Rows.At(i).Cells(1).value = True Then
            CargarDatos sgEmpresas.Rows.At(i).Cells(4).value, sgEmpresas.Rows.At(i).Cells(5).value
        End If
    Next
    
    cn.Open SGEPOB
    cn.Close
    
End Sub

Private Sub cmdSimplifReg_Click()

    For i = 1 To sgEmpresas.DataRowCount
        
        If sgEmpresas.Rows.At(i).Cells(1).value = True Then
            GenerarArchivoSimplifReg sgEmpresas.Rows.At(i).Cells(4).value, sgEmpresas.Rows.At(i).Cells(5).value
        End If
    Next
    

End Sub

Private Sub Form_Load()
    
    sgEmpresas.Columns(1).Control.Type = sgCellCheckBox
    sgEmpresas.Columns(1).Style.PictureAlignment = sgPicAlignCenterCenter
    
    LlenarEmpresas
    
End Sub

