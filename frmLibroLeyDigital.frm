VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form frmLibroLeyDigital 
   Caption         =   "Libro Ley Digital"
   ClientHeight    =   4770
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9390
   LinkTopic       =   "Form3"
   ScaleHeight     =   4770
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8160
      TabIndex        =   20
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "    Eliminar liquidación      (no presentada en AFIP)"
      Height          =   615
      Left            =   4920
      TabIndex        =   17
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar liquidación (Aceptada en AFIP)"
      Height          =   615
      Left            =   2640
      TabIndex        =   16
      Top             =   3960
      Width           =   1935
   End
   Begin DDSharpGridOLEDB2.SGGrid sgEmpresas 
      Height          =   2175
      Left            =   120
      TabIndex        =   14
      Top             =   1560
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
      StylesCollection=   $"frmLibroLeyDigital.frx":0000
      ColumnsCollection=   $"frmLibroLeyDigital.frx":1DA5
      ValueItems      =   $"frmLibroLeyDigital.frx":3D8A
   End
   Begin MSComCtl2.DTPicker dtpFecPago 
      Height          =   375
      Left            =   7440
      TabIndex        =   12
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   54132737
      CurrentDate     =   43803
   End
   Begin VB.TextBox txtNumero 
      Height          =   405
      Left            =   5880
      MaxLength       =   5
      TabIndex        =   10
      Top             =   480
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de liquidación"
      Height          =   1335
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton optSemanal 
         Caption         =   "Semanal"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optQuincenal 
         Caption         =   "Quincenal"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optMensual 
         Caption         =   "Mensual"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComCtl2.DTPicker dtpPeriodo 
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MM/yyyy"
      Format          =   54132739
      CurrentDate     =   43802
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de envío"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton optRE 
         Caption         =   "Rectificativa F931"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton optSJ 
         Caption         =   "Liquidación y F931"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "Generar Archivo para AFIP"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   3960
      Width           =   1935
   End
   Begin ComctlLib.ProgressBar pbar 
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3600
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label txtUsuario 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label5"
      Height          =   255
      Left            =   8640
      TabIndex        =   19
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label txtEstado 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "En proceso Sueldos"
      Height          =   375
      Left            =   5160
      TabIndex        =   18
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "Empresas"
      Height          =   255
      Left            =   2160
      TabIndex        =   15
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha de pago"
      Height          =   255
      Left            =   7440
      TabIndex        =   13
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Número Liq"
      Height          =   255
      Left            =   5880
      TabIndex        =   11
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Período"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmLibroLeyDigital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCerrar_Click()
    
    Dim cn As Connection
    Dim sSQL As String
    
    Set cn = New Connection
    
    cn.Open SGINST
    
    For i = 1 To sgEmpresas.DataRowCount
        If sgEmpresas.Rows.At(i).Cells(1).value = True Then
            sSQL = "Update SgEmpresa set EmprTipo = 'S', EmprComenta = '" & Right("000" & Trim(Str(txtNumero.Text)), 3) & " - "
            sSQL = sSQL & Mid(dtpPeriodo.value, 1, 5) & "', EmprUsuari = '" & user.USUACODIGO & "'"
            sSQL = sSQL & " where Empr_Id = " & sgEmpresas.Rows.At(i).Cells(3).value
            cn.Execute sSQL
        End If
    Next
    
    cn.Close
    
    cn.Open SGEPOB
    sSQL = "update LibroLey_Historial set Terminada = 'S' where Terminada = 'N'"
    cn.Execute sSQL
    
    cn.Close
    Set cn = Nothing
    
End Sub

Private Sub cmdEliminar_Click()
    
    Dim cn As Connection
    Set cn = New Connection
    
    cn.Open SGEPOB
    
    If MsgBox("Confirma que elimina la presentación en proceso?", vbYesNo, "Confirmación") = vbYes Then
        cn.Execute "Delete from LibroLey_Historial where Terminada='N'"
    End If
    
    cn.Close
    Set cn = Nothing
    
End Sub

Private Sub cmdProcesar_Click()
    
    Dim cn As Connection
    Dim rs As Recordset
    Dim rs3 As Recordset
    Dim rs4 As Recordset
    Dim sSQL As String
    
    Dim sCUIT As String
    Dim sFechaPago As String
    Dim sLegajo As String
    Dim iCuantos As Integer
    
    Dim sReg01 As String
    Dim sReg02 As String
    Dim sReg03 As String
    Dim sReg04 As String
    
    Dim oSueConc As Sue118
    Dim oSueR4 As Suelld4
    
    Dim sArchivo As String
    Dim Arch As Integer
    
    Dim dIngreso As Date
    Dim dEgreso As Date
    
    Set cn = New Connection
    Set rs = New Recordset
    Set rs3 = New Recordset
    Set rs4 = New Recordset
    
    iCuantos = 0
    
    cn.Open SGEPOB
    rs.Open "SELECT * FROM TABLA_GENERAL WHERE GRAL_TABLA = 110001 AND GRAL_ITEM = 1", cn, adOpenDynamic, adLockBatchOptimistic, adCmdText
    If rs.EOF Then
        MsgBox "Falta definir el parametro de CUIT de la empresa. No se puede continuar.", vbCritical, "Información"
        Exit Sub
        rs.Close
        cn.Close
    Else
        sCUIT = rs!GRAL_DATO3
        sCUIT = Replace(sCUIT, "-", "")
        rs.Close
    End If
    
    cn.Execute "Delete from LibroLey_Reg02"
    cn.Execute "Delete from LibroLey_Reg03"
    cn.Execute "Delete from LibroLey_Reg04"
    
    sReg01 = "01"
    sReg01 = sReg01 & Mid(sCUIT, 1, 11)
    
    If optSJ.value = True Then
        sReg01 = sReg01 & "SJ"
    Else
        sReg01 = sReg01 & "RE"
    End If
    
    sReg01 = sReg01 & Mid(ConvertirFecha(dtpPeriodo.value), 1, 6)
    
    If optRE.value = True Then
        sReg01 = sReg01 & " "
    Else
        If optMensual.value = True Then
            sReg01 = sReg01 & "M"
        Else
            If optQuincenal.value = True Then
                sReg01 = sReg01 & "Q"
            Else
                sReg01 = sReg01 & "S"
            End If
        End If
    End If
    
    sReg01 = sReg01 & IIf(optSJ.value = True, Right("00000" & Trim(txtNumero.Text), 5), Space(5))
    sReg01 = sReg01 & IIf(optSJ.value = True, "30", Space(2))
        
    For i = 1 To sgEmpresas.DataRowCount
    
        If sgEmpresas.Rows.At(i).Cells(1).value = True Then
            
            sCarpeta = sgEmpresas.Rows.At(i).Cells(6).value
            CargarDatos sgEmpresas.Rows.At(i).Cells(6).value, sgEmpresas.Rows.At(i).Cells(3).value

            Set oSueConc = New Sue118
            Set oSueR4 = New Suelld4

            sFechaPago = ConvertirFecha(dtpFecPago.value)
        
            While Not oSueConc.rsSue118.EOF
                
                sLegajo = oSueConc.rsSue118!legajo
                iCuantos = iCuantos + 1
            
                If rs.State = adStateOpen Then rs.Close
                sSQL = "SELECT PERSSUE.*, ISNULL(CBU,'') AS CBU, ISNULL(MAIL,'') AS MAIL FROM PERSSUE"
                sSQL = sSQL & " LEFT JOIN PERSANEX ON PERSSUE.LEGAJO = PERSANEX.LEGAJO"
                sSQL = sSQL & " WHERE PERSSUE.LEGAJO = " & sLegajo
                
                rs.Open sSQL, cn, adOpenDynamic, adLockBatchOptimistic, adCmdText
                If rs.EOF Then
                    MsgBox "No se encuentran los datos correspondientes al legajo " & sLegajo & Chr(13) & Chr(10) & "No se puede continuar", vbCritical
                    rs.Close
                    cn.Close
                    Exit Sub
                End If
            
                sReg02 = "02"
                sReg02 = sReg02 & Trim(Replace(rs!Cuil, "-", ""))
                
'                If rs!legajo = "25107" Then
'                    sReg02 = "02"
'                End If
                
                sReg02 = sReg02 & Right(Space(10) & Trim(oSueConc.rsSue118!legajo), 10)
                sReg02 = sReg02 & Space(50)
                
                If Trim(rs!Cbu) = "" Or Val(rs!Cbu) = 0 Then
                    sReg02 = sReg02 & Space(22)
                Else
                    sReg02 = sReg02 & Right(Trim(rs!Cbu), 22)
                End If
                
                'aca hay que ver si es liq de baja o si la persona entrò con el mes empezado
                dIngreso = CDate(Mid(rs!INGRESO, 1, 2) & "/" & Mid(rs!INGRESO, 3, 2) & "/" & Mid(rs!INGRESO, 5, 4))
                
                If ValidarFecha_txt(Trim(rs!Egreso), 1900, 2100) = True Then
                    dEgreso = CDate(Mid(rs!Egreso, 1, 2) & "/" & Mid(rs!Egreso, 3, 2) & "/" & Mid(rs!Egreso, 5, 4))
                Else
                    dEgreso = CDate("31/12/2100")
                End If
                
                iDiasNoTrab = 0
                If dIngreso > DateAdd("d", (Day(dtpPeriodo.value) - 1) * (-1), dtpPeriodo.value) Then
                    iDiasNoTrab = Day(dIngreso) - 1
'                    sReg02 = sReg02 & Right("000" & Trim(Str(DateDiff("d", dIngreso, DateAdd("d", -1, DateAdd("m", 1, DateAdd("d", (Day(dtpPeriodo.value) - 1) * (-1), dtpPeriodo.value)))))), 3)
                End If
                
                If dEgreso < DateAdd("d", -1, DateAdd("m", 1, DateAdd("d", (Day(dtpPeriodo.value) - 1) * (-1), dtpPeriodo.value))) Then
                    iDiasNoTrab = iDiasNoTrab + (30 - Day(dEgreso))
'                        sReg02 = sReg02 & Right("000" & Trim(Str(DateDiff("d", DateAdd("d", (Day(dtpPeriodo.value) - 1) * (-1), dtpPeriodo.value), dEgreso))), 3)
                End If
                
                If iDiasNoTrab = 0 Then
                    sReg02 = sReg02 & "000"
                Else
                    sReg02 = sReg02 & Right("000" & Trim(Str((30 - iDiasNoTrab))), 3)
                End If
                'ojo con el dato de arriba
        
                sReg02 = sReg02 & sFechaPago
                sReg02 = sReg02 & Space(8)
                                
                'Antes
                'sReg02 = sReg02 & "3"
                If Trim(rs!Cbu) = "" Or Val(rs!Cbu) = 0 Then
                    sReg02 = sReg02 & "1"
                Else
                    sReg02 = sReg02 & "3"
                End If

            
                cn.Execute "Insert into LibroLey_Reg02 values('" & Trim(sReg02) & "'," & Replace(rs!Cuil, "-", "") & ")"
            
                While oSueConc.rsSue118!legajo = sLegajo
                
                    If oSueConc.rsSue118!TIPOCONCEPTO <> 9 And oSueConc.rsSue118!TIPOCONCEPTO <> 0 And oSueConc.rsSue118!Concepto <> 259 Then
                
                        sReg03 = "03"
                        sReg03 = sReg03 & Trim(Replace(rs!Cuil, "-", ""))
                        sReg03 = sReg03 & Right(Space(10) & Trim(oSueConc.rsSue118!Concepto), 10)
                        sReg03 = sReg03 & Right("000" & Trim(Str(Int(Abs(Val(oSueConc.rsSue118!cantidad))))), 3) & Right("00" & Trim(Str(Abs(Val(oSueConc.rsSue118!cantidad)) - Abs(Int(Val(oSueConc.rsSue118!cantidad))))), 2)
'                        If (Abs(Val(oSueConc.rsSue118!cantidad)) * 100) - Int(Abs(Val(oSueConc.rsSue118!cantidad)) * 100) > 0 Then
'                            sReg03 = sReg03 & Right("00000" & Trim(Str(Abs(Round(Val(oSueConc.rsSue118!cantidad), 4)) * 10000)), 5)
'                        Else
'                            sReg03 = sReg03 & Right("00000" & Trim(Str(Abs(Val(oSueConc.rsSue118!cantidad)) * 100)), 5)
'                        End If
                        sReg03 = sReg03 & Space(1)
                        sReg03 = sReg03 & Right("000000000000000" & Trim(Str(Abs(Val(oSueConc.rsSue118!Importe) * 100))), 15)
                        If oSueConc.rsSue118!TIPOCONCEPTO < 6 Then
                            sReg03 = sReg03 & IIf(Val(oSueConc.rsSue118!Importe) > 0, "C", "D")
                        Else
                            sReg03 = sReg03 & IIf(Val(oSueConc.rsSue118!Importe) > 0, "D", "C")
                        End If
                        
                        'de donde saco el periodo por el que es el retroactivo???
'                        sReg03 = sReg03 & IIf(oSueConc.rsSue118!Caracter = 2, "MMAAAA", Space(6))
                        sReg03 = sReg03 & Space(6)
                        'ojo con el dato de arriba
                        
                        cn.Execute "Insert into LibroLey_Reg03 values('" & sReg03 & "'," & Replace(rs!Cuil, "-", "") & ")"
    
                    End If
                
                    oSueConc.rsSue118.MoveNext
                    
                    If oSueConc.rsSue118.EOF Then
                        GoTo salgo
                    End If
            
                Wend
            
            Wend
        
salgo:
        
            While Not oSueR4.rsReg04.EOF
                
                cn.Execute "Insert into LibroLey_Reg04 values ('" & oSueR4.rsReg04!Reg04 & "'," & Mid(oSueR4.rsReg04!Reg04, 3, 11) & ")"
                oSueR4.rsReg04.MoveNext
        
            Wend
    
        End If
    
    Next
    
    'aca tengo que poner cuantos empleados hay en la liquidacion
    sReg01 = sReg01 & Right("000000" & Trim(Str(iCuantos)), 6)
    
    Arch = FreeFile
    Open sCarpeta & "\LibroLey.txt" For Output As Arch
     
    Print #Arch, sReg01
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from LibroLey_Reg02 order by Cuil", cn, adOpenDynamic, adLockBatchOptimistic, adCmdText
    
    While Not rs.EOF
    
        Print #Arch, rs!Registro02
        rs3.Open "Select * from LibroLey_Reg03 where Cuil = " & rs!Cuil, cn, adOpenDynamic, adLockBatchOptimistic, adCmdText
        
        While Not rs3.EOF
            
            Print #Arch, rs3!Registro03
            rs3.MoveNext
            
        Wend
        rs3.Close
        
        rs.MoveNext
        
    Wend
    
    rs4.Open "Select * from LibroLey_Reg04 order by Cuil", cn, adOpenDynamic, adLockBatchOptimistic, adCmdText
        
    rs.MoveFirst
    While Not rs.EOF
    
        rs4.MoveFirst
        While Not rs4.EOF
        
            If rs!Cuil = rs4!Cuil Then
                Print #Arch, rs4!Registro04
                GoTo Avanzo
            Else
                rs4.MoveNext
            End If
        
        Wend
Avanzo:
        rs.MoveNext
    Wend
    
    rs4.Close
    rs.Close
    
    Close Arch
    
    rs.Open "Select * from LibroLey_Historial where Terminada='N'", cn, adOpenDynamic, adLockBatchOptimistic, adCmdText
    If Not rs.EOF Then
        sSQL = "Update LibroLey_Historial set FechaPago = Convert(datetime,'" & dtpFecPago.value & "',103), "
        sSQL = sSQL & "Usuario = '" & user.USUACODIGO & "' where Terminada='N'"
        cn.Execute sSQL
    Else
        sSQL = "Insert into LibroLey_Historial values("
        sSQL = sSQL & "Convert(datetime, '" & DateAdd("d", (Day(dtpPeriodo.value) - 1) * (-1), dtpPeriodo.value) & "',103), "
        sSQL = sSQL & txtNumero.Text & ", "
        sSQL = sSQL & "Convert(datetime, '" & dtpFecPago.value & "',103), "
        For i = 1 To sgEmpresas.DataRowCount
            If sgEmpresas.Rows.At(i).Cells(1).value Then
                sSQL = sSQL & sgEmpresas.Rows.At(i).Cells(3).value & ","
            End If
        Next
        sSQL = sSQL & "'N', '" & user.USUACODIGO & "') "
        cn.Execute sSQL
    End If
    
    MsgBox "Archivo generado con éxito"
    
    rs.Close
    cn.Close
    
    Set rs = Nothing
    Set rs3 = Nothing
    Set rs4 = Nothing
    Set cn = Nothing
    
End Sub



Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    optSJ.value = True
    optMensual.value = True
    
    sgEmpresas.Columns(1).Control.Type = sgCellCheckBox
    sgEmpresas.Columns(1).Style.PictureAlignment = sgPicAlignCenterCenter
    
    LlenarEmpresas
    
End Sub
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
            .Cells(4).value = rs!EMPRCOMENTA
            .Cells(5).value = rs!EmprUsuari
            .Cells(6).value = rs!EmprCarpeta
        End With
        
        rs.MoveNext
    Wend
    
    rs.Close
    cn.Close
    
    cn.Open SGEPOB
    rs.Open "Select * from LibroLey_Historial where Terminada='N'", cn, adOpenDynamic, adLockBatchOptimistic, adCmdText
    
    If Not rs.EOF Then
        
        dtpPeriodo.value = rs!Periodo
        txtNumero.Text = rs!Numero
        dtpFecPago.value = rs!FechaPago
        txtUsuario.caption = rs!Usuario
        txtEstado.caption = "En Proceso" & Space(2) & rs!Usuario
            
        For i = 1 To sgEmpresas.DataRowCount
            
            If InStr(Trim(sgEmpresas.Rows.At(i).Cells(3).value), rs!Empresas) > 0 Then
                sgEmpresas.Rows.At(i).Cells(1).value = True
            End If
            sgEmpresas.Rows.At(i).Cells(1).ReadOnly = True
            
        Next
        
    Else
        
        rs.Close
        rs.Open "Select max(Numero) as UltNum, max(Periodo) as UltPeriod from LibroLey_Historial", cn, adOpenDynamic, adLockBatchOptimistic, adCmdText
        
        If DateAdd("m", 1, rs!UltPeriod) > Date Then
            dtpPeriodo.value = rs!UltPeriod
            txtNumero.Text = rs!UltNum + 1
        Else
            dtpPeriodo.value = DateSerial(Year(Date), Month(Date) + 0, 1)
            txtNumero.Text = 1
        End If
        
        dtpFecPago.value = Date - 10
        txtEstado.caption = "A generar" & Space(2) & user.USUACODIGO
            
    End If
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
            If rs2!EMPR_ID = Val(sgEmpresas.Rows.At(i).Cells(3).value) Then
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
            If sgEmpresas.Rows.At(i).Cells(1).value = True Then
                bApto = False
            End If
        End If
    Next
    
    If bApto = False Then
        
        sSQL = "El libro ley que se está trabajando tiene empresas a las que su usuario no puede acceder." & Chr(13) & Chr(10)
        sSQL = sSQL & "No puede continuar con el proceso hasta que esta presentación se cierre"
        
        MsgBox sSQL, vbCritical, "Información"
        
        cmdProcesar.Enabled = False
        cmdCerrar.Enabled = False
        cmdEliminar.Enabled = False
    Else
        
        cmdProcesar.Enabled = True
        cmdCerrar.Enabled = True
        cmdEliminar.Enabled = True
    
    End If
    
    rs2.Close
    cn2.Close
    
    Set rs = Nothing
    Set cn = Nothing
    
    Set rs2 = Nothing
    Set cn2 = Nothing
    
End Sub

