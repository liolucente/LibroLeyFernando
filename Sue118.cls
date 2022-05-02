VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sue118"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public rsSue118 As ADODB.Recordset

Private Sub Class_GetDataMember(DataMember As String, Data As Object)
  Set Data = rsSue118
End Sub

Private Sub Class_Initialize()

   Dim fld As ADODB.field
   Dim strRow As String * 130
   Dim strField As String
   Dim intPos As Integer

   Set rsSue118 = New ADODB.Recordset

   With rsSue118
      .Fields.Append "Liquidacion", adChar, 5, adFldUpdatable
      .Fields.Append "Legajo", adChar, 5, adFldUpdatable
      .Fields.Append "Concepto", adChar, 3, adFldUpdatable
      .Fields.Append "Caracter", adChar, 1, adFldUpdatable
      .Fields.Append "Cantidad", adChar, 9, adFldUpdatable
      .Fields.Append "Tarifa", adChar, 9, adFldUpdatable
      .Fields.Append "Importe", adChar, 9, adFldUpdatable
      .Fields.Append "TipoConcepto", adChar, 1, adFldUpdatable
      .Fields.Append "EmprSucu", adChar, 2, adFldUpdatable
      .Fields.Append "Departamento", adChar, 2, adFldUpdatable
      .Fields.Append "LugarPago", adChar, 1, adFldUpdatable
      .Fields.Append "CtoCosto", adChar, 2, adFldUpdatable
      .Fields.Append "Filler", adChar, 1, adFldUpdatable
      .Fields.Append "Sindicato", adChar, 1, adFldUpdatable
      .Fields.Append "Status", adChar, 1, adFldUpdatable
      .Fields.Append "ObraSocial", adChar, 2, adFldUpdatable
      .Fields.Append "NroRecibo", adChar, 5, adFldUpdatable
      .Fields.Append "NroPlanilla", adChar, 5, adFldUpdatable
      .Fields.Append "Filler2", adChar, 64, adFldUpdatable
      ' Utilice el tipo de cursor Keyset para permitir la actualización
      ' de los registros.
      .CursorType = adOpenKeyset
      .LockType = adLockOptimistic
      .Open
   End With
   
   Open sCarpeta & "\Sue218A.Arc" For Random As #1 Len = 130
   
   CantReg = LOF(1) / 130

  ' Line Input #1, strRow

   For i = 1 To CantReg
    
    
    Get #1, i, strRow
        
        With rsSue118
         
            If Mid(strRow, 52, 1) <> "*" Then
                
                .AddNew
                
                For Each fld In .Fields
                    strField = strRow
                    fld.value = Mid(strField, 1, fld.DefinedSize)
                
                    ' Elimina el texto del campo de la fila de texto.
                    strRow = Right(strRow, Len(strRow) - fld.DefinedSize)
                Next
                .Update
                .MoveFirst
            
            End If
            
        End With
            
    Next
      
    Close #1

End Sub




Public Sub GuardarEnArchivo()

'    Dim fld As ADODB.field
'    Dim strRow As String
'    Dim strField As String
'    Dim iReg As Long
'
'    Open SGSUEL & "\sue103.arc" For Output As #1
'
'    With rsPersonal
'
'      .MoveFirst
'      iReg = 1
'      Do While Not .EOF
'
'         For Each fld In .Fields
'            ' Si un campo tiene un valor, agrega comillas.
'            If Len(fld.value) > 0 Then
'                strField = fld.value
'            Else
'                strField = Space(fld.DefinedSize)
'            End If
'            ' Agrega el valor del campo y un delimitador de tabulación
'            ' a la cadena de salida.
'            strRow = strRow & strField
'        Next
'        ' Imprime la cadena de salida.
''       Seek #1, iReg  ' Establece la posición.
'        Print #1, strRow
'
'        strRow = ""
'        .MoveNext
'        iReg = iReg + 1
'
'      Loop
'
'   End With
'
'   Close

End Sub



