VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sgDSueldosAdic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public legajo As String
Public Nombre As String
Public Domicilio As String
Public CodPostal As String
Public Nacionalidad As String
Public Sexo As String
Public TipoDoc As String
Public NumeroDoc As String
Public EstCivil As String
Public Localidad As String
Public Provincia As String
Public FechaNac As String
Public FechaIngPais As String
Public FechaIngreso As String
Public FECHAALTA As String
Public FechaEgreso As String
Public Objetivo As String
Public ModCont As String
Public RebajaProm As String
Public ActReduc As String
Public CodSindicato As String
Public NroSindical As String
Public PorcSindicato As String
Public PlanOsoc As String
Public NroOsoc As String
Public PorcOsoc As String
Public Empresa As String
Public Departamento As String
Public Seccion As String
Public CtoCosto As String
Public LugarPago As String
Public TipoPers As String
Public tarea As String
Public Turno As String
Public FormaPago As String
Public Banco As String
Public JornalHrEsp As String
Public Bloqueo As String
Public MotivoCese As String
Public CargasLey18610 As String
Public TipoAntig As String
Public Titulo As String
Public CaracterServ As String
Public Coeficiente1 As String
Public Coeficiente2 As String
Public Coeficiente3 As String
Public Conyuge As String
Public Hijos As String
Public HijosIncap As String
Public Indicador3 As String
Public AdherSIJP As String
Public Libre1 As String
Public IndFranjaSalFam As String
Public Prenatal As String
Public LicMaternidad As String
Public EscHijosIncap As String
Public Libre2 As String
Public Sueldo As String
Public AcumSAC As String
Public AcumU44 As String
Public Cuil As String
Public Indicador1 As String
Public Indicador2 As String
Public ActividadSIJP As String
Public CtoContable As String
Public OSocial As String
Public ImpAux1 As String
Public ImpAux2 As String
Public ImpAux3 As String
Public DiasNoTrabSem As String
Public Afjp As String
Public CpoAux1 As String
Public CpoAux2 As String
Public CpoAux3 As String
Public AniosAntCalc As String
Public FechaAltaPuesto As String
Public CtaBancaria As String
Public SituacionSIJP As String
Public CondicionSIJP As String
Public ContratacionSIJP As String
Public COMENTARIO As String
Public Extension1 As String
Public Extension2 As String
Public Extension3 As String

Public CBUBloque1 As String
Public IndComplBonos As String
Public Indef1 As String
Public Indef2 As String
Public Indef3 As String
Public Indef4 As String
Public Indef5 As String
Public Indef6 As String
Public Indef7 As String
Public Inex1 As String
Public Inex2 As String
Public UltLiquid As String
Public ImpAdicLiq As String
Public UltNetoPag As String
Public AcumMes As String
Public AcumLiq As String
Public AcumCorriente As String
Public FechaUltPago As String
Public ImpPlanOSocial As String
Public CtaBonos As String


Public Function GrabarModificacion() As Boolean
    
    Dim mRegistro As String * 490   'antes 492
    Dim mIndice As String * 12
    Dim mReg As String
    
    'prueba de relleno del indice
    Dim i As Long
    
    
    'relleno de sue101am.arc
    Dim mRegistro3 As String * 247

    mReg = Me.legajo                                        '1-001-005
    mReg = mReg & Me.Nombre                                 '1-006-024
    mReg = mReg & Me.Domicilio                              '1-030-024
    mReg = mReg & Me.CodPostal                              '1-054-004
            
    mReg = mReg & Me.ModCont                                '1-058-008
'    mReg = mReg & CompletarDer(Me.RebajaProm, " ", 3)       '1-058-008
                   
    mReg = mReg & Me.Nacionalidad                           '1-066-001
    mReg = mReg & Me.Sexo                                   '1-067-001
    mReg = mReg & Me.ImpPlanOSocial                         '1-068-006
    mReg = mReg & Me.CBUBloque1                             '1-074-006
    mReg = mReg & Me.TipoDoc                                '1-080-002
    mReg = mReg & Me.NumeroDoc                              '1-082-008
    mReg = mReg & Me.EstCivil                               '1-090-001
    mReg = mReg & Me.IndComplBonos                          '1-091-001
    mReg = mReg & Me.Localidad                              '1-092-010
    mReg = mReg & Me.ActReduc                               '1-102-001
    mReg = mReg & Me.CodSindicato                           '1-103-001
    mReg = mReg & Me.NroSindical                            '1-104-010
    mReg = mReg & Me.PorcSindicato                          '1-114-005
    mReg = mReg & Me.Indef1                                 '1-119-001
    mReg = mReg & Me.PlanOsoc                               '1-120-001
    mReg = mReg & Me.NroOsoc                                '1-121-010
    mReg = mReg & Me.PorcOsoc                               '1-131-005
    mReg = mReg & Me.Indef2                                 '1-136-006
    mReg = mReg & Me.Empresa                                '1-142-002
    mReg = mReg & Me.Departamento                           '1-144-002
    mReg = mReg & Me.Seccion                                '1-146-002
    mReg = mReg & Me.CtoCosto                               '1-148-002
    mReg = mReg & Me.LugarPago                              '1-150-001
    mReg = mReg & Me.TipoPers                               '1-151-001
    mReg = mReg & Me.Indef3                                 '1-152-006
    mReg = mReg & Me.tarea                                  '1-158-002
    mReg = mReg & Me.Turno                                  '1-160-001
    mReg = mReg & Me.FormaPago                              '1-161-001
    mReg = mReg & Me.Banco                                  '1-162-002
    mReg = mReg & Me.JornalHrEsp                            '1-164-010
    mReg = mReg & Me.Bloqueo                                '1-174-001
    mReg = mReg & Me.MotivoCese                             '1-175-002
    mReg = mReg & Me.Indef4                                 '1-177-006
    mReg = mReg & Me.CargasLey18610                         '1-183-001
    mReg = mReg & Me.TipoAntig                              '1-184-001
    mReg = mReg & Me.Titulo                                 '1-185-001
    mReg = mReg & Me.CaracterServ                           '1-186-001
    mReg = mReg & Me.Provincia                              '1-187-002
    
    mReg = mReg & Me.Coeficiente1                           '1-189-005
    mReg = mReg & Me.Coeficiente2                           '1-194-005
    mReg = mReg & Me.Coeficiente3                           '1-199-005
    
    mReg = mReg & Me.Conyuge                                '1-204-001
    mReg = mReg & Me.Hijos                                  '1-205-001
    mReg = mReg & Me.HijosIncap                             '1-206-001
    mReg = mReg & Me.Indef5                                 '1-207-001
    mReg = mReg & Me.Indicador3                             '1-208-001
    mReg = mReg & Me.Indef6                                 '1-209-001
    mReg = mReg & Me.AdherSIJP                              '1-210-001
    mReg = mReg & Me.Indef7                                 '1-211-001
    mReg = mReg & Me.Libre1                                 '1-212-001
    mReg = mReg & Me.IndFranjaSalFam                        '1-213-001
    mReg = mReg & Me.Prenatal                               '1-214-001
    mReg = mReg & Me.LicMaternidad                          '1-215-001
    mReg = mReg & Me.EscHijosIncap                          '1-216-001
    mReg = mReg & Me.Libre2                                 '1-217-001
    mReg = mReg & Me.Sueldo                                 '1-218-008
    mReg = mReg & Me.AcumSAC                                '1-226-008
    mReg = mReg & Me.AcumU44                                '1-234-009
    mReg = mReg & Me.Cuil                                   '2-001-013
    mReg = mReg & Me.Inex1                                  '2-014-003
    mReg = mReg & Me.Indicador1                             '2-017-001
    mReg = mReg & Me.Indicador2                             '2-018-001
    mReg = mReg & Me.ActividadSIJP                          '2-019-002
    mReg = mReg & Me.Inex2                                  '2-021-002
    mReg = mReg & Me.CtoContable                            '2-023-002
    mReg = mReg & Me.OSocial                                '2-025-002
    mReg = mReg & Me.UltLiquid                              '2-027-005
    mReg = mReg & Me.ImpAdicLiq                             '2-032-006
    mReg = mReg & Me.UltNetoPag                             '2-038-008
    mReg = mReg & Me.AcumMes                                '2-046-008
    mReg = mReg & Me.AcumLiq                                '2-054-008
    mReg = mReg & Me.ImpAux1                                '2-062-008
    mReg = mReg & Me.AcumCorriente                          '2-070-008
    mReg = mReg & Me.ImpAux2                                '2-078-008
    mReg = mReg & Me.ImpAux3                                '2-086-008
    mReg = mReg & Me.DiasNoTrabSem                          '2-094-003
    mReg = mReg & Me.Afjp                                   '2-097-002
    
    mReg = mReg & Me.CpoAux1                                '2-099-005
    mReg = mReg & Me.CpoAux2                                '2-104-005
    mReg = mReg & Me.CpoAux3                                '2-109-005
    
    mReg = mReg & Me.AniosAntCalc                           '2-114-005
    mReg = mReg & Me.FechaNac                               '2-119-008
    mReg = mReg & Me.FechaIngPais                           '2-127-008
    mReg = mReg & Me.FechaIngreso                           '2-135-008
    mReg = mReg & Me.FechaAltaPuesto                        '2-143-008
    mReg = mReg & Me.CtaBancaria                            '2-151-016
    mReg = mReg & Me.FechaEgreso                            '2-167-008
    mReg = mReg & Me.SituacionSIJP                          '2-175-002
    mReg = mReg & Me.CondicionSIJP                          '2-177-002
    mReg = mReg & Me.ContratacionSIJP                       '2-179-003
    mReg = mReg & Me.FechaUltPago                           '2-182-008
'CAMBIO TEMPORAL HASTA SINCRONIZAR OBJETIVOS
'    mReg = mReg & Me.Objetivo
    mReg = mReg & Me.CtaBonos                               '2-190-016
    mReg = mReg & Me.COMENTARIO                             '2-206-016
    mReg = mReg & Me.Extension1                             '2-222-009
    mReg = mReg & Me.Extension2                             '2-231-009
    mReg = mReg & Me.Extension3                             '2-240-009
   
    mRegistro = mReg
'MsgBox "Abriendo sue101.ind"
'    Open SGSUEL & "\sue101.ind" For Random Access Read As #1 Len = 12
'MsgBox "sue101.ind OK"
    Seek #1, Me.legajo
    Get #1, , mIndice
    
    If Val(Mid(mIndice, 6, 5)) > 0 Then
'MsgBox "Abriendo sue101.arc"
'        Open SGSUEL & "\sue101.arc" For Random Access Read Write As #2 Len = 492    'antes 494
'MsgBox "sue101.arc OK"
        Seek #2, Val(Mid(mIndice, 6, 5))
        Put #2, , mRegistro
        Close #2
        
    End If
    
    Close #1
    
    'grabacion en el SUE101AM.ARC
'    Open SGSUEL & "\sue101AM.ARC" For Random Access Read Write As #3 Len = 249
'    Seek #3, Me.legajo
'    mRegistro3 = CompletarDer("", " ", 1) & CompletarIzq(Me.legajo, "0", 5) & CompletarIzq(frmPeDatosPersonal.txtCBU.Text, " ", 25)
'    Put #3, , mRegistro3
'    Close #3

    GrabarModificacion = True
    
End Function

'sueldos_grabacion_sue101
'nueva funcion que hace la grabacion de la misma manera que la modificacion
Public Function GrabarAlta() As Boolean
    
    
    Dim mRegistro As String * 490   'antes 492
    Dim mRegistro2 As String * 492
    Dim mIndice As String * 10
    Dim mReg As String
    Dim i As Long
    Dim Indice As Long
    'Para el relleno del sue101am.arc
    Dim mRegistro3 As String * 247
    Dim mCbu As String * 31
    

    mReg = Me.legajo
    mReg = mReg & Me.Nombre
    mReg = mReg & Me.Domicilio
    mReg = mReg & Me.CodPostal
    
    mReg = mReg & Me.ModCont
    mReg = mReg & Me.RebajaProm
    
    mReg = mReg & Me.Nacionalidad
    mReg = mReg & Me.Sexo
    
    mReg = mReg & Me.ImpPlanOSocial
    mReg = mReg & Me.CBUBloque1
    
    mReg = mReg & Me.TipoDoc
    mReg = mReg & Me.NumeroDoc
    mReg = mReg & Me.EstCivil
    
    mReg = mReg & Me.IndComplBonos
    
    mReg = mReg & Me.Localidad
    
    mReg = mReg & Me.ActReduc
    mReg = mReg & Me.CodSindicato
    mReg = mReg & Me.NroSindical
    mReg = mReg & Me.PorcSindicato
    
    mReg = mReg & Me.Indef1
    mReg = mReg & Me.PlanOsoc
    mReg = mReg & Me.NroOsoc
    mReg = mReg & Me.PorcOsoc
    mReg = mReg & Me.Indef2
    mReg = mReg & Me.Empresa
    mReg = mReg & Me.Departamento
    mReg = mReg & Me.Seccion
    mReg = mReg & Me.CtoCosto
    mReg = mReg & Me.LugarPago
    mReg = mReg & Me.TipoPers
    mReg = mReg & Me.Indef3
    mReg = mReg & Me.tarea
    mReg = mReg & Me.Turno
    mReg = mReg & Me.FormaPago
    mReg = mReg & Me.Banco
    mReg = mReg & Me.JornalHrEsp
    mReg = mReg & Me.Bloqueo
    mReg = mReg & Me.MotivoCese
    mReg = mReg & Me.Indef4
    mReg = mReg & Me.CargasLey18610
    mReg = mReg & Me.TipoAntig
    mReg = mReg & Me.Titulo
    mReg = mReg & Me.CaracterServ
    
    mReg = mReg & Me.Provincia
    mReg = mReg & Me.Coeficiente1
    mReg = mReg & Me.Coeficiente2
    mReg = mReg & Me.Coeficiente3
    mReg = mReg & Me.Conyuge
    mReg = mReg & Me.Hijos
    mReg = mReg & Me.HijosIncap
    mReg = mReg & Me.Indef5
    mReg = mReg & Me.Indicador3
    mReg = mReg & Me.Indef6
    mReg = mReg & Me.AdherSIJP
    mReg = mReg & Me.Indef7
    mReg = mReg & Me.Libre1
    mReg = mReg & Me.IndFranjaSalFam
    mReg = mReg & Me.Prenatal
    mReg = mReg & Me.LicMaternidad
    mReg = mReg & Me.EscHijosIncap
    mReg = mReg & Me.Libre2
    mReg = mReg & Me.Sueldo
    mReg = mReg & Me.AcumSAC
    mReg = mReg & Me.AcumU44
    mReg = mReg & Me.Cuil
    
    mReg = mReg & Me.Inex1
    mReg = mReg & Me.Indicador1
    mReg = mReg & Me.Indicador2
    mReg = mReg & Me.ActividadSIJP
    mReg = mReg & Me.Inex2
    mReg = mReg & Me.CtoContable
    mReg = mReg & Me.OSocial
    mReg = mReg & Me.UltLiquid
    mReg = mReg & Me.ImpAdicLiq
    
    mReg = mReg & Me.UltNetoPag
    mReg = mReg & Me.AcumMes
    mReg = mReg & Me.AcumLiq
    mReg = mReg & Me.ImpAux1
    mReg = mReg & Me.AcumCorriente
    mReg = mReg & Me.ImpAux2
    mReg = mReg & Me.ImpAux3
    mReg = mReg & Me.DiasNoTrabSem
    mReg = mReg & Me.Afjp
    mReg = mReg & Me.CpoAux1
    mReg = mReg & Me.CpoAux2
    mReg = mReg & Me.CpoAux3
    mReg = mReg & Me.AniosAntCalc
    mReg = mReg & Me.FechaNac
    mReg = mReg & Me.FechaIngPais
    mReg = mReg & Me.FechaIngreso
    mReg = mReg & Me.FechaAltaPuesto
    mReg = mReg & Me.CtaBancaria
    mReg = mReg & Me.FechaEgreso
    mReg = mReg & Me.SituacionSIJP
    mReg = mReg & Me.CondicionSIJP
    mReg = mReg & Me.ContratacionSIJP
    mReg = mReg & Me.FechaUltPago
' CAMBIO TEMPORAL HASTA SINCRONIZAR OBJETIVOS
'    mReg = mReg & Me.Objetivo

    mReg = mReg & Me.CtaBonos
    mReg = mReg & Me.COMENTARIO
    mReg = mReg & Me.Extension1
    mReg = mReg & Me.Extension2
    mReg = mReg & Me.Extension3
   
    mRegistro = mReg    'len(mReg)=348 DEBE SER = 490

'relleno el archivo indice hasta 99999
'Open SGSUEL & "\sue101.ind" For Append As #1 Len = 12
'For i = 2001 To 10000
''    indice_Leg.legajo = CompletarIzq(i, "0", 5)
''    indice_Leg.posicion = CompletarIzq("", " ", 5)
'mIndice = CompletarIzq(i, "0", 5) & CompletarIzq("", " ", 5)
'Print #1, mIndice
'
'Next i
'Close #1
'aca va la grabacion en el archivo de indice

'RELLENO EL SUE101AM.ARC HASTA 99999
'Open SGSUEL & "\sue101am.arc" For Output As #1 Len = 248
'For i = 1 To 10000
''''    indice_Leg.legajo = CompletarIzq(i, "0", 5)
''''    indice_Leg.posicion = CompletarIzq("", " ", 5)
'mRegistro3 = CompletarDer("", " ", 1) & CompletarIzq(i, "0", 5) & CompletarIzq("", " ", 25)
'Print #1, mRegistro3
'Next i
'Close #1
'ACA VA LA GRABACION EN EL ARCHIVO DE LA CBU





'primero grabo en el sue101.arc
'Open SGSUEL & "\sue101.arc" For Append As #2 Len = 491  'antes 493
Print #2, mRegistro
Close #2
        
'¿como sé donde se grabo? numero de linea
'uso la funcion Contar_lineas que devuelve el num. de lineas
'Indice = Contar_Lineas(SGSUEL & "\sue101.arc")

'hago una prueba de posicionamiento en el sue101.arc
'Open SGSUEL & "\sue101.arc" For Random Access Read Write As #2 Len = 492    'antes 494
Seek #2, Indice
Get #2, , mRegistro2
Close #2


'Open SGSUEL & "\sue101.ind" For Random Access Read Write As #1 Len = 12
'Seek #1, Me.Legajo + 1  'me posiciono en el legajo
'mIndice = Me.legajo & CompletarIzq(Indice, "0", 5)
Put #1, Me.legajo, mIndice    'y actualizo el indice: ESTA GRABACION ESTA PERFECTA



Close #1
    
    
'grabacion en el SUE101AM.ARC

On Error GoTo sue101am
'Open SGSUEL & "\sue101AM.ARC" For Random Access Read Write As #4 Len = 249
Seek #4, Val(Me.legajo)
'mRegistro3 = CompletarDer("", " ", 1) & CompletarIzq(Me.legajo, "0", 5) & CompletarIzq(frmPeDatosPersonal.txtCBU.Text, " ", 25)
Put #4, , mRegistro3
Close #4


GrabarAlta = True
    
Exit Function
    
sue101am:
Close #4
'Open SGSUEL & "\sue101AM.ARC" For Random Access Read Write As #5 Len = 249
Seek #5, Val(Me.legajo)
'mRegistro3 = CompletarDer("", " ", 1) & CompletarIzq(Me.legajo, "0", 5) & CompletarIzq(frmPeDatosPersonal.txtCBU.Text, " ", 25)
Put #5, , mRegistro3
Close #5
    
End Function

Public Function Contar_Lineas(ByVal strTextFile As String) As Long

    Dim F As Integer
    Dim Lineas As Long
    Dim str_Linea As String

    ' Número de archivo libre
    F = FreeFile

    ' Abre el archivo de texto
    Open strTextFile For Input As #F

    'Recorre todo el archivo de texto _
     linea por linea hasta el final
    Do
        'Lee una línea
        Line Input #F, str_Linea
        
        ' Incrementa la cantidad de lineas leidas
        Lineas = Lineas + 1
    
    ' Leerá hasta que llegue al fin de archivo
    Loop While Not EOF(F)

    ' Cierra el archivo de texto abierto
    Close #F

    ' Retorna a la función el número de lineas del fichero
    Contar_Lineas = Lineas

End Function


