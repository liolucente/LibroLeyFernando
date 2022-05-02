VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sgTratamiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Inicio As Boolean

Public Titulo As String

Public ClaveNombre As String
Public ClaveCampo As String

Public colCriterios As sgTRCriterios

Public GrillaDatos As ADODB.Recordset
Public GrillaColumnas As sgGrillaColColumnas

Public colBotones As sgColBotones

