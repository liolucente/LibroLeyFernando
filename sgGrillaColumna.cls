VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sgGrillaColumna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Titulo As String
Public Campo As String
Public Ancho As Integer

Public TipoDato As Integer
    'sgtEmpty       0 Data type is undefined.
    'sgtShort       2 16 bit integer.
    'sgtLong        3 32 bit integer.
    'sgtFloat       4 Single precision number.
    'sgtDouble      5 Double precision number.
    'sgtCurrency    6 Currency.
    'sgtDateTime    7 Date-time value.
    'sgtString      8 Character string.
    'sgtBoolean     11 Boolean.
    'sgtDecimal     14 Decimal.
    'sgtByte        17 8 bit integer.
    'sgtBinary      65 Binary data.

Public Alineacion As Integer
    'sgAlignGeneral         0 Alignment is based on the underlying data type.
    'sgAlignLeftTop         1 Align to the left-top position.
    'sgAlignCenterTop       2 Align to the center-top position.
    'sgAlignRightTop        3 Align to the right-top position.
    'sgAlignLeftCenter      4 Align to the left-center position.
    'sgAlignCenterCenter    5 Align to the center-center position.
    'sgAlignRightCenter     6 Align to the right-center position.
    'sgAlignLeftBottom      7 Align to the left-bottom position.
    'sgAlignCenterBottom    8 Align to the center-bottom position.
    'sgAlignRightBottom     9 Align to the right-bottom position.

Public TipoComparacion As Integer
    'sgSortTypeDefault      0 Convert values to the column's data type before comparing.
    'sgSortTypeCustom       1 Fire the CompareCell event to compare two values.
    'sgSortTypeNumber       2 Convert to double-precision values before comparing.
    'sgSortTypeString       3 Convert to string values before comparing.
    'sgSortTypeNoCaseString 4 Convert to string values before comparing. Ignore case.
    'sgSortTypeDateTime     5 Convert to date-time values before comparing.
 
Public Oculta As Boolean

Public key As String




