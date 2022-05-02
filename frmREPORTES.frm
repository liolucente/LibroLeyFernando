VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmREPORTES 
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9510
   Icon            =   "frmREPORTES.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin CRVIEWERLibCtl.CRViewer crv 
      Height          =   3165
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8880
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Menu mnuExportaciones 
      Caption         =   "Exportaciones"
      Visible         =   0   'False
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAXLS 
         Caption         =   "Exportar a Excel"
      End
      Begin VB.Menu mnuAHTML 
         Caption         =   "Exportar a HTML"
      End
      Begin VB.Menu mnuARTF 
         Caption         =   "Exportar a texto con formato"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "frmREPORTES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bReloj  As Boolean
Dim mImpresion  As frmConfigReportes

Public Property Set FormularioDeImpresion(ByVal F As frmConfigReportes)
Set mImpresion = F
End Property

Private Sub crv_DownloadFinished(ByVal loadingType As CRVIEWERLibCtl.CRLoadingType)
On Error Resume Next
If loadingType = crLoadingNothing Then
    Me.crv.Zoom 110
End If
Screen.MousePointer = vbNormal
End Sub

Private Sub crv_DownloadStarted(ByVal loadingType As CRVIEWERLibCtl.CRLoadingType)
Screen.MousePointer = vbHourglass
End Sub

Private Sub crv_PrintButtonClicked(UseDefault As Boolean)
UseDefault = False

mImpresion.Show vbModal, Me
End Sub

Private Sub Form_Activate()
'Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
crv.Top = 0
crv.Left = 0
crv.Zoom 80
Screen.MousePointer = vbHourglass
End Sub

Private Sub Form_Resize()
crv.Height = Me.ScaleHeight
crv.Width = Me.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide

On Error Resume Next
Screen.MousePointer = vbNormal
Set mImpresion.Reporte = Nothing
Unload mImpresion
Set mImpresion = Nothing
Screen.MousePointer = vbNormal
End Sub
