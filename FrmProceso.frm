VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmProceso 
   Caption         =   "Procesando información"
   ClientHeight    =   555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   555
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar pbar 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "FrmProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
