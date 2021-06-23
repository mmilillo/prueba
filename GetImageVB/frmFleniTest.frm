VERSION 5.00
Object = "{DFB67E2B-A2F6-418B-8152-1F01509E1801}#2.0#0"; "GetImageOCXProyect.ocx"
Begin VB.Form frmFleniTest 
   Caption         =   "Start Window"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin GetImagexOCXProyect.UserControl1 UserControl1 
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1508
   End
End
Attribute VB_Name = "frmFleniTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()
UserControl1_GotFocus
End Sub

