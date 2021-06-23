VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   1440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   ScaleHeight     =   1440
   ScaleWidth      =   5385
   Begin VB.CommandButton cmdGetImageOCX 
      Caption         =   "Get Image From Url"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Event Click()

Private Sub cmdGetImageOCX_Click()
    
    RaiseEvent Click
    
    Load frmGetImageUrl
    
    frmGetImageUrl.Show vbModal, Me 'UserControl.Parent
    Set frmGetImageUrl = Nothing
    
End Sub

