VERSION 5.00
Begin VB.Form frmGetImageUrl 
   Caption         =   "Desafío QR vía OCX"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetImage 
      Caption         =   "GetImage"
      Height          =   855
      Left            =   2640
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.PictureBox picImage 
      Height          =   2415
      Left            =   240
      ScaleHeight     =   2355
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox txtUrl 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmGetImageUrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdGetImage_Click()
    If txtUrl.Text = "" Then
        MsgBox "Debe ingresar una dirección correspondiente a una imagen", vbInformation
    Else
        '' Cargo Imagen
        Getimage txtUrl.Text
    End If
    
End Sub


Private Sub Getimage(path As String)

    Dim bResult As Boolean
    Dim sDestPath As String
    
    On Error GoTo error_handler
 
     sDestPath = "C:\url_Image.jpg"
 
     bResult = DownloadFile(txtUrl.Text, sDestPath) ' Descargo imagen
 
    If bResult = True Then 'Si la descarga fue satisfactoria muestro en el picturebox
    With picImage
            .Picture = LoadPicture(sDestPath)
            .AutoSize = True
            .Width = Me.Width / 2
            .Height = Me.Height / 2
            .ScaleMode = 3
            .AutoRedraw = True
            .PaintPicture .Picture, 0, 0, .ScaleWidth, .ScaleHeight
    End With

    Kill sDestPath 'Borro el archivo descargado

    End If

Exit Sub
error_handler:
    If Err.Number = 53 Then
        Kill sDestPath
        MsgBox "La ruta no es válida. Verifique el archivo", vbCritical, "Error"
    Else
        Kill sDestPath
        MsgBox Err.Description, vbCritical
    End If

End Sub
