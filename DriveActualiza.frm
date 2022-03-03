VERSION 5.00
Begin VB.Form DriveActualiza 
   Caption         =   "Actualizar Datos"
   ClientHeight    =   3540
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   6015
      Begin VB.DirListBox Dir1 
         Height          =   1665
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   2880
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "DriveActualiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
RutaDir = Dir1.Path
RESP = MsgBox("DESEA ACTUALIZAR LA INFORMACIÓN HACIA " & RutaDir, vbYesNo + vbQuestion + vbDefaultButton1, "ACTUALIZAR INFORMACION")
If RESP = vbYes Then
    I = 0
    PASSW.Show 1
    If I = 1 Then
        Unload Me
        ACTUDISKPRO.Show
    End If
End If

End Sub

Private Sub Dir1_Change()
Frame1.Caption = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub
