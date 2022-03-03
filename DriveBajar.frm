VERSION 5.00
Begin VB.Form DriveBajar 
   Caption         =   "Bajar Datos"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3570
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5895
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   5655
      End
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "DriveBajar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
RutaDir = Dir1.Path
If Dir(RutaDir & "\inicial.edu") = "" Then
    MsgBox "LA RUTA SELECCIONADA NO CONTIENE DATOS DE NOTAS PARA DESCARGAR, VERIFIQUE NUEVAMENTE.", 48, "ADVERTENCIA"
    Exit Sub
End If
RESP = MsgBox("DESEA BAJAR LA INFORMACIÓN DESDE " & RutaDir, vbYesNo + vbQuestion + vbDefaultButton1, "BAJAR INFORMACION")
If RESP = vbYes Then
    'I = 0
    'PASSW.Show 1
    'If I = 1 Then
        Unload Me
        BAJAR_DISCOPRO.Show
    'End If
End If
End Sub

Private Sub Dir1_Change()
Frame1.Caption = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub
