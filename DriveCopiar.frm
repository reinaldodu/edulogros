VERSION 5.00
Begin VB.Form DriveCopiar 
   Caption         =   "Copiar Datos"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5175
      Begin VB.DirListBox Dir1 
         Height          =   990
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   4815
      End
   End
End
Attribute VB_Name = "DriveCopiar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
RutaDir = Dir1.Path
If Dir(RutaDir & "\datos", vbDirectory) <> "" Then
    MsgBox "Ya existen datos creados en este directorio, seleccione otro directorio para hacer la copia", 48, "Copiar datos"
    Exit Sub
End If
If Dir("c:\datos", vbDirectory) <> "" Then
    MsgBox "Necesita borrar o mover primero el directorio C:\DATOS\ para realizar la copia", 48, "Copiar datos"
    Exit Sub
End If
RESP = MsgBox("DESEA COPIAR LA INFORMACIÓN EN " & RutaDir, vbYesNo + vbQuestion + vbDefaultButton1, "COPIAR INFORMACION")
If RESP = vbYes Then
    'I = 0
    'PASSW.Show 1
    'If I = 1 Then
        Unload Me
        DISCO_PROFES.Show
    'End If
End If
End Sub

Private Sub Dir1_Change()
Frame1.Caption = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub
