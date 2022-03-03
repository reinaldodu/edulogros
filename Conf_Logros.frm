VERSION 5.00
Begin VB.Form Conf_Logros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Porcentaje de logros"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2670
   Icon            =   "Conf_Logros.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   2670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.OptionButton Option2 
         Caption         =   "Porcentaje manual"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Porcentaje automático"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Conf_Logros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ConfLgr As Byte
Private Sub Command1_Click()
'Logros de igual porcentaje = 0
'Logros de diferente porcentaje = 1
    If Option1.Value = True Then
        ConfLgr = 0
    Else
        ConfLgr = 1
    End If
    Open Ruta & "conf_logro.edu" For Output As #NAR
    Print #NAR, ConfLgr
    Close #NAR
    Unload Me
End Sub

Private Sub Form_Load()
    If Dir(Ruta & "conf_logro.edu") = "" Then
        Option1.Value = True
        Exit Sub
    End If
    Open Ruta & "conf_logro.edu" For Input As #NAR
    Input #NAR, ConfLgr
    Close #NAR
    If ConfLgr = 1 Then
        Option2.Value = True
    Else
        Option1.Value = True
    End If
End Sub
