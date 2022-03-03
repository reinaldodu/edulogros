VERSION 5.00
Begin VB.Form conf_encabeza2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encabezado reporte mitad de periodo"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7335
   Icon            =   "Conf_Encabeza2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
   End
End
Attribute VB_Name = "conf_encabeza2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ConfTexto = Text1
    Open Ruta & "conf_encabeza2.txt" For Output As #NAR
    Print #NAR, ConfTexto
    Close #NAR
    Unload Me
End Sub

Private Sub Form_Load()
    If Dir(Ruta & "conf_encabeza2.txt") = "" Then
        Exit Sub
    End If
    Open Ruta & "conf_encabeza2.txt" For Input As #NAR
    Input #NAR, ConfTexto
    Close #NAR
    Text1 = ConfTexto
End Sub
