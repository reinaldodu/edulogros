VERSION 5.00
Begin VB.Form PASSW 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contraseña de acceso"
   ClientHeight    =   1815
   ClientLeft      =   3045
   ClientTop       =   2895
   ClientWidth     =   3255
   Icon            =   "PASSW.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.TextBox Text1 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "PASSW.frx":0442
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contraseña:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   855
      End
   End
End
Attribute VB_Name = "PASSW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim contra As clave
Dim PASS_CLAV As String

If Text1.Text = "" Then
    MsgBox "ESCRIBA SU PASSWORD", 48, "CONTRASEÑA"
    Text1.SetFocus
    Exit Sub
End If
CLA = 0
NAR = FreeFile
Open App.Path & "\clase.edu" For Random As #NAR Len = Len(contra)
While Not EOF(NAR)
    CLA = CLA + 1
    Get #NAR, CLA, contra
    
    PASS_CLAV = ""

    For I = 1 To Len(Trim(contra.PASSW)) Step 3
        PASS_CLAV = PASS_CLAV & Chr(Val(Mid(contra.PASSW, I, 3)))
    Next
    
    If (UCase(Text1) = PASS_CLAV) Then
        Close #NAR
        Unload Me
        I = 1
        Exit Sub
    End If
Wend
Close #NAR
MsgBox "CONTRASEÑA INCORRECTA", 64, "CONTRASEÑA"
Text1.Text = ""
Text1.SetFocus
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Escriba su password de entrada."
End Sub

Private Sub Form_Load()
Text1.MaxLength = 15
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub
