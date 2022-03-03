VERSION 5.00
Begin VB.Form CAMBIO_DIRECT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar director de grupo"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4335
   Icon            =   "CAMBIO_DIRECT.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cambiar"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   320
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cambiar por:"
      ForeColor       =   &H00C00000&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.TextBox Text2 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   2895
      End
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
         Height          =   320
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Profesor No."
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   885
      End
   End
End
Attribute VB_Name = "CAMBIO_DIRECT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim profe As maestropro
If Text1.Text = "" Then
    MsgBox "ESCRIBA EL NUMERO DEL DIRECTOR", 48, "CAMBIAR"
    Text1.SetFocus
    Exit Sub
End If
NAR = FreeFile
Open Ruta & "contpro.edu" For Input As #NAR
Input #NAR, r
Close #NAR
dire = Text1.Text
If ((dire > r - 1) Or (dire < 1)) Then
    MsgBox "PROFESOR NO EXISTE", 64, "ADVERTENCIA"
    Text2.Text = ""
    Text1.SetFocus
    Exit Sub
End If
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Get #NAR, dire, profe
Close #NAR
Text2.Text = RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
Command2.SetFocus
End Sub

Private Sub Command2_Click()
'Dim icur As inforcur
'Dim profe As maestropro
If Text2.Text = "" Then
MsgBox "ESCRIBA EL NUMERO DEL PROFESOR", 64, "ADVERTENCIA"
Text1.SetFocus
Exit Sub
End If
RESP = MsgBox("DESEA CAMBIAR EL DIRECTOR DE GRUPO?", vbYesNo + vbQuestion + vbDefaultButton1, "CAMBIAR DIRECTOR")
If RESP = vbYes Then
NAR = FreeFile
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
If RTrim(icur.nom) = RTrim(CONS_GRUP.Combo1.Text) Then
icur.director = Text1.Text
End If
NAR = FreeFile
Open Ruta & "infcur2.edu" For Append As #NAR
Write #NAR, icur.nom, icur.jornada, icur.grado, icur.director
Close #NAR
NAR = NAR - 1
Wend
Close #NAR
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Get #NAR, Text1.Text, profe
Close #NAR
CONS_GRUP.Label7.Caption = "Directora: " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
Kill Ruta & "INFCUR.EDU"
Name Ruta & "INFCUR2.EDU" As Ruta & "INFCUR.EDU"
End If
Unload Me
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Cambia el director de grupo, escribiendo el código del profesor y dando click en cambiar."
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command1_Click
End If
C$ = Chr(KeyAscii)
If KeyAscii = 8 Then
    Exit Sub
End If
If C$ < "0" Or C$ > "9" Then
    KeyAscii = 0
    Beep
End If
End Sub

Private Sub Form_Load()
Text1.MaxLength = 3
End Sub
