VERSION 5.00
Begin VB.Form INS_NOTA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insertar"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2535
   Icon            =   "INS_NOTA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   2535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text2 
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
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No.Carnet:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   765
      End
   End
End
Attribute VB_Name = "INS_NOTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim alugru As grupoalu
Dim alumno As maestroalum
Screen.MousePointer = 11
If Text2.Text = "" Then
    MsgBox "ESCRIBA UN NUMERO DE CARNET", 64, "ADVERTENCIA"
    Text2.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If
If Val(Text2.Text) > 32000 Then
    MsgBox "No. DE CARNET INVALIDO", 64, "ADVERTENCIA"
    Text2.Text = ""
    Text2.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If
NAR = FreeFile
Open "c:\windows\datos\cont.edu" For Input As #NAR
Input #NAR, I
Close #NAR
h = Val(Text2.Text)
If ((h > I - 1) Or (h < 1)) Then
    MsgBox "REGISTRO NO EXISTE", 32
    Text2.Text = ""
    Text2.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If
GRABAR_OBSER.MATI12.Col = 14
For J = 1 To Val(GRABAR_OBSER.Text7.Text)
GRABAR_OBSER.MATI12.Row = J
If Val(Right(GRABAR_OBSER.MATI12.Text, 5)) = Val(Text2.Text) Then
    MsgBox "ALUMNO YA EXISTE", 32
    Text2.Text = ""
    Text2.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If
Next J
Open "c:\windows\datos\prinalu.edu" For Random As #NAR Len = Len(alumno)
Get #NAR, h, alumno
Close #NAR
If RTrim(alumno.n_carnet) = "" Then
    MsgBox "ALUMNO ESTA RETIRADO", 32
    Text2.Text = ""
    Text2.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If
t = 0
s = 0
Open "c:\windows\datos\" & RTrim(GRABAR_OBSER.Combo2.Text) & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
t = t + 1
Get #NAR, t, alugru
If Val(alugru.num_carnet) = h Then
s = s + 1
End If
Wend
Close #NAR
If s = 0 Then
    MsgBox "ALUMNO NO PERTENECE A ESTE GRUPO", 16, "ADVERTENCIA"
    Text2.Text = ""
    Text2.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If
GRABAR_OBSER.Text7.Text = GRABAR_OBSER.Text7.Text + 1
GRABAR_OBSER.MATI12.Rows = GRABAR_OBSER.MATI12.Rows + 1
GRABAR_OBSER.MATI12.Row = Val(GRABAR_OBSER.Text7.Text)
GRABAR_OBSER.MATI12.Col = 13
GRABAR_OBSER.MATI12.Text = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
GRABAR_OBSER.MATI12.Col = 14
GRABAR_OBSER.MATI12.Text = alumno.n_carnet
GRABAR_OBSER.MATI12.Col = 13
GRABAR_OBSER.MATI12.Sort = 5
GRABAR_OBSER.MATI12.Col = 0
For TT = 1 To Val(GRABAR_OBSER.Text7.Text)
GRABAR_OBSER.MATI12.Row = TT
GRABAR_OBSER.MATI12.Text = TT
Next TT
Text2.Text = ""
Text2.SetFocus
VALI4 = False
Screen.MousePointer = 0
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Inserta un alumno en la lista, escribiendo el carnet que le corresponde."
End Sub

Private Sub Form_Load()
Text2.MaxLength = 5
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
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
