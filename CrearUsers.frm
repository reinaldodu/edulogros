VERSION 5.00
Begin VB.Form CrearUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas de usuario en red"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "CrearUsers.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4095
      Begin VB.TextBox Text2 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "PASSWORD:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "PROFESOR No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1485
      End
   End
End
Attribute VB_Name = "CrearUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
    MsgBox "ESCRIBA EL NUMERO DE PROFESOR", 64, "ADVERTENCIA"
    Text1.SetFocus
    Exit Sub
End If
If Text2.Text = "" Then
    MsgBox "ESCRIBA EL PASSWORD DEL PROFESOR", 64, "ADVERTENCIA"
    Text2.SetFocus
    Exit Sub
End If
NAR = FreeFile
Open Ruta & "contpro.edu" For Input As #NAR
Input #NAR, r
Close #NAR
w = Val(Text1.Text)
If ((w > r - 1) Or (w < 1)) Then
    MsgBox "PROFESOR NO EXISTE", 32, "Creación de usuarios"
    Text1.SetFocus
    Exit Sub
End If
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Get #NAR, w, profe
Close #NAR
If ((RTrim(profe.nombres) = "") And (RTrim(profe.especiali) = "")) Then
    MsgBox "REGISTRO NO EXISTE", 16, "Creación de usuarios"
    Text1.SetFocus
    Exit Sub
End If
RESP = MsgBox("Desea crear la cuenta de usuario para " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos) & "?", vbYesNo + vbQuestion + vbDefaultButton1, "Creación de usuarios")
If RESP = vbYes Then
    On Error Resume Next
    Err.Clear
    Screen.MousePointer = 11
    que = 0
    Open Ruta & "CLAPRO.EDU" For Random As #NAR Len = Len(CLAV)
    While Not EOF(NAR)
        que = que + 1
        Get #NAR, que, CLAV
        If CLAV.NUMERO = w Then
            GoTo SAIRS
        End If
    Wend
SAIRS:
    CLAV.NUMERO = w
    CLAV.PASSWW = RTrim(Text2.Text)
    Put #NAR, que, CLAV
    If Err.Number <> 0 Then
        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
        Screen.MousePointer = 0
        Close #NAR
        Exit Sub
    End If
    Close #NAR
    MsgBox "La cuenta fue creada satisfactoriamente", 64, "Creación de usuarios"
End If
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
'I = 0
'PASSW.Show 1
'If I = 1 Then
'    PASSWS_PROFES.MATI14.ColWidth(0) = 4000
'    PASSWS_PROFES.MATI14.ColWidth(1) = 2000
'    PASSWS_PROFES.MATI14.Row = 0
'    PASSWS_PROFES.MATI14.Col = 0
'    PASSWS_PROFES.MATI14.CellForeColor = RGB(255, 255, 255)
'    PASSWS_PROFES.MATI14.CellBackColor = RGB(0, 0, 150)
'    PASSWS_PROFES.MATI14.Text = "NOMBRE DEL PROFESOR"
'    PASSWS_PROFES.MATI14.Col = 1
'    PASSWS_PROFES.MATI14.CellForeColor = RGB(255, 255, 255)
'    PASSWS_PROFES.MATI14.CellBackColor = RGB(0, 0, 150)
'    PASSWS_PROFES.MATI14.Text = "PASSWORD"
'    que = 0
'    NAR = FreeFile
'    Open Ruta & "CLAPRO.EDU" For Random As #NAR Len = Len(CLAV)
'    While Not EOF(NAR)
'        que = que + 1
'        Get #NAR, que, CLAV
'    Wend
'    Close #NAR
'    For J = 1 To que - 1
'        Open Ruta & "CLAPRO.EDU" For Random As #NAR Len = Len(CLAV)
'        Get #NAR, J, CLAV
'        Close #NAR
'        Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
'        Get #NAR, CLAV.NUMERO, profe
'        Close #NAR
'        PASSWS_PROFES.MATI14.Rows = J + 1
'        PASSWS_PROFES.MATI14.TextMatrix(J, 0) = RTrim(profe.nombres) & " " & RTrim(profe.apellidos) & "(" & CLAV.NUMERO & ")"
'        PASSWS_PROFES.MATI14.TextMatrix(J, 1) = CLAV.PASSWW
'    Next J
'    PASSWS_PROFES.Show 1
'End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Creación de cuentas de acceso en red para profesores."
End Sub

Private Sub Form_Load()
If Dir(Ruta & "infcur.edu") = "" Then
    Command1.Enabled = False
    'Command2.Enabled = False
Else
    Command1.Enabled = True
    'Command2.Enabled = True
End If
Text1.MaxLength = 3
Text2.MaxLength = 15
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
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

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Command1.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command1_Click
End If
End Sub
