VERSION 5.00
Begin VB.Form DISKINISUB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear Disco Inicial"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   Icon            =   "DISKINISUB.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Crear disco"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione el Subsistema"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         ItemData        =   "DISKINISUB.frx":0442
         Left            =   240
         List            =   "DISKINISUB.frx":0464
         TabIndex        =   1
         Text            =   "SUBSISTEMA No.1"
         Top             =   480
         Width           =   2295
      End
   End
End
Attribute VB_Name = "DISKINISUB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
If Combo1.Text <> Combo1.List(0) Then
    Combo1.Text = Combo1.List(0)
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If Command1.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command1_Click
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
Err.Clear
If Combo1.ListIndex = -1 Then
    Combo1.ListIndex = 0
End If
If Dir(Ruta & "subsis" & Combo1.ListIndex + 1 & ".sub") = "" Then
    MsgBox "No se puede crear el Disco Inicial.  El Subsistema no tiene grupos creados", 64, "Disco Inicial"
    Exit Sub
End If
RESP = MsgBox("Desea crear el Disco Inicial para el Subsistema No." & Combo1.ListIndex + 1 & "?", vbYesNo + vbQuestion + vbDefaultButton1, "Disco Inicial")
If RESP = vbYes Then
    Screen.MousePointer = 11
    If (Dir("A:\", vbDirectory) <> "") Or Dir("A:\", vbArchive) <> "" Then
        MsgBox "INSERTE UN DISKETTE QUE NO CONTENGA INFORMACION", 48, "ADVERTENCIA"
        Screen.MousePointer = 0
        Exit Sub
    End If
    If Dir("a:\disksub", vbDirectory) = "" Then
        MkDir "a:\disksub"
        If Err.Number <> 0 Then
            MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    NAR = FreeFile
    Open Ruta & "subsis" & Combo1.ListIndex + 1 & ".sub" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, TTT
        FileCopy Ruta & RTrim(TTT) & ".gru", "A:\DISKSUB\" & RTrim(TTT) & ".gru"
        If Err.Number <> 0 Then
            MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
            Screen.MousePointer = 0
            Close #NAR
            Exit Sub
        End If
    Wend
    Close #NAR
    If Dir(Ruta & "inicial.edu") <> "" Then
        FileCopy Ruta & "inicial.edu", "a:\disksub\inicial.edu"
    End If
    If Dir(Ruta & "cont.edu") <> "" Then
        FileCopy Ruta & "cont.edu", "a:\disksub\cont.edu"
    End If
    If Dir(Ruta & "prinalu.edu") <> "" Then
        FileCopy Ruta & "prinalu.edu", "a:\disksub\prinalu.edu"
    End If
    If Dir(Ruta & "quegru.edu") <> "" Then
        FileCopy Ruta & "quegru.edu", "a:\disksub\quegru.edu"
    End If
    If Dir(Ruta & "contpro.edu") <> "" Then
        FileCopy Ruta & "contpro.edu", "a:\disksub\contpro.edu"
    End If
    If Dir(Ruta & "prinpro.edu") <> "" Then
        FileCopy Ruta & "prinpro.edu", "a:\disksub\prinpro.edu"
    End If
    If Dir(Ruta & "materia.edu") <> "" Then
        FileCopy Ruta & "materia.edu", "a:\disksub\materia.edu"
    End If
    If Dir(Ruta & "subsis" & Combo1.ListIndex + 1 & ".sub") <> "" Then
        FileCopy Ruta & "subsis" & Combo1.ListIndex + 1 & ".sub", "a:\disksub\subsis" & Combo1.ListIndex + 1 & ".sub"
    End If
    If Dir(Ruta & "infcur.edu") <> "" Then
        FileCopy Ruta & "infcur.edu", "a:\disksub\infcur.edu"
    End If
    If Dir(Ruta & "areagra.edu") <> "" Then
        FileCopy Ruta & "areagra.edu", "a:\disksub\areagra.edu"
    End If
    If Dir(Ruta & "webhelp.txt") <> "" Then
        FileCopy Ruta & "webhelp.txt", "a:\disksub\webhelp.txt"
    End If
    If Dir(Ruta & "subsist.bat") <> "" Then
        FileCopy Ruta & "subsist.bat", "a:\disksub\subsist.bat"
    End If
    Screen.MousePointer = 0
    MsgBox "Disco creado con éxito", 64, "Disco Inicial"
    Unload Me
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Disco Inicial Subsistema: Se utiliza para terminar de instalar el Subsistema."
End Sub

Private Sub Form_Load()
If (Dir(Ruta & "infcur.edu") <> "") And (Dir(Ruta & "areagra.edu") <> "") And (Dir(Ruta & "materia.edu") <> "") Then
    Command1.Enabled = True
Else
    Command1.Enabled = False
End If
End Sub
