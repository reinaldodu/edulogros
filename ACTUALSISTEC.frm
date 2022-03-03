VERSION 5.00
Begin VB.Form ACTUALSISTEC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizar Subsistema"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   Icon            =   "ACTUALSISTEC.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   3495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Actualizar"
      Height          =   320
      Left            =   1080
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "ACTUALSISTEC.frx":0442
         Left            =   1200
         List            =   "ACTUALSISTEC.frx":0464
         TabIndex        =   1
         Text            =   "SUBSISTEMA No.1"
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "ACTUALSISTEC.frx":0513
         Left            =   1200
         List            =   "ACTUALSISTEC.frx":0526
         TabIndex        =   2
         Text            =   "PRIMERO"
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "SUBSISTEMA"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PERIODO"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
   End
End
Attribute VB_Name = "ACTUALSISTEC"
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

Private Sub Combo3_Change()
If Combo3.Text <> Combo3.List(0) Then
    Combo3.Text = Combo3.List(0)
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo1.SetFocus
End If
End Sub

Private Sub Command1_Click()
VVAA = False
PENTI = False
If Combo3.ListIndex = -1 Then
    Combo3.ListIndex = 0
End If
If Dir(Ruta & "subsis" & Combo3.ListIndex + 1 & ".sub") = "" Then
    MsgBox "No se puede actualizar Subsistema.  El Subsistema no tiene grupos creados", 64, "Actualizar Subsistema"
    Exit Sub
End If
RESP = MsgBox("Desea actualizar la información del Subsistema No." & Combo3.ListIndex + 1 & " (Inserte un diskette sin información en la unidad A)?", vbYesNo + vbQuestion + vbDefaultButton1, "Actualizar")
If RESP = vbYes Then
    Screen.MousePointer = 11
    If Combo1.Text = "PRIMERO" Then
        lw = 1
    End If
    If Combo1.Text = "SEGUNDO" Then
        lw = 2
    End If
    If Combo1.Text = "TERCERO" Then
        lw = 3
    End If
    If Combo1.Text = "CUARTO" Then
        lw = 4
    End If
    If Combo1.Text = "FINAL" Then
        lw = 5
    End If
    On Error Resume Next
    Err.Clear
    If (Dir("A:\", vbDirectory) <> "") Or Dir("A:\", vbArchive) <> "" Then
            MsgBox "INSERTE UN DISKETTE QUE NO CONTENGA INFORMACION", 48, "ADVERTENCIA"
            Screen.MousePointer = 0
            Exit Sub
    End If
    If Dir("a:\subactu", vbDirectory) = "" Then
        MkDir "a:\subactu"
        If Err.Number <> 0 Then
            MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    If Dir(Ruta & "inicial.edu") <> "" Then
        FileCopy Ruta & "inicial.edu", "a:\subactu\inicial.edu"
    End If
    If Dir(Ruta & "cont.edu") <> "" Then
        FileCopy Ruta & "cont.edu", "a:\subactu\cont.edu"
    End If
    If Dir(Ruta & "prinalu.edu") <> "" Then
        FileCopy Ruta & "prinalu.edu", "a:\subactu\prinalu.edu"
    End If
    If Dir(Ruta & "quegru.edu") <> "" Then
        FileCopy Ruta & "quegru.edu", "a:\subactu\quegru.edu"
    End If
    If Dir(Ruta & "contpro.edu") <> "" Then
        FileCopy Ruta & "contpro.edu", "a:\subactu\contpro.edu"
    End If
    If Dir(Ruta & "prinpro.edu") <> "" Then
        FileCopy Ruta & "prinpro.edu", "a:\subactu\prinpro.edu"
    End If
    If Dir(Ruta & "materia.edu") <> "" Then
        FileCopy Ruta & "materia.edu", "a:\subactu\materia.edu"
    End If
    If Dir(Ruta & "infcur.edu") <> "" Then
        FileCopy Ruta & "infcur.edu", "a:\subactu\infcur.edu"
    End If
    If Dir(Ruta & "areagra.edu") <> "" Then
        FileCopy Ruta & "areagra.edu", "a:\subactu\areagra.edu"
    End If
    If Dir(Ruta & "webhelp.txt") <> "" Then
        FileCopy Ruta & "webhelp.txt", "a:\subactu\webhelp.txt"
    End If
    NAR = FreeFile
    Open Ruta & "subsis" & Combo3.ListIndex + 1 & ".sub" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, TTT
        If Dir(Ruta & RTrim(TTT) & ".gru") <> "" Then
            FileCopy Ruta & RTrim(TTT) & ".gru", "A:\SUBACTU\" & RTrim(TTT) & ".gru"
            If Err.Number <> 0 Then
                MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                Screen.MousePointer = 0
                Close #NAR
                Exit Sub
            End If
        End If
        NAR = FreeFile
        CERD = 0
        Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
        While Not EOF(NAR)
            CERD = CERD + 1
            Get #NAR, CERD, argra
            If (RTrim(argra.nom_grup) = RTrim(TTT)) Then
                If Dir(Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".obs") <> "" Then
                    FileCopy Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".obs", "a:\subactu\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".obs"
                    If Err.Number <> 0 Then
                        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                        Screen.MousePointer = 0
                        Close #NAR
                        Exit Sub
                    End If
                    VVAA = True
                End If
                fl = Left(argra.nom_grup, 1)
                If Dir(Ruta & fl & Left(argra.grado, 3) & argra.num_area & lw & ".lgr") <> "" Then
                    FileCopy Ruta & fl & Left(argra.grado, 3) & argra.num_area & lw & ".lgr", "a:\subactu\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".lgr"
                    If Err.Number <> 0 Then
                        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                        Screen.MousePointer = 0
                        Close #NAR
                        Exit Sub
                    End If
                    VVAA = True
                End If
                PENTI = True
            End If
        Wend
        Close #NAR
        NAR = NAR - 1
        If Dir(Ruta & "lrf" & RTrim(TTT) & ".lrf") <> "" Then
            FileCopy Ruta & "lrf" & RTrim(TTT) & ".lrf", "a:\subactu\lrf" & RTrim(TTT) & ".lrf"
            If Err.Number <> 0 Then
                MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                Screen.MousePointer = 0
                Close #NAR
                Exit Sub
            End If
        End If
        If Dir(Ruta & "orf" & RTrim(TTT) & ".orf") <> "" Then
            FileCopy Ruta & "orf" & RTrim(TTT) & ".orf", "a:\subactu\orf" & RTrim(TTT) & ".orf"
            If Err.Number <> 0 Then
                MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                Screen.MousePointer = 0
                Close #NAR
                Exit Sub
            End If
        End If
    Wend
    Close #NAR
    If PENTI = False Then
        MsgBox "No se le han asignado áreas a los grupos del Subsistema (áreas por grupo)", 32, "Actualizar Subsistema"
        Screen.MousePointer = 0
        Exit Sub
    End If
    If VVAA = False Then
        MsgBox "No existe información para actualizar el Subsistema en este periodo", 32, "Actualizar Subsistema"
    Else
        FileCopy Ruta & "subsis" & Combo3.ListIndex + 1 & ".sub", "a:\subactu\subsis" & Combo3.ListIndex + 1 & ".sub"
        infsub.subsistema = Combo3.ListIndex + 1
        infsub.actualsub = Format(Date, "mmm/dd/yyyy")
        infsub.bajasub = ""
        Open Ruta & "infosub.edu" For Append As #NAR
        Write #NAR, infsub.subsistema, infsub.actualsub, infsub.bajasub
        Close #NAR
        MsgBox "Copia Exitosa", 64, "Actualizar Subsistema"
        Unload Me
    End If
End If
Screen.MousePointer = 0
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Copia la información necesaria en un diskette para actualizar el Subsistema de acuerdo al periodo seleccionado."
End Sub

Private Sub Form_Load()
If (Dir(Ruta & "infcur.edu") <> "") And (Dir(Ruta & "areagra.edu") <> "") And (Dir(Ruta & "materia.edu") <> "") Then
    Command1.Enabled = True
Else
    Command1.Enabled = False
End If
End Sub
