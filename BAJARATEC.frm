VERSION 5.00
Begin VB.Form BAJARATEC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bajar Disco-Subsistema"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   Icon            =   "BAJARATEC.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   3495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   320
      Left            =   1080
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "BAJARATEC.frx":0442
         Left            =   1200
         List            =   "BAJARATEC.frx":0464
         TabIndex        =   1
         Text            =   "SUBSISTEMA No.1"
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "BAJARATEC.frx":0513
         Left            =   1200
         List            =   "BAJARATEC.frx":0526
         TabIndex        =   2
         Text            =   "PRIMERO"
         Top             =   600
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
         Top             =   720
         Width           =   735
      End
   End
End
Attribute VB_Name = "BAJARATEC"
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
    MsgBox "No se puede bajar el Disco.  El Subsistema no tiene grupos creados", 64, "Bajar Disco"
    Exit Sub
End If
RESP = MsgBox("Desea bajar el Disco del Subsistema No." & Combo3.ListIndex + 1 & "?", vbYesNo + vbQuestion + vbDefaultButton1, "Bajar Disco")
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
    If Dir("a:\subsist\subsis" & Combo3.ListIndex + 1 & ".sub") = "" Then
        MsgBox "INSERTE EL DISKETTE DEL SUBSISTEMA No." & Combo3.ListIndex + 1 & " EN LA UNIDAD A", 16, "ADVERTENCIA"
        Screen.MousePointer = 0
        Exit Sub
    End If
    NAR = FreeFile
    Open Ruta & "subsis" & Combo3.ListIndex + 1 & ".sub" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, TTT
        NAR = FreeFile
        CERD = 0
        Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
        While Not EOF(NAR)
            CERD = CERD + 1
            Get #NAR, CERD, argra
            If (RTrim(argra.nom_grup) = RTrim(TTT)) Then
                If Dir("a:\subsist\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".obs") = "" Then
                    RESP = MsgBox("Grupo " & Format(RTrim(argra.nom_grup), "<") & " no tiene notas para el área No." & argra.num_area & ".  Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2, "Diskette incompleto")
                    If RESP = vbYes Then
                        GoTo bacontec
                    Else
                        Close #NAR
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                Else
                    If FileLen("a:\subsist\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".obs") = 0 Then
                        RESP = MsgBox("Grupo " & Format(RTrim(argra.nom_grup), "<") & " no tiene notas para el área No." & argra.num_area & ".  Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2, "Diskette incompleto")
                        If RESP = vbYes Then
                            GoTo bacontec
                        Else
                            Close #NAR
                            Screen.MousePointer = 0
                            Exit Sub
                        End If
                    End If
                    FileCopy "a:\subsist\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".obs", Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".obs"
                    If Err.Number <> 0 Then
                        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                        Screen.MousePointer = 0
                        Close #NAR
                        Exit Sub
                    End If
                    VVAA = True
                End If
bacontec:
                fl = Left(argra.nom_grup, 1)
                If Dir("a:\subsist\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".lgr") = "" Then
                    RESP = MsgBox("No existen observaciones para el grado " & Format(RTrim(argra.grado), "<") & " área No." & argra.num_area & ".  Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2, "Diskette incompleto")
                    If RESP = vbYes Then
                        GoTo bacontec2
                    Else
                        Close #NAR
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                Else
                    If FileLen("a:\subsist\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".lgr") = 0 Then
                        RESP = MsgBox("No existen observaciones para el grado " & Format(RTrim(argra.grado), "<") & " área No." & argra.num_area & ".  Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2, "Diskette incompleto")
                        If RESP = vbYes Then
                            GoTo bacontec2
                        Else
                            Close #NAR
                            Screen.MousePointer = 0
                            Exit Sub
                        End If
                    End If
                    FileCopy "a:\subsist\" & fl & Left(argra.grado, 3) & argra.num_area & lw & ".lgr", Ruta & fl & Left(argra.grado, 3) & argra.num_area & lw & ".lgr"
                    If Err.Number <> 0 Then
                        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                        Screen.MousePointer = 0
                        Close #NAR
                        Exit Sub
                    End If
                    VVAA = True
                End If
bacontec2:
            PENTI = True
            End If
        Wend
        Close #NAR
        NAR = NAR - 1
        If Dir("a:\subsist\lrf" & RTrim(TTT) & ".lrf") <> "" Then
            FileCopy "a:\subsist\lrf" & RTrim(TTT) & ".lrf", Ruta & "lrf" & RTrim(TTT) & ".lrf"
            If Err.Number <> 0 Then
                MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                Screen.MousePointer = 0
                Close #NAR
                Exit Sub
            End If
        End If
        If Dir("a:\subsist\orf" & RTrim(TTT) & ".orf") <> "" Then
            FileCopy "a:\subsist\orf" & RTrim(TTT) & ".orf", Ruta & "orf" & RTrim(TTT) & ".orf"
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
        MsgBox "No se le han asignado áreas a los grupos del Subsistema (áreas por grupo)", 32, "Bajar Disco"
        Screen.MousePointer = 0
        Exit Sub
    End If
    If VVAA = False Then
        MsgBox "No existe información del Subsistema para bajar en este periodo", 32, "Bajar Disco"
    Else
        infsub.subsistema = Combo3.ListIndex + 1
        infsub.bajasub = Format(Date, "mmm/dd/yyyy")
        infsub.actualsub = ""
        Open Ruta & "infosub.edu" For Append As #NAR
        Write #NAR, infsub.subsistema, infsub.actualsub, infsub.bajasub
        Close #NAR
        MsgBox "Copia Exitosa", 64, "Bajar Disco"
        Unload Me
    End If
End If
Screen.MousePointer = 0
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Copia la información que contiene el diskette de boletines del Subsistema seleccionado al sistema principal."
End Sub

Private Sub Form_Load()
If (Dir(Ruta & "infcur.edu") <> "") And (Dir(Ruta & "areagra.edu") <> "") And (Dir(Ruta & "materia.edu") <> "") Then
    Command1.Enabled = True
Else
    Command1.Enabled = False
End If
End Sub
