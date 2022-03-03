VERSION 5.00
Begin VB.Form ACTUDISKPRO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizar Disco-profesor"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "ACTUDISKPRO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.TextBox Text1 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         ItemData        =   "ACTUDISKPRO.frx":0442
         Left            =   1560
         List            =   "ACTUDISKPRO.frx":0455
         TabIndex        =   2
         Text            =   "PRIMERO"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "PERIODO:"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PROFESOR No."
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   1185
      End
   End
End
Attribute VB_Name = "ACTUDISKPRO"
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
Frame1.Caption = ""
If Text1.Text = "" Then
    MsgBox "Escriba el número de profesor", 64, "Actualizar"
    Text1.SetFocus
    Exit Sub
End If
On Error Resume Next
Err.Clear
'If (Dir("a:\datos\inicial.edu") = "") And (Dir("a:\datos\clase.edu") = "") Then
If (Dir(RutaDir & "\inicial.edu") = "") And (Dir(RutaDir & "\clase.edu") = "") Then
    MsgBox "Directorio inválido para actualizar", 16, "Actualizar"
    Exit Sub
End If
w = Val(Text1.Text)
NAR = FreeFile
'Open "a:\datos\clase.edu" For Random As #NAR Len = Len(CLAV)
Open RutaDir & "\clase.edu" For Random As #NAR Len = Len(CLAV)
Get #NAR, 1, CLAV
Close #NAR
If CLAV.NUMERO <> w Then
    MsgBox "Diskette no corresponde al número de profesor", 16, "Actualizar"
    Text1.SetFocus
    Exit Sub
End If
Open Ruta & "contpro.edu" For Input As #NAR
Input #NAR, r
Close #NAR
If ((w > r - 1) Or (w < 1)) Then
    MsgBox "Profesor no existe", 32, "Actualizar"
    Text1.SetFocus
    Exit Sub
End If
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Get #NAR, w, profe
Close #NAR
If ((RTrim(profe.nombres) = "") And (RTrim(profe.especiali) = "")) Then
    MsgBox "REGISTRO NO EXISTE", 64, "Actualizar"
    Text1.SetFocus
    Exit Sub
End If
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
Frame1.Caption = RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
RESP = MsgBox("Desea actualizar el disco-datos de este profesor?", vbYesNo + vbQuestion + vbDefaultButton1, "Actualizar")
If RESP = vbYes Then
    Screen.MousePointer = 11
    If Dir(Ruta & "inicial.edu") <> "" Then
        'FileCopy Ruta & "inicial.edu", "a:\datos\inicial.edu"
        FileCopy Ruta & "inicial.edu", RutaDir & "\inicial.edu"
        If Err.Number <> 0 Then
            MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    If Dir(Ruta & "infcur.edu") <> "" Then
        'FileCopy Ruta & "infcur.edu", "a:\datos\infcur.edu"
        FileCopy Ruta & "infcur.edu", RutaDir & "\infcur.edu"
        If Err.Number <> 0 Then
            MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    If Dir(Ruta & "prinpro.edu") <> "" Then
        'FileCopy Ruta & "prinpro.edu", "a:\datos\prinpro.edu"
        FileCopy Ruta & "prinpro.edu", RutaDir & "\prinpro.edu"
        If Err.Number <> 0 Then
            MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    If Dir(Ruta & "prinalu.edu") <> "" Then
        'FileCopy Ruta & "prinalu.edu", "a:\datos\prinalu.edu"
        FileCopy Ruta & "prinalu.edu", RutaDir & "\prinalu.edu"
        If Err.Number <> 0 Then
            MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    If Dir(Ruta & "cont.edu") <> "" Then
        'FileCopy Ruta & "cont.edu", "a:\datos\cont.edu"
        FileCopy Ruta & "cont.edu", RutaDir & "\cont.edu"
        If Err.Number <> 0 Then
            MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    If Dir(Ruta & "materia.edu") <> "" Then
        'FileCopy Ruta & "materia.edu", "a:\datos\materia.edu"
        FileCopy Ruta & "materia.edu", RutaDir & "\materia.edu"
        If Err.Number <> 0 Then
            MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    If Dir(Ruta & "areagra.edu") <> "" Then
        'FileCopy Ruta & "areagra.edu", "a:\datos\areagra.edu"
        FileCopy Ruta & "areagra.edu", RutaDir & "\areagra.edu"
        If Err.Number <> 0 Then
            MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    If Dir(Ruta & "contpro.edu") <> "" Then
        'FileCopy Ruta & "contpro.edu", "a:\datos\contpro.edu"
        FileCopy Ruta & "contpro.edu", RutaDir & "\contpro.edu"
        If Err.Number <> 0 Then
            MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
'    If Dir(Ruta & "retialu.edu") <> "" Then
'        'FileCopy Ruta & "retialu.edu", "a:\datos\retialu.edu"
'        FileCopy Ruta & "retialu.edu", RutaDir & "\retialu.edu"
'        If Err.Number <> 0 Then
'            MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
'            Screen.MousePointer = 0
'            Exit Sub
'        End If
'    End If
    CERD = 0
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
        CERD = CERD + 1
        Get #NAR, CERD, argra
        If argra.num_pro = w Then
            If Dir(Ruta & RTrim(argra.nom_grup) & ".gru") <> "" Then
                'FileCopy Ruta & RTrim(argra.nom_grup) & ".gru", "a:\datos\" & RTrim(argra.nom_grup) & ".gru"
                FileCopy Ruta & RTrim(argra.nom_grup) & ".gru", RutaDir & "\" & RTrim(argra.nom_grup) & ".gru"
                If Err.Number <> 0 Then
                    MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                    Screen.MousePointer = 0
                    Close #NAR
                    Exit Sub
                End If
            End If
            If Dir(Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".obs") <> "" Then
                'FileCopy Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".obs", "a:\datos\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".obs"
                FileCopy Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".obs", RutaDir & "\" & RTrim(argra.nom_grup) & argra.num_area & lw & ".obs"
                If Err.Number <> 0 Then
                    MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                    Screen.MousePointer = 0
                    Close #NAR
                    Exit Sub
                End If
            End If
            If Dir(Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".lgr") <> "" Then
                If FileLen(Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".lgr") <> 0 Then
                    'FileCopy Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".lgr", "a:\datos\" & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".lgr"
                    FileCopy Ruta & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".lgr", RutaDir & "\" & Left(argra.nom_grup, 1) & Left(argra.grado, 3) & argra.num_area & lw & ".lgr"
                    If Err.Number <> 0 Then
                        MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                        Screen.MousePointer = 0
                        Close #NAR
                        Exit Sub
                    End If
                End If
            End If
            If Dir(Ruta & "lrf" & RTrim(argra.nom_grup) & ".lrf") <> "" Then
                'FileCopy Ruta & "lrf" & RTrim(argra.nom_grup) & ".lrf", "a:\datos\lrf" & RTrim(argra.nom_grup) & ".lrf"
                FileCopy Ruta & "lrf" & RTrim(argra.nom_grup) & ".lrf", RutaDir & "\lrf" & RTrim(argra.nom_grup) & ".lrf"
                If Err.Number <> 0 Then
                    MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                    Screen.MousePointer = 0
                    Close #NAR
                    Exit Sub
                End If
            End If
            If Dir(Ruta & "orf" & RTrim(argra.nom_grup) & ".orf") <> "" Then
                'FileCopy Ruta & "orf" & RTrim(argra.nom_grup) & ".orf", "a:\datos\orf" & RTrim(argra.nom_grup) & ".orf"
                FileCopy Ruta & "orf" & RTrim(argra.nom_grup) & ".orf", RutaDir & "\orf" & RTrim(argra.nom_grup) & ".orf"
                If Err.Number <> 0 Then
                    MsgBox "Error #" & Str(Err.Number) & " " & Err.Description, , "Error"
                    Screen.MousePointer = 0
                    Close #NAR
                    Exit Sub
                End If
            End If
        End If
    Wend
    Close #NAR
    Screen.MousePointer = 0
    MsgBox "Actualización de disco finalizada", 64, "Actualizar"
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Actualiza la información del diskette de un profesor de acuerdo al periodo seleccionado."
End Sub

Private Sub Form_Load()
Text1.MaxLength = 3
If (Dir(Ruta & "prinpro.edu") <> "") And (Dir(Ruta & "materia.edu") <> "") And (Dir(Ruta & "areagra.edu") <> "") And (Dir(Ruta & "infcur.edu") <> "") Then
    Command1.Enabled = True
Else
    Command1.Enabled = False
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo1.SetFocus
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
