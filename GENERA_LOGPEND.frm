VERSION 5.00
Begin VB.Form GENERA_LOGPEND 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar logros pendientes"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "GENERA_LOGPEND.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.CheckBox Check1 
         Caption         =   "Generar toda la jornada"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2055
      End
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   3135
         Begin VB.ComboBox Combo3 
            BackColor       =   &H00800000&
            ForeColor       =   &H0000FFFF&
            Height          =   315
            ItemData        =   "GENERA_LOGPEND.frx":0442
            Left            =   1200
            List            =   "GENERA_LOGPEND.frx":0476
            TabIndex        =   3
            Text            =   "PREKINDER"
            Top             =   960
            Width           =   1695
         End
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00800000&
            ForeColor       =   &H0000FFFF&
            Height          =   315
            ItemData        =   "GENERA_LOGPEND.frx":04FF
            Left            =   1200
            List            =   "GENERA_LOGPEND.frx":050F
            TabIndex        =   2
            Text            =   "UNICA"
            Top             =   600
            Width           =   1695
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00800000&
            ForeColor       =   &H0000FFFF&
            Height          =   315
            ItemData        =   "GENERA_LOGPEND.frx":0530
            Left            =   1200
            List            =   "GENERA_LOGPEND.frx":0543
            TabIndex        =   1
            Text            =   "PRIMERO"
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "GRADO:"
            Height          =   195
            Left            =   240
            TabIndex        =   9
            Top             =   1080
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "JORNADA:"
            Height          =   195
            Left            =   240
            TabIndex        =   8
            Top             =   720
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "PERIODO:"
            Height          =   195
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   780
         End
      End
   End
End
Attribute VB_Name = "GENERA_LOGPEND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gnlp As Boolean
Private Sub gene_lopend()
If Dir(Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".obp") = "" Then
    If Dir(Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".obs") <> "" Then
        If FileLen(Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".obs") <> 0 Then
            gnlp = True
            FileCopy Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".obs", Ruta & RTrim(argra.nom_grup) & argra.num_area & lw & ".obp"
        End If
    End If
End If
End Sub

Private Sub Combo1_Change()
If Combo1.Text <> Combo1.List(0) Then
    Combo1.Text = Combo1.List(0)
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo2.SetFocus
End If
End Sub

Private Sub Combo2_Change()
If Combo2.Text <> Combo2.List(0) Then
    Combo2.Text = Combo2.List(0)
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo3.SetFocus
End If
End Sub

Private Sub Combo3_Change()
If Combo3.Text <> Combo3.List(0) Then
    Combo3.Text = Combo3.List(0)
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If Command1.Enabled = False Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Command1_Click
End If
End Sub

Private Sub Command1_Click()
If Check1.Value = 1 Then
    MS1 = "Desea generar los logros pendientes de la jornada " & Format(Combo2.Text, "<") & ", periodo " & Format(Combo1.Text, "<") & "?"
Else
    MS1 = "Desea generar los logros pendientes del grado " & Format(Combo3.Text, "<") & ", jornada " & Format(Combo2.Text, "<") & ", periodo " & Format(Combo1.Text, "<") & "?"
End If
RESP = MsgBox(MS1, vbYesNo + vbQuestion + vbDefaultButton2, "Generar logros pendientes")
If RESP = vbYes Then
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
    If Combo2.Text = "UNICA" Then
        fl = "1"
    End If
    If Combo2.Text = "MAÑANA" Then
        fl = "2"
    End If
    If Combo2.Text = "TARDE" Then
        fl = "3"
    End If
    If Combo2.Text = "NOCHE" Then
        fl = "4"
    End If
    IMPOK = False
    gnlp = False
    Screen.MousePointer = 11
    cona = 0
    NAR = FreeFile
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, argra
        If Check1.Value = 1 Then
            If Left((argra.nom_grup), 1) = fl Then
                Call gene_lopend
                IMPOK = True
            End If
        Else
            If (Left((argra.nom_grup), 1) = fl) And (RTrim(argra.grado) = Combo3.Text) Then
                Call gene_lopend
                IMPOK = True
            End If
        End If
    Wend
    Close #NAR
    If IMPOK = False Then
        MsgBox "No existe información para generar logros pendientes", 48, "Generar logros pendientes"
        Screen.MousePointer = 0
        Exit Sub
    End If
    If gnlp = False Then
        MsgBox "Ya se habian generado estos logros pendientes", 64, "Generar logros pendientes"
        Screen.MousePointer = 0
        Exit Sub
    End If
    Screen.MousePointer = 0
    MsgBox "Logros pendientes generados con éxito", 64, "Generar logros pendientes"
    Unload Me
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Genera los archivos para el control de logros pendientes por jornada y/o grado."
End Sub

Private Sub Form_Load()
If (Dir(Ruta & "materia.edu") <> "") And (Dir(Ruta & "areagra.edu") <> "") And (Dir(Ruta & "infcur.edu") <> "") Then
    Command1.Enabled = True
Else
    Command1.Enabled = False
End If
End Sub
