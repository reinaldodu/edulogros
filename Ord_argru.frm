VERSION 5.00
Begin VB.Form Ord_argru 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenar áreas por grupo"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3615
   Icon            =   "Ord_argru.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   2880
      Picture         =   "Ord_argru.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Guardar este orden"
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   2880
      Picture         =   "Ord_argru.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Bajar un nivel"
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   2880
      Picture         =   "Ord_argru.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Subir un nivel"
      Top             =   240
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.ComboBox Ver_grupo 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3360
         Width           =   2415
      End
      Begin VB.ListBox List1 
         Height          =   2985
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "GRUPO:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   3480
         Width           =   630
      End
   End
End
Attribute VB_Name = "Ord_argru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VerArea As String
Private Sub Command1_Click()
If List1.ListIndex > 0 Then
    VerArea = List1.List(List1.ListIndex)
    List1.List(List1.ListIndex) = List1.List(List1.ListIndex - 1)
    List1.ListIndex = List1.ListIndex - 1
    List1.List(List1.ListIndex) = VerArea
    Command3.Enabled = True
End If
End Sub

Private Sub Command2_Click()
If (List1.ListIndex < List1.ListCount - 1) And (List1.ListIndex >= 0) Then
    VerArea = List1.List(List1.ListIndex)
    List1.List(List1.ListIndex) = List1.List(List1.ListIndex + 1)
    List1.ListIndex = List1.ListIndex + 1
    List1.List(List1.ListIndex) = VerArea
    Command3.Enabled = True
End If
End Sub

Private Sub Command3_Click()
RESP = MsgBox("Desea guardar este orden?", vbYesNo + vbQuestion + vbDefaultButton1, "Guardar")
If RESP = vbYes Then
    Screen.MousePointer = 11
    NAR = FreeFile
    cona = 0
    h = 1
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, argra
        If RTrim(argra.nom_grup) <> Ver_grupo.Text Then
            NAR = FreeFile
            Open Ruta & "areagra2.edu" For Random As #NAR Len = Len(argra)
            Put #NAR, h, argra
            Close #NAR
            h = h + 1
            NAR = NAR - 1
        End If
    Wend
    Close #NAR
    For I = 0 To List1.ListCount - 1
        cona = 0
        NAR = FreeFile
        Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
        While Not EOF(NAR)
            cona = cona + 1
            Get #NAR, cona, mate
            If RTrim(mate.nom) = List1.List(I) Then
                cona2 = 0
                NAR = FreeFile
                Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
                While Not EOF(NAR)
                    cona2 = cona2 + 1
                    Get #NAR, cona2, argra
                    If (RTrim(argra.nom_grup) = Ver_grupo.Text) And (argra.num_area = mate.num) Then
                        NAR = FreeFile
                        Open Ruta & "areagra2.edu" For Random As #NAR Len = Len(argra)
                        Put #NAR, h, argra
                        Close #NAR
                        h = h + 1
                        NAR = NAR - 1
                    End If
                Wend
                Close #NAR
                NAR = NAR - 1
            End If
        Wend
        Close #NAR
    Next I
    If Dir(Ruta & "areagra2.edu") <> "" Then
        FileCopy Ruta & "areagra2.edu", Ruta & "areagra.edu"
        Kill Ruta & "areagra2.edu"
    End If
    Screen.MousePointer = 0
End If
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
If Dir(Ruta & "infcur.edu") <> "" Then
    NAR = FreeFile
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
        Ver_grupo.AddItem RTrim(icur.nom)
    Wend
    Close #NAR
    Ver_grupo.Text = Ver_grupo.List(0)
Else
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Ver_grupo.Enabled = False
End If
End Sub

Private Sub Ver_grupo_Click()
List1.Clear
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
If Dir(Ruta & "areagra.edu") <> "" Then
    cona = 0
    OkArea = False
    NAR = FreeFile
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, argra
        If RTrim(argra.nom_grup) = Ver_grupo.Text Then
            NAR = FreeFile
            Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
            Get #NAR, argra.num_area, mate
            Close #NAR
            NAR = NAR - 1
            List1.AddItem RTrim(mate.nom)
            OkArea = True
        End If
    Wend
    Close #NAR
    If OkArea = True Then
        Command1.Enabled = True
        Command2.Enabled = True
    End If
End If
End Sub
