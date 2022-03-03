VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Control_Planillas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de planillas cerradas"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   13980
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   13980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   6240
      TabIndex        =   2
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13695
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Control_Planillas.frx":0000
         Left            =   11040
         List            =   "Control_Planillas.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2535
      End
      Begin MSFlexGridLib.MSFlexGrid MT_Control 
         Height          =   5175
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   9128
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PERIODO:"
         Height          =   195
         Left            =   10200
         TabIndex        =   3
         Top             =   360
         Width           =   780
      End
   End
End
Attribute VB_Name = "Control_Planillas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Screen.MousePointer = 11
If Dir(Ruta & "infcur.edu") = "" Then
Screen.MousePointer = 0
Exit Sub
End If
If Dir(Ruta & "areagra.edu") = "" Then
Screen.MousePointer = 0
Exit Sub
End If
If Dir(Ruta & "materia.edu") = "" Then
Screen.MousePointer = 0
Exit Sub
End If
NAR = FreeFile
plo = 2
If RTrim(Combo1.Text) = "PRIMERO" Then
    lw = 1
End If
If RTrim(Combo1.Text) = "SEGUNDO" Then
    lw = 2
End If
If RTrim(Combo1.Text) = "TERCERO" Then
    lw = 3
End If
If RTrim(Combo1.Text) = "CUARTO" Then
    lw = 4
End If
If RTrim(Combo1.Text) = "FINAL" Then
    lw = 5
End If

Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
    Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
    MT_Control.Rows = plo
    'If plo = 2 Then
    'MT_Control.FixedRows = 1
    'End If
    MT_Control.Row = plo - 1
    MT_Control.Col = 0
    MT_Control.CellFontBold = True
    MT_Control.Text = RTrim(icur.nom)
    NAR = FreeFile
    cona = 0
    CONTAREA = 0
    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, argra
        If RTrim(icur.nom) = RTrim(argra.nom_grup) Then
        CONTAREA = CONTAREA + 1
        NAR = FreeFile
        Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
        Get #NAR, argra.num_area, mate
        Close #NAR
        Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
        Get #NAR, argra.num_pro, profe
        Close #NAR
        NAR = NAR - 1
        MT_Control.Rows = plo
        MT_Control.Row = plo - 1
        MT_Control.Col = 1
        'mt_control.Text = RTrim(mate.nom) & " (" & mate.num & ")"
        'MT_Control.Text = RTrim(mate.nom) & " (I.H:" & argra.ih & ")"
        MT_Control.Text = RTrim(mate.nom)
        MT_Control.Col = 2
        'MT_Control.Text = RTrim(profe.nombres) & " " & RTrim(profe.apellidos) & " (" & argra.num_pro & ")"
        MT_Control.Text = RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
        'Verificar archivos de control de planillas
        If Dir(Ruta & RTrim(argra.nom_grup) & mate.num & lw & ".fnp") <> "" Then
            MT_Control.Col = 3
            MT_Control.Text = "OK"
        End If
        plo = plo + 1
        End If
    Wend
    Close #NAR
    NAR = NAR - 1
    MT_Control.Rows = plo
    MT_Control.Row = plo - 1
    MT_Control.Col = 1
    MT_Control.CellFontBold = True
    MT_Control.CellForeColor = RGB(0, 0, 255)
    MT_Control.Text = "TOTAL MATERIAS..." & CONTAREA
    plo = plo + 1
Wend
Close #NAR
Screen.MousePointer = 0

End Sub

Private Sub Form_Load()
Combo1.Text = Combo1.List(0)
MT_Control.Row = 0
MT_Control.Col = 0
MT_Control.ColWidth(0) = 2500
MT_Control.CellFontBold = True
MT_Control.CellForeColor = RGB(255, 255, 255)
MT_Control.CellBackColor = RGB(0, 0, 150)
MT_Control.Text = "GRUPO"
MT_Control.Col = 1
MT_Control.ColWidth(1) = 5200
MT_Control.CellFontBold = True
MT_Control.CellForeColor = RGB(255, 255, 255)
MT_Control.CellBackColor = RGB(0, 0, 150)
MT_Control.Text = "MATERIA"
MT_Control.Col = 2
MT_Control.ColWidth(2) = 4000
MT_Control.CellFontBold = True
MT_Control.CellForeColor = RGB(255, 255, 255)
MT_Control.CellBackColor = RGB(0, 0, 150)
MT_Control.Text = "PROFESOR"
MT_Control.Col = 3
MT_Control.ColWidth(3) = 1000
MT_Control.CellFontBold = True
MT_Control.CellForeColor = RGB(255, 255, 255)
MT_Control.CellBackColor = RGB(0, 0, 150)
MT_Control.Text = "PLANILLA"
End Sub
