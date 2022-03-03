VERSION 5.00
Begin VB.Form CVSEXO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta por sexo"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2655
   Icon            =   "CVSEXO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   2655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   2415
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "CVSEXO.frx":0442
         Left            =   1200
         List            =   "CVSEXO.frx":044C
         TabIndex        =   0
         Text            =   "M"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SEXO:"
         Height          =   195
         Left            =   600
         TabIndex        =   3
         Top             =   360
         Width           =   480
      End
   End
End
Attribute VB_Name = "CVSEXO"
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
'Dim alumno As maestroalum
'Dim aluper As pertgrup
CONSVARIAS.MATRICON.Rows = 1
CONSVARIAS.MATRICON.Cols = 14
CONSVARIAS.MATRICON.Col = 0
CONSVARIAS.MATRICON.ColWidth(0) = 1000
CONSVARIAS.MATRICON.CellForeColor = RGB(255, 255, 255)
CONSVARIAS.MATRICON.CellBackColor = RGB(0, 0, 150)
CONSVARIAS.MATRICON.Text = "CARNET"
CONSVARIAS.MATRICON.Col = 1
CONSVARIAS.MATRICON.ColWidth(1) = 3900
CONSVARIAS.MATRICON.CellForeColor = RGB(255, 255, 255)
CONSVARIAS.MATRICON.CellBackColor = RGB(0, 0, 150)
CONSVARIAS.MATRICON.Text = "APELLIDOS Y NOMBRES"
CONSVARIAS.MATRICON.Col = 2
CONSVARIAS.MATRICON.ColWidth(2) = 800
CONSVARIAS.MATRICON.CellForeColor = RGB(255, 255, 255)
CONSVARIAS.MATRICON.CellBackColor = RGB(0, 0, 150)
CONSVARIAS.MATRICON.Text = "#MATR."
CONSVARIAS.MATRICON.Col = 3
CONSVARIAS.MATRICON.ColWidth(3) = 1200
CONSVARIAS.MATRICON.CellForeColor = RGB(255, 255, 255)
CONSVARIAS.MATRICON.CellBackColor = RGB(0, 0, 150)
CONSVARIAS.MATRICON.Text = "F_NACIM."
CONSVARIAS.MATRICON.Col = 4
CONSVARIAS.MATRICON.ColWidth(4) = 600
CONSVARIAS.MATRICON.CellForeColor = RGB(255, 255, 255)
CONSVARIAS.MATRICON.CellBackColor = RGB(0, 0, 150)
CONSVARIAS.MATRICON.Text = "EDAD"
CONSVARIAS.MATRICON.Col = 5
CONSVARIAS.MATRICON.ColWidth(5) = 500
CONSVARIAS.MATRICON.CellForeColor = RGB(255, 255, 255)
CONSVARIAS.MATRICON.CellBackColor = RGB(0, 0, 150)
CONSVARIAS.MATRICON.Text = "R.H."
CONSVARIAS.MATRICON.Col = 6
CONSVARIAS.MATRICON.ColWidth(6) = 600
CONSVARIAS.MATRICON.CellForeColor = RGB(255, 255, 255)
CONSVARIAS.MATRICON.CellBackColor = RGB(0, 0, 150)
CONSVARIAS.MATRICON.Text = "SEXO"
CONSVARIAS.MATRICON.Col = 7
CONSVARIAS.MATRICON.ColWidth(7) = 1200
CONSVARIAS.MATRICON.CellForeColor = RGB(255, 255, 255)
CONSVARIAS.MATRICON.CellBackColor = RGB(0, 0, 150)
CONSVARIAS.MATRICON.Text = "DOC. ID."
CONSVARIAS.MATRICON.Col = 8
CONSVARIAS.MATRICON.ColWidth(8) = 3600
CONSVARIAS.MATRICON.CellForeColor = RGB(255, 255, 255)
CONSVARIAS.MATRICON.CellBackColor = RGB(0, 0, 150)
CONSVARIAS.MATRICON.Text = "ACUDIENTE"
CONSVARIAS.MATRICON.Col = 9
CONSVARIAS.MATRICON.ColWidth(9) = 3600
CONSVARIAS.MATRICON.CellForeColor = RGB(255, 255, 255)
CONSVARIAS.MATRICON.CellBackColor = RGB(0, 0, 150)
CONSVARIAS.MATRICON.Text = "DIRECCION"
CONSVARIAS.MATRICON.Col = 10
CONSVARIAS.MATRICON.ColWidth(10) = 1000
CONSVARIAS.MATRICON.CellForeColor = RGB(255, 255, 255)
CONSVARIAS.MATRICON.CellBackColor = RGB(0, 0, 150)
CONSVARIAS.MATRICON.Text = "TELEFONO"
CONSVARIAS.MATRICON.Col = 11
CONSVARIAS.MATRICON.ColWidth(11) = 700
CONSVARIAS.MATRICON.CellForeColor = RGB(255, 255, 255)
CONSVARIAS.MATRICON.CellBackColor = RGB(0, 0, 150)
CONSVARIAS.MATRICON.Text = "A_INGR."
CONSVARIAS.MATRICON.Col = 12
CONSVARIAS.MATRICON.ColWidth(12) = 1200
CONSVARIAS.MATRICON.CellForeColor = RGB(255, 255, 255)
CONSVARIAS.MATRICON.CellBackColor = RGB(0, 0, 150)
CONSVARIAS.MATRICON.Text = "GRADO"
CONSVARIAS.MATRICON.Col = 13
CONSVARIAS.MATRICON.ColWidth(13) = 1600
CONSVARIAS.MATRICON.CellForeColor = RGB(255, 255, 255)
CONSVARIAS.MATRICON.CellBackColor = RGB(0, 0, 150)
CONSVARIAS.MATRICON.Text = "GRUPO"
NAR = FreeFile
cona = 2
J = 0
Screen.MousePointer = 11
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
While Not EOF(NAR)
    J = J + 1
    Get #NAR, J, alumno
    If (RTrim(alumno.n_carnet) <> "") And (RTrim(alumno.sexo) = RTrim(Combo1.Text)) Then
        NAR = FreeFile
        Open Ruta & "quegru.edu" For Random As #NAR Len = Len(aluper)
        Get #NAR, J, aluper
        Close #NAR
        CONSVARIAS.MATRICON.Rows = cona
        CONSVARIAS.MATRICON.TextMatrix((CONSVARIAS.MATRICON.Rows - 1), 0) = alumno.n_carnet
        CONSVARIAS.MATRICON.TextMatrix((CONSVARIAS.MATRICON.Rows - 1), 1) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
        CONSVARIAS.MATRICON.TextMatrix((CONSVARIAS.MATRICON.Rows - 1), 2) = alumno.n_matricula
        CONSVARIAS.MATRICON.TextMatrix((CONSVARIAS.MATRICON.Rows - 1), 3) = RTrim(alumno.f_nacimiento)
        dd = Val(Left(alumno.f_nacimiento, 2))
        mm2 = Right(alumno.f_nacimiento, 7)
        mm = Val(Left(mm2, 2))
        aaaa = Val(Right(alumno.f_nacimiento, 4))
        aaaa = Year(Date) - aaaa
        If mm > Month(Date) Then
            aaaa = aaaa - 1
        End If
        If mm = Month(Date) Then
            If dd > Day(Date) Then
                aaaa = aaaa - 1
            End If
        End If
        CONSVARIAS.MATRICON.TextMatrix((CONSVARIAS.MATRICON.Rows - 1), 4) = aaaa
        CONSVARIAS.MATRICON.TextMatrix((CONSVARIAS.MATRICON.Rows - 1), 5) = RTrim(alumno.rh)
        CONSVARIAS.MATRICON.TextMatrix((CONSVARIAS.MATRICON.Rows - 1), 6) = RTrim(alumno.sexo)
        CONSVARIAS.MATRICON.TextMatrix((CONSVARIAS.MATRICON.Rows - 1), 7) = RTrim(alumno.documento)
        CONSVARIAS.MATRICON.TextMatrix((CONSVARIAS.MATRICON.Rows - 1), 8) = RTrim(alumno.acudiente)
        CONSVARIAS.MATRICON.TextMatrix((CONSVARIAS.MATRICON.Rows - 1), 9) = RTrim(alumno.direccion)
        CONSVARIAS.MATRICON.TextMatrix((CONSVARIAS.MATRICON.Rows - 1), 10) = RTrim(alumno.tel_acu)
        CONSVARIAS.MATRICON.TextMatrix((CONSVARIAS.MATRICON.Rows - 1), 11) = RTrim(alumno.año_ingre)
        CONSVARIAS.MATRICON.TextMatrix((CONSVARIAS.MATRICON.Rows - 1), 12) = RTrim(alumno.grado)
        CONSVARIAS.MATRICON.TextMatrix((CONSVARIAS.MATRICON.Rows - 1), 13) = RTrim(aluper.grupo)
        cona = cona + 1
        NAR = NAR - 1
    End If
Wend
Close #NAR
If cona = 2 Then
    MsgBox "NO SE ENCONTRARON REGISTROS", 64
    Screen.MousePointer = 0
    Exit Sub
End If
CONSVARIAS.Caption = "Consultas opcionales - [total registros encontrados = " & (CONSVARIAS.MATRICON.Rows - 1) & "]"
CONSVARIAS.Frame1.Caption = "CONSULTA POR SEXO (SEXO = " & Combo1.Text & ")"
If CONSVARIAS.MATRICON.Rows > 1 Then
    CONSVARIAS.MATRICON.FixedRows = 1
    CONSVARIAS.MATRICON.FixedCols = 2
End If
Screen.MousePointer = 0
Unload Me
CONSVARIAS.Show
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Consulta de alumnos por sexo."
End Sub

Private Sub Form_Load()
If Dir(Ruta & "prinalu.edu") = "" Then
    Command1.Enabled = False
End If
End Sub
