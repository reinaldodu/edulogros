VERSION 5.00
Begin VB.Form CVRANGO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta por rango de carnets"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "CVRANGO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.VScrollBar VScroll2 
         Height          =   750
         Left            =   3000
         TabIndex        =   6
         Top             =   240
         Width           =   200
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   750
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   200
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   1800
         TabIndex        =   4
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   510
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "CVRANGO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Screen.MousePointer = 11
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
For J = Val(Text1.Text) To Val(Text2.Text)
    Get #NAR, J, alumno
    If RTrim(alumno.n_carnet) <> "" Then
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
Next J
Close #NAR
CONSVARIAS.Caption = "Consultas opcionales - [total registros encontrados = " & (CONSVARIAS.MATRICON.Rows - 1) & "]"
CONSVARIAS.Frame1.Caption = "CONSULTA POR RANGO DE CARNETS (DESDE EL No." & Text1.Text & " HASTA EL No." & Text2.Text & ")"
If CONSVARIAS.MATRICON.Rows > 1 Then
    CONSVARIAS.MATRICON.FixedRows = 1
    CONSVARIAS.MATRICON.FixedCols = 2
End If
Screen.MousePointer = 0
Unload Me
CONSVARIAS.Show
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Consulta de alumnos por rango de carnets."
End Sub

Private Sub Form_Load()
NAR = FreeFile
Open Ruta & "cont.edu" For Input As #NAR
Input #NAR, I
Close #NAR
If (I - 1) = 0 Then
    VScroll1.Enabled = False
    VScroll2.Enabled = False
    Command1.Enabled = False
    Exit Sub
End If
VScroll1.Min = 1
VScroll1.Max = I - 1
VScroll2.Min = 1
VScroll2.Max = I - 1
VScroll1.LargeChange = 10
VScroll2.LargeChange = 10
VScroll1.SmallChange = 1
VScroll2.SmallChange = 1
End Sub

Private Sub VScroll1_Change()
    Text1.Text = VScroll1.Value
    If Val(Text1.Text) > Val(Text2.Text) Then
        Text2.Text = Text1.Text
        VScroll2.Value = VScroll1.Value
    End If
End Sub

Private Sub VScroll2_Change()
    Text2.Text = VScroll2.Value
    If Val(Text2.Text) < Val(Text1.Text) Then
        Text1.Text = Text2.Text
        VScroll1.Value = VScroll2.Value
    End If
End Sub
