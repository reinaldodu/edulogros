VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Import_CSV 
   Caption         =   "Importar archivo CSV"
   ClientHeight    =   5850
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Ayuda SQL"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Seleccionar archivo CSV"
      Height          =   615
      Left            =   3600
      TabIndex        =   3
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar datos"
      Height          =   615
      Left            =   8040
      TabIndex        =   2
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Importar archivo CSV"
      Height          =   615
      Left            =   5880
      TabIndex        =   1
      Top             =   5040
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MTCSV 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   8281
      _Version        =   393216
      Rows            =   3
      Cols            =   21
   End
End
Attribute VB_Name = "Import_CSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Dir(RutaCSV) <> "" Then
    Screen.MousePointer = 11
    Open RutaCSV For Input As #NAR
    CONT = 1
    While Not EOF(NAR)
        Input #NAR, VerColum
        CONT = CONT + 1
        MTCSV.Rows = CONT
        ArrayColumn = Split(VerColum, ";")
        X = UBound(ArrayColumn)
        If X <> 19 Then
            MsgBox "El registro No." & CONT - 1 & " no puede ser importado, revisélo e intente nuevamente.  Es probable que algún campo contenga un separador CSV como coma (,) o punto y coma (;).  Verifique también que el archivo CSV cumple con la estructura de datos necesaria.", 16
            Close #NAR
            Screen.MousePointer = 0
            Exit Sub
        End If
        MTCSV.TextMatrix(CONT - 1, 0) = CONT - 1
        For I = 0 To 19
            MTCSV.TextMatrix(CONT - 1, I + 1) = ArrayColumn(I)
        Next I
    Wend
    Close #NAR
    Screen.MousePointer = 0
    Command2.Enabled = True
End If
End Sub

Private Sub Command2_Click()
RESP = MsgBox("Desea guardar en el sistema los registros que aparecen en la tabla?", vbYesNo + vbQuestion + vbDefaultButton2, "Guardar")
If RESP = vbYes Then
    Screen.MousePointer = 11
    NAR = FreeFile
    Open Ruta & "cont.edu" For Input As #NAR
    Input #NAR, I
    Close #NAR
    NAR = FreeFile
    For X = 1 To MTCSV.Rows - 1
        alumno.nombres = MTCSV.TextMatrix(X, 1)
        alumno.apellidos = MTCSV.TextMatrix(X, 2)
        AdiCampo.Tel_casa = MTCSV.TextMatrix(X, 3)
        alumno.direccion = MTCSV.TextMatrix(X, 4)
        alumno.f_nacimiento = MTCSV.TextMatrix(X, 6)
        alumno.rh = MTCSV.TextMatrix(X, 7)
        alumno.sexo = MTCSV.TextMatrix(X, 8)
        alumno.documento = MTCSV.TextMatrix(X, 9)
        AdiCampo.salud = MTCSV.TextMatrix(X, 10)
        alumno.grado = MTCSV.TextMatrix(X, 11)
        alumno.padre = MTCSV.TextMatrix(X, 12) & " " & MTCSV.TextMatrix(X, 13)
        alumno.tel_pa = MTCSV.TextMatrix(X, 14)
        alumno.madre = MTCSV.TextMatrix(X, 15) & " " & MTCSV.TextMatrix(X, 16)
        alumno.tel_ma = MTCSV.TextMatrix(X, 17)
        alumno.acudiente = MTCSV.TextMatrix(X, 18) & " " & MTCSV.TextMatrix(X, 19)
        alumno.tel_acu = MTCSV.TextMatrix(X, 20)
        alumno.jornada = "UNICA"
        alumno.n_carnet = (I - 1) + X
        alumno.año_ingre = Year(Date)
        alumno.n_matricula = (I - 1) + X
        AdiCampo.otras = ""
        AdiCampo.email = MTCSV.TextMatrix(X, 5)
        'Guardar información de estudiantes
        Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
        Put #NAR, (I - 1) + X, alumno
        Close #NAR
        ' Guardar información estudiantes pendientes de grupo
        aluper.grupo = "PENDIENTE"
        Open Ruta & "quegru.edu" For Random As #NAR Len = Len(aluper)
        Put #NAR, (I - 1) + X, aluper
        Close #NAR
        ' Guardar información adicional del estudiante
        Open Ruta & "mascampo.edu" For Random As #NAR Len = Len(AdiCampo)
        Put #NAR, (I - 1) + X, AdiCampo
        Close #NAR
        'Habilitar espacio para la información de pensiones
        If Dir(Ruta & "pensi.edu") <> "" Then
            For J = 1 To 12
                pens(J) = 0
            Next J
            Open Ruta & "pensi.edu" For Random As #NAR Len = 96
            Put #NAR, (I - 1) + X, pens
            Close #NAR
        End If
        'Habilitar espacio para la información del historial de colegios
        If Dir(Ruta & "infcol.edu") <> "" Then
            For J = 1 To 14
                newmatri.nombre(J) = ""
                newmatri.grado(J) = ""
                newmatri.año(J) = ""
                newmatri.ciudad(J) = ""
            Next J
            Open Ruta & "infcol.edu" For Random As #NAR Len = Len(newmatri)
            Put #NAR, (I - 1) + X, newmatri
            Close #NAR
        End If
    Next X
    'Guargar contador de estudiantes
    Open Ruta & "cont.edu" For Output As #NAR
    Print #NAR, (I + X - 1)
    Close #NAR
    Screen.MousePointer = 0
End If
End Sub

Private Sub Command3_Click()
I = 0
PASSW.Show 1
If I = 1 Then
    DriveCSV.Show
End If
End Sub

Private Sub Command4_Click()
AyudaSQL.Show
End Sub

Private Sub Form_Load()
MTCSV.Row = 0
MTCSV.Col = 0
MTCSV.ColWidth(0) = 500
MTCSV.CellForeColor = RGB(255, 255, 255)
MTCSV.CellBackColor = RGB(0, 0, 150)
MTCSV.Text = "No."
MTCSV.Col = 1
MTCSV.ColWidth(1) = 2500
MTCSV.CellForeColor = RGB(255, 255, 255)
MTCSV.CellBackColor = RGB(0, 0, 150)
MTCSV.Text = "NOMBRES"
MTCSV.Col = 2
MTCSV.ColWidth(2) = 2500
MTCSV.CellForeColor = RGB(255, 255, 255)
MTCSV.CellBackColor = RGB(0, 0, 150)
MTCSV.Text = "APELLIDOS"
MTCSV.Col = 3
MTCSV.ColWidth(3) = 1500
MTCSV.CellForeColor = RGB(255, 255, 255)
MTCSV.CellBackColor = RGB(0, 0, 150)
MTCSV.Text = "TELEFONO"
MTCSV.Col = 4
MTCSV.ColWidth(4) = 3500
MTCSV.CellForeColor = RGB(255, 255, 255)
MTCSV.CellBackColor = RGB(0, 0, 150)
MTCSV.Text = "DIRECCIÓN"
MTCSV.Col = 5
MTCSV.ColWidth(5) = 2000
MTCSV.CellForeColor = RGB(255, 255, 255)
MTCSV.CellBackColor = RGB(0, 0, 150)
MTCSV.Text = "EMAIL"

MTCSV.Col = 6
MTCSV.ColWidth(6) = 1500
MTCSV.CellForeColor = RGB(255, 255, 255)
MTCSV.CellBackColor = RGB(0, 0, 150)
MTCSV.Text = "FNACIMIENTO"
MTCSV.Col = 7
MTCSV.ColWidth(7) = 1000
MTCSV.CellForeColor = RGB(255, 255, 255)
MTCSV.CellBackColor = RGB(0, 0, 150)
MTCSV.Text = "RH"
MTCSV.Col = 8
MTCSV.ColWidth(8) = 1000
MTCSV.CellForeColor = RGB(255, 255, 255)
MTCSV.CellBackColor = RGB(0, 0, 150)
MTCSV.Text = "SEXO"
MTCSV.Col = 9
MTCSV.ColWidth(9) = 1500
MTCSV.CellForeColor = RGB(255, 255, 255)
MTCSV.CellBackColor = RGB(0, 0, 150)
MTCSV.Text = "DOCUMENTO"
MTCSV.Col = 10
MTCSV.ColWidth(10) = 2000
MTCSV.CellForeColor = RGB(255, 255, 255)
MTCSV.CellBackColor = RGB(0, 0, 150)
MTCSV.Text = "EPS"
MTCSV.Col = 11
MTCSV.ColWidth(11) = 1500
MTCSV.CellForeColor = RGB(255, 255, 255)
MTCSV.CellBackColor = RGB(0, 0, 150)
MTCSV.Text = "GRADO"
MTCSV.Col = 12
MTCSV.ColWidth(12) = 2500
MTCSV.CellForeColor = RGB(255, 255, 255)
MTCSV.CellBackColor = RGB(0, 0, 150)
MTCSV.Text = "NOMBRES DEL PADRE"
MTCSV.Col = 13
MTCSV.ColWidth(13) = 2500
MTCSV.CellForeColor = RGB(255, 255, 255)
MTCSV.CellBackColor = RGB(0, 0, 150)
MTCSV.Text = "APELLIDOS DEL PADRE"
MTCSV.Col = 14
MTCSV.ColWidth(14) = 2000
MTCSV.CellForeColor = RGB(255, 255, 255)
MTCSV.CellBackColor = RGB(0, 0, 150)
MTCSV.Text = "TELEFONO DEL PADRE"

MTCSV.Col = 15
MTCSV.ColWidth(15) = 2500
MTCSV.CellForeColor = RGB(255, 255, 255)
MTCSV.CellBackColor = RGB(0, 0, 150)
MTCSV.Text = "NOMBRES DE LA MADRE"
MTCSV.Col = 16
MTCSV.ColWidth(16) = 2500
MTCSV.CellForeColor = RGB(255, 255, 255)
MTCSV.CellBackColor = RGB(0, 0, 150)
MTCSV.Text = "APELLIDOS DE LA MADRE"
MTCSV.Col = 17
MTCSV.ColWidth(17) = 2500
MTCSV.CellForeColor = RGB(255, 255, 255)
MTCSV.CellBackColor = RGB(0, 0, 150)
MTCSV.Text = "TELEFONO DE LA MADRE"

MTCSV.Col = 18
MTCSV.ColWidth(18) = 2500
MTCSV.CellForeColor = RGB(255, 255, 255)
MTCSV.CellBackColor = RGB(0, 0, 150)
MTCSV.Text = "NOMBRES DEL ACUDIENTE"
MTCSV.Col = 19
MTCSV.ColWidth(19) = 2500
MTCSV.CellForeColor = RGB(255, 255, 255)
MTCSV.CellBackColor = RGB(0, 0, 150)
MTCSV.Text = "APELLIDOS DEL ACUDIENTE"
MTCSV.Col = 20
MTCSV.ColWidth(20) = 2500
MTCSV.CellForeColor = RGB(255, 255, 255)
MTCSV.CellBackColor = RGB(0, 0, 150)
MTCSV.Text = "TELEFONO DEL ACUDIENTE"

Command1.Enabled = False
Command2.Enabled = False
End Sub
