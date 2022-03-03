VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form planeacion_semanal 
   Caption         =   "Planeación semanal"
   ClientHeight    =   7740
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13950
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   13950
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   240
      TabIndex        =   10
      Top             =   6480
      Width           =   13575
      Begin VB.CommandButton Command7 
         Caption         =   "Salir"
         Height          =   435
         Left            =   11760
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Eliminar"
         Height          =   435
         Left            =   2280
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Agregar"
         Height          =   435
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   10080
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   45
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   13575
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   11760
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   8040
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   3375
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   3015
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "planeacion_semanal.frx":0000
         Left            =   960
         List            =   "planeacion_semanal.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Materia:"
         Height          =   195
         Left            =   7320
         TabIndex        =   7
         Top             =   480
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Grupo:"
         Height          =   195
         Left            =   3360
         TabIndex        =   5
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodo:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   13575
      Begin MSFlexGridLib.MSFlexGrid MTPlan 
         Height          =   4695
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   8281
         _Version        =   393216
         Rows            =   1
         Cols            =   6
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "planeacion_semanal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Pln1 As Integer, Pln2 As Integer, Pln3 As Integer, Pln4 As Integer, ContRow As Integer
MTPlan.Rows = 1
Label4.Caption = ""

Command2.Enabled = False
Command4.Enabled = False
'Command6.Enabled = False
If Dir(Ruta & Combo2.Text & ".gru") = "" Then
    MsgBox "GRUPO INCORRECTO", 48
    Exit Sub
End If
Screen.MousePointer = 11
NAR = FreeFile
TN = 0
Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
While Not EOF(NAR)
    TN = TN + 1
    Get #NAR, TN, mate
    If RTrim(mate.nom) = Combo3.Text Then
        que = mate.num
    End If
Wend
Close #NAR
ret = 0
Open Ruta & Combo2.Text & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
    ret = ret + 1
    Get #NAR, ret, alugru
Wend
Close #NAR
NAR = FreeFile
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
    Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
    If RTrim(icur.nom) = RTrim(Combo2.Text) Then
        RE22 = RTrim(icur.grado)
        JOJI = RTrim(icur.jornada)
        GoTo ALTU33
    End If
Wend
ALTU33:
Close #NAR
pio = 0
cona = 0
Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
While Not EOF(NAR)
    cona = cona + 1
    Get #NAR, cona, argra
    If ((RTrim(argra.grado) = RE22) And (RTrim(argra.nom_grup) = Combo2.Text) And (argra.num_area = que)) Then
        NAR = FreeFile
        Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
        Get #NAR, (argra.num_pro), profe
        Close #NAR
        PRO = RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
        pio = 1
        NAR = NAR - 1
    End If
Wend
Close #NAR
If pio = 0 Then
MsgBox "NO SE HA CREADO EL AREA " & Combo3.Text & " PARA ESTE GRUPO O NO LE CORRESPONDE", 64, "ADVERTENCIA"
    Combo3.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If
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
fl = "1"
ser = Left(RE22, 3)

Frame1.Caption = "PLANEACIÓN SEMANAL - GRUPO:" & Combo2.Text & " / MATERIA:" & Combo3.Text
Label4.Caption = Combo2.Text

'***Verificar cantidad de registros de cada uno de los archivos***

'Archivo de la planeación semanal
h = 0
Open Ruta & Label4.Caption & que & lw & ".pln" For Random As #NAR Len = Len(semanal_planeacion)
While Not EOF(NAR)
    h = h + 1
    Get #NAR, h, semanal_planeacion
Wend
Close #NAR

'Archivo de los ejes temáticos
Pln1 = 0
Open Ruta & fl & ser & que & lw & ".eje" For Random As #NAR Len = Len(semanal_ejetematico)
While Not EOF(NAR)
    Pln1 = Pln1 + 1
    Get #NAR, Pln1, semanal_ejetematico
Wend
Close #NAR

'Archivo de los contenidos
Pln2 = 0
Open Ruta & fl & ser & que & lw & ".ctd" For Random As #NAR Len = Len(semanal_contenidos)
While Not EOF(NAR)
    Pln2 = Pln2 + 1
    Get #NAR, Pln2, semanal_contenidos
Wend
Close #NAR

'Archivo de las competencias
Pln3 = 0
Open Ruta & fl & ser & que & lw & ".cpt" For Random As #NAR Len = Len(semanal_competencias)
While Not EOF(NAR)
    Pln3 = Pln3 + 1
    Get #NAR, Pln3, semanal_competencias
Wend
Close #NAR

'Archivo de los logros
Pln4 = 0
Open Ruta & fl & ser & que & lw & ".lgr" For Random As #NAR Len = Len(logru)
While Not EOF(NAR)
    Pln4 = Pln4 + 1
    Get #NAR, Pln4, logru
Wend
Close #NAR

ContRow = 1
For I = 1 To h - 1
    NAR = FreeFile
    Open Ruta & Label4.Caption & que & lw & ".pln" For Random As #NAR Len = Len(semanal_planeacion)
    Get #NAR, I, semanal_planeacion
    If Trim(semanal_planeacion.fecha) <> "" Then
        MTPlan.Rows = ContRow + 1
        MTPlan.TextMatrix(ContRow, 0) = Trim(semanal_planeacion.fecha)
        NAR = FreeFile
        If Val(semanal_planeacion.eje) < Pln1 Then
            Open Ruta & fl & ser & que & lw & ".eje" For Random As #NAR Len = Len(semanal_ejetematico)
            Get #NAR, Val(semanal_planeacion.eje), semanal_ejetematico
            Close #NAR
            MTPlan.TextMatrix(ContRow, 1) = Trim(semanal_ejetematico.txt_eje)
        End If
        ArrCont = Split(semanal_planeacion.contenidos, ",")
        For r = 0 To UBound(ArrCont) - 1
            If Val(ArrCont(r)) < Pln2 Then
                Open Ruta & fl & ser & que & lw & ".ctd" For Random As #NAR Len = Len(semanal_contenidos)
                Get #NAR, Val(ArrCont(r)), semanal_contenidos
                Close #NAR
                MTPlan.TextMatrix(ContRow, 2) = MTPlan.TextMatrix(ContRow, 2) & Trim(semanal_contenidos.txt_cont) & vbCrLf
            End If
        Next r
        If Val(semanal_planeacion.competencia) < Pln3 Then
            Open Ruta & fl & ser & que & lw & ".cpt" For Random As #NAR Len = Len(semanal_competencias)
            Get #NAR, Val(semanal_planeacion.competencia), semanal_competencias
            Close #NAR
            MTPlan.TextMatrix(ContRow, 3) = Trim(semanal_competencias.txt_comp)
        End If
        ArrCont2 = Split(semanal_planeacion.logros, ",")
        For r = 0 To UBound(ArrCont2) - 1
            If Val(ArrCont2(r)) < Pln4 Then
                Open Ruta & fl & ser & que & lw & ".lgr" For Random As #NAR Len = Len(logru)
                Get #NAR, Val(ArrCont2(r)), logru
                Close #NAR
                MTPlan.TextMatrix(ContRow, 4) = MTPlan.TextMatrix(ContRow, 4) & "(" & ArrCont2(r) & ") " & Trim(logru.observ) & vbCrLf
            End If
        Next r
        If Val(UBound(ArrCont)) < Val(UBound(ArrCont2)) Then
            MTPlan.RowHeight(ContRow) = 240 * (Val(UBound(ArrCont2)) + 2)
        Else
            MTPlan.RowHeight(ContRow) = 240 * (Val(UBound(ArrCont)) + 2)
        End If
        
        NAR = NAR - 1
        ArrFecha = Split(Trim(semanal_planeacion.fecha), "/")
        MTPlan.TextMatrix(ContRow, 5) = ArrFecha(2) & ArrFecha(1) & ArrFecha(0)
        ContRow = ContRow + 1
    End If
    Close #NAR
Next I
MTPlan.Col = 5
MTPlan.Sort = 3
Screen.MousePointer = 0

Command2.Enabled = True
'Command3.Enabled = True
Command4.Enabled = True
'Command5.Enabled = True
'Command6.Enabled = True
End Sub

Private Sub Command2_Click()
NewPlaneador.Show 1
End Sub

Private Sub Command4_Click()
If Val(MTPlan.Rows - 1) = 0 Then
    MsgBox "NO EXISTE INFORMACION PARA ELIMINAR", 64
    Exit Sub
End If
If Val(MTPlan.Rows - 1) = 1 Then
    MsgBox "No se puede Eliminar el último ítem de la planeación", 32, "Eliminar ítem"
    Exit Sub
End If
For I = 1 To MTPlan.Rows - 1
    Del_Planeador.List1.AddItem MTPlan.TextMatrix(I, 0)
Next I
Del_Planeador.Show 1
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Form_Load()
MTPlan.Row = 0
MTPlan.Col = 0
MTPlan.ColWidth(0) = 1000
MTPlan.CellForeColor = RGB(255, 255, 255)
MTPlan.CellBackColor = RGB(0, 0, 150)
MTPlan.Text = "FECHA"
MTPlan.Col = 1
MTPlan.ColWidth(1) = 3000
MTPlan.CellForeColor = RGB(255, 255, 255)
MTPlan.CellBackColor = RGB(0, 0, 150)
MTPlan.Text = "EJES"
MTPlan.Col = 2
MTPlan.ColWidth(2) = 3000
MTPlan.CellForeColor = RGB(255, 255, 255)
MTPlan.CellBackColor = RGB(0, 0, 150)
MTPlan.Text = "CONTENIDOS"
MTPlan.Col = 3
MTPlan.ColWidth(3) = 5000
MTPlan.CellForeColor = RGB(255, 255, 255)
MTPlan.CellBackColor = RGB(0, 0, 150)
MTPlan.Text = "COMPETENCIAS"
MTPlan.Col = 4
MTPlan.ColWidth(4) = 5000
MTPlan.CellForeColor = RGB(255, 255, 255)
MTPlan.CellBackColor = RGB(0, 0, 150)
MTPlan.Text = "LOGROS"
'Columna oculta para ordenar las fechas del planeador
MTPlan.Col = 5
MTPlan.ColWidth(5) = 0
planeacion_semanal.MTPlan.ColAlignment(0) = 0
planeacion_semanal.MTPlan.ColAlignment(1) = 0
planeacion_semanal.MTPlan.ColAlignment(2) = 0
planeacion_semanal.MTPlan.ColAlignment(3) = 0
planeacion_semanal.MTPlan.ColAlignment(4) = 0


Command2.Enabled = False
'Command3.Enabled = False
Command4.Enabled = False
'Command5.Enabled = False
'Command6.Enabled = False

If (Dir(Ruta & "infcur.edu") <> "") And (Dir(Ruta & "materia.edu") <> "") Then
    Command1.Enabled = True
    NAR = FreeFile
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
        Combo2.AddItem RTrim(icur.nom)
    Wend
    Close #NAR
    Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
    cona = 0
    While Not EOF(NAR)
        cona = cona + 1
        Get #NAR, cona, mate
    Wend
    Close #NAR
    Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
    For I = 1 To cona - 1
        Get #NAR, I, mate
        If RTrim(mate.nom) <> "" Then
            Combo3.AddItem RTrim(mate.nom)
        End If
    Next I
    Close #NAR
    Combo2.Text = Combo2.List(0)
    Combo3.Text = Combo3.List(0)
Else
    Command1.Enabled = False
End If
Combo1 = Combo1.List(0)
Combo2 = Combo2.List(0)
Combo3 = Combo3.List(0)

''If (Dir(Ruta & "materia.edu") <> "") And (Dir(Ruta & "areagra.edu") <> "") Then
''    Command1.Enabled = True
''    NAR = FreeFile
''    cona = 0
''    Open Ruta & "areagra.edu" For Random As #NAR Len = Len(argra)
''    While Not EOF(NAR)
''        cona = cona + 1
''        Get #NAR, cona, argra
''        If argra.num_pro = Val(MENUPROFE.LBLNumProfe.Caption) Then
''            VALI2 = False
''            For I = 0 To (Combo2.ListCount - 1)
''                If Combo2.List(I) = RTrim(argra.nom_grup) Then
''                    VALI2 = True
''                    Exit For
''                End If
''            Next I
''            If VALI2 = False Then
''                Combo2.AddItem RTrim(argra.nom_grup)
''            End If
''            NAR = FreeFile
''            Open Ruta & "materia.edu" For Random As #NAR Len = Len(mate)
''            Get #NAR, argra.num_area, mate
''            Close #NAR
''            NAR = NAR - 1
''            VALI2 = False
''            For I = 0 To (Combo3.ListCount - 1)
''                If Combo3.List(I) = RTrim(mate.nom) Then
''                    VALI2 = True
''                    Exit For
''                End If
''            Next I
''            If VALI2 = False Then
''                Combo3.AddItem RTrim(mate.nom)
''            End If
''        End If
''    Wend
''    Close #NAR
''    Combo1.Text = Combo1.List(0)
''    Combo2.Text = Combo2.List(0)
''    Combo3.Text = Combo3.List(0)
''        If (RTrim(Combo2.Text) = "") Or (RTrim(Combo3.Text) = "") Then
''        Command1.Enabled = False
''    End If
''Else
''    Command1.Enabled = False
''End If
End Sub
