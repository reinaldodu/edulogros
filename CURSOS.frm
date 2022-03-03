VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CURSOS 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear grupo"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8895
   Icon            =   "CURSOS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "ELIMINAR GRUPO"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Frame Frame7 
      Caption         =   "Borrar estudiante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   26
      Top             =   5160
      Width           =   2055
      Begin VB.CommandButton Command7 
         Caption         =   "Ok"
         Height          =   320
         Left            =   1440
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   840
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "CODIGO:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&GUARDAR"
      Height          =   740
      Left            =   5520
      Picture         =   "CURSOS.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "GUARDAR EL GRUPO EXISTENTE"
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H00800000&
      Height          =   3375
      Left            =   2640
      TabIndex        =   16
      Top             =   3120
      Width           =   6015
      Begin VB.CommandButton Command1 
         Caption         =   "Seleccionar estudiantes"
         Height          =   495
         Left            =   360
         TabIndex        =   29
         Top             =   2520
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&NUEVO"
         Height          =   735
         Left            =   4560
         Picture         =   "CURSOS.frx":2E0C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "CREAR NUEVO GRUPO"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Frame Frame6 
         Caption         =   "3. Director de grupo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2880
         TabIndex        =   23
         Top             =   1320
         Width           =   3015
         Begin VB.CommandButton Command5 
            Caption         =   "Ok"
            Height          =   330
            Left            =   2280
            TabIndex        =   4
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H00800000&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   960
            TabIndex        =   3
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "NUMERO:"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   765
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "2. Nombre del grupo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2880
         TabIndex        =   18
         Top             =   240
         Width           =   3015
         Begin VB.CommandButton Command2 
            Caption         =   "Ok"
            Height          =   320
            Left            =   2400
            TabIndex        =   2
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00800000&
            ForeColor       =   &H00FFFFFF&
            Height          =   320
            Left            =   960
            TabIndex        =   1
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "NOMBRE:"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   480
            Width           =   750
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "4. Agregar estudiantes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         ItemData        =   "CURSOS.frx":324E
         Left            =   1200
         List            =   "CURSOS.frx":327C
         TabIndex        =   0
         Text            =   "PREJARDIN"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         ItemData        =   "CURSOS.frx":32FC
         Left            =   1200
         List            =   "CURSOS.frx":330C
         TabIndex        =   10
         Text            =   "UNICA"
         Top             =   840
         Width           =   1335
      End
      Begin VB.Frame Frame5 
         Caption         =   "1. Jornada y grado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   2535
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "GRADO:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   960
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "JORNADA:"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   480
            Width           =   810
         End
      End
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00800000&
      Height          =   2535
      Left            =   2640
      TabIndex        =   12
      Top             =   480
      Width           =   6015
      Begin VB.TextBox Text1 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2040
         Width           =   495
      End
      Begin MSFlexGridLib.MSFlexGrid MATI3 
         Height          =   1815
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   3201
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         BackColorBkg    =   -2147483633
         GridColor       =   12582912
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   2640
         TabIndex        =   28
         Top             =   2160
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL ESTUDIANTES..."
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   1845
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   240
      Picture         =   "CURSOS.frx":332D
      ScaleHeight     =   4635
      ScaleWidth      =   2235
      TabIndex        =   11
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2760
      TabIndex        =   25
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "CURSOS"
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
    Text2.SetFocus
End If
End Sub

Private Sub Command1_Click()
If Label7.Caption = "" Then
    MsgBox "NO HA ASIGNADO DIRECTOR DE GRUPO", 32, "ADVERTENCIA"
    Text4.SetFocus
    Exit Sub
End If
NAR = FreeFile
Open Ruta & "cont.edu" For Input As #NAR
Input #NAR, I
Close #NAR
J = 0
Open Ruta & "quegru.edu" For Random As #NAR Len = Len(aluper)
For h = 1 To (I - 1)
Get #NAR, h, aluper
If (RTrim(aluper.grupo) = "PENDIENTE") Or (RTrim(aluper.grupo) = "SIN GRUPO") Then
    NAR = FreeFile
    Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
    Get #NAR, h, alumno
    If (RTrim(alumno.nombres) <> "") And (RTrim(alumno.apellidos) <> "") And (RTrim(alumno.grado) = J5) Then
        Est_Grado.List1.AddItem RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " - " & Format(alumno.n_carnet, "0000")
        J = J + 1
    End If
    Close #NAR
    NAR = NAR - 1
End If
Next h
Close #NAR
If J <> 0 Then
    Est_Grado.Caption = "Estudiantes del grado " & J5 & " - Total..." & J
    Est_Grado.Frame1.Caption = "(Presione <Ctrl> para seleccionar estudiantes o <Shift> para seleccionar en bloque)"
    Est_Grado.Show 1
Else
    MsgBox "NO HAY ESTUDIANTES DISPONIBLES PARA ESTE GRADO", 32, "CREAR GRUPO"
End If
End Sub

Private Sub Command4_Click()
If Val(Text1.Text) <> 0 Then
    Call Command6_Click
End If
MATI3.Rows = 1
CONT = 1
Frame1.Caption = ""
Frame2.Caption = ""
Label7.Caption = ""
Text2.Text = ""
Text4.Text = ""
Text1.Text = "0"
Command2.Enabled = True
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Crear o eliminar grupos."
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Command2.Enabled = True Then
    If KeyAscii = 13 Then
        Call Command2_Click
    End If
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command5_Click
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
Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command7_Click
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

Private Sub Command2_Click()
'Dim alumno As maestroalum
If RTrim(Text2.Text) = "" Then
    MsgBox "ESCRIBA UN NOMBRE DE GRUPO", 32, "CREAR GRUPO"
    Text2.Text = ""
    Text2.SetFocus
    Exit Sub
End If
Frame1.Caption = "JORNADA: " & Combo1.Text & "  GRADO: " & Combo2.Text
J4 = RTrim(Combo1.Text)
J5 = RTrim(Combo2.Text)
Text2.Text = Format(Text2.Text, ">")
    If J4 = "UNICA" Then
        ar = "1"
    End If
    If J4 = "MAÑANA" Then
        ar = "2"
    End If
    If J4 = "TARDE" Then
        ar = "3"
    End If
    If J4 = "NOCHE" Then
        ar = "4"
    End If
    If Dir(Ruta & RTrim(ar & Text2.Text) & ".gru") <> "" Then
        MsgBox "YA SE CREO ESTE GRUPO", 16
        Text2.Text = ""
        Text2.SetFocus
        Exit Sub
    End If
    PARCHI = RTrim(ar & Text2.Text)
    Frame2.Caption = "GRUPO A CREAR: " & PARCHI
    Text4.SetFocus
End Sub

Private Sub Command5_Click()
'Dim profe As maestropro
If Frame2.Caption = "" Then
    MsgBox "NO LE HA DADO UN NOMBRE AL GRUPO", 32, "ADVERTENCIA"
    Text2.SetFocus
    Exit Sub
End If
If Text4.Text = "" Then
    MsgBox "ESCRIBA EL NUMERO DEL DIRECTOR", 48, "DIRECTOR"
    Text4.SetFocus
    Exit Sub
End If
NAR = FreeFile
Open Ruta & "contpro.edu" For Input As #NAR
Input #NAR, r
Close #NAR
If ((Val(Text4.Text) > r - 1) Or (Val(Text4.Text) < 1)) Then
    MsgBox "PROFESOR NO EXISTE", 64, "ADVERTENCIA"
    Text4.Text = ""
    Text4.SetFocus
    Exit Sub
End If
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Get #NAR, Val(Text4.Text), profe
Close #NAR
If RTrim(profe.nombres) = "" Then
    MsgBox "REGISTRO NO EXISTE", 16, "DIRECTOR DE GRUPO"
    Text4.SetFocus
    Exit Sub
End If
dire2 = Val(Text4.Text)
Label7.Caption = "DIRECTOR: " & RTrim(profe.nombres) & " " & profe.apellidos
End Sub

Private Sub Command6_Click()
'Dim alugru As grupoalu
'Dim icur As inforcur
'Dim aluper As pertgrup
If CONT = 1 Then
    MsgBox "NO HAY INFORMACION PARA GUARDAR", 32, "GUARDAR"
    Text2.SetFocus
    Exit Sub
End If
RESP = MsgBox("DESEA GUARDAR EL GRUPO " & PARCHI & "?", vbYesNo + vbQuestion + vbDefaultButton1, "GUARDAR GRUPO")
If RESP = vbYes Then
    MATI3.Col = 2
    MATI3.Sort = 5
    NAR = FreeFile
    Open Ruta & PARCHI & ".gru" For Random As #NAR Len = Len(alugru)
    NAR = FreeFile
    Open Ruta & "quegru.edu" For Random As #NAR Len = Len(aluper)
    For GH = 1 To CONT - 1
        alugru.num_carnet = MATI3.TextMatrix(GH, 1)
        Put #(NAR - 1), GH, alugru
        aluper.grupo = RTrim(PARCHI)
        Put #NAR, Val(alugru.num_carnet), aluper
    Next GH
    Close #NAR
    Close #(NAR - 1)
    icur.nom = PARCHI
    icur.jornada = J4
    icur.grado = J5
    icur.director = dire2
    Open Ruta & "infcur.edu" For Append As #NAR
    Write #NAR, icur.nom, icur.jornada, icur.grado, icur.director
    Close #NAR
    MATI3.Rows = 1
    CONT = 1
    Frame1.Caption = ""
    Frame2.Caption = ""
    Label7.Caption = ""
    Text2.Text = ""
    Text4.Text = ""
    Text1.Text = "0"
    Command2.Enabled = True
End If
End Sub

Private Sub Command7_Click()
If Text5.Text = "" Then
    MsgBox "ESCRIBA EL CODIGO A ELIMINAR", 48, "ELIMINAR"
    Text5.SetFocus
    Exit Sub
End If
If ((Text5.Text > CONT - 1) Or (Text5.Text < 1)) Then
    MsgBox "NO EXISTE EL CÓDIGO", 32, "ELIMINAR CÓDIGO"
    Text5.Text = ""
    Text5.SetFocus
    Exit Sub
End If
If Val(Text1.Text) <> 1 Then
    MATI3.RemoveItem Val(Text5.Text)
    CONT = CONT - 1
    For TT = 1 To (CONT - 1)
        MATI3.TextMatrix(TT, 0) = TT
    Next TT
Else
    MATI3.Rows = 1
    CONT = CONT - 1
End If
Text1.Text = Val(Text1.Text) - 1
End Sub

Private Sub Command8_Click()
I = 0
PASSW.Show 1
If I = 1 Then
ELIMNAR_CURSO.Show 1
End If
End Sub

Private Sub Form_Load()
Text2.MaxLength = 20
Text4.MaxLength = 3
Text5.MaxLength = 2
MATI3.Row = 0
MATI3.Col = 0
MATI3.ColWidth(0) = 450
MATI3.CellForeColor = RGB(255, 255, 255)
MATI3.CellBackColor = RGB(0, 0, 150)
MATI3.Text = "CÓD"
MATI3.Col = 1
MATI3.ColWidth(1) = 800
MATI3.CellForeColor = RGB(255, 255, 255)
MATI3.CellBackColor = RGB(0, 0, 150)
MATI3.Text = "CARNET"
'MATI3.Col = 2
'MATI3.ColWidth(2) = 1100
'MATI3.CellForeColor = RGB(255, 255, 255)
'MATI3.CellBackColor = RGB(0, 0, 150)
'MATI3.Text = "MATRICULA"
MATI3.Col = 2
MATI3.ColWidth(2) = 4000
MATI3.CellForeColor = RGB(255, 255, 255)
MATI3.CellBackColor = RGB(0, 0, 150)
MATI3.Text = "APELLIDOS Y NOMBRES"
'MATI3.Col = 3
'MATI3.ColWidth(3) = 1900
'MATI3.CellForeColor = RGB(255, 255, 255)
'MATI3.CellBackColor = RGB(0, 0, 150)
'MATI3.Text = "NOMBRES"
'MATI3.Col = 5
'MATI3.ColWidth(5) = 2800
'MATI3.CellForeColor = RGB(255, 255, 255)
'MATI3.CellBackColor = RGB(0, 0, 150)
'MATI3.Text = "ACUDIENTE"
'MATI3.Col = 6
'MATI3.ColWidth(6) = 1200
'MATI3.CellForeColor = RGB(255, 255, 255)
'MATI3.CellBackColor = RGB(0, 0, 150)
'MATI3.Text = "TELEFONO"
Text1.Text = "0"
CONT = 1
End Sub
