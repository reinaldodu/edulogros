VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CONS_GRUP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar grupo"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7590
   Icon            =   "CONS_GRUP.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "&CAMBIAR DE DIRECTOR"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&IMPRIMIR GRUPO"
      Height          =   615
      Left            =   2520
      TabIndex        =   4
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      ForeColor       =   &H00C00000&
      Height          =   2055
      Left            =   4560
      TabIndex        =   10
      Top             =   4440
      Width           =   2895
      Begin VB.Frame Frame5 
         Caption         =   "Borrar estudiante"
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   2655
         Begin VB.CommandButton Command3 
            Caption         =   "Ok"
            Height          =   320
            Left            =   1680
            TabIndex        =   3
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00800000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   320
            Left            =   960
            TabIndex        =   2
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Left            =   360
            TabIndex        =   13
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Agregar estudiantes"
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   2655
         Begin VB.CommandButton Command8 
            Caption         =   "Seleccionar estudiantes"
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   2175
         End
      End
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H00C00000&
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   4215
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   1920
         TabIndex        =   0
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&ACEPTAR"
         Height          =   320
         Left            =   1440
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "NOMBRE DEL GRUPO:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1740
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7335
      Begin MSFlexGridLib.MSFlexGrid MATI9 
         Height          =   3375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5953
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         BackColorBkg    =   -2147483633
         GridColor       =   12582912
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total estudiantes..."
         Height          =   195
         Left            =   5160
         TabIndex        =   17
         Top             =   3720
         Width           =   1350
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "CONS_GRUP"
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
'Dim profe As maestropro
'Dim icur As inforcur
'Dim alugru As grupoalu
If Dir(Ruta & Combo1.Text & ".gru") = "" Then
    YO = 0
    MsgBox "GRUPO INCORRECTO", 48
    Exit Sub
End If
Screen.MousePointer = 11
MATI9.Rows = 1
Label7.Caption = ""
Frame1.Caption = ""
YO = 1
NAR = FreeFile
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
If RTrim(icur.nom) = Combo1.Text Then
J2 = RTrim(icur.grado)
J1 = RTrim(icur.jornada)
J3 = icur.director
GoTo ALTU86
End If
Wend
ALTU86:
Close #NAR
Open Ruta & "prinpro.edu" For Random As #NAR Len = Len(profe)
Get #NAR, J3, profe
Close #NAR
Label7.Caption = "Directora: " & RTrim(profe.nombres) & " " & RTrim(profe.apellidos)
Frame1.Caption = "JORNADA: " & J1 & "   GRADO: " & J2 & "   GRUPO: " & Combo1.Text
leo = 0
Open Ruta & Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
While Not EOF(NAR)
leo = leo + 1
Get #NAR, leo, alugru
Wend
Close #NAR
Open Ruta & Combo1.Text & ".gru" For Random As #NAR Len = Len(alugru)
NAR = FreeFile
Open Ruta & "prinalu.edu" For Random As #NAR Len = Len(alumno)
For TN = 1 To leo - 1
Get #(NAR - 1), TN, alugru
Get #NAR, (Val(alugru.num_carnet)), alumno
MATI9.Rows = YO + 1
MATI9.TextMatrix(YO, 0) = TN
MATI9.TextMatrix(YO, 1) = alumno.n_carnet
'MATI9.TextMatrix(YO, 2) = alumno.n_matricula
MATI9.TextMatrix(YO, 2) = RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres)
'MATI9.TextMatrix(YO, 4) = RTrim(alumno.nombres)
'MATI9.TextMatrix(YO, 5) = RTrim(alumno.f_nacimiento)
'dd = Val(Left(alumno.f_nacimiento, 2))
'mm2 = Right(alumno.f_nacimiento, 7)
'mm = Val(Left(mm2, 2))
'aaaa = Val(Right(alumno.f_nacimiento, 4))
'aaaa = Year(Date) - aaaa
'If mm > Month(Date) Then
'aaaa = aaaa - 1
'End If
'If mm = Month(Date) Then
'   If dd > Day(Date) Then
'   aaaa = aaaa - 1
'   End If
'End If
'MATI9.TextMatrix(YO, 6) = aaaa
'MATI9.TextMatrix(YO, 7) = RTrim(alumno.acudiente)
'MATI9.TextMatrix(YO, 8) = RTrim(alumno.tel_acu)
YO = YO + 1
Next TN
Close #NAR
Close #(NAR - 1)
RESC = Combo1.Text
Text4.Text = YO - 1
Screen.MousePointer = 0
End Sub

Private Sub Command3_Click()
'Dim alugru As grupoalu
'Dim aluper As pertgrup
If Text3.Text = "" Then
    MsgBox "ESCRIBA EL CODIGO A ELIMINAR", 48, "ELIMINAR"
    Text3.SetFocus
    Exit Sub
End If
If ((Val(Text3.Text) > YO - 1) Or (Val(Text3.Text) < 1)) Then
    MsgBox "NO EXISTE ESTE CÓDIGO", 32, "ELIMINAR CÓDIGO"
    Text3.Text = ""
    Text3.SetFocus
    Exit Sub
End If
Screen.MousePointer = 11
aluper.grupo = "SIN GRUPO"
NAR = FreeFile
Open Ruta & "quegru.edu" For Random As #NAR Len = Len(aluper)
Put #NAR, Val(MATI9.TextMatrix(Val(Text3.Text), 1)), aluper
Close #NAR
If Val(Text4.Text) <> 1 Then
    MATI9.RemoveItem Val(Text3.Text)
    YO = YO - 1
    For TT = 1 To (YO - 1)
        MATI9.TextMatrix(TT, 0) = TT
    Next TT
Else
    MATI9.Rows = 1
    YO = YO - 1
End If
Text4.Text = Val(Text4.Text) - 1
    Kill Ruta & RESC & ".gru"
    Open Ruta & RESC & ".gru" For Random As #NAR Len = Len(alugru)
    For we = 1 To (YO - 1)
        alugru.num_carnet = MATI9.TextMatrix(we, 1)
        Put #NAR, we, alugru
    Next we
    Close #NAR
Text3.Text = ""
Text3.SetFocus
Screen.MousePointer = 0
End Sub

Private Sub Command6_Click()
IMP_GRUP.Show
End Sub

Private Sub Command7_Click()
If YO = 0 Then
    MsgBox "ESCOJA EL NOMBRE DEL GRUPO", 32, "CAMBIAR DIRECTOR"
    Combo1.SetFocus
Else
    CAMBIO_DIRECT.Show 1
End If
End Sub

Private Sub Command8_Click()
If YO = 0 Then
    MsgBox "DEBE ESCOGER PRIMERO EL NOMBRE DEL GRUPO", 48, "ADICIONAR"
    Combo1.SetFocus
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
    If (RTrim(alumno.nombres) <> "") And (RTrim(alumno.apellidos) <> "") And (RTrim(alumno.grado) = J2) Then
        Est_Grado2.List1.AddItem RTrim(alumno.apellidos) & " " & RTrim(alumno.nombres) & " - " & Format(alumno.n_carnet, "0000")
        J = J + 1
    End If
    Close #NAR
    NAR = NAR - 1
End If
Next h
Close #NAR
If J <> 0 Then
    Est_Grado2.Caption = "Estudiantes del grado " & J2 & " - Total..." & J
    Est_Grado2.Frame1.Caption = "(Presione <Ctrl> para seleccionar estudiantes o <Shift> para seleccionar en bloque)"
    Est_Grado2.Show 1
Else
    MsgBox "NO HAY ESTUDIANTES DISPONIBLES PARA ESTE GRADO", 32, "CREAR GRUPO"
End If

End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Puede adicionar o retirar alumnos del grupo, como también cambiar de director de grupo."
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command3_Click
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
Private Sub Form_Load()
'Dim icur As inforcur
'MATI9.Row = 0
'MATI9.Col = 0
'MATI9.ColWidth(0) = 450
'MATI9.CellForeColor = RGB(255, 255, 255)
'MATI9.CellBackColor = RGB(0, 0, 150)
'MATI9.Text = "CÓD"
'MATI9.Col = 1
'MATI9.ColWidth(1) = 800
'MATI9.CellForeColor = RGB(255, 255, 255)
'MATI9.CellBackColor = RGB(0, 0, 150)
'MATI9.Text = "CARNET"
'MATI9.Col = 2
'MATI9.ColWidth(2) = 1100
'MATI9.CellForeColor = RGB(255, 255, 255)
'MATI9.CellBackColor = RGB(0, 0, 150)
'MATI9.Text = "MATRICULA"
'MATI9.Col = 3
'MATI9.ColWidth(3) = 2200
'MATI9.CellForeColor = RGB(255, 255, 255)
'MATI9.CellBackColor = RGB(0, 0, 150)
'MATI9.Text = "APELLIDOS"
'MATI9.Col = 4
'MATI9.ColWidth(4) = 2200
'MATI9.CellForeColor = RGB(255, 255, 255)
'MATI9.CellBackColor = RGB(0, 0, 150)
'MATI9.Text = "NOMBRES"
'MATI9.Col = 5
'MATI9.ColWidth(5) = 1100
'MATI9.CellForeColor = RGB(255, 255, 255)
'MATI9.CellBackColor = RGB(0, 0, 150)
'MATI9.Text = "FECH_NACIM"
'MATI9.Col = 6
'MATI9.ColWidth(6) = 600
'MATI9.CellForeColor = RGB(255, 255, 255)
'MATI9.CellBackColor = RGB(0, 0, 150)
'MATI9.Text = "EDAD"
'MATI9.Col = 7
'MATI9.ColWidth(7) = 3300
'MATI9.CellForeColor = RGB(255, 255, 255)
'MATI9.CellBackColor = RGB(0, 0, 150)
'MATI9.Text = "ACUDIENTE"
'MATI9.Col = 8
'MATI9.ColWidth(8) = 1300
'MATI9.CellForeColor = RGB(255, 255, 255)
'MATI9.CellBackColor = RGB(0, 0, 150)
'MATI9.Text = "TELEFONO"
MATI9.Row = 0
MATI9.Col = 0
MATI9.ColWidth(0) = 450
MATI9.CellForeColor = RGB(255, 255, 255)
MATI9.CellBackColor = RGB(0, 0, 150)
MATI9.Text = "CÓD"
MATI9.Col = 1
MATI9.ColWidth(1) = 800
MATI9.CellForeColor = RGB(255, 255, 255)
MATI9.CellBackColor = RGB(0, 0, 150)
MATI9.Text = "CARNET"
MATI9.Col = 2
MATI9.ColWidth(2) = 4000
MATI9.CellForeColor = RGB(255, 255, 255)
MATI9.CellBackColor = RGB(0, 0, 150)
MATI9.Text = "APELLIDOS Y NOMBRES"

'Text2.MaxLength = 5
Text3.MaxLength = 2
If Dir(Ruta & "infcur.edu") <> "" Then
Command1.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
NAR = FreeFile
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
Combo1.AddItem RTrim(icur.nom)
Wend
Close #NAR
Combo1.Text = Combo1.List(0)
Else
Command1.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
End If
YO = 0
End Sub
